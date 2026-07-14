import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { storage } from "./storage";

/* ============================================================
   LMA OCR EXTRACTOR — Registro de Experiencia Básica Part 66
   Flujo: Foto/PDF de OT → Claude Vision → Validación → Tabla → XLSX
   Persistencia: storage (clave única "lma-entries")
   ============================================================ */

const STORAGE_KEY = "lma-entries";
const BACKUP_KEY = "lma-entries-bak"; // copia de seguridad automática redundante
const LEARN_KEY = "lma-learning"; // reglas aprendidas + historial de correcciones
const PHOTO_PREFIX = "lma-foto-"; // una clave por foto adjunta (límite 5MB/clave)
const QUEUE_KEY = "lma-queue"; // cola offline de escaneos pendientes de OCR
const SYNCCODE_KEY = "lma-sync-code";
const TITULAR_KEY = "lma-titular"; // nombre y nº licencia para portada del logbook
/* Sincronización opcional (Supabase). La app standalone inyecta estas variables
   desde el entorno de Vercel en main.jsx; en el artifact quedan vacías y la UI se oculta. */
const SUPA_URL = (typeof globalThis !== "undefined" && globalThis.LMA_SUPABASE_URL) || "";
const SUPA_KEY = (typeof globalThis !== "undefined" && globalThis.LMA_SUPABASE_ANON_KEY) || "";
const SYNC_ENABLED = Boolean(SUPA_URL && SUPA_KEY);
const MESES = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];

const BASE_EXTRACTION_PROMPT = `Eres especialista en órdenes de trabajo de mantenimiento aeronáutico (Part 145, entorno Ryanair/B737).

Lee este documento (foto o PDF de una orden de trabajo) y extrae SOLO estos datos:
- Matrícula del avión sin guion (ej: EIEMD, 9HQFM, SPRZC)
- Fecha de ejecución en formato DD/MM/YYYY
- Código ATA (2 dígitos, ej: 22, 33, 36)
- Descripción de la tarea EN INGLÉS y en MAYÚSCULAS, incluyendo referencia AMM y revisión si figuran (ej: "STARTER REPLACED IAW AMM 80-11-01-400-801-F00 REV.89"), máx 120 caracteres
- Modelo de avión (ej: B737-8, B737-800, A320)
- Número de Workorder (alfanumérico)
- Números de tarea/SEQ si hay varios (array de enteros)
- Certificador: nombre o número de licencia/sello del técnico que certifica la tarea (firma, stamp, "Certified by", "CRS")

RESPONDE ÚNICAMENTE CON JSON VÁLIDO, sin markdown ni texto adicional:
{
  "matricula": "EIEMD",
  "fecha": "29/06/2026",
  "ata": 22,
  "descripcion": "MAIN BATTERY REPLACED IAW AMM 24-31-11/401 REV.89",
  "modelo": "B737-8",
  "workorder": "191387705",
  "seq": [45, 67],
  "certificador": "J. GARCIA / ES.66.12345",
  "confidence": 0.95,
  "warnings": []
}

Si un campo no es legible, pon null y añade una nota en "warnings".`;

/**
 * Prompt dinámico: base + reglas aprendidas (hints) + correcciones previas (few-shot).
 * El "aprendizaje" del sistema es a nivel de prompt: cada corrección del usuario
 * se acumula y se inyecta en futuras extracciones.
 */
function buildExtractionPrompt(learning) {
  let prompt = BASE_EXTRACTION_PROMPT;
  if (learning?.hints?.length) {
    prompt += `\n\nREGLAS APRENDIDAS DE ESTE OPERADOR (prioridad alta, aplícalas siempre):\n`;
    prompt += learning.hints.map((h, i) => `${i + 1}. ${h.text}`).join("\n");
  }
  const ex = (learning?.examples || []).slice(-12); // últimas 12 correcciones
  if (ex.length) {
    prompt += `\n\nCORRECCIONES PREVIAS DEL USUARIO (el valor extraído era erróneo, aprende el patrón):\n`;
    prompt += ex.map((e) => `- Campo "${e.field}": extraje "${e.extracted}" → correcto era "${e.corrected}"`).join("\n");
  }
  return prompt;
}

/* ---------- Utilidades ---------- */

const uuid = () => (crypto?.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2));

function parseFecha(str) {
  if (!str) return null;
  const m = String(str).trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  if (d.getDate() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1) return null;
  return d;
}

function fmtFecha(str) {
  const d = parseFecha(str);
  if (!d) return str || "";
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
}

function monthKey(str) {
  const d = parseFecha(str);
  if (!d) return "SIN FECHA";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(key) {
  if (key === "SIN FECHA") return key;
  const [y, m] = key.split("-");
  return `${MESES[Number(m) - 1]} ${y}`;
}

function seqToString(seq) {
  if (Array.isArray(seq)) return seq.join(", ");
  return seq ? String(seq) : "";
}

function stringToSeq(str) {
  if (!str) return [];
  return String(str)
    .split(/[,;\s]+/)
    .map((s) => parseInt(s, 10))
    .filter((n) => !isNaN(n));
}

/* ---------- Validación (según especificación LMA) ---------- */

function validateRow(row) {
  const errors = {};
  const mat = (row.matricula || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!/^[A-Z0-9]{4,7}$/.test(mat)) errors.matricula = "Matrícula inválida (4–7 alfanuméricos, sin guion)";

  const d = parseFecha(row.fecha);
  if (!d) errors.fecha = "Formato DD/MM/YYYY requerido";
  else if (d > new Date()) errors.fecha = "La fecha no puede ser futura";

  const ata = parseInt(row.ata, 10);
  if (isNaN(ata) || ata < 1 || ata > 99) errors.ata = "ATA debe ser un entero 1–99";

  if (!row.descripcion || !String(row.descripcion).trim()) errors.descripcion = "Descripción obligatoria";
  else if (String(row.descripcion).length > 120) errors.descripcion = "Máx. 120 caracteres";

  if (!row.modelo || !/^[A-Z]\d{3}/.test(String(row.modelo).toUpperCase())) errors.modelo = "Modelo inválido (ej: B737-8)";

  if (!row.workorder || String(row.workorder).length > 20) errors.workorder = "WO obligatorio, máx. 20 caracteres";

  if (row.certificador && String(row.certificador).length > 60) errors.certificador = "Máx. 60 caracteres";

  return { valid: Object.keys(errors).length === 0, errors };
}

/**
 * Detección de duplicados: mismo Workorder + misma fecha y, si ambos tienen
 * SEQ, al menos un SEQ en común. Devuelve la entrada existente o null.
 */
function findDuplicate(row, rows) {
  const wo = String(row.workorder || "").trim().toUpperCase();
  if (!wo) return null;
  const f = fmtFecha(row.fecha);
  const seqA = stringToSeq(row.seq);
  return (
    rows.find((r) => {
      if (String(r.workorder).trim().toUpperCase() !== wo) return false;
      if (r.fecha !== f) return false;
      const seqB = Array.isArray(r.seq) ? r.seq : [];
      if (!seqA.length || !seqB.length) return true;
      return seqA.some((s) => seqB.includes(s));
    }) || null
  );
}

/* ---------- Claude Vision ---------- */

async function fileToBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("No se pudo leer el archivo"));
    r.readAsDataURL(file);
  });
}

/* ---------- Transporte API unificado ---------- */

async function callClaude(content) {
  const response = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{ role: "user", content }],
    }),
  });
  const raw = await response.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    /* Respuestas no-JSON del servidor: "Request Entity Too Large" (413), HTML de error, etc. */
    throw new Error(`Respuesta no válida del servidor (HTTP ${response.status}): ${raw.slice(0, 80)}`);
  }
  if (data.error) throw new Error(data.error.message || "Error de API");
  return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
}

function parseClaudeJson(text) {
  const clean = text.replace(/```json|```/g, "").trim();
  const m = clean.match(/\{[\s\S]*\}/);
  if (!m) throw new Error("La respuesta del modelo no contiene JSON válido");
  return JSON.parse(m[0]);
}

function normalizeExtraction(parsed) {
  return {
    matricula: parsed.matricula ? String(parsed.matricula).toUpperCase().replace(/[^A-Z0-9]/g, "") : "",
    fecha: parsed.fecha ? fmtFecha(parsed.fecha) : "",
    ata: parsed.ata ?? "",
    descripcion: parsed.descripcion || "",
    modelo: parsed.modelo ? String(parsed.modelo).toUpperCase() : "",
    workorder: parsed.workorder ? String(parsed.workorder) : "",
    seq: Array.isArray(parsed.seq) ? parsed.seq : stringToSeq(parsed.seq),
    certificador: parsed.certificador ? String(parsed.certificador) : "",
    confidence: typeof parsed.confidence === "number" ? parsed.confidence : null,
    warnings: Array.isArray(parsed.warnings) ? parsed.warnings : [],
  };
}

/** OCR desde base64: reutilizado por el upload directo y por la cola offline */
async function scanFromBase64(base64, mediaType, prompt) {
  const isPdf = mediaType === "application/pdf";
  const block = isPdf
    ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } }
    : { type: "image", source: { type: "base64", media_type: mediaType || "image/jpeg", data: base64 } };
  return normalizeExtraction(parseClaudeJson(await callClaude([block, { type: "text", text: prompt }])));
}


/**
 * Comprime una imagen a JPEG (máx. 1600 px lado mayor, calidad 0,8):
 * suficiente resolución para OCR de texto de OT y muy por debajo del
 * límite del endpoint (~4,5MB) y del almacenamiento (5MB/clave).
 */
async function compressImage(file, maxDim = 1600, quality = 0.8) {
  return new Promise((res, rej) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
      const c = document.createElement("canvas");
      c.width = Math.round(img.width * scale);
      c.height = Math.round(img.height * scale);
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      res(c.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      rej(new Error("Imagen no válida"));
    };
    img.src = url;
  });
}

/* ---------- Componentes de campo ---------- */

const FIELD_DEFS = [
  { key: "matricula", label: "Matrícula", mono: true, w: "w-24" },
  { key: "fecha", label: "Fecha", mono: true, w: "w-28" },
  { key: "ata", label: "ATA", mono: true, w: "w-16" },
  { key: "descripcion", label: "Descripción", mono: false, w: "w-full" },
  { key: "modelo", label: "Modelo", mono: true, w: "w-24" },
  { key: "workorder", label: "Workorder", mono: true, w: "w-32" },
  { key: "seq", label: "SEQ", mono: true, w: "w-24" },
  { key: "certificador", label: "Certificador", mono: false, w: "w-32" },
];

function FieldInput({ def, value, error, onChange }) {
  return (
    <div className="flex flex-col gap-1">
      <label className="text-xs font-semibold uppercase tracking-wider text-slate-500">{def.label}</label>
      <input
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value)}
        className={`px-2 py-1.5 rounded border text-sm bg-white ${def.mono ? "font-mono" : ""} ${
          error ? "border-red-500 bg-red-50" : "border-slate-300"
        } focus:outline-none focus:border-slate-600`}
      />
      {error && <span className="text-xs text-red-600">{error}</span>}
    </div>
  );
}

/* ---------- App principal ---------- */

export default function LmaOcrExtractor() {
  const [rows, setRows] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [pending, setPending] = useState([]); // extracciones pendientes de confirmar
  const [queue, setQueue] = useState(0); // archivos en proceso OCR
  const [search, setSearch] = useState("");
  const [filterAta, setFilterAta] = useState("");
  const [filterMonth, setFilterMonth] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const [toast, setToast] = useState(null);
  const [learning, setLearning] = useState({ hints: [], examples: [] });
  const [showLearn, setShowLearn] = useState(false);
  const [showDash, setShowDash] = useState(false);
  const [newHint, setNewHint] = useState("");
  const [synthesizing, setSynthesizing] = useState(false);
  const [photoModal, setPhotoModal] = useState(null); // { rowId, url } | { loading: true }
  const [cola, setCola] = useState([]); // escaneos offline pendientes de OCR
  const [selectedIds, setSelectedIds] = useState([]); // marcadas para export AESA/Logbook
  const [syncCode, setSyncCode] = useState("");
  const [syncBusy, setSyncBusy] = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const [titular, setTitular] = useState({ nombre: "", licencia: "" });
  const fileRef = useRef(null);
  const cameraRef = useRef(null); // input con capture: abre la cámara directamente en móvil
  const attachRef = useRef(null); // input oculto para adjuntar foto a una fila existente
  const attachRowId = useRef(null);
  const lastSaved = useRef({}); // último valor persistido por clave (dedupe)
  const saveTimers = useRef({}); // timers de debounce por clave
  const storageWarned = useRef(false);

  const storageAvailable = () => typeof indexedDB !== "undefined";

  /* Carga inicial desde storage */
  useEffect(() => {
    (async () => {
      if (!storageAvailable()) {
        setLoaded(true);
        showToast("⚠️ Almacenamiento no disponible: los datos viven solo en esta sesión. Usa Backup JSON.", true);
        storageWarned.current = true;
        return;
      }
      try {
        const result = await storage.get(STORAGE_KEY);
        if (result?.value) {
          setRows(JSON.parse(result.value));
          lastSaved.current[STORAGE_KEY] = result.value; // evita re-escritura inmediata
        }
      } catch (e) {
        /* clave principal inexistente: intentar restaurar desde la copia de seguridad */
        try {
          const bak = await storage.get(BACKUP_KEY);
          if (bak?.value) {
            setRows(JSON.parse(bak.value));
            showToast("↺ Datos restaurados desde la copia de seguridad automática");
          }
        } catch (e2) {
          /* primer uso: sin datos */
        }
      }
      try {
        const learn = await storage.get(LEARN_KEY);
        if (learn?.value) {
          setLearning(JSON.parse(learn.value));
          lastSaved.current[LEARN_KEY] = learn.value;
        }
      } catch (e) {
        /* sin aprendizaje previo */
      }
      try {
        const q = await storage.get(QUEUE_KEY);
        if (q?.value) {
          setCola(JSON.parse(q.value));
          lastSaved.current[QUEUE_KEY] = q.value;
        }
      } catch (e) {
        /* sin cola offline */
      }
      try {
        const sc = await storage.get(SYNCCODE_KEY);
        if (sc?.value) {
          setSyncCode(JSON.parse(sc.value));
          lastSaved.current[SYNCCODE_KEY] = sc.value;
        }
      } catch (e) {
        /* sin código de sincronización */
      }
      try {
        const t = await storage.get(TITULAR_KEY);
        if (t?.value) {
          setTitular(JSON.parse(t.value));
          lastSaved.current[TITULAR_KEY] = t.value;
        }
      } catch (e) {
        /* sin datos de titular */
      }
      setLoaded(true);
    })();
    return () => Object.values(saveTimers.current).forEach(clearTimeout);
  }, []);

  /**
   * Persistidor centralizado:
   * - Dedupe: no escribe si el valor no cambió desde el último guardado exitoso.
   * - Debounce 800 ms: agrupa cambios rápidos en una sola escritura (evita rate limit).
   * - shared=false explícito, verificación de resultado y un reintento a los 1,5 s.
   */
  function schedulePersist(key, obj, label) {
    const value = JSON.stringify(obj);
    if (lastSaved.current[key] === value) return;
    if (!storageAvailable()) {
      if (!storageWarned.current) {
        showToast("⚠️ Almacenamiento no disponible: usa Backup JSON para no perder datos.", true);
        storageWarned.current = true;
      }
      return;
    }
    clearTimeout(saveTimers.current[key]);
    saveTimers.current[key] = setTimeout(async () => {
      const attempt = async () => {
        const res = await storage.set(key, value, false);
        if (!res) throw new Error("respuesta vacía del almacenamiento");
        lastSaved.current[key] = value;
        if (key === STORAGE_KEY) {
          /* copia redundante best-effort: si falla, no interrumpe */
          try { await storage.set(BACKUP_KEY, value, false); } catch (e) {}
        }
      };
      try {
        await attempt();
      } catch (e1) {
        setTimeout(async () => {
          try {
            await attempt();
          } catch (e2) {
            showToast(`⚠️ Error al guardar ${label} (${e2.message}). Exporta Backup JSON como respaldo.`, true);
          }
        }, 1500);
      }
    }, 800);
  }

  /* Guardado automático (debounced) */
  useEffect(() => {
    if (loaded) schedulePersist(STORAGE_KEY, rows, "entradas");
  }, [rows, loaded]);

  useEffect(() => {
    if (loaded) schedulePersist(LEARN_KEY, learning, "aprendizaje");
  }, [learning, loaded]);

  useEffect(() => {
    if (loaded) schedulePersist(QUEUE_KEY, cola, "cola offline");
  }, [cola, loaded]);

  useEffect(() => {
    if (loaded && syncCode) schedulePersist(SYNCCODE_KEY, syncCode, "código de sincronización");
  }, [syncCode, loaded]);

  useEffect(() => {
    if (loaded && (titular.nombre || titular.licencia)) schedulePersist(TITULAR_KEY, titular, "datos del titular");
  }, [titular, loaded]);

  function showToast(msg, isError = false) {
    setToast({ msg, isError });
    setTimeout(() => setToast(null), 3500);
  }

  /* ---------- OCR ---------- */

  async function handleFiles(files) {
    const list = Array.from(files || []);
    if (!list.length) return;
    /* Sin conexión: encolar fotos comprimidas; se procesan al recuperar red */
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      for (const file of list) {
        if (file.type === "application/pdf") {
          showToast(`⚠️ ${file.name}: los PDF requieren conexión (no encolado)`, true);
          continue;
        }
        try {
          const dataUrl = await compressImage(file);
          setCola((c) => [...c, { id: uuid(), fileName: file.name, dataUrl, ts: new Date().toISOString() }]);
        } catch (e) {
          showToast(`⚠️ ${file.name}: ${e.message}`, true);
        }
      }
      showToast("📥 Sin conexión: escaneos guardados en cola offline");
      if (fileRef.current) fileRef.current.value = "";
      if (cameraRef.current) cameraRef.current.value = "";
      return;
    }
    setQueue((q) => q + list.length);
    const prompt = buildExtractionPrompt(learning); // inyecta reglas + correcciones acumuladas
    for (const file of list) {
      try {
        /* Comprimir SIEMPRE la imagen antes del OCR: las fotos de cámara móvil
           (3–8MB) superan el límite del endpoint y provocan errores 413 no-JSON.
           La misma imagen comprimida sirve de adjunto (una sola compresión). */
        let extracted;
        let photoData = null;
        if (file.type === "application/pdf") {
          const base64 = await fileToBase64(file);
          if (base64.length > 4.2 * 1024 * 1024) throw new Error("PDF demasiado grande (>4MB): fotografía la página en su lugar");
          extracted = await scanFromBase64(base64, "application/pdf", prompt);
        } else {
          photoData = await compressImage(file);
          extracted = await scanFromBase64(photoData.split(",")[1], "image/jpeg", prompt);
        }
        const flat = { ...extracted, seq: seqToString(extracted.seq) };
        setPending((p) => [
          ...p,
          { ...flat, id: uuid(), fileName: file.name, ocrOriginal: { ...flat }, photoData },
        ]);
      } catch (e) {
        showToast(`Error OCR en ${file.name}: ${e.message}`, true);
      } finally {
        setQueue((q) => q - 1);
      }
    }
    if (fileRef.current) fileRef.current.value = "";
    if (cameraRef.current) cameraRef.current.value = "";
  }

  /** Procesa la cola offline: OCR de cada foto encolada, con la foto como adjunto */
  async function processQueue() {
    if (!cola.length) return;
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      showToast("Sin conexión: la cola se procesará automáticamente al recuperar red", true);
      return;
    }
    const items = [...cola];
    setQueue((q) => q + items.length);
    const prompt = buildExtractionPrompt(learning);
    for (const item of items) {
      try {
        const extracted = await scanFromBase64(item.dataUrl.split(",")[1], "image/jpeg", prompt);
        const flat = { ...extracted, seq: seqToString(extracted.seq) };
        setPending((p) => [
          ...p,
          { ...flat, id: uuid(), fileName: `${item.fileName} (cola)`, ocrOriginal: { ...flat }, photoData: item.dataUrl },
        ]);
        setCola((c) => c.filter((x) => x.id !== item.id));
      } catch (e) {
        showToast(`Error OCR en cola (${item.fileName}): ${e.message}`, true);
      } finally {
        setQueue((q) => q - 1);
      }
    }
  }

  const processQueueRef = useRef(null);
  processQueueRef.current = processQueue;

  /* Al recuperar conexión, procesar la cola automáticamente */
  useEffect(() => {
    const onOnline = () => {
      showToast("🌐 Conexión recuperada: procesando cola offline…");
      processQueueRef.current?.();
    };
    window.addEventListener("online", onOnline);
    return () => window.removeEventListener("online", onOnline);
  }, []);

  function addManual() {
    setPending((p) => [
      ...p,
      { id: uuid(), fileName: "ENTRADA MANUAL", matricula: "", fecha: "", ata: "", descripcion: "", modelo: "B737-8", workorder: "", seq: "", certificador: "", confidence: null, warnings: [] },
    ]);
  }

  function updatePending(id, key, value) {
    setPending((p) => p.map((r) => (r.id === id ? { ...r, [key]: value } : r)));
  }

  function confirmPending(id) {
    const row = pending.find((r) => r.id === id);
    const { valid } = validateRow(row);
    if (!valid) {
      showToast("Corrige los campos marcados antes de confirmar", true);
      setPending((p) => p.map((r) => (r.id === id ? { ...r, showErrors: true } : r)));
      return;
    }
    /* Duplicado: primer intento avisa, segundo intento fuerza el guardado */
    const dup = findDuplicate(row, rows);
    if (dup && !row.forceDup) {
      setPending((p) => p.map((r) => (r.id === id ? { ...r, forceDup: true } : r)));
      showToast(`⚠️ Posible duplicado: WO ${dup.workorder} ya registrado el ${dup.fecha}. Pulsa Confirmar de nuevo para forzar.`, true);
      return;
    }
    const entry = {
      id: uuid(),
      matricula: row.matricula.toUpperCase(),
      fecha: fmtFecha(row.fecha),
      ata: parseInt(row.ata, 10),
      descripcion: String(row.descripcion).trim(),
      modelo: row.modelo.toUpperCase(),
      workorder: String(row.workorder),
      seq: stringToSeq(row.seq),
      certificador: String(row.certificador || "").trim(),
      hasPhoto: Boolean(row.photoData),
      status: "Confirmado",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    setRows((r) => [...r, entry]);
    setPending((p) => p.filter((r) => r.id !== id));

    /* Guardar la foto de la OT en su propia clave (una clave por foto) */
    if (row.photoData) savePhoto(entry.id, row.photoData);

    /* Aprendizaje: si el usuario corrigió campos respecto al OCR, registrar el diff */
    if (row.ocrOriginal) {
      const fields = ["matricula", "fecha", "ata", "descripcion", "modelo", "workorder", "seq", "certificador"];
      const corrections = fields
        .filter((f) => {
          const a = String(row.ocrOriginal[f] ?? "").trim().toUpperCase();
          const b = String(row[f] ?? "").trim().toUpperCase();
          return a !== b && a !== "";
        })
        .map((f) => ({
          field: f,
          extracted: String(row.ocrOriginal[f] ?? ""),
          corrected: String(row[f] ?? ""),
          ts: new Date().toISOString(),
        }));
      if (corrections.length) {
        setLearning((l) => ({ ...l, examples: [...l.examples, ...corrections].slice(-30) }));
        showToast(`✓ Guardada · ${corrections.length} corrección(es) registrada(s) para aprendizaje`);
        return;
      }
    }
    showToast(`✓ Entrada ${entry.matricula} · WO ${entry.workorder} guardada`);
  }

  function discardPending(id) {
    setPending((p) => p.filter((r) => r.id !== id));
  }

  /** Confirma en lote todos los pendientes válidos y sin duplicado; deja el resto para revisión manual */
  function confirmAllPending() {
    const snapshot = [...pending];
    let ok = 0, skipped = 0;
    snapshot.forEach((p) => {
      if (validateRow(p).valid && !findDuplicate(p, rows)) {
        confirmPending(p.id);
        ok++;
      } else {
        skipped++;
      }
    });
    if (skipped) showToast(`✓ ${ok} confirmadas · ${skipped} requieren revisión manual (errores o duplicado)`, ok === 0);
  }

  /* ---------- Fotos adjuntas (una clave de almacenamiento por foto) ---------- */

  async function savePhoto(entryId, dataUrl) {
    if (!storageAvailable()) {
      showToast("⚠️ Almacenamiento no disponible: la foto no puede persistir", true);
      return false;
    }
    if (dataUrl.length > 4.5 * 1024 * 1024) {
      showToast("⚠️ Foto demasiado grande incluso comprimida (>4,5MB): no adjuntada", true);
      return false;
    }
    try {
      const res = await storage.set(PHOTO_PREFIX + entryId, dataUrl, false);
      if (!res) throw new Error("respuesta vacía");
      return true;
    } catch (e) {
      showToast("⚠️ Error al guardar la foto: " + e.message, true);
      return false;
    }
  }

  async function viewPhoto(row) {
    setPhotoModal({ loading: true });
    try {
      const res = await storage.get(PHOTO_PREFIX + row.id);
      if (!res?.value) throw new Error("foto no encontrada");
      setPhotoModal({ rowId: row.id, url: res.value, label: `${row.matricula} · WO ${row.workorder}` });
    } catch (e) {
      setPhotoModal(null);
      showToast("⚠️ No se pudo cargar la foto: " + e.message, true);
      /* Coherencia: si la clave no existe, corregir el flag */
      setRows((r) => r.map((x) => (x.id === row.id ? { ...x, hasPhoto: false } : x)));
    }
  }

  function requestAttach(rowId) {
    attachRowId.current = rowId;
    attachRef.current?.click();
  }

  async function handleAttachFile(e) {
    const file = e.target.files?.[0];
    const rowId = attachRowId.current;
    if (!file || !rowId) return;
    try {
      const dataUrl = await compressImage(file);
      const ok = await savePhoto(rowId, dataUrl);
      if (ok) {
        setRows((r) => r.map((x) => (x.id === rowId ? { ...x, hasPhoto: true, updatedAt: new Date().toISOString() } : x)));
        showToast("✓ Foto adjuntada y guardada");
      }
    } catch (err) {
      showToast("⚠️ " + err.message, true);
    } finally {
      attachRowId.current = null;
      if (attachRef.current) attachRef.current.value = "";
    }
  }

  async function removePhoto(row) {
    try {
      await storage.delete(PHOTO_PREFIX + row.id, false);
    } catch (e) {
      /* clave ya inexistente: continuar */
    }
    setRows((r) => r.map((x) => (x.id === row.id ? { ...x, hasPhoto: false, updatedAt: new Date().toISOString() } : x)));
    setPhotoModal(null);
    showToast("Foto eliminada de la entrada");
  }

  /* ---------- Aprendizaje: reglas manuales + síntesis automática ---------- */

  function addHint() {
    const text = newHint.trim();
    if (!text) return;
    setLearning((l) => ({ ...l, hints: [...l.hints, { id: uuid(), text, source: "manual" }] }));
    setNewHint("");
    showToast("✓ Regla añadida: se aplicará en los próximos escaneos");
  }

  function removeHint(id) {
    setLearning((l) => ({ ...l, hints: l.hints.filter((h) => h.id !== id) }));
  }

  /**
   * Sintetiza el historial de correcciones en reglas de extracción concisas
   * (llamada a Claude) y las incorpora como hints permanentes.
   */
  async function synthesizeRules() {
    if (!learning.examples.length) {
      showToast("No hay correcciones registradas todavía", true);
      return;
    }
    setSynthesizing(true);
    try {
      const response = await fetch("/api/extract", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          messages: [
            {
              role: "user",
              content: `Contexto: sistema OCR de órdenes de trabajo de mantenimiento aeronáutico (Part 145, flota B737 Ryanair/Malta Air/Buzz).

Historial de correcciones del usuario (valor extraído por OCR → valor correcto):
${learning.examples.map((e) => `- Campo "${e.field}": "${e.extracted}" → "${e.corrected}"`).join("\n")}

Reglas ya existentes (no las repitas):
${learning.hints.map((h) => `- ${h.text}`).join("\n") || "(ninguna)"}

Sintetiza los patrones de error en reglas de extracción concisas y generalizables (máx 5), en español. Ejemplos de buen formato: "La matrícula EI puede confundirse con El: verificar que el segundo carácter es I mayúscula", "El workorder aparece bajo la etiqueta W/O No. en la esquina superior derecha".

RESPONDE SOLO CON JSON: {"rules": ["regla 1", "regla 2"]}`,
            },
          ],
        }),
      });
      const data = await response.json();
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
      const parsed = JSON.parse(text.replace(/```json|```/g, "").trim().match(/\{[\s\S]*\}/)[0]);
      const existing = new Set(learning.hints.map((h) => h.text.toLowerCase()));
      const nuevas = (parsed.rules || []).filter((r) => !existing.has(String(r).toLowerCase()));
      if (!nuevas.length) {
        showToast("Sin patrones nuevos que sintetizar");
        return;
      }
      setLearning((l) => ({
        ...l,
        hints: [...l.hints, ...nuevas.map((text) => ({ id: uuid(), text, source: "auto" }))],
      }));
      showToast(`✓ ${nuevas.length} regla(s) sintetizada(s) desde correcciones`);
    } catch (e) {
      showToast("Error al sintetizar reglas: " + e.message, true);
    } finally {
      setSynthesizing(false);
    }
  }

  /* ---------- Edición inline en tabla ---------- */

  function startEdit(row) {
    setEditingId(row.id);
    setEditDraft({ ...row, seq: seqToString(row.seq) });
  }

  function saveEdit() {
    const { valid, errors } = validateRow(editDraft);
    if (!valid) {
      showToast("Errores: " + Object.values(errors).join(" · "), true);
      return;
    }
    setRows((r) =>
      r.map((row) =>
        row.id === editingId
          ? {
              ...row,
              matricula: editDraft.matricula.toUpperCase(),
              fecha: fmtFecha(editDraft.fecha),
              ata: parseInt(editDraft.ata, 10),
              descripcion: String(editDraft.descripcion).trim(),
              modelo: editDraft.modelo.toUpperCase(),
              workorder: String(editDraft.workorder),
              seq: stringToSeq(editDraft.seq),
              certificador: String(editDraft.certificador || "").trim(),
              updatedAt: new Date().toISOString(),
            }
          : row
      )
    );
    setEditingId(null);
    setEditDraft(null);
    showToast("✓ Cambios guardados");
  }

  function toggleStatus(id) {
    /* Soft delete: alterna Confirmado ↔ Borrador. No hay borrado permanente. */
    setRows((r) =>
      r.map((row) =>
        row.id === id
          ? { ...row, status: row.status === "Confirmado" ? "Borrador" : "Confirmado", updatedAt: new Date().toISOString() }
          : row
      )
    );
  }

  /* ---------- Filtros y búsqueda ---------- */

  const months = useMemo(() => {
    const set = new Set(rows.map((r) => monthKey(r.fecha)));
    return Array.from(set).sort().reverse();
  }, [rows]);

  const atas = useMemo(() => {
    const set = new Set(rows.map((r) => r.ata));
    return Array.from(set).sort((a, b) => a - b);
  }, [rows]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return rows
      .filter((r) => {
        if (filterAta && String(r.ata) !== filterAta) return false;
        if (filterMonth && monthKey(r.fecha) !== filterMonth) return false;
        if (filterStatus && r.status !== filterStatus) return false;
        if (!q) return true;
        return [r.matricula, r.fecha, r.ata, r.descripcion, r.modelo, r.workorder, seqToString(r.seq), r.certificador]
          .some((v) => String(v ?? "").toLowerCase().includes(q));
      })
      .sort((a, b) => (parseFecha(b.fecha)?.getTime() || 0) - (parseFecha(a.fecha)?.getTime() || 0));
  }, [rows, search, filterAta, filterMonth, filterStatus]);

  /* ---------- Exportación ---------- */

  function exportXlsx() {
    const confirmed = rows.filter((r) => r.status === "Confirmado");
    if (!confirmed.length) {
      showToast("No hay entradas confirmadas para exportar", true);
      return;
    }
    const wb = XLSX.utils.book_new();
    const byMonth = {};
    confirmed.forEach((r) => {
      const k = monthKey(r.fecha);
      (byMonth[k] = byMonth[k] || []).push(r);
    });
    Object.keys(byMonth)
      .sort()
      .forEach((k) => {
        const data = byMonth[k]
          .sort((a, b) => (parseFecha(a.fecha)?.getTime() || 0) - (parseFecha(b.fecha)?.getTime() || 0))
          .map((r) => ({
            "Matrícula": r.matricula,
            "Fecha": r.fecha,
            "ATA": r.ata,
            "Descripción": r.descripcion,
            "Modelo": r.modelo,
            "Workorder": r.workorder,
            "SEQ": seqToString(r.seq),
            "Certificador": r.certificador || "",
          }));
        const ws = XLSX.utils.json_to_sheet(data);
        ws["!cols"] = [{ wch: 10 }, { wch: 12 }, { wch: 6 }, { wch: 50 }, { wch: 10 }, { wch: 16 }, { wch: 14 }, { wch: 22 }];
        XLSX.utils.book_append_sheet(wb, ws, monthLabel(k).replace(/\s/g, "_"));
      });
    XLSX.writeFile(wb, `LMA_Experiencia_Basica_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast(`✓ Exportadas ${confirmed.length} entradas (${Object.keys(byMonth).length} hojas)`);
  }

  function exportJson() {
    const blob = new Blob([JSON.stringify(rows, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `LMA_backup_${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  /* ---------- Selección para exportaciones ---------- */

  function toggleSelect(id) {
    setSelectedIds((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  }

  function toggleSelectAllFiltered() {
    const ids = filtered.map((r) => r.id);
    const all = ids.length > 0 && ids.every((id) => selectedIds.includes(id));
    setSelectedIds(all ? selectedIds.filter((id) => !ids.includes(id)) : Array.from(new Set([...selectedIds, ...ids])));
  }

  /* ---------- Sincronización opcional entre dispositivos (Supabase) ---------- */

  async function syncPush() {
    if (!syncCode.trim()) return showToast("Define un código de sincronización primero", true);
    setSyncBusy(true);
    try {
      const payload = { rows, learning, ts: new Date().toISOString() };
      const res = await fetch(`${SUPA_URL}/rest/v1/lma_sync`, {
        method: "POST",
        headers: {
          apikey: SUPA_KEY,
          Authorization: `Bearer ${SUPA_KEY}`,
          "Content-Type": "application/json",
          Prefer: "resolution=merge-duplicates",
        },
        body: JSON.stringify({ id: syncCode.trim(), data: payload, updated_at: payload.ts }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setLastSync(payload.ts);
      showToast(`✓ ${rows.length} entradas subidas a la nube (las fotos no viajan)`);
    } catch (e) {
      showToast("⚠️ Error al subir: " + e.message, true);
    } finally {
      setSyncBusy(false);
    }
  }

  async function syncPull() {
    if (!syncCode.trim()) return showToast("Define un código de sincronización primero", true);
    setSyncBusy(true);
    try {
      const res = await fetch(`${SUPA_URL}/rest/v1/lma_sync?id=eq.${encodeURIComponent(syncCode.trim())}&select=data`, {
        headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` },
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const arr = await res.json();
      if (!arr.length || !arr[0]?.data) return showToast("No hay datos en la nube con ese código", true);
      const remote = arr[0].data;
      const ok = window.confirm(
        `Descargar ${remote.rows?.length ?? 0} entradas de la nube (${remote.ts || "sin fecha"}) y SUSTITUIR las locales?`
      );
      if (!ok) return;
      setRows(remote.rows || []);
      if (remote.learning) setLearning(remote.learning);
      setLastSync(remote.ts || new Date().toISOString());
      showToast("✓ Datos descargados de la nube");
    } catch (e) {
      showToast("⚠️ Error al descargar: " + e.message, true);
    } finally {
      setSyncBusy(false);
    }
  }

  /* ---------- Export AESA (calibrado con el formato oficial) ---------- */

  /** Matrícula con guion para el formato AESA: EIEMD -> EI-EMD, 9HQBR -> 9H-QBR */
  const fmtMatriculaGuion = (m) => (m && m.length >= 4 ? `${m.slice(0, 2)}-${m.slice(2)}` : m || "");
  /** Modelo según AESA: B737-800 -> 737-800 */
  const fmtModeloAesa = (m) => String(m || "").replace(/^B(?=7)/i, "");

  /** Formato oficial: ATA · Descripción Tarea · Modelo de aeronave · Matricula · Fecha de realización · Workorder · Stamp */
  function exportAesa() {
    const sel = rows
      .filter((r) => selectedIds.includes(r.id) && r.status === "Confirmado")
      .sort((a, b) => (parseFecha(a.fecha)?.getTime() || 0) - (parseFecha(b.fecha)?.getTime() || 0));
    if (!sel.length) {
      showToast("Marca en la tabla las tareas a incluir (casillas de la primera columna)", true);
      return;
    }
    const header = [["ATA", "Descripción Tarea", "Modelo de aeronave", "Matricula", "Fecha de realización", "Workorder", "Stamp"]];
    const body = sel.map((r) => [
      r.ata,
      String(r.descripcion || "").toUpperCase(),
      fmtModeloAesa(r.modelo),
      fmtMatriculaGuion(r.matricula),
      r.fecha,
      r.workorder,
      r.certificador || "",
    ]);
    /* Filas vacías al final, como en la hoja oficial */
    const blanks = Array.from({ length: 3 }, () => ["", "", "", "", "", "", ""]);
    const ws = XLSX.utils.aoa_to_sheet([...header, ...body, ...blanks]);
    ws["!cols"] = [{ wch: 6 }, { wch: 62 }, { wch: 14 }, { wch: 11 }, { wch: 14 }, { wch: 13 }, { wch: 16 }];
    ws["!rows"] = [{ hpt: 34 }, ...body.map(() => ({ hpt: 26 })), ...blanks.map(() => ({ hpt: 22 }))];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "AESA");
    XLSX.writeFile(wb, `AESA_ExperienciaBasica_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast(`✓ ${sel.length} tareas exportadas en formato AESA oficial`);
  }

  /** Abre un documento HTML en ventana de impresión; fallback: descarga .html */
  function printHtmlDoc(html, filename, okMsg) {
    const w = typeof window !== "undefined" ? window.open("", "_blank") : null;
    if (w) {
      w.document.write(html);
      w.document.close();
      setTimeout(() => w.print(), 500);
      showToast(okMsg);
    } else {
      const blob = new Blob([html], { type: "text/html" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
      showToast("Popup bloqueado: documento descargado como HTML — ábrelo e imprime a PDF");
    }
  }

  /** Réplica imprimible del formato oficial AESA (A4 apaisado, celdas Stamp para sello físico) */
  function exportAesaPrint() {
    const sel = rows
      .filter((r) => selectedIds.includes(r.id) && r.status === "Confirmado")
      .sort((a, b) => (parseFecha(a.fecha)?.getTime() || 0) - (parseFecha(b.fecha)?.getTime() || 0));
    if (!sel.length) {
      showToast("Marca en la tabla las tareas a incluir (casillas de la primera columna)", true);
      return;
    }
    const rowsHtml = sel
      .map(
        (r) =>
          `<tr><td class="c">${r.ata}</td><td>${String(r.descripcion || "").toUpperCase()}</td><td class="c">${fmtModeloAesa(r.modelo)}</td><td class="c mono">${fmtMatriculaGuion(r.matricula)}</td><td class="c">${r.fecha}</td><td class="c mono">${r.workorder}</td><td class="stamp">${r.certificador || ""}</td></tr>`
      )
      .join("");
    const blanks = Array.from({ length: 4 }, () => `<tr><td class="c">&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td class="stamp"></td></tr>`).join("");
    const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Registro de experiencia — formato AESA</title>
<style>
  @page { size: A4 landscape; margin: 12mm; }
  body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; color: #000; margin: 0; }
  .mono { font-family: "Courier New", monospace; }
  .logos { height: 20mm; text-align: right; color: #cbd5e1; font-size: 8pt; padding-top: 4mm; }
  table { width: 100%; border-collapse: collapse; }
  th, td { border: 1px solid #000; padding: 1.6mm 2mm; vertical-align: middle; }
  th { background: #bfbfbf; font-weight: 700; height: 16mm; text-align: center; }
  td { height: 9mm; }
  td.c { text-align: center; }
  td.stamp { width: 11%; font-size: 6.5pt; vertical-align: top; color: #333; }
  col.ata { width: 6%; } col.desc { width: 44%; } col.mod { width: 11%; }
  col.mat { width: 9%; } col.fec { width: 10%; } col.wo { width: 9%; }
</style></head><body>
  <div class="logos">(espacio reservado: escudo · logo AESA de la plantilla oficial)</div>
  <table>
    <colgroup><col class="ata"/><col class="desc"/><col class="mod"/><col class="mat"/><col class="fec"/><col class="wo"/><col/></colgroup>
    <thead><tr><th>ATA</th><th>Descripción Tarea</th><th>Modelo de<br/>aeronave</th><th>Matricula</th><th>Fecha de<br/>realización</th><th>Workorder</th><th>Stamp</th></tr></thead>
    <tbody>${rowsHtml}${blanks}</tbody>
  </table>
</body></html>`;
    printHtmlDoc(html, `AESA_ExperienciaBasica_${new Date().toISOString().slice(0, 10)}.html`, "✓ Hoja AESA abierta: Imprimir → Guardar como PDF");
  }

  /** Logbook personal imprimible (HTML con estilo profesional → Imprimir → PDF) */
  function exportLogbook() {
    const base = selectedIds.length ? rows.filter((r) => selectedIds.includes(r.id)) : rows;
    const src = base
      .filter((r) => r.status === "Confirmado")
      .sort((a, b) => (parseFecha(a.fecha)?.getTime() || 0) - (parseFecha(b.fecha)?.getTime() || 0));
    if (!src.length) return showToast("No hay entradas confirmadas para el logbook", true);
    const byMonth = {};
    const byAta = {};
    const regs = new Set();
    src.forEach((r) => {
      const k = monthKey(r.fecha);
      (byMonth[k] = byMonth[k] || []).push(r);
      byAta[r.ata] = (byAta[r.ata] || 0) + 1;
      regs.add(r.matricula);
    });
    const days = new Set(src.map((r) => r.fecha)).size;
    const monthsHtml = Object.keys(byMonth)
      .sort()
      .map(
        (k) => `
      <section class="month">
        <h2>${monthLabel(k)}</h2>
        <div class="owner">Titular: ${titular.nombre || "____________________"} · Ref.: ${titular.licencia || "__________"} · Firma titular: ____________________</div>
        <table>
          <thead><tr><th>Fecha</th><th>Modelo</th><th>Matrícula</th><th>ATA</th><th>Descripción de la tarea</th><th>Workorder</th><th>SEQ</th><th>Tipo</th><th>Certificador</th><th>Nº Lic.</th><th>Firma / Sello</th></tr></thead>
          <tbody>${byMonth[k]
            .map(
              (r) =>
                `<tr><td>${r.fecha}</td><td class="mono">${r.modelo}</td><td class="mono">${fmtMatriculaGuion(r.matricula)}</td><td class="mono">${String(r.ata).padStart(2, "0")}</td><td>${r.descripcion}</td><td class="mono">${r.workorder}</td><td class="mono">${seqToString(r.seq)}</td><td>LÍNEA</td><td>${r.certificador || ""}</td><td class="firma"></td><td class="firma"></td></tr>`
            )
            .join("")}</tbody>
        </table>
      </section>`
      )
      .join("");
    const ataHtml = Object.entries(byAta)
      .sort((a, b) => Number(a[0]) - Number(b[0]))
      .map(([a, v]) => `<tr><td class="mono">ATA ${String(a).padStart(2, "0")}</td><td>${v}</td></tr>`)
      .join("");
    const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Logbook de Experiencia Básica — Part 66</title>
<style>
  @page { size: A4; margin: 16mm 14mm; }
  body { font-family: Georgia, "Times New Roman", serif; color: #1a2230; margin: 0; }
  .mono { font-family: "Courier New", monospace; }
  .cover { text-align: center; padding-top: 70mm; page-break-after: always; }
  .cover .rule { width: 60mm; height: 3px; background: #b45309; margin: 10mm auto; }
  .cover h1 { font-size: 26pt; letter-spacing: 2px; margin: 0; }
  .cover .sub { font-size: 11pt; color: #64748b; letter-spacing: 4px; text-transform: uppercase; }
  .cover .meta { margin-top: 22mm; font-size: 11pt; line-height: 2; }
  .summary { page-break-after: always; }
  h2 { font-size: 13pt; border-bottom: 2px solid #b45309; padding-bottom: 2mm; letter-spacing: 1px; }
  table { width: 100%; border-collapse: collapse; font-size: 8.2pt; margin-bottom: 6mm; }
  th { background: #0f172a; color: #fff; text-align: left; padding: 1.6mm 2mm; font-size: 7.6pt; text-transform: uppercase; letter-spacing: 0.5px; }
  td { border-bottom: 0.4pt solid #cbd5e1; padding: 1.4mm 2mm; vertical-align: top; }
  tr:nth-child(even) td { background: #f8fafc; }
  .firma { min-width: 22mm; }
  .month { page-break-inside: avoid; }
  .owner { font-size: 8.5pt; color: #475569; margin: -2mm 0 2mm; }
  .kpis { display: flex; gap: 8mm; margin: 8mm 0; }
  .kpi { border: 1px solid #cbd5e1; padding: 4mm 6mm; text-align: center; flex: 1; }
  .kpi b { display: block; font-size: 18pt; }
  footer { font-size: 8pt; color: #94a3b8; margin-top: 10mm; }
</style></head><body>
  <div class="cover">
    <div class="sub">Aircraft Maintenance Technician</div>
    <h1>LOGBOOK DE EXPERIENCIA BÁSICA</h1>
    <div class="rule"></div>
    <div class="sub">EASA Part 66 · Categoría B2</div>
    <div class="meta">
      Titular: ${titular.nombre || "______________________________________"}<br/>
      Licencia / Referencia: ${titular.licencia || "____________________"}<br/>
      Organización Part 145: JC Aircraft Maintenance<br/>
      Periodo: ${src[0].fecha} — ${src[src.length - 1].fecha}<br/>
      Flota: ${Array.from(regs).join(" · ")}
    </div>
  </div>
  <div class="summary">
    <h2>Resumen de experiencia</h2>
    <div class="kpis">
      <div class="kpi"><b>${src.length}</b>tareas certificadas</div>
      <div class="kpi"><b>${days}</b>días documentados</div>
      <div class="kpi"><b>${Object.keys(byAta).length}</b>capítulos ATA</div>
      <div class="kpi"><b>${regs.size}</b>matrículas</div>
    </div>
    <h2>Índice por capítulo ATA</h2>
    <table><thead><tr><th>Capítulo</th><th>Tareas certificadas</th></tr></thead><tbody>${ataHtml}</tbody></table>
  </div>
  ${monthsHtml}
  <footer>Documento generado el ${new Date().toLocaleDateString("es-ES")} · Registro personal de experiencia básica Part 66 · Cada tarea se valida con la firma del certificador</footer>
</body></html>`;
    printHtmlDoc(html, `Logbook_Part66_${new Date().toISOString().slice(0, 10)}.html`, "✓ Logbook abierto: usa Imprimir → Guardar como PDF");
  }

  /* Dashboard de progreso Part 66: solo entradas confirmadas */
  const dash = useMemo(() => {
    const conf = rows.filter((r) => r.status === "Confirmado");
    const byMonth = {}, byAta = {}, monthAtas = {};
    const days = new Set(), regs = new Set();
    conf.forEach((r) => {
      const mk = monthKey(r.fecha);
      byMonth[mk] = (byMonth[mk] || 0) + 1;
      byAta[r.ata] = (byAta[r.ata] || 0) + 1;
      const ma = (monthAtas[mk] = monthAtas[mk] || { count: 0, atas: {} });
      ma.count++;
      ma.atas[r.ata] = (ma.atas[r.ata] || 0) + 1;
      days.add(r.fecha);
      regs.add(r.matricula);
    });
    const monthsArr = Object.entries(byMonth).sort(([a], [b]) => a.localeCompare(b)).slice(-12);
    const atasArr = Object.entries(byAta).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const maxM = Math.max(1, ...monthsArr.map(([, v]) => v));
    const maxA = Math.max(1, ...atasArr.map(([, v]) => v));
    /* Sesgo: meses con >=5 tareas donde un ATA concentra >50% o hay <3 capítulos distintos.
       AESA valora variedad de sistemas, no volumen del mismo tipo de tarea. */
    const biased = Object.entries(monthAtas)
      .filter(([, m]) => m.count >= 5)
      .map(([k, m]) => {
        const top = Object.entries(m.atas).sort((a, b) => b[1] - a[1])[0];
        return { month: k, topAta: top[0], pct: Math.round((top[1] / m.count) * 100), distinct: Object.keys(m.atas).length, count: m.count };
      })
      .filter((b) => b.pct > 50 || b.distinct < 3)
      .sort((a, b) => b.month.localeCompare(a.month));
    return { monthsArr, atasArr, maxM, maxA, biased, days: days.size, regs: regs.size, atasDistinct: Object.keys(byAta).length, total: conf.length };
  }, [rows]);

  const stats = useMemo(() => {
    const conf = rows.filter((r) => r.status === "Confirmado").length;
    const thisMonth = rows.filter((r) => monthKey(r.fecha) === monthKey(fmtFecha(`${new Date().getDate()}/${new Date().getMonth() + 1}/${new Date().getFullYear()}`))).length;
    return { total: rows.length, conf, draft: rows.length - conf, thisMonth };
  }, [rows]);

  /* ---------- Render ---------- */

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      {/* Cabecera tipo formulario técnico */}
      <header className="bg-slate-900 text-white border-b-4 border-amber-500">
        <div className="max-w-6xl mx-auto px-4 py-4 flex flex-wrap items-end justify-between gap-3">
          <div>
            <div className="text-xs font-mono tracking-widest text-amber-400 uppercase">Part 66 · Experiencia Básica · EASA Form LIC-P66</div>
            <h1 className="text-xl font-bold tracking-tight">LMA — OCR Workorder & Logbook</h1>
          </div>
          <div className="flex gap-4 font-mono text-sm">
            <div><span className="text-slate-400">TOTAL </span><span className="font-bold">{stats.total}</span></div>
            <div><span className="text-slate-400">CONF </span><span className="font-bold text-emerald-400">{stats.conf}</span></div>
            <div><span className="text-slate-400">BORR </span><span className="font-bold text-amber-400">{stats.draft}</span></div>
            <div><span className="text-slate-400">MES ACTUAL </span><span className="font-bold">{stats.thisMonth}</span></div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-6 flex flex-col gap-6">
        {/* Zona de carga */}
        <section className="bg-white rounded-lg border border-slate-300 p-4">
          <div className="flex flex-wrap items-center gap-3">
            <input
              ref={fileRef}
              type="file"
              accept="image/*,application/pdf"
              multiple
              className="hidden"
              onChange={(e) => handleFiles(e.target.files)}
            />
            {/* capture="environment" fuerza la cámara trasera en Android/iOS */}
            <input
              ref={cameraRef}
              type="file"
              accept="image/*"
              capture="environment"
              className="hidden"
              onChange={(e) => handleFiles(e.target.files)}
            />
            <button
              onClick={() => cameraRef.current?.click()}
              className="px-4 py-2 bg-amber-500 hover:bg-amber-600 text-slate-900 font-semibold rounded shadow-sm"
            >
              📷 Cámara
            </button>
            <button
              onClick={() => fileRef.current?.click()}
              className="px-4 py-2 bg-white hover:bg-slate-50 border border-slate-400 font-semibold rounded"
            >
              🖼 Galería / PDF
            </button>
            <button
              onClick={addManual}
              className="px-4 py-2 bg-white hover:bg-slate-50 border border-slate-400 font-semibold rounded"
            >
              + Entrada manual
            </button>
            {queue > 0 && (
              <span className="font-mono text-sm text-slate-600 animate-pulse">
                ⏳ Procesando {queue} archivo{queue > 1 ? "s" : ""} con Vision…
              </span>
            )}
            {cola.length > 0 && (
              <button onClick={processQueue} className="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded">
                📥 Procesar cola offline ({cola.length})
              </button>
            )}
            <div className="w-full lg:w-auto lg:ml-auto flex flex-wrap gap-2">
              <button onClick={exportXlsx} className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white font-semibold rounded">
                ⬇ Excel
              </button>
              <button onClick={exportAesa} title="Tareas marcadas → Excel formato AESA" className="px-3 py-2 border border-slate-400 hover:bg-slate-50 rounded text-sm">
                🏛 AESA XLSX
              </button>
              <button onClick={exportAesaPrint} title="Tareas marcadas → réplica imprimible del formato oficial (PDF)" className="px-3 py-2 border border-slate-400 hover:bg-slate-50 rounded text-sm">
                🖨 AESA PDF
              </button>
              <button onClick={exportLogbook} title="Logbook imprimible profesional (selección, o todo si no hay selección)" className="px-3 py-2 border border-slate-400 hover:bg-slate-50 rounded text-sm">
                📖 Logbook
              </button>
              <button onClick={exportJson} className="px-3 py-2 border border-slate-400 hover:bg-slate-50 rounded text-sm">
                Backup JSON
              </button>
            </div>
          </div>
        </section>

        {/* Dashboard de progreso Part 66 */}
        <section className="bg-white rounded-lg border border-slate-300">
          <button onClick={() => setShowDash((s) => !s)} className="w-full px-4 py-3 flex items-center justify-between text-left">
            <span className="text-sm font-bold uppercase tracking-wider text-slate-700">📊 Progreso Part 66</span>
            <span className="font-mono text-xs text-slate-500">
              {dash.biased.length > 0 && <span className="text-red-600 font-bold">⚠ sesgo ATA · </span>}
              {dash.days} días · {dash.atasDistinct} ATAs · {dash.regs} matrículas {showDash ? "▲" : "▼"}
            </span>
          </button>
          {showDash && (
            <div className="px-4 pb-4 border-t border-slate-200 pt-3 grid md:grid-cols-2 gap-6">
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Entradas por mes (últimos 12)</h3>
                {dash.monthsArr.length === 0 && <p className="text-sm text-slate-400">Sin datos confirmados.</p>}
                {dash.monthsArr.map(([m, v]) => (
                  <div key={m} className="flex items-center gap-2 mb-1">
                    <span className="font-mono text-xs w-16 text-slate-600">{monthLabel(m)}</span>
                    <div className="flex-1 bg-slate-100 rounded h-4">
                      <div className="bg-amber-500 h-4 rounded" style={{ width: `${(v / dash.maxM) * 100}%` }} />
                    </div>
                    <span className="font-mono text-xs w-8 text-right font-bold">{v}</span>
                  </div>
                ))}
              </div>
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Cobertura ATA (top 10)</h3>
                {dash.atasArr.map(([a, v]) => (
                  <div key={a} className="flex items-center gap-2 mb-1">
                    <span className="font-mono text-xs w-16 text-slate-600">ATA {String(a).padStart(2, "0")}</span>
                    <div className="flex-1 bg-slate-100 rounded h-4">
                      <div className="bg-slate-700 h-4 rounded" style={{ width: `${(v / dash.maxA) * 100}%` }} />
                    </div>
                    <span className="font-mono text-xs w-8 text-right font-bold">{v}</span>
                  </div>
                ))}
                <p className="text-xs text-slate-500 mt-3">
                  <span className="font-bold">{dash.total}</span> tareas confirmadas en <span className="font-bold">{dash.days}</span> días de trabajo documentados.
                </p>
              </div>
              {dash.biased.length > 0 && (
                <div className="md:col-span-2 border border-red-300 bg-red-50 rounded p-3">
                  <h3 className="text-xs font-bold uppercase tracking-wider text-red-700 mb-1">⚠ Alerta de sesgo mensual</h3>
                  {dash.biased.map((b) => (
                    <p key={b.month} className="text-xs text-red-800">
                      <span className="font-mono font-bold">{monthLabel(b.month)}</span>: el {b.pct}% de las {b.count} tareas son ATA {String(b.topAta).padStart(2, "0")} ({b.distinct} capítulo{b.distinct !== 1 ? "s" : ""} distinto{b.distinct !== 1 ? "s" : ""}). AESA valora variedad de sistemas: busca tareas de otros capítulos este mes.
                    </p>
                  ))}
                </div>
              )}
              <div className="md:col-span-2 flex flex-wrap items-center gap-2 border-t border-slate-200 pt-3">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-500">👤 Titular (portada y páginas del logbook)</span>
                <input
                  value={titular.nombre}
                  onChange={(e) => setTitular({ ...titular, nombre: e.target.value })}
                  placeholder="Nombre completo"
                  className="flex-1 min-w-44 px-3 py-1.5 border border-slate-300 rounded text-sm focus:outline-none focus:border-slate-600"
                />
                <input
                  value={titular.licencia}
                  onChange={(e) => setTitular({ ...titular, licencia: e.target.value })}
                  placeholder="Nº licencia / referencia AESA"
                  className="flex-1 min-w-44 px-3 py-1.5 border border-slate-300 rounded text-sm font-mono focus:outline-none focus:border-slate-600"
                />
              </div>
            </div>
          )}
        </section>

        {/* Sincronización nube (solo si la standalone tiene Supabase configurado) */}
        {SYNC_ENABLED && (
          <section className="bg-white rounded-lg border border-slate-300 p-4 flex flex-wrap items-center gap-3">
            <span className="text-sm font-bold uppercase tracking-wider text-slate-700">☁️ Sincronización</span>
            <input
              value={syncCode}
              onChange={(e) => setSyncCode(e.target.value)}
              placeholder="Código de sincronización (el mismo en móvil y PC)"
              className="flex-1 min-w-52 px-3 py-2 border border-slate-300 rounded text-sm font-mono focus:outline-none focus:border-slate-600"
            />
            <button onClick={syncPush} disabled={syncBusy} className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-white text-sm font-semibold rounded disabled:opacity-50">
              ⬆ Subir
            </button>
            <button onClick={syncPull} disabled={syncBusy} className="px-3 py-2 border border-slate-400 hover:bg-slate-50 text-sm rounded disabled:opacity-50">
              ⬇ Descargar
            </button>
            <span className="font-mono text-xs text-slate-400 w-full">
              Sincroniza entradas y aprendizaje entre dispositivos (las fotos no viajan). Última sync: {lastSync ? new Date(lastSync).toLocaleString("es-ES") : "—"}
            </span>
          </section>
        )}

        {/* Panel de aprendizaje */}
        <section className="bg-white rounded-lg border border-slate-300">
          <button
            onClick={() => setShowLearn((s) => !s)}
            className="w-full px-4 py-3 flex items-center justify-between text-left"
          >
            <span className="text-sm font-bold uppercase tracking-wider text-slate-700">
              🧠 Aprendizaje del extractor
            </span>
            <span className="font-mono text-xs text-slate-500">
              {learning.hints.length} regla{learning.hints.length !== 1 ? "s" : ""} · {learning.examples.length} corrección{learning.examples.length !== 1 ? "es" : ""} {showLearn ? "▲" : "▼"}
            </span>
          </button>
          {showLearn && (
            <div className="px-4 pb-4 flex flex-col gap-3 border-t border-slate-200 pt-3">
              <p className="text-xs text-slate-500">
                Cada corrección que haces antes de confirmar se registra y se inyecta en el prompt de extracción de los siguientes escaneos.
                Puedes añadir reglas manuales (formato de la OT, ubicación de campos, matrículas frecuentes) o sintetizar reglas automáticamente desde el historial.
              </p>
              <div className="flex gap-2">
                <input
                  value={newHint}
                  onChange={(e) => setNewHint(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addHint()}
                  placeholder='Ej: "El WO está bajo la etiqueta W/O No. arriba a la derecha"'
                  className="flex-1 px-3 py-2 border border-slate-300 rounded text-sm focus:outline-none focus:border-slate-600"
                />
                <button onClick={addHint} className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-white text-sm font-semibold rounded">
                  + Regla
                </button>
                <button
                  onClick={synthesizeRules}
                  disabled={synthesizing}
                  className="px-3 py-2 border border-slate-400 hover:bg-slate-50 text-sm rounded disabled:opacity-50"
                >
                  {synthesizing ? "⏳ Sintetizando…" : "⚡ Sintetizar desde correcciones"}
                </button>
              </div>
              {learning.hints.length > 0 && (
                <ul className="flex flex-col gap-1">
                  {learning.hints.map((h) => (
                    <li key={h.id} className="flex items-start gap-2 text-sm bg-slate-50 border border-slate-200 rounded px-3 py-1.5">
                      <span className={`font-mono text-xs px-1.5 py-0.5 rounded mt-0.5 ${h.source === "auto" ? "bg-blue-100 text-blue-800" : "bg-slate-200 text-slate-700"}`}>
                        {h.source === "auto" ? "AUTO" : "MANUAL"}
                      </span>
                      <span className="flex-1">{h.text}</span>
                      <button onClick={() => removeHint(h.id)} className="text-slate-400 hover:text-red-600 text-xs font-bold">✕</button>
                    </li>
                  ))}
                </ul>
              )}
              {learning.examples.length > 0 && (
                <details className="text-xs text-slate-500">
                  <summary className="cursor-pointer font-semibold">Historial de correcciones ({learning.examples.length}, últimas 30)</summary>
                  <ul className="mt-1 font-mono">
                    {learning.examples.slice().reverse().map((e, i) => (
                      <li key={i}>· {e.field}: "{e.extracted}" → "{e.corrected}"</li>
                    ))}
                  </ul>
                </details>
              )}
            </div>
          )}
        </section>

        {/* Pendientes de validación */}
        {pending.length > 0 && (
          <section className="flex flex-col gap-3">
            <h2 className="text-sm font-bold uppercase tracking-wider text-slate-600 flex items-center justify-between">
              <span>Pendientes de validación ({pending.length})</span>
              {pending.length > 1 && (
                <button onClick={confirmAllPending} className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold rounded normal-case tracking-normal">
                  ✓ Confirmar todos los válidos
                </button>
              )}
            </h2>
            {pending.map((p) => {
              const { errors } = validateRow(p);
              const dup = findDuplicate(p, rows);
              return (
                <div key={p.id} className="bg-amber-50 border-l-4 border-amber-500 border border-amber-200 rounded-lg p-4">
                  {dup && (
                    <div className="mb-2 px-3 py-1.5 bg-red-100 border border-red-300 rounded text-xs text-red-800 font-mono">
                      ⚠ POSIBLE DUPLICADO: WO {dup.workorder} ya registrado el {dup.fecha} ({dup.matricula})
                    </div>
                  )}
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                    <div className="font-mono text-xs text-slate-600">
                      📄 {p.fileName}
                      {p.photoData && <span className="ml-2 text-emerald-700">📎 foto adjunta al confirmar</span>}
                      {p.confidence != null && (
                        <span className={`ml-3 px-2 py-0.5 rounded font-bold ${p.confidence >= 0.9 ? "bg-emerald-100 text-emerald-800" : "bg-red-100 text-red-800"}`}>
                          Confianza {(p.confidence * 100).toFixed(0)}%
                        </span>
                      )}
                    </div>
                    <div className="flex gap-2">
                      <button onClick={() => confirmPending(p.id)} className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold rounded">
                        ✓ Confirmar
                      </button>
                      <button onClick={() => discardPending(p.id)} className="px-3 py-1.5 bg-white border border-slate-400 hover:bg-slate-50 text-sm rounded">
                        ✕ Rechazar
                      </button>
                    </div>
                  </div>
                  {p.warnings?.length > 0 && (
                    <div className="mb-3 text-xs text-amber-800 font-mono">⚠ {p.warnings.join(" · ")}</div>
                  )}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                    {FIELD_DEFS.map((def) => (
                      <FieldInput
                        key={def.key}
                        def={def}
                        value={p[def.key]}
                        error={p.showErrors ? errors[def.key] : null}
                        onChange={(v) => updatePending(p.id, def.key, v)}
                      />
                    ))}
                  </div>
                </div>
              );
            })}
          </section>
        )}

        {/* Búsqueda y filtros */}
        <section className="bg-white rounded-lg border border-slate-300 p-4 flex flex-wrap gap-3 items-center">
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="🔍 Buscar en todos los campos (matrícula, WO, descripción…)"
            className="flex-1 min-w-64 px-3 py-2 border border-slate-300 rounded font-mono text-sm focus:outline-none focus:border-slate-600"
          />
          <select value={filterAta} onChange={(e) => setFilterAta(e.target.value)} className="px-2 py-2 border border-slate-300 rounded text-sm font-mono bg-white">
            <option value="">ATA: todos</option>
            {atas.map((a) => <option key={a} value={String(a)}>ATA {a}</option>)}
          </select>
          <select value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)} className="px-2 py-2 border border-slate-300 rounded text-sm font-mono bg-white">
            <option value="">Mes: todos</option>
            {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
          </select>
          <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="px-2 py-2 border border-slate-300 rounded text-sm font-mono bg-white">
            <option value="">Estado: todos</option>
            <option value="Confirmado">Confirmado</option>
            <option value="Borrador">Borrador</option>
          </select>
          <span className="font-mono text-xs text-slate-500">
            {filtered.length} resultado{filtered.length !== 1 ? "s" : ""}
            {selectedIds.length > 0 && (
              <>
                {" · "}
                <span className="text-amber-700 font-bold">{selectedIds.length} seleccionadas para export</span>{" "}
                <button onClick={() => setSelectedIds([])} className="underline">limpiar</button>
              </>
            )}
          </span>
        </section>

        {/* Tabla */}
        <section className="bg-white rounded-lg border border-slate-300 overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-900 text-white text-left">
                <th className="px-3 py-2">
                  <input
                    type="checkbox"
                    checked={filtered.length > 0 && filtered.every((r) => selectedIds.includes(r.id))}
                    onChange={toggleSelectAllFiltered}
                    title="Seleccionar todo lo filtrado (para AESA F02 / Logbook)"
                  />
                </th>
                {["Matrícula", "Fecha", "ATA", "Descripción", "Modelo", "Workorder", "SEQ", "Certificador", "Foto", "Estado", ""].map((h) => (
                  <th key={h} className="px-3 py-2 font-semibold text-xs uppercase tracking-wider whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={12} className="px-4 py-10 text-center text-slate-500">
                    {rows.length === 0
                      ? "Sin entradas. Escanea tu primera orden de trabajo para empezar."
                      : "Ningún resultado con los filtros actuales."}
                  </td>
                </tr>
              )}
              {filtered.map((row) =>
                editingId === row.id ? (
                  <tr key={row.id} className="bg-blue-50 border-b border-slate-200">
                    <td className="px-3 py-1.5"></td>
                    {FIELD_DEFS.map((def) => (
                      <td key={def.key} className="px-2 py-1.5">
                        <input
                          value={editDraft[def.key] ?? ""}
                          onChange={(e) => setEditDraft({ ...editDraft, [def.key]: e.target.value })}
                          className={`w-full px-1.5 py-1 border border-blue-400 rounded text-sm bg-white ${def.mono ? "font-mono" : ""}`}
                        />
                      </td>
                    ))}
                    <td className="px-3 py-1.5">
                      <span className="text-xs text-slate-400">—</span>
                    </td>
                    <td className="px-3 py-1.5">
                      <span className="text-xs text-slate-500">editando…</span>
                    </td>
                    <td className="px-3 py-1.5 whitespace-nowrap">
                      <button onClick={saveEdit} className="text-emerald-700 font-bold mr-2 hover:underline text-xs">GUARDAR</button>
                      <button onClick={() => { setEditingId(null); setEditDraft(null); }} className="text-slate-500 hover:underline text-xs">CANCELAR</button>
                    </td>
                  </tr>
                ) : (
                  <tr key={row.id} className={`border-b border-slate-200 hover:bg-slate-50 ${row.status === "Borrador" ? "opacity-50" : ""}`}>
                    <td className="px-3 py-2">
                      <input type="checkbox" checked={selectedIds.includes(row.id)} onChange={() => toggleSelect(row.id)} />
                    </td>
                    <td className="px-3 py-2 font-mono font-bold">{row.matricula}</td>
                    <td className="px-3 py-2 font-mono whitespace-nowrap">{row.fecha}</td>
                    <td className="px-3 py-2">
                      <span className="inline-block px-1.5 py-0.5 bg-slate-900 text-amber-400 font-mono font-bold rounded text-xs">{String(row.ata).padStart(2, "0")}</span>
                    </td>
                    <td className="px-3 py-2 max-w-md truncate" title={row.descripcion}>{row.descripcion}</td>
                    <td className="px-3 py-2 font-mono">{row.modelo}</td>
                    <td className="px-3 py-2 font-mono">{row.workorder}</td>
                    <td className="px-3 py-2 font-mono text-xs">{seqToString(row.seq)}</td>
                    <td className="px-3 py-2 text-xs">{row.certificador || <span className="text-slate-400">—</span>}</td>
                    <td className="px-3 py-2 whitespace-nowrap">
                      {row.hasPhoto ? (
                        <button onClick={() => viewPhoto(row)} className="text-blue-700 hover:underline text-xs font-semibold">📷 VER</button>
                      ) : (
                        <button onClick={() => requestAttach(row.id)} className="text-slate-400 hover:text-slate-700 text-xs">📎 Adjuntar</button>
                      )}
                    </td>
                    <td className="px-3 py-2">
                      <span className={`text-xs font-semibold px-2 py-0.5 rounded ${row.status === "Confirmado" ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"}`}>
                        {row.status}
                      </span>
                    </td>
                    <td className="px-3 py-2 whitespace-nowrap">
                      <button onClick={() => startEdit(row)} className="text-blue-700 hover:underline text-xs font-semibold mr-2">EDITAR</button>
                      <button onClick={() => toggleStatus(row.id)} className="text-slate-500 hover:underline text-xs">
                        {row.status === "Confirmado" ? "→ BORRADOR" : "→ CONFIRMAR"}
                      </button>
                    </td>
                  </tr>
                )
              )}
            </tbody>
          </table>
        </section>

        <footer className="text-xs font-mono text-slate-400 text-center pb-6">
          Soft delete vía Borrador · Guardado automático + backup redundante · Cola offline con proceso automático al recuperar red · Selección con casillas → AESA F02 / Logbook · Export XLSX: 1 hoja/mes, solo confirmadas
        </footer>
      </main>

      {/* Input oculto para adjuntar foto a una entrada existente */}
      <input ref={attachRef} type="file" accept="image/*" className="hidden" onChange={handleAttachFile} />

      {/* Modal de foto */}
      {photoModal && (
        <div
          className="fixed inset-0 bg-black bg-opacity-70 flex items-center justify-center p-4 z-50"
          onClick={() => setPhotoModal(null)}
        >
          <div className="bg-white rounded-lg max-w-3xl w-full max-h-full overflow-auto" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-center justify-between px-4 py-2 border-b border-slate-200">
              <span className="font-mono text-sm font-bold">{photoModal.label || "Foto de la OT"}</span>
              <div className="flex gap-3">
                {photoModal.rowId && (
                  <button
                    onClick={() => removePhoto(filtered.find((r) => r.id === photoModal.rowId) || { id: photoModal.rowId })}
                    className="text-red-600 hover:underline text-xs font-semibold"
                  >
                    ELIMINAR FOTO
                  </button>
                )}
                <button onClick={() => setPhotoModal(null)} className="text-slate-500 hover:text-slate-900 font-bold">✕</button>
              </div>
            </div>
            <div className="p-3">
              {photoModal.loading ? (
                <div className="py-16 text-center font-mono text-sm text-slate-500 animate-pulse">Cargando foto…</div>
              ) : (
                <img src={photoModal.url} alt="Orden de trabajo" className="w-full h-auto rounded" />
              )}
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div className={`fixed bottom-4 right-4 px-4 py-3 rounded-lg shadow-lg font-mono text-sm text-white ${toast.isError ? "bg-red-600" : "bg-slate-900"}`}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
