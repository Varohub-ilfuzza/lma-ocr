# LMA OCR — App Standalone (Vercel)

Registro de experiencia básica Part 66 con OCR de órdenes de trabajo.
Migración directa del artifact de Claude a app de producción:
cámara nativa en Android, persistencia local IndexedDB, API key protegida en servidor.

## Arquitectura

```
Navegador (React + Vite, PWA)
  ├─ IndexedDB (entradas + backup + fotos, local al dispositivo)
  └─ POST /api/extract ──► Función serverless Vercel ──► API Anthropic
                            (ANTHROPIC_API_KEY solo aquí)
```

## Despliegue paso a paso (~20 min)

### 1. API key de Anthropic
1. Crear cuenta en https://console.anthropic.com
2. Billing → cargar crédito prepago (mínimo 5 USD, dura meses con este uso)
3. API Keys → Create Key → copiar (empieza por `sk-ant-`). No compartirla nunca.

### 2. Subir a GitHub
1. Crear repositorio nuevo (privado) en https://github.com/new — nombre: `lma-ocr`
2. En esta carpeta:
```bash
git init && git add . && git commit -m "LMA OCR v1"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/lma-ocr.git
git push -u origin main
```

### 3. Desplegar en Vercel (gratis)
1. https://vercel.com → Sign up con GitHub
2. Add New → Project → importar `lma-ocr` (detecta Vite automáticamente)
3. Antes de Deploy: **Environment Variables** →
   - Name: `ANTHROPIC_API_KEY` · Value: tu clave `sk-ant-...`
4. Deploy → URL tipo `https://lma-ocr.vercel.app`

### 4. Instalar en el móvil (Android)
1. Abrir la URL en **Chrome**
2. Menú ⋮ → **Añadir a pantalla de inicio** (PWA)
3. El botón 📷 Cámara abre la cámara trasera de forma nativa.

## Desarrollo local
```bash
npm install
npx vercel dev        # sirve frontend + /api/extract juntos
```
Nota: `npm run dev` solo levanta el frontend; el endpoint /api requiere `vercel dev`
(exporta antes `ANTHROPIC_API_KEY` o usa `vercel env pull`).

## Costes (verificados jul 2026)
- Vercel Hobby: 0 €
- API Anthropic: pago por uso. Claude Sonnet 4.6: 3 $/M tokens entrada, 15 $/M salida.
  Un escaneo ≈ 2.000 tokens entrada + 200 salida ≈ 0,009 $ → **~0,70 $/mes con 75 OTs**.
- Alternativa low-cost: variable `CLAUDE_MODEL=claude-haiku-4-5-20251001` (1 $/5 $ por M) ≈ 0,25 $/mes;
  probar precisión antes de fijarlo.
- Fuente oficial: https://docs.claude.com/en/api/overview y claude.com/pricing

## Datos y respaldo
- Las entradas y fotos viven en IndexedDB **del dispositivo** (offline tras la primera carga).
- Móvil y PC no comparten datos entre sí (sin backend de datos por diseño).
- Disciplina: exportar Backup JSON semanal. Fase 2 posible: sincronización vía backend (Supabase/Postgres).

## Límites conocidos
- PDFs > ~4 MB fallan en /api/extract (límite de body serverless). Las fotos van comprimidas y no tienen este problema.
- Si Vercel devuelve 500 "ANTHROPIC_API_KEY no configurada": revisar Environment Variables y redeploy.

## Modo offline (hangar)
- La PWA carga sin conexión (service worker network-first: online siempre sirve la última versión).
- Escaneos sin red → **cola offline** (foto comprimida en IndexedDB). Al recuperar conexión se procesan
  automáticamente; también hay botón "📥 Procesar cola".
- Los PDF requieren conexión (no se encolan).

## Export AESA F02 y Logbook
- Marca con las casillas de la tabla las 5–12 tareas del mes → botón **🏛 AESA F02** (XLSX provisional
  basado en LIC-P66-P01-F02; adjuntar el formulario oficial para calibrar el formato exacto).
- **📖 Logbook**: documento imprimible profesional (portada, KPIs, índice ATA, tablas mensuales con
  columna de firma) → Imprimir → Guardar como PDF. Usa la selección; sin selección, todo lo confirmado.

## Sincronización móvil ↔ PC (opcional, Supabase gratuito)
1. Crear proyecto en https://supabase.com (plan free).
2. SQL Editor → ejecutar:
```sql
create table lma_sync (id text primary key, data jsonb, updated_at timestamptz);
alter table lma_sync enable row level security;
create policy "lma_anon" on lma_sync for all using (true) with check (true);
```
3. Project Settings → API → copiar URL y anon key.
4. En Vercel → Environment Variables:
   - `VITE_SUPABASE_URL` = https://xxxx.supabase.co
   - `VITE_SUPABASE_ANON_KEY` = eyJ...
5. Redeploy. Aparece la sección "☁️ Sincronización": define un **código** (actúa como contraseña,
   usa uno largo y no trivial) y usa ⬆ Subir / ⬇ Descargar en cada dispositivo.

**Seguridad, sin rodeos:** este esquema es de nivel personal (cualquiera con tu anon key + código
podría leer los datos). No hay datos de terceros ni fotos en la nube. Para nivel superior: RLS con
auth por email (fase futura).

**Límite:** last-write-wins — sube desde el dispositivo con los datos buenos y descarga en el otro;
no hay merge automático.
