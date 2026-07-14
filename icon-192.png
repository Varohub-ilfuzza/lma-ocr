import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";

/* Variables de entorno de Vercel inyectadas al ámbito global (el código
   compartido con el artifact no puede usar import.meta directamente). */
globalThis.LMA_SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "";
globalThis.LMA_SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || "";

/* Almacenamiento persistente: en iOS/Safari evita que el sistema purgue
   IndexedDB por inactividad. Best-effort: si el navegador no lo soporta, no falla. */
if (navigator.storage?.persist) {
  navigator.storage.persist().catch(() => {});
}

/* Service worker: la app carga offline (los escaneos offline van a la cola) */
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js").catch(() => {});
  });
}

createRoot(document.getElementById("root")).render(<App />);
