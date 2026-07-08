/**
 * Adaptador de almacenamiento: misma interfaz que window.storage de los
 * artifacts de Claude, implementada sobre IndexedDB (idb-keyval).
 * Persistencia local en el dispositivo, sin límites de rate ni de red.
 */
import { get, set, del } from "idb-keyval";

export const storage = {
  async get(key) {
    const value = await get(key);
    if (value === undefined) throw new Error(`clave "${key}" no encontrada`);
    return { key, value };
  },
  async set(key, value) {
    await set(key, value);
    return { key, value };
  },
  async delete(key) {
    await del(key);
    return { key, deleted: true };
  },
};
