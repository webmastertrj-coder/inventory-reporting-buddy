import { useSyncExternalStore } from "react";

export type InventoryRow = Record<string, string | number> & { __id: string };

export interface InventoryState {
  columns: string[];
  rows: InventoryRow[];
  uploadedAt: string | null;
  sales: Record<string, number>; // __id -> units sold
}

const KEY = "inv_state_v1";

const initial: InventoryState = { columns: [], rows: [], uploadedAt: null, sales: {} };

let state: InventoryState = load();
const listeners = new Set<() => void>();

function load(): InventoryState {
  if (typeof window === "undefined") return initial;
  try {
    const raw = window.localStorage.getItem(KEY);
    if (!raw) return initial;
    return { ...initial, ...JSON.parse(raw) };
  } catch {
    return initial;
  }
}

function persist() {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(KEY, JSON.stringify(state));
  listeners.forEach((l) => l());
}

export function setInventory(columns: string[], rows: InventoryRow[]) {
  state = { columns, rows, uploadedAt: new Date().toISOString(), sales: {} };
  persist();
}

export function setSale(id: string, units: number) {
  const sales = { ...state.sales };
  if (!units || units <= 0) delete sales[id];
  else sales[id] = units;
  state = { ...state, sales };
  persist();
}

export function resetSales() {
  state = { ...state, sales: {} };
  persist();
}

export function clearAll() {
  state = initial;
  persist();
}

export function useInventory(): InventoryState {
  return useSyncExternalStore(
    (cb) => {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    () => state,
    () => initial,
  );
}

/** Try to find a sensible "name" and "sku" column heuristically */
export function detectKeyColumns(columns: string[]) {
  const lower = columns.map((c) => c.toLowerCase());
  const find = (needles: string[]) => {
    const idx = lower.findIndex((c) => needles.some((n) => c.includes(n)));
    return idx >= 0 ? columns[idx] : null;
  };
  return {
    sku: find(["sku", "codigo", "código", "code", "ref"]),
    name: find(["nombre", "producto", "name", "descrip"]),
    stock: find(["stock", "existencia", "cantidad", "inventario"]),
  };
}