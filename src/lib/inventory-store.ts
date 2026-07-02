import { useSyncExternalStore } from "react";

export type InventoryRow = Record<string, string | number> & { __id: string };

export interface SellerInventory {
  columns: string[];
  rows: InventoryRow[];
  uploadedAt: string | null;
  sales: Record<string, number>; // __id -> units sold
}

export interface InventoryState {
  sellers: Record<string, SellerInventory>; // sellerEmail -> SellerInventory
}

const KEY = "inv_state_v2"; // Increment version for safe migrations

const initialSellerInventory: SellerInventory = {
  columns: [],
  rows: [],
  uploadedAt: null,
  sales: {},
};

const initial: InventoryState = {
  sellers: {},
};

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

export function setInventory(sellerEmail: string, columns: string[], rows: InventoryRow[]) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  sellers[email] = {
    columns,
    rows,
    uploadedAt: new Date().toISOString(),
    sales: {},
  };
  state = { ...state, sellers };
  persist();
}

export function setSale(sellerEmail: string, id: string, units: number) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  const sellerInv = sellers[email] || { ...initialSellerInventory };
  const sales = { ...sellerInv.sales };
  
  if (!units || units <= 0) {
    delete sales[id];
  } else {
    sales[id] = units;
  }

  sellers[email] = {
    ...sellerInv,
    sales,
  };
  state = { ...state, sellers };
  persist();
}

export function resetSales(sellerEmail: string) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  if (sellers[email]) {
    sellers[email] = {
      ...sellers[email],
      sales: {},
    };
    state = { ...state, sellers };
    persist();
  }
}

export function clearSellerInventory(sellerEmail: string) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  delete sellers[email];
  state = { ...state, sellers };
  persist();
}

export function clearAll() {
  state = initial;
  persist();
}

export function useInventory(sellerEmail: string): SellerInventory {
  const globalState = useSyncExternalStore(
    (cb) => {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    () => state,
    () => initial,
  );

  return globalState.sellers[sellerEmail.toLowerCase()] || initialSellerInventory;
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
    ref: find(["referencia", "ref", "modelo"]),
    color: find(["color", "colour"]),
    size: find(["talla", "size", "tamaño", "tamano"]),
  };
}