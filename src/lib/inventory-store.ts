import { useSyncExternalStore } from "react";
import { supabase, isCloudEnabled } from "./supabase-client";

export type InventoryRow = Record<string, string | number> & { __id: string };

export interface SellerInventory {
  columns: string[];
  rows: InventoryRow[];
  uploadedAt: string | null;
  sales: Record<string, number>; // __id -> units sold
}

export interface ExportLogItem {
  __id: string;
  ref: string;
  name: string;
  variantDesc: string;
  qty: number;
  price: number;
  stock: number | null;
  size: string;
  color: string;
}

export interface ExportLog {
  id: string;
  sellerEmail: string;
  timestamp: string;
  totalUnits: number;
  totalAmount: number;
  commissionPercent: number;
  warehouseId: string;
  items: ExportLogItem[];
}

export interface InventoryState {
  sellers: Record<string, SellerInventory>; // sellerEmail -> SellerInventory
  exportLogs?: ExportLog[];
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
  exportLogs: [],
};

function migrateState(loaded: any): InventoryState {
  if (!loaded || !loaded.sellers) return loaded;
  
  let changed = false;
  const sellers = { ...loaded.sellers };

  for (const email of Object.keys(sellers)) {
    const inv = sellers[email];
    if (!inv || !inv.rows || !inv.rows.length) continue;

    // Check if any row has an invalid ID or if there are duplicate IDs in the list
    const ids = inv.rows.map((r: any) => r.__id);
    const hasDuplicates = ids.length !== new Set(ids).size;
    const hasBadId = inv.rows.some((r: any) => !r.__id || r.__id.startsWith("row__"));

    if (hasBadId || hasDuplicates) {
      const tempKeys = detectKeyColumns(inv.columns);
      const mergedRowsMap: Record<string, any> = {};

      inv.rows.forEach((r: any, i: number) => {
        const refPart = tempKeys.ref ? String(r[tempKeys.ref] ?? "").trim() : "";
        const colorPart = tempKeys.color ? String(r[tempKeys.color] ?? "").trim() : "";
        const sizePart = tempKeys.size ? String(r[tempKeys.size] ?? "").trim() : "";
        const stableId = (refPart || colorPart || sizePart)
          ? `row_${refPart}_${colorPart}_${sizePart}`.replace(/\s+/g, "_")
          : `row_${i}`;

        if (mergedRowsMap[stableId]) {
          if (tempKeys.stock) {
            const existingStock = Number(mergedRowsMap[stableId][tempKeys.stock]) || 0;
            const newStock = Number(r[tempKeys.stock]) || 0;
            mergedRowsMap[stableId][tempKeys.stock] = existingStock + newStock;
          }
        } else {
          mergedRowsMap[stableId] = {
            ...r,
            __id: stableId,
          };
        }
      });

      sellers[email] = {
        ...inv,
        rows: Object.values(mergedRowsMap),
        sales: {}, // Reset sales as they are corrupted/colliding anyway due to clashing
      };
      changed = true;
    }
  }

  if (changed) {
    return { ...loaded, sellers };
  }
  return loaded;
}

let state: InventoryState = load();
const listeners = new Set<() => void>();

function load(): InventoryState {
  if (typeof window === "undefined") return initial;
  try {
    const raw = window.localStorage.getItem(KEY);
    if (!raw) return initial;
    const parsed = JSON.parse(raw);
    const migrated = migrateState(parsed);
    
    // Persist migrated state if changes were made
    if (JSON.stringify(migrated) !== JSON.stringify(parsed)) {
      window.localStorage.setItem(KEY, JSON.stringify(migrated));
    }
    return migrated;
  } catch {
    return initial;
  }
}

function persist() {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(KEY, JSON.stringify(state));
  listeners.forEach((l) => l());
}

export async function syncFromCloud() {
  if (!isCloudEnabled || !supabase) return;

  try {
    // 1. Fetch all sellers
    const { data: dbSellers, error: sellersErr } = await supabase
      .from("sellers")
      .select("*");
      
    if (sellersErr) throw sellersErr;

    // 2. Fetch all inventories
    const { data: dbInventories, error: invsErr } = await supabase
      .from("inventories")
      .select("*");
      
    if (invsErr) throw invsErr;

    // 3. Fetch all sales
    const { data: dbSales, error: salesErr } = await supabase
      .from("sales")
      .select("*");
      
    if (salesErr) throw salesErr;

    // 4. Fetch all export logs
    const { data: dbLogs, error: logsErr } = await supabase
      .from("export_logs")
      .select("*");
      
    if (logsErr) throw logsErr;

    // Reconstruct state
    const sellersMap: Record<string, SellerInventory> = {};

    const salesBySeller: Record<string, Record<string, number>> = {};
    if (dbSales) {
      dbSales.forEach((sale) => {
        const email = sale.seller_email.toLowerCase();
        if (!salesBySeller[email]) salesBySeller[email] = {};
        salesBySeller[email][sale.product_id] = sale.qty;
      });
    }

    // Gather all unique seller emails from sellers, inventories, and sales tables
    const allEmails = new Set<string>();
    if (dbSellers) dbSellers.forEach((s) => allEmails.add(s.email.toLowerCase()));
    if (dbInventories) dbInventories.forEach((i) => allEmails.add(i.seller_email.toLowerCase()));
    if (dbSales) dbSales.forEach((sa) => allEmails.add(sa.seller_email.toLowerCase()));

    allEmails.forEach((email) => {
      const invData = dbInventories?.find((i) => i.seller_email.toLowerCase() === email);
      
      sellersMap[email] = {
        columns: invData?.columns || [],
        rows: invData?.rows || [],
        uploadedAt: invData?.uploaded_at || null,
        sales: salesBySeller[email] || {},
      };
    });

    const exportLogs = dbLogs ? dbLogs.map((log) => ({
      id: log.id,
      sellerEmail: log.seller_email.toLowerCase(),
      timestamp: log.timestamp,
      totalUnits: Number(log.total_units),
      totalAmount: Number(log.total_amount),
      commissionPercent: Number(log.commission_percent),
      warehouseId: log.warehouse_id,
      items: log.items || [],
    })) : [];

    state = {
      sellers: sellersMap,
      exportLogs,
    };
    
    persist();
  } catch (err) {
    console.error("Failed to sync state from cloud:", err);
  }
}

// Initial fetch and polling configuration
if (isCloudEnabled) {
  syncFromCloud();
  if (typeof window !== "undefined") {
    // Poll every 10 seconds
    setInterval(syncFromCloud, 10000);
    // Sync on tab focus
    window.addEventListener("focus", () => {
      syncFromCloud();
    });
  }
}

export function setInventory(sellerEmail: string, columns: string[], rows: InventoryRow[]) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  const existingInv = sellers[email];
  sellers[email] = {
    columns,
    rows,
    uploadedAt: new Date().toISOString(),
    sales: existingInv ? existingInv.sales : {},
  };
  state = { ...state, sellers };
  persist();

  if (isCloudEnabled && supabase) {
    supabase
      .from("inventories")
      .upsert({
        seller_email: email,
        columns,
        rows,
        uploaded_at: new Date().toISOString(),
      })
      .then(({ error }) => {
        if (error) console.error("Error saving inventory to cloud:", error);
      });
  }
}

export function setSale(sellerEmail: string, id: string, units: number) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  const sellerInv = sellers[email] || { ...initialSellerInventory };
  const sales = { ...(sellerInv.sales || {}) };
  
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

  if (isCloudEnabled && supabase) {
    if (!units || units <= 0) {
      supabase
        .from("sales")
        .delete()
        .eq("seller_email", email)
        .eq("product_id", id)
        .then(({ error }) => {
          if (error) console.error("Error deleting sale from cloud:", error);
        });
    } else {
      supabase
        .from("sales")
        .upsert({
          seller_email: email,
          product_id: id,
          qty: units,
        })
        .then(({ error }) => {
          if (error) console.error("Error saving sale to cloud:", error);
        });
    }
  }
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

    if (isCloudEnabled && supabase) {
      supabase
        .from("sales")
        .delete()
        .eq("seller_email", email)
        .then(({ error }) => {
          if (error) console.error("Error resetting sales in cloud:", error);
        });
    }
  }
}

export function clearSellerInventory(sellerEmail: string) {
  const email = sellerEmail.toLowerCase();
  const sellers = { ...state.sellers };
  delete sellers[email];
  state = { ...state, sellers };
  persist();

  if (isCloudEnabled && supabase) {
    supabase
      .from("inventories")
      .delete()
      .eq("seller_email", email)
      .then(({ error }) => {
        if (error) console.error("Error deleting inventory from cloud:", error);
      });
  }
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

  const inv = globalState.sellers[sellerEmail.toLowerCase()];
  if (!inv) return initialSellerInventory;

  return {
    columns: inv.columns || [],
    rows: inv.rows || [],
    uploadedAt: inv.uploadedAt || null,
    sales: inv.sales || {},
  };
}

export function useExportLogs(): ExportLog[] {
  const globalState = useSyncExternalStore(
    (cb) => {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    () => state,
    () => initial,
  );

  return globalState.exportLogs || [];
}

/** Try to find a sensible "name" and "sku" column heuristically */
export function detectKeyColumns(columns: string[]) {
  const trimmedCols = columns.map((c) => c.trim());
  const lower = trimmedCols.map((c) => c.toLowerCase());
  
  const find = (needles: string[], exclude: string | null = null) => {
    const idx = lower.findIndex((c, i) => 
      trimmedCols[i] !== exclude && needles.some((n) => c.includes(n))
    );
    return idx >= 0 ? trimmedCols[idx] : null;
  };

  // Find refCol by searching first for strproducto, referencia, modelo
  let refCol = find(["strproducto", "referencia", "modelo"]);
  if (!refCol) {
    // Robust fallbacks for other standard reference column names
    refCol = find(["ref", "sku", "codigo", "código", "code", "producto"]);
  }

  // Prioritize "pvp" for price, otherwise search standard price keys
  const pvpIdx = lower.findIndex((c) => c === "pvp" || c.includes("pvp"));
  const priceCol = pvpIdx >= 0
    ? trimmedCols[pvpIdx]
    : find(["precio", "valor", "price", "costo", "cost", "unitario"]);

  return {
    sku: find(["sku", "codigo", "código", "code", "ref"], refCol),
    name: find(["nombre", "producto", "name", "descrip"], refCol),
    stock: find(["stock", "existencia", "cantidad", "inventario"]),
    ref: refCol,
    color: find(["color", "colour"]),
    size: find(["talla", "size", "tamaño", "tamano", "lote", "strlote"]),
    price: priceCol,
  };
}

export function parsePrice(val: any): number {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === "number") return val;
  
  let str = String(val).trim();
  // Remove currency signs, spaces
  str = str.replace(/[\$\s]/g, "");
  
  // Match 1 or 2 digits at the end preceded by dot or comma (decimals)
  const matchDecimal = str.match(/[\.,](\d{1,2})$/);
  if (matchDecimal) {
    const decimals = matchDecimal[1];
    // Remove all dots/commas from the main part
    const mainPart = str.substring(0, matchDecimal.index!).replace(/[\.,]/g, "");
    str = mainPart + "." + decimals;
  } else {
    // Remove all dots/commas
    str = str.replace(/[\.,]/g, "");
  }
  
  const num = Number(str);
  return isNaN(num) ? 0 : num;
}

export function addExportLog(
  sellerEmail: string,
  warehouseId: string,
  commissionPercent: number,
  items: ExportLogItem[]
) {
  const email = sellerEmail.toLowerCase();
  const logs = state.exportLogs || [];
  
  const totalUnits = items.reduce((sum, item) => sum + item.qty, 0);
  const totalAmount = items.reduce((sum, item) => sum + item.qty * item.price, 0);

  const logId = `log_${Date.now()}_${Math.random().toString(36).substring(2, 11)}`;
  const timestampStr = new Date().toISOString();

  const newLog: ExportLog = {
    id: logId,
    sellerEmail: email,
    timestamp: timestampStr,
    totalUnits,
    totalAmount,
    commissionPercent,
    warehouseId,
    items,
  };

  state = {
    ...state,
    exportLogs: [newLog, ...logs],
  };
  persist();

  if (isCloudEnabled && supabase) {
    supabase
      .from("export_logs")
      .insert({
        id: logId,
        seller_email: email,
        timestamp: timestampStr,
        total_units: totalUnits,
        total_amount: totalAmount,
        commission_percent: commissionPercent,
        warehouse_id: warehouseId,
        items,
      })
      .then(({ error }) => {
        if (error) console.error("Error saving export log to cloud:", error);
      });
  }
}

export function deleteExportLog(logId: string) {
  const logs = state.exportLogs || [];
  state = {
    ...state,
    exportLogs: logs.filter((log) => log.id !== logId),
  };
  persist();

  if (isCloudEnabled && supabase) {
    supabase
      .from("export_logs")
      .delete()
      .eq("id", logId)
      .then(({ error }) => {
        if (error) console.error("Error deleting export log from cloud:", error);
      });
  }
}