import * as XLSX from "xlsx";
import { getColorCode, translateColor } from "./color-mapping";

export interface ExcelSalesItem {
  ref: string;
  qty: number;
  size: string;
  color: string;
}

export interface ExcelCommissionsItem {
  ref: string;
  name: string;
  size: string;
  color: string;
  qty: number;
  price: number;
}

export function generateSalesExcel(
  items: ExcelSalesItem[],
  warehouseId: string,
  sellerName: string
) {
  const header = [
    "StrProducto",
    "IntCantidaddoc",
    "IntvalorUnitario",
    "IntBodega",
    "StrLote",
    "StrColor",
  ];
  const data: (string | number)[][] = [header];
  
  for (const item of items) {
    const colorCode = getColorCode(item.color);
    data.push([
      item.ref,
      item.qty,
      0,
      warehouseId,
      item.size,
      colorCode,
    ]);
  }
  
  const ws = XLSX.utils.aoa_to_sheet(data);
  const range = XLSX.utils.decode_range(ws["!ref"]!);
  const textCols = [0, 3, 4, 5];
  for (let R = 1; R <= range.e.r; R++) {
    for (const C of textCols) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (cell) {
        cell.t = "s";
        cell.v = String(cell.v);
      }
    }
  }
  
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "TblDetalleDocumentos");
  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `REPORTE_VENTAS_${sellerName.toUpperCase()}_${stamp}.xls`, { bookType: "biff8" });
}

export function generateCommissionsExcel(
  items: ExcelCommissionsItem[],
  commPercent: number,
  sellerName: string
) {
  const header = [
    "Referencia",
    "StrProducto",
    "Descripción",
    "StrLote",
    "StrColor",
    "Cantidad",
    "Precio Unitario",
    "VENTA CON IVA",
    "VENTA SIN IVA",
    `COMISION ${commPercent}%`,
    "IVA COMISION (19%)",
    "RETEFUENTE (11%)",
    "TOTAL COMISION",
    "TOTAL A REEMBOLSAR",
  ];
  const data: (string | number)[][] = [header];

  for (const item of items) {
    const rawVal = item.qty * item.price;
    const valWithIva = Math.round(rawVal);
    const valWithoutIva = Math.round(rawVal / 1.19);
    const commission = Math.round(valWithoutIva * (commPercent / 100));
    const ivaComm = Math.round(commission * 0.19);
    const retSource = Math.round(commission * 0.11);
    const totalComm = commission + ivaComm - retSource;
    const totalReimburse = valWithIva - totalComm;

    const colorDesc = translateColor(item.color);
    data.push([
      item.ref || "",
      item.ref || "",
      item.name || item.ref || "Producto",
      item.size || "",
      colorDesc,
      item.qty,
      item.price,
      valWithIva,
      valWithoutIva,
      commission,
      ivaComm,
      retSource,
      totalComm,
      totalReimburse,
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const range = XLSX.utils.decode_range(ws["!ref"]!);
  const textCols = [0, 1, 3, 4];
  for (let R = 1; R <= range.e.r; R++) {
    for (const C of textCols) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (cell) {
        cell.t = "s";
        cell.v = String(cell.v);
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Comisiones");
  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `REPORTE_COMISIONES_${sellerName.toUpperCase()}_${stamp}.xls`, { bookType: "biff8" });
}
