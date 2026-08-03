import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useRef, useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import { LogOut, Loader2, Plus, User, FileSpreadsheet, Trash2, Download, RefreshCw, FileText, CheckCircle2, AlertCircle, Search, ShoppingBag } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { setInventory, useInventory, clearSellerInventory, resetSales, detectKeyColumns, type InventoryRow, parsePrice, useExportLogs, deleteExportLog, setSale, addExportLog } from "@/lib/inventory-store";
import { useAuth } from "@/lib/auth-store";
import { translateColor, getColorCode } from "@/lib/color-mapping";
import { generateSalesExcel, generateCommissionsExcel } from "@/lib/excel-helpers";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Inventario | Panel del Administrador" },
      {
        name: "description",
        content:
          "Gestiona los vendedores autorizados, sube archivos de inventario Excel para cada uno y consolida sus reportes de ventas.",
      },
    ],
  }),
  component: Index,
});

function Index() {
  const { user, loading: authLoading, logout, sellers, addSeller, updateSellerCommission, updateSellerWarehouseId } = useAuth();
  const navigate = useNavigate();
  const inputRef = useRef<HTMLInputElement>(null);
  const [fileLoading, setFileLoading] = useState(false);
  const [selectedSellerEmail, setSelectedSellerEmail] = useState<string>("");
  const [newSellerName, setNewSellerName] = useState("");
  const [newSellerEmail, setNewSellerEmail] = useState("");
  const [newSellerCommission, setNewSellerCommission] = useState("10");
  const [newSellerWarehouseId, setNewSellerWarehouseId] = useState("01");
  const [showAddForm, setShowAddForm] = useState(false);

  // States for editing commission
  const [isEditingCommission, setIsEditingCommission] = useState(false);
  const [editCommissionValue, setEditCommissionValue] = useState("10");

  // States for editing warehouse
  const [isEditingWarehouse, setIsEditingWarehouse] = useState(false);
  const [editWarehouseValue, setEditWarehouseValue] = useState("01");

  // State for filtering loaded products preview table
  const [tableFilter, setTableFilter] = useState("");

  // Hook to get the inventory of the selected seller
  const selectedInv = useInventory(selectedSellerEmail);
  const allLogs = useExportLogs();

  const handleDownloadSalesFromLog = (log: any) => {
    const exportItems = log.items.map((item: any) => ({
      ref: item.ref,
      qty: item.qty,
      size: item.size,
      color: item.color,
    }));
    generateSalesExcel(exportItems, log.warehouseId, log.sellerEmail.split("@")[0]);
    toast.success("Reporte de ventas descargado del historial.");
  };

  const handleDownloadCommissionsFromLog = (log: any) => {
    const exportItems = log.items.map((item: any) => ({
      ref: item.ref,
      name: item.name,
      size: item.size,
      color: item.color,
      qty: item.qty,
      price: item.price,
    }));
    generateCommissionsExcel(exportItems, log.commissionPercent, log.sellerEmail.split("@")[0]);
    toast.success("Liquidación de comisiones descargada del historial.");
  };

  const handleDeleteLog = (logId: string) => {
    deleteExportLog(logId);
    toast.success("Registro de exportación eliminado.");
  };

  // Guard routing
  useEffect(() => {
    if (!authLoading) {
      if (!user) {
        navigate({ to: "/login" });
      } else if (user.role !== "admin") {
        navigate({ to: "/proveedor" });
      }
    }
  }, [user, authLoading, navigate]);

  // Set the first seller as selected by default if available
  useEffect(() => {
    if (sellers.length > 0 && !selectedSellerEmail) {
      setSelectedSellerEmail(sellers[0].email);
    }
  }, [sellers, selectedSellerEmail]);

  const handleAddSeller = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newSellerName || !newSellerEmail) {
      toast.error("Por favor completa el nombre y el correo.");
      return;
    }
    const success = addSeller(newSellerName, newSellerEmail, Number(newSellerCommission) || 0, newSellerWarehouseId);
    if (success) {
      toast.success(`Vendedor "${newSellerName}" registrado con éxito. Contraseña por defecto: vendedor123`);
      setSelectedSellerEmail(newSellerEmail);
      setNewSellerName("");
      setNewSellerEmail("");
      setNewSellerCommission("10");
      setNewSellerWarehouseId("01");
      setShowAddForm(false);
    } else {
      toast.error("Este vendedor ya existe en el sistema.");
    }
  };

  const handleSaveCommission = () => {
    const val = Number(editCommissionValue);
    if (isNaN(val) || val < 0 || val > 100) {
      toast.error("Por favor ingresa un porcentaje válido entre 0 y 100.");
      return;
    }
    const success = updateSellerCommission(selectedSellerEmail, val);
    if (success) {
      toast.success("Comisión actualizada con éxito.");
      setIsEditingCommission(false);
    } else {
      toast.error("No se pudo actualizar la comisión.");
    }
  };

  const handleSaveWarehouse = () => {
    const val = editWarehouseValue.trim();
    if (!val || val.length > 2) {
      toast.error("Por favor ingresa un ID de bodega válido de hasta 2 dígitos.");
      return;
    }
    const success = updateSellerWarehouseId(selectedSellerEmail, val);
    if (success) {
      toast.success("Bodega actualizada con éxito.");
      setIsEditingWarehouse(false);
    } else {
      toast.error("No se pudo actualizar la bodega.");
    }
  };

  const handleFile = async (file: File) => {
    if (!selectedSellerEmail) {
      toast.error("Selecciona un vendedor primero.");
      return;
    }
    setFileLoading(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rawJson = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
      if (!rawJson.length) {
        toast.error("La hoja está vacía");
        return;
      }
      
      // Clean keys in every row object (trim whitespace from header names)
      const json: Record<string, string | number>[] = rawJson.map((row) => {
        const cleanRow: Record<string, string | number> = {};
        Object.keys(row).forEach((k) => {
          cleanRow[k.trim()] = row[k] as string | number;
        });
        return cleanRow;
      });

      const columns = Object.keys(json[0]);
      const tempKeys = detectKeyColumns(columns);
      const mergedRowsMap: Record<string, InventoryRow> = {};
      
      json.forEach((r, i) => {
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
            ...(r as Record<string, string | number>),
            __id: stableId,
          };
        }
      });

      const rows = Object.values(mergedRowsMap);
      setInventory(selectedSellerEmail, columns, rows);
      toast.success(`Inventario cargado: ${rows.length} productos para ${selectedSellerEmail}`);
    } catch (e) {
      console.error(e);
      toast.error("No se pudo leer el archivo");
    } finally {
      setFileLoading(false);
    }
  };

  // Exporter for the selected seller
  const keys = useMemo(() => detectKeyColumns(selectedInv.columns), [selectedInv.columns]);
  const totalSoldUnits = useMemo(() => Object.values(selectedInv.sales).reduce((a, b) => a + b, 0), [selectedInv.sales]);
  const productsWithSales = useMemo(() => Object.keys(selectedInv.sales).length, [selectedInv.sales]);
  const totalStockUnits = useMemo(() => {
    const stockKey = keys.stock;
    if (!stockKey) return 0;
    return selectedInv.rows.reduce((sum, r) => sum + (Number(r[stockKey]) || 0), 0);
  }, [selectedInv.rows, keys.stock]);

  const filteredAdminRows = useMemo(() => {
    const term = tableFilter.trim().toLowerCase();
    if (!term) return selectedInv.rows;
    return selectedInv.rows.filter((r) =>
      selectedInv.columns.some((c) =>
        String(r[c] ?? "").toLowerCase().includes(term)
      )
    );
  }, [selectedInv.rows, selectedInv.columns, tableFilter]);

  const handleExportSellerSales = () => {
    if (!keys.ref) {
      toast.error("No se encontró la columna de Referencia en el inventario.");
      return;
    }
    const header = [
      "StrProducto",
      "IntCantidaddoc",
      "IntvalorUnitario",
      "IntBodega",
      "StrLote",
      "StrColor",
    ];
    const data: (string | number)[][] = [header];
    for (const r of selectedInv.rows) {
      const qty = selectedInv.sales[r.__id];
      if (!qty || qty <= 0) continue;
      const rawColor = keys.color ? String(r[keys.color] ?? "") : "";
      data.push([
        String(r[keys.ref!] ?? ""),
        qty,
        0,
        activeSellerObj?.warehouseId ?? "01",
        keys.size ? String(r[keys.size] ?? "") : "",
        getColorCode(rawColor),
      ]);
    }
    if (data.length === 1) {
      toast.error("No hay ventas registradas para exportar.");
      return;
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
    const sellerSlug = selectedSellerEmail.split("@")[0];
    XLSX.writeFile(wb, `REPORTE_${sellerSlug.toUpperCase()}_${stamp}.xls`, { bookType: "biff8" });

    // Save log entry to history
    const loggedItems = [];
    for (const r of selectedInv.rows) {
      const qty = selectedInv.sales[r.__id];
      if (!qty || qty <= 0) continue;
      const refVal = keys.ref ? String(r[keys.ref] ?? "") : "";
      const nameVal = keys.name ? String(r[keys.name] ?? "") : (refVal || "Producto");
      const sizeVal = keys.size ? String(r[keys.size] ?? "") : "";
      const colorVal = keys.color ? String(r[keys.color] ?? "") : "";
      loggedItems.push({
        __id: r.__id,
        ref: refVal,
        name: nameVal,
        variantDesc: `${translateColor(colorVal)} / Talla ${sizeVal}`,
        qty,
        price: keys.price ? parsePrice(r[keys.price]) : 0,
        stock: keys.stock ? Number(r[keys.stock]) || null : null,
        size: sizeVal,
        color: colorVal,
      });
    }

    if (loggedItems.length > 0) {
      addExportLog(
        selectedSellerEmail,
        activeSellerObj?.warehouseId ?? "01",
        activeSellerObj?.commission ?? 10,
        loggedItems
      );
    }

    toast.success(`Archivo exportado con ${data.length - 1} líneas para el vendedor.`);
  };

  const handleExportSellerCommissions = () => {
    if (!keys.ref) {
      toast.error("No se encontró la columna de Referencia en el inventario.");
      return;
    }
    const commPercent = activeSellerObj?.commission ?? 0;
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

    let totalQty = 0;
    let totalVentaConIva = 0;
    let totalVentaSinIva = 0;
    let totalComision = 0;
    let totalIvaComision = 0;
    let totalRetefuente = 0;
    let totalTotalComision = 0;
    let totalTotalReembolsar = 0;

    for (const r of selectedInv.rows) {
      const qty = selectedInv.sales[r.__id];
      if (!qty || qty <= 0) continue;
      const price = keys.price ? parsePrice(r[keys.price]) : 0;
      
      const ventaConIva = qty * price;
      const ventaSinIva = Math.round(ventaConIva / 1.19);
      const comision = Math.round(ventaSinIva * (commPercent / 100));
      const ivaComision = Math.round(comision * 0.19);
      const retefuente = Math.round(comision * 0.11);
      const rowTotalComision = comision + ivaComision - retefuente;
      const rowTotalReembolsar = ventaConIva - rowTotalComision;

      totalQty += qty;
      totalVentaConIva += ventaConIva;
      totalVentaSinIva += ventaSinIva;
      totalComision += comision;
      totalIvaComision += ivaComision;
      totalRetefuente += retefuente;
      totalTotalComision += rowTotalComision;
      totalTotalReembolsar += rowTotalReembolsar;

      const refVal = keys.ref ? String(r[keys.ref] ?? "") : "";
      const nameVal = keys.name ? String(r[keys.name] ?? "") : (refVal || "Producto");

      data.push([
        refVal,
        refVal,
        nameVal,
        keys.size ? String(r[keys.size] ?? "") : "",
        keys.color ? translateColor(String(r[keys.color] ?? "")) : "",
        qty,
        price,
        ventaConIva,
        ventaSinIva,
        comision,
        ivaComision,
        retefuente,
        rowTotalComision,
        rowTotalReembolsar,
      ]);
    }

    if (data.length === 1) {
      toast.error("No hay ventas registradas para exportar.");
      return;
    }

    // Add Totals row
    data.push([
      "TOTAL",
      "",
      "",
      "",
      "",
      totalQty,
      "",
      totalVentaConIva,
      totalVentaSinIva,
      totalComision,
      totalIvaComision,
      totalRetefuente,
      totalTotalComision,
      totalTotalReembolsar,
    ]);

    // Add empty row and summary card
    data.push([]);
    data.push(["TOTAL A REEMBOLSAR", "Reembolso a la marca"]);
    data.push([totalTotalReembolsar]);

    const ws = XLSX.utils.aoa_to_sheet(data);
    const range = XLSX.utils.decode_range(ws["!ref"]!);
    
    const textCols = [0, 1, 2, 3, 4];
    const currencyCols = [6, 7, 8, 9, 10, 11, 12, 13];

    for (let R = 1; R <= range.e.r; R++) {
      for (let C = 0; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr];
        if (!cell) continue;

        if (textCols.includes(C) && R < data.length - 3) {
          cell.t = "s";
          cell.v = String(cell.v);
        } else if (currencyCols.includes(C) && typeof cell.v === "number") {
          cell.t = "n";
          cell.z = "$#,##0";
        }
      }
    }

    // Format summary cell at the bottom
    const lastRowIndex = data.length - 1;
    const summaryCellAddr = XLSX.utils.encode_cell({ r: lastRowIndex, c: 0 });
    const summaryCell = ws[summaryCellAddr];
    if (summaryCell && typeof summaryCell.v === "number") {
      summaryCell.t = "n";
      summaryCell.z = "$#,##0";
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Comisiones");
    const stamp = new Date().toISOString().slice(0, 10);
    const sellerSlug = selectedSellerEmail.split("@")[0];
    XLSX.writeFile(wb, `REPORTE_COMISIONES_${sellerSlug.toUpperCase()}_${stamp}.xls`, { bookType: "biff8" });

    // Save log entry to history
    const loggedItems = [];
    for (const r of selectedInv.rows) {
      const qty = selectedInv.sales[r.__id];
      if (!qty || qty <= 0) continue;
      const refVal = keys.ref ? String(r[keys.ref] ?? "") : "";
      const nameVal = keys.name ? String(r[keys.name] ?? "") : (refVal || "Producto");
      const sizeVal = keys.size ? String(r[keys.size] ?? "") : "";
      const colorVal = keys.color ? String(r[keys.color] ?? "") : "";
      loggedItems.push({
        __id: r.__id,
        ref: refVal,
        name: nameVal,
        variantDesc: `${translateColor(colorVal)} / Talla ${sizeVal}`,
        qty,
        price: keys.price ? parsePrice(r[keys.price]) : 0,
        stock: keys.stock ? Number(r[keys.stock]) || null : null,
        size: sizeVal,
        color: colorVal,
      });
    }

    if (loggedItems.length > 0) {
      addExportLog(
        selectedSellerEmail,
        activeSellerObj?.warehouseId ?? "01",
        activeSellerObj?.commission ?? 10,
        loggedItems
      );
    }

    toast.success("¡Reporte de comisiones exportado con éxito!");
  };

  const handleDownloadTemplate = () => {
    const headers = [
      "StrProducto",
      "Descripción",
      "StrLote",
      "StrColor",
      "IntCantidaddoc",
      "PVP"
    ];
    const sampleRows = [
      ["B01097081", "Blusa Encaje Manga Larga", "S", "01", 10, 45000],
      ["B01097081", "Blusa Encaje Manga Larga", "M", "01", 15, 45000],
      ["B01097081", "Blusa Encaje Manga Larga", "S", "02", 5, 45000],
      ["B01097082", "Jeans Skinny Tiro Alto", "06", "08", 20, 85000],
      ["B01097082", "Jeans Skinny Tiro Alto", "08", "08", 25, 85000],
      ["B01097082", "Jeans Skinny Tiro Alto", "06", "51", 12, 85000],
    ];
    const data = [headers, ...sampleRows];
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    const range = XLSX.utils.decode_range(ws["!ref"]!);
    for (let R = 1; R <= range.e.r; R++) {
      const qtyAddr = XLSX.utils.encode_cell({ r: R, c: 4 });
      const qtyCell = ws[qtyAddr];
      if (qtyCell) {
        qtyCell.t = "n";
      }

      const pvpAddr = XLSX.utils.encode_cell({ r: R, c: 5 });
      const pvpCell = ws[pvpAddr];
      if (pvpCell) {
        pvpCell.t = "n";
        pvpCell.z = "$#,##0";
      }

      for (const C of [0, 2, 3]) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr];
        if (cell) {
          cell.t = "s";
          cell.v = String(cell.v);
        }
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventario_Plantilla");
    XLSX.writeFile(wb, "PLANTILLA_INVENTARIO_CLIENTE.xls", { bookType: "biff8" });
    toast.success("¡Plantilla de inventario descargada con éxito!");
  };

  if (authLoading || !user || user.role !== "admin") {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background">
        <div className="flex flex-col items-center gap-2">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <p className="text-sm text-muted-foreground">Verificando acceso...</p>
        </div>
      </div>
    );
  }

  // Hook details of active seller object
  const activeSellerObj = sellers.find((s) => s.email === selectedSellerEmail);

  return (
    <div className="min-h-screen bg-background animate-in fade-in duration-300">
      {/* Header */}
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
          <div>
            <h1 className="text-xl font-bold tracking-tight text-foreground">Gestión de Inventario</h1>
            <p className="text-sm text-muted-foreground flex items-center gap-1.5 flex-wrap">
              <span>Panel del administrador</span>
              <span className="h-1 w-1 rounded-full bg-muted-foreground/45" />
              <span className="text-xs font-semibold text-primary">Hola, {user.name}</span>
            </p>
          </div>
          <div className="flex items-center gap-3">
            <Link
              to="/proveedor"
              className="rounded-md border border-border bg-background px-4 py-2 text-sm font-medium text-foreground hover:bg-accent transition-colors"
            >
              Ver Vista de Ventas →
            </Link>
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                logout();
                toast.success("Sesión cerrada");
              }}
              className="gap-2 border-destructive/20 text-destructive hover:bg-destructive/10 hover:text-destructive transition-all"
            >
              <LogOut className="h-4 w-4" />
              Salir
            </Button>
          </div>
        </div>
      </header>

      {/* Main Body */}
      <main className="mx-auto max-w-7xl px-6 py-8">
        <div className="grid gap-6 md:grid-cols-12">
          
          {/* Sidebar: Sellers List - 4 columns */}
          <div className="md:col-span-4 space-y-4">
            <Card className="border-border/60 shadow-md">
              <CardHeader className="pb-3 flex flex-row items-center justify-between space-y-0">
                <div>
                  <CardTitle className="text-lg">Vendedores</CardTitle>
                  <CardDescription>Selecciona un vendedor para ver su estado</CardDescription>
                </div>
                <Button
                  variant="outline"
                  size="icon"
                  className="h-8 w-8 rounded-full"
                  onClick={() => setShowAddForm(!showAddForm)}
                >
                  <Plus className={`h-4 w-4 transition-transform duration-300 ${showAddForm ? "rotate-45" : ""}`} />
                </Button>
              </CardHeader>
              <CardContent className="space-y-3">
                {/* Form to add seller */}
                {showAddForm && (
                  <form onSubmit={handleAddSeller} className="border border-border/80 rounded-lg p-3 bg-muted/40 space-y-3 animate-in slide-in-from-top-2 duration-300">
                    <p className="text-xs font-semibold text-foreground uppercase tracking-wider">Nuevo Vendedor</p>
                    <div className="space-y-1.5">
                      <Label htmlFor="sellerName" className="text-xs">Nombre</Label>
                      <Input
                        id="sellerName"
                        placeholder="Ej. Juan Pérez"
                        value={newSellerName}
                        onChange={(e) => setNewSellerName(e.target.value)}
                        className="h-8 text-xs"
                        required
                      />
                    </div>
                    <div className="space-y-1.5">
                      <Label htmlFor="sellerEmail" className="text-xs">Correo</Label>
                      <Input
                        id="sellerEmail"
                        type="email"
                        placeholder="juan@comodatos.com"
                        value={newSellerEmail}
                        onChange={(e) => setNewSellerEmail(e.target.value)}
                        className="h-8 text-xs"
                        required
                      />
                    </div>
                    <div className="space-y-1.5">
                      <Label htmlFor="sellerCommission" className="text-xs">Comisión (%)</Label>
                      <Input
                        id="sellerCommission"
                        type="number"
                        min="0"
                        max="100"
                        step="0.1"
                        placeholder="Ej. 10"
                        value={newSellerCommission}
                        onChange={(e) => setNewSellerCommission(e.target.value)}
                        className="h-8 text-xs"
                        required
                      />
                    </div>
                    <div className="space-y-1.5">
                      <Label htmlFor="sellerWarehouseId" className="text-xs">Bodega ID (2 dígitos)</Label>
                      <Input
                        id="sellerWarehouseId"
                        placeholder="Ej. 01"
                        maxLength={2}
                        value={newSellerWarehouseId}
                        onChange={(e) => setNewSellerWarehouseId(e.target.value.replace(/\D/g, "").slice(0, 2))}
                        className="h-8 text-xs"
                        required
                      />
                    </div>
                    <div className="flex justify-end gap-2 pt-1">
                      <Button type="button" variant="ghost" size="sm" className="h-7 text-xs" onClick={() => setShowAddForm(false)}>
                        Cancelar
                      </Button>
                      <Button type="submit" size="sm" className="h-7 text-xs">
                        Guardar
                      </Button>
                    </div>
                  </form>
                )}

                {/* List of sellers */}
                <div className="space-y-1.5 max-h-[350px] overflow-y-auto">
                  {sellers.map((s) => {
                    const isSelected = s.email === selectedSellerEmail;
                    return (
                      <SellerListItem
                        key={s.email}
                        seller={s}
                        isSelected={isSelected}
                        onSelect={() => setSelectedSellerEmail(s.email)}
                      />
                    );
                  })}
                  {sellers.length === 0 && (
                    <p className="text-center text-sm text-muted-foreground py-6">No hay vendedores registrados.</p>
                  )}
                </div>
              </CardContent>
            </Card>

            {/* HISTORIAL GLOBAL DE EXPORTACIONES */}
            <Card className="border-border/60 shadow-md">
              <CardHeader className="pb-3 flex flex-row items-center justify-between space-y-0">
                <div>
                  <CardTitle className="text-sm font-bold flex items-center gap-1.5">
                    <FileText className="h-4 w-4 text-primary" />
                    Historial de Ventas
                  </CardTitle>
                  <CardDescription className="text-3xs">Reportes de ventas exportados por vendedores</CardDescription>
                </div>
              </CardHeader>
              <CardContent className="space-y-3">
                {allLogs.length === 0 ? (
                  <p className="text-center text-xs text-muted-foreground py-6">No hay registros de exportación aún.</p>
                ) : (
                  <div className="space-y-2.5 max-h-[350px] overflow-y-auto pr-1">
                    {allLogs.map((log) => {
                      const seller = sellers.find(s => s.email.toLowerCase() === log.sellerEmail.toLowerCase());
                      return (
                        <div key={log.id} className="p-3 border border-border/50 rounded-lg hover:bg-muted/10 transition-colors space-y-2 text-xs">
                          <div className="flex justify-between items-start">
                            <div>
                              <p className="font-extrabold text-foreground">{seller?.name || log.sellerEmail.split("@")[0]}</p>
                              <p className="text-3xs text-muted-foreground">
                                {new Date(log.timestamp).toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" })}
                              </p>
                            </div>
                            <Button
                              size="icon"
                              variant="ghost"
                              onClick={() => handleDeleteLog(log.id)}
                              className="h-6 w-6 text-muted-foreground hover:text-red-500 hover:bg-red-500/5 rounded-md"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </Button>
                          </div>
                          <div className="text-3xs text-muted-foreground flex justify-between bg-muted/40 p-2 rounded-md">
                            <span>Cant: <strong>{log.totalUnits} uds</strong></span>
                            <span>Valor: <strong>${log.totalAmount.toLocaleString("es-CO")}</strong></span>
                            {log.commissionPercent > 0 && <span>Com: <strong>{log.commissionPercent}%</strong></span>}
                          </div>
                          <div className="flex gap-2">
                            <Button
                              size="sm"
                              variant="outline"
                              onClick={() => handleDownloadSalesFromLog(log)}
                              className="h-7 w-full text-3xs font-bold gap-1 rounded-md border-emerald-600/30 text-emerald-700 hover:bg-emerald-50/50 dark:hover:bg-emerald-950/20"
                            >
                              <FileSpreadsheet className="h-3.5 w-3.5" />
                              Ventas
                            </Button>
                            {log.commissionPercent > 0 && (
                              <Button
                                size="sm"
                                variant="outline"
                                onClick={() => handleDownloadCommissionsFromLog(log)}
                                className="h-7 w-full text-3xs font-bold gap-1 rounded-md border-indigo-600/30 text-indigo-700 hover:bg-indigo-50/50 dark:hover:bg-indigo-950/20"
                              >
                                <FileSpreadsheet className="h-3.5 w-3.5" />
                                Comisiones
                              </Button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </CardContent>
            </Card>
          </div>

          {/* Main Panel: Seller Inventory Actions - 8 columns */}
          <div className="md:col-span-8 space-y-6">
            {activeSellerObj ? (
              <div className="space-y-6">
                
                {/* Active Seller Banner */}
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 p-5 rounded-lg border border-border/60 bg-card shadow-sm">
                  <div>
                    <div className="flex items-center gap-2">
                      <div className="h-2 w-2 rounded-full bg-primary" />
                      <h2 className="text-xl font-bold text-foreground">{activeSellerObj.name}</h2>
                    </div>
                    <p className="text-sm text-muted-foreground mt-0.5">{activeSellerObj.email}</p>
                    <div className="flex items-center gap-3 mt-2 flex-wrap text-xs text-muted-foreground">
                      {/* Comisión Edit block */}
                      <div className="flex items-center gap-1.5">
                        <span>Comisión:</span>
                        {isEditingCommission ? (
                          <span className="inline-flex items-center gap-1.5">
                            <Input
                              type="number"
                              min="0"
                              max="100"
                              step="0.1"
                              value={editCommissionValue}
                              onChange={(e) => setEditCommissionValue(e.target.value)}
                              className="h-7 w-16 text-xs py-0 px-2 border-border/80 bg-background"
                              required
                            />
                            <span className="text-xs font-bold text-foreground">%</span>
                            <Button
                              size="sm"
                              className="h-7 px-2 py-0 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold gap-1 rounded-md"
                              onClick={handleSaveCommission}
                            >
                              Guardar
                            </Button>
                            <Button
                              size="sm"
                              variant="ghost"
                              className="h-7 px-2 py-0 text-xs font-medium text-muted-foreground hover:bg-accent rounded-md"
                              onClick={() => setIsEditingCommission(false)}
                            >
                              Cancelar
                            </Button>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1.5">
                            <span className="font-extrabold text-foreground bg-primary/10 text-primary px-2 py-0.5 rounded text-2xs border border-primary/15">
                              {activeSellerObj.commission ?? 0}%
                            </span>
                            <button
                              onClick={() => {
                                setEditCommissionValue(String(activeSellerObj.commission ?? 0));
                                setIsEditingCommission(true);
                              }}
                              className="text-xs text-primary hover:underline font-bold"
                            >
                              Editar
                            </button>
                          </span>
                        )}
                      </div>

                      <span className="h-3 w-px bg-border hidden sm:block" />

                      {/* Bodega Edit block */}
                      <div className="flex items-center gap-1.5">
                        <span>Bodega:</span>
                        {isEditingWarehouse ? (
                          <span className="inline-flex items-center gap-1.5">
                            <Input
                              type="text"
                              maxLength={2}
                              value={editWarehouseValue}
                              onChange={(e) => setEditWarehouseValue(e.target.value.replace(/\D/g, "").slice(0, 2))}
                              className="h-7 w-12 text-xs py-0 px-2 border-border/80 bg-background"
                              required
                            />
                            <Button
                              size="sm"
                              className="h-7 px-2 py-0 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold gap-1 rounded-md"
                              onClick={handleSaveWarehouse}
                            >
                              Guardar
                            </Button>
                            <Button
                              size="sm"
                              variant="ghost"
                              className="h-7 px-2 py-0 text-xs font-medium text-muted-foreground hover:bg-accent rounded-md"
                              onClick={() => setIsEditingWarehouse(false)}
                            >
                              Cancelar
                            </Button>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1.5">
                            <span className="font-extrabold text-foreground bg-indigo-50 text-indigo-600 dark:bg-indigo-950/20 dark:text-indigo-400 px-2 py-0.5 rounded text-2xs border border-indigo-150">
                              {activeSellerObj.warehouseId ?? "01"}
                            </span>
                            <button
                              onClick={() => {
                                setEditWarehouseValue(activeSellerObj.warehouseId ?? "01");
                                setIsEditingWarehouse(true);
                              }}
                              className="text-xs text-primary hover:underline font-bold"
                            >
                              Editar
                            </button>
                          </span>
                        )}
                      </div>
                    </div>
                  </div>
                  {selectedInv.uploadedAt ? (
                    <div className="flex items-center gap-1.5 text-xs text-emerald-600 bg-emerald-50 dark:bg-emerald-950/20 dark:text-emerald-400 border border-emerald-200/55 rounded-full px-3 py-1 font-medium w-fit">
                      <CheckCircle2 className="h-3.5 w-3.5" />
                      <span>Inventario cargado ({selectedInv.rows.length} refs)</span>
                    </div>
                  ) : (
                    <div className="flex items-center gap-1.5 text-xs text-amber-600 bg-amber-50 dark:bg-amber-950/20 dark:text-amber-400 border border-amber-200/55 rounded-full px-3 py-1 font-medium w-fit">
                      <AlertCircle className="h-3.5 w-3.5" />
                      <span>Sin inventario asignado</span>
                    </div>
                  )}
                </div>

                {/* Dashboard Stats */}
                <div className="grid gap-4 grid-cols-2 lg:grid-cols-4">
                  <StatCard label="Total Productos" value={selectedInv.rows.length} />
                  <StatCard label="Unidades Inventario" value={totalStockUnits} />
                  <StatCard label="Refs Con Ventas" value={productsWithSales} />
                  <StatCard label="Unidades Vendidas" value={totalSoldUnits} />
                </div>

                {/* Inventory upload card */}
                <Card className="border-border/60 shadow-sm">
                  <CardHeader>
                    <CardTitle className="text-base flex items-center gap-2">
                      <FileSpreadsheet className="h-5 w-5 text-primary" />
                      1. Subir inventario Excel (.xlsx)
                    </CardTitle>
                    <CardDescription>
                      Asigna o reemplaza el inventario de {activeSellerObj.name}. La primera fila del archivo debe contener los nombres de las columnas.
                    </CardDescription>
                  </CardHeader>
                  <CardContent>
                    <input
                      ref={inputRef}
                      type="file"
                      accept=".xlsx,.xls"
                      className="hidden"
                      onChange={(e) => {
                        const f = e.target.files?.[0];
                        if (f) handleFile(f);
                        e.target.value = "";
                      }}
                    />
                    <div className="flex flex-wrap items-center gap-3">
                      <Button onClick={() => inputRef.current?.click()} disabled={fileLoading}>
                        {fileLoading ? "Procesando…" : selectedInv.rows.length ? "Reemplazar archivo" : "Seleccionar archivo"}
                      </Button>
                      
                      <Button
                        variant="outline"
                        onClick={handleDownloadTemplate}
                        className="gap-1.5 border-primary/20 hover:border-primary/45 text-primary hover:bg-primary/5 font-semibold"
                      >
                        <FileSpreadsheet className="h-4 w-4" />
                        Descargar Plantilla
                      </Button>
                      
                      {selectedInv.uploadedAt && (
                        <>
                          <Button
                            variant="outline"
                            onClick={handleExportSellerSales}
                            disabled={productsWithSales === 0}
                            className="gap-1.5"
                          >
                            <Download className="h-4 w-4" />
                            Consolidar Reporte
                          </Button>
                          <Button
                            variant="ghost"
                            onClick={() => {
                              clearSellerInventory(selectedSellerEmail);
                              toast.success("Inventario eliminado");
                            }}
                            className="text-destructive hover:bg-destructive/10 hover:text-destructive gap-1.5 ml-auto"
                          >
                            <Trash2 className="h-4 w-4" />
                            Borrar Inventario
                          </Button>
                          {productsWithSales > 0 && (
                            <Button
                              variant="ghost"
                              onClick={() => {
                                resetSales(selectedSellerEmail);
                                toast.success("Ventas reiniciadas");
                              }}
                              className="text-muted-foreground hover:bg-accent gap-1.5"
                            >
                              <RefreshCw className="h-3.5 w-3.5" />
                              Reiniciar Ventas
                            </Button>
                          )}
                        </>
                      )}
                    </div>
                    {selectedInv.uploadedAt && (
                      <p className="text-xs text-muted-foreground mt-3">
                        Última actualización: {new Date(selectedInv.uploadedAt).toLocaleString()}
                      </p>
                    )}
                  </CardContent>
                </Card>

                {/* Table of loaded products */}
                {selectedInv.rows.length > 0 && (
                  <Card className="border-border/60 shadow-sm">
                    <CardHeader className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                      <div>
                        <CardTitle className="text-base flex items-center gap-2">
                          <FileText className="h-5 w-5 text-primary" />
                          2. Productos cargados en sistema
                        </CardTitle>
                        <CardDescription>
                          Vista preliminar de las referencias del vendedor. Se muestran las primeras 25 filas.
                        </CardDescription>
                      </div>
                      <div className="flex flex-wrap gap-2 sm:ml-auto">
                        <Button
                          onClick={handleExportSellerSales}
                          disabled={productsWithSales === 0}
                          className="gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold shadow-xs shrink-0"
                          title="Descargar reporte estándar de ventas para cargar al sistema"
                        >
                          <FileSpreadsheet className="h-4 w-4" />
                          Reporte de Ventas
                        </Button>
                        <Button
                          onClick={handleExportSellerCommissions}
                          disabled={productsWithSales === 0}
                          className="gap-2 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold shadow-xs shrink-0"
                          title="Descargar reporte detallado de ventas con comisiones de vendedor"
                        >
                          <FileSpreadsheet className="h-4 w-4" />
                          Reporte de Comisiones
                        </Button>
                      </div>
                    </CardHeader>
                    <CardContent>
                      {/* Search box for preview table */}
                      <div className="relative mb-3.5 max-w-sm">
                        <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-muted-foreground" />
                        <Input
                          placeholder="Buscar por referencia, color, talla, nombre..."
                          value={tableFilter}
                          onChange={(e) => setTableFilter(e.target.value)}
                          className="pl-8 h-8 text-xs rounded-lg"
                        />
                      </div>

                      <div className="overflow-auto rounded-md border border-border bg-card">
                        <table className="w-full text-xs">
                          <thead className="bg-muted/80 text-muted-foreground border-b border-border">
                            <tr>
                              {selectedInv.columns.map((c) => (
                                <th key={c} className="whitespace-nowrap px-3 py-2.5 text-left font-semibold">
                                  {c}
                                </th>
                              ))}
                              <th className="whitespace-nowrap px-3 py-2.5 text-right font-semibold min-w-[140px]">
                                Ventas (Montar / Editar)
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {filteredAdminRows.slice(0, 50).map((r) => {
                              const unitsSold = selectedInv.sales[r.__id] || 0;
                              return (
                                <tr key={r.__id} className="border-b border-border/50 last:border-0 hover:bg-muted/30 transition-colors">
                                  {selectedInv.columns.map((c) => (
                                    <td key={c} className="whitespace-nowrap px-3 py-2.5 text-foreground/80">
                                      {c === keys.color ? translateColor(r[c]) : String(r[c] ?? "")}
                                    </td>
                                  ))}
                                  <td className="whitespace-nowrap px-3 py-2.5 text-right">
                                    <div className="flex items-center justify-end gap-1">
                                      <Button
                                        type="button"
                                        size="icon"
                                        variant="outline"
                                        onClick={() => setSale(selectedSellerEmail, r.__id, Math.max(0, unitsSold - 1))}
                                        disabled={unitsSold === 0}
                                        className="h-6 w-6 rounded text-xs font-bold shrink-0 disabled:opacity-30"
                                        title="Restar 1 unidad"
                                      >
                                        -
                                      </Button>
                                      <Input
                                        type="number"
                                        min={0}
                                        value={unitsSold || ""}
                                        placeholder="0"
                                        onChange={(e) => {
                                          const val = parseInt(e.target.value, 10);
                                          setSale(selectedSellerEmail, r.__id, isNaN(val) ? 0 : Math.max(0, val));
                                        }}
                                        className={`h-6 w-12 text-center font-mono font-bold text-xs rounded transition-colors px-1 ${
                                          unitsSold > 0
                                            ? "bg-emerald-50 dark:bg-emerald-950/40 text-emerald-600 dark:text-emerald-400 border-emerald-500/50"
                                            : "bg-background text-foreground border-border"
                                        }`}
                                      />
                                      <Button
                                        type="button"
                                        size="icon"
                                        variant="outline"
                                        onClick={() => setSale(selectedSellerEmail, r.__id, unitsSold + 1)}
                                        className="h-6 w-6 rounded text-xs font-bold shrink-0 bg-primary/10 border-primary/30 text-primary hover:bg-primary/20"
                                        title="Sumar 1 unidad"
                                      >
                                        +
                                      </Button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                      <div className="mt-2.5 flex items-center justify-between text-2xs text-muted-foreground">
                        <span>
                          {filteredAdminRows.length === 0 ? "No se encontraron coincidencias." : `Mostrando ${Math.min(50, filteredAdminRows.length)} de ${filteredAdminRows.length} resultados.`}
                        </span>
                        {selectedInv.rows.length > 50 && (
                          <span className="italic">
                            (Total inventario: {selectedInv.rows.length} referencias)
                          </span>
                        )}
                      </div>
                    </CardContent>
                  </Card>
                )}

              </div>
            ) : (
              <div className="flex flex-col items-center justify-center p-12 rounded-lg border border-dashed border-border/80 text-center min-h-[400px]">
                <User className="h-10 w-10 text-muted-foreground/60 mb-3 bg-muted p-2 rounded-full" />
                <h3 className="font-bold text-lg text-foreground">Ningún vendedor seleccionado</h3>
                <p className="text-sm text-muted-foreground mt-1 max-w-sm">
                  Crea un vendedor en el panel lateral o selecciona uno de la lista para gestionar su inventario y consolidar reportes.
                </p>
              </div>
            )}
          </div>

        </div>
      </main>
    </div>
  );
}

// Seller List Item Component to optimize code layout
function SellerListItem({
  seller,
  isSelected,
  onSelect,
}: {
  seller: { email: string; name: string; commission?: number; warehouseId?: string };
  isSelected: boolean;
  onSelect: () => void;
}) {
  const sellerInv = useInventory(seller.email);
  const hasInventory = sellerInv.rows.length > 0;

  const sellerKeys = useMemo(() => detectKeyColumns(sellerInv.columns), [sellerInv.columns]);
  const totalSellerUnits = useMemo(() => {
    const stockKey = sellerKeys.stock;
    if (!stockKey) return 0;
    return sellerInv.rows.reduce((sum, r) => sum + (Number(r[stockKey]) || 0), 0);
  }, [sellerInv.rows, sellerKeys.stock]);

  return (
    <button
      onClick={onSelect}
      className={`w-full flex items-center justify-between p-3 rounded-lg border text-left transition-all ${
        isSelected
          ? "bg-primary/10 border-primary text-foreground shadow-sm"
          : "bg-card hover:bg-muted/40 border-border/60 text-muted-foreground hover:text-foreground"
      }`}
    >
      <div className="truncate pr-2">
        <p className={`font-semibold text-xs truncate ${isSelected ? "text-primary" : "text-foreground"}`}>
          {seller.name} {seller.commission !== undefined && `(${seller.commission}%)`}
        </p>
        <p className="text-2xs text-muted-foreground truncate">
          Bodega: {seller.warehouseId ?? "01"} • {seller.email}
        </p>
      </div>
      <div className="shrink-0 flex flex-col items-end gap-0.5">
        {hasInventory ? (
          <>
            <span className="text-2xs bg-emerald-50 text-emerald-600 dark:bg-emerald-950/20 dark:text-emerald-400 font-semibold px-2 py-0.5 rounded-full border border-emerald-100 dark:border-emerald-900/30">
              {sellerInv.rows.length} refs
            </span>
            <span className="text-3xs text-muted-foreground font-mono font-medium">
              {totalSellerUnits.toLocaleString()} uds
            </span>
          </>
        ) : (
          <span className="text-2xs bg-muted text-muted-foreground font-medium px-2 py-0.5 rounded-full border border-border">
            Vacío
          </span>
        )}
      </div>
    </button>
  );
}

function StatCard({ label, value }: { label: string; value: number }) {
  return (
    <Card className="border-border/60 shadow-sm hover:shadow-md transition-shadow">
      <CardContent className="p-4 pt-5">
        <p className="text-2xs text-muted-foreground uppercase font-semibold tracking-wider truncate">{label}</p>
        <p className="mt-1 text-xl font-bold text-foreground">{value.toLocaleString()}</p>
      </CardContent>
    </Card>
  );
}