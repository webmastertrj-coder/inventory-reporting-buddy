import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useRef, useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import { LogOut, Loader2, Plus, User, FileSpreadsheet, Trash2, Download, RefreshCw, FileText, CheckCircle2, AlertCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { setInventory, useInventory, clearSellerInventory, resetSales, detectKeyColumns, type InventoryRow } from "@/lib/inventory-store";
import { useAuth } from "@/lib/auth-store";
import { translateColor } from "@/lib/color-mapping";

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
  const { user, loading: authLoading, logout, sellers, addSeller } = useAuth();
  const navigate = useNavigate();
  const inputRef = useRef<HTMLInputElement>(null);
  const [fileLoading, setFileLoading] = useState(false);
  const [selectedSellerEmail, setSelectedSellerEmail] = useState<string>("");
  const [newSellerName, setNewSellerName] = useState("");
  const [newSellerEmail, setNewSellerEmail] = useState("");
  const [showAddForm, setShowAddForm] = useState(false);

  // Hook to get the inventory of the selected seller
  const selectedInv = useInventory(selectedSellerEmail);

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
    const success = addSeller(newSellerName, newSellerEmail);
    if (success) {
      toast.success(`Vendedor "${newSellerName}" registrado con éxito. Contraseña por defecto: vendedor123`);
      setSelectedSellerEmail(newSellerEmail);
      setNewSellerName("");
      setNewSellerEmail("");
      setShowAddForm(false);
    } else {
      toast.error("Este vendedor ya existe en el sistema.");
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
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
      if (!json.length) {
        toast.error("La hoja está vacía");
        return;
      }
      const columns = Object.keys(json[0]);
      const rows: InventoryRow[] = json.map((r, i) => ({
        ...(r as Record<string, string | number>),
        __id: `row_${i}_${Date.now()}`,
      }));
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
      data.push([
        String(r[keys.ref!] ?? ""),
        qty,
        0,
        "01",
        keys.size ? String(r[keys.size] ?? "") : "",
        keys.color ? String(r[keys.color] ?? "") : "",
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
    toast.success(`Archivo exportado con ${data.length - 1} líneas para el vendedor.`);
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
                    <div className="flex gap-2 justify-end pt-1">
                      <Button type="button" variant="ghost" size="sm" onClick={() => setShowAddForm(false)} className="h-7 text-xs">
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
                <div className="grid gap-4 grid-cols-3">
                  <StatCard label="Total Productos" value={selectedInv.rows.length} />
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
                    <CardHeader>
                      <CardTitle className="text-base flex items-center gap-2">
                        <FileText className="h-5 w-5 text-primary" />
                        2. Productos cargados en sistema
                      </CardTitle>
                      <CardDescription>
                        Vista preliminar de las referencias del vendedor. Se muestran las primeras 25 filas.
                      </CardDescription>
                    </CardHeader>
                    <CardContent>
                      <div className="overflow-auto rounded-md border border-border bg-card">
                        <table className="w-full text-xs">
                          <thead className="bg-muted/80 text-muted-foreground border-b border-border">
                            <tr>
                              {selectedInv.columns.map((c) => (
                                <th key={c} className="whitespace-nowrap px-3 py-2.5 text-left font-semibold">
                                  {c}
                                </th>
                              ))}
                              <th className="whitespace-nowrap px-3 py-2.5 text-right font-semibold">
                                Ventas Reportadas
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {selectedInv.rows.slice(0, 25).map((r) => {
                              const unitsSold = selectedInv.sales[r.__id] || 0;
                              return (
                                <tr key={r.__id} className="border-b border-border/50 last:border-0 hover:bg-muted/30 transition-colors">
                                  {selectedInv.columns.map((c) => (
                                    <td key={c} className="whitespace-nowrap px-3 py-2.5 text-foreground/80">
                                      {c === keys.color ? translateColor(r[c]) : String(r[c] ?? "")}
                                    </td>
                                  ))}
                                  <td className="whitespace-nowrap px-3 py-2.5 text-right font-medium text-primary">
                                    {unitsSold > 0 ? (
                                      <span className="bg-primary/10 text-primary px-2.5 py-0.5 rounded-full text-2xs font-semibold">
                                        {unitsSold} u.
                                      </span>
                                    ) : (
                                      <span className="text-muted-foreground/45">-</span>
                                    )}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                      {selectedInv.rows.length > 25 && (
                        <p className="mt-2.5 text-2xs text-muted-foreground text-right italic">
                          Mostrando 25 de {selectedInv.rows.length} productos totales.
                        </p>
                      )}
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
  seller: { email: string; name: string };
  isSelected: boolean;
  onSelect: () => void;
}) {
  const sellerInv = useInventory(seller.email);
  const hasInventory = sellerInv.rows.length > 0;

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
          {seller.name}
        </p>
        <p className="text-2xs text-muted-foreground truncate">{seller.email}</p>
      </div>
      <div className="shrink-0">
        {hasInventory ? (
          <span className="text-2xs bg-emerald-50 text-emerald-600 dark:bg-emerald-950/20 dark:text-emerald-400 font-semibold px-2 py-0.5 rounded-full border border-emerald-100 dark:border-emerald-900/30">
            {sellerInv.rows.length} refs
          </span>
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