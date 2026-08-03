import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useMemo, useState, useEffect } from "react";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import {
  LogOut,
  Loader2,
  Search,
  Plus,
  Minus,
  ShoppingBag,
  TrendingUp,
  AlertTriangle,
  RefreshCw,
  FileSpreadsheet,
  Trash2,
  Info,
  Shirt,
  Sparkles,
  Tag,
  CheckCircle2,
  Package,
  ChevronLeft,
  ChevronRight,
  History
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { detectKeyColumns, resetSales, setSale, useInventory, InventoryRow, parsePrice, addExportLog, deleteExportLog, useExportLogs } from "@/lib/inventory-store";
import { useAuth } from "@/lib/auth-store";
import { translateColor, getColorCode } from "@/lib/color-mapping";
import { generateSalesExcel, generateCommissionsExcel } from "@/lib/excel-helpers";

export const Route = createFileRoute("/proveedor")({
  head: () => ({
    meta: [
      { title: "Tienda de Reportes | Proveedor" },
      {
        name: "description",
        content: "Registra tus ventas de forma interactiva en tu tienda virtual.",
      },
    ],
  }),
  component: ProveedorPage,
});

interface GroupedProduct {
  key: string;
  name: string;
  ref: string;
  sku: string;
  totalStock: number;
  variants: InventoryRow[];
}

function ProveedorPage() {
  const { user, loading: authLoading, sellers } = useAuth();
  const navigate = useNavigate();
  const inv = useInventory(user?.email || "");
  const allLogs = useExportLogs();
  
  const myLogs = useMemo(() => {
    return allLogs.filter((log) => log.sellerEmail === user?.email?.toLowerCase());
  }, [allLogs, user]);

  const handleDownloadSalesFromLog = (log: any) => {
    const exportItems = log.items.map((item: any) => ({
      ref: item.ref,
      qty: item.qty,
      size: item.size,
      color: item.color,
    }));
    generateSalesExcel(exportItems, log.warehouseId, user?.name || user?.email || "vendedor");
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
    generateCommissionsExcel(exportItems, log.commissionPercent, user?.name || user?.email || "vendedor");
    toast.success("Liquidación de comisiones descargada del historial.");
  };

  const handleDeleteLog = (logId: string) => {
    deleteExportLog(logId);
    toast.success("Registro eliminado del historial.");
  };
  
  // UI States
  const [activeTab, setActiveTab] = useState<"register" | "report">("register");
  const [filter, setFilter] = useState("");
  const [showOnlySales, setShowOnlySales] = useState(false);
  const [isResetConfirmOpen, setIsResetConfirmOpen] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);

  // Reset pagination on filter change
  useEffect(() => {
    setCurrentPage(1);
  }, [filter, showOnlySales]);
  
  // E-commerce states: selected color and size per product
  // key: product.key -> { color: string, size: string }
  const [selectedVariants, setSelectedVariants] = useState<Record<string, { color: string; size: string }>>({});

  useEffect(() => {
    if (!authLoading && !user) {
      navigate({ to: "/login" });
    }
  }, [user, authLoading, navigate]);

  const keys = useMemo(() => detectKeyColumns(inv.columns), [inv.columns]);

  const commPercent = useMemo(() => {
    const s = sellers?.find((s) => s.email === user?.email);
    return s?.commission ?? 0;
  }, [sellers, user]);

  const warehouseId = useMemo(() => {
    const s = sellers?.find((s) => s.email === user?.email);
    return s?.warehouseId ?? "01";
  }, [sellers, user]);

  // Group products by Reference first, then SKU, then Name
  const groupedProducts = useMemo(() => {
    const groups: Record<string, GroupedProduct> = {};

    inv.rows.forEach((row) => {
      const nameVal = keys.name ? String(row[keys.name] ?? "").trim() : "";
      const refVal = keys.ref ? String(row[keys.ref] ?? "").trim() : "";
      const skuVal = keys.sku ? String(row[keys.sku] ?? "").trim() : "";
      
      // Prioritize Reference code as the primary grouping key
      const groupKey = (refVal || skuVal || nameVal || row.__id).trim().toLowerCase();
      const stockVal = keys.stock ? Number(row[keys.stock]) || 0 : 0;

      if (!groups[groupKey]) {
        groups[groupKey] = {
          key: groupKey,
          name: nameVal || (refVal ? `Ref: ${refVal}` : "Producto sin nombre"),
          ref: refVal || skuVal || nameVal,
          sku: skuVal,
          totalStock: 0,
          variants: [],
        };
      }

      groups[groupKey].variants.push(row);
      groups[groupKey].totalStock += stockVal;
    });

    return Object.values(groups);
  }, [inv.rows, keys]);

  // Filter grouped products
  const filteredProducts = useMemo(() => {
    const term = filter.trim().toLowerCase();
    const cleanTerm = term.replace(/[\s\-_]/g, "");

    return groupedProducts.filter((gp) => {
      if (showOnlySales) {
        const hasSales = gp.variants.some((v) => (inv.sales[v.__id] ?? 0) > 0);
        if (!hasSales) return false;
      }

      if (!term) return true;

      // Check product header (ref, name, sku) with both raw term and cleanTerm
      const gpRefLower = gp.ref.toLowerCase();
      const gpNameLower = gp.name.toLowerCase();
      const gpSkuLower = gp.sku.toLowerCase();

      const matchesHeader = 
        gpRefLower.includes(term) ||
        gpNameLower.includes(term) ||
        gpSkuLower.includes(term) ||
        (cleanTerm.length > 2 && (
          gpRefLower.replace(/[\s\-_]/g, "").includes(cleanTerm) ||
          gpNameLower.replace(/[\s\-_]/g, "").includes(cleanTerm)
        ));

      if (matchesHeader) return true;

      // Check if any variant cell matches search term
      return gp.variants.some((v) =>
        inv.columns.some((col) => {
          const val = String(v[col] ?? "").toLowerCase();
          return val.includes(term) || (cleanTerm.length > 2 && val.replace(/[\s\-_]/g, "").includes(cleanTerm));
        })
      );
    });
  }, [groupedProducts, filter, showOnlySales, inv.sales, inv.columns]);

  // Calculate totals
  const totalSoldUnits = Object.values(inv.sales).reduce((a, b) => a + b, 0);
  const productsWithSalesCount = Object.keys(inv.sales).length;

  const salesCartItems = useMemo(() => {
    const items: {
      row: InventoryRow;
      qty: number;
      name: string;
      ref: string;
      variantDesc: string;
      price: number;
      stock: number | null;
    }[] = [];

    inv.rows.forEach((row) => {
      const qty = inv.sales[row.__id] ?? 0;
      if (qty > 0) {
        const nameVal = keys.name ? String(row[keys.name] ?? "") : "";
        const refVal = keys.ref ? String(row[keys.ref] ?? "") : "";
        const colorVal = keys.color ? String(row[keys.color] ?? "") : "";
        const sizeVal = keys.size ? String(row[keys.size] ?? "") : "";
        const priceVal = keys.price ? parsePrice(row[keys.price]) : 0;
        const stockVal = keys.stock ? Number(row[keys.stock]) || 0 : null;

        const descParts = [];
        if (colorVal) descParts.push(translateColor(colorVal));
        if (sizeVal) descParts.push(`Talla ${sizeVal}`);

        items.push({
          row,
          qty,
          name: nameVal || refVal || "Producto",
          ref: refVal,
          variantDesc: descParts.join(" / ") || "Única",
          price: priceVal,
          stock: stockVal,
        });
      }
    });

    return items;
  }, [inv.rows, inv.sales, keys]);

  const totalEstimatedRevenue = useMemo(() => {
    return salesCartItems.reduce((sum, item) => sum + (item.qty * item.price), 0);
  }, [salesCartItems]);

  const commissionSummary = useMemo(() => {
    let totalQty = 0;
    let totalVentaConIva = 0;
    let totalVentaSinIva = 0;
    let totalComision = 0;
    let totalIvaComision = 0;
    let totalRetefuente = 0;
    let totalTotalComision = 0;
    let totalTotalReembolsar = 0;

    salesCartItems.forEach((item) => {
      const ventaConIva = item.qty * item.price;
      const ventaSinIva = Math.round(ventaConIva / 1.19);
      const comision = Math.round(ventaSinIva * (commPercent / 100));
      const ivaComision = Math.round(comision * 0.19);
      const retefuente = Math.round(comision * 0.11);
      const rowTotalComision = comision + ivaComision - retefuente;
      const rowTotalReembolsar = ventaConIva - rowTotalComision;

      totalQty += item.qty;
      totalVentaConIva += ventaConIva;
      totalVentaSinIva += ventaSinIva;
      totalComision += comision;
      totalIvaComision += ivaComision;
      totalRetefuente += retefuente;
      totalTotalComision += rowTotalComision;
      totalTotalReembolsar += rowTotalReembolsar;
    });

    return {
      commPercent,
      totalQty,
      totalVentaConIva,
      totalVentaSinIva,
      totalComision,
      totalIvaComision,
      totalRetefuente,
      totalTotalComision,
      totalTotalReembolsar,
    };
  }, [salesCartItems, commPercent]);

  // Pagination Helper calculations
  const itemsPerPage = 40;
  const totalPages = Math.ceil(filteredProducts.length / itemsPerPage);

  const paginatedProducts = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return filteredProducts.slice(startIndex, startIndex + itemsPerPage);
  }, [filteredProducts, currentPage]);

  const handleExport = () => {
    if (!keys.ref) {
      toast.error("No se encontró la columna de Referencia.");
      return;
    }
    if (salesCartItems.length === 0) {
      toast.error("No hay ventas para exportar.");
      return;
    }

    // Generate Excel data
    const exportItems = salesCartItems.map(item => ({
      ref: item.ref,
      qty: item.qty,
      size: keys.size ? String(item.row[keys.size] ?? "") : "",
      color: keys.color ? String(item.row[keys.color] ?? "") : "",
    }));

    // Generate and download
    generateSalesExcel(exportItems, warehouseId, user?.name || user?.email || "vendedor");

    // Add to export logs
    const loggedItems = salesCartItems.map(item => ({
      __id: item.row.__id,
      ref: item.ref,
      name: item.name,
      variantDesc: item.variantDesc,
      qty: item.qty,
      price: item.price,
      stock: item.stock,
      size: keys.size ? String(item.row[keys.size] ?? "") : "",
      color: keys.color ? String(item.row[keys.color] ?? "") : "",
    }));
    addExportLog(user?.email || "", warehouseId, commPercent, loggedItems);

    // Reset current sales
    if (user) {
      resetSales(user.email);
    }
    toast.success("Ventas exportadas y archivadas en el historial.");
  };

  const handleExportCommissions = () => {
    if (!keys.ref) {
      toast.error("No se encontró la columna de Referencia.");
      return;
    }
    if (salesCartItems.length === 0) {
      toast.error("No hay ventas para exportar.");
      return;
    }

    // Generate Excel data
    const exportItems = salesCartItems.map(item => ({
      ref: item.ref,
      name: item.name,
      size: keys.size ? String(item.row[keys.size] ?? "") : "",
      color: keys.color ? String(item.row[keys.color] ?? "") : "",
      qty: item.qty,
      price: item.price,
    }));

    // Generate and download
    generateCommissionsExcel(exportItems, commPercent, user?.name || user?.email || "vendedor");

    // Add to export logs
    const loggedItems = salesCartItems.map(item => ({
      __id: item.row.__id,
      ref: item.ref,
      name: item.name,
      variantDesc: item.variantDesc,
      qty: item.qty,
      price: item.price,
      stock: item.stock,
      size: keys.size ? String(item.row[keys.size] ?? "") : "",
      color: keys.color ? String(item.row[keys.color] ?? "") : "",
    }));
    addExportLog(user?.email || "", warehouseId, commPercent, loggedItems);

    // Reset current sales
    if (user) {
      resetSales(user.email);
    }
    toast.success("Liquidación exportada y archivada en el historial.");
  };

  const handleIncrement = (id: string, currentVal: number, stock: number | null) => {
    if (!user) return;
    const newVal = currentVal + 1;
    if (stock !== null && newVal > stock) {
      toast.error("Alerta: Estás reportando más unidades de las disponibles en stock.");
    }
    setSale(user.email, id, newVal);
  };

  const handleDecrement = (id: string, currentVal: number) => {
    if (!user) return;
    const newVal = Math.max(0, currentVal - 1);
    setSale(user.email, id, newVal);
  };

  const handleReset = () => {
    if (user) {
      resetSales(user.email);
      toast.success("Se han borrado todos los registros de ventas.");
      setIsResetConfirmOpen(false);
      setActiveTab("register");
    }
  };

  if (authLoading || !user) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background">
        <div className="flex flex-col items-center gap-2">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <p className="text-sm font-semibold">Verificando acceso...</p>
        </div>
      </div>
    );
  }

  if (inv.rows.length === 0) {
    return (
      <div className="min-h-screen bg-neutral-50/50 dark:bg-background">
        <Header />
        <main className="mx-auto max-w-md px-4 py-16 text-center">
          <div className="bg-white dark:bg-card p-6 rounded-2xl border border-border shadow-md space-y-4">
            <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-full bg-neutral-100 dark:bg-muted text-muted-foreground">
              <Package className="h-7 w-7" />
            </div>
            <h3 className="text-lg font-bold text-foreground">Sin inventario asignado</h3>
            <p className="text-sm text-muted-foreground leading-relaxed">
              Hola <strong>{user.name}</strong>, el administrador aún no ha cargado los productos para tu cuenta. 
              Por favor contáctalo para poder ingresar tus ventas.
            </p>
          </div>
        </main>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-neutral-100 dark:bg-neutral-950 pb-28 animate-in fade-in duration-200">
      <Header />

      {/* TABS NAVIGATION */}
      <div className="sticky top-14 z-40 bg-white dark:bg-neutral-900 border-b border-border shadow-3xs">
        <div className="max-w-2xl mx-auto flex">
          <button
            onClick={() => setActiveTab("register")}
            className={`flex-1 py-4 text-center font-bold text-sm border-b-2 transition-all flex items-center justify-center gap-2 ${
              activeTab === "register"
                ? "border-primary text-primary"
                : "border-transparent text-muted-foreground hover:text-foreground"
            }`}
          >
            <Shirt className="h-4.5 w-4.5" />
            Catálogo
          </button>
          <button
            onClick={() => setActiveTab("report")}
            className={`flex-1 py-4 text-center font-bold text-sm border-b-2 transition-all flex items-center justify-center gap-2 relative ${
              activeTab === "report"
                ? "border-primary text-primary"
                : "border-transparent text-muted-foreground hover:text-foreground"
            }`}
          >
            <ShoppingBag className="h-4.5 w-4.5" />
            Ver Reporte
            {productsWithSalesCount > 0 && (
              <span className="absolute top-2.5 right-6 md:right-12 h-5 min-w-[20px] px-1.5 flex items-center justify-center text-3xs font-black bg-emerald-500 text-white rounded-full animate-pulse">
                {productsWithSalesCount}
              </span>
            )}
          </button>
        </div>
      </div>

      <main className="max-w-7xl mx-auto px-4 pt-6 space-y-6">

        {/* TAB 1: E-COMMERCE REGISTRATION VIEW */}
        {activeTab === "register" && (
          <div className="space-y-6">
            
            {/* SEARCH AND FILTERS BAR */}
            <div className="bg-white dark:bg-neutral-900 p-4 rounded-2xl border border-border/80 shadow-2xs space-y-3.5">
              <div className="relative">
                <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 h-4.5 w-4.5 text-muted-foreground" />
                <Input
                  placeholder="Buscar prendas por nombre, color, referencia..."
                  value={filter}
                  onChange={(e) => setFilter(e.target.value)}
                  className="pl-10 h-11 bg-neutral-50 dark:bg-neutral-950 border-border/80 rounded-xl"
                />
                {filter && (
                  <button
                    onClick={() => setFilter("")}
                    className="absolute right-3.5 top-1/2 -translate-y-1/2 text-xs font-bold text-muted-foreground hover:text-foreground bg-neutral-200 dark:bg-muted p-1 px-2.5 rounded-lg"
                  >
                    Borrar
                  </button>
                )}
              </div>

              <div className="flex items-center justify-between pt-1">
                <span className="text-2xs text-muted-foreground font-semibold uppercase tracking-wider">Filtrar prendas:</span>
                <button
                  onClick={() => setShowOnlySales(!showOnlySales)}
                  className={`text-xs font-bold px-3 py-1.5 rounded-full border transition-all ${
                    showOnlySales
                      ? "bg-primary/10 border-primary/30 text-primary"
                      : "bg-background border-border text-muted-foreground hover:text-foreground"
                  }`}
                >
                  {showOnlySales ? "✓ Viendo prendas con ventas" : "Ver prendas con ventas"}
                </button>
              </div>
            </div>

            {/* VIRTUAL STORE PRODUCT GRID */}
            {filteredProducts.length === 0 ? (
              <div className="py-16 text-center bg-white dark:bg-neutral-900 rounded-2xl border border-border/80 p-6 space-y-3">
                <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-full bg-neutral-100 dark:bg-muted text-muted-foreground">
                  <Search className="h-6 w-6" />
                </div>
                <h3 className="font-bold text-base">No hay productos coincidentes</h3>
                <p className="text-xs text-muted-foreground max-w-sm mx-auto">
                  Prueba cambiando la descripción en el buscador o desactiva el filtro de modificados.
                </p>
              </div>
            ) : (
              <>
                <div className="grid gap-6 grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4">
                  {paginatedProducts.map((product) => {
                    return (
                      <ProductVirtualCard
                        key={product.key}
                        product={product}
                        keys={keys}
                        inv={inv}
                        user={user}
                        selectedVariants={selectedVariants}
                        setSelectedVariants={setSelectedVariants}
                        handleIncrement={handleIncrement}
                        handleDecrement={handleDecrement}
                      />
                    );
                  })}
                </div>

                {/* PAGINATION CONTROLS */}
                {totalPages > 1 && (
                  <div className="flex items-center justify-center gap-2 pt-6 pb-2 animate-in fade-in duration-300">
                    <Button
                      variant="outline"
                      size="icon"
                      className="h-9 w-9 rounded-xl border-border bg-card text-foreground hover:bg-accent disabled:opacity-50"
                      onClick={() => setCurrentPage((prev) => Math.max(1, prev - 1))}
                      disabled={currentPage === 1}
                    >
                      <ChevronLeft className="h-4 w-4" />
                    </Button>

                    {Array.from({ length: totalPages }, (_, i) => i + 1).map((page) => {
                      const isCurrent = page === currentPage;
                      const isNear = Math.abs(page - currentPage) <= 1;
                      const isFirstOrLast = page === 1 || page === totalPages;

                      if (!isNear && !isFirstOrLast) {
                        if (page === 2 || page === totalPages - 1) {
                          return <span key={page} className="px-1 text-muted-foreground text-xs font-semibold">...</span>;
                        }
                        return null;
                      }

                      return (
                        <Button
                          key={page}
                          variant={isCurrent ? "default" : "outline"}
                          className={`h-9 w-9 text-xs font-bold rounded-xl ${
                            isCurrent
                              ? "bg-primary text-primary-foreground shadow-xs font-black"
                              : "border-border bg-card text-muted-foreground hover:text-foreground"
                          }`}
                          onClick={() => setCurrentPage(page)}
                        >
                          {page}
                        </Button>
                      );
                    })}

                    <Button
                      variant="outline"
                      size="icon"
                      className="h-9 w-9 rounded-xl border-border bg-card text-foreground hover:bg-accent disabled:opacity-50"
                      onClick={() => setCurrentPage((prev) => Math.min(totalPages, prev + 1))}
                      disabled={currentPage === totalPages}
                    >
                      <ChevronRight className="h-4 w-4" />
                    </Button>
                  </div>
                )}
              </>
            )}

          </div>
        )}

        {/* TAB 2: DETAILED SALES REPORT VIEW */}
        {activeTab === "report" && (
          <div className="max-w-2xl mx-auto space-y-4">
            
            <Card className="border-border/80 shadow-md bg-white dark:bg-neutral-900 rounded-2xl overflow-hidden">
              <CardHeader className="bg-neutral-50 dark:bg-muted/15 border-b border-border p-4 flex flex-row items-center justify-between">
                <div>
                  <CardTitle className="text-sm font-bold">Reporte Consolidado</CardTitle>
                  <CardDescription className="text-3xs">Resumen de ventas listas para exportar</CardDescription>
                </div>
                <span className="text-xs font-black bg-primary/10 text-primary px-3 py-1 rounded-full border border-primary/20">
                  {totalSoldUnits} unidades
                </span>
              </CardHeader>

              <CardContent className="p-4 space-y-4">
                {salesCartItems.length === 0 ? (
                  <div className="py-12 flex flex-col items-center justify-center text-center space-y-3">
                    <ShoppingBag className="h-10 w-10 text-muted-foreground/30 stroke-1" />
                    <p className="text-sm text-muted-foreground font-semibold">Tu reporte está vacío</p>
                    <p className="text-xs text-muted-foreground max-w-[240px]">
                      Aún no has agregado ventas en tu catálogo. Regresa al catálogo y suma unidades.
                    </p>
                    <Button 
                      variant="outline" 
                      onClick={() => setActiveTab("register")}
                      className="mt-2 text-xs font-bold rounded-xl"
                    >
                      Ver Catálogo
                    </Button>
                  </div>
                ) : (
                  <>
                    <div className="divide-y divide-border/50 max-h-[350px] overflow-y-auto pr-1">
                      {salesCartItems.map((item) => {
                        const over = item.stock !== null && item.qty > item.stock;
                        return (
                          <div key={item.row.__id} className="py-3 flex items-center justify-between gap-3">
                            <div className="min-w-0 flex-1">
                              <p className="font-bold text-xs text-foreground truncate">{item.name}</p>
                              <div className="flex items-center gap-1.5 flex-wrap mt-0.5">
                                {item.ref && <span className="font-mono text-3xs text-muted-foreground">REF: {item.ref}</span>}
                                {item.variantDesc && (
                                  <>
                                    <span className="h-0.5 w-0.5 rounded-full bg-muted-foreground/40" />
                                    <span className="text-3xs text-muted-foreground truncate">{item.variantDesc}</span>
                                  </>
                                )}
                              </div>
                              {keys.price && item.price > 0 && (
                                <p className="text-3xs text-emerald-600 dark:text-emerald-400 font-medium">
                                  {item.qty} × ${item.price.toLocaleString()}
                                </p>
                              )}
                            </div>

                            <div className="flex items-center gap-3 shrink-0">
                              <div className="text-right">
                                <span className={`font-mono text-xs font-black px-2.5 py-1 rounded-lg border ${
                                  over
                                    ? "bg-red-50 text-red-600 border-red-200 animate-pulse"
                                    : "bg-neutral-50 dark:bg-muted text-foreground border-border"
                                }`}>
                                  {item.qty} ud
                                </span>
                                {keys.price && item.price > 0 && (
                                  <p className="text-3xs font-bold text-foreground mt-1">
                                    ${(item.qty * item.price).toLocaleString()}
                                  </p>
                                )}
                              </div>
                              <button
                                onClick={() => {
                                  if (user) setSale(user.email, item.row.__id, 0);
                                  toast.success("Producto removido del reporte");
                                }}
                                className="p-1.5 text-muted-foreground hover:text-destructive hover:bg-destructive/5 rounded-lg transition-colors"
                                title="Eliminar registro"
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    <div className="pt-4 border-t border-border space-y-2">
                      <div className="flex justify-between text-xs font-medium text-muted-foreground">
                        <span>Referencias Únicas:</span>
                        <span className="text-foreground font-bold">{productsWithSalesCount}</span>
                      </div>
                      {keys.price && (
                        <div className="flex justify-between text-xs font-medium text-muted-foreground">
                          <span>Subtotal Estimado:</span>
                          <span className="text-emerald-600 dark:text-emerald-400 font-extrabold text-sm">
                            ${totalEstimatedRevenue.toLocaleString()}
                          </span>
                        </div>
                      )}
                      
                      {keys.price && commPercent > 0 && (
                        <div className="pt-3 border-t border-dashed border-border mt-3 space-y-2">
                          <p className="text-2xs font-extrabold text-indigo-600 dark:text-indigo-400 uppercase tracking-wider">
                            Liquidación de Comisiones ({commPercent}%)
                          </p>
                          <div className="grid grid-cols-2 gap-2 text-2xs">
                            <div className="flex justify-between text-muted-foreground">
                              <span>Venta Sin IVA (base):</span>
                              <span className="font-semibold text-foreground">
                                ${commissionSummary.totalVentaSinIva.toLocaleString()}
                              </span>
                            </div>
                            <div className="flex justify-between text-muted-foreground">
                              <span>Comisión base:</span>
                              <span className="font-semibold text-foreground">
                                ${commissionSummary.totalComision.toLocaleString()}
                              </span>
                            </div>
                            <div className="flex justify-between text-muted-foreground">
                              <span>IVA Comisión (19%):</span>
                              <span className="font-semibold text-foreground">
                                +${commissionSummary.totalIvaComision.toLocaleString()}
                              </span>
                            </div>
                            <div className="flex justify-between text-muted-foreground">
                              <span>Retefuente (11%):</span>
                              <span className="font-semibold text-destructive">
                                -${commissionSummary.totalRetefuente.toLocaleString()}
                              </span>
                            </div>
                          </div>
                          
                          <div className="pt-2 border-t border-border/50 flex justify-between text-xs font-bold">
                            <span className="text-indigo-600 dark:text-indigo-400">Total Comisión Ganada:</span>
                            <span className="text-indigo-600 dark:text-indigo-400">
                              ${commissionSummary.totalTotalComision.toLocaleString()}
                            </span>
                          </div>
                          
                          <div className="flex justify-between text-xs font-bold">
                            <span className="text-foreground">Total a Reembolsar a Marca:</span>
                            <span className="text-foreground">
                              ${commissionSummary.totalTotalReembolsar.toLocaleString()}
                            </span>
                          </div>
                        </div>
                      )}
                      
                      {salesCartItems.some(item => item.stock !== null && item.qty > item.stock) && (
                        <div className="flex items-center gap-2 p-2.5 bg-red-50 dark:bg-red-950/20 border border-red-200 dark:border-red-900/30 rounded-xl text-red-600 dark:text-red-400 text-3xs font-semibold">
                          <AlertTriangle className="h-4 w-4 shrink-0" />
                          <span>Atención: Algunas cantidades superan el stock físico disponible.</span>
                        </div>
                      )}
                    </div>

                    <div className="space-y-2 pt-2">
                      <Button
                        onClick={handleExport}
                        className="w-full h-12 font-bold text-sm gap-2 rounded-xl shadow-sm bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer"
                      >
                        <FileSpreadsheet className="h-5 w-5" />
                        Descargar Reporte de Ventas (.xls)
                      </Button>

                      {commPercent > 0 && (
                        <Button
                          onClick={handleExportCommissions}
                          className="w-full h-12 font-bold text-sm gap-2 rounded-xl shadow-sm bg-indigo-600 hover:bg-indigo-700 text-white cursor-pointer"
                        >
                          <FileSpreadsheet className="h-5 w-5" />
                          Descargar Liquidación de Comisiones (.xls)
                        </Button>
                      )}

                      {isResetConfirmOpen ? (
                        <div className="p-3 rounded-xl border border-red-200 dark:border-red-950 bg-red-50/40 dark:bg-red-950/10 flex items-center justify-between">
                          <span className="text-2xs font-bold text-red-700 dark:text-red-400">¿Estás seguro de borrar todas las ventas?</span>
                          <div className="flex gap-2">
                            <Button
                              variant="destructive"
                              size="sm"
                              onClick={handleReset}
                              className="text-3xs font-extrabold px-3 h-8 rounded-lg"
                            >
                              Sí, Borrar
                            </Button>
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() => setIsResetConfirmOpen(false)}
                              className="text-3xs font-bold px-3 h-8 rounded-lg"
                            >
                              Cancelar
                            </Button>
                          </div>
                        </div>
                      ) : (
                        <Button
                          variant="ghost"
                          onClick={() => setIsResetConfirmOpen(true)}
                          className="w-full text-xs font-bold text-muted-foreground hover:text-red-500 hover:bg-red-500/5 h-10 rounded-xl"
                        >
                          <RefreshCw className="h-3.5 w-3.5 mr-1.5" />
                          Limpiar todo el reporte
                        </Button>
                      )}
                    </div>
                  </>
                )}
              </CardContent>
            </Card>

            {/* HISTORIAL DE EXPORTACIONES */}
            <Card className="border-border/80 shadow-md bg-white dark:bg-neutral-900 rounded-2xl overflow-hidden">
              <CardHeader className="bg-neutral-50 dark:bg-muted/15 border-b border-border p-4 flex flex-row items-center justify-between">
                <div>
                  <CardTitle className="text-sm font-bold flex items-center gap-1.5">
                    <History className="h-4.5 w-4.5 text-primary" />
                    Historial de Exportaciones
                  </CardTitle>
                  <CardDescription className="text-3xs">Registro de reportes descargados y archivados</CardDescription>
                </div>
              </CardHeader>
              <CardContent className="p-4">
                {myLogs.length === 0 ? (
                  <p className="text-xs text-muted-foreground text-center py-6">No hay reportes exportados en el historial.</p>
                ) : (
                  <div className="divide-y divide-border/50 max-h-[300px] overflow-y-auto pr-1">
                    {myLogs.map((log) => (
                      <div key={log.id} className="py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                        <div className="space-y-1">
                          <p className="font-bold text-foreground">
                            {new Date(log.timestamp).toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" })}
                          </p>
                          <p className="text-3xs text-muted-foreground">
                            <strong>{log.totalUnits}</strong> {log.totalUnits === 1 ? "unidad" : "unidades"} · 
                            Valor: <strong>${log.totalAmount.toLocaleString("es-CO")}</strong>
                            {log.commissionPercent > 0 && ` · Comisión: ${log.commissionPercent}%`}
                          </p>
                        </div>
                        <div className="flex gap-2 shrink-0">
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => handleDownloadSalesFromLog(log)}
                            className="h-8 text-3xs font-bold gap-1 rounded-lg border-emerald-600/30 text-emerald-700 hover:bg-emerald-50/50 dark:hover:bg-emerald-950/20"
                          >
                            <FileSpreadsheet className="h-3.5 w-3.5" />
                            Ventas
                          </Button>
                          {log.commissionPercent > 0 && (
                            <Button
                              size="sm"
                              variant="outline"
                              onClick={() => handleDownloadCommissionsFromLog(log)}
                              className="h-8 text-3xs font-bold gap-1 rounded-lg border-indigo-600/30 text-indigo-700 hover:bg-indigo-50/50 dark:hover:bg-indigo-950/20"
                            >
                              <FileSpreadsheet className="h-3.5 w-3.5" />
                              Comisiones
                            </Button>
                          )}
                          <Button
                            size="icon"
                            variant="ghost"
                            onClick={() => handleDeleteLog(log.id)}
                            className="h-8 w-8 text-muted-foreground hover:text-red-500 hover:bg-red-500/5 rounded-lg"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </Button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </CardContent>
            </Card>

            <div className="bg-neutral-100 dark:bg-neutral-900/60 p-4 rounded-2xl flex gap-3 text-xs border border-border/40">
              <Info className="h-5 w-5 text-primary shrink-0 mt-0.5" />
              <div className="space-y-1">
                <p className="font-bold text-foreground">Instrucciones de Carga</p>
                <p className="text-muted-foreground leading-relaxed">
                  El archivo Excel generado contiene las columnas exactas requeridas para cargarlo en el sistema administrativo del administrador. No modifiques el nombre de las columnas.
                </p>
              </div>
            </div>

          </div>
        )}

      </main>

      {/* STICKY BOTTOM CHECKOUT BAR (Only on Tab 1 & when there are items) */}
      {activeTab === "register" && totalSoldUnits > 0 && (
        <div className="fixed bottom-0 left-0 right-0 z-50 bg-white/95 dark:bg-neutral-900/95 backdrop-blur-md border-t border-border p-4 shadow-lg animate-in slide-in-from-bottom duration-250">
          <div className="max-w-md mx-auto flex items-center justify-between gap-4">
            <div className="truncate">
              <p className="text-xs font-extrabold text-foreground">
                🛒 {productsWithSalesCount} {productsWithSalesCount === 1 ? "referencia" : "referencias"} en reporte
              </p>
              <p className="text-3xs text-muted-foreground">
                Total unidades reportadas: <strong>{totalSoldUnits}</strong>
              </p>
            </div>
            
            <Button
              onClick={() => setActiveTab("report")}
              className="font-bold text-xs gap-1.5 py-5 px-4 rounded-xl bg-primary hover:bg-primary/95 text-primary-foreground shadow-xs cursor-pointer shrink-0 animate-bounce-subtle"
            >
              Ver Reporte
              <ShoppingBag className="h-4.5 w-4.5" />
            </Button>
          </div>
        </div>
      )}

    </div>
  );
}

// REDESIGNED VIRTUAL STORE CARD COMPONENT
function ProductVirtualCard({
  product,
  keys,
  inv,
  user,
  selectedVariants,
  setSelectedVariants,
  handleIncrement,
  handleDecrement,
}: {
  product: GroupedProduct;
  keys: ReturnType<typeof detectKeyColumns>;
  inv: any;
  user: any;
  selectedVariants: Record<string, { color: string; size: string }>;
  setSelectedVariants: React.Dispatch<React.SetStateAction<Record<string, { color: string; size: string }>>>;
  handleIncrement: (id: string, currentVal: number, stock: number | null) => void;
  handleDecrement: (id: string, currentVal: number) => void;
}) {
  // Extract unique colors and sizes for selector (keeping empty values as selectable options)
  const uniqueColors = useMemo(() => {
    return Array.from(new Set(product.variants.map((v) => keys.color ? String(v[keys.color] ?? "").trim() : "")));
  }, [product.variants, keys.color]);

  const uniqueSizes = useMemo(() => {
    return Array.from(new Set(product.variants.map((v) => keys.size ? String(v[keys.size] ?? "").trim() : "")));
  }, [product.variants, keys.size]);

  // Current selections
  const currentSelection = selectedVariants[product.key] || {
    color: uniqueColors[0] || "",
    size: uniqueSizes[0] || "",
  };

  const activeColor = currentSelection.color;
  const activeSize = currentSelection.size;

  // Active Variant row based on selected color and size
  const activeVariant = useMemo(() => {
    return product.variants.find((v) => {
      const matchColor = !keys.color || String(v[keys.color] ?? "") === activeColor;
      const matchSize = !keys.size || String(v[keys.size] ?? "") === activeSize;
      return matchColor && matchSize;
    }) || product.variants[0];
  }, [product.variants, keys.color, keys.size, activeColor, activeSize]);

  // Specific data for the active variant
  const soldQty = activeVariant ? (inv.sales[activeVariant.__id] ?? 0) : 0;
  const stock = activeVariant && keys.stock ? Number(activeVariant[keys.stock]) || 0 : null;
  const over = stock !== null && soldQty > stock;
  const priceVal = activeVariant && keys.price ? parsePrice(activeVariant[keys.price]) : 0;

  // Cumulative quantities for the visual indicator on the card
  const totalSoldInProduct = useMemo(() => {
    return product.variants.reduce((sum, v) => sum + (inv.sales[v.__id] ?? 0), 0);
  }, [product.variants, inv.sales]);

  const handleColorSelect = (color: string) => {
    setSelectedVariants((prev) => {
      // Find a size that is available for this color in the stock
      const availableSizesForColor = product.variants
        .filter((v) => !keys.color || String(v[keys.color] ?? "") === color)
        .map((v) => keys.size ? String(v[keys.size] ?? "") : "");
      
      const nextSize = availableSizesForColor.includes(activeSize)
        ? activeSize
        : availableSizesForColor[0] || "";

      return {
        ...prev,
        [product.key]: { color, size: nextSize }
      };
    });
  };

  const handleSizeSelect = (size: string) => {
    setSelectedVariants((prev) => {
      const availableColorsForSize = product.variants
        .filter((v) => !keys.size || String(v[keys.size] ?? "") === size)
        .map((v) => keys.color ? String(v[keys.color] ?? "") : "");

      const nextColor = availableColorsForSize.includes(activeColor)
        ? activeColor
        : availableColorsForSize[0] || "";

      return {
        ...prev,
        [product.key]: { color: nextColor, size }
      };
    });
  };

  return (
    <div
      className={`bg-white dark:bg-neutral-900 rounded-3xl border transition-all duration-200 overflow-hidden flex flex-col shadow-sm relative hover:shadow-md ${
        totalSoldInProduct > 0
          ? "border-emerald-500 ring-2 ring-emerald-500/10 shadow-emerald-500/5 bg-emerald-50/2 dark:bg-emerald-950/1"
          : "border-border/80"
      }`}
    >
      {/* VIRTUAL STORE PRODUCT IMAGE CONTAINER (CSS Gradient mock) */}
      <div className="h-32 relative flex items-center justify-center overflow-hidden bg-gradient-to-tr from-neutral-100 to-neutral-200 dark:from-neutral-800 dark:to-neutral-900 border-b border-border/55">
        <div className="absolute inset-0 bg-radial-gradient opacity-30" />
        
        {/* Category apparel icon */}
        <div className="p-4 bg-white/70 dark:bg-neutral-950/70 backdrop-blur-md rounded-2xl shadow-sm text-neutral-600 dark:text-neutral-300">
          <Shirt className="h-8 w-8 stroke-[1.5]" />
        </div>

        {/* Floating price badge */}
        {priceVal > 0 && (
          <span className="absolute bottom-3 right-3 bg-emerald-600 text-white font-extrabold text-xs px-3 py-1 rounded-xl shadow-xs">
            $ {priceVal.toLocaleString()}
          </span>
        )}

        {/* Floating ref badge */}
        {product.ref && (
          <span className="absolute top-3 left-3 bg-white/90 dark:bg-neutral-950/90 text-muted-foreground font-mono text-3xs font-bold px-2 py-0.5 rounded-lg border border-border shadow-3xs">
            REF: {product.ref}
          </span>
        )}

        {/* Floating shopping indicator if units are sold */}
        {totalSoldInProduct > 0 && (
          <span className="absolute top-3 right-3 bg-emerald-500 text-white font-black text-2xs px-2.5 py-1 rounded-xl flex items-center gap-1 shadow-sm">
            <CheckCircle2 className="h-3.5 w-3.5" />
            Reportado: {totalSoldInProduct} ud
          </span>
        )}
      </div>

      {/* CARD CONTENT */}
      <div className="p-4 flex-1 flex flex-col justify-between space-y-4">
        
        {/* Product Title & Ref */}
        <div className="space-y-0.5">
          <h4 className="font-extrabold text-sm text-foreground tracking-tight line-clamp-2" title={product.name}>
            {product.name}
          </h4>
          {product.ref && product.ref !== product.name && (
            <p className="text-3xs font-mono font-semibold text-muted-foreground">
              Ref: <span className="text-foreground font-bold">{product.ref}</span>
            </p>
          )}
        </div>

        {/* SELECTORS WRAPPER */}
        <div className="space-y-3.5">
          
          {/* COLOR SELECTOR */}
          {uniqueColors.length > 0 && (
            <div className="space-y-1">
              <p className="text-3xs text-muted-foreground font-bold uppercase tracking-wider">
                Color: <span className="text-foreground normal-case font-black">{translateColor(activeColor)}</span>
              </p>
              {uniqueColors.length > 1 && (
                <div className="flex flex-wrap gap-1.5 pt-0.5">
                  {uniqueColors.map((color) => {
                    const isActive = color === activeColor;
                    return (
                      <button
                        key={color}
                        onClick={() => handleColorSelect(color)}
                        className={`text-3xs px-2.5 py-1 rounded-full border font-bold transition-all truncate max-w-[120px] ${
                          isActive
                            ? "bg-primary/10 border-primary/45 text-primary"
                            : "bg-background border-border text-muted-foreground hover:text-foreground"
                        }`}
                        type="button"
                      >
                        {translateColor(color)}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {/* SIZE SELECTOR */}
          {uniqueSizes.length > 0 && (
            <div className="space-y-1">
              <p className="text-3xs text-muted-foreground font-bold uppercase tracking-wider">
                Talla seleccionada: <span className="text-foreground normal-case font-black">{activeSize || "Única"}</span>
              </p>
              {uniqueSizes.length > 1 && (
                <div className="flex flex-wrap gap-1.5 pt-0.5">
                  {uniqueSizes.map((size) => {
                    const isActive = size === activeSize;
                    return (
                      <button
                        key={size}
                        onClick={() => handleSizeSelect(size)}
                        className={`w-9 h-9 rounded-xl border text-2xs font-extrabold transition-all flex items-center justify-center ${
                          isActive
                            ? "bg-primary border-primary text-primary-foreground shadow-2xs font-black"
                            : "bg-background border-border text-muted-foreground hover:text-foreground"
                        }`}
                        type="button"
                      >
                        {size || "Única"}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          )}

        </div>

        {/* DYNAMIC STOCK & INPUT ADJUSTER */}
        <div className="pt-3 border-t border-border/50 flex items-center justify-between gap-2">
          
          {/* Stock Display */}
          <div className="min-w-0">
            {keys.stock && (
              <>
                <p className="text-3xs text-muted-foreground font-medium">Disponible</p>
                <p className={`text-xs font-black ${
                  stock === 0 
                    ? "text-red-500" 
                    : stock !== null && stock < 5 
                      ? "text-amber-600 dark:text-amber-400" 
                      : "text-foreground"
                }`}>
                  {stock === 0 ? "Agotado" : `${stock} unidades`}
                </p>
              </>
            )}
          </div>

          {/* Quantity Controls */}
          {activeVariant ? (
            <div className="flex flex-col items-end gap-1">
              <div className="flex items-center border border-border rounded-xl overflow-hidden bg-background shadow-3xs">
                <button
                  onClick={() => handleDecrement(activeVariant.__id, soldQty)}
                  disabled={soldQty <= 0}
                  className="w-10 h-10 flex items-center justify-center hover:bg-neutral-50 dark:hover:bg-muted text-muted-foreground hover:text-foreground disabled:opacity-30 transition-colors"
                  type="button"
                >
                  <Minus className="h-4 w-4 stroke-[2.5]" />
                </button>
                
                <input
                  type="number"
                  min={0}
                  value={soldQty || ""}
                  onChange={(e) => {
                    if (user) setSale(user.email, activeVariant.__id, Number(e.target.value) || 0);
                  }}
                  className={`w-10 h-10 text-center text-xs font-black focus:outline-hidden border-x border-border/30 ${
                    over ? "text-destructive bg-destructive/5 animate-pulse font-black" : "text-foreground"
                  }`}
                  placeholder="0"
                  disabled={stock === 0 && soldQty === 0}
                />

                <button
                  onClick={() => handleIncrement(activeVariant.__id, soldQty, stock)}
                  disabled={stock === 0}
                  className="w-10 h-10 flex items-center justify-center hover:bg-neutral-50 dark:hover:bg-muted text-muted-foreground hover:text-foreground disabled:opacity-30 transition-colors"
                  type="button"
                >
                  <Plus className="h-4 w-4 stroke-[2.5]" />
                </button>
              </div>

              {over && (
                <span className="text-3xs text-destructive font-black flex items-center gap-0.5 animate-pulse">
                  <AlertTriangle className="h-3 w-3" /> Excede stock
                </span>
              )}
            </div>
          ) : (
            <span className="text-3xs text-muted-foreground italic">No disponible</span>
          )}

        </div>

      </div>
    </div>
  );
}

function Header() {
  const { user, logout } = useAuth();

  return (
    <header className="border-b border-border bg-card sticky top-0 z-45 shadow-3xs">
      <div className="max-w-2xl mx-auto flex items-center justify-between px-4 py-3">
        <div className="flex flex-col">
          <h1 className="text-sm font-black text-foreground tracking-tight flex items-center gap-1.5">
            <ShoppingBag className="h-4.5 w-4.5 text-primary" />
            Reporte Ventas
          </h1>
          {user && (
            <span className="text-3xs text-muted-foreground font-semibold">
              Proveedor: <span className="text-primary font-bold">{user.name}</span>
            </span>
          )}
        </div>
        
        <div className="flex items-center gap-2">
          {user?.role === "admin" && (
            <Link
              to="/"
              className="rounded-xl border border-border bg-background px-3 py-1.5 text-3xs font-extrabold text-foreground hover:bg-accent transition-colors"
            >
              Panel Admin
            </Link>
          )}
          <Button
            variant="outline"
            size="sm"
            onClick={() => {
              logout();
              toast.success("Sesión cerrada");
            }}
            className="border-destructive/15 text-destructive hover:bg-destructive/10 hover:text-destructive transition-all h-8 px-2.5 rounded-xl text-3xs font-extrabold"
          >
            <LogOut className="h-3.5 w-3.5" />
            Salir
          </Button>
        </div>
      </div>
    </header>
  );
}