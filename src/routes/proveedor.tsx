import { createFileRoute, Link } from "@tanstack/react-router";
import { useMemo, useState } from "react";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { detectKeyColumns, resetSales, setSale, useInventory } from "@/lib/inventory-store";

export const Route = createFileRoute("/proveedor")({
  head: () => ({
    meta: [
      { title: "Ventas del mes | Proveedor" },
      {
        name: "description",
        content: "Registra las unidades vendidas de cada producto durante el mes.",
      },
    ],
  }),
  component: ProveedorPage,
});

function ProveedorPage() {
  const inv = useInventory();
  const [filter, setFilter] = useState("");

  const keys = useMemo(() => detectKeyColumns(inv.columns), [inv.columns]);

  const rows = useMemo(() => {
    if (!filter.trim()) return inv.rows;
    const f = filter.toLowerCase();
    return inv.rows.filter((r) =>
      inv.columns.some((c) => String(r[c] ?? "").toLowerCase().includes(f)),
    );
  }, [inv.rows, inv.columns, filter]);

  const totalSoldUnits = Object.values(inv.sales).reduce((a, b) => a + b, 0);
  const productsWithSales = Object.keys(inv.sales).length;

  const handleExport = () => {
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
    for (const r of inv.rows) {
      const qty = inv.sales[r.__id];
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
    // Force text format for string columns to preserve leading zeros
    const range = XLSX.utils.decode_range(ws["!ref"]!);
    const textCols = [0, 3, 4, 5]; // StrProducto, IntBodega, StrLote, StrColor
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
    XLSX.writeFile(wb, `CARGA_COMODATOS_${stamp}.xls`, { bookType: "biff8" });
    toast.success(`Archivo exportado con ${data.length - 1} líneas.`);
  };

  if (inv.rows.length === 0) {
    return (
      <div className="min-h-screen bg-background">
        <Header />
        <main className="mx-auto max-w-3xl px-6 py-16 text-center">
          <Card>
            <CardHeader>
              <CardTitle>No hay inventario disponible</CardTitle>
              <CardDescription>
                El administrador todavía no ha subido el inventario del mes.
              </CardDescription>
            </CardHeader>
          </Card>
        </main>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background">
      <Header />
      <main className="mx-auto max-w-6xl space-y-6 px-6 py-8">
        <div className="grid gap-4 sm:grid-cols-3">
          <Stat label="Productos" value={inv.rows.length} />
          <Stat label="Con ventas" value={productsWithSales} />
          <Stat label="Unidades vendidas" value={totalSoldUnits} />
        </div>

        <Card>
          <CardHeader className="flex flex-row items-start justify-between gap-4">
            <div>
              <CardTitle>Registrar ventas del mes</CardTitle>
              <CardDescription>
                Ingresa las unidades vendidas para cada producto. Se guarda automáticamente.
              </CardDescription>
            </div>
            <div className="flex gap-2">
              <Button
                variant="ghost"
                onClick={() => {
                  resetSales();
                  toast.success("Ventas reiniciadas");
                }}
              >
                Reiniciar
              </Button>
              <Button onClick={handleExport} disabled={productsWithSales === 0}>
                Exportar archivo
              </Button>
            </div>
          </CardHeader>
          <CardContent className="space-y-4">
            <Input
              placeholder="Buscar producto…"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              className="max-w-sm"
            />
            <div className="overflow-auto rounded-md border border-border">
              <table className="w-full text-sm">
                <thead className="bg-muted text-muted-foreground">
                  <tr>
                    {keys.ref && <th className="whitespace-nowrap px-3 py-2 text-left font-medium">{keys.ref}</th>}
                    {keys.sku && keys.sku !== keys.ref && <th className="whitespace-nowrap px-3 py-2 text-left font-medium">{keys.sku}</th>}
                    {keys.name && <th className="whitespace-nowrap px-3 py-2 text-left font-medium">{keys.name}</th>}
                    {keys.color && <th className="whitespace-nowrap px-3 py-2 text-left font-medium">{keys.color}</th>}
                    {keys.size && <th className="whitespace-nowrap px-3 py-2 text-left font-medium">{keys.size}</th>}
                    {!keys.sku && !keys.name && !keys.ref && !keys.color && !keys.size && inv.columns.slice(0, 2).map((c) => (
                      <th key={c} className="whitespace-nowrap px-3 py-2 text-left font-medium">{c}</th>
                    ))}
                    {keys.stock && <th className="whitespace-nowrap px-3 py-2 text-right font-medium">{keys.stock}</th>}
                    <th className="whitespace-nowrap px-3 py-2 text-right font-medium">Unidades vendidas</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r) => {
                    const sold = inv.sales[r.__id] ?? 0;
                    const stock = keys.stock ? Number(r[keys.stock]) || 0 : null;
                    const over = stock !== null && sold > stock;
                    return (
                      <tr key={r.__id} className="border-t border-border">
                        {keys.ref && <td className="whitespace-nowrap px-3 py-2 font-mono text-xs">{String(r[keys.ref] ?? "")}</td>}
                        {keys.sku && keys.sku !== keys.ref && <td className="whitespace-nowrap px-3 py-2 font-mono text-xs">{String(r[keys.sku] ?? "")}</td>}
                        {keys.name && <td className="px-3 py-2">{String(r[keys.name] ?? "")}</td>}
                        {keys.color && <td className="whitespace-nowrap px-3 py-2">{String(r[keys.color] ?? "")}</td>}
                        {keys.size && <td className="whitespace-nowrap px-3 py-2">{String(r[keys.size] ?? "")}</td>}
                        {!keys.sku && !keys.name && !keys.ref && !keys.color && !keys.size && inv.columns.slice(0, 2).map((c) => (
                          <td key={c} className="px-3 py-2">{String(r[c] ?? "")}</td>
                        ))}
                        {keys.stock && (
                          <td className="whitespace-nowrap px-3 py-2 text-right text-muted-foreground">
                            {String(r[keys.stock] ?? "")}
                          </td>
                        )}
                        <td className="px-3 py-2 text-right">
                          <Input
                            type="number"
                            min={0}
                            value={sold || ""}
                            onChange={(e) => setSale(r.__id, Number(e.target.value) || 0)}
                            className={`ml-auto w-24 text-right ${over ? "border-destructive" : ""}`}
                            placeholder="0"
                          />
                          {over && (
                            <p className="mt-1 text-xs text-destructive">Excede stock</p>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {rows.length === 0 && (
              <p className="py-8 text-center text-sm text-muted-foreground">
                No se encontraron productos para "{filter}".
              </p>
            )}
          </CardContent>
        </Card>
      </main>
    </div>
  );
}

function Header() {
  return (
    <header className="border-b border-border bg-card">
      <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Ventas del mes</h1>
          <p className="text-sm text-muted-foreground">Vista del proveedor</p>
        </div>
        <Link
          to="/"
          className="rounded-md border border-border bg-background px-4 py-2 text-sm font-medium text-foreground hover:bg-accent"
        >
          ← Panel admin
        </Link>
      </div>
    </header>
  );
}

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <Card>
      <CardContent className="pt-6">
        <p className="text-sm text-muted-foreground">{label}</p>
        <p className="mt-1 text-2xl font-semibold text-foreground">{value.toLocaleString()}</p>
      </CardContent>
    </Card>
  );
}