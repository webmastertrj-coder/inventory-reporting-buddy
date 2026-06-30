import { createFileRoute, Link } from "@tanstack/react-router";
import { useRef, useState } from "react";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { setInventory, useInventory, clearAll, type InventoryRow } from "@/lib/inventory-store";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Inventario | Panel del Administrador" },
      {
        name: "description",
        content:
          "Sube tu inventario en Excel y compártelo con tu proveedor para registrar las ventas del mes.",
      },
    ],
  }),
  component: Index,
});

function Index() {
  const inv = useInventory();
  const inputRef = useRef<HTMLInputElement>(null);
  const [loading, setLoading] = useState(false);

  const handleFile = async (file: File) => {
    setLoading(true);
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
      setInventory(columns, rows);
      toast.success(`Inventario cargado: ${rows.length} productos`);
    } catch (e) {
      console.error(e);
      toast.error("No se pudo leer el archivo");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <div>
            <h1 className="text-xl font-semibold text-foreground">Gestión de Inventario</h1>
            <p className="text-sm text-muted-foreground">Panel del administrador</p>
          </div>
          <Link
            to="/proveedor"
            className="rounded-md border border-border bg-background px-4 py-2 text-sm font-medium text-foreground hover:bg-accent"
          >
            Vista del proveedor →
          </Link>
        </div>
      </header>

      <main className="mx-auto max-w-6xl space-y-6 px-6 py-8">
        <Card>
          <CardHeader>
            <CardTitle>1. Sube tu inventario (.xlsx)</CardTitle>
            <CardDescription>
              La primera fila debe contener los nombres de las columnas. Se aceptan las columnas que tu archivo tenga.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
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
              <Button onClick={() => inputRef.current?.click()} disabled={loading}>
                {loading ? "Procesando…" : inv.rows.length ? "Reemplazar archivo" : "Seleccionar archivo"}
              </Button>
              {inv.uploadedAt && (
                <span className="text-sm text-muted-foreground">
                  Última carga: {new Date(inv.uploadedAt).toLocaleString()}
                </span>
              )}
              {inv.rows.length > 0 && (
                <Button
                  variant="ghost"
                  onClick={() => {
                    clearAll();
                    toast.success("Inventario limpiado");
                  }}
                >
                  Limpiar
                </Button>
              )}
            </div>
          </CardContent>
        </Card>

        {inv.rows.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle>2. Productos cargados</CardTitle>
              <CardDescription>
                {inv.rows.length} filas · {inv.columns.length} columnas
              </CardDescription>
            </CardHeader>
            <CardContent>
              <div className="overflow-auto rounded-md border border-border">
                <table className="w-full text-sm">
                  <thead className="bg-muted text-muted-foreground">
                    <tr>
                      {inv.columns.map((c) => (
                        <th key={c} className="whitespace-nowrap px-3 py-2 text-left font-medium">
                          {c}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {inv.rows.slice(0, 50).map((r) => (
                      <tr key={r.__id} className="border-t border-border">
                        {inv.columns.map((c) => (
                          <td key={c} className="whitespace-nowrap px-3 py-2">
                            {String(r[c] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {inv.rows.length > 50 && (
                <p className="mt-2 text-xs text-muted-foreground">
                  Mostrando 50 de {inv.rows.length} filas.
                </p>
              )}
            </CardContent>
          </Card>
        )}
      </main>
    </div>
  );
}