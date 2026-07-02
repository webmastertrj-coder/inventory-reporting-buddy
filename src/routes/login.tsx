import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState, useEffect } from "react";
import { toast } from "sonner";
import { Eye, EyeOff, Lock, Mail, Loader2, ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent } from "@/components/ui/card";
import { useAuth } from "@/lib/auth-store";

export const Route = createFileRoute("/login")({
  head: () => ({
    meta: [
      { title: "Iniciar Sesión | Inventory Reporting Buddy" },
      { name: "description", content: "Ingresa al sistema para gestionar tu inventario o reportar ventas del mes." },
    ],
  }),
  component: LoginPage,
});

function LoginPage() {
  const { user, login, loading: authLoading } = useAuth();
  const navigate = useNavigate();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);

  // If already logged in, redirect to respective panel
  useEffect(() => {
    if (!authLoading && user) {
      if (user.role === "admin") {
        navigate({ to: "/" });
      } else {
        navigate({ to: "/proveedor" });
      }
    }
  }, [user, authLoading, navigate]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !password) {
      toast.error("Por favor, completa todos los campos.");
      return;
    }

    setIsSubmitting(true);
    // Simulate a brief network latency for premium feel
    await new Promise((resolve) => setTimeout(resolve, 800));

    const success = login(email, password);

    if (success) {
      toast.success("Sesión iniciada con éxito");
    } else {
      toast.error("Correo o contraseña incorrectos.");
    }
    setIsSubmitting(false);
  };

  if (authLoading || (user && !isSubmitting)) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background">
        <div className="flex flex-col items-center gap-2">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <p className="text-sm text-muted-foreground">Verificando sesión...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex min-h-screen w-full flex-col md:flex-row bg-background animate-in fade-in duration-500">
      {/* Visual Section - Left Side */}
      <div className="relative hidden w-1/2 flex-col justify-between bg-zinc-950 p-12 text-white md:flex overflow-hidden">
        {/* Absolute Background Image */}
        <div className="absolute inset-0 z-0">
          <img
            src="/login_bg.png"
            alt="Decorative business network background"
            className="h-full w-full object-cover opacity-60 transition-transform duration-[10000ms] hover:scale-105"
          />
          <div className="absolute inset-0 bg-gradient-to-tr from-slate-950/90 via-slate-900/40 to-slate-950/30" />
        </div>

        {/* Brand header */}
        <div className="relative z-10 flex items-center gap-2">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-primary/20 backdrop-blur-md border border-primary/30">
            <span className="text-lg font-bold text-primary-foreground">IRB</span>
          </div>
          <span className="text-xl font-bold tracking-tight bg-gradient-to-r from-white to-slate-300 bg-clip-text text-transparent">
            Inventory Reporting Buddy
          </span>
        </div>

        {/* Big quote/text */}
        <div className="relative z-10 mt-auto max-w-lg space-y-4">
          <h2 className="text-4xl font-extrabold tracking-tight leading-tight">
            Optimiza tu stock. <br />
            Simplifica tus ventas.
          </h2>
          <p className="text-lg text-slate-300 font-light leading-relaxed">
            Nuestra plataforma unifica la gestión de inventario y el reporte de ventas mensuales entre administradores y proveedores autorizados.
          </p>
          <div className="h-[2px] w-24 bg-primary" />
        </div>
      </div>

      {/* Form Section - Right Side */}
      <div className="flex w-full items-center justify-center p-8 md:w-1/2 bg-background">
        <div className="mx-auto w-full max-w-md space-y-6">
          <div className="space-y-2 text-center md:text-left">
            {/* Mobile logo header */}
            <div className="flex items-center justify-center gap-2 md:hidden mb-6">
              <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-primary text-white">
                <span className="text-lg font-bold">IRB</span>
              </div>
              <span className="text-xl font-bold tracking-tight">Inventory Reporting Buddy</span>
            </div>

            <h1 className="text-3xl font-bold tracking-tight">Iniciar Sesión</h1>
            <p className="text-muted-foreground">
              Ingresa tus credenciales para acceder a tu panel de control.
            </p>
          </div>

          <Card className="border-border/60 shadow-lg bg-card/50 backdrop-blur-sm">
            <CardContent className="pt-6">
              <form onSubmit={handleSubmit} className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="email">Correo electrónico</Label>
                  <div className="relative">
                    <Mail className="absolute left-3 top-3 h-5 w-5 text-muted-foreground/75" />
                    <Input
                      id="email"
                      type="email"
                      placeholder="nombre@ejemplo.com"
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      className="pl-10"
                      required
                      disabled={isSubmitting}
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <Label htmlFor="password">Contraseña</Label>
                  </div>
                  <div className="relative">
                    <Lock className="absolute left-3 top-3 h-5 w-5 text-muted-foreground/75" />
                    <Input
                      id="password"
                      type={showPassword ? "text" : "password"}
                      placeholder="••••••••"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      className="pl-10 pr-10"
                      required
                      disabled={isSubmitting}
                    />
                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-3 top-3 text-muted-foreground hover:text-foreground focus:outline-none"
                      tabIndex={-1}
                      disabled={isSubmitting}
                    >
                      {showPassword ? <EyeOff className="h-5 w-5" /> : <Eye className="h-5 w-5" />}
                    </button>
                  </div>
                </div>

                <Button type="submit" className="w-full mt-2" disabled={isSubmitting}>
                  {isSubmitting ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      Iniciando sesión...
                    </>
                  ) : (
                    <>
                      Ingresar
                      <ArrowRight className="ml-2 h-4 w-4" />
                    </>
                  )}
                </Button>
              </form>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
