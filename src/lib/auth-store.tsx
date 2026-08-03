import React, { createContext, useContext, useState, useEffect } from "react";
import { supabase, isCloudEnabled } from "./supabase-client";

export interface User {
  email: string;
  role: "admin" | "vendedor";
  name: string;
  commission?: number;
  warehouseId?: string;
}

interface UserDatabaseEntry extends User {
  password?: string;
}

interface AuthContextType {
  user: User | null;
  loading: boolean;
  login: (email: string, password: string) => boolean;
  logout: () => void;
  sellers: User[];
  addSeller: (name: string, email: string, commission: number, warehouseId: string) => boolean;
  updateSellerCommission: (email: string, commission: number) => boolean;
  updateSellerWarehouseId: (email: string, warehouseId: string) => boolean;
  updateSellerEmail: (oldEmail: string, newEmail: string) => boolean;
  deleteSeller: (email: string) => boolean;
}

const AuthContext = createContext<AuthContextType | null>(null);

const SESSION_KEY = "auth_session_v1";
const USERS_KEY = "auth_users_v1";

const SEED_USERS: Record<string, UserDatabaseEntry> = {
  "admin@comodatos.com": { email: "admin@comodatos.com", role: "admin", name: "Administrador", password: "admin123", commission: 0, warehouseId: "00" },
  "vendedor@comodatos.com": { email: "vendedor@comodatos.com", role: "vendedor", name: "Vendedor Comodatos", password: "vendedor123", commission: 10, warehouseId: "01" },
  "vendedor2@comodatos.com": { email: "vendedor2@comodatos.com", role: "vendedor", name: "Vendedor Alterno", password: "vendedor123", commission: 8, warehouseId: "02" },
};

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [users, setUsers] = useState<Record<string, UserDatabaseEntry>>({});

  useEffect(() => {
    try {
      // 1. Load users database
      let storedUsers = localStorage.getItem(USERS_KEY);
      let parsedUsers: Record<string, UserDatabaseEntry> = {};
      if (storedUsers) {
        parsedUsers = JSON.parse(storedUsers);
      } else {
        parsedUsers = SEED_USERS;
        localStorage.setItem(USERS_KEY, JSON.stringify(SEED_USERS));
      }
      setUsers(parsedUsers);

      // 2. Load active session
      const rawSession = localStorage.getItem(SESSION_KEY);
      if (rawSession) {
        setUser(JSON.parse(rawSession));
      }
    } catch (e) {
      console.error("Error reading auth / users state:", e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    const client = supabase;
    if (!isCloudEnabled || !client) return;

    const syncSellers = () => {
      client
        .from("sellers")
        .select("*")
        .then(({ data, error }) => {
          if (error) {
            console.error("Error loading sellers from cloud:", error);
            return;
          }
          if (data) {
            setUsers((prev) => {
              const mappedUsers: Record<string, UserDatabaseEntry> = {};
              data.forEach((row) => {
                mappedUsers[row.email.toLowerCase()] = {
                  email: row.email,
                  name: row.name,
                  role: row.role as any,
                  password: row.password,
                  commission: Number(row.commission),
                  warehouseId: row.warehouse_id,
                };
              });

              // Detect local-only sellers that are not in the cloud
              const localOnlySellers = Object.values(prev).filter(
                (u) => u.role === "vendedor" && !mappedUsers[u.email.toLowerCase()]
              );

              if (localOnlySellers.length > 0) {
                localOnlySellers.forEach((s) => {
                  client
                    .from("sellers")
                    .insert({
                      email: s.email.toLowerCase(),
                      name: s.name,
                      password: s.password || "vendedor123",
                      role: s.role,
                      commission: s.commission ?? 10,
                      warehouse_id: s.warehouseId ?? "01",
                    })
                    .then(({ error: insErr }) => {
                      if (insErr) console.error("Error backing up seller to cloud:", insErr);
                    });
                });
              }
              
              const merged = { ...prev, ...mappedUsers };
              localStorage.setItem(USERS_KEY, JSON.stringify(merged));
              return merged;
            });
          }
        });
    };

    syncSellers();
    
    // Poll sellers list from cloud every 12 seconds
    const interval = setInterval(syncSellers, 12000);
    
    // Sync on window focus
    window.addEventListener("focus", syncSellers);
    
    return () => {
      clearInterval(interval);
      window.removeEventListener("focus", syncSellers);
    };
  }, []);

  const login = (email: string, password: string): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    const foundUser = users[normalizedEmail];
    if (foundUser && foundUser.password === password) {
      const sessionUser: User = {
        email: foundUser.email,
        role: foundUser.role,
        name: foundUser.name,
        warehouseId: foundUser.warehouseId ?? "01",
      };
      localStorage.setItem(SESSION_KEY, JSON.stringify(sessionUser));
      setUser(sessionUser);
      return true;
    }
    return false;
  };

  const logout = () => {
    localStorage.removeItem(SESSION_KEY);
    setUser(null);
  };

  const addSeller = (name: string, email: string, commission: number, warehouseId: string): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    if (users[normalizedEmail]) {
      // User already exists
      return false;
    }
    const formattedWarehouseId = warehouseId.trim().padStart(2, "0");
    const newUser: UserDatabaseEntry = {
      name,
      email: normalizedEmail,
      role: "vendedor",
      password: "vendedor123", // Default password for new sellers
      commission,
      warehouseId: formattedWarehouseId,
    };
    const updatedUsers = { ...users, [normalizedEmail]: newUser };
    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);

    const client = supabase;
    if (isCloudEnabled && client) {
      client
        .from("sellers")
        .insert({
          email: normalizedEmail,
          name,
          password: "vendedor123",
          role: "vendedor",
          commission,
          warehouse_id: formattedWarehouseId,
        })
        .then(({ error }) => {
          if (error) console.error("Error adding seller to cloud:", error);
        });
    }
    return true;
  };

  const sellers = Object.values(users)
    .filter((u) => u.role === "vendedor")
    .map((u) => ({ 
      email: u.email, 
      role: u.role, 
      name: u.name, 
      commission: u.commission ?? 0, 
      warehouseId: u.warehouseId ?? "01" 
    }));

  const updateSellerCommission = (email: string, commission: number): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!users[normalizedEmail]) {
      return false;
    }
    const updatedUser = { ...users[normalizedEmail], commission };
    const updatedUsers = { ...users, [normalizedEmail]: updatedUser };
    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);

    const client = supabase;
    if (isCloudEnabled && client) {
      client
        .from("sellers")
        .update({ commission })
        .eq("email", normalizedEmail)
        .then(({ error }) => {
          if (error) console.error("Error updating seller commission in cloud:", error);
        });
    }
    return true;
  };

  const updateSellerWarehouseId = (email: string, warehouseId: string): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!users[normalizedEmail]) {
      return false;
    }
    const formattedWarehouseId = warehouseId.trim().padStart(2, "0");
    const updatedUser = { ...users[normalizedEmail], warehouseId: formattedWarehouseId };
    const updatedUsers = { ...users, [normalizedEmail]: updatedUser };
    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);

    const client = supabase;
    if (isCloudEnabled && client) {
      client
        .from("sellers")
        .update({ warehouse_id: formattedWarehouseId })
        .eq("email", normalizedEmail)
        .then(({ error }) => {
          if (error) console.error("Error updating seller warehouse in cloud:", error);
        });
    }
    return true;
  };

  const updateSellerEmail = (oldEmail: string, newEmail: string): boolean => {
    const oldNormalized = oldEmail.trim().toLowerCase();
    const newNormalized = newEmail.trim().toLowerCase();

    if (!newNormalized || !newNormalized.includes("@")) {
      return false;
    }

    if (!users[oldNormalized]) {
      return false;
    }

    if (oldNormalized !== newNormalized && users[newNormalized]) {
      return false; // Email already taken
    }

    const existingUser = users[oldNormalized];
    const updatedUser: UserDatabaseEntry = {
      ...existingUser,
      email: newNormalized,
    };

    const updatedUsers = { ...users };
    delete updatedUsers[oldNormalized];
    updatedUsers[newNormalized] = updatedUser;

    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);

    if (user && user.email.trim().toLowerCase() === oldNormalized) {
      const updatedSession: User = {
        ...user,
        email: newNormalized,
      };
      localStorage.setItem(SESSION_KEY, JSON.stringify(updatedSession));
      setUser(updatedSession);
    }

    const client = supabase;
    if (isCloudEnabled && client) {
      client
        .from("sellers")
        .update({ email: newNormalized })
        .eq("email", oldNormalized)
        .then(({ error }) => {
          if (error) console.error("Error updating seller email in cloud:", error);
        });
    }

    return true;
  };

  const deleteSeller = (email: string): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!users[normalizedEmail]) {
      return false;
    }

    const updatedUsers = { ...users };
    delete updatedUsers[normalizedEmail];

    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);

    if (user && user.email.trim().toLowerCase() === normalizedEmail) {
      logout();
    }

    const client = supabase;
    if (isCloudEnabled && client) {
      client
        .from("sellers")
        .delete()
        .eq("email", normalizedEmail)
        .then(({ error }) => {
          if (error) console.error("Error deleting seller from cloud:", error);
        });
    }

    return true;
  };

  return (
    <AuthContext.Provider value={{ user, loading, login, logout, sellers, addSeller, updateSellerCommission, updateSellerWarehouseId, updateSellerEmail, deleteSeller }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
}
