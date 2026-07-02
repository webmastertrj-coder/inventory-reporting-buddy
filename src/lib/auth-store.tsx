import React, { createContext, useContext, useState, useEffect } from "react";

export interface User {
  email: string;
  role: "admin" | "vendedor";
  name: string;
  commission?: number;
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
  addSeller: (name: string, email: string, commission: number) => boolean;
  updateSellerCommission: (email: string, commission: number) => boolean;
}

const AuthContext = createContext<AuthContextType | null>(null);

const SESSION_KEY = "auth_session_v1";
const USERS_KEY = "auth_users_v1";

const SEED_USERS: Record<string, UserDatabaseEntry> = {
  "admin@comodatos.com": { email: "admin@comodatos.com", role: "admin", name: "Administrador", password: "admin123", commission: 0 },
  "vendedor@comodatos.com": { email: "vendedor@comodatos.com", role: "vendedor", name: "Vendedor Comodatos", password: "vendedor123", commission: 10 },
  "vendedor2@comodatos.com": { email: "vendedor2@comodatos.com", role: "vendedor", name: "Vendedor Alterno", password: "vendedor123", commission: 8 },
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

  const login = (email: string, password: string): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    const foundUser = users[normalizedEmail];
    if (foundUser && foundUser.password === password) {
      const sessionUser: User = {
        email: foundUser.email,
        role: foundUser.role,
        name: foundUser.name,
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

  const addSeller = (name: string, email: string, commission: number): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    if (users[normalizedEmail]) {
      // User already exists
      return false;
    }
    const newUser: UserDatabaseEntry = {
      name,
      email: normalizedEmail,
      role: "vendedor",
      password: "vendedor123", // Default password for new sellers
      commission,
    };
    const updatedUsers = { ...users, [normalizedEmail]: newUser };
    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);
    return true;
  };

  const sellers = Object.values(users)
    .filter((u) => u.role === "vendedor")
    .map((u) => ({ email: u.email, role: u.role, name: u.name, commission: u.commission ?? 0 }));

  const updateSellerCommission = (email: string, commission: number): boolean => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!users[normalizedEmail]) {
      return false;
    }
    const updatedUser = { ...users[normalizedEmail], commission };
    const updatedUsers = { ...users, [normalizedEmail]: updatedUser };
    localStorage.setItem(USERS_KEY, JSON.stringify(updatedUsers));
    setUsers(updatedUsers);
    return true;
  };

  return (
    <AuthContext.Provider value={{ user, loading, login, logout, sellers, addSeller, updateSellerCommission }}>
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
