import React, { useEffect, useMemo, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
} from "recharts";
import {
  onAuthStateChanged,
  signInWithPopup,
  signOut,
  signInWithEmailAndPassword,
  sendPasswordResetEmail,
  type User,
} from "firebase/auth";
import {
  collection,
  addDoc,
  getDocs,
  query,
  where,
  Timestamp,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  serverTimestamp,
} from "firebase/firestore";
import { auth, googleProvider, db } from "./firebase/config";

type RawTradeRow = {
  name?: string;
  order_id?: string;
  symbol?: string;
  mov_time?: string;
  mov_type?: string | number;
  exec_qty?: string | number;
  price_done?: string | number;
  points?: string | number;
  profit?: string | number;
  created_on?: string;
  account?: string;
  orderId?: string;
  movTime?: string;
  movType?: string;
  qty?: string | number;
  price?: string | number;
  createdOn?: string;
};

type Trade = {
  account: string;
  orderId: string;
  symbol: string;
  movTime: string;
  movType: string;
  qty: number;
  price: number;
  points: number;
  profit: number;
  createdOn: string;
  day: string;
  importedFileName?: string;
  importedAt?: string;
};

type UserData = {
  name: string;
  email: string;
  photo: string;
};

type AccountConfig = {
  balanceStart: number;
  trailingDrawdown: number;
  profitTarget: number;
  typeLabel: string;
  fixedLiquidationThreshold?: number;
  payoutResetBalance?: number;
};

const COLORS = ["#22c55e", "#ef4444"];

const STORAGE_KEYS = {
  selectedAccount: "qboard_selected_account",
  search: "qboard_search",
  selectedFile: "qboard_selected_file",
  startDate: "qboard_start_date",
  endDate: "qboard_end_date",
  accountCategory: "qboard_account_category",
  accountStatus: "qboard_account_status",
};

const ACCOUNT_RULES: Record<string, AccountConfig> = {
  PA: {
    balanceStart: 0,
    trailingDrawdown: 0,
    profitTarget: 0,
    typeLabel: "PA",
  },
  APEX25: {
    balanceStart: 25000,
    trailingDrawdown: 1500,
    profitTarget: 1500,
    typeLabel: "25k",
  },
  APEX50: {
    balanceStart: 50000,
    trailingDrawdown: 2500,
    profitTarget: 3000,
    typeLabel: "50k",
  },
  APEX150: {
    balanceStart: 150000,
    trailingDrawdown: 5000,
    profitTarget: 9000,
    typeLabel: "150k",
  },
};

const PA_ACCOUNT_CONFIGS: Record<string, AccountConfig> = {
  "PA-APEX-244134-47": {
    balanceStart: 25000,
    trailingDrawdown: 1500,
    profitTarget: 0,
    typeLabel: "PA 25k",
    fixedLiquidationThreshold: 25100,
    payoutResetBalance: 26600,
  },
  // Adicione aqui outras PAs quando quiser que o resumo calcule
  // balance, liquidation threshold e drawdown com base no tamanho correto.
  // Exemplo:
  // "PA-APEX-XXXXXXXX-YY": {
  //   balanceStart: 50000,
  //   trailingDrawdown: 2500,
  //   profitTarget: 0,
  //   typeLabel: "PA 50k",
  // },
};



function getTradeDateOnly(trade: Trade): Date | null {
  if (trade.createdOn) {
    const d = new Date(trade.createdOn);
    if (!Number.isNaN(d.getTime())) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
  }

  if (trade.movTime) {
    const d = new Date(trade.movTime);
    if (!Number.isNaN(d.getTime())) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
  }

  if (trade.day) {
    const [dd, mm, yyyy] = trade.day.split("/");
    if (dd && mm && yyyy) {
      const d = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
      if (!Number.isNaN(d.getTime())) {
        return d;
      }
    }
  }

  return null;
}

function getDateKey(date: Date): string {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

const PA_DAILY_BALANCE_OVERRIDES: Record<string, Record<string, number>> = {
  "PA-APEX-244134-47": {
    "2026-03-17": 26900.62,
    "2026-03-18": 27127.20,
    "2026-03-19": 27241.60,
    "2026-03-20": 27558.90,
    "2026-03-23": 27771.86,
    "2026-03-24": 27947.56,
    "2026-03-25": 28187.74,
    "2026-03-26": 28262.92,
    "2026-03-27": 28452.52,
    "2026-03-30": 28707.58,
    "2026-04-01": 27207.58,
    "2026-04-02": 27446.38,
    "2026-04-03": 27617.40,
    "2026-04-06": 27782.08,
    "2026-04-07": 27951.32,
    "2026-04-08": 28156.44,
    "2026-04-09": 28186.90,
    "2026-04-10": 28131.86,
    "2026-04-13": 28110.82,
    "2026-04-15": 26610.82,
    "2026-04-16": 26782.92,
    "2026-04-17": 26938.42,
    "2026-04-20": 27022.18,
    "2026-04-21": 27216.18,
    "2026-04-22": 27407.48,
  },
};

function getPADailyBalanceOverride(
  account: string,
  referenceDate: Date | null
): number | null {
  const balances = PA_DAILY_BALANCE_OVERRIDES[account];
  if (!balances) return null;

  const referenceKey = referenceDate ? getDateKey(referenceDate) : null;
  const ordered = Object.entries(balances).sort((a, b) =>
    a[0].localeCompare(b[0])
  );

  let current: number | null = null;

  for (const [dateKey, balance] of ordered) {
    if (!referenceKey || dateKey <= referenceKey) {
      current = balance;
    } else {
      break;
    }
  }

  return current;
}

function detectAccountConfig(account: string): AccountConfig {
  if (account.startsWith("PA-")) {
    return PA_ACCOUNT_CONFIGS[account] || ACCOUNT_RULES.PA;
  }

  if (
    account.includes("-303") ||
    account.includes("-304") ||
    account.includes("-305")
  ) {
    return ACCOUNT_RULES.APEX150;
  }

  if (
    account.includes("-300") ||
    account.includes("-301") ||
    account.includes("-302")
  ) {
    return ACCOUNT_RULES.APEX50;
  }

  return ACCOUNT_RULES.APEX25;
}

function isPAAccount(account: string) {
  return account.trim().toUpperCase().startsWith("PA-");
}

function getAccountCategoryLabel(account: string) {
  return isPAAccount(account) ? "Aprovada / PA" : "Avaliação";
}

function formatCurrency(value: number) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function formatNumber(value: number) {
  return new Intl.NumberFormat("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function parseMoney(value: unknown): number {
  if (typeof value === "number") return value;
  if (typeof value !== "string") return NaN;

  const cleaned = value
    .trim()
    .replace(/\$/g, "")
    .replace(/\s/g, "")
    .replace(/,/g, "");

  if (!cleaned) return NaN;
  return Number(cleaned);
}

function parseValue(value: unknown): number {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;

  const text = String(value).trim();
  const commaDecimal = /^-?\d+,\d+$/.test(text);
  const normalized = commaDecimal
    ? text.replace(/\./g, "").replace(",", ".")
    : text.replace(/,/g, "");

  const parsed = Number(normalized);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function getDayLabel(row: RawTradeRow): string {
  const createdOn = row.created_on || row.createdOn;
  if (createdOn) {
    const d = new Date(createdOn);
    if (!Number.isNaN(d.getTime())) {
      return d.toLocaleDateString("pt-BR");
    }
  }

  const movTime = row.mov_time || row.movTime;
  if (movTime) {
    const d = new Date(movTime);
    if (!Number.isNaN(d.getTime())) {
      return d.toLocaleDateString("pt-BR");
    }
  }

  return "Sem data";
}

function buildTrades(rows: RawTradeRow[], importedFileName: string): Trade[] {
  const importedAt = new Date().toISOString();

  return rows
    .map((row) => {
      const profit = parseMoney(row.profit);

      return {
        account: row.name?.trim() || row.account?.trim() || "Sem conta",
        orderId: row.order_id?.trim() || row.orderId?.trim() || "",
        symbol: row.symbol?.trim() || "",
        movTime: row.mov_time?.trim() || row.movTime?.trim() || "",
        movType: String(row.mov_type ?? row.movType ?? ""),
        qty: Number(row.exec_qty ?? row.qty ?? 0),
        price: parseValue(row.price_done ?? row.price ?? 0),
        points: parseValue(row.points ?? 0),
        profit,
        createdOn: row.created_on?.trim() || row.createdOn?.trim() || "",
        day: getDayLabel(row),
        importedFileName,
        importedAt,
      };
    })
    .filter(
      (trade) =>
        !Number.isNaN(trade.profit) &&
        (trade.account || trade.symbol || trade.orderId)
    );
}

function mapFirebaseUser(user: User): UserData {
  return {
    name: user.displayName || "Usuário",
    email: user.email || "",
    photo: user.photoURL || "",
  };
}

async function ensureAppUser(firebaseUser: User) {
  const email = firebaseUser.email;

  if (!email) {
    throw new Error("Usuário sem email.");
  }

  const userRef = doc(db, "app_users", email);
  const userSnap = await getDoc(userRef);

  if (!userSnap.exists()) {
    await setDoc(userRef, {
      email,
      name: firebaseUser.displayName || "",
      photo: firebaseUser.photoURL || "",
      active: false,
      createdAt: serverTimestamp(),
      lastLoginAt: serverTimestamp(),
    });

    return {
      exists: false,
      active: false,
    };
  }

  const data = userSnap.data();

  await updateDoc(userRef, {
    name: firebaseUser.displayName || data.name || "",
    photo: firebaseUser.photoURL || data.photo || "",
    lastLoginAt: serverTimestamp(),
  });

  return {
    exists: true,
    active: Boolean(data.active),
  };
}

function MetricCard({
  title,
  value,
  subtitle,
  color = "#e2e8f0",
}: {
  title: string;
  value: string;
  subtitle?: string;
  color?: string;
}) {
  return (
    <div style={styles.card}>
      <div style={styles.cardTitle}>{title}</div>
      <div style={{ ...styles.metricValue, color }}>{value}</div>
      {subtitle ? <div style={styles.metricSub}>{subtitle}</div> : null}
    </div>
  );
}

function LoginScreen({
  onGoogleLogin,
  onEmailLogin,
  onForgotPassword,
  loading,
  emailLogin,
  setEmailLogin,
  passwordLogin,
  setPasswordLogin,
}: {
  onGoogleLogin: () => void;
  onEmailLogin: () => void;
  onForgotPassword: () => void;
  loading: boolean;
  emailLogin: string;
  setEmailLogin: React.Dispatch<React.SetStateAction<string>>;
  passwordLogin: string;
  setPasswordLogin: React.Dispatch<React.SetStateAction<string>>;
}) {
  return (
    <div style={styles.loginWrap}>
      <div style={styles.loginPanel}>
        <div style={styles.brandBadge}>QiBoard</div>
        <h1 style={styles.loginTitle}>
          Painel profissional para contas mesas proprietárias
        </h1>
        <p style={styles.loginText}>
          Faça login com sua conta Google ou com email e senha para acessar o
          dashboard.
        </p>

        <input
          type="email"
          placeholder="Digite seu email"
          value={emailLogin}
          onChange={(e) => setEmailLogin(e.target.value)}
          style={styles.loginInput}
        />

        <input
          type="password"
          placeholder="Digite sua senha"
          value={passwordLogin}
          onChange={(e) => setPasswordLogin(e.target.value)}
          style={styles.loginInput}
        />

        <div style={{ textAlign: "right", marginTop: -2, marginBottom: 12 }}>
          <span
            style={styles.loginLink}
            onClick={onForgotPassword}
          >
            Esqueci a senha
          </span>
        </div>

        <button
          style={styles.googleButton}
          onClick={onEmailLogin}
          disabled={loading}
        >
          {loading ? "Entrando..." : "Entrar com email"}
        </button>

        <button
  style={styles.googleLoginButton}
  onClick={onGoogleLogin}
  disabled={loading}
>
  <span style={styles.googleIconWrap}>
    <svg width="18" height="18" viewBox="0 0 48 48">
      <path
        fill="#FFC107"
        d="M43.611 20.083H42V20H24v8h11.303C33.655 32.657 29.195 36 24 36c-6.627 0-12-5.373-12-12s5.373-12 12-12c3.059 0 5.842 1.154 7.96 3.04l5.657-5.657C34.046 6.053 29.277 4 24 4 12.955 4 4 12.955 4 24s8.955 20 20 20 20-8.955 20-20c0-1.341-.138-2.65-.389-3.917z"
      />
      <path
        fill="#FF3D00"
        d="M6.306 14.691l6.571 4.819C14.655 15.108 18.961 12 24 12c3.059 0 5.842 1.154 7.96 3.04l5.657-5.657C34.046 6.053 29.277 4 24 4 16.318 4 9.656 8.337 6.306 14.691z"
      />
      <path
        fill="#4CAF50"
        d="M24 44c5.176 0 9.86-1.977 13.409-5.192l-6.19-5.238C29.143 35.091 26.715 36 24 36c-5.174 0-9.623-3.326-11.283-7.946l-6.522 5.025C9.505 39.556 16.227 44 24 44z"
      />
      <path
        fill="#1976D2"
        d="M43.611 20.083H42V20H24v8h11.303c-.792 2.237-2.231 4.166-4.084 5.571l.003-.002 6.19 5.238C36.971 39.205 44 34 44 24c0-1.341-.138-2.65-.389-3.917z"
      />
    </svg>
  </span>

  <span>{loading ? "Entrando..." : "Entrar com Google"}</span>
</button>
      </div>
    </div>
  );
}

function DashboardScreen({
  onLogout,
  user,
}: {
  onLogout: () => void;
  user: UserData;
}) {
  const [trades, setTrades] = useState<Trade[]>([]);
  const [selectedAccount, setSelectedAccount] = useState("Todas");
  const [search, setSearch] = useState("");
  const [selectedFile, setSelectedFile] = useState("Todos");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [accountCategory, setAccountCategory] = useState("Todas");
  const [accountStatus, setAccountStatus] = useState("Todas");
  const [lastImportedFileName, setLastImportedFileName] = useState("");
  const [screenWidth, setScreenWidth] = useState(window.innerWidth);
  const [loadingData, setLoadingData] = useState(false);

  useEffect(() => {
    function handleResize() {
      setScreenWidth(window.innerWidth);
    }

    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    const savedSelectedAccount = localStorage.getItem(
      STORAGE_KEYS.selectedAccount
    );
    const savedSearch = localStorage.getItem(STORAGE_KEYS.search);
    const savedSelectedFile = localStorage.getItem(STORAGE_KEYS.selectedFile);
    const savedStartDate = localStorage.getItem(STORAGE_KEYS.startDate);
    const savedEndDate = localStorage.getItem(STORAGE_KEYS.endDate);
    const savedAccountCategory = localStorage.getItem(
      STORAGE_KEYS.accountCategory
    );
    const savedAccountStatus = localStorage.getItem(
      STORAGE_KEYS.accountStatus
    );

    if (savedSelectedAccount) setSelectedAccount(savedSelectedAccount);
    if (savedSearch) setSearch(savedSearch);
    if (savedSelectedFile) setSelectedFile(savedSelectedFile);
    if (savedStartDate) setStartDate(savedStartDate);
    if (savedEndDate) setEndDate(savedEndDate);
    if (savedAccountCategory) setAccountCategory(savedAccountCategory);
    if (savedAccountStatus) setAccountStatus(savedAccountStatus);
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.selectedAccount, selectedAccount);
  }, [selectedAccount]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.search, search);
  }, [search]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.selectedFile, selectedFile);
  }, [selectedFile]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.startDate, startDate);
  }, [startDate]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.endDate, endDate);
  }, [endDate]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.accountCategory, accountCategory);
  }, [accountCategory]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.accountStatus, accountStatus);
  }, [accountStatus]);

  async function saveTradesToFirestore(importedTrades: Trade[]) {
    if (!user?.email) return;

    try {
      const existingQuery = query(
        collection(db, "trades"),
        where("userEmail", "==", user.email)
      );

      const existingSnapshot = await getDocs(existingQuery);
      const existingKeys = new Set(
        existingSnapshot.docs.map((docItem) => {
          const data = docItem.data();
          return `${data.userEmail || ""}__${data.orderId || ""}`;
        })
      );

      let insertedCount = 0;

      for (const trade of importedTrades) {
        const dedupeKey = `${user.email}__${trade.orderId || ""}`;

        if (trade.orderId && existingKeys.has(dedupeKey)) {
          continue;
        }

        await addDoc(collection(db, "trades"), {
          userEmail: user.email,
          account: trade.account,
          orderId: trade.orderId,
          symbol: trade.symbol,
          movTime: trade.movTime,
          movType: trade.movType,
          qty: trade.qty,
          price: trade.price,
          points: trade.points,
          profit: trade.profit,
          createdOn: trade.createdOn,
          day: trade.day,
          importedFileName: trade.importedFileName || "",
          importedAt: trade.importedAt || new Date().toISOString(),
        });

        insertedCount += 1;
      }

      alert(
        insertedCount > 0
          ? `${insertedCount} trade(s) novo(s) salvo(s) no banco.`
          : "Nenhum trade novo foi salvo. Os orderId já existem no banco."
      );
    } catch (error) {
      console.error("Erro ao salvar trades no Firestore:", error);
      alert("Erro ao salvar trades no banco.");
    }
  }

  async function loadTradesFromFirestore() {
    if (!user?.email) return;

    try {
      setLoadingData(true);

      const q = query(
        collection(db, "trades"),
        where("userEmail", "==", user.email)
      );

      const snapshot = await getDocs(q);

      const loadedTrades: Trade[] = snapshot.docs.map((docItem) => {
        const data = docItem.data();

        return {
          account: data.account || "",
          orderId: data.orderId || "",
          symbol: data.symbol || "",
          movTime: data.movTime || "",
          movType: data.movType || "",
          qty: Number(data.qty || 0),
          price: Number(data.price || 0),
          points: Number(data.points || 0),
          profit: Number(data.profit || 0),
          createdOn: data.createdOn || "",
          day: data.day || "",
          importedFileName: data.importedFileName || "",
          importedAt:
            data.importedAt instanceof Timestamp
              ? data.importedAt.toDate().toISOString()
              : data.importedAt || "",
        };
      });

      loadedTrades.sort((a, b) =>
        String(b.importedAt || "").localeCompare(String(a.importedAt || ""))
      );

      setTrades(loadedTrades);

      if (loadedTrades.length > 0) {
        setLastImportedFileName(loadedTrades[0].importedFileName || "");
      }
    } catch (error) {
      console.error("Erro ao carregar trades do Firestore:", error);
      alert("Erro ao recarregar dados do banco.");
    } finally {
      setLoadingData(false);
    }
  }

  useEffect(() => {
    loadTradesFromFirestore();
  }, [user?.email]);

  const importedFiles = useMemo(() => {
    return Array.from(
      new Set(
        trades
          .map((trade) => trade.importedFileName)
          .filter((value): value is string => Boolean(value))
      )
    ).sort((a, b) => a.localeCompare(b));
  }, [trades]);

  const accounts = useMemo(() => {
    return Array.from(new Set(trades.map((trade) => trade.account))).sort();
  }, [trades]);

  function tradeMatchesDateRange(trade: Trade) {
    if (!startDate && !endDate) return true;

    let tradeDate: Date | null = null;

    if (trade.createdOn) {
      const d = new Date(trade.createdOn);
      if (!Number.isNaN(d.getTime())) {
        tradeDate = d;
      }
    }

    if (!tradeDate && trade.movTime) {
      const d = new Date(trade.movTime);
      if (!Number.isNaN(d.getTime())) {
        tradeDate = d;
      }
    }

    if (!tradeDate) return false;

    const tradeOnlyDate = new Date(
      tradeDate.getFullYear(),
      tradeDate.getMonth(),
      tradeDate.getDate()
    );

    if (startDate) {
      const start = new Date(`${startDate}T00:00:00`);
      if (tradeOnlyDate < start) return false;
    }

    if (endDate) {
      const end = new Date(`${endDate}T23:59:59`);
      if (tradeOnlyDate > end) return false;
    }

    return true;
  }

  const filteredTrades = useMemo(() => {
    return trades.filter((trade) => {
      const accountOk =
        selectedAccount === "Todas" || trade.account === selectedAccount;

      const fileOk =
        selectedFile === "Todos" || trade.importedFileName === selectedFile;

      const categoryOk =
        accountCategory === "Todas" ||
        (accountCategory === "Aprovadas" && isPAAccount(trade.account)) ||
        (accountCategory === "Avaliação" && !isPAAccount(trade.account));

      const text = [
        trade.account,
        trade.symbol,
        trade.day,
        trade.orderId,
        trade.importedFileName || "",
      ]
        .join(" ")
        .toLowerCase();

      const searchOk = text.includes(search.toLowerCase());
      const dateOk = tradeMatchesDateRange(trade);

      return accountOk && fileOk && categoryOk && searchOk && dateOk;
    });
  }, [
    trades,
    selectedAccount,
    selectedFile,
    accountCategory,
    search,
    startDate,
    endDate,
  ]);

  const riskReferenceDate = useMemo(() => {
    const value = endDate || startDate;
    return value ? new Date(`${value}T00:00:00`) : null;
  }, [startDate, endDate]);

  const cumulativeNetByAccount = useMemo(() => {
    const map = new Map<string, number>();

    trades.forEach((trade) => {
      const accountOk =
        selectedAccount === "Todas" || trade.account === selectedAccount;

      const categoryOk =
        accountCategory === "Todas" ||
        (accountCategory === "Aprovadas" && isPAAccount(trade.account)) ||
        (accountCategory === "Avaliação" && !isPAAccount(trade.account));

      if (!accountOk || !categoryOk) return;

      if (riskReferenceDate) {
        const tradeDate = getTradeDateOnly(trade);
        if (!tradeDate || tradeDate.getTime() > riskReferenceDate.getTime()) {
          return;
        }
      }

      map.set(trade.account, (map.get(trade.account) || 0) + trade.profit);
    });

    return map;
  }, [trades, selectedAccount, accountCategory, riskReferenceDate]);

  const metrics = useMemo(() => {
    const positive = filteredTrades.filter((t) => t.profit > 0);
    const negative = filteredTrades.filter((t) => t.profit < 0);

    const grossProfit = positive.reduce((sum, t) => sum + t.profit, 0);
    const grossLoss = negative.reduce((sum, t) => sum + t.profit, 0);
    const net = filteredTrades.reduce((sum, t) => sum + t.profit, 0);
    const wins = positive.length;
    const losses = negative.length;
    const total = filteredTrades.length;
    const winRate = total ? (wins / total) * 100 : 0;
    const avgWin = wins ? grossProfit / wins : 0;
    const avgLoss = losses ? Math.abs(grossLoss / losses) : 0;
    const payoff = avgLoss !== 0 ? avgWin / avgLoss : 0;

    return {
      grossProfit,
      grossLoss,
      net,
      wins,
      losses,
      total,
      winRate,
      avgWin,
      avgLoss,
      payoff,
    };
  }, [filteredTrades]);

  const byAccount = useMemo(() => {
    const map = new Map<
      string,
      {
        account: string;
        positive: number;
        negative: number;
        net: number;
        trades: number;
      }
    >();

    filteredTrades.forEach((trade) => {
      const current = map.get(trade.account) || {
        account: trade.account,
        positive: 0,
        negative: 0,
        net: 0,
        trades: 0,
      };

      current.net += trade.profit;
      current.trades += 1;
      if (trade.profit > 0) current.positive += trade.profit;
      if (trade.profit < 0) current.negative += trade.profit;

      map.set(trade.account, current);
    });

    return Array.from(map.values()).sort((a, b) => b.net - a.net);
  }, [filteredTrades]);

  const riskByAccount = useMemo(() => {
    return byAccount.map((row) => {
      const cfg = detectAccountConfig(row.account);
      const isPA = isPAAccount(row.account);
      const cumulativeNet = cumulativeNetByAccount.get(row.account) || 0;

      let liquidationThresholdBase =
        cfg.fixedLiquidationThreshold ??
        (cfg.balanceStart > 0 ? cfg.balanceStart - cfg.trailingDrawdown : 0);

      let balanceCurrent =
        cfg.balanceStart > 0 ? cfg.balanceStart + cumulativeNet : cumulativeNet;

      let liquidationThresholdCurrent =
        cfg.fixedLiquidationThreshold ??
        (cfg.balanceStart > 0
          ? liquidationThresholdBase + cumulativeNet
          : liquidationThresholdBase);

      if (isPA) {
        const overrideBalance = getPADailyBalanceOverride(
          row.account,
          riskReferenceDate
        );

        if (overrideBalance !== null) {
          balanceCurrent = overrideBalance;
        }

        liquidationThresholdCurrent =
          cfg.fixedLiquidationThreshold ?? liquidationThresholdBase;
      }

      const dailyDrawdown = balanceCurrent - liquidationThresholdCurrent;
      const drawdownAvailable = dailyDrawdown;

      const distanceToTarget =
        cfg.profitTarget > 0 ? cfg.profitTarget - cumulativeNet : 0;

      const progressToTarget =
        cfg.profitTarget > 0
          ? Math.max(0, Math.min(100, (cumulativeNet / cfg.profitTarget) * 100))
          : 0;

      const riskPerDay =
        dailyDrawdown > 0 ? Math.max(0, dailyDrawdown * 0.05) : 0;

      let status = "Ativa";

      if (!isPA && cfg.balanceStart > 0) {
        if (dailyDrawdown <= 0) {
          status = "Reprovada";
        } else if (dailyDrawdown < cfg.trailingDrawdown * 0.2) {
          status = "Crítica";
        } else if (dailyDrawdown < cfg.trailingDrawdown * 0.4) {
          status = "Atenção";
        } else {
          status = "Ativa";
        }
      }

      return {
        account: row.account,
        category: getAccountCategoryLabel(row.account),
        isPA,
        typeLabel: cfg.typeLabel,
        balanceStart: cfg.balanceStart,
        liquidationThresholdBase,
        balanceCurrent,
        liquidationThresholdCurrent,
        dailyDrawdown,
        drawdownAvailable,
        trailingDrawdown: cfg.trailingDrawdown,
        target: cfg.profitTarget,
        distanceToTarget,
        progressToTarget,
        riskPerDay,
        status,
        net: row.net,
        cumulativeNet,
        positive: row.positive,
        negative: row.negative,
        trades: row.trades,
      };
    });
  }, [byAccount, cumulativeNetByAccount, riskReferenceDate]);

  const filteredRiskAccounts = useMemo(() => {
    return riskByAccount.filter((row) => {
      const categoryOk =
        accountCategory === "Todas" ||
        (accountCategory === "Aprovadas" && row.isPA) ||
        (accountCategory === "Avaliação" && !row.isPA);

      const statusOk =
        accountStatus === "Todas" ||
        (accountStatus === "Ativas" && row.status !== "Reprovada") ||
        (accountStatus === "Reprovadas" && row.status === "Reprovada");

      const selectedAccountOk =
        selectedAccount === "Todas" || row.account === selectedAccount;

      return categoryOk && statusOk && selectedAccountOk;
    });
  }, [riskByAccount, accountCategory, accountStatus, selectedAccount]);

  const selectedRiskAccount = useMemo(() => {
    if (selectedAccount !== "Todas") {
      return (
        filteredRiskAccounts.find((r) => r.account === selectedAccount) || null
      );
    }
    return filteredRiskAccounts[0] || null;
  }, [filteredRiskAccounts, selectedAccount]);

  const byDay = useMemo(() => {
    const map = new Map<string, number>();

    filteredTrades.forEach((trade) => {
      map.set(trade.day, (map.get(trade.day) || 0) + trade.profit);
    });

    const result = Array.from(map.entries()).map(([day, total]) => {
      const [d, m, y] = day.split("/");
      const dateObj = new Date(`${y}-${m}-${d}`);

      return {
        day,
        total,
        dateObj,
      };
    });

    return result
      .sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime())
      .map(({ day, total }) => ({ day, total }));
  }, [filteredTrades]);

  const payoffByDay = useMemo(() => {
    const map = new Map<
      string,
      {
        day: string;
        gains: number;
        losses: number;
        grossProfit: number;
        grossLossAbs: number;
        net: number;
        avgGain: number;
        avgLoss: number;
        payoff: number | null;
        payoffLabel: string;
      }
    >();

    filteredTrades.forEach((trade) => {
      const current = map.get(trade.day) || {
        day: trade.day,
        gains: 0,
        losses: 0,
        grossProfit: 0,
        grossLossAbs: 0,
        net: 0,
        avgGain: 0,
        avgLoss: 0,
        payoff: null,
        payoffLabel: "-",
      };

      current.net += trade.profit;

      if (trade.profit > 0) {
        current.gains += 1;
        current.grossProfit += trade.profit;
      }

      if (trade.profit < 0) {
        current.losses += 1;
        current.grossLossAbs += Math.abs(trade.profit);
      }

      map.set(trade.day, current);
    });

    const result = Array.from(map.values()).map((item) => {
      const avgGain = item.gains > 0 ? item.grossProfit / item.gains : 0;
      const avgLoss = item.losses > 0 ? item.grossLossAbs / item.losses : 0;

      let payoff: number | null = null;
      let payoffLabel = "-";

      if (item.gains > 0 && item.losses === 0) {
        payoff = null;
        payoffLabel = "∞";
      } else if (item.gains === 0 && item.losses > 0) {
        payoff = 0;
        payoffLabel = "0,00";
      } else if (avgLoss > 0) {
        payoff = avgGain / avgLoss;
        payoffLabel = formatNumber(payoff);
      }

      return {
        ...item,
        avgGain,
        avgLoss,
        payoff,
        payoffLabel,
      };
    });

    return result.sort((a, b) => {
      const da = a.day.split("/").reverse().join("-");
      const db = b.day.split("/").reverse().join("-");
      return db.localeCompare(da);
    });
  }, [filteredTrades]);

  const bestDay = useMemo(() => {
    if (!payoffByDay.length) return null;
    return [...payoffByDay].sort((a, b) => b.net - a.net)[0];
  }, [payoffByDay]);

  const worstDay = useMemo(() => {
    if (!payoffByDay.length) return null;
    return [...payoffByDay].sort((a, b) => a.net - b.net)[0];
  }, [payoffByDay]);

  function getPayoffQuality(
    payoff: number | null,
    gains: number,
    losses: number
  ) {
    if (gains > 0 && losses === 0) {
      return { label: "Perfeito", color: "#22c55e" };
    }

    if (gains === 0 && losses > 0) {
      return { label: "Só loss", color: "#ef4444" };
    }

    if (payoff === null) {
      return { label: "-", color: "#94a3b8" };
    }

    if (payoff >= 2) {
      return { label: "Excelente", color: "#22c55e" };
    }

    if (payoff >= 1.5) {
      return { label: "Bom", color: "#84cc16" };
    }

    if (payoff >= 1) {
      return { label: "Ok", color: "#f59e0b" };
    }

    return { label: "Ruim", color: "#ef4444" };
  }

  function getDayScore(
    payoff: number | null,
    net: number,
    gains: number,
    losses: number
  ) {
    if (gains > 0 && losses === 0 && net > 0) return "A+";
    if (payoff !== null && payoff >= 2 && net > 0) return "A";
    if (payoff !== null && payoff >= 1.5 && net > 0) return "B";
    if (payoff !== null && payoff >= 1 && net >= 0) return "C";
    return "D";
  }

  const pieData = useMemo(
    () => [
      { name: "Gains", value: metrics.wins },
      { name: "Losses", value: metrics.losses },
    ],
    [metrics.wins, metrics.losses]
  );

  const layout = useMemo(() => {
    const isMobile = screenWidth < 900;
    const isNotebook = screenWidth >= 900 && screenWidth < 1440;

    return {
      appShell: {
        ...styles.appShell,
        gridTemplateColumns: isMobile ? "1fr" : "480px 1fr",
      } as React.CSSProperties,
      gridCards: {
        ...styles.gridCards,
        gridTemplateColumns: isMobile
          ? "1fr"
          : isNotebook
          ? "repeat(4, minmax(0, 1fr))"
          : "repeat(8, minmax(0, 1fr))",
      } as React.CSSProperties,
      chartGrid: {
        ...styles.chartGrid,
        gridTemplateColumns: isMobile ? "1fr" : isNotebook ? "1fr" : "2fr 1fr",
      } as React.CSSProperties,
      tableGrid: {
        ...styles.tableGrid,
        gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr",
      } as React.CSSProperties,
      riskGrid: {
        ...styles.riskGrid,
        gridTemplateColumns: isMobile ? "1fr" : "repeat(5, minmax(0, 1fr))",
      } as React.CSSProperties,
      main: {
        ...styles.main,
        padding: isMobile ? 16 : 28,
      } as React.CSSProperties,
      headerRow: {
        ...styles.headerRow,
        flexDirection: isMobile ? "column" : "row",
        alignItems: isMobile ? "flex-start" : "center",
        gap: isMobile ? 12 : 0,
      } as React.CSSProperties,
      sidebar: {
        ...styles.sidebar,
        minHeight: isMobile ? "auto" : "100vh",
      } as React.CSSProperties,
      dateFilterRow: {
        ...styles.dateFilterRow,
        gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr",
      } as React.CSSProperties,
    };
  }, [screenWidth]);

  async function handleParsedRows(
    rows: RawTradeRow[],
    importedFileName: string
  ) {
    const parsedTrades = buildTrades(rows, importedFileName);

    await saveTradesToFirestore(parsedTrades);
    await loadTradesFromFirestore();

    setLastImportedFileName(importedFileName);
    setSelectedFile(importedFileName);
  }

  function handleCsvFile(file: File) {
    Papa.parse<RawTradeRow>(file, {
      header: true,
      skipEmptyLines: true,
      complete: async (results) => {
        await handleParsedRows(results.data || [], file.name);
      },
    });
  }

  function handleXlsxFile(file: File) {
    const reader = new FileReader();

    reader.onload = async (event) => {
      const data = event.target?.result;
      if (!data) return;

      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json<RawTradeRow>(sheet, { defval: "" });

      await handleParsedRows(rows, file.name);
    };

    reader.readAsArrayBuffer(file);
  }

  function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    const extension = file.name.toLowerCase();

    if (extension.endsWith(".csv")) {
      handleCsvFile(file);
      return;
    }

    if (extension.endsWith(".xlsx") || extension.endsWith(".xls")) {
      handleXlsxFile(file);
      return;
    }

    alert("Formato não suportado. Use .csv, .xlsx ou .xls");
  }

  function exportSummaryExcel() {
    const summaryRows = filteredRiskAccounts.map((row) => ({
      Conta: row.account,
      Categoria: row.category,
      Tipo: row.typeLabel,
      SaldoInicial: row.balanceStart,
      LiquidationThresholdBase: row.liquidationThresholdBase,
      SaldoAtual: row.balanceCurrent,
      LiquidationThresholdAtual: row.liquidationThresholdCurrent,
      DrawdownDiario: row.dailyDrawdown,
      Positivos: row.positive,
      Negativos: row.negative,
      Liquido: row.net,
      Trades: row.trades,
      DrawdownDisponivel: row.drawdownAvailable,
      Meta: row.target,
      DistanciaMeta: row.distanceToTarget,
      RiscoDia: row.riskPerDay,
      Status: row.status,
    }));

    const payoffRows = payoffByDay.map((row) => ({
      Data: row.day,
      Gains: row.gains,
      Losses: row.losses,
      MediaGain: row.avgGain,
      MediaLoss: row.avgLoss,
      Payoff: row.payoffLabel,
      Liquido: row.net,
    }));

    const wb = XLSX.utils.book_new();
    const wsResumo = XLSX.utils.json_to_sheet(summaryRows);
    const wsPayoff = XLSX.utils.json_to_sheet(payoffRows);

    XLSX.utils.book_append_sheet(wb, wsResumo, "Risco e Metas");
    XLSX.utils.book_append_sheet(wb, wsPayoff, "Payoff por Dia");

    const excelBuffer = XLSX.write(wb, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(blob, "qiboard-risco-metas.xlsx");
  }

  function clearLocalFilters() {
    localStorage.removeItem(STORAGE_KEYS.selectedAccount);
    localStorage.removeItem(STORAGE_KEYS.search);
    localStorage.removeItem(STORAGE_KEYS.selectedFile);
    localStorage.removeItem(STORAGE_KEYS.startDate);
    localStorage.removeItem(STORAGE_KEYS.endDate);
    localStorage.removeItem(STORAGE_KEYS.accountCategory);
    localStorage.removeItem(STORAGE_KEYS.accountStatus);

    setSelectedAccount("Todas");
    setSearch("");
    setSelectedFile("Todos");
    setStartDate("");
    setEndDate("");
    setAccountCategory("Todas");
    setAccountStatus("Todas");
  }

  return (
    <div style={layout.appShell}>
      <aside style={layout.sidebar}>
        <div>
          <div style={styles.logo}>QiBoard</div>
          <div style={styles.sidebarSub}>Dashboard para qualquer prop firm</div>
        </div>

        <div style={styles.userBox}>
          {user.photo ? (
            <img src={user.photo} alt={user.name} style={styles.avatar} />
          ) : null}
          <div>
            <div style={styles.userName}>{user.name}</div>
            <div style={styles.userEmail}>{user.email}</div>
          </div>
        </div>

        <div style={styles.sideBox}>
          <div style={styles.sideBoxTitle}>Importação</div>
          <label style={styles.uploadLabel}>
            Importar arquivo
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              onChange={handleFileChange}
              style={{ display: "none" }}
            />
          </label>
          <div style={styles.fileName}>
            Último arquivo: {lastImportedFileName || "Nenhum"}
          </div>
        </div>

        <div style={styles.sideBox}>
          <div style={styles.sideBoxTitle}>Consulta do histórico</div>
          <select
            value={selectedFile}
            onChange={(e) => setSelectedFile(e.target.value)}
            style={styles.select}
          >
            <option value="Todos">Todos os arquivos</option>
            {importedFiles.map((file) => (
              <option key={file} value={file}>
                {file}
              </option>
            ))}
          </select>

          <select
            value={selectedAccount}
            onChange={(e) => setSelectedAccount(e.target.value)}
            style={styles.select}
          >
            <option value="Todas">Todas as contas</option>
            {accounts.map((account) => (
              <option key={account} value={account}>
                {account}
              </option>
            ))}
          </select>

          <select
            value={accountCategory}
            onChange={(e) => setAccountCategory(e.target.value)}
            style={styles.select}
          >
            <option value="Todas">Todas as categorias</option>
            <option value="Aprovadas">Contas aprovadas / PA</option>
            <option value="Avaliação">Contas de avaliação</option>
          </select>

          <select
            value={accountStatus}
            onChange={(e) => setAccountStatus(e.target.value)}
            style={styles.select}
          >
            <option value="Todas">Todos os status</option>
            <option value="Ativas">Ativas</option>
            <option value="Reprovadas">Reprovadas</option>
          </select>

          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Buscar conta, símbolo, data, arquivo..."
            style={styles.inputDark}
          />

          <div style={layout.dateFilterRow}>
            <input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              style={styles.inputDark}
            />

            <input
              type="date"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              style={styles.inputDark}
            />
          </div>
        </div>

        <div style={styles.sideBox}>
          <div style={styles.sideBoxTitle}>Ações</div>
          <button style={styles.actionButton} onClick={loadTradesFromFirestore}>
            Pesquisar
          </button>
          <button style={styles.actionButton} onClick={exportSummaryExcel}>
            Exportar risco e metas
          </button>
          <button style={styles.clearButton} onClick={clearLocalFilters}>
            Limpar
          </button>
          <button style={styles.secondaryButton} onClick={onLogout}>
            Sair
          </button>
        </div>
      </aside>

      <main style={layout.main}>
        <div style={layout.headerRow}>
          <div>
            <h1 style={styles.title}>QiBoard - Painel Principal</h1>
            <p style={styles.subtitle}>
              Relatório profissional de performance, risco, metas e execução.
            </p>
            <p style={styles.subtitle}>
              {loadingData
                ? "Carregando dados do banco..."
                : `Histórico carregado: ${filteredTrades.length} trade(s)`}
            </p>
          </div>
        </div>

        <div style={layout.gridCards}>
          <MetricCard
            title="Lucro bruto"
            value={formatCurrency(metrics.grossProfit)}
            color="#22c55e"
          />
          <MetricCard
            title="Prejuízo bruto"
            value={formatCurrency(metrics.grossLoss)}
            color="#ef4444"
          />
          <MetricCard
            title="Resultado líquido"
            value={formatCurrency(metrics.net)}
            color={metrics.net >= 0 ? "#22c55e" : "#ef4444"}
          />
          <MetricCard
            title="Trades"
            value={String(metrics.total)}
            subtitle={`${metrics.wins} gains / ${metrics.losses} losses`}
          />
          <MetricCard
            title="Win rate"
            value={`${formatNumber(metrics.winRate)}%`}
          />
          <MetricCard
            title="Payoff geral"
            value={formatNumber(metrics.payoff)}
            subtitle="ganho médio ÷ perda média"
          />
          <MetricCard
            title="Melhor dia"
            value={bestDay ? bestDay.day : "-"}
            subtitle={bestDay ? formatCurrency(bestDay.net) : "Sem dados"}
            color="#22c55e"
          />
          <MetricCard
            title="Pior dia"
            value={worstDay ? worstDay.day : "-"}
            subtitle={worstDay ? formatCurrency(worstDay.net) : "Sem dados"}
            color="#ef4444"
          />
        </div>

        {selectedRiskAccount ? (
          <div style={styles.panel}>
            <div style={styles.panelTitle}>Metas e risco da conta</div>
            <div style={layout.riskGrid}>
              <MetricCard
                title="Conta"
                value={selectedRiskAccount.account}
                subtitle={`${selectedRiskAccount.typeLabel} • ${selectedRiskAccount.category}`}
              />
              <MetricCard
                title="Saldo atual"
                value={formatCurrency(selectedRiskAccount.balanceCurrent)}
                color={
                  selectedRiskAccount.balanceCurrent >= 0
                    ? "#22c55e"
                    : "#ef4444"
                }
              />
              <MetricCard
                title="Liquidation Threshold"
                value={formatCurrency(
                  selectedRiskAccount.liquidationThresholdCurrent
                )}
                subtitle={`Base: ${formatCurrency(
                  selectedRiskAccount.liquidationThresholdBase
                )}`}
                color={
                  selectedRiskAccount.status === "Reprovada"
                    ? "#ef4444"
                    : selectedRiskAccount.status === "Crítica"
                    ? "#fb923c"
                    : selectedRiskAccount.status === "Atenção"
                    ? "#f59e0b"
                    : "#22c55e"
                }
              />
              <MetricCard
                title="Drawdown diário"
                value={formatCurrency(selectedRiskAccount.dailyDrawdown)}
                color={
                  selectedRiskAccount.status === "Reprovada"
                    ? "#ef4444"
                    : selectedRiskAccount.status === "Crítica"
                    ? "#fb923c"
                    : selectedRiskAccount.status === "Atenção"
                    ? "#f59e0b"
                    : "#22c55e"
                }
              />
              <MetricCard
                title="Distância da meta"
                value={formatCurrency(selectedRiskAccount.distanceToTarget)}
                color={
                  selectedRiskAccount.distanceToTarget <= 0
                    ? "#22c55e"
                    : "#93c5fd"
                }
              />
              <MetricCard
                title="Risco/dia sugerido"
                value={formatCurrency(selectedRiskAccount.riskPerDay)}
                subtitle={`Status: ${selectedRiskAccount.status}`}
                color={
                  selectedRiskAccount.status === "Reprovada"
                    ? "#ef4444"
                    : selectedRiskAccount.status === "Crítica"
                    ? "#fb923c"
                    : selectedRiskAccount.status === "Atenção"
                    ? "#f59e0b"
                    : "#22c55e"
                }
              />
            </div>

            {selectedRiskAccount.typeLabel !== "PA" ? (
              <div style={styles.progressWrap}>
                <div style={styles.progressHeader}>
                  <span>Progresso até a meta</span>
                  <span>
                    {formatNumber(selectedRiskAccount.progressToTarget)}%
                  </span>
                </div>
                <div style={styles.progressBarBg}>
                  <div
                    style={{
                      ...styles.progressBarFill,
                      width: `${selectedRiskAccount.progressToTarget}%`,
                    }}
                  />
                </div>
              </div>
            ) : null}
          </div>
        ) : null}

        <div style={layout.chartGrid}>
          <div style={styles.panel}>
            <div style={styles.panelTitle}>Resultado por dia</div>
            <div style={{ width: "100%", height: 280 }}>
              <ResponsiveContainer>
                <BarChart data={byDay}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#243042" />
                  <XAxis dataKey="day" stroke="#94a3b8" />
                  <YAxis stroke="#94a3b8" />
                  <Tooltip
                    formatter={(value) => formatCurrency(Number(value ?? 0))}
                    contentStyle={{
                      backgroundColor: "#0f172a",
                      border: "1px solid #334155",
                      borderRadius: 12,
                      color: "#ffffff",
                    }}
                    labelStyle={{ color: "#93c5fd" }}
                    itemStyle={{ color: "#ffffff" }}
                  />
                  <Bar dataKey="total" radius={[8, 8, 0, 0]}>
                    {byDay.map((entry, index) => (
                      <Cell
                        key={`cell-${index}`}
                        fill={entry.total >= 0 ? "#22c55e" : "#ef4444"}
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div style={styles.panel}>
            <div style={styles.panelTitle}>Distribuição gains x losses</div>
            <div style={{ width: "100%", height: 280 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie
                    data={pieData}
                    dataKey="value"
                    nameKey="name"
                    innerRadius={65}
                    outerRadius={95}
                    paddingAngle={4}
                  >
                    {pieData.map((entry, index) => (
                      <Cell
                        key={entry.name}
                        fill={COLORS[index % COLORS.length]}
                      />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value) => Number(value ?? 0)}
                    contentStyle={{
                      backgroundColor: "#0f172a",
                      border: "1px solid #334155",
                      borderRadius: 12,
                      color: "#ffffff",
                    }}
                    labelStyle={{ color: "#93c5fd" }}
                    itemStyle={{ color: "#ffffff" }}
                  />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <div style={styles.legendRow}>
              <span style={styles.legendItem}>
                <span
                  style={{ ...styles.legendDot, background: "#22c55e" }}
                />{" "}
                Gains
              </span>
              <span style={styles.legendItem}>
                <span
                  style={{ ...styles.legendDot, background: "#ef4444" }}
                />{" "}
                Losses
              </span>
            </div>
          </div>
        </div>

        <div style={layout.tableGrid}>
          <div style={styles.panel}>
            <div style={styles.panelTitle}>Resumo por conta</div>
            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Conta</th>
                    <th style={styles.th}>Categoria</th>
                    <th style={styles.th}>Tipo</th>
                    <th style={styles.th}>Líquido</th>
                    <th style={styles.th}>Balance</th>
                    <th style={styles.th}>Liquidation Threshold</th>
                    <th style={styles.th}>Drawdown diário</th>
                    <th style={styles.th}>Meta</th>
                    <th style={styles.th}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRiskAccounts.map((row) => (
                    <tr key={row.account}>
                      <td style={styles.td}>{row.account}</td>
                      <td style={styles.td}>{row.category}</td>
                      <td style={styles.td}>{row.typeLabel}</td>
                      <td
                        style={{
                          ...styles.td,
                          color: row.net >= 0 ? "#22c55e" : "#ef4444",
                        }}
                      >
                        {formatCurrency(row.net)}
                      </td>
                      <td style={styles.td}>
                        {formatCurrency(row.balanceCurrent)}
                      </td>
                      <td style={styles.td}>
                        {formatCurrency(row.liquidationThresholdCurrent)}
                      </td>
                      <td style={styles.td}>
                        {formatCurrency(row.dailyDrawdown)}
                      </td>
                      <td style={styles.td}>
                        {formatCurrency(row.distanceToTarget)}
                      </td>
                      <td
                        style={{
                          ...styles.td,
                          color:
                            row.status === "Ativa"
                              ? "#22c55e"
                              : row.status === "Atenção"
                              ? "#f59e0b"
                              : row.status === "Crítica"
                              ? "#fb923c"
                              : "#ef4444",
                        }}
                      >
                        {row.status}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div style={styles.panel}>
            <div style={styles.panelTitle}>Histórico de trades</div>
            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Arquivo</th>
                    <th style={styles.th}>Data</th>
                    <th style={styles.th}>Conta</th>
                    <th style={styles.th}>Símbolo</th>
                    <th style={styles.th}>Pontos</th>
                    <th style={styles.th}>Profit</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTrades
                    .slice()
                    .sort((a, b) =>
                      String(b.importedAt || "").localeCompare(
                        String(a.importedAt || "")
                      )
                    )
                    .slice(0, 200)
                    .map((trade) => (
                      <tr
                        key={`${trade.orderId}-${trade.createdOn}-${trade.importedAt}-${trade.importedFileName}`}
                      >
                        <td style={styles.td}>
                          {trade.importedFileName || "-"}
                        </td>
                        <td style={styles.td}>{trade.day}</td>
                        <td style={styles.td}>{trade.account}</td>
                        <td style={styles.td}>{trade.symbol}</td>
                        <td style={styles.td}>{formatNumber(trade.points)}</td>
                        <td
                          style={{
                            ...styles.td,
                            color: trade.profit >= 0 ? "#22c55e" : "#ef4444",
                          }}
                        >
                          {formatCurrency(trade.profit)}
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </div>

          <div style={styles.panel}>
            <div style={styles.panelTitle}>Payoff por dia</div>
            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Data</th>
                    <th style={styles.th}>Gains</th>
                    <th style={styles.th}>Losses</th>
                    <th style={styles.th}>Média gain</th>
                    <th style={styles.th}>Média loss</th>
                    <th style={styles.th}>Payoff</th>
                    <th style={styles.th}>Qualidade</th>
                    <th style={styles.th}>Score</th>
                    <th style={styles.th}>Líquido</th>
                  </tr>
                </thead>
                <tbody>
                  {payoffByDay.map((row) => {
                    const quality = getPayoffQuality(
                      row.payoff,
                      row.gains,
                      row.losses
                    );
                    const score = getDayScore(
                      row.payoff,
                      row.net,
                      row.gains,
                      row.losses
                    );

                    return (
                      <tr key={row.day}>
                        <td style={styles.td}>{row.day}</td>
                        <td style={styles.td}>{row.gains}</td>
                        <td style={styles.td}>{row.losses}</td>
                        <td style={{ ...styles.td, color: "#22c55e" }}>
                          {row.gains > 0 ? formatCurrency(row.avgGain) : "-"}
                        </td>
                        <td style={{ ...styles.td, color: "#ef4444" }}>
                          {row.losses > 0 ? formatCurrency(row.avgLoss) : "-"}
                        </td>
                        <td
                          style={{
                            ...styles.td,
                            color:
                              row.payoffLabel === "∞"
                                ? "#22c55e"
                                : row.payoff !== null && row.payoff >= 1
                                ? "#22c55e"
                                : "#ef4444",
                            fontWeight: 700,
                          }}
                        >
                          {row.payoffLabel}
                        </td>
                        <td style={{ ...styles.td, color: quality.color }}>
                          {quality.label}
                        </td>
                        <td
                          style={{
                            ...styles.td,
                            color:
                              score === "A+" || score === "A"
                                ? "#22c55e"
                                : score === "B"
                                ? "#84cc16"
                                : score === "C"
                                ? "#f59e0b"
                                : "#ef4444",
                            fontWeight: 700,
                          }}
                        >
                          {score}
                        </td>
                        <td
                          style={{
                            ...styles.td,
                            color: row.net >= 0 ? "#22c55e" : "#ef4444",
                          }}
                        >
                          {formatCurrency(row.net)}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}


export default function QiBoard() {
  const [user, setUser] = useState<UserData | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginLoading, setLoginLoading] = useState(false);
  const [emailLogin, setEmailLogin] = useState("");
  const [passwordLogin, setPasswordLogin] = useState("");
  const [showResetPopup, setShowResetPopup] = useState(false);
  const [resetEmail, setResetEmail] = useState("");
  const [resetLoading, setResetLoading] = useState(false);
  const [message, setMessage] = useState<{
    text: string;
    type: "success" | "error" | "info";
  } | null>(null);

  function showMessage(
    text: string,
    type: "success" | "error" | "info" = "info"
  ) {
    setMessage({ text, type });
  }

  async function handleResetPassword() {
    if (!resetEmail.trim()) {
      showMessage("Informe um email válido.", "error");
      return;
    }

    try {
      setResetLoading(true);
      await sendPasswordResetEmail(auth, resetEmail.trim());
      showMessage(
        "Se o email existir, enviamos um link para redefinição de senha.",
        "success"
      );
      setShowResetPopup(false);
      setResetEmail("");
    } catch (error: any) {
      console.error("Erro ao enviar redefinição de senha:", error);
      showMessage(
        error?.message || "Não foi possível enviar o link de redefinição.",
        "error"
      );
    } finally {
      setResetLoading(false);
    }
  }

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (firebaseUser) => {
      try {
        if (firebaseUser) {
          const appUser = await ensureAppUser(firebaseUser);

          if (!appUser.active) {
            setUser(null);
            await signOut(auth);
            showMessage("Seu acesso ainda não está ativo.", "error");
          } else {
            setUser(mapFirebaseUser(firebaseUser));
          }
        } else {
          setUser(null);
        }
      } catch (error) {
        console.error("Erro ao validar usuário:", error);
        setUser(null);
      } finally {
        setAuthLoading(false);
      }
    });

    return () => unsubscribe();
  }, []);

  async function handleGoogleLogin() {
    try {
      setLoginLoading(true);

      const result = await signInWithPopup(auth, googleProvider);
      const firebaseUser = result.user;

      const appUser = await ensureAppUser(firebaseUser);

      if (!appUser.active) {
        showMessage(
          `O usuário ${firebaseUser.email} foi cadastrado, mas ainda está inativo.`,
          "error"
        );
        await signOut(auth);
        return;
      }
    } catch (error) {
      console.error("Erro ao fazer login com Google:", error);
      showMessage("Não foi possível entrar com Google.", "error");
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleEmailLogin() {
    try {
      setLoginLoading(true);

      const result = await signInWithEmailAndPassword(
        auth,
        emailLogin,
        passwordLogin
      );

      const firebaseUser = result.user;
      const appUser = await ensureAppUser(firebaseUser);

      if (!appUser.active) {
        showMessage(
          `O usuário ${firebaseUser.email} está cadastrado, mas ainda está inativo.`,
          "error"
        );
        await signOut(auth);
        return;
      }
    } catch (error) {
      console.error("Erro ao fazer login com email e senha:", error);
      showMessage("Não foi possível entrar com email e senha.", "error");
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleLogout() {
    await signOut(auth);
  }

  if (authLoading) {
    return (
      <div style={styles.loadingScreen}>
        <div style={styles.loadingText}>Carregando sessão...</div>
      </div>
    );
  }

  if (!user) {
    return (
      <>
        <LoginScreen
          onGoogleLogin={handleGoogleLogin}
          onEmailLogin={handleEmailLogin}
          onForgotPassword={() => {
            setResetEmail(emailLogin || "");
            setShowResetPopup(true);
          }}
          loading={loginLoading}
          emailLogin={emailLogin}
          setEmailLogin={setEmailLogin}
          passwordLogin={passwordLogin}
          setPasswordLogin={setPasswordLogin}
        />

        {showResetPopup && (
          <div style={styles.modalOverlay}>
            <div style={styles.modalBox}>
              <div style={styles.modalTitle}>Redefinir senha</div>
              <p style={styles.modalText}>
                Digite seu email para receber o link de redefinição de senha.
              </p>

              <input
                type="email"
                placeholder="Digite seu email"
                value={resetEmail}
                onChange={(e) => setResetEmail(e.target.value)}
                style={styles.loginInput}
              />

              <div style={styles.modalActions}>
                <button
                  style={styles.secondaryButton}
                  onClick={() => {
                    setShowResetPopup(false);
                    setResetEmail("");
                  }}
                  disabled={resetLoading}
                >
                  Cancelar
                </button>

                <button
                  style={styles.actionButton}
                  onClick={handleResetPassword}
                  disabled={resetLoading}
                >
                  {resetLoading ? "Enviando..." : "Enviar link"}
                </button>
              </div>
            </div>
          </div>
        )}

        {message && (
          <div style={styles.popupOverlay}>
            <div
              style={{
                ...styles.popupBox,
                borderColor:
                  message.type === "success"
                    ? "#22c55e"
                    : message.type === "error"
                    ? "#ef4444"
                    : "#3b82f6",
              }}
            >
              <div style={styles.popupText}>{message.text}</div>

              <button
                style={styles.actionButton}
                onClick={() => setMessage(null)}
              >
                OK
              </button>
            </div>
          </div>
        )}
      </>
    );
  }

  return <DashboardScreen onLogout={handleLogout} user={user} />;
}

const styles: Record<string, React.CSSProperties> = {
  appShell: {
    minHeight: "100vh",
    width: "100%",
    display: "grid",
    background: "#020617",
    color: "#e2e8f0",
    fontFamily: "Arial, sans-serif",
  },
  sidebar: {
    borderRight: "1px solid #1e293b",
    background: "#0f172a",
    padding: 24,
    display: "flex",
    flexDirection: "column",
    gap: 18,
  },
  logo: {
    fontSize: 28,
    fontWeight: 700,
    color: "#f8fafc",
  },
  sidebarSub: {
    fontSize: 13,
    color: "#94a3b8",
    marginTop: 6,
  },
  userBox: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    background: "#111c2e",
    border: "1px solid #223048",
    borderRadius: 16,
    padding: 14,
  },
  avatar: {
    width: 42,
    height: 42,
    borderRadius: "50%",
    objectFit: "cover",
  },
  userName: {
    fontSize: 14,
    fontWeight: 700,
    color: "#f8fafc",
  },
  userEmail: {
    fontSize: 12,
    color: "#94a3b8",
    wordBreak: "break-word",
  },
  sideBox: {
    background: "#111c2e",
    border: "1px solid #223048",
    borderRadius: 16,
    padding: 16,
    display: "flex",
    flexDirection: "column",
    gap: 12,
  },
  sideBoxTitle: {
    fontSize: 14,
    fontWeight: 700,
    color: "#f8fafc",
  },
  uploadLabel: {
    background: "#2563eb",
    color: "white",
    padding: "12px 14px",
    borderRadius: 12,
    textAlign: "center",
    cursor: "pointer",
    fontWeight: 600,
  },
  fileName: {
    fontSize: 12,
    color: "#94a3b8",
    wordBreak: "break-word",
  },
  select: {
    padding: 12,
    borderRadius: 12,
    background: "#020617",
    color: "#e2e8f0",
    border: "1px solid #334155",
  },
  inputDark: {
    padding: 12,
    borderRadius: 12,
    background: "#020617",
    color: "#e2e8f0",
    border: "1px solid #334155",
  },
  dateFilterRow: {
    display: "grid",
    gap: 12,
  },
  secondaryButton: {
    marginTop: "auto",
    padding: 12,
    borderRadius: 12,
    background: "transparent",
    color: "#e2e8f0",
    border: "1px solid #334155",
    cursor: "pointer",
    fontWeight: 600,
  },
  actionButton: {
    padding: 12,
    borderRadius: 12,
    background: "#2563eb",
    color: "#fff",
    border: "none",
    cursor: "pointer",
    fontWeight: 600,
  },
  clearButton: {
    padding: 12,
    borderRadius: 12,
    background: "#7f1d1d",
    color: "#fff",
    border: "none",
    cursor: "pointer",
    fontWeight: 600,
  },
  main: {
    width: "100%",
    boxSizing: "border-box",
  },
  headerRow: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: 20,
  },
  title: {
    margin: 0,
    fontSize: 34,
    color: "#f8fafc",
  },
  subtitle: {
    marginTop: 6,
    color: "#94a3b8",
  },
  gridCards: {
    display: "grid",
    gap: 16,
    marginBottom: 20,
  },
  riskGrid: {
    display: "grid",
    gap: 16,
    marginTop: 16,
  },
  card: {
    background: "#0f172a",
    border: "1px solid #223048",
    borderRadius: 18,
    padding: 18,
    boxShadow: "0 8px 30px rgba(0,0,0,0.18)",
  },
  cardTitle: {
    fontSize: 13,
    color: "#94a3b8",
    marginBottom: 10,
  },
  metricValue: {
    fontSize: 28,
    fontWeight: 700,
    wordBreak: "break-word",
  },
  metricSub: {
    marginTop: 8,
    fontSize: 12,
    color: "#94a3b8",
  },
  progressWrap: {
    marginTop: 18,
  },
  progressHeader: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: 8,
    color: "#cbd5e1",
    fontSize: 13,
  },
  progressBarBg: {
    width: "100%",
    height: 12,
    borderRadius: 999,
    background: "#1e293b",
    overflow: "hidden",
  },
  progressBarFill: {
    height: "100%",
    borderRadius: 999,
    background: "linear-gradient(90deg, #2563eb, #22c55e)",
  },
  chartGrid: {
    display: "grid",
    gap: 16,
    marginBottom: 20,
  },
  panel: {
    background: "#0f172a",
    border: "1px solid #223048",
    borderRadius: 18,
    padding: 18,
  },
  panelTitle: {
    fontSize: 18,
    fontWeight: 700,
    color: "#f8fafc",
    marginBottom: 14,
  },
  legendRow: {
    display: "flex",
    justifyContent: "center",
    gap: 16,
    marginTop: -8,
  },
  legendItem: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    color: "#cbd5e1",
    fontSize: 13,
  },
  legendDot: {
    width: 10,
    height: 10,
    borderRadius: 999,
    display: "inline-block",
  },
  tableGrid: {
    display: "grid",
    gap: 16,
  },
  tableWrap: {
    overflowX: "auto",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    minWidth: 650,
  },
  th: {
    textAlign: "left",
    padding: "12px 10px",
    borderBottom: "1px solid #223048",
    color: "#94a3b8",
    fontSize: 13,
  },
  td: {
    padding: "12px 10px",
    borderBottom: "1px solid #1e293b",
    fontSize: 14,
    color: "#e2e8f0",
  },
  loginWrap: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background:
      "linear-gradient(135deg, #020617 0%, #0f172a 50%, #111827 100%)",
    padding: 20,
    fontFamily: "Arial, sans-serif",
  },
  loginPanel: {
    width: "100%",
    maxWidth: 700,
    background: "rgba(15, 23, 42, 0.95)",
    border: "1px solid #223048",
    borderRadius: 24,
    padding: 32,
    display: "flex",
    flexDirection: "column",
    gap: 14,
    boxShadow: "0 20px 60px rgba(0,0,0,0.35)",
  },
  brandBadge: {
    width: "fit-content",
    padding: "8px 14px",
    borderRadius: 999,
    background: "rgba(37, 99, 235, 0.18)",
    color: "#93c5fd",
    fontWeight: 700,
  },
  loginTitle: {
    margin: 0,
    color: "#f8fafc",
    fontSize: 28,
    lineHeight: 1.2,
  },
  loginText: {
    margin: 0,
    color: "#94a3b8",
    lineHeight: 1.6,
    fontSize: 16,
  },
  loginInput: {
    padding: 14,
    borderRadius: 14,
    border: "1px solid #334155",
    background: "#020617",
    color: "#ffffff",
    fontSize: 15,
    outline: "none",
  },
  googleButton: {
    marginTop: 8,
    padding: 14,
    borderRadius: 14,
    border: "none",
    background: "#2563eb",
    color: "white",
    fontWeight: 700,
    cursor: "pointer",
    fontSize: 15,
  },
  secondaryLoginButton: {
    marginTop: 4,
    padding: 14,
    borderRadius: 14,
    border: "1px solid #334155",
    background: "transparent",
    color: "white",
    fontWeight: 700,
    cursor: "pointer",
    fontSize: 15,
  },
  loadingScreen: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#020617",
    color: "#e2e8f0",
    fontFamily: "Arial, sans-serif",
  },
  loadingText: {
    fontSize: 24,
    fontWeight: 700,
  },
  googleLoginButton: {
  marginTop: 4,
  padding: 14,
  borderRadius: 14,
  border: "1px solid #334155",
  background: "#ffffff",
  color: "#111827",
  fontWeight: 700,
  cursor: "pointer",
  fontSize: 15,
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 10,
},

googleIconWrap: {
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  width: 20,
  height: 20,
},
  modalOverlay: {
    position: "fixed",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    backgroundColor: "rgba(0,0,0,0.65)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1000,
    padding: 16,
    boxSizing: "border-box",
  },
  modalBox: {
    width: "100%",
    maxWidth: 420,
    background: "#0f172a",
    border: "1px solid #223048",
    borderRadius: 18,
    padding: 24,
    boxShadow: "0 20px 60px rgba(0,0,0,0.45)",
  },
  modalTitle: {
    fontSize: 20,
    fontWeight: 700,
    color: "#f8fafc",
    marginBottom: 8,
  },
  modalText: {
    fontSize: 14,
    color: "#94a3b8",
    marginBottom: 16,
    lineHeight: 1.5,
  },
  modalActions: {
    display: "flex",
    gap: 12,
    marginTop: 8,
  },
  popupOverlay: {
    position: "fixed",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    backgroundColor: "rgba(0,0,0,0.55)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1100,
    padding: 16,
    boxSizing: "border-box",
  },
  popupBox: {
    width: "100%",
    maxWidth: 420,
    backgroundColor: "#0f172a",
    padding: 24,
    borderRadius: 18,
    border: "1px solid #223048",
    boxShadow: "0 20px 60px rgba(0,0,0,0.45)",
    textAlign: "center",
  },
  popupText: {
    color: "#e2e8f0",
    marginBottom: 18,
    fontSize: 15,
    lineHeight: 1.5,
  },
};