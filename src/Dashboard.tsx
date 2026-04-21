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
  type User,
} from "firebase/auth";
import { auth, googleProvider } from "./firebase/config";

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
};

const COLORS = ["#22c55e", "#ef4444"];
const STORAGE_KEYS = {
  trades: "qboard_trades",
  fileName: "qboard_file_name",
  selectedAccount: "qboard_selected_account",
  search: "qboard_search",
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

function detectAccountConfig(account: string): AccountConfig {
  if (account.startsWith("PA-")) return ACCOUNT_RULES.PA;

  if (account.includes("-303") || account.includes("-304") || account.includes("-305")) {
    return ACCOUNT_RULES.APEX150;
  }

  if (
    account.includes("-301") ||
    account.includes("-302") ||
    account.includes("-300")
  ) {
    return ACCOUNT_RULES.APEX50;
  }

  return ACCOUNT_RULES.APEX25;
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

  const cleaned = value.trim().replace(/\$/g, "").replace(/\s/g, "").replace(/,/g, "");
  if (!cleaned) return NaN;
  return Number(cleaned);
}

function getDayLabel(row: RawTradeRow): string {
  if (row.created_on) {
    const d = new Date(row.created_on);
    if (!Number.isNaN(d.getTime())) return d.toLocaleDateString("pt-BR");
  }

  if (row.mov_time) {
    const d = new Date(row.mov_time);
    if (!Number.isNaN(d.getTime())) return d.toLocaleDateString("pt-BR");
  }

  return "Sem data";
}

function buildTrades(rows: RawTradeRow[]): Trade[] {
  return rows
    .map((row) => {
      const profit = parseMoney(row.profit);

      return {
        account: row.name?.trim() || "Sem conta",
        orderId: row.order_id?.trim() || "",
        symbol: row.symbol?.trim() || "",
        movTime: row.mov_time?.trim() || "",
        movType: String(row.mov_type ?? ""),
        qty: Number(row.exec_qty ?? 0),
        price: Number(row.price_done ?? 0),
        points: Number(row.points ?? 0),
        profit,
        createdOn: row.created_on?.trim() || "",
        day: getDayLabel(row),
      };
    })
    .filter((trade) => !Number.isNaN(trade.profit));
}

function mapFirebaseUser(user: User): UserData {
  return {
    name: user.displayName || "Usuário",
    email: user.email || "",
    photo: user.photoURL || "",
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
  loading,
}: {
  onGoogleLogin: () => void;
  loading: boolean;
}) {
  return (
    <div style={styles.loginWrap}>
      <div style={styles.loginPanel}>
        <div style={styles.brandBadge}>QBoard</div>
        <h1 style={styles.loginTitle}>Painel profissional para contas proprietárias</h1>
        <p style={styles.loginText}>
          Faça login com sua conta Google para acessar o dashboard, importar
          planilhas e acompanhar a performance das suas contas.
        </p>

        <button style={styles.googleButton} onClick={onGoogleLogin} disabled={loading}>
          {loading ? "Entrando..." : "Entrar com Google"}
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
  const [selectedAccount, setSelectedAccount] = useState<string>("Todas");
  const [search, setSearch] = useState("");
  const [fileName, setFileName] = useState("");
  const [screenWidth, setScreenWidth] = useState<number>(window.innerWidth);

  useEffect(() => {
    function handleResize() {
      setScreenWidth(window.innerWidth);
    }
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    const savedTrades = localStorage.getItem(STORAGE_KEYS.trades);
    const savedFileName = localStorage.getItem(STORAGE_KEYS.fileName);
    const savedSelectedAccount = localStorage.getItem(STORAGE_KEYS.selectedAccount);
    const savedSearch = localStorage.getItem(STORAGE_KEYS.search);

    if (savedTrades) {
      try {
        setTrades(JSON.parse(savedTrades));
      } catch {
        localStorage.removeItem(STORAGE_KEYS.trades);
      }
    }

    if (savedFileName) setFileName(savedFileName);
    if (savedSelectedAccount) setSelectedAccount(savedSelectedAccount);
    if (savedSearch) setSearch(savedSearch);
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.trades, JSON.stringify(trades));
  }, [trades]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.fileName, fileName);
  }, [fileName]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.selectedAccount, selectedAccount);
  }, [selectedAccount]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.search, search);
  }, [search]);

  const accounts = useMemo(() => {
    return Array.from(new Set(trades.map((trade) => trade.account))).sort();
  }, [trades]);

  const filteredTrades = useMemo(() => {
    return trades.filter((trade) => {
      const accountOk =
        selectedAccount === "Todas" || trade.account === selectedAccount;

      const text =
        `${trade.account} ${trade.symbol} ${trade.day} ${trade.orderId}`.toLowerCase();

      const searchOk = text.includes(search.toLowerCase());

      return accountOk && searchOk;
    });
  }, [trades, selectedAccount, search]);

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
      const balanceCurrent =
        cfg.typeLabel === "PA" ? row.net : cfg.balanceStart + row.net;

      const drawdownAvailable =
        cfg.typeLabel === "PA"
          ? 0
          : balanceCurrent - (cfg.balanceStart - cfg.trailingDrawdown);

      const distanceToTarget =
        cfg.typeLabel === "PA" ? 0 : cfg.profitTarget - row.net;

      const progressToTarget =
        cfg.typeLabel === "PA"
          ? 0
          : Math.max(0, Math.min(100, (row.net / cfg.profitTarget) * 100));

      const riskPerDay =
        cfg.typeLabel === "PA"
          ? 0
          : Math.max(0, drawdownAvailable * 0.05);

      let status = "Saudável";
      if (cfg.typeLabel !== "PA") {
        if (drawdownAvailable <= 0) status = "Falhou";
        else if (drawdownAvailable < cfg.trailingDrawdown * 0.2) status = "Crítico";
        else if (drawdownAvailable < cfg.trailingDrawdown * 0.4) status = "Atenção";
      }

      return {
        account: row.account,
        typeLabel: cfg.typeLabel,
        balanceStart: cfg.balanceStart,
        balanceCurrent,
        drawdownAvailable,
        trailingDrawdown: cfg.trailingDrawdown,
        target: cfg.profitTarget,
        distanceToTarget,
        progressToTarget,
        riskPerDay,
        status,
        net: row.net,
        positive: row.positive,
        negative: row.negative,
        trades: row.trades,
      };
    });
  }, [byAccount]);

  const selectedRiskAccount = useMemo(() => {
    if (selectedAccount !== "Todas") {
      return riskByAccount.find((r) => r.account === selectedAccount) || null;
    }
    return riskByAccount[0] || null;
  }, [riskByAccount, selectedAccount]);

  const byDay = useMemo(() => {
    const map = new Map<string, number>();

    filteredTrades.forEach((trade) => {
      map.set(trade.day, (map.get(trade.day) || 0) + trade.profit);
    });

    return Array.from(map.entries()).map(([day, total]) => ({ day, total }));
  }, [filteredTrades]);

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
        gridTemplateColumns: isMobile ? "1fr" : "280px 1fr",
      } as React.CSSProperties,
      gridCards: {
        ...styles.gridCards,
        gridTemplateColumns: isMobile
          ? "1fr"
          : isNotebook
          ? "repeat(3, minmax(0, 1fr))"
          : "repeat(6, minmax(0, 1fr))",
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
    };
  }, [screenWidth]);

  function handleParsedRows(rows: RawTradeRow[], importedFileName: string) {
    const parsedTrades = buildTrades(rows);
    setTrades(parsedTrades);
    setFileName(importedFileName);
  }

  function handleCsvFile(file: File) {
    Papa.parse<RawTradeRow>(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        handleParsedRows(results.data || [], file.name);
      },
    });
  }

  function handleXlsxFile(file: File) {
    const reader = new FileReader();

    reader.onload = (event) => {
      const data = event.target?.result;
      if (!data) return;

      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json<RawTradeRow>(sheet, { defval: "" });

      handleParsedRows(rows, file.name);
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
    const summaryRows = riskByAccount.map((row) => ({
      Conta: row.account,
      Tipo: row.typeLabel,
      SaldoInicial: row.balanceStart,
      SaldoAtual: row.balanceCurrent,
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

    const wb = XLSX.utils.book_new();
    const wsResumo = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsResumo, "Risco e Metas");

    const excelBuffer = XLSX.write(wb, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(blob, "qboard-risco-metas.xlsx");
  }

  function clearLocalData() {
    localStorage.removeItem(STORAGE_KEYS.trades);
    localStorage.removeItem(STORAGE_KEYS.fileName);
    localStorage.removeItem(STORAGE_KEYS.selectedAccount);
    localStorage.removeItem(STORAGE_KEYS.search);

    setTrades([]);
    setFileName("");
    setSelectedAccount("Todas");
    setSearch("");
  }

  return (
    <div style={layout.appShell}>
      <aside style={layout.sidebar}>
        <div>
          <div style={styles.logo}>QBoard</div>
          <div style={styles.sidebarSub}>Dashboard para qualquer prop firm</div>
        </div>

        <div style={styles.userBox}>
          {user.photo ? <img src={user.photo} alt={user.name} style={styles.avatar} /> : null}
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
            {fileName || "Nenhum arquivo importado"}
          </div>
        </div>

        <div style={styles.sideBox}>
          <div style={styles.sideBoxTitle}>Filtro</div>
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

          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Buscar conta, símbolo, data..."
            style={styles.inputDark}
          />
        </div>

        <div style={styles.sideBox}>
          <div style={styles.sideBoxTitle}>Ações</div>
          <button style={styles.actionButton} onClick={exportSummaryExcel}>
            Exportar risco e metas
          </button>
          <button style={styles.clearButton} onClick={clearLocalData}>
            Limpar dados locais
          </button>
        </div>

        <button style={styles.secondaryButton} onClick={onLogout}>
          Sair
        </button>
      </aside>

      <main style={layout.main}>
        <div style={layout.headerRow}>
          <div>
            <h1 style={styles.title}>QBoard - Painel Principal</h1>
            <p style={styles.subtitle}>
              Resumo profissional de performance, risco, metas e execução.
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
            title="Payoff"
            value={formatNumber(metrics.payoff)}
            subtitle="ganho médio ÷ perda média"
          />
        </div>

        {selectedRiskAccount ? (
          <div style={styles.panel}>
            <div style={styles.panelTitle}>Metas e risco da conta</div>
            <div style={layout.riskGrid}>
              <MetricCard
                title="Conta"
                value={selectedRiskAccount.account}
                subtitle={selectedRiskAccount.typeLabel}
              />
              <MetricCard
                title="Saldo atual"
                value={formatCurrency(selectedRiskAccount.balanceCurrent)}
                color={selectedRiskAccount.balanceCurrent >= 0 ? "#22c55e" : "#ef4444"}
              />
              <MetricCard
                title="Drawdown disponível"
                value={formatCurrency(selectedRiskAccount.drawdownAvailable)}
                color={
                  selectedRiskAccount.status === "Crítico"
                    ? "#ef4444"
                    : selectedRiskAccount.status === "Atenção"
                    ? "#f59e0b"
                    : "#22c55e"
                }
              />
              <MetricCard
                title="Distância da meta"
                value={formatCurrency(selectedRiskAccount.distanceToTarget)}
                color={selectedRiskAccount.distanceToTarget <= 0 ? "#22c55e" : "#93c5fd"}
              />
              <MetricCard
                title="Risco/dia sugerido"
                value={formatCurrency(selectedRiskAccount.riskPerDay)}
                subtitle={`Status: ${selectedRiskAccount.status}`}
                color={
                  selectedRiskAccount.status === "Falhou"
                    ? "#ef4444"
                    : selectedRiskAccount.status === "Crítico"
                    ? "#f59e0b"
                    : "#22c55e"
                }
              />
            </div>

            {selectedRiskAccount.typeLabel !== "PA" ? (
              <div style={styles.progressWrap}>
                <div style={styles.progressHeader}>
                  <span>Progresso até a meta</span>
                  <span>{formatNumber(selectedRiskAccount.progressToTarget)}%</span>
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
                    }}
                  />
                  <Bar dataKey="total" fill="#3b82f6" radius={[8, 8, 0, 0]} />
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
                      <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />
                    ))}
                  </Pie>
                  <Tooltip
                    contentStyle={{
                      backgroundColor: "#0f172a",
                      border: "1px solid #334155",
                      borderRadius: 12,
                    }}
                  />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <div style={styles.legendRow}>
              <span style={styles.legendItem}>
                <span style={{ ...styles.legendDot, background: "#22c55e" }} /> Gains
              </span>
              <span style={styles.legendItem}>
                <span style={{ ...styles.legendDot, background: "#ef4444" }} /> Losses
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
                    <th style={styles.th}>Tipo</th>
                    <th style={styles.th}>Líquido</th>
                    <th style={styles.th}>Drawdown disp.</th>
                    <th style={styles.th}>Meta</th>
                    <th style={styles.th}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {riskByAccount.map((row) => (
                    <tr key={row.account}>
                      <td style={styles.td}>{row.account}</td>
                      <td style={styles.td}>{row.typeLabel}</td>
                      <td style={{ ...styles.td, color: row.net >= 0 ? "#22c55e" : "#ef4444" }}>
                        {formatCurrency(row.net)}
                      </td>
                      <td style={styles.td}>{formatCurrency(row.drawdownAvailable)}</td>
                      <td style={styles.td}>{formatCurrency(row.distanceToTarget)}</td>
                      <td
                        style={{
                          ...styles.td,
                          color:
                            row.status === "Saudável"
                              ? "#22c55e"
                              : row.status === "Atenção"
                              ? "#f59e0b"
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
            <div style={styles.panelTitle}>Últimos trades com profit</div>
            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Data</th>
                    <th style={styles.th}>Conta</th>
                    <th style={styles.th}>Símbolo</th>
                    <th style={styles.th}>Pontos</th>
                    <th style={styles.th}>Profit</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTrades
                    .slice(-12)
                    .reverse()
                    .map((trade) => (
                      <tr key={`${trade.orderId}-${trade.createdOn}-${trade.profit}`}>
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
        </div>
      </main>
    </div>
  );
}

export default function QBoard() {
  const [user, setUser] = useState<UserData | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginLoading, setLoginLoading] = useState(false);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (firebaseUser) => {
      if (firebaseUser) {
        setUser(mapFirebaseUser(firebaseUser));
      } else {
        setUser(null);
      }
      setAuthLoading(false);
    });

    return () => unsubscribe();
  }, []);

  async function handleGoogleLogin() {
    try {
      setLoginLoading(true);
      await signInWithPopup(auth, googleProvider);
    } catch (error) {
      console.error("Erro ao fazer login com Google:", error);
      alert("Não foi possível entrar com Google.");
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
    return <LoginScreen onGoogleLogin={handleGoogleLogin} loading={loginLoading} />;
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
    background: "linear-gradient(135deg, #020617 0%, #0f172a 50%, #111827 100%)",
    padding: 20,
    fontFamily: "Arial, sans-serif",
  },
  loginPanel: {
    width: "100%",
    maxWidth: 460,
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
    fontSize: 34,
    lineHeight: 1.1,
  },
  loginText: {
    margin: 0,
    color: "#94a3b8",
    lineHeight: 1.6,
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
};