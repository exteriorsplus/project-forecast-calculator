import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const UNKNOWN_KEY = "Unknown-Work Type";
const UNKNOWN_RPP = 15191.27;
const UNKNOWN_MARGIN = 0.277;
const COMPANY_EXPENSE_RATE = 0.1;

const categories = [
  {
    title: "Roofing",
    className: "roofing",
    items: [
      { label: "Insurance", rpp: 16691.43, margin: 0.28218532 },
      { label: "Repair", rpp: 2352.01, margin: 0.35404618 },
      { label: "Retail", rpp: 13064.05, margin: 0.23808109 },
    ],
  },
  {
    title: "Siding",
    className: "siding",
    items: [
      { label: "Insurance", rpp: 18816.72, margin: 0.35401739 },
      { label: "Repair", rpp: 1870.49, margin: 0.33731542 },
      { label: "Retail", rpp: 22143.99, margin: 0.24288058 },
    ],
  },
  {
    title: "Roofing & Siding",
    className: "combo",
    items: [
      { label: "Insurance", rpp: 34034.1, margin: 0.26774496 },
      { label: "Repair", rpp: 5727.9, margin: 0.2741249 },
      { label: "Retail", rpp: 33258.86, margin: 0.22728198 },
    ],
  },
];

const money = (value) =>
  value.toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
  });

const percent = (value) => `${(value * 100).toFixed(4)}%`;

const formatDate = (date) => date.toISOString().slice(0, 10);

function getMonthDateRange() {
  const today = new Date();
  const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  return {
    start: formatDate(firstOfMonth),
    end: formatDate(today),
  };
}

function AnimatedMoney({ value }) {
  const [displayValue, setDisplayValue] = useState(value);
  const frameRef = useRef(null);

  useEffect(() => {
    const isMobile =
      typeof window !== "undefined" &&
      window.matchMedia("(max-width: 768px)").matches;

    if (isMobile) {
      setDisplayValue(value);
      return;
    }

    const start = displayValue;
    const end = value;
    const duration = 250;
    const startTime = performance.now();

    const animate = (time) => {
      const progress = Math.min((time - startTime) / duration, 1);
      const current = start + (end - start) * progress;

      setDisplayValue(current);

      if (progress < 1) {
        frameRef.current = requestAnimationFrame(animate);
      }
    };

    frameRef.current = requestAnimationFrame(animate);

    return () => {
      if (frameRef.current) {
        cancelAnimationFrame(frameRef.current);
      }
    };
  }, [value]);

  return <>{money(displayValue)}</>;
}

function parseExcelDate(value) {
  if (!value) return null;

  if (value instanceof Date) return value;

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const date = new Date(String(value).trim());
  return Number.isNaN(date.getTime()) ? null : date;
}

function dateOnly(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function normalizeTrade(value) {
  const text = String(value || "").trim().toLowerCase();

  if (text === "roofing") return "Roofing";
  if (text === "siding") return "Siding";
  if (text === "roofing & siding") return "Roofing & Siding";

  if (text.includes("roofing") && text.includes("siding")) {
    return "Roofing & Siding";
  }

  if (text.includes("roofing")) return "Roofing";
  if (text.includes("siding")) return "Siding";

  return "";
}

function normalizeWorkType(value) {
  const text = String(value || "").trim().toLowerCase();

  if (text.includes("insurance")) return "Insurance";
  if (text.includes("repair")) return "Repair";
  if (text.includes("retail")) return "Retail";

  return "";
}

function getProfitBreakdown(revenue, companyMargin) {
  const trueProjectMargin = companyMargin + COMPANY_EXPENSE_RATE;
  const trueProjectProfit = revenue * trueProjectMargin;
  const companyExpense = revenue * COMPANY_EXPENSE_RATE;
  const companyProfit = revenue * companyMargin;

  return {
    trueProjectMargin,
    trueProjectProfit,
    companyExpense,
    companyProfit,
  };
}

export default function App() {
  const [screen, setScreen] = useState("home");

  const [leads, setLeads] = useState({});
  const [projects, setProjects] = useState({});
  const [closeRate, setCloseRate] = useState(0.25);
  const [customCloseRate, setCustomCloseRate] = useState("");
  const [flash, setFlash] = useState(false);

  const [leadRows, setLeadRows] = useState([]);
  const [dataStatus, setDataStatus] = useState("Loading leads.xlsx...");

  const initialDateRange = getMonthDateRange();
  const [startDate, setStartDate] = useState(initialDateRange.start);
  const [endDate, setEndDate] = useState(initialDateRange.end);

  const flashTimeoutRef = useRef(null);

  const rawRate =
    closeRate === "custom" ? Number(customCloseRate || 0) / 100 : closeRate;

  const activeCloseRate = Math.min(Math.max(rawRate, 0), 1);

  const triggerFlash = () => {
    setFlash(true);

    if (flashTimeoutRef.current) {
      clearTimeout(flashTimeoutRef.current);
    }

    flashTimeoutRef.current = setTimeout(() => {
      setFlash(false);
    }, 200);
  };

  const update = (setter, state, key, value) => {
    setter({ ...state, [key]: value === "" ? "" : Number(value) });
  };

  const clearLeads = () => {
    const cleared = {};

    categories.forEach((category) => {
      category.items.forEach((item) => {
        cleared[`${category.title}-${item.label}`] = 0;
      });
    });

    cleared[UNKNOWN_KEY] = 0;

    setLeads(cleared);
    triggerFlash();
  };

  const clearProjects = () => {
    const cleared = {};

    categories.forEach((category) => {
      category.items.forEach((item) => {
        cleared[`${category.title}-${item.label}`] = 0;
      });
    });

    setProjects(cleared);
    triggerFlash();
  };

  const loadLeadFile = async () => {
    try {
      setDataStatus("Loading leads.xlsx...");

      const response = await fetch(`/leads.xlsx?t=${Date.now()}`);

      if (!response.ok) {
        throw new Error("Could not find public/leads.xlsx");
      }

      const buffer = await response.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
      });

      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
      });

      setLeadRows(rows);
      setDataStatus(`${rows.length} rows loaded from leads.xlsx`);
    } catch (error) {
      setLeadRows([]);
      setDataStatus("No leads.xlsx file found yet.");
      console.error(error);
    }
  };

  const applyLeadCounts = (rows = leadRows) => {
    const start = dateOnly(new Date(startDate));
    const end = dateOnly(new Date(endDate));

    const counts = {};

    categories.forEach((category) => {
      category.items.forEach((item) => {
        counts[`${category.title}-${item.label}`] = 0;
      });
    });

    counts[UNKNOWN_KEY] = 0;

    rows.forEach((row) => {
      const rowDate = parseExcelDate(
        row["Initial Appointment Date"] ||
          row["Prospect Milestone Date"] ||
          row["Closed Milestone Date"]
      );

      if (!rowDate) return;

      const cleanDate = dateOnly(rowDate);
      if (cleanDate < start || cleanDate > end) return;

      const rawTrade = String(row["Job Trade Type 2"] || "")
        .trim()
        .toLowerCase();

      let trade = "";

      if (rawTrade === "roofing") {
        trade = "Roofing";
      } else if (rawTrade === "siding") {
        trade = "Siding";
      } else if (rawTrade === "roofing & siding") {
        trade = "Roofing & Siding";
      } else {
        trade = normalizeTrade(row["Job Trade Type"]);
      }

      const workType = normalizeWorkType(row["Work Type"]);

      if (!workType) {
        counts[UNKNOWN_KEY] += 1;
        return;
      }

      const key = `${trade}-${workType}`;

      if (Object.prototype.hasOwnProperty.call(counts, key)) {
        counts[key] += 1;
      }
    });

    setLeads(counts);
  };

  const getLeadTotals = () => {
    let totalLeads = Number(leads[UNKNOWN_KEY] || 0);
    let revenue = totalLeads * activeCloseRate * UNKNOWN_RPP;

    const unknownBreakdown = getProfitBreakdown(revenue, UNKNOWN_MARGIN);

    let trueProjectProfit = unknownBreakdown.trueProjectProfit;
    let companyExpense = unknownBreakdown.companyExpense;
    let companyProfit = unknownBreakdown.companyProfit;

    categories.forEach((category) => {
      category.items.forEach((item) => {
        const key = `${category.title}-${item.label}`;
        const quantity = Number(leads[key] || 0);
        const itemRevenue = quantity * activeCloseRate * item.rpp;
        const breakdown = getProfitBreakdown(itemRevenue, item.margin);

        totalLeads += quantity;
        revenue += itemRevenue;
        trueProjectProfit += breakdown.trueProjectProfit;
        companyExpense += breakdown.companyExpense;
        companyProfit += breakdown.companyProfit;
      });
    });

    return {
      totalLeads,
      revenue,
      trueProjectProfit,
      companyExpense,
      companyProfit,
    };
  };

  const getProjectTotals = () => {
    let revenue = 0;
    let trueProjectProfit = 0;
    let companyExpense = 0;
    let companyProfit = 0;

    categories.forEach((category) => {
      category.items.forEach((item) => {
        const key = `${category.title}-${item.label}`;
        const quantity = Number(projects[key] || 0);
        const itemRevenue = quantity * item.rpp;
        const breakdown = getProfitBreakdown(itemRevenue, item.margin);

        revenue += itemRevenue;
        trueProjectProfit += breakdown.trueProjectProfit;
        companyExpense += breakdown.companyExpense;
        companyProfit += breakdown.companyProfit;
      });
    });

    return {
      revenue,
      trueProjectProfit,
      companyExpense,
      companyProfit,
    };
  };

  useEffect(() => {
    loadLeadFile();

    return () => {
      if (flashTimeoutRef.current) {
        clearTimeout(flashTimeoutRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!leadRows.length || !startDate || !endDate) return;

    const timeout = setTimeout(() => {
      applyLeadCounts(leadRows);
    }, 150);

    return () => clearTimeout(timeout);
  }, [leadRows, startDate, endDate]);

  const leadTotals = getLeadTotals();
  const projectTotals = getProjectTotals();
  const unknownLeadCount = Number(leads[UNKNOWN_KEY] || 0);

  const Header = () => (
    <header className="header">
      {screen !== "home" && (
        <button className="back-button-header" onClick={() => setScreen("home")}>
          ← Home
        </button>
      )}

      <div className="header-top">
        <img src="/logo.png" alt="Logo" className="logo" />
        <h1>Lead & Project Forecast Calculator</h1>
      </div>
    </header>
  );

  if (screen === "home") {
    return (
      <div className="page">
        <Header />

        <section className="home-grid">
          <button className="home-card" onClick={() => setScreen("leads")}>
            <span>Lead Forecast</span>
            <strong>Forecast revenue from lead counts</strong>
            <p>Uses close rate × revenue per project × margin.</p>
          </button>

          <button
            className="home-card projects-card"
            onClick={() => setScreen("projects")}
          >
            <span>Project Forecast</span>
            <strong>Forecast revenue from project counts</strong>
            <p>Uses confirmed project count × revenue per project × margin.</p>
          </button>
        </section>
      </div>
    );
  }

  return (
    <div className="page">
      <Header />

      {screen === "leads" && (
        <section className="calculator-section">
          <div className="data-panel">
            <div className="date-controls">
              <label>
                Start Date
                <input
                  type="date"
                  value={startDate}
                  onChange={(event) => setStartDate(event.target.value)}
                />
              </label>

              <label>
                End Date
                <input
                  type="date"
                  value={endDate}
                  onChange={(event) => setEndDate(event.target.value)}
                />
              </label>

              <button onClick={() => applyLeadCounts()}>Apply Lead Counts</button>

              <button className="clear-button" onClick={clearLeads}>
                Clear Leads
              </button>

              <button onClick={loadLeadFile}>Reload File</button>
            </div>
          </div>

          <div className="grid">
            {categories.map((category) => (
              <div className={`card ${category.className}`} key={category.title}>
                <h3>{category.title}</h3>

                {category.items.map((item) => {
                  const key = `${category.title}-${item.label}`;
                  const quantity = Number(leads[key] || 0);
                  const revenue = quantity * activeCloseRate * item.rpp;
                  const breakdown = getProfitBreakdown(revenue, item.margin);

                  return (
                    <div className="row" key={key}>
                      <div>
                        <strong>{item.label}</strong>
                        <div className="sub-label">
                          Rev/Project: {money(item.rpp)}
                        </div>
                        <div className="sub-label">
                          True Margin: {percent(breakdown.trueProjectMargin)}
                        </div>
                        <div className="sub-label">
                          Company Margin: {percent(item.margin)}
                        </div>
                      </div>

                      <input
                        type="number"
                        min="0"
                        placeholder="Leads"
                        value={leads[key] ?? ""}
                        onChange={(event) =>
                          update(setLeads, leads, key, event.target.value)
                        }
                      />

                      <div className="return">
  <div className="metric">
    <span>Revenue</span>
    <strong>{money(revenue)}</strong>
  </div>

  <div className="metric true">
    <span>True Profit</span>
    <strong>{money(breakdown.trueProjectProfit)}</strong>
  </div>

  <div className="metric company">
    <span>Company Profit</span>
    <strong>{money(breakdown.companyProfit)}</strong>
  </div>
</div>
                    </div>
                  );
                })}
              </div>
            ))}
          </div>

          <div className="lead-control-grid">
            <div className="unknown-leads-panel">
              <h3>Leads - Unknown Work Type</h3>
              <p>Using {money(UNKNOWN_RPP)} avg revenue/project</p>
              <p>True Margin: {percent(UNKNOWN_MARGIN + COMPANY_EXPENSE_RATE)}</p>
              <p>Company Margin: {percent(UNKNOWN_MARGIN)}</p>
              <strong>{unknownLeadCount}</strong>

              <div className="unknown-divider"></div>

              <h3>Total Leads</h3>
              <strong>{leadTotals.totalLeads}</strong>
            </div>

            <div className="close-rate-panel">
              <h3>Lead Close Rate</h3>
              <p>
                Lead totals are calculated as Leads × Close Rate × Revenue /
                Project.
              </p>

              <div className="rate-buttons">
                {[0.2, 0.25, 0.3, 0.35].map((rate) => (
                  <button
                    key={rate}
                    className={
                      closeRate === rate ? "rate-button active" : "rate-button"
                    }
                    onClick={() => setCloseRate(rate)}
                  >
                    {(rate * 100).toFixed(0)}%
                  </button>
                ))}

                <div className="other-rate">
                  <button
                    className={
                      closeRate === "custom"
                        ? "rate-button active"
                        : "rate-button"
                    }
                    onClick={() => setCloseRate("custom")}
                  >
                    Other
                  </button>

                  <input
                    type="number"
                    min="0"
                    max="100"
                    placeholder="%"
                    value={customCloseRate}
                    onChange={(event) => {
                      let value = Number(event.target.value);
                      if (value > 100) value = 100;
                      if (value < 0) value = 0;

                      setCloseRate("custom");
                      setCustomCloseRate(value);
                    }}
                  />
                </div>
              </div>

              <div className="active-rate">
                Close Rate: {(activeCloseRate * 100).toFixed(0)}%
              </div>
            </div>
          </div>

          <div className="section-summary-grid">
            <div className={`section-total red ${flash ? "glow" : ""}`}>
              <strong>
                <AnimatedMoney value={leadTotals.revenue} />
              </strong>
              <span>Lead Revenue</span>
            </div>

            <div className={`section-total white profit-total ${flash ? "glow" : ""}`}>
              <strong>
                <AnimatedMoney value={leadTotals.trueProjectProfit} />
              </strong>
              <span>True Project Profit</span>

              <div className="profit-breakdown">
                <div className="expense-line">
                  Company Expense 10%: -{money(leadTotals.companyExpense)}
                </div>
                <div>
                  Company Profit: {money(leadTotals.companyProfit)}
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {screen === "projects" && (
        <section className="calculator-section">
          <h2 className="section-title">Projects Forecast</h2>

          <div className="projects-controls">
            <button className="apply-button" onClick={clearProjects}>
              Clear Project Counts
            </button>
          </div>

          <div className="grid">
            {categories.map((category) => (
              <div className={`card ${category.className}`} key={category.title}>
                <h3>{category.title}</h3>

                {category.items.map((item) => {
                  const key = `${category.title}-${item.label}`;
                  const quantity = Number(projects[key] || 0);
                  const revenue = quantity * item.rpp;
                  const breakdown = getProfitBreakdown(revenue, item.margin);

                  return (
                    <div className="row" key={key}>
                      <div>
                        <strong>{item.label}</strong>
                        <div className="sub-label">
                          Rev/Project: {money(item.rpp)}
                        </div>
                        <div className="sub-label">
                          True Margin: {percent(breakdown.trueProjectMargin)}
                        </div>
                        <div className="sub-label">
                          Company Margin: {percent(item.margin)}
                        </div>
                      </div>

                      <input
                        type="number"
                        min="0"
                        placeholder="Projects"
                        value={projects[key] ?? ""}
                        onChange={(event) =>
                          update(setProjects, projects, key, event.target.value)
                        }
                      />

<div className="return">
  <div className="metric">
    <span>Revenue</span>
    <strong>{money(revenue)}</strong>
  </div>

  <div className="metric true">
    <span>True Profit</span>
    <strong>{money(breakdown.trueProjectProfit)}</strong>
  </div>

  <div className="metric company">
    <span>Company Profit</span>
    <strong>{money(breakdown.companyProfit)}</strong>
  </div>
</div>
                    </div>
                  );
                })}
              </div>
            ))}
          </div>

          <div className="section-summary-grid">
            <div className={`section-total red ${flash ? "glow" : ""}`}>
              <strong>
                <AnimatedMoney value={projectTotals.revenue} />
              </strong>
              <span>Project Revenue</span>
            </div>

            <div className={`section-total white profit-total ${flash ? "glow" : ""}`}>
              <strong>
                <AnimatedMoney value={projectTotals.trueProjectProfit} />
              </strong>
              <span>True Project Profit</span>

              <div className="profit-breakdown">
                <div className="expense-line">
                  Company Expense 10%: -{money(projectTotals.companyExpense)}
                </div>
                <div>
                  Company Profit: {money(projectTotals.companyProfit)}
                </div>
              </div>
            </div>
          </div>
        </section>
      )}
    </div>
  );
}