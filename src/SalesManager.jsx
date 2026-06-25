import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import "./SalesManager.css";

const PM_GOALS = {
  "William Dye": 1750000,
  "Jamie Jenkins": 1750000,
  "Andrew Painter": 1750000,
  "George Anim": 1000000,
  "John Fincher": 1500000,
  "Dani Cole": 1500000,
  "Megan Rice": 1000000,
  "Mike Harr": 500000,
};

const projectManagers = [
  {
    name: "Jamie Jenkins",
    slug: "jamiejenkins",
    image: "/pm/jamiejenkins.jpg",
    activeGoal: true,
  },
  {
    name: "Megan Rice",
    slug: "meganrice",
    image: "/pm/meganrice.jpg",
    activeGoal: true,
  },
  {
    name: "Dani Cole",
    slug: "danicole",
    image: "/pm/danicole.jpg",
    activeGoal: true,
  },
  {
    name: "John Fincher",
    slug: "johnfincher",
    image: "/pm/johnfincher.jpg",
    activeGoal: true,
  },
  {
    name: "Andrew Painter",
    slug: "andrewpainter",
    image: "/pm/andrewpainter.jpg",
    activeGoal: true,
  },
  {
    name: "George Anim",
    slug: "georgeanim",
    image: "/pm/georgeanim.jpg",
    activeGoal: true,
  },
  {
    name: "William Dye",
    slug: "williamdye",
    image: "/pm/williamdye.jpg",
    activeGoal: true,
  },
  {
    name: "Mike Harr",
    slug: "mikeharr",
    image: "/pm/mikeharr.jpg",
    activeGoal: true,
  },
];

const monthNames = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const fiscalMonths = [
  "November 2025",
  "December 2025",
  "January 2026",
  "February 2026",
  "March 2026",
  "April 2026",
  "May 2026",
  "June 2026",
  "July 2026",
  "August 2026",
  "September 2026",
  "October 2026",
];

const money = (value) =>
  Number(value || 0).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
  });

const displayPercent = (value, decimals = 1) => {
  const number = Number(value || 0);
  return `${(number * 100).toFixed(decimals)}%`;
};

function compareNumbers(current, comparison, isRate = false) {
  const currentNumber = Number(current || 0);
  const comparisonNumber = Number(comparison || 0);

  if (!comparisonNumber) {
    return {
      label: "N/A",
      className: "neutral",
      rawDifference: 0,
    };
  }

  if (isRate) {
    const pointDifference = currentNumber - comparisonNumber;

    return {
      label: `${pointDifference >= 0 ? "+" : ""}${(
        pointDifference * 100
      ).toFixed(1)} pts`,
      className: pointDifference >= 0 ? "positive" : "negative",
      rawDifference: pointDifference,
    };
  }

  const percentDifference =
    (currentNumber - comparisonNumber) / comparisonNumber;

  return {
    label: `${percentDifference >= 0 ? "+" : ""}${(
      percentDifference * 100
    ).toFixed(1)}%`,
    className: percentDifference >= 0 ? "positive" : "negative",
    rawDifference: percentDifference,
  };
}

function formatPMMonth(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return `${monthNames[value.getMonth()]} ${value.getFullYear()}`;
  }

  const text = String(value).trim();

  if (text.toUpperCase() === "YTD") return "YTD";
  if (text === "November 205") return "November 2025";

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return `${monthNames[parsed.getMonth()]} ${parsed.getFullYear()}`;
  }

  return text.replace(" 205", " 2025");
}

function getMonthSortValue(monthLabel) {
  if (monthLabel === "YTD") return 999999;

  const [month, year] = String(monthLabel).split(" ");
  const monthIndex = monthNames.indexOf(month);

  return Number(year || 0) * 100 + monthIndex;
}

function getMonthDateBounds(monthLabel) {
  const [monthName, yearText] = String(monthLabel || "").split(" ");
  const monthIndex = monthNames.indexOf(monthName);
  const year = Number(yearText);

  if (monthIndex < 0 || !year) {
    return null;
  }

  return {
    start: new Date(year, monthIndex, 1),
    end: new Date(year, monthIndex + 1, 0),
  };
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

function parseMoney(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;

  const cleaned = String(value).replace(/[$,]/g, "").trim();
  const number = Number(cleaned);

  return Number.isNaN(number) ? 0 : number;
}

function parseMetricValue(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return value > 1 ? value / 100 : value;
  }

  const text = String(value).trim();

  if (text.includes("%")) {
    const number = Number(text.replace("%", "").trim());
    return Number.isNaN(number) ? 0 : number / 100;
  }

  const number = Number(text);

  return Number.isNaN(number) ? 0 : number;
}

function dateOnly(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

const normalizeMetricLabel = (value) =>
  String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ");

const metricLabelMatches = (rowLabel, requestedMetric) => {
  const rowMetric = normalizeMetricLabel(rowLabel);
  const requested = normalizeMetricLabel(requestedMetric);

  if (rowMetric === requested) return true;

  if (requested === "closing rate" && rowMetric === "monthly closing rate") {
    return true;
  }

  return false;
};

export default function SalesManager() {
  const [salesRows, setSalesRows] = useState([]);
  const [pmRows, setPmRows] = useState([]);
  const [invoiceRows, setInvoiceRows] = useState([]);
  const [dataStatus, setDataStatus] = useState("Loading manager data...");

  useEffect(() => {
    const loadSalesFile = async () => {
      try {
        const response = await fetch(`/sales.xlsx?t=${Date.now()}`);

        if (!response.ok) {
          throw new Error("Could not find public/sales.xlsx");
        }

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, {
          type: "array",
          cellDates: true,
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          defval: "",
        });

        setSalesRows(rows);
      } catch (error) {
        setSalesRows([]);
        console.error(error);
      }
    };

    const loadPMFile = async () => {
      try {
        const response = await fetch(`/pm-contract-data.xlsx?t=${Date.now()}`);

        if (!response.ok) {
          throw new Error("Could not find public/pm-contract-data.xlsx");
        }

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, {
          type: "array",
          cellDates: true,
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
        });

        setPmRows(rows);
      } catch (error) {
        setPmRows([]);
        console.error(error);
      }
    };
    const loadInvoiceFile = async () => {
  try {
    const response = await fetch(`/Invoicing.xlsx?t=${Date.now()}`);

    if (!response.ok) {
      throw new Error("Could not find public/Invoicing.xlsx");
    }

    const buffer = await response.arrayBuffer();

    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
    });

    setInvoiceRows(rows);
  } catch (error) {
    setInvoiceRows([]);
    console.error(error);
  }
};

Promise.all([
    loadSalesFile(),
    loadPMFile(),
    loadInvoiceFile(),
]).finally(() => {
    setDataStatus("");
});
  }, []);

  const getPMBlockStart = (pmName) => {
    return pmRows.findIndex(
      (row) => String(row?.[0] || "").trim() === pmName
    );
  };

  const getTeamMetricRow = (metric) => {
    const teamSectionStart = pmRows.findIndex(
      (row) =>
        String(row?.[0] || "").trim().toLowerCase() ===
        "contract total average"
    );

    if (teamSectionStart < 0) return null;

    return pmRows
      .slice(teamSectionStart, teamSectionStart + 14)
      .find((row) => metricLabelMatches(row?.[0], metric));
  };

  const getPMMetric = (pmName, metric, monthLabel) => {
    const startIndex = getPMBlockStart(pmName);

    if (startIndex < 0 || !monthLabel) return 0;

    const headerRow = pmRows[startIndex] || [];

    const nextSectionIndex = pmRows.findIndex((row, index) => {
      if (index <= startIndex) return false;

      const label = String(row?.[0] || "").trim();

      return (
        projectManagers.some((pm) => pm.name === label) ||
        label.toLowerCase() === "contract total average"
      );
    });

    const blockRows = pmRows.slice(
      startIndex + 1,
      nextSectionIndex > startIndex ? nextSectionIndex : startIndex + 20
    );

    const metricRow = blockRows.find((row) =>
      metricLabelMatches(row?.[0], metric)
    );

    if (!metricRow) return 0;

    const columnIndex = headerRow.findIndex(
      (cell) => formatPMMonth(cell) === monthLabel
    );

    if (columnIndex < 0) return 0;

    return parseMetricValue(metricRow[columnIndex]);
  };

  const getTeamMetric = (metric, monthLabel) => {
    const row = getTeamMetricRow(metric);

    if (!row || !monthLabel) return 0;

    const headerSource =
      pmRows.find(
        (item) => String(item?.[0] || "").trim() === projectManagers[0].name
      ) || [];

    const columnIndex = headerSource.findIndex(
      (cell) => formatPMMonth(cell) === monthLabel
    );

    if (columnIndex < 0) return 0;

    return parseMetricValue(row[columnIndex]);
  };

  const getPMSalesDataForRange = (pmName, startDate, endDate) => {
    if (!pmName || !startDate || !endDate) {
      return {
        contractTotal: 0,
        contracts: 0,
        averageContract: 0,
      };
    }

    const start = dateOnly(startDate);
    const end = dateOnly(endDate);

    let contractTotal = 0;
    let contracts = 0;

    salesRows.forEach((row) => {
      const rowPMValues = [
        row["Project Manager"],
        row["Salesperson"],
        row["Sales Rep"],
        row["Sales Representative"],
        row["Primary Salesperson"],
        row["Estimator"],
      ].map((value) => String(value || "").trim());

      if (!rowPMValues.includes(pmName)) return;

      const rowDate = parseExcelDate(row["Approved Date"]);
      if (!rowDate) return;

      const cleanDate = dateOnly(rowDate);
      if (cleanDate < start || cleanDate > end) return;

      const amount = parseMoney(row["Contract Amount"]);

      if (amount <= 0) return;

      contractTotal += amount;
      contracts += 1;
    });

    return {
      contractTotal,
      contracts,
      averageContract: contracts > 0 ? contractTotal / contracts : 0,
    };
  };

const getManagerSalesReps = () =>
  projectManagers.filter((pm) => pm.activeGoal && pm.name !== "Mike Harr");

const getFiscalYearBounds = () => ({
  start: new Date(2025, 10, 1),
  end: new Date(2026, 9, 31),
});

const getCurrentQuarterBounds = () => {
  const today = dateOnly(new Date());
  const fiscalQuarterStarts = [
    new Date(2025, 10, 1),
    new Date(2026, 1, 1),
    new Date(2026, 4, 1),
    new Date(2026, 7, 1),
  ];

  const quarterStart =
    fiscalQuarterStarts
      .slice()
      .reverse()
      .find((start) => today >= start) || fiscalQuarterStarts[0];

  return {
    start: quarterStart,
    end: new Date(quarterStart.getFullYear(), quarterStart.getMonth() + 3, 0),
  };
};

const getCurrentMonthBounds = () => {
  const today = new Date();

  return {
    start: new Date(today.getFullYear(), today.getMonth(), 1),
    end: new Date(today.getFullYear(), today.getMonth() + 1, 0),
  };
};

const getManagerMonthsInRange = (startDate, endDate) =>
  fiscalMonths.filter((month) => {
    const bounds = getMonthDateBounds(month);
    if (!bounds) return false;
    return bounds.end >= startDate && bounds.start <= endDate;
  });

const getManagerClosingRateForRange = (pmName, startDate, endDate) => {
  const months = getManagerMonthsInRange(startDate, endDate);
  const rates = months
    .map((month) => getPMMetric(pmName, "Monthly Closing Rate", month))
    .filter((rate) => Number(rate || 0) > 0);

  if (!rates.length) return 0;

  return rates.reduce((sum, rate) => sum + rate, 0) / rates.length;
};

const getManagerMetrics = () => {
  const reps = getManagerSalesReps();
  const fiscalYear = getFiscalYearBounds();
  const currentMonth = getCurrentMonthBounds();
  const currentQuarter = getCurrentQuarterBounds();
  const priorFiscalYear = {
    start: new Date(
      fiscalYear.start.getFullYear() - 1,
      fiscalYear.start.getMonth(),
      fiscalYear.start.getDate()
    ),
    end: new Date(
      fiscalYear.end.getFullYear() - 1,
      fiscalYear.end.getMonth(),
      fiscalYear.end.getDate()
    ),
  };

  const rows = reps.map((pm) => {
    const annualSales = getPMSalesDataForRange(
      pm.name,
      fiscalYear.start,
      fiscalYear.end
    );
    const monthlySales = getPMSalesDataForRange(
      pm.name,
      currentMonth.start,
      currentMonth.end
    );
    const quarterlySales = getPMSalesDataForRange(
      pm.name,
      currentQuarter.start,
      currentQuarter.end
    );
    const lastYearSales = getPMSalesDataForRange(
      pm.name,
      priorFiscalYear.start,
      priorFiscalYear.end
    );
    const goal = Number(PM_GOALS[pm.name] || 0);
    const closingRate = getManagerClosingRateForRange(
      pm.name,
      fiscalYear.start,
      fiscalYear.end
    );

    return {
      name: pm.name,
      image: pm.image,
      annualRevenue: annualSales.contractTotal,
      annualContracts: annualSales.contracts,
      averageContract: annualSales.averageContract,
      monthlyRevenue: monthlySales.contractTotal,
      quarterlyRevenue: quarterlySales.contractTotal,
      lyRevenue: lastYearSales.contractTotal,
      closingRate,
      goal,
      monthlyGoal: goal / 12,
      quarterlyGoal: goal / 4,
      annualGoalPercent: goal > 0 ? annualSales.contractTotal / goal : 0,
      monthlyGoalPercent: goal > 0 ? monthlySales.contractTotal / (goal / 12) : 0,
      quarterlyGoalPercent: goal > 0 ? quarterlySales.contractTotal / (goal / 4) : 0,
    };
  });

  const sortedRows = rows.sort((a, b) => b.annualRevenue - a.annualRevenue);
  const activeRows = rows.filter((row) => row.annualRevenue > 0);
  const totalRevenue = rows.reduce((sum, row) => sum + row.annualRevenue, 0);
  const totalGoal = rows.reduce((sum, row) => sum + row.goal, 0);
  const totalContracts = rows.reduce((sum, row) => sum + row.annualContracts, 0);
  const teamAverageRevenue = activeRows.length > 0 ? totalRevenue / activeRows.length : 0;
  const teamClosingRate =
    activeRows.length > 0
      ? activeRows.reduce((sum, row) => sum + row.closingRate, 0) / activeRows.length
      : 0;
  const teamAverageContract = totalContracts > 0 ? totalRevenue / totalContracts : 0;
  const teamLYRevenue = rows.reduce((sum, row) => sum + row.lyRevenue, 0);

  return {
    rows: sortedRows.map((row, index) => ({
      ...row,
      rank: index + 1,
      revenueVsLY: compareNumbers(row.annualRevenue, row.lyRevenue),
      revenueVsTeam: compareNumbers(row.annualRevenue, teamAverageRevenue),
      closingVsTeam: compareNumbers(row.closingRate, teamClosingRate, true),
    })),
    totals: {
      totalRevenue,
      totalGoal,
      totalGoalPercent: totalGoal > 0 ? totalRevenue / totalGoal : 0,
      totalContracts,
      teamAverageRevenue,
      teamClosingRate,
      teamAverageContract,
      teamLYRevenue,
      teamVsLY: compareNumbers(totalRevenue, teamLYRevenue),
    },
    ranges: {
      currentMonth,
      currentQuarter,
      fiscalYear,
    },
  };
};
const getInvoicePipeline = () => {
  const pipeline = {
    scheduledForBuild: 0,
    needsToBeInvoiced: 0,
    balanceDue: 0,
    totalPipeline: 0,
  };

  invoiceRows.forEach((row) => {
    const values = Object.values(row);

    values.forEach((cell, index) => {
      const label = String(cell || "").trim().toLowerCase();
      const nextValue = parseMoney(values[index + 1]);

      if (label.includes("scheduled for build")) {
        pipeline.scheduledForBuild += nextValue;
      }

      if (label.includes("needs to be invoiced")) {
        pipeline.needsToBeInvoiced += nextValue;
      }

      if (label.includes("balance due") || label.includes("money invoiced")) {
        pipeline.balanceDue += nextValue;
      }

      if (label === "total") {
        pipeline.totalPipeline += nextValue;
      }
    });
  });

  if (!pipeline.totalPipeline) {
    pipeline.totalPipeline =
      pipeline.scheduledForBuild +
      pipeline.needsToBeInvoiced +
      pipeline.balanceDue;
  }

  return pipeline;
};
const ProgressBar = ({ value }) => {
  const cappedValue = Math.min(Math.max(Number(value || 0), 0), 1.25);

  return (
    <div className="manager-progress-bar">
      <div style={{ width: `${Math.min(cappedValue, 1) * 100}%` }} />
    </div>
  );
};

const SalesManagerDashboard = () => {
  const managerMetrics = getManagerMetrics();
  const { rows, totals } = managerMetrics;
  const invoicePipeline = getInvoicePipeline();

  return (
    <div className="sales-manager-page">
      <header className="manager-header">
        <img src="/logo.png" alt="Logo" className="manager-logo" />
        <div>
          <span>Sales Manager Portal</span>
          <h1>Sales Performance Command Center</h1>
          <p>Revenue, closing rate, year-over-year movement, team averages, and goal progress.</p>
        </div>
      </header>

      <section className="manager-kpi-grid">
        <div className="manager-kpi-card primary">
          <span>Team Revenue</span>
          <strong>{money(totals.totalRevenue)}</strong>
          <small className={`manager-difference ${totals.teamVsLY.className}`}>
            {totals.teamVsLY.label} vs LY
          </small>
        </div>

        <div className="manager-kpi-card">
          <span>Team Closing Rate</span>
          <strong>{displayPercent(totals.teamClosingRate, 1)}</strong>
          <small>Active rep average</small>
        </div>

        <div className="manager-kpi-card">
          <span>Annual Goal Progress</span>
          <strong>{displayPercent(totals.totalGoalPercent, 1)}</strong>
          <ProgressBar value={totals.totalGoalPercent} />
        </div>

        <div className="manager-kpi-card">
          <span>Average Contract</span>
          <strong>{money(totals.teamAverageContract)}</strong>
          <small>{totals.totalContracts} contracts</small>
        </div>
      </section>

<section className="manager-panel">
  <div className="manager-panel-header">
    <div>
      <span>Cash Flow</span>
      <h2>Future Cash Pipeline</h2>
    </div>
    <p>Expected future payments from Invoicing.xlsx.</p>
  </div>

  <div className="manager-kpi-grid">
    <div className="manager-kpi-card">
      <span>Scheduled for Build</span>
      <strong>{money(invoicePipeline.scheduledForBuild)}</strong>
      <small>Future production money</small>
    </div>

    <div className="manager-kpi-card">
      <span>Needs to be Invoiced</span>
      <strong>{money(invoicePipeline.needsToBeInvoiced)}</strong>
      <small>Ready to invoice</small>
    </div>

    <div className="manager-kpi-card">
      <span>Balance Due</span>
      <strong>{money(invoicePipeline.balanceDue)}</strong>
      <small>Already invoiced / owed</small>
    </div>

    <div className="manager-kpi-card primary">
      <span>Total Future Cash</span>
      <strong>{money(invoicePipeline.totalPipeline)}</strong>
      <small>Build + invoice + receivables</small>
    </div>
  </div>
</section>

      <section className="manager-panel">
        <div className="manager-panel-header">
          <div>
            <span>Leaderboard</span>
            <h2>Salesperson Scoreboard</h2>
          </div>
          <p>Ranked by annual fiscal-year revenue.</p>
        </div>

        <div className="manager-table-wrap">
          <table className="manager-table">
            <thead>
              <tr>
                <th>Rank</th>
                <th>Salesperson</th>
                <th>Revenue</th>
                <th>Close Rate</th>
                <th>vs LY</th>
                <th>vs Team Avg</th>
                <th>Annual Goal</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.name}>
                  <td>#{row.rank}</td>
                  <td>
                    <div className="manager-rep-cell">
                      <img src={row.image} alt={row.name} />
                      <strong>{row.name}</strong>
                    </div>
                  </td>
                  <td>{money(row.annualRevenue)}</td>
                  <td>
                    {displayPercent(row.closingRate, 1)}
                    <div className={`manager-difference ${row.closingVsTeam.className}`}>
                      {row.closingVsTeam.label} vs team
                    </div>
                  </td>
                  <td>
                    <div className={`manager-difference ${row.revenueVsLY.className}`}>
                      {row.revenueVsLY.label}
                    </div>
                    <small>{money(row.lyRevenue)} LY</small>
                  </td>
                  <td>
                    <div className={`manager-difference ${row.revenueVsTeam.className}`}>
                      {row.revenueVsTeam.label}
                    </div>
                    <small>{money(totals.teamAverageRevenue)} avg</small>
                  </td>
                  <td>
                    <strong>{displayPercent(row.annualGoalPercent, 1)}</strong>
                    <ProgressBar value={row.annualGoalPercent} />
                    <small>{money(row.goal)}</small>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="manager-goal-grid">
        {rows.map((row) => (
          <div className="manager-goal-card" key={`${row.name}-goals`}>
            <div className="manager-goal-title">
              <img src={row.image} alt={row.name} />
              <h3>{row.name}</h3>
            </div>

            <div className="manager-goal-row">
              <span>Monthly</span>
              <strong>{money(row.monthlyRevenue)}</strong>
              <b>{displayPercent(row.monthlyGoalPercent, 1)}</b>
              <ProgressBar value={row.monthlyGoalPercent} />
            </div>

            <div className="manager-goal-row">
              <span>Quarterly</span>
              <strong>{money(row.quarterlyRevenue)}</strong>
              <b>{displayPercent(row.quarterlyGoalPercent, 1)}</b>
              <ProgressBar value={row.quarterlyGoalPercent} />
            </div>

            <div className="manager-goal-row">
              <span>Annual</span>
              <strong>{money(row.annualRevenue)}</strong>
              <b>{displayPercent(row.annualGoalPercent, 1)}</b>
              <ProgressBar value={row.annualGoalPercent} />
            </div>
          </div>
        ))}
      </section>
    </div>
  );
};



  return (
    <div className="page manager-shell">
      {dataStatus && <div className="manager-status">{dataStatus}</div>}
      <SalesManagerDashboard />
    </div>
  );
}
