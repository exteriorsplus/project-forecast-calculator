import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import "./SalesManager.css";

const MANAGER_PASSWORD = "Lakeview2910";
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
const REPAIR_SPECIALISTS = [
  "George Anim",
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

const formatShortDate = (date) =>
  date
    ? date.toLocaleDateString("en-US", {
        month: "numeric",
        day: "numeric",
        year: "2-digit",
      })
    : "";

export default function SalesManager() {
  const [salesRows, setSalesRows] = useState([]);
  const [pmRows, setPmRows] = useState([]);
  const [invoiceRows, setInvoiceRows] = useState([]);
  const [dataStatus, setDataStatus] = useState("Loading manager data...");
  const [managerAuthorized, setManagerAuthorized] = useState(
  () => sessionStorage.getItem("managerAuthorized") === "true"
);
const [managerPassword, setManagerPassword] = useState("");
const [managerLoginError, setManagerLoginError] = useState("");
const [activeSummaryTab, setActiveSummaryTab] = useState("Monthly Production");
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

    Promise.all([loadSalesFile(), loadPMFile(), loadInvoiceFile()]).finally(
      () => {
        setDataStatus("");
      }
    );
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

  const getCompanySalesDataForRange = (startDate, endDate) => {
    if (!startDate || !endDate) {
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

  const getInvoicePipelineReps = () =>
    projectManagers.filter((pm) => pm.activeGoal);

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

const getLatestPMMetric = (pmName, metric) => {
  const availableMonths = fiscalMonths.filter(
    (month) => Number(getPMMetric(pmName, metric, month) || 0) > 0
  );

  const latestMonth = availableMonths[availableMonths.length - 1];

  return latestMonth ? getPMMetric(pmName, metric, latestMonth) : 0;
};

const getCurrentPMMonthLabel = () => {
  const today = new Date();
  return `${monthNames[today.getMonth()]} ${today.getFullYear()}`;
};

const getTeamMetricForMonth = (metric, monthLabel) => {
  const requestedMetric = normalizeMetricLabel(metric);

  const teamRowIndex = pmRows.findIndex((row) =>
    row.some((cell) => normalizeMetricLabel(cell) === requestedMetric)
  );

  if (teamRowIndex < 0) return 0;

  const teamRow = pmRows[teamRowIndex];

  let headerRow = null;

  for (let index = teamRowIndex; index >= 0; index -= 1) {
    const possibleHeader = pmRows[index] || [];

    if (possibleHeader.some((cell) => formatPMMonth(cell) === monthLabel)) {
      headerRow = possibleHeader;
      break;
    }
  }

  if (!headerRow) return 0;

  const columnIndex = headerRow.findIndex(
    (cell) => formatPMMonth(cell) === monthLabel
  );

  if (columnIndex < 0) return 0;

  return parseMetricValue(teamRow[columnIndex]);
};
  const getManagerMetrics = () => {
    const reps = getManagerSalesReps();

    const fiscalYear = getFiscalYearBounds();
    const currentMonth = getCurrentMonthBounds();
    const currentQuarter = getCurrentQuarterBounds();
    const today = dateOnly(new Date());

    const fiscalYTD = {
      start: fiscalYear.start,
      end: today > fiscalYear.end ? fiscalYear.end : today,
    };

    const priorFiscalYTD = {
      start: new Date(
        fiscalYTD.start.getFullYear() - 1,
        fiscalYTD.start.getMonth(),
        fiscalYTD.start.getDate()
      ),
      end: new Date(
        fiscalYTD.end.getFullYear() - 1,
        fiscalYTD.end.getMonth(),
        fiscalYTD.end.getDate()
      ),
    };

    const rolling90End = today;
    const rolling90Start = new Date(rolling90End);
    rolling90Start.setDate(rolling90Start.getDate() - 89);

    const quarterMTD = {
      start: currentQuarter.start,
      end: today > currentQuarter.end ? currentQuarter.end : today,
    };

    const rows = reps.map((pm) => {
      const annualSales = getPMSalesDataForRange(
        pm.name,
        fiscalYTD.start,
        fiscalYTD.end
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

      const quarterMTDSales = getPMSalesDataForRange(
        pm.name,
        quarterMTD.start,
        quarterMTD.end
      );

      const rolling90Sales = getPMSalesDataForRange(
        pm.name,
        rolling90Start,
        rolling90End
      );

      const lastYearSales = getPMSalesDataForRange(
        pm.name,
        priorFiscalYTD.start,
        priorFiscalYTD.end
      );

      const goal = Number(PM_GOALS[pm.name] || 0);

      const closingRate = getManagerClosingRateForRange(
        pm.name,
        fiscalYTD.start,
        fiscalYTD.end
      );

      const lyClosingRate = getManagerClosingRateForRange(
        pm.name,
        priorFiscalYTD.start,
        priorFiscalYTD.end
      );

      const monthlyClosingRate = getManagerClosingRateForRange(
        pm.name,
        currentMonth.start,
        currentMonth.end
      );

      const quarterlyClosingRate = getManagerClosingRateForRange(
        pm.name,
        currentQuarter.start,
        currentQuarter.end
      );

      const quarterMTDClosingRate = getManagerClosingRateForRange(
        pm.name,
        quarterMTD.start,
        quarterMTD.end
      );

const rolling90ClosingRate = getLatestPMMetric(
  pm.name,
  "Rolling 90-Day Closing Rate"
);

return {
  name: pm.name,
  image: pm.image,

  annualRevenue: annualSales.contractTotal,
  annualContracts: annualSales.contracts,
  averageContract: annualSales.averageContract,

  monthlyRevenue: monthlySales.contractTotal,
  monthlyContracts: monthlySales.contracts,
  monthlyAverageContract: monthlySales.averageContract,

  quarterlyRevenue: quarterlySales.contractTotal,
  quarterlyContracts: quarterlySales.contracts,
  quarterlyAverageContract: quarterlySales.averageContract,

  quarterMTDRevenue: quarterMTDSales.contractTotal,
  rolling90Revenue: rolling90Sales.contractTotal,
  lyRevenue: lastYearSales.contractTotal,

  closingRate,
  lyClosingRate,
  monthlyClosingRate,
  quarterlyClosingRate,
  quarterMTDClosingRate,
  rolling90ClosingRate,

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

const trackedTeamRevenue = rows.reduce(
  (sum, row) => sum + row.annualRevenue,
  0
);

const companyFYTD = getCompanySalesDataForRange(
  fiscalYTD.start,
  fiscalYTD.end
);

const companyLYFYTD = getCompanySalesDataForRange(
  priorFiscalYTD.start,
  priorFiscalYTD.end
);

const totalRevenue = companyFYTD.contractTotal;
const totalContracts = companyFYTD.contracts;
const teamAverageContract = companyFYTD.averageContract;

const totalGoal = rows.reduce((sum, row) => sum + row.goal, 0);

const teamAverageRevenue =
  activeRows.length > 0 ? trackedTeamRevenue / activeRows.length : 0;

const teamClosingRate =
  activeRows.length > 0
    ? activeRows.reduce((sum, row) => sum + row.closingRate, 0) /
      activeRows.length
    : 0;

const teamLYRevenue = companyLYFYTD.contractTotal;

return {
  rows: sortedRows.map((row, index) => ({
    ...row,
    rank: index + 1,
    revenueVsLY: compareNumbers(row.annualRevenue, row.lyRevenue),
    closingVsLY: compareNumbers(row.closingRate, row.lyClosingRate, true),
    revenueVsTeam: compareNumbers(row.annualRevenue, teamAverageRevenue),
    closingVsTeam: compareNumbers(row.closingRate, teamClosingRate, true),
  })),
  totals: {
    totalRevenue,
    totalGoal,
    totalGoalPercent: totalGoal > 0 ? trackedTeamRevenue / totalGoal : 0,
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
    quarterMTD,
    fiscalYear,
    fiscalYTD,
    priorFiscalYTD,
    rolling90: {
      start: rolling90Start,
      end: rolling90End,
    },
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

   if (label.includes("money invoiced")) {
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

  const normalizeInvoicePMName = (value) => {
    const text = String(value || "").trim().toLowerCase();

    if (text === "will" || text === "william") return "William Dye";
    if (text === "jamie") return "Jamie Jenkins";
    if (text === "andrew") return "Andrew Painter";
    if (text === "george") return "George Anim";
    if (text === "john") return "John Fincher";
    if (text === "dani") return "Dani Cole";
    if (text === "megan") return "Megan Rice";
    if (text === "mike") return "Mike Harr";

    return "";
  };

  const getInvoiceCell = (row, index) => {
    const key = index === 0 ? "__EMPTY" : `__EMPTY_${index}`;

    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key];
    }

    return Object.values(row)[index];
  };

  const getInvoicePipelineByRep = () => {
    const totalsByRep = {};

    getInvoicePipelineReps().forEach((rep) => {
      totalsByRep[rep.name] = {
        name: rep.name,
        image: rep.image,
        scheduledForBuild: 0,
        needsToBeInvoiced: 0,
        balanceDue: 0,
        totalPipeline: 0,
      };
    });

    invoiceRows.forEach((row) => {
      const scheduledPM = normalizeInvoicePMName(getInvoiceCell(row, 2));
      const scheduledAmount = parseMoney(getInvoiceCell(row, 4));

      if (scheduledPM && totalsByRep[scheduledPM]) {
        totalsByRep[scheduledPM].scheduledForBuild += scheduledAmount;
      }

      const needsInvoicePM = normalizeInvoicePMName(getInvoiceCell(row, 8));
      const needsInvoiceAmount = parseMoney(getInvoiceCell(row, 10));

      if (needsInvoicePM && totalsByRep[needsInvoicePM]) {
        totalsByRep[needsInvoicePM].needsToBeInvoiced += needsInvoiceAmount;
      }

      const balanceDuePM = normalizeInvoicePMName(getInvoiceCell(row, 15));
      const balanceDueAmount = parseMoney(getInvoiceCell(row, 20));

      if (balanceDuePM && totalsByRep[balanceDuePM]) {
        totalsByRep[balanceDuePM].balanceDue += balanceDueAmount;
      }
    });

    return Object.values(totalsByRep)
      .map((rep) => ({
        ...rep,
        totalPipeline:
          rep.scheduledForBuild + rep.needsToBeInvoiced + rep.balanceDue,
      }))
      .sort((a, b) => b.totalPipeline - a.totalPipeline);
  };

  const ProgressBar = ({ value }) => {
    const cappedValue = Math.min(Math.max(Number(value || 0), 0), 1.25);

    return (
      <div className="manager-progress-bar">
        <div style={{ width: `${Math.min(cappedValue, 1) * 100}%` }} />
      </div>
    );
  };

  const MetricMiniBlock = ({ label, value, difference, subValue }) => (
    <div className="manager-mini-metric">
      <span>{label}</span>
      <strong>{value}</strong>
      {difference && (
        <div className={`manager-difference ${difference.className}`}>
          {difference.label}
        </div>
      )}
      {subValue && <small>{subValue}</small>}
    </div>
  );

  const GoalThermometer = ({ label, amount, percentValue }) => (
    <div className="manager-thermometer-row">
      <div className="manager-thermometer-topline">
        <span>{label}</span>
        <strong>{money(amount)}</strong>
        <b>{displayPercent(percentValue, 1)}</b>
      </div>
      <ProgressBar value={percentValue} />
    </div>
  );

  const getRepStatus = (row) => {
    const revenueWellAboveLY = row.revenueVsLY.rawDifference >= 0.15;
    const revenueAboveTeam = row.revenueVsTeam.rawDifference >= 0.15;
    const closingAboveTeam = row.closingVsTeam.rawDifference >= 0;
    const quarterStrong = row.quarterlyGoalPercent >= 1;
    const monthStrong = row.monthlyGoalPercent >= 1;
    const annualHealthy = row.annualGoalPercent >= 0.5;
    const annualWeak = row.annualGoalPercent < 0.35;
    const closingWeak = row.closingVsTeam.rawDifference <= -0.03;
    const revenueBehindLY = row.revenueVsLY.rawDifference <= -0.1;

    if (
      revenueWellAboveLY &&
      revenueAboveTeam &&
      quarterStrong &&
      (closingAboveTeam || monthStrong)
    ) {
      return { label: "Exceeding", className: "positive" };
    }

    if ((annualHealthy && !closingWeak) || quarterStrong || monthStrong) {
      return { label: "On Pace", className: "positive" };
    }

    if (annualWeak || (revenueBehindLY && closingWeak)) {
      return { label: "Needs Coaching", className: "negative" };
    }

    return { label: "Watch", className: "neutral" };
  };

  const getManagerSynopsis = (row) => {
    const notes = [];

    if (row.revenueVsLY.className === "positive") {
      notes.push(
        `Revenue is ${row.revenueVsLY.label} versus the same fiscal period last year`
      );
    } else if (row.revenueVsLY.className === "negative") {
      notes.push(
        `Revenue is ${row.revenueVsLY.label} versus the same fiscal period last year`
      );
    } else {
      notes.push("Prior-year revenue comparison is limited");
    }

    if (row.revenueVsTeam.className === "positive") {
      notes.push(`${row.revenueVsTeam.label} above the team revenue average`);
    } else if (row.revenueVsTeam.className === "negative") {
      notes.push(`${row.revenueVsTeam.label} below the team revenue average`);
    }

    if (row.closingVsTeam.className === "positive") {
      notes.push(`close rate is ${row.closingVsTeam.label} above team average`);
    } else if (row.closingVsTeam.className === "negative") {
      notes.push(
        `close rate trails the team average by ${row.closingVsTeam.label.replace("-", "")}`
      );
    } else {
      notes.push("close rate is roughly aligned with team average");
    }

    if (row.quarterlyGoalPercent >= 1) {
      notes.push("quarterly goal has already been exceeded");
    } else if (row.monthlyGoalPercent >= 1) {
      notes.push("monthly pace is ahead of goal");
    } else if (row.annualGoalPercent < 0.35) {
      notes.push("annual goal progress needs attention");
    } else {
      notes.push("goal progress should be monitored closely through the next month");
    }

    return `${notes.join(", ")}.`;
  };

const getExecutiveSummary = (rows, totals, invoicePipeline) => {
  const activeRows = rows.filter((row) => Number(row.annualRevenue || 0) > 0);

  const revenueCoachingRows = activeRows.filter(
    (row) => !REPAIR_SPECIALISTS.includes(row.name)
  );

  const hour = new Date().getHours();

  const greeting =
    hour < 12
      ? "Good Morning Mike -"
      : hour < 17
      ? "Good Afternoon Mike -"
      : "Good Evening Mike -";

  const getHighest = (list, key) =>
    list
      .slice()
      .filter((row) => Number(row[key] || 0) > 0)
      .sort((a, b) => Number(b[key] || 0) - Number(a[key] || 0))[0];
const getLeaders = (list, key) => {
  const values = list
    .map((row) => Number(row[key] || 0))
    .filter((value) => value > 0);

  if (!values.length) return [];

  const max = Math.max(...values);

  return list.filter((row) => Number(row[key] || 0) === max);
};

const formatLeaderNames = (leaders) => {
  if (leaders.length === 1) return leaders[0].name;

  if (leaders.length === 2) {
    return `${leaders[0].name} and ${leaders[1].name}`;
  }

  return `${leaders
    .slice(0, -1)
    .map((row) => row.name)
    .join(", ")}, and ${leaders[leaders.length - 1].name}`;
};
  const getLowest = (list, key) =>
    list
      .slice()
      .filter((row) => Number(row[key] || 0) > 0)
      .sort((a, b) => Number(a[key] || 0) - Number(b[key] || 0))[0];

  const getAverage = (list, key) => {
    const values = list
      .map((row) => Number(row[key] || 0))
      .filter((value) => value > 0);

    if (!values.length) return 0;

    return values.reduce((sum, value) => sum + value, 0) / values.length;
  };

  const getLargestRevenueGap = (list, revenueKey) => {
    const values = list.filter((row) => Number(row[revenueKey] || 0) > 0);
    const average = getAverage(values, revenueKey);

    if (!values.length || !average) return null;

    const row = values
      .slice()
      .sort((a, b) => Number(a[revenueKey] || 0) - Number(b[revenueKey] || 0))[0];

    return {
      row,
      average,
      gap: Number(row[revenueKey] || 0) - average,
    };
  };

  const formatPointGap = (value, teamAverage) => {
    const gap = Number(value || 0) - Number(teamAverage || 0);
    return `${Math.abs(gap * 100).toFixed(1)} pts ${gap >= 0 ? "above" : "below"} team average`;
  };

  const getStatus = ({ good, warning }) => {
    if (good) return "good";
    if (warning) return "watch";
    return "attention";
  };

  const statusLabel = {
    good: "ahead of plan",
    watch: "slightly behind plan",
    attention: "needs immediate attention",
  };

  const monthlyRevenueLeader = getHighest(activeRows, "monthlyRevenue");
  const monthlyContractsLeaders = getLeaders(activeRows, "monthlyContracts");
  const monthlyAverageContractLeader = getHighest(activeRows, "monthlyAverageContract");
  const monthlyGoalLeader = getHighest(activeRows, "monthlyGoalPercent");
  const monthlyRevenueLow = getLowest(revenueCoachingRows, "monthlyRevenue");
  const monthlyGoalLow = getLowest(activeRows, "monthlyGoalPercent");
  const monthlyAverageContractLow = getLowest(revenueCoachingRows, "monthlyAverageContract");
  const monthlyAtGoalCount = activeRows.filter(
    (row) => Number(row.monthlyGoalPercent || 0) >= 1
  ).length;

  const rolling90ClosingLeader = getHighest(activeRows, "rolling90ClosingRate");
  const rolling90RevenueLeader = getHighest(activeRows, "rolling90Revenue");
  const rolling90ClosingLow = getLowest(activeRows, "rolling90ClosingRate");
const currentPMMonth = getCurrentPMMonthLabel();

const rolling90ClosingAverage = getTeamMetricForMonth(
  "Rolling 90-Day Closing Rate Average",
  currentPMMonth
);
  const rolling90RevenueGap = getLargestRevenueGap(revenueCoachingRows, "rolling90Revenue");

  const quarterlyRevenueLeader = getHighest(activeRows, "quarterMTDRevenue");
  const quarterlyGoalLeader = getHighest(activeRows, "quarterlyGoalPercent");
  const quarterlyGoalLow = getLowest(activeRows, "quarterlyGoalPercent");
  const quarterlyRevenueGap = getLargestRevenueGap(revenueCoachingRows, "quarterMTDRevenue");
  const quarterlyAtGoalCount = activeRows.filter(
    (row) => Number(row.quarterlyGoalPercent || 0) >= 1
  ).length;

  const annualRevenueLeader = getHighest(activeRows, "annualRevenue");
  const annualClosingLeader = getHighest(activeRows, "closingRate");
  const annualGoalLeader = getHighest(activeRows, "annualGoalPercent");
  const annualClosingLow = getLowest(activeRows, "closingRate");
  const annualGoalLow = getLowest(activeRows, "annualGoalPercent");
  const annualRevenueLow = revenueCoachingRows
    .slice()
    .sort(
      (a, b) =>
        a.revenueVsTeam.rawDifference - b.revenueVsTeam.rawDifference
    )[0];
  const annualClosingAverage = getAverage(activeRows, "closingRate");
  const aboveTeamRevenue = activeRows
    .filter((row) => row.revenueVsTeam.className === "positive")
    .sort((a, b) => b.revenueVsTeam.rawDifference - a.revenueVsTeam.rawDifference);

  const averageRolling90Revenue = getAverage(activeRows, "rolling90Revenue");

  const weeklyTopDawgCandidates = activeRows
    .map((row) => {
      const annualGoalScore = Math.min(Number(row.annualGoalPercent || 0), 1.5) * 35;
      const quarterlyGoalScore = Math.min(Number(row.quarterlyGoalPercent || 0), 1.5) * 25;
      const closingScore = rolling90ClosingAverage
        ? Math.min(Number(row.rolling90ClosingRate || 0) / rolling90ClosingAverage, 1.5) * 25
        : 0;
      const rollingRevenueScore = averageRolling90Revenue
        ? Math.min(Number(row.rolling90Revenue || 0) / averageRolling90Revenue, 1.5) * 15
        : 0;

      return {
        ...row,
        weeklyTopDawgScore:
          annualGoalScore + quarterlyGoalScore + closingScore + rollingRevenueScore,
      };
    })
    .sort((a, b) => b.weeklyTopDawgScore - a.weeklyTopDawgScore);

  const weeklyTopDawg = weeklyTopDawgCandidates[0];

  const getWeeklyTopDawgReasons = (row) => {
    if (!row) return [];

    const reasons = [];
    const teamRolling90Gap =
      Number(row.rolling90ClosingRate || 0) - Number(rolling90ClosingAverage || 0);

    if (annualRevenueLeader?.name === row.name) {
      reasons.push(`Leads FYTD production at ${money(row.annualRevenue)}`);
    }

    if (rolling90ClosingLeader?.name === row.name) {
      reasons.push(`Leads the team in Rolling 90-Day Closing Rate at ${displayPercent(row.rolling90ClosingRate, 1)}`);
    }

    if (rolling90RevenueLeader?.name === row.name) {
      reasons.push(`Leads Rolling 90-Day revenue at ${money(row.rolling90Revenue)}`);
    }

    if (Number(row.quarterlyGoalPercent || 0) >= 1) {
      reasons.push(`Is above quarterly goal pace at ${displayPercent(row.quarterlyGoalPercent, 1)}`);
    }

    if (annualGoalLeader?.name === row.name) {
      reasons.push(`Has the strongest annual goal progress at ${displayPercent(row.annualGoalPercent, 1)}`);
    }

    if (teamRolling90Gap > 0) {
      reasons.push(`Is ${Math.abs(teamRolling90Gap * 100).toFixed(1)} pts above the team Rolling 90-Day Closing Rate average`);
    }

    if (Number(row.rank || 0) > 0 && Number(row.rank || 0) <= 3) {
      reasons.push(`Currently ranks #${row.rank} in FYTD production`);
    }

    if (Number(row.revenueVsTeam?.rawDifference || 0) > 0) {
      reasons.push(`Is ${row.revenueVsTeam.label} above the team revenue average`);
    }

    if (!reasons.length) {
      reasons.push("Shows the strongest all-around score across revenue, goal pace, and Rolling 90-Day performance");
    }
    const backupReasons = [
      `Has produced ${money(row.quarterMTDRevenue)} quarter-to-date`,
      `Is tracking at ${displayPercent(row.annualGoalPercent, 1)} of annual goal`,
      `Has generated ${money(row.annualRevenue)} in FYTD revenue`,
      `Has a Rolling 90-Day Closing Rate of ${displayPercent(row.rolling90ClosingRate, 1)}`,
    ];

    backupReasons.forEach((reason) => {
      if (reasons.length < 8 && !reasons.includes(reason)) {
        reasons.push(reason);
      }
    });
    return reasons.slice(0, 8);
  };
  const weeklyTopDawgReasons = getWeeklyTopDawgReasons(weeklyTopDawg);

  const scheduledForBuild = Number(invoicePipeline.scheduledForBuild || 0);
  const needsToBeInvoiced = Number(invoicePipeline.needsToBeInvoiced || 0);
  const balanceDue = Number(invoicePipeline.balanceDue || 0);
  const totalPipeline = Number(invoicePipeline.totalPipeline || 0);

const revenueStatus = getStatus({
  good:
    Number(totals.totalGoalPercent || 0) >= 1 ||
    totals.teamVsLY.className === "positive",
  warning:
    Number(totals.totalGoalPercent || 0) >= 0.9 ||
    totals.teamVsLY.rawDifference >= -0.05,
});

const monthlyLeadingStatus = getStatus({
  good: monthlyAtGoalCount >= 4,
  warning: monthlyAtGoalCount >= 2,
});

const rolling90LeadingStatus = getStatus({
  good: rolling90ClosingAverage >= 0.25,
  warning: rolling90ClosingAverage >= 0.22,
});

const annualLeadingStatus = getStatus({
  good:
    annualClosingAverage >= 0.25 &&
    totals.teamAverageContract >= 12000 &&
    totals.totalContracts >= 300,
  warning:
    annualClosingAverage >= 0.22 &&
    totals.teamAverageContract >= 10000 &&
    totals.totalContracts >= 250,
});
const closingStatus = getStatus({
  good: rolling90ClosingAverage >= 0.25,
  warning: rolling90ClosingAverage >= 0.22,
});

  const cashFlowStatus = getStatus({
    good: totalPipeline >= 1000000,
    warning: totalPipeline >= 500000,
  });

  const collectionsStatus = getStatus({
    good: balanceDue <= 150000,
    warning: balanceDue <= 300000,
  });

  const healthCards = [
{
  label: "Goal Gap",
  value: money(Math.max(Number(totals.totalGoal || 0) - Number(totals.totalRevenue || 0), 0)),
  note: `${displayPercent(Math.max(1 - Number(totals.totalGoalPercent || 0), 0), 1)} left to goal`,
  status: totals.totalGoalPercent >= 1 ? "good" : totals.totalGoalPercent >= 0.75 ? "watch" : "attention",
},
{
  label: "Close Rate Gap",
value:
  rolling90ClosingAverage >= 0.25
    ? "On Goal"
    : `${((0.25 - rolling90ClosingAverage) * 100).toFixed(1)} pts below`,

note: "Compared to 25% Rolling 90 goal",

status:
  rolling90ClosingAverage >= 0.25
    ? "good"
    : rolling90ClosingAverage >= 0.20
    ? "watch"
    : "attention",
},
    {
      label: "Cash Flow",
      status: cashFlowStatus,
      value: money(totalPipeline),
      note: "Future cash pipeline",
    },
    {
      label: "Collections",
      status: collectionsStatus,
      value: money(balanceDue),
      note: "Balance due",
    },
  ];

  const weakestHealthArea =
    healthCards.find((card) => card.status === "attention") ||
    healthCards.find((card) => card.status === "watch");

  const coachingOpportunityText = rolling90ClosingLow
    ? `Meet with ${rolling90ClosingLow.name} to improve Rolling 90-Day Closing Rate. Current rate is ${displayPercent(
        rolling90ClosingLow.rolling90ClosingRate,
        1
      )}, which is ${formatPointGap(
        rolling90ClosingLow.rolling90ClosingRate,
        rolling90ClosingAverage
      )}.`
    : "Review Rolling 90-Day Closing Rate once metrics are loaded.";

  const secondaryRolling90ClosingLow = activeRows
    .slice()
    .filter(
      (row) =>
        row.name !== rolling90ClosingLow?.name &&
        Number(row.rolling90ClosingRate || 0) > 0
    )
    .sort(
      (a, b) =>
        Number(a.rolling90ClosingRate || 0) -
        Number(b.rolling90ClosingRate || 0)
    )[0];

  const annualGoalOpportunityText = annualGoalLow
    ? `Review ${annualGoalLow.name}'s annual production plan. ${money(
        Math.max(
          Number(annualGoalLow.goal || 0) -
            Number(annualGoalLow.annualRevenue || 0),
          0
        )
      )} remains to reach annual goal.`
    : "Review annual goal pace once revenue data is loaded.";

  const secondaryCoachingOpportunityText = secondaryRolling90ClosingLow
    ? `Work with ${secondaryRolling90ClosingLow.name} on improving Rolling 90-Day Closing Rate. Current rate is ${displayPercent(
        secondaryRolling90ClosingLow.rolling90ClosingRate,
        1
      )}, which is ${formatPointGap(
        secondaryRolling90ClosingLow.rolling90ClosingRate,
        rolling90ClosingAverage
      )}.`
    : "Review the next Rolling 90-Day Closing Rate coaching opportunity once data is loaded.";

  let recognitionOpportunityText = annualRevenueLeader
    ? `Recognize ${annualRevenueLeader.name} for leading FYTD production at ${money(
        annualRevenueLeader.annualRevenue
      )}. Reinforce the habits driving that success.`
    : "Recognize the current revenue leader once rankings are available.";

  if (quarterlyGoalLeader && Number(quarterlyGoalLeader.quarterlyGoalPercent || 0) >= 1) {
    recognitionOpportunityText = `Recognize ${quarterlyGoalLeader.name} for exceeding quarterly goal pace at ${displayPercent(
      quarterlyGoalLeader.quarterlyGoalPercent,
      1
    )}. Reinforce the behaviors creating that momentum.`;
  } else if (rolling90ClosingLeader && rolling90ClosingLeader.name !== rolling90ClosingLow?.name) {
    recognitionOpportunityText = `Recognize ${rolling90ClosingLeader.name} for leading Rolling 90-Day Closing Rate at ${displayPercent(
      rolling90ClosingLeader.rolling90ClosingRate,
      1
    )}. Reinforce the appointment and follow-up habits behind that performance.`;
  }

  const weeklyTopDawgText = weeklyTopDawg
    ? `Weekly Top Dawg: ${weeklyTopDawg.name}. Why: ${weeklyTopDawgReasons.join("; ")}.`
    : "Weekly Top Dawg will appear once enough performance data is available.";

  const primaryLeadershipFocus = rolling90ClosingLow
    ? `${rolling90ClosingLow.name}'s Rolling 90-Day Closing Rate is the biggest coaching opportunity at ${displayPercent(
        rolling90ClosingLow.rolling90ClosingRate,
        1
      )} (${formatPointGap(
        rolling90ClosingLow.rolling90ClosingRate,
        rolling90ClosingAverage
      )}).`
    : "Rolling 90-Day Closing Rate coaching data is still loading.";

  const prioritySentence = weeklyTopDawg
    ? `Today's leadership focus: coach where performance is lagging and recognize ${weeklyTopDawg.name} as Weekly Top Dawg.`
    : "Today's leadership focus: coach where performance is lagging and recognize strong execution.";

  const executiveBrief = [
    `${greeting} Revenue remains ${statusLabel[revenueStatus]} at ${money(
      totals.totalRevenue
    )} (${totals.teamVsLY.label} versus last fiscal year), while cash flow shows ${money(
      totalPipeline
    )} in visible future pipeline.`,
    annualRevenueLeader
      ? `${annualRevenueLeader.name} continues to lead FYTD production at ${money(
          annualRevenueLeader.annualRevenue
        )}.`
      : "FYTD production leader is not available yet.",
    primaryLeadershipFocus,
    prioritySentence,
  ].join(" ");

  const priorities = [
    {
      label: "Coaching Opportunity",
      text: coachingOpportunityText,
    },
    {
      label: "Coaching Opportunity",
      text: annualGoalOpportunityText,
    },
    {
      label: "Coaching Opportunity",
      text: secondaryCoachingOpportunityText,
    },
    {
      label: "Recognition Opportunity",
      text: recognitionOpportunityText,
    },
  ];

return {
  brief: executiveBrief,
  rolling90ClosingAverage,
  annualLeadingStatus,
  monthlyAtGoalCount,
  healthCards,
  priorities,
  weeklyTopDawg: weeklyTopDawg
      ? {
          name: weeklyTopDawg.name,
          image: weeklyTopDawg.image,
          score: weeklyTopDawg.weeklyTopDawgScore,
          reasons: weeklyTopDawgReasons,
        }
      : null,
    sections: [
      {
        title: "Monthly Production",
        eyebrow: "Current Month",
status: monthlyLeadingStatus,
insight: `Monthly Production is currently highlighted in ${
  monthlyAtGoalCount >= 4
    ? "green (good to go)"
    : monthlyAtGoalCount >= 2
    ? "yellow (keep an eye on it)"
    : "red (needs attention)"
} because ${monthlyAtGoalCount} of ${activeRows.length} salespeople are currently on monthly goal pace. ${
  monthlyAtGoalCount >= 4
    ? "The team has enough salespeople producing at goal pace to maintain healthy monthly performance."
    : monthlyAtGoalCount >= 2
    ? "Increasing the number of PM's on pace to 4 would move this section to green (good to go)."
    : "Increasing the number of PM's on pace to at least 2 would move this section to yellow (keep an eye on it). Reaching 4 PM's on pace would move it to green (good to go)."
}`,
        positive: [
          monthlyRevenueLeader
            ? `${monthlyRevenueLeader.name} leads monthly revenue at ${money(monthlyRevenueLeader.monthlyRevenue)}.`
            : "Monthly revenue leader is not available yet.",
monthlyContractsLeaders.length === 1
  ? `${monthlyContractsLeaders[0].name} has the most contracts sold this month at ${monthlyContractsLeaders[0].monthlyContracts}.`
  : monthlyContractsLeaders.length > 1
    ? `${formatLeaderNames(monthlyContractsLeaders)} are tied for the most contracts sold this month with ${monthlyContractsLeaders[0].monthlyContracts} each.`
    : "Monthly contract count is not available yet.",
          monthlyAverageContractLeader
            ? `${monthlyAverageContractLeader.name} has the highest monthly average contract at ${money(monthlyAverageContractLeader.monthlyAverageContract)}.`
            : "Monthly average contract leader is not available yet.",
          monthlyGoalLeader
            ? `${monthlyGoalLeader.name} is strongest against monthly goal pace at ${displayPercent(monthlyGoalLeader.monthlyGoalPercent, 1)}.`
            : "Monthly goal pace leader is not available yet.",
        ],
        coaching: [
          monthlyRevenueLow
            ? `${monthlyRevenueLow.name} has the lowest monthly production at ${money(monthlyRevenueLow.monthlyRevenue)}. Recommended action: review current appointment volume and open estimates.`
            : "Monthly revenue coaching data is not available yet.",
          monthlyGoalLow
            ? `${monthlyGoalLow.name} is lowest versus monthly goal pace at ${displayPercent(monthlyGoalLow.monthlyGoalPercent, 1)} (${money(Math.max(Number(monthlyGoalLow.monthlyGoal || 0) - Number(monthlyGoalLow.monthlyRevenue || 0), 0))} behind monthly goal). Recommended action: review this month's pipeline.`
            : "Monthly goal coaching data is not available yet.",
          monthlyAverageContractLow
            ? `${monthlyAverageContractLow.name} has the lowest monthly average contract at ${money(monthlyAverageContractLow.monthlyAverageContract)}. Recommended action: review upsell opportunities and job mix.`
            : "Monthly average-contract coaching data is not available yet.",
        ],
      },
      {
        title: "Rolling 90-Day Performance",
        eyebrow: "Rolling 90 Days",
status: rolling90LeadingStatus,
            insight: `Rolling 90-Day Performance is currently highlighted in ${
  rolling90ClosingAverage >= 0.25
    ? "green (good to go)"
    : rolling90ClosingAverage >= 0.22
    ? "yellow (keep an eye on it)"
    : "red (needs attention)"
} because the team Rolling 90-Day Closing Rate is ${displayPercent(
  rolling90ClosingAverage,
  1
)}. ${
  rolling90ClosingAverage >= 0.25
    ? "The team is meeting or exceeding the 25% closing-rate goal."
    : rolling90ClosingAverage >= 0.22
    ? "Improving the team closing rate to 25% would move this section to green (good to go)."
    : "Improving the team closing rate to at least 22% would move this section to yellow (keep an eye on it). Reaching 25% would move it to green (good to go)."
}`,
        positive: [
          rolling90ClosingLeader
            ? `${rolling90ClosingLeader.name} leads rolling 90-day closing rate at ${displayPercent(rolling90ClosingLeader.rolling90ClosingRate, 1)}.`
            : "Rolling 90-day closing-rate leader is not available yet.",
          rolling90RevenueLeader
            ? `${rolling90RevenueLeader.name} leads rolling 90-day revenue at ${money(rolling90RevenueLeader.rolling90Revenue)}.`
            : "Rolling 90-day revenue leader is not available yet.",
          quarterlyRevenueLeader
            ? `${quarterlyRevenueLeader.name} leads quarter-to-date revenue at ${money(quarterlyRevenueLeader.quarterMTDRevenue)}.`
            : "Quarter-to-date revenue leader is not available yet.",
(() => {
  const quarterlyGoalLeaders = activeRows.filter(
    (row) => Number(row.quarterlyGoalPercent || 0) >= 1
  );

  if (!quarterlyGoalLeaders.length) {
    return "No salesperson is currently at or above quarterly goal pace.";
  }

  const names = formatLeaderNames(quarterlyGoalLeaders);

  return quarterlyGoalLeaders.length === 1
    ? `${names} is at or above quarterly goal pace.`
    : `${names} are all at or above quarterly goal pace.`;
})(),
        ],
        coaching: [
          rolling90ClosingLow
            ? `${rolling90ClosingLow.name} has the lowest rolling 90-day closing rate at ${displayPercent(rolling90ClosingLow.rolling90ClosingRate, 1)} (${formatPointGap(rolling90ClosingLow.rolling90ClosingRate, rolling90ClosingAverage)}). Recommended action: review the last five unsold appointments.`
            : "Rolling 90-day closing-rate coaching data is not available yet.",
          quarterlyGoalLow
            ? `${quarterlyGoalLow.name} is lowest versus quarterly goal pace at ${displayPercent(quarterlyGoalLow.quarterlyGoalPercent, 1)} (${money(Math.max(Number(quarterlyGoalLow.quarterlyGoal || 0) - Number(quarterlyGoalLow.quarterlyRevenue || 0), 0))} behind quarterly goal). Recommended action: review quarter pipeline and close dates.`
            : "Quarterly goal coaching data is not available yet.",
          rolling90RevenueGap?.row
            ? `${rolling90RevenueGap.row.name} has the largest rolling 90-day revenue gap versus the team average at ${money(rolling90RevenueGap.gap)}. Recommended action: review lead volume, appointment quality, and average contract.`
            : quarterlyRevenueGap?.row
            ? `${quarterlyRevenueGap.row.name} has the largest quarter-to-date revenue gap versus the team average at ${money(quarterlyRevenueGap.gap)}. Recommended action: review lead volume, appointment quality, and average contract.`
            : "Revenue-gap coaching data is not available yet.",
        ],
      },
      {
        title: "Annual Performance",
        eyebrow: "Fiscal YTD",
        status: annualLeadingStatus,
insight: `Annual Performance is currently highlighted in ${
  annualLeadingStatus === "good"
    ? "green (good to go)"
    : annualLeadingStatus === "watch"
    ? "yellow (keep an eye on it)"
    : "red (needs attention)"
} because annual success is being measured by the leading indicators that drive revenue: closing rate, average contract value, and contract volume. The team is currently at ${displayPercent(
  annualClosingAverage,
  1
)} FYTD closing rate, ${money(
  totals.teamAverageContract
)} average contract value, and ${totals.totalContracts} contracts. ${
  annualLeadingStatus === "good"
    ? "These inputs are strong enough to support annual goal achievement."
    : annualLeadingStatus === "watch"
    ? "Improving FYTD closing rate to 25%, average contract value to $12,000, or contract volume above 300 would move this section to green (good to go)."
    : "Improving FYTD closing rate to at least 22%, average contract value to $10,000, and contract volume to 250 would move this section to yellow (keep an eye on it)."
}`,
        positive: [
          `Team revenue is ${totals.teamVsLY.label} versus the same fiscal period last year.`,
          annualRevenueLeader
            ? `${annualRevenueLeader.name} leads FYTD production at ${money(annualRevenueLeader.annualRevenue)}.`
            : "FYTD revenue leader is not available yet.",
          annualClosingLeader
            ? `${annualClosingLeader.name} leads FYTD closing rate at ${displayPercent(annualClosingLeader.closingRate, 1)}.`
            : "FYTD closing-rate leader is not available yet.",
          annualGoalLeader
            ? `${annualGoalLeader.name} is strongest against annual goal pace at ${displayPercent(annualGoalLeader.annualGoalPercent, 1)}.`
            : "Annual goal pace leader is not available yet.",
aboveTeamRevenue.length
  ? `${formatLeaderNames(aboveTeamRevenue)} ${
      aboveTeamRevenue.length === 1 ? "is" : "are"
    } above the team revenue average.`
  : "No salesperson is currently above the team revenue average.",
        ],
        coaching: [
          annualClosingLow
            ? `${annualClosingLow.name} has the lowest FYTD closing rate at ${displayPercent(annualClosingLow.closingRate, 1)} (${formatPointGap(annualClosingLow.closingRate, annualClosingAverage)}). Recommended action: review close-rate trend and unsold appointment notes.`
            : "FYTD closing-rate coaching data is not available yet.",
          annualGoalLow
            ? `${annualGoalLow.name} has the lowest annual goal progress at ${displayPercent(annualGoalLow.annualGoalPercent, 1)} (${money(Math.max(Number(annualGoalLow.goal || 0) - Number(annualGoalLow.annualRevenue || 0), 0))} remaining to annual goal). Recommended action: rebuild the goal recovery plan.`
            : "Annual goal coaching data is not available yet.",
          annualRevenueLow
            ? `${annualRevenueLow.name} is lowest versus team revenue average at ${annualRevenueLow.revenueVsTeam.label}. Recommended action: review production pace and average contract strategy.`
            : "Revenue versus team average data is not available yet.",
        ],
      },
      {
        title: "Business Health",
        eyebrow: "Cash Pipeline",
        status: balanceDue > 300000 ? "attention" : balanceDue > 150000 ? "watch" : "good",
        positive: [
          `${money(totalPipeline)} total future cash pipeline is currently visible.`,
          `${money(scheduledForBuild)} is already sold and scheduled for production.`,
          `${money(needsToBeInvoiced)} is ready to be invoiced.`,
          `${money(balanceDue)} is currently invoiced and owed.`,
        ],
        coaching: [
          needsToBeInvoiced > 0
            ? `${money(needsToBeInvoiced)} should be watched closely because it can convert to invoiceable cash quickly. Recommended action: confirm what can be invoiced this week.`
            : "No immediate invoice-ready cash is showing right now.",
          balanceDue > 0
            ? `${money(balanceDue)} in balance due should stay visible until collected. Recommended action: prioritize collection follow-up until cleared.`
            : "No balance-due exposure is showing right now.",
        ],
      },
    ],
  };
};  const SalesManagerDashboard = () => {
    const managerMetrics = getManagerMetrics();
    const { rows, totals, ranges } = managerMetrics;
    const invoicePipeline = getInvoicePipeline();
    const invoiceRepRows = getInvoicePipelineByRep();
    const executiveSummary = getExecutiveSummary(rows, totals, invoicePipeline);
    const weeklyTopDawg = executiveSummary.weeklyTopDawg;
    const activeSummarySection =
      executiveSummary.sections.find(
        (section) => section.title === activeSummaryTab
      ) || executiveSummary.sections[0];
    const todayLabel = new Date().toLocaleDateString("en-US", {
      weekday: "long",
      month: "long",
      day: "numeric",
    });

    return (
      <div className="sales-manager-page">
        <header className="manager-header executive-command-header">
          <img src="/logo.png" alt="Logo" className="manager-logo" />
          <div>
            <span>Exteriors Plus</span>
            <h1>Executive Command Center</h1>
            <p>{todayLabel}</p>
          </div>
        </header>

        <section className="manager-kpi-stack">
  <div className="manager-kpi-grid">
    <div className="manager-kpi-card primary">
      <span>FYTD Team Revenue</span>
      <strong>{money(totals.totalRevenue)}</strong>
      <small className={`manager-difference ${totals.teamVsLY.className}`}>
        {totals.teamVsLY.label} vs LY
      </small>
    </div>

    <div className="manager-kpi-card">
      <span>Cash Pipeline</span>
      <strong>{money(invoicePipeline.totalPipeline)}</strong>
      <small>Scheduled + invoice-ready + owed</small>
    </div>

    <div className="manager-kpi-card">
      <span>Money Owed</span>
      <strong>{money(invoicePipeline.balanceDue)}</strong>
      <small>Invoiced / owed</small>
    </div>

    <div className="manager-kpi-card">
      <span>Scheduled for Build</span>
      <strong>{money(invoicePipeline.scheduledForBuild)}</strong>
      <small>Future production money</small>
    </div>
  </div>

  <div className="manager-kpi-grid">
    <div className="manager-kpi-card">
      <span>Rolling 90-Day Closing Rate</span>
      <strong>{displayPercent(executiveSummary.rolling90ClosingAverage, 1)}</strong>
      <small
        className={`manager-difference ${
          executiveSummary.rolling90ClosingAverage >= 0.25 ? "positive" : "negative"
        }`}
      >
        {executiveSummary.rolling90ClosingAverage >= 0.25
          ? `${((executiveSummary.rolling90ClosingAverage - 0.25) * 100).toFixed(1)} pts above 25% goal`
          : `${((0.25 - executiveSummary.rolling90ClosingAverage) * 100).toFixed(1)} pts below 25% goal`}
      </small>
    </div>

    <div className="manager-kpi-card">
      <span>Annual Performance</span>
      <strong>
  {executiveSummary.annualLeadingStatus === "good"
    ? "Good"
    : executiveSummary.annualLeadingStatus === "watch"
    ? "Watch"
    : "Attention"}
</strong>
      <small>{displayPercent(totals.totalGoalPercent, 1)} annual progress</small>
    </div>

    <div className="manager-kpi-card">
      <span>Average Contract</span>
      <strong>{money(totals.teamAverageContract)}</strong>
      <small>{totals.totalContracts} contracts</small>
    </div>

    <div className="manager-kpi-card">
      <span>PMs on Goal Pace</span>
      <strong>{executiveSummary.monthlyAtGoalCount} / {rows.length}</strong>
      <small>Monthly goal pace</small>
    </div>
  </div>
</section>

        <section className="manager-panel manager-executive-summary-panel executive-command-panel">
          <div className="manager-panel-header">
            <div>
              <span>Executive Intelligence</span>
              <h2>Today's Executive Brief</h2>
            </div>
            <p>What the business needs from you today.</p>
          </div>

          <div className="executive-brief-card">
            <p>{executiveSummary.brief}</p>
          </div>

          <div className="executive-health-grid">
            {executiveSummary.healthCards.map((card) => (
              <div
                className={`executive-health-card ${card.status}`}
                key={card.label}
              >
                <span>{card.label}</span>
                <strong>{card.value}</strong>
                <small>{card.note}</small>
                <b>{card.status === "good" ? "Healthy" : card.status === "watch" ? "Watch" : "Attention"}</b>
              </div>
            ))}
          </div>

          <div className="executive-priority-card">
            <div className="manager-summary-section-header">
              <span>Leadership Focus</span>
              <h3>Leadership Opportunities</h3>
            </div>

            <div className="executive-priority-list">
              {executiveSummary.priorities.map((item, index) => (
                <div className="executive-priority-item" key={`priority-${index}`}>
                  <span>{index + 1}</span>
                  {typeof item === "string" ? (
                    <p>{item}</p>
                  ) : (
                    <div>
                      <strong>{item.label}</strong>
                      <p>{item.text}</p>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div className="manager-summary-tabs">
            {executiveSummary.sections.map((section) => (
              <button
                type="button"
                key={section.title}
                className={`${activeSummarySection.title === section.title ? "active" : ""} ${section.status}`}
                onClick={() => setActiveSummaryTab(section.title)}
              >
                <span className="summary-tab-status" />
                {section.title.replace(" Performance", "").replace("Fiscal YTD", "Annual")}
              </button>
            ))}
          </div>

          {activeSummarySection && (
            <div className={`manager-summary-section-card active ${activeSummarySection.status}`}>
              <div className="manager-summary-section-header">
                <span>{activeSummarySection.eyebrow}</span>
                <h3>{activeSummarySection.title}</h3>
              </div>
              {activeSummarySection.insight && (
  <div className={`manager-section-insight ${activeSummarySection.status}`}>
    <strong>Executive Interpretation</strong>
    <p>{activeSummarySection.insight}</p>
  </div>
)}

              <div className="manager-executive-summary-grid">
                <div className="manager-summary-column positive">
                  <h4>What's Going Well</h4>

                  {activeSummarySection.positive.map((item, index) => (
                    <div
                      className="manager-executive-summary-item positive"
                      key={`${activeSummarySection.title}-positive-${index}`}
                    >
                      <span>{String(index + 1).padStart(2, "0")}</span>
                      <p>{item}</p>
                    </div>
                  ))}
                </div>

                <div className="manager-summary-column coaching">
                  <h4>Leadership Opportunities</h4>

                  {activeSummarySection.coaching.map((item, index) => (
                    <div
                      className="manager-executive-summary-item coaching"
                      key={`${activeSummarySection.title}-coaching-${index}`}
                    >
                      <span>{String(index + 1).padStart(2, "0")}</span>
                      <p>{item}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
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

          <div className="manager-table-wrap manager-invoice-table-wrap">
            <table className="manager-table">
              <thead>
                <tr>
                  <th>Salesperson</th>
                  <th>Scheduled</th>
                  <th>Needs Invoicing</th>
                  <th>Balance Due</th>
                  <th>Total Pipeline</th>
                </tr>
              </thead>
              <tbody>
                {invoiceRepRows
                  .filter((row) => row.totalPipeline > 0)
                  .map((row) => (
                    <tr key={`${row.name}-invoice-pipeline`}>
                      <td>
                        <div className="manager-rep-cell">
                          <img src={row.image} alt={row.name} />
                          <strong>{row.name}</strong>
                        </div>
                      </td>
                      <td>{money(row.scheduledForBuild)}</td>
                      <td>{money(row.needsToBeInvoiced)}</td>
                      <td>{money(row.balanceDue)}</td>
                      <td>
                        <strong>{money(row.totalPipeline)}</strong>
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </section>

        <section className="manager-panel manager-salesperson-panel">
          <div className="manager-panel-header">
            <div>
              <span>Leaderboard</span>
              <h2>Salesperson Performance</h2>
            </div>
            <p>
              FYTD {formatShortDate(ranges.fiscalYTD.start)} – {formatShortDate(ranges.fiscalYTD.end)} compared to LY {formatShortDate(ranges.priorFiscalYTD.start)} – {formatShortDate(ranges.priorFiscalYTD.end)}
            </p>
          </div>

          {weeklyTopDawg && (
            <div className="employee-of-week-card">
              <div className="employee-of-week-person">
                <span>Weekly Top Dawg</span>
                <img src={weeklyTopDawg.image} alt={weeklyTopDawg.name} />
                <div>
                  <h3>{weeklyTopDawg.name}</h3>
                </div>
              </div>

<div className="employee-of-week-right">
  <div className="employee-of-week-heading">
    Performance Highlights
  </div>

  <div className="employee-of-week-reasons">
<ul className="employee-of-week-reason-grid">
  {weeklyTopDawg.reasons.map((reason, index) => (
    <li key={`weekly-top-dawg-reason-${index}`}>{reason}</li>
  ))}
</ul>
  </div>
</div>
            </div>
          )}

          <div className="manager-performance-list v2">
            {rows.map((row) => {
              const status = getRepStatus(row);

              return (
                <article
                  className={`manager-performance-report ${status.className}`}
                  key={`${row.name}-performance`}
                >
                  <div className="manager-performance-topline">
                    <div className="manager-rank-name-block">
                      <div className="manager-performance-rank">#{row.rank}</div>
<div className="manager-performance-person">
  <div>
    <h3>{row.name}</h3>
    <div className={`manager-status-pill ${status.className}`}>
      {status.label}
    </div>
  </div>

  <img src={row.image} alt={row.name} />
</div>
                    </div>

                    <div className="manager-metric-strip v21">
                      <div className="manager-metric-row revenue-row">
                        <MetricMiniBlock label="FYTD Revenue" value={money(row.annualRevenue)} />
                        <MetricMiniBlock
                          label="vs LY Revenue"
                          value={row.revenueVsLY.label}
                          difference={row.revenueVsLY}
                          subValue={`${money(row.lyRevenue)} LY`}
                        />
                        <MetricMiniBlock
                          label="vs Team Revenue"
                          value={row.revenueVsTeam.label}
                          difference={row.revenueVsTeam}
                          subValue={`${money(totals.teamAverageRevenue)} avg`}
                        />
                        <MetricMiniBlock label="Rolling 90 Revenue" value={money(row.rolling90Revenue)} />
                        <MetricMiniBlock label="Quarter MTD Revenue" value={money(row.quarterMTDRevenue)} />
                      </div>

                      <div className="manager-metric-row-divider" />

                      <div className="manager-metric-row close-row">
                        <MetricMiniBlock
                          label="FYTD Close Rate"
                          value={displayPercent(row.closingRate, 1)}
                          difference={row.closingVsTeam}
                        />
                        <MetricMiniBlock
                          label="vs LY Close Rate"
                          value={row.closingVsLY.label}
                          difference={row.closingVsLY}
                          subValue={`${displayPercent(row.lyClosingRate, 1)} LY`}
                        />
                        <MetricMiniBlock
                          label="vs Team Close Rate"
                          value={row.closingVsTeam.label}
                          difference={row.closingVsTeam}
                          subValue={`${displayPercent(totals.teamClosingRate, 1)} avg`}
                        />
                        <MetricMiniBlock label="Rolling 90 Close Rate" value={displayPercent(row.rolling90ClosingRate, 1)} />
                        <MetricMiniBlock label="Quarter MTD Close Rate" value={displayPercent(row.quarterMTDClosingRate, 1)} />
                      </div>
                    </div>
                  </div>

                  <div className="manager-performance-bottomline">
                    <div className="manager-synopsis-card">
                      <span>Manager Insight</span>
                      <p>{getManagerSynopsis(row)}</p>
                    </div>

                    <div className="manager-goal-thermometers v2">
                      <GoalThermometer
                        label="Monthly"
                        amount={row.monthlyRevenue}
                        percentValue={row.monthlyGoalPercent}
                      />
                      <GoalThermometer
                        label="Quarterly"
                        amount={row.quarterlyRevenue}
                        percentValue={row.quarterlyGoalPercent}
                      />
                      <GoalThermometer
                        label="Annual"
                        amount={row.annualRevenue}
                        percentValue={row.annualGoalPercent}
                      />
                    </div>
                  </div>
                </article>
              );
            })}
          </div>
        </section>
      </div>
    );
  };
if (!managerAuthorized) {
  return (
    <div className="page manager-shell">
      <div className="manager-login-card">
        <img src="/logo.png" alt="Logo" className="manager-login-logo" />

        <h1>Sales Manager Portal</h1>
        <p>Enter password to continue.</p>

        <form
          onSubmit={(event) => {
            event.preventDefault();

            if (managerPassword === MANAGER_PASSWORD) {
              sessionStorage.setItem("managerAuthorized", "true");
              setManagerAuthorized(true);
              setManagerPassword("");
              setManagerLoginError("");
            } else {
              setManagerLoginError("Incorrect password. Please try again.");
            }
          }}
        >
          <input
            type="password"
            value={managerPassword}
            onChange={(event) => setManagerPassword(event.target.value)}
            placeholder="Password"
          />

          <button type="submit">Enter</button>
        </form>

        {managerLoginError && (
          <div className="manager-login-error">{managerLoginError}</div>
        )}
      </div>
    </div>
  );
}
  return (
    <div className="page manager-shell">
      {dataStatus && <div className="manager-status">{dataStatus}</div>}
      <SalesManagerDashboard />
    </div>
  );
}
