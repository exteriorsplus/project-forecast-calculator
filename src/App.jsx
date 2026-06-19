import { useCallback, useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import confetti from "canvas-confetti";
import JSConfetti from "js-confetti";
import "./App.css";

const APP_PASSWORD = "Lakeview2910";

const UNKNOWN_KEY = "Unknown-Work Type";
const UNKNOWN_RPP = 15191.27;
const UNKNOWN_MARGIN = 0.277;
const COMPANY_EXPENSE_RATE = 0.1;

const MARKETING_COMPANY_MARGIN = 0.21;
const MARKETING_TRUE_PROJECT_MARGIN =
  MARKETING_COMPANY_MARGIN + COMPANY_EXPENSE_RATE;

const MARKETING_SPEND_SPLIT = {
  googleLocalService: 0.7284,
  googleAdsSeo: 0.1377,
  thumbtack: 0.1339,
};

const marketingChannels = [
  {
    key: "googleLocalService",
    label: "Google Local Service",
    defaultSpend: 25985.91,
    leads: 165,
    appointments: 31,
    averageRevenuePerAppointment: 2232.24,
  },
  {
    key: "googleAdsSeo",
    label: "Google Ads / SEO",
    defaultSpend: 7976.88,
    leads: null,
    appointments: 121,
    averageRevenuePerAppointment: 3574.19,
  },
  {
    key: "thumbtack",
    label: "Thumbtack",
    defaultSpend: 8833.12,
    leads: 78,
    appointments: 41,
    averageRevenuePerAppointment: 1139.22,
  },
];

const marketingScenarios = [
  { key: "conservative", label: "Conservative", factor: 0.85 },
  { key: "expected", label: "Expected", factor: 1 },
  { key: "aggressive", label: "Aggressive", factor: 1.15 },
];
const historicalMarketingData = [
  {
    channel: "Google Local Service",
    rows: [
      { month: "Oct 2025", spend: 21707.27, revenue: 81856.9 },
      { month: "Nov 2025", spend: 10728.89, revenue: 15033.3 },
      { month: "Dec 2025", spend: 8686.2, revenue: 44918.5 },
      { month: "Jan 2026", spend: 12634.11, revenue: 49926.44 },
      { month: "Feb 2026", spend: 16854.55, revenue: 30022.96 },
      { month: "Mar 2026", spend: 39419.01, revenue: 254652.52 },
      { month: "Apr 2026", spend: 21500, revenue: 49993.06 },
    ],
  },
  {
    channel: "Google Ads / SEO",
    rows: [
      { month: "Oct 2025", spend: 4839.65, revenue: 24395.18 },
      { month: "Nov 2025", spend: 4390.23, revenue: 62146.38 },
      { month: "Dec 2025", spend: 317.17, revenue: 99771.9 },
      { month: "Jan 2026", spend: 892.94, revenue: 152525.13 },
      { month: "Feb 2026", spend: 1043.44, revenue: 32266.07 },
      { month: "Mar 2026", spend: 12438.75, revenue: 210804.96 },
      { month: "Apr 2026", spend: 7976.88, revenue: 214601.44 },
    ],
  },
  {
    channel: "Thumbtack",
    rows: [
      { month: "Oct 2025", spend: 2033.45, revenue: 31945.92 },
      { month: "Nov 2025", spend: 2685.1, revenue: 23946.03 },
      { month: "Dec 2025", spend: 1072.79, revenue: 0 },
      { month: "Jan 2026", spend: 1182.32, revenue: 0 },
      { month: "Feb 2026", spend: 2334.4, revenue: 3100 },
      { month: "Mar 2026", spend: 11574.3, revenue: 80014.51 },
      { month: "Apr 2026", spend: 8833.12, revenue: 54165.59 },
    ],
  },
];

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

const PM_COMPANY_GOAL = Object.values(PM_GOALS).reduce(
  (sum, goal) => sum + goal,
  0
);
const PM_COMMISSION_RATE = 0.25;

const getPMCommissionRate = (pmName) => {
  if (pmName === "George Anim") return 0.15;
  return 0.25;
};

const getOverheadRate = (tradeType) => {
  if (tradeType === "James Hardie Siding") return 0.15;
  if (tradeType === "Metal Roofing") return 0.15;
  return 0.10;
};

const projectManagers = [
  {
    name: "Jamie Jenkins",
    slug: "jamiejenkins",
    image: "/pm/jamiejenkins.jpg",
    password: "Jamie123",
    activeGoal: true,
  },
  {
    name: "Megan Rice",
    slug: "meganrice",
    image: "/pm/meganrice.jpg",
    password: "Megan123",
    activeGoal: true,
  },
  {
    name: "Dani Cole",
    slug: "danicole",
    image: "/pm/danicole.jpg",
    password: "Dani123",
    activeGoal: true,
  },
  {
    name: "John Fincher",
    slug: "johnfincher",
    image: "/pm/johnfincher.jpg",
    password: "John123",
    activeGoal: true,
  },
  {
    name: "Andrew Painter",
    slug: "andrewpainter",
    image: "/pm/andrewpainter.jpg",
    password: "Andrew123",
    activeGoal: true,
  },
  {
    name: "George Anim",
    slug: "georgeanim",
    image: "/pm/georgeanim.jpg",
    password: "George123",
    activeGoal: true,
  },
  {
    name: "William Dye",
    slug: "williamdye",
    image: "/pm/williamdye.jpg",
    password: "William123",
    activeGoal: true,
  },
{
  name: "Mike Harr",
  slug: "mikeharr",
  image: "/pm/mikeharr.jpg",
  password: "Mike123",
  activeGoal: true,
},
];

const getCurrentProjectManager = () => {
  if (typeof window === "undefined") return null;

  const hostname = window.location.hostname.toLowerCase();

  return (
    projectManagers.find((pm) => hostname.startsWith(`${pm.slug}.`)) || null
  );
};

const categories = [
  {
    title: "Roofing",
    className: "roofing",
    items: [
      { label: "Insurance", rpp: 19082.88, margin: 0.258049 },
      { label: "Repair", rpp: 1142.38, margin: 0.364825 },
      { label: "Retail", rpp: 12984.44, margin: 0.212359 },
    ],
  },
  {
    title: "Roofing & Gutters",
    className: "roofing-gutters",
    items: [
      { label: "Insurance", rpp: 20999.14, margin: 0.255887 },
      { label: "Repair", rpp: 4198.92, margin: 0.350644 },
      { label: "Retail", rpp: 14127.34, margin: 0.210625 },
    ],
  },
  {
    title: "Roofing & Siding",
    className: "combo",
    items: [
      { label: "Insurance", rpp: 41468.5, margin: 0.267046 },
      { label: "Repair", rpp: 2097.63, margin: 0.289835 },
      { label: "Retail", rpp: 32113.64, margin: 0.188107 },
    ],
  },
  {
    title: "Siding",
    className: "siding",
    items: [
      { label: "Insurance", rpp: 21568.3, margin: 0.272371 },
      { label: "Repair", rpp: 1457.44, margin: 0.30295 },
      { label: "Retail", rpp: 17962.46, margin: 0.211199 },
    ],
  },
  {
    title: "James Hardie Siding",
    className: "james-hardie",
    items: [
      { label: "Insurance", rpp: 32561.58, margin: 0.146764 },
      { label: "Repair", rpp: 848.52, margin: 0.351721 },
      { label: "Retail", rpp: 41849.3, margin: 0.154653 },
    ],
  },
  {
    title: "Metal Roofing",
    className: "metal-roofing",
    items: [
      { label: "Insurance", rpp: 30710.23, margin: 0.308307 },
      { label: "Repair", rpp: 4448.43, margin: 0.202305 },
      { label: "Retail", rpp: 20678.49, margin: 0.172779 },
    ],
  },
  {
    title: "Windows",
    className: "windows",
    items: [
      { label: "Insurance", rpp: 27727.17, margin: 0.154518 },
      { label: "Repair", rpp: 559.78, margin: 0.231567 },
      { label: "Retail", rpp: 11016.29, margin: 0.191793 },
    ],
  },
  {
    title: "Gutters",
    className: "gutters",
    items: [
      { label: "Insurance", rpp: 4557.48, margin: 0.400291 },
      { label: "Repair", rpp: 2732.52, margin: 0.398716 },
      { label: "Retail", rpp: 4077.33, margin: 0.300995 },
    ],
  },
  {
    title: "Doors",
    className: "doors",
    items: [
      { label: "Retail", rpp: 4996.31, margin: 0.19385 },
      { label: "Service", rpp: 200.0, margin: 0.85 },
    ],
  },
];

const getMarginForTradeType = (tradeType) => {
  const category = categories.find(
    (c) => c.title === tradeType
  );

  if (!category) return 0;

  const margins = category.items
    .map((item) => item.margin)
    .filter((margin) => margin > 0);

  if (!margins.length) return 0;

  return (
    margins.reduce((sum, margin) => sum + margin, 0) /
    margins.length
  );
};

function normalizeMarginTradeType(value) {
  return normalizeTrade(value);
}

function normalizeMarginWorkType(value) {
  return normalizeWorkType(value);
}

function getMarginRowPayment(row) {
  return parseMoney(
    row["Payments Received Total"] ??
      row["SUM of Paymen"] ??
      row["SUM of Payments Received Total"] ??
      row["Contract Amount"] ??
      row["Payment"]
  );
}

function getMarginRowProfit(row) {
  return parseMoney(
    row["Profit"] ??
      row["SUM of Profit"] ??
      row["Gross Profit"]
  );
}

function getMarginRowPercent(row) {
  const raw =
    row["Profit %"] ??
    row["Margin per job"] ??
    row["Margin"] ??
    row["Profit Percent"];

  if (raw === null || raw === undefined || raw === "") return null;

  if (typeof raw === "number") {
    return raw > 1 ? raw / 100 : raw;
  }

  const cleaned = String(raw).replace("%", "").trim();
  const number = Number(cleaned);

  if (Number.isNaN(number)) return null;

  return number > 1 ? number / 100 : number;
}


const WORK_TYPE_ORDER = ["Retail", "Insurance", "Repair", "Service"];
const COMMISSION_WORK_TYPE_OPTIONS = ["Retail", "Insurance", "Repair"];

const money = (value) =>
  Number(value || 0).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
  });

const formatMoneyInput = (value) => {
  const cleaned = String(value || "").replace(/[^\d.]/g, "");

  if (!cleaned) return "";

  const parts = cleaned.split(".");

  const wholePart = parts[0];
  const decimalPart = parts[1];

  const formattedWhole = Number(
    wholePart || 0
  ).toLocaleString("en-US");

  return decimalPart !== undefined
    ? `$${formattedWhole}.${decimalPart}`
    : `$${formattedWhole}`;
};
const cleanMoneyInput = (value) => {
  const cleaned = String(value || "").replace(/[^\d.]/g, "");
  const parts = cleaned.split(".");

  if (parts.length <= 1) return cleaned;

  return `${parts[0]}.${parts.slice(1).join("")}`;
};

const percent = (value) => `${(value * 100).toFixed(4)}%`;


const displayPercent = (value, decimals = 1) => {
  const number = Number(value || 0);

  return `${(number * 100).toFixed(decimals)}%`;
};

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

function getPreviousYearMonth(monthLabel) {
  if (!monthLabel || monthLabel === "YTD") return "";

  const parts = monthLabel.split(" ");
  const year = Number(parts[parts.length - 1]);
  const month = parts.slice(0, -1).join(" ");

  if (!year || !month) return "";

  return `${month} ${year - 1}`;
}

function getMonthSortValue(monthLabel) {
  if (monthLabel === "YTD") return 999999;

  const [month, year] = String(monthLabel).split(" ");
  const monthIndex = monthNames.indexOf(month);

  return Number(year || 0) * 100 + monthIndex;
}

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

function getRankSuffix(rank) {
  if (!rank) return "";

  const lastDigit = rank % 10;
  const lastTwoDigits = rank % 100;

  if (lastTwoDigits >= 11 && lastTwoDigits <= 13) return "th";
  if (lastDigit === 1) return "st";
  if (lastDigit === 2) return "nd";
  if (lastDigit === 3) return "rd";

  return "th";
}


function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseInputDate(value) {
  if (!value) return null;
  const [year, month, day] = String(value).split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

function getMonthDateRange() {
  const today = new Date();
  const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  return {
    start: formatDate(firstOfMonth),
    end: formatDate(today),
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

function getInclusiveMonthCount(startDate, endDate) {
  if (!startDate || !endDate) return 0;

  const start = dateOnly(startDate);
  const end = dateOnly(endDate);

  if (end < start) return 0;

  return (
    (end.getFullYear() - start.getFullYear()) * 12 +
    end.getMonth() -
    start.getMonth() +
    1
  );
}

function normalizeTrade(value) {
  const text = String(value || "").trim().toLowerCase();

    // EXCLUDED TRADE TYPES
if (
  text === "other" ||
  text === "other, skylights" ||
  text === "other, fascia" ||
  text === "wraps, fascia" ||
  text === "storm damage, fascia" ||
  text === "soffits, fascia" ||
  text === "repair" ||
  text === "wraps"
) {
  return "";
}

  if (text === "roofing & gutters") return "Roofing & Gutters";
  if (text === "roofing & siding") return "Roofing & Siding";
  if (text === "james hardie siding") return "James Hardie Siding";
  if (text === "metal roofing") return "Metal Roofing";
  if (text === "roofing") return "Roofing";
  if (text === "siding") return "Siding";
  if (text === "gutters") return "Gutters";
  if (text === "doors") return "Doors";
  if (text === "windows") return "Windows";

  if (text.includes("roofing") && text.includes("gutters")) return "Roofing & Gutters";
  if (text.includes("roofing") && text.includes("siding")) return "Roofing & Siding";
  if (text.includes("james hardie")) return "James Hardie Siding";
  if (text.includes("metal roofing")) return "Metal Roofing";
  if (text.includes("roofing")) return "Roofing";
  if (text.includes("siding")) return "Siding";
  if (text.includes("gutters")) return "Gutters";
  if (text.includes("doors")) return "Doors";
  if (text.includes("windows")) return "Windows";

  return "";
}

function normalizeWorkType(value) {
  const text = String(value || "").trim().toLowerCase();

  if (text.includes("insurance")) return "Insurance";
  if (text.includes("repair")) return "Repair";
  if (text.includes("retail")) return "Retail";
  if (text.includes("service")) return "Service";

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

function findProjectConfig(trade, workType) {
  for (const category of categories) {
    for (const item of category.items) {
      if (category.title === trade && item.label === workType) {
        return {
          key: `${category.title}-${item.label}`,
          rpp: item.rpp,
          margin: item.margin,
        };
      }
    }
  }

  return {
    key: UNKNOWN_KEY,
    rpp: UNKNOWN_RPP,
    margin: UNKNOWN_MARGIN,
  };
}

function buildClearedCounts(includeUnknown = false) {
  const cleared = {};

  categories.forEach((category) => {
    category.items.forEach((item) => {
      cleared[`${category.title}-${item.label}`] = 0;
    });
  });

  if (includeUnknown) cleared[UNKNOWN_KEY] = 0;

  return cleared;
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
      if (frameRef.current) cancelAnimationFrame(frameRef.current);
    };
  }, [value]);

  return <>{money(displayValue)}</>;
}

function ForecastRow({
  item,
  keyName,
  quantity,
  revenue,
  breakdown,
  type,
  onLeadQuantityChange,
}) {
  const [localValue, setLocalValue] = useState(quantity ?? "");

  useEffect(() => {
    setLocalValue(quantity ?? "");
  }, [quantity]);

  useEffect(() => {
    if (type !== "leads") return;

    if (String(quantity ?? "") === String(localValue ?? "")) {
      return;
    }

    const timeout = setTimeout(() => {
      onLeadQuantityChange(keyName, localValue);
    }, 600);

    return () => clearTimeout(timeout);
  }, [localValue, quantity, keyName, type, onLeadQuantityChange]);

  return (
    <div className="row">
      <div className="row-label">{item.label}</div>

      <div className="row-content">
        <div className="row-left">
          <div className="sub-label">
            Rev / Project
            <strong>{money(item.rpp)}</strong>
          </div>

          <div className="sub-label">
            True Margin
            <strong>{percent(breakdown.trueProjectMargin)}</strong>
          </div>

          <div className="sub-label">
            Company Margin
            <strong>{percent(item.margin)}</strong>
          </div>
        </div>

        <div className="row-middle">
          <input
            type="number"
            min="0"
            placeholder={type === "leads" ? "Leads" : "Projects"}
            value={localValue}
            readOnly={type === "projects"}
            onChange={(event) => {
              if (type === "projects") return;
              setLocalValue(event.target.value);
            }}
          />
        </div>

        <div className="row-right">
          <div className="metric">
            <span>Revenue</span>
            <strong>{money(revenue)}</strong>
          </div>

          <div className="metric true">
            <span>Project Profit</span>
            <strong>{money(breakdown.trueProjectProfit)}</strong>
          </div>

          <div className="metric company">
            <span>Company Profit</span>
            <strong>{money(breakdown.companyProfit)}</strong>
          </div>
        </div>
      </div>
    </div>
  );
}


function PMMetricCard({
  label,
  value,
  comparisonLabel,
  comparisonValue,
  difference,
  customMessage,
  messageTitle = "✨MAGIC MIKE MOMENT✨",
}) {
  return (
    <div className="pm-metric-card">
      <span>{label}</span>
      <strong>{value}</strong>

      {comparisonLabel && (
        <div className="pm-comparison-line">
          <small>{comparisonLabel}</small>
          <b>{comparisonValue}</b>
        </div>
      )}

      {difference && (
        <div className={`pm-difference ${difference.className}`}>
          {difference.label}
        </div>
      )}

      {customMessage && (
        <div className="mike-moment-mini">
          <img src="/pm/mikeharr.jpg" alt="Mike Harr" />

          <div className="mike-moment-mini-bubble">
            <div className="mike-moment-mini-title">
              {messageTitle}
            </div>

            <p>{customMessage}</p>
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [authorized, setAuthorized] = useState(
    () => sessionStorage.getItem("dashboardAuthorized") === "true"
  );
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const [screen, setScreen] = useState("home");
  const [leads, setLeads] = useState({});
  const [salesRows, setSalesRows] = useState([]);
  const [closeRate, setCloseRate] = useState(0.25);
  const [customCloseRate, setCustomCloseRate] = useState("");
  const [flash, setFlash] = useState(false);
  const [leadRows, setLeadRows] = useState([]);
  const [marginRows, setMarginRows] = useState([]);
  const [dataStatus, setDataStatus] = useState("Loading files...");

  const [pmRows, setPmRows] = useState([]);
  const [pmAuthorized, setPmAuthorized] = useState(() => {
    const initialPM = getCurrentProjectManager();
    if (!initialPM) return false;
    return sessionStorage.getItem(`${initialPM.slug}PMAuthorized`) === "true";
  });
  const [pmPassword, setPmPassword] = useState("");
  const [pmLoginError, setPmLoginError] = useState("");
 const [pmSaleAmount, setPmSaleAmount] = useState("0");
const [debouncedPmSaleAmount, setDebouncedPmSaleAmount] = useState("0");
const [pmJobCost, setPmJobCost] = useState("0");
const [debouncedPmJobCost, setDebouncedPmJobCost] = useState("0");
  const [selectedCommissionTrades, setSelectedCommissionTrades] = useState([]);
  const [selectedCommissionWorkTypes, setSelectedCommissionWorkTypes] = useState([]);
  const [selectedPMMonth, setSelectedPMMonth] = useState("");
  const [pmStartDate, setPmStartDate] = useState("");
const [pmEndDate, setPmEndDate] = useState("");
const [pmDateMode, setPmDateMode] = useState("month");


  const defaultTotalMarketingSpend = marketingChannels.reduce(
    (sum, channel) => sum + channel.defaultSpend,
    0
  );

const [totalMarketingSpend, setTotalMarketingSpend] = useState("0");

const [marketingSpendByChannel, setMarketingSpendByChannel] = useState(() =>
  marketingChannels.reduce((totals, channel) => {
    totals[channel.key] = "0";
    return totals;
  }, {})
);

  const initialDateRange = getMonthDateRange();

  const [startDate, setStartDate] = useState(initialDateRange.start);
  const [endDate, setEndDate] = useState(initialDateRange.end);

  const [projectStartDate, setProjectStartDate] = useState(initialDateRange.start);
  const [projectEndDate, setProjectEndDate] = useState(initialDateRange.end);

  const flashTimeoutRef = useRef(null);
  const commissionCardRef = useRef(null);
  const previousCommissionRef = useRef(0);
  const jsConfettiRef = useRef(null);
  useEffect(() => {
  jsConfettiRef.current = new JSConfetti();
}, []);

 const currentPM = getCurrentProjectManager();
const isPMPortal = Boolean(currentPM);

const renderTopDawgNames = (names) =>
  names.map((name, index) => (
    <span
      key={`${name}-${index}`}
      className={name === currentPM?.name ? "top-dawg-current-pm" : ""}
    >
      {index > 0 ? ", " : ""}
      {name}
    </span>
  ));

const rawRate =
  closeRate === "custom" ? Number(customCloseRate || 0) / 100 : closeRate;

  const activeCloseRate = Math.min(Math.max(rawRate, 0), 1);

  const applyHistoricalSpendSplit = () => {
    const total = Number(totalMarketingSpend || 0);

    setMarketingSpendByChannel({
      googleLocalService: String(
        (total * MARKETING_SPEND_SPLIT.googleLocalService).toFixed(2)
      ),
      googleAdsSeo: String(
        (total * MARKETING_SPEND_SPLIT.googleAdsSeo).toFixed(2)
      ),
      thumbtack: String(
        (total * MARKETING_SPEND_SPLIT.thumbtack).toFixed(2)
      ),
    });
  };

  const getMarketingForecast = () => {
    const channelForecasts = marketingChannels.map((channel) => {
      const spend = Number(marketingSpendByChannel[channel.key] || 0);

      const costPerLead =
        channel.leads && channel.leads > 0
          ? channel.defaultSpend / channel.leads
          : null;

      const appointmentsPerSpend =
        channel.defaultSpend > 0
          ? channel.appointments / channel.defaultSpend
          : 0;

      const expectedLeads = costPerLead ? spend / costPerLead : null;
      const expectedAppointments = spend * appointmentsPerSpend;

      const scenarios = marketingScenarios.map((scenario) => {
        const appointments = expectedAppointments * scenario.factor;
        const leads =
          expectedLeads === null ? null : expectedLeads * scenario.factor;

        const revenue =
          appointments *
          channel.averageRevenuePerAppointment *
          scenario.factor;

        return {
          ...scenario,
          leads,
          appointments,
          revenue,
          trueProjectProfit: revenue * MARKETING_TRUE_PROJECT_MARGIN,
          companyProfit: revenue * MARKETING_COMPANY_MARGIN,
          returnPerSpend: spend > 0 ? revenue / spend : 0,
        };
      });

      return {
        ...channel,
        spend,
        splitPercent: MARKETING_SPEND_SPLIT[channel.key],
        costPerLead,
        appointmentsPerSpend,
        scenarios,
      };
    });

    const totals = marketingScenarios.map((scenario) => {
      const matchingScenarios = channelForecasts.map((channel) =>
        channel.scenarios.find((item) => item.key === scenario.key)
      );

      const spend = channelForecasts.reduce(
        (sum, channel) => sum + channel.spend,
        0
      );

      const revenue = matchingScenarios.reduce(
        (sum, item) => sum + item.revenue,
        0
      );

      return {
        ...scenario,
        spend,
        leads: matchingScenarios.reduce(
          (sum, item) => sum + Number(item.leads || 0),
          0
        ),
        appointments: matchingScenarios.reduce(
          (sum, item) => sum + item.appointments,
          0
        ),
        revenue,
        trueProjectProfit: matchingScenarios.reduce(
          (sum, item) => sum + item.trueProjectProfit,
          0
        ),
        companyProfit: matchingScenarios.reduce(
          (sum, item) => sum + item.companyProfit,
          0
        ),
        returnPerSpend: spend > 0 ? revenue / spend : 0,
      };
    });

    const actualTotalSpend = channelForecasts.reduce(
      (sum, channel) => sum + channel.spend,
      0
    );

    return {
      channelForecasts,
      totals,
      actualTotalSpend,
    };
  };

  const handleLogin = (event) => {
    event.preventDefault();

    if (password === APP_PASSWORD) {
      sessionStorage.setItem("dashboardAuthorized", "true");
      setAuthorized(true);
      setLoginError("");
      setPassword("");
    } else {
      setLoginError("Incorrect password. Please try again.");
    }
  };

  const logout = () => {
    sessionStorage.removeItem("dashboardAuthorized");
    setAuthorized(false);
    setScreen("home");
    setLeadRows([]);
    setSalesRows([]);
    setMarginRows([]);
    setLeads({});
  };

  const handleLeadQuantityChange = useCallback((keyName, value) => {
    setLeads((currentLeads) => {
      const nextValue = value === "" ? "" : Number(value);

      if (currentLeads[keyName] === nextValue) {
        return currentLeads;
      }

      return {
        ...currentLeads,
        [keyName]: nextValue,
      };
    });
  }, []);

  const triggerFlash = () => {
    setFlash(true);

    if (flashTimeoutRef.current) clearTimeout(flashTimeoutRef.current);

    flashTimeoutRef.current = setTimeout(() => {
      setFlash(false);
    }, 200);
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

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
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

  const loadSalesFile = async () => {
    try {
      setDataStatus("Loading sales.xlsx...");

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
      setDataStatus(`${rows.length} rows loaded from sales.xlsx`);
    } catch (error) {
      setSalesRows([]);
      setDataStatus("No public/sales.xlsx file found yet.");
      console.error(error);
    }
  };


  const loadMarginFile = async () => {
    try {
      const response = await fetch(`/margin.csv?t=${Date.now()}`);

      if (!response.ok) {
        throw new Error("Could not find public/margin.csv");
      }

      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: "array",
      });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
      });

      setMarginRows(rows);
    } catch (error) {
      setMarginRows([]);
      console.error("No public/margin.csv file found yet.", error);
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

  const getPMBlockStart = (pmName) => {
    return pmRows.findIndex(
      (row) => String(row?.[0] || "").trim() === pmName
    );
  };

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

  // Backward compatibility: old code/data used "Closing Rate".
  // New monthly PM sheet row is "Monthly Closing Rate".
  if (requested === "closing rate" && rowMetric === "monthly closing rate") {
    return true;
  }

  return false;
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

  const getPMMonthOptions = (pmName = currentPM?.name || projectManagers[0].name) => {
    const startIndex = getPMBlockStart(pmName);

    if (startIndex < 0) return [];

    const headerRow = pmRows[startIndex] || [];

    return headerRow
      .slice(1)
      .map(formatPMMonth)
      .filter(Boolean)
      .filter((month) => month !== "YTD")
      .sort((a, b) => getMonthSortValue(b) - getMonthSortValue(a));
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
      pmRows.find((item) => String(item?.[0] || "").trim() === (currentPM?.name || projectManagers[0].name)) ||
      [];

    const columnIndex = headerSource.findIndex(
      (cell) => formatPMMonth(cell) === monthLabel
    );

    if (columnIndex < 0) return 0;

    return parseMetricValue(row[columnIndex]);
  };

  const getRankForMetric = (
    metric,
    monthLabel,
    pmName = currentPM?.name || projectManagers[0].name
  ) => {
    const ranked = projectManagers
      .filter((pm) => pm.activeGoal)
      .map((pm) => ({
        name: pm.name,
        value: getPMMetric(pm.name, metric, monthLabel),
      }))
      .filter((pm) => Number(pm.value || 0) > 0)
      .sort((a, b) => b.value - a.value);

    const index = ranked.findIndex((pm) => pm.name === pmName);

    return {
      rank: index >= 0 ? index + 1 : null,
      total: ranked.length,
    };
  };
    const getRevenueRankForRange = (
  startDate,
  endDate,
  pmName = currentPM?.name || projectManagers[0].name
) => {
  const ranked = projectManagers
    .filter((pm) => pm.activeGoal)
    .map((pm) => ({
      name: pm.name,
      value: getPMSalesDataForRange(pm.name, startDate, endDate).contractTotal,
    }))
    .filter((pm) => Number(pm.value || 0) > 0)
    .sort((a, b) => b.value - a.value);

  const index = ranked.findIndex((pm) => pm.name === pmName);

  return {
    rank: index >= 0 ? index + 1 : null,
    total: ranked.length,
  };
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

  const getPMCustomSalesData = (pmName) => {
    if (!pmStartDate || !pmEndDate) {
      return {
        contractTotal: 0,
        contracts: 0,
        averageContract: 0,
      };
    }

    return getPMSalesDataForRange(
      pmName,
      parseInputDate(pmStartDate),
      parseInputDate(pmEndDate)
    );
  };

const getPMReferralDataForRange = (pmName, startDate, endDate) => {
  if (!pmName || !startDate || !endDate) {
    return {
      referralTotal: 0,
      referralGoal: 0,
      referralDelta: 0,
      referralPercent: 0,
    };
  }

  const start = dateOnly(startDate);
  const end = dateOnly(endDate);
  const referralGoal = getInclusiveMonthCount(start, end);

  let referralTotal = 0;

  leadRows.forEach((row) => {
    const milestone = String(row["Current Milestone"] || "")
      .trim()
      .toLowerCase();

    if (milestone.includes("dead")) return;

    const rowPMValues = [
      row["Project Manager"],
      row["Salesperson"],
      row["Sales Rep"],
      row["Sales Representative"],
      row["Primary Salesperson"],
      row["Sales Owner"],
      row["Estimator"],
    ].map((value) => String(value || "").trim());

    if (!rowPMValues.includes(pmName)) return;

    const leadSource = String(row["Lead Source"] || "")
      .trim()
      .toLowerCase();

    if (
      !leadSource ||
      leadSource.replace(/\s+/g, "").indexOf("referral") === -1
    ) {
      return;
    }

    const rowDate = parseExcelDate(
      row["Initial Appointment Date"] ||
        row["Prospect Milestone Date"] ||
        row["Closed Milestone Date"] ||
        row["Created Date"] ||
        row["Date Created"] ||
        row["Approved Date"]
    );

    if (!rowDate) return;

    const cleanDate = dateOnly(rowDate);
    if (cleanDate < start || cleanDate > end) return;

    referralTotal += 1;
  });

  const referralDelta = referralTotal - referralGoal;

  return {
    referralTotal,
    referralGoal,
    referralDelta,
    referralPercent: referralGoal > 0 ? referralTotal / referralGoal : 0,
  };
};

const getCommissionMarginSummary = () => {
  const summary = {};

  marginRows.forEach((row) => {
    const tradeType =
      normalizeMarginTradeType(row["Job Trade Type"]) ||
      normalizeMarginTradeType(row["Job Trade Type 2"]);

    const workType = normalizeMarginWorkType(row["Work Type"]);

    if (!tradeType || !workType || workType === "Service") return;

    const key = `${tradeType}-${workType}`;
    const payment = getMarginRowPayment(row);
    const profit = getMarginRowProfit(row);
    const rowMargin = getMarginRowPercent(row);

    if (!summary[key]) {
      summary[key] = {
        tradeType,
        workType,
        payments: 0,
        profit: 0,
        count: 0,
        marginSum: 0,
        marginCount: 0,
      };
    }

    summary[key].payments += payment;
    summary[key].profit += profit;
    summary[key].count += 1;

    if (rowMargin !== null) {
      summary[key].marginSum += rowMargin;
      summary[key].marginCount += 1;
    }
  });

  return Object.values(summary).map((item) => {
    const calculatedMargin =
      item.payments > 0
        ? item.profit / item.payments
        : item.marginCount > 0
        ? item.marginSum / item.marginCount
        : 0;

    return {
      ...item,
      margin: calculatedMargin,
      rpp: item.count > 0 ? item.payments / item.count : 0,
    };
  });
};

const getJobTradeTypeOptions = () => {
  const marginSummary = getCommissionMarginSummary();

  const EXCLUDED_TRADE_TYPES = [
    "Other",
    "Other, Skylights",
    "Other Fascia",
    "Wraps, Fascia",
    "Storm Damage, Fascia",
    "Repair",
    "Wraps",
  ];

  if (!marginSummary.length) {
    return categories
      .map((category) => category.title)
      .filter(
        (trade) => !EXCLUDED_TRADE_TYPES.includes(trade)
      )
      .sort();
  }

  const tradeSalesTotals = marginSummary.reduce((totals, item) => {
    if (
      !item.tradeType ||
      EXCLUDED_TRADE_TYPES.includes(item.tradeType)
    ) {
      return totals;
    }

    totals[item.tradeType] =
      Number(totals[item.tradeType] || 0) +
      Number(item.payments || 0);

    return totals;
  }, {});

  return Object.keys(tradeSalesTotals).sort((a, b) => {
    const salesDifference =
      Number(tradeSalesTotals[b] || 0) -
      Number(tradeSalesTotals[a] || 0);

    if (salesDifference !== 0) return salesDifference;

    return a.localeCompare(b);
  });
};

const getCommissionWorkTypeOptions = () => {
  const marginSummary = getCommissionMarginSummary();
  const selectedTrades = selectedCommissionTrades || [];

  const availableWorkTypes = marginSummary
    .filter(
      (item) =>
        selectedTrades.length === 0 || selectedTrades.includes(item.tradeType)
    )
    .map((item) => item.workType)
    .filter((workType) => workType && workType !== "Service");

  const uniqueWorkTypes = [...new Set(availableWorkTypes)];

  if (!uniqueWorkTypes.length) return COMMISSION_WORK_TYPE_OPTIONS;

  return uniqueWorkTypes.sort((a, b) => {
    const aIndex = COMMISSION_WORK_TYPE_OPTIONS.indexOf(a);
    const bIndex = COMMISSION_WORK_TYPE_OPTIONS.indexOf(b);

    if (aIndex !== -1 || bIndex !== -1) {
      return (aIndex === -1 ? 999 : aIndex) - (bIndex === -1 ? 999 : bIndex);
    }

    return a.localeCompare(b);
  });
};

const getWorkTypeOptionsForTrade = (tradeType) => {
  const marginSummary = getCommissionMarginSummary();
  const marginWorkTypes = marginSummary
    .filter((item) => item.tradeType === tradeType)
    .map((item) => item.workType)
    .filter((workType) => workType && workType !== "Service");

  if (marginWorkTypes.length) {
    return [...new Set(marginWorkTypes)].sort(
      (a, b) =>
        WORK_TYPE_ORDER.indexOf(a) - WORK_TYPE_ORDER.indexOf(b)
    );
  }

  const category = categories.find((item) => item.title === tradeType);

  if (!category) return [];

  return category.items
    .map((item) => item.label)
    .filter((label) => label !== "Service")
    .sort(
      (a, b) =>
        WORK_TYPE_ORDER.indexOf(a) - WORK_TYPE_ORDER.indexOf(b)
    );
};

const getMarginForTradeAndWorkType = (tradeType, workType) => {
  const marginSummary = getCommissionMarginSummary();
  const marginMatch = marginSummary.find(
    (item) => item.tradeType === tradeType && item.workType === workType
  );

  if (marginMatch) return Number(marginMatch.margin || 0);

  const category = categories.find((item) => item.title === tradeType);

  if (!category) return 0;

  const matchingWorkType = category.items.find(
    (item) => item.label === workType
  );

  return Number(matchingWorkType?.margin || 0);
};
const roundCommission = (amount) => {
  const cents = amount * 100;
  const wholeCents = Math.floor(cents);
  const decimalPart = cents - wholeCents;

  if (Math.abs(decimalPart - 0.5) < 0.0001) {
    return (wholeCents + 1) / 100;
  }

  return wholeCents / 100;
};
  const getPMDashboardData = () => {
    const monthOptions = getPMMonthOptions((currentPM?.name || projectManagers[0].name));
    const selectedMonth = selectedPMMonth || monthOptions[0] || "";
    const lastYearMonth = getPreviousYearMonth(selectedMonth);

const pmName = currentPM?.name || projectManagers[0].name;
const fiscalMonthIndex = fiscalMonths.indexOf(selectedMonth);

const ytdMonths =
  fiscalMonthIndex >= 0
    ? fiscalMonths.slice(0, fiscalMonthIndex + 1)
    : [];

const getMonthDateBoundsForSales = (monthLabel) => {
  const [monthName, yearText] = String(monthLabel || "").split(" ");
  const monthIndex = monthNames.indexOf(monthName);
  const year = Number(yearText);

  if (monthIndex < 0 || !year) return null;

  return {
    start: new Date(year, monthIndex, 1),
    end: new Date(year, monthIndex + 1, 0),
  };
};

const getSalesDataForMonths = (months) => {
  const ranges = months
    .map(getMonthDateBoundsForSales)
    .filter(Boolean);

  if (!ranges.length) {
    return {
      contractTotal: 0,
      contracts: 0,
      averageContract: 0,
    };
  }

  const start = ranges[0].start;
  const end = ranges[ranges.length - 1].end;

  return getPMSalesDataForRange(pmName, start, end);
};

const fiscalYTDStartDate = new Date(2025, 10, 1);
const fiscalYTDEndDate = dateOnly(new Date());

const ytdSalesData = getPMSalesDataForRange(
  pmName,
  fiscalYTDStartDate,
  fiscalYTDEndDate
);

const ytdRevenue = ytdSalesData.contractTotal;
const ytdContracts = ytdSalesData.contracts;
const ytdAverageContract = ytdSalesData.averageContract;
const ytdClosingRateMonths = monthOptions.filter(
  (month) => getMonthSortValue(month) >= getMonthSortValue("January 2025")
);

const ytdClosingRate =
  ytdClosingRateMonths.length > 0
    ? ytdClosingRateMonths.reduce(
        (sum, month) =>
          sum + getPMMetric(pmName, "Monthly Closing Rate", month),
        0
      ) / ytdClosingRateMonths.length
    : 0;

const customMonths = (() => {
  if (pmDateMode !== "custom" || !pmStartDate || !pmEndDate) {
    return [];
  }

  const start = new Date(`${pmStartDate}T00:00:00`);
  const end = new Date(`${pmEndDate}T23:59:59`);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return [];
  }

  return monthOptions.filter((month) => {
    const [monthName, yearText] = String(month || "").split(" ");
    const monthIndex = monthNames.indexOf(monthName);
    const year = Number(yearText);

    if (monthIndex < 0 || !year) return false;

    const monthStart = new Date(year, monthIndex, 1);
    const monthEnd = new Date(year, monthIndex + 1, 0, 23, 59, 59);

    return monthEnd >= start && monthStart <= end;
  });
})();

const activeMonths =
  pmDateMode === "fiscalYTD"
    ? ytdMonths
    : pmDateMode === "custom"
    ? customMonths
    : [selectedMonth];

const customSalesData =
  pmDateMode === "custom" && pmStartDate && pmEndDate
    ? getPMSalesDataForRange(
        pmName,
        parseInputDate(pmStartDate),
        parseInputDate(pmEndDate)
      )
    : {
        contractTotal: 0,
        contracts: 0,
        averageContract: 0,
      };

const monthSalesData = getSalesDataForMonths([selectedMonth]);

const activeSalesData =
  pmDateMode === "fiscalYTD"
    ? ytdSalesData
    : pmDateMode === "custom"
    ? customSalesData
    : monthSalesData;

const contractTotal = activeSalesData.contractTotal;
const contracts = activeSalesData.contracts;
const averageContract = activeSalesData.averageContract;

const closingRate =
  pmDateMode === "fiscalYTD"
    ? ytdClosingRate
    : pmDateMode === "custom"
    ? activeMonths.length > 0
      ? activeMonths.reduce(
          (sum, month) =>
            sum + getPMMetric(pmName, "Monthly Closing Rate", month),
          0
        ) / activeMonths.length
      : 0
    : getPMMetric(pmName, "Monthly Closing Rate", selectedMonth);


const priorFiscalYTDStartDate = new Date(
  fiscalYTDStartDate.getFullYear() - 1,
  fiscalYTDStartDate.getMonth(),
  fiscalYTDStartDate.getDate()
);

const priorFiscalYTDEndDate = new Date(
  fiscalYTDEndDate.getFullYear() - 1,
  fiscalYTDEndDate.getMonth(),
  fiscalYTDEndDate.getDate()
);

const lastYearSalesData =
  pmDateMode === "fiscalYTD"
    ? getPMSalesDataForRange(
        pmName,
        priorFiscalYTDStartDate,
        priorFiscalYTDEndDate
      )
    : getSalesDataForMonths([lastYearMonth]);

const lyContractTotal = lastYearSalesData.contractTotal;
const lyContracts = lastYearSalesData.contracts;
const lyAverageContract = lastYearSalesData.averageContract;

const priorFiscalYTDMonths = monthOptions.filter(
  (month) =>
    getMonthSortValue(month) >= getMonthSortValue("January 2025") &&
    getMonthSortValue(month) <= getMonthSortValue("June 2025")
);

const priorFiscalYTDClosingRates = priorFiscalYTDMonths
  .map((month) => getPMMetric(pmName, "Monthly Closing Rate", month))
  .filter((rate) => Number(rate || 0) > 0);

const lyClosingRate =
  pmDateMode === "fiscalYTD"
    ? priorFiscalYTDClosingRates.length > 0
      ? priorFiscalYTDClosingRates.reduce((sum, rate) => sum + rate, 0) /
        priorFiscalYTDClosingRates.length
      : 0
    : getPMMetric(pmName, "Monthly Closing Rate", lastYearMonth);

const activeGoalProjectManagers = projectManagers.filter(
  (pm) =>
    pm.activeGoal &&
    pm.name !== "George Anim"
);

const customTeamSalesData = activeGoalProjectManagers.map((pm) =>
  getPMCustomSalesData(pm.name)
);

const activeCustomTeamSalesData = customTeamSalesData.filter(
  (item) => Number(item.contractTotal || 0) > 0
);

const customTeamContractTotal =
  activeCustomTeamSalesData.length > 0
    ? activeCustomTeamSalesData.reduce(
        (sum, item) => sum + Number(item.contractTotal || 0),
        0
      ) / activeCustomTeamSalesData.length
    : 0;

const customTeamContracts =
  activeCustomTeamSalesData.length > 0
    ? activeCustomTeamSalesData.reduce(
        (sum, item) => sum + Number(item.contracts || 0),
        0
      ) / activeCustomTeamSalesData.length
    : 0;

const customTeamTotalRevenue = customTeamSalesData.reduce(
  (sum, item) => sum + Number(item.contractTotal || 0),
  0
);

const customTeamTotalContracts = customTeamSalesData.reduce(
  (sum, item) => sum + Number(item.contracts || 0),
  0
);

const customTeamAverageContract =
  customTeamTotalContracts > 0
    ? customTeamTotalRevenue / customTeamTotalContracts
    : 0;

const customTeamClosingRate =
  activeMonths.length > 0 && activeGoalProjectManagers.length > 0
    ? activeGoalProjectManagers.reduce((pmSum, pm) => {
        const pmClosingRate =
          activeMonths.reduce(
            (monthSum, month) =>
              monthSum + getPMMetric(pm.name, "Monthly Closing Rate", month),
            0
          ) / activeMonths.length;

        return pmSum + pmClosingRate;
      }, 0) / activeGoalProjectManagers.length
    : 0;

const getTeamSalesDataForRange = (startDate, endDate) => {
  const teamSalesData = activeGoalProjectManagers.map((pm) =>
    getPMSalesDataForRange(pm.name, startDate, endDate)
  );

  const totalRevenue = teamSalesData.reduce(
    (sum, item) => sum + Number(item.contractTotal || 0),
    0
  );

  const totalContracts = teamSalesData.reduce(
    (sum, item) => sum + Number(item.contracts || 0),
    0
  );

const activeTeamSalesData = teamSalesData.filter(
    (item) => Number(item.contractTotal || 0) > 0
  );

  const activeTeamRevenue = activeTeamSalesData.reduce(
    (sum, item) => sum + Number(item.contractTotal || 0),
    0
  );

  const activeTeamContracts = activeTeamSalesData.reduce(
    (sum, item) => sum + Number(item.contracts || 0),
    0
  );

  return {
    contractTotal:
      activeTeamSalesData.length > 0
        ? activeTeamRevenue / activeTeamSalesData.length
        : 0,

    contracts:
      activeTeamSalesData.length > 0
        ? activeTeamContracts / activeTeamSalesData.length
        : 0,

    averageContract:
      activeTeamContracts > 0
        ? activeTeamRevenue / activeTeamContracts
        : 0,
  };
};

const getTeamSalesDataForMonths = (months) => {
  const ranges = months
    .map(getMonthDateBoundsForSales)
    .filter(Boolean);

  if (!ranges.length) {
    return {
      contractTotal: 0,
      contracts: 0,
      averageContract: 0,
    };
  }

  return getTeamSalesDataForRange(
    ranges[0].start,
    ranges[ranges.length - 1].end
  );
};

const ytdTeamClosingRate =
  ytdMonths.length > 0 && activeGoalProjectManagers.length > 0
    ? activeGoalProjectManagers.reduce((pmSum, pm) => {
        const pmClosingRate =
          ytdMonths.reduce(
            (monthSum, month) =>
              monthSum + getPMMetric(pm.name, "Monthly Closing Rate", month),
            0
          ) / ytdMonths.length;

        return pmSum + pmClosingRate;
      }, 0) / activeGoalProjectManagers.length
    : 0;

const teamSalesData =
  pmDateMode === "fiscalYTD"
    ? getTeamSalesDataForRange(fiscalYTDStartDate, fiscalYTDEndDate)
    : pmDateMode === "custom"
    ? customTeamSalesData.length > 0
      ? {
          contractTotal: customTeamContractTotal,
          contracts: customTeamContracts,
          averageContract: customTeamAverageContract,
        }
      : getTeamSalesDataForMonths(customMonths)
    : getTeamSalesDataForMonths([selectedMonth]);

const teamContractTotal = teamSalesData.contractTotal;
const teamContracts = teamSalesData.contracts;
const teamAverageContract = teamSalesData.averageContract;

const teamClosingRate =
  pmDateMode === "custom"
    ? customTeamClosingRate
    : pmDateMode === "fiscalYTD"
    ? ytdTeamClosingRate
    : getTeamMetric("Monthly Closing Rate Average", selectedMonth);


const ninetyDayEndDate = dateOnly(new Date());
const ninetyDayStartDate = new Date(ninetyDayEndDate);
ninetyDayStartDate.setDate(ninetyDayStartDate.getDate() - 89);

const ninetyDaySalesData = getPMSalesDataForRange(
  pmName,
  ninetyDayStartDate,
  ninetyDayEndDate
);

const ninetyDayTeamSalesData = getTeamSalesDataForRange(
  ninetyDayStartDate,
  ninetyDayEndDate
);

const ninetyDayClosingRate = getPMMetric(
  pmName,
  "Rolling 90-Day Closing Rate",
  selectedMonth
);

const ninetyDayTeamClosingRate = getTeamMetric(
  "Rolling 90-Day Closing Rate Average",
  selectedMonth
);

const rankDateRange =
  pmDateMode === "fiscalYTD"
    ? { start: fiscalYTDStartDate, end: fiscalYTDEndDate }
    : pmDateMode === "custom" && pmStartDate && pmEndDate
    ? {
        start: parseInputDate(pmStartDate),
        end: parseInputDate(pmEndDate),
      }
    : getMonthDateBoundsForSales(selectedMonth);

const revenueRank = rankDateRange
  ? getRevenueRankForRange(rankDateRange.start, rankDateRange.end, pmName)
  : { rank: null, total: 0 };
const closingRateRank = getRankForMetric(
  "Monthly Closing Rate",
  selectedMonth,
  (currentPM?.name || projectManagers[0].name)
);

    const companyTake = 0.10;

const saleAmount = Number(debouncedPmSaleAmount || 0);
const jobCost = Number(debouncedPmJobCost || 0);

const marginSummary = getCommissionMarginSummary();

const selectedTradeConfigs = selectedCommissionTrades
  .flatMap((tradeType) =>
    selectedCommissionWorkTypes
      .map((workType) => {
        const marginMatch = marginSummary.find(
          (item) => item.tradeType === tradeType && item.workType === workType
        );

        if (marginMatch) {
          return {
            tradeType,
            workType,
            rpp: Number(marginMatch.rpp || 0),
            margin: Number(marginMatch.margin || 0),
          };
        }

        const category = categories.find((item) => item.title === tradeType);
        const matchingItem = category?.items.find(
          (item) => item.label === workType
        );

        if (!matchingItem || workType === "Service") return null;

        return {
          tradeType,
          workType,
          rpp: Number(matchingItem.rpp || 0),
          margin: Number(matchingItem.margin || 0),
        };
      })
      .filter(Boolean)
  )
  .filter(Boolean);

const totalSelectedRpp = selectedTradeConfigs.reduce(
  (sum, item) => sum + Number(item.rpp || 0),
  0
);

const pmCommissionName = currentPM?.name || projectManagers[0].name;

const overheadAmount = Math.round(saleAmount * 0.1 * 100) / 100;

const grossProfit = Math.max(
  Math.round((saleAmount - jobCost - overheadAmount) * 100) / 100,
  0
);

const commissionRate = getPMCommissionRate(pmCommissionName);

const commission = roundCommission(
  grossProfit * commissionRate
);

const totalGrossProfit = grossProfit;

const effectiveCommissionRate = commissionRate;
const individualGoal =
  PM_GOALS[currentPM?.name || projectManagers[0].name] || 0;

const goalPercent = individualGoal > 0 ? ytdRevenue / individualGoal : 0;
const remainingToGoal = Math.max(individualGoal - ytdRevenue, 0);
const monthlyGoal = individualGoal / 12;

const monthlyGoalPercent =
  monthlyGoal > 0 ? contractTotal / monthlyGoal : 0;

const monthlyRemaining =
  monthlyGoal - contractTotal;
const now = new Date();

const selectedMonthBounds = getMonthDateBoundsForSales(selectedMonth);
const selectedMonthStart = selectedMonthBounds?.start || null;

const fiscalYearStart = selectedMonthStart
  ? selectedMonthStart.getMonth() >= 10
    ? new Date(selectedMonthStart.getFullYear(), 10, 1)
    : new Date(selectedMonthStart.getFullYear() - 1, 10, 1)
  : null;

const monthsSinceFiscalStart =
  selectedMonthStart && fiscalYearStart
    ? (selectedMonthStart.getFullYear() - fiscalYearStart.getFullYear()) * 12 +
      selectedMonthStart.getMonth() -
      fiscalYearStart.getMonth()
    : -1;

const quarterStartOffset =
  monthsSinceFiscalStart >= 0
    ? Math.floor(monthsSinceFiscalStart / 3) * 3
    : 0;

const quarterStartDate = fiscalYearStart
  ? new Date(
      fiscalYearStart.getFullYear(),
      fiscalYearStart.getMonth() + quarterStartOffset,
      1
    )
  : null;

const quarterEndDate = quarterStartDate
  ? new Date(
      quarterStartDate.getFullYear(),
      quarterStartDate.getMonth() + 3,
      0
    )
  : null;

  const formatShortDate = (date) =>
  date
    ? date.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
        year: "numeric",
      })
    : "";

const quarterDateRangeLabel =
  quarterStartDate && quarterEndDate
    ? `${formatShortDate(quarterStartDate)} - ${formatShortDate(quarterEndDate)}`
    : "";

const quarterDaysRemaining =
  quarterEndDate
    ? Math.max(
        Math.ceil(
          (dateOnly(quarterEndDate) - dateOnly(new Date())) /
            (1000 * 60 * 60 * 24)
        ),
        0
      )
    : 0;

const quarterMonths = quarterStartDate
  ? [0, 1, 2].map((offset) => {
      const date = new Date(
        quarterStartDate.getFullYear(),
        quarterStartDate.getMonth() + offset,
        1
      );

      return `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
    })
  : [];

const quarterSalesData =
  quarterStartDate && quarterEndDate
    ? getPMSalesDataForRange(pmName, quarterStartDate, quarterEndDate)
    : {
        contractTotal: 0,
        contracts: 0,
        averageContract: 0,
      };

const quarterRevenue = quarterSalesData.contractTotal;
const quarterTeamSalesData =
  quarterStartDate && quarterEndDate
    ? getTeamSalesDataForRange(quarterStartDate, quarterEndDate)
    : {
        contractTotal: 0,
        contracts: 0,
        averageContract: 0,
      };

const quarterClosingRate =
  quarterMonths.length > 0
    ? quarterMonths.reduce(
        (sum, month) =>
          sum + getPMMetric(pmName, "Monthly Closing Rate", month),
        0
      ) / quarterMonths.length
    : 0;

const quarterTeamClosingRate =
  quarterMonths.length > 0
    ? quarterMonths.reduce(
        (sum, month) =>
          sum + getTeamMetric("Monthly Closing Rate Average", month),
        0
      ) / quarterMonths.length
    : 0;

const getInclusiveDays = (start, end) => {
  const msPerDay = 1000 * 60 * 60 * 24;
  return Math.floor((dateOnly(end) - dateOnly(start)) / msPerDay) + 1;
};

const totalQuarterDays =
  quarterStartDate && quarterEndDate
    ? getInclusiveDays(quarterStartDate, quarterEndDate)
    : 0;

const elapsedQuarterDays =
  !quarterStartDate || !quarterEndDate
    ? 0
    : now < quarterStartDate
    ? 0
    : now > quarterEndDate
    ? totalQuarterDays
    : getInclusiveDays(quarterStartDate, now);

const quarterlyGoal = individualGoal / 4;

const quarterlyProjectedRevenue =
  elapsedQuarterDays > 0
    ? (quarterRevenue / elapsedQuarterDays) * totalQuarterDays
    : 0;

const quarterlyActualPercent =
  quarterlyGoal > 0 ? quarterRevenue / quarterlyGoal : 0;

const quarterlyProjectedPercent =
  quarterlyGoal > 0 ? quarterlyProjectedRevenue / quarterlyGoal : 0;

const quarterlyRemaining =
  quarterlyGoal - quarterRevenue;

const referralRange = (() => {
  if (pmDateMode === "fiscalYTD") {
    return {
      start: fiscalYTDStartDate,
      end: fiscalYTDEndDate,
    };
  }

  if (pmDateMode === "custom" && pmStartDate && pmEndDate) {
    return {
      start: parseInputDate(pmStartDate),
      end: parseInputDate(pmEndDate),
    };
  }

  return selectedMonthBounds;
})();

const referralData = referralRange
  ? getPMReferralDataForRange(pmName, referralRange.start, referralRange.end)
  : {
      referralTotal: 0,
      referralGoal: 0,
      referralDelta: 0,
      referralPercent: 0,
    };

const quarterlyReferralData =
  quarterStartDate && quarterEndDate
    ? getPMReferralDataForRange(pmName, quarterStartDate, quarterEndDate)
    : {
        referralTotal: 0,
        referralGoal: 0,
        referralDelta: 0,
        referralPercent: 0,
      };

const getReferralStatusForRange = (data, range) => {
  if (!range) {
    return data.referralDelta >= 0 ? "goalMet" : "goalNotMet";
  }

  const currentMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const currentMonthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  if (dateOnly(range.end) < currentMonthStart) {
    return data.referralDelta >= 0 ? "goalMet" : "goalNotMet";
  }

if (dateOnly(range.start) > currentMonthEnd) {
  return "future";
}

return data.referralDelta >= 0 ? "goalMet" : "inProgress";
};

const referralStatus =
  pmDateMode === "fiscalYTD" || pmDateMode === "custom"
    ? referralData.referralDelta >= 0
      ? "goalMet"
      : "goalNotMet"
    : getReferralStatusForRange(referralData, referralRange);

const quarterlyReferralStatus = getReferralStatusForRange(
  quarterlyReferralData,
  quarterStartDate && quarterEndDate
    ? { start: quarterStartDate, end: quarterEndDate }
    : null
);

const getClosedMonthOptions = () => {
  const today = dateOnly(new Date());

  return monthOptions
    .filter((month) => {
      const bounds = getMonthDateBoundsForSales(month);
      return bounds && dateOnly(bounds.end) < today;
    })
    .sort((a, b) => getMonthSortValue(b) - getMonthSortValue(a));
};

const getTopRevenueDawgForMonth = (month) => {
  const bounds = getMonthDateBoundsForSales(month);
  if (!bounds) return null;

  const ranked = projectManagers
    .filter((pm) => pm.activeGoal)
    .map((pm) => ({
      name: pm.name,
      value: getPMSalesDataForRange(pm.name, bounds.start, bounds.end)
        .contractTotal,
    }))
    .filter((pm) => Number(pm.value || 0) > 0)
    .sort((a, b) => b.value - a.value);

  if (!ranked.length) return null;

  const topValue = ranked[0].value;

  return {
    month,
    names: ranked.filter((pm) => pm.value === topValue).map((pm) => pm.name),
    value: topValue,
  };
};

const getTopClosingDawgForMonth = (month) => {
  const ranked = projectManagers
    .filter((pm) => pm.activeGoal)
    .map((pm) => ({
      name: pm.name,
      value: getPMMetric(pm.name, "Monthly Closing Rate", month),
    }))
    .filter((pm) => Number(pm.value || 0) > 0)
    .sort((a, b) => b.value - a.value);

  if (!ranked.length) return null;

  const topValue = ranked[0].value;

  return {
    month,
    names: ranked.filter((pm) => pm.value === topValue).map((pm) => pm.name),
    value: topValue,
  };
};

const closedMonths = getClosedMonthOptions();

const topDawgLeaderboards = {
  revenue: closedMonths.map(getTopRevenueDawgForMonth).filter(Boolean),
  closingRate: closedMonths.map(getTopClosingDawgForMonth).filter(Boolean),
};
const bestEverTopDawgs = {
  revenue: topDawgLeaderboards.revenue.reduce(
    (best, item) => (!best || item.value > best.value ? item : best),
    null
  ),
  closingRate: topDawgLeaderboards.closingRate.reduce(
    (best, item) => (!best || item.value > best.value ? item : best),
    null
  ),
};
  
return {
  monthOptions,
  selectedMonth,
  individualGoal,
      lastYearMonth,
      ytdRevenue,
      ytdContracts,
      ytdAverageContract,
      ytdClosingRate,
      contractTotal,
      contracts,
      averageContract,
      closingRate,
      lyContractTotal,
      lyContracts,
      lyAverageContract,
      lyClosingRate,
      teamContractTotal,
      teamContracts,
      teamAverageContract,
      teamClosingRate,
      ninetyDayContractTotal: ninetyDaySalesData.contractTotal,
      ninetyDayContracts: ninetyDaySalesData.contracts,
      ninetyDayAverageContract: ninetyDaySalesData.averageContract,
      ninetyDayClosingRate,
      ninetyDayTeamContractTotal: ninetyDayTeamSalesData.contractTotal,
      ninetyDayTeamContracts: ninetyDayTeamSalesData.contracts,
      ninetyDayTeamAverageContract: ninetyDayTeamSalesData.averageContract,
      ninetyDayTeamClosingRate,
      revenueRank,
      closingRateRank,
      topDawgLeaderboards,
      bestEverTopDawgs,
      saleAmount,
jobCost,
commission,
commissionRate: effectiveCommissionRate,
grossProfit: totalGrossProfit,
goalPercent,
      remainingToGoal,
      monthlyGoal,
      monthlyPace: contractTotal,
      quarterlyProjectedRevenue,
monthlyGoalPercent,
monthlyRemaining,
quarterRevenue,
quarterTeamContractTotal: quarterTeamSalesData.contractTotal,
quarterTeamContracts: quarterTeamSalesData.contracts,
quarterTeamAverageContract: quarterTeamSalesData.averageContract,
quarterClosingRate,
quarterTeamClosingRate,
quarterlyGoal,
quarterlyGoalPercent: quarterlyActualPercent,
quarterlyProjectedPercent,
quarterlyRemaining,
quarterDateRangeLabel,
quarterDaysRemaining,
referralTotal: referralData.referralTotal,
referralGoal: referralData.referralGoal,
referralDelta: referralData.referralDelta,
referralPercent: referralData.referralPercent,
referralStatus,
quarterlyReferralTotal: quarterlyReferralData.referralTotal,
quarterlyReferralGoal: quarterlyReferralData.referralGoal,
quarterlyReferralDelta: quarterlyReferralData.referralDelta,
quarterlyReferralPercent: quarterlyReferralData.referralPercent,
quarterlyReferralStatus,
    };
  };

  const handlePMLogin = (event) => {
    event.preventDefault();

    if (currentPM && pmPassword === currentPM.password) {
      sessionStorage.setItem(`${currentPM.slug}PMAuthorized`, "true");
      setPmAuthorized(true);
      setPmLoginError("");
      setPmPassword("");
    } else {
      setPmLoginError("Incorrect password. Please try again.");
    }
  };

  const pmLogout = () => {
    if (currentPM) sessionStorage.removeItem(`${currentPM.slug}PMAuthorized`);
    setPmAuthorized(false);
    setPmPassword("");
  };

  const reloadFiles = () => {
    loadLeadFile();
    loadSalesFile();
    loadPMFile();
    loadMarginFile();
  };

  const clearLeads = () => {
    setLeads(buildClearedCounts(true));
    triggerFlash();
  };

  const clearProjects = () => {
    setProjectStartDate("");
    setProjectEndDate("");
    triggerFlash();
  };

  const applyLeadCounts = (rows = leadRows) => {
    const start = dateOnly(parseInputDate(startDate));
    const end = dateOnly(parseInputDate(endDate));

    const counts = buildClearedCounts(true);

    rows.forEach((row) => {
      const milestone = String(row["Current Milestone"] || "")
        .trim()
        .toLowerCase();

      if (milestone.includes("dead")) return;

      const rowDate = parseExcelDate(
        row["Initial Appointment Date"] ||
          row["Prospect Milestone Date"] ||
          row["Closed Milestone Date"]
      );

      if (!rowDate) return;

      const cleanDate = dateOnly(rowDate);
      if (cleanDate < start || cleanDate > end) return;

      const trade =
        normalizeTrade(row["Job Trade Type 2"]) ||
        normalizeTrade(row["Job Trade Type"]);

      const workType = normalizeWorkType(row["Work Type"]);

      if (!trade || !workType) {
        counts[UNKNOWN_KEY] += 1;
        return;
      }

      const key = `${trade}-${workType}`;

      if (Object.prototype.hasOwnProperty.call(counts, key)) {
        counts[key] += 1;
      } else {
        counts[UNKNOWN_KEY] += 1;
      }
    });

    setLeads(counts);
  };

  const getCurrentProspectCount = () => {
    const start = dateOnly(parseInputDate(startDate));
    const end = dateOnly(parseInputDate(endDate));

    let count = 0;

    leadRows.forEach((row) => {
      const milestone = String(row["Current Milestone"] || "")
        .trim()
        .toLowerCase();

      if (!milestone.includes("prospect")) return;
      if (milestone.includes("dead")) return;

      const rowDate = parseExcelDate(
        row["Initial Appointment Date"] ||
          row["Prospect Milestone Date"] ||
          row["Closed Milestone Date"]
      );

      if (!rowDate) return;

      const cleanDate = dateOnly(rowDate);
      if (cleanDate < start || cleanDate > end) return;

      count += 1;
    });

    return count;
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

  const getProjectData = () => {
    if (!projectStartDate || !projectEndDate) {
      return {
        projectCounts: buildClearedCounts(true),
        projectRevenue: buildClearedCounts(true),
        totals: {
          revenue: 0,
          trueProjectProfit: 0,
          companyExpense: 0,
          companyProfit: 0,
          projectCount: 0,
          zeroProjectCount: 0,
        },
      };
    }

    const start = dateOnly(parseInputDate(projectStartDate));
    const end = dateOnly(parseInputDate(projectEndDate));

    const projectCounts = buildClearedCounts(true);
    const projectRevenue = buildClearedCounts(true);

    let revenue = 0;
    let trueProjectProfit = 0;
    let companyExpense = 0;
    let companyProfit = 0;
    let projectCount = 0;
    let zeroProjectCount = 0;

    salesRows.forEach((row) => {
      const rawAmount = row["Contract Amount"];

      if (rawAmount === "" || rawAmount === null || rawAmount === undefined) return;

      const contractAmount = parseMoney(rawAmount);

      const rowDate = parseExcelDate(row["Approved Date"]);
      if (!rowDate) return;

      const cleanDate = dateOnly(rowDate);
      if (cleanDate < start || cleanDate > end) return;

      const trade =
        normalizeTrade(row["Job Trade Type 2"]) ||
        normalizeTrade(row["Job Trade Type"]);

      const workType = normalizeWorkType(row["Work Type"]);
      const config = findProjectConfig(trade, workType);

      projectCounts[config.key] = Number(projectCounts[config.key] || 0) + 1;

      projectRevenue[config.key] =
        Number(projectRevenue[config.key] || 0) + contractAmount;

      const breakdown = getProfitBreakdown(contractAmount, config.margin);

      if (contractAmount === 0) zeroProjectCount += 1;

      projectCount += 1;
      revenue += contractAmount;
      trueProjectProfit += breakdown.trueProjectProfit;
      companyExpense += breakdown.companyExpense;
      companyProfit += breakdown.companyProfit;
    });

    return {
      projectCounts,
      projectRevenue,
      totals: {
        revenue,
        trueProjectProfit,
        companyExpense,
        companyProfit,
        projectCount,
        zeroProjectCount,
      },
    };
  };

  const getLeadCategoryTotals = (category) => {
    return category.items.reduce(
      (totals, item) => {
        const key = `${category.title}-${item.label}`;
        const quantity = Number(leads[key] || 0);
        const revenue = quantity * activeCloseRate * item.rpp;
        const breakdown = getProfitBreakdown(revenue, item.margin);

        return {
          quantity: totals.quantity + quantity,
          revenue: totals.revenue + revenue,
          trueProjectProfit: totals.trueProjectProfit + breakdown.trueProjectProfit,
          companyProfit: totals.companyProfit + breakdown.companyProfit,
        };
      },
      {
        quantity: 0,
        revenue: 0,
        trueProjectProfit: 0,
        companyProfit: 0,
      }
    );
  };

  const getProjectCategoryTotals = (category) => {
    return category.items.reduce(
      (totals, item) => {
        const key = `${category.title}-${item.label}`;
        const quantity = Number(projectData.projectCounts[key] || 0);
        const revenue = Number(projectData.projectRevenue[key] || 0);
        const breakdown = getProfitBreakdown(revenue, item.margin);

        return {
          quantity: totals.quantity + quantity,
          revenue: totals.revenue + revenue,
          trueProjectProfit: totals.trueProjectProfit + breakdown.trueProjectProfit,
          companyProfit: totals.companyProfit + breakdown.companyProfit,
        };
      },
      {
        quantity: 0,
        revenue: 0,
        trueProjectProfit: 0,
        companyProfit: 0,
      }
    );
  };

  useEffect(() => {
    if (!authorized) return;

    reloadFiles();

    return () => {
      if (flashTimeoutRef.current) clearTimeout(flashTimeoutRef.current);
    };
  }, [authorized]);

  useEffect(() => {
    if (!authorized || !leadRows.length || !startDate || !endDate) return;

    const timeout = setTimeout(() => {
      applyLeadCounts(leadRows);
    }, 150);

    return () => clearTimeout(timeout);
  }, [authorized, leadRows, startDate, endDate]);


useEffect(() => {
  if (!pmAuthorized) return;

  loadPMFile();
  loadSalesFile();
  loadLeadFile();
  loadMarginFile();
}, [pmAuthorized]);

useEffect(() => {
  const timer = setTimeout(() => {
    setDebouncedPmSaleAmount(pmSaleAmount);
  }, 500);

  return () => clearTimeout(timer);
}, [pmSaleAmount]);

useEffect(() => {
  const timer = setTimeout(() => {
    setDebouncedPmJobCost(pmJobCost);
  }, 500);

  return () => clearTimeout(timer);
}, [pmJobCost]);
  useEffect(() => {
  if (!pmRows.length || selectedPMMonth) return;

  const options = getPMMonthOptions((currentPM?.name || projectManagers[0].name));

  if (!options.length) return;

  const currentMonth = new Date().toLocaleDateString("en-US", {
    month: "long",
    year: "numeric",
  });

  const matchingMonth = options.find(
    (month) => month === currentMonth
  );

  setSelectedPMMonth(
    matchingMonth || options[0]
  );
}, [pmRows, selectedPMMonth]);

  const pmData = currentPM ? getPMDashboardData() : null;

  const revenueGoatPM = projectManagers.find((pm) =>
  pmData?.bestEverTopDawgs?.revenue?.names?.includes(pm.name)
);

const closingRateGoatPM = projectManagers.find((pm) =>
  pmData?.bestEverTopDawgs?.closingRate?.names?.includes(pm.name)
);

  useEffect(() => {
    const hasSale = Number(cleanMoneyInput(pmSaleAmount)) > 0;
    const hasJobCost = Number(cleanMoneyInput(pmJobCost)) > 0;
    const currentCommission = Number(pmData?.commission || 0);

    if (!hasSale || !hasJobCost || currentCommission <= 0) {
      previousCommissionRef.current = currentCommission;
      return;
    }

    if (currentCommission === previousCommissionRef.current) {
      return;
    }

    const confettiDelay = setTimeout(() => {
      const rect = commissionCardRef.current?.getBoundingClientRect();

      if (rect) {
jsConfettiRef.current?.addConfetti({
  emojis: ["💵", "💸", "💰"],
  emojiSize: 75,
  confettiNumber: 100,
});
      }

      previousCommissionRef.current = currentCommission;
    }, 900);

    return () => clearTimeout(confettiDelay);
  }, [pmSaleAmount, pmJobCost, pmData?.commission]);

  const leadTotals = getLeadTotals();
  const projectData = getProjectData();
  const projectTotals = projectData.totals;
  const marketingForecast = getMarketingForecast();
  const unknownLeadCount = Number(leads[UNKNOWN_KEY] || 0);
  const currentProspectCount = getCurrentProspectCount();

  const Header = () => (
    <header className="header">
      {screen !== "home" && (
        <button className="back-button-header" onClick={() => setScreen("home")}>
          ← Home
        </button>
      )}

      <button className="header-action-button" onClick={logout}>
        Logout
      </button>

      <div className="header-top">
        <img src="/logo.png" alt="Logo" className="pm-header-logo" />
        <h1>
          {screen === "leads"
            ? "Lead Forecasted Revenue & Profit"
            : screen === "projects"
            ? "Projects Revenue and Margin Forecast"
            : screen === "marketing"
            ? "Marketing Forecast"
            : "Lead & Project Dashboard"}
        </h1>
      </div>
    </header>
  );

  const CategoryLeadTotal = ({ category }) => {
    const totals = getLeadCategoryTotals(category);

    return (
      <div className="category-total">
        <h4>{category.title} Total</h4>

        <div>
          <span>Leads</span>
          <strong>{totals.quantity}</strong>
        </div>

        <div>
          <span>Revenue</span>
          <strong>{money(totals.revenue)}</strong>
        </div>

        <div>
          <span>Project Profit</span>
          <strong className="total-green">{money(totals.trueProjectProfit)}</strong>
        </div>

        <div>
          <span>Company Profit</span>
          <strong className="total-red">{money(totals.companyProfit)}</strong>
        </div>
      </div>
    );
  };

  const CategoryProjectTotal = ({ category }) => {
    const totals = getProjectCategoryTotals(category);

    return (
      <div className="category-total">
        <h4>{category.title} Total</h4>

        <div>
          <span>Projects</span>
          <strong>{totals.quantity}</strong>
        </div>

        <div>
          <span>Revenue</span>
          <strong>{money(totals.revenue)}</strong>
        </div>

        <div>
          <span>Project Profit</span>
          <strong className="total-green">{money(totals.trueProjectProfit)}</strong>
        </div>

        <div>
          <span>Company Profit</span>
          <strong className="total-red">{money(totals.companyProfit)}</strong>
        </div>
      </div>
    );
  };

  const SummaryTotals = ({
    revenue,
    trueProjectProfit,
    companyProfit,
    companyExpense,
    revenueLabel,
  }) => (
    <div className="section-summary-grid">
      <div className={`section-total red ${flash ? "glow" : ""}`}>
        <strong>
          <AnimatedMoney value={revenue} />
        </strong>
        <span>{revenueLabel}</span>
      </div>

      <div
        className={`section-total white profit-total finance-total ${
          flash ? "glow" : ""
        }`}
      >
        <div className="finance-main-row">
          <span>True Project Profit</span>
          <strong className="true-profit-number">
            <AnimatedMoney value={trueProjectProfit} />
          </strong>
        </div>

        <div className="finance-divider" />

        <div className="finance-row expense">
          <span>Company Expense (10%)</span>
          <strong className="expense-number">-{money(companyExpense)}</strong>
        </div>

        <div className="finance-row company">
          <span>True Company Profit</span>
          <strong className="company-number">{money(companyProfit)}</strong>
        </div>
      </div>
    </div>
  );


  if (isPMPortal && !pmAuthorized) {
    return (
      <div className="login-screen">
        <form className="login-box" onSubmit={handlePMLogin}>
          <img src={currentPM.image} alt={(currentPM?.name || projectManagers[0].name)} className="pm-login-photo" />

          <h1>{(currentPM?.name || projectManagers[0].name)}</h1>
          <p>Enter your password to view your dashboard.</p>

          <input
            type="text"
            placeholder="Password"
            value={pmPassword}
            autoFocus
            onChange={(event) => {
              setPmPassword(event.target.value);
              setPmLoginError("");
            }}
          />

          {pmLoginError && <div className="login-error">{pmLoginError}</div>}

          <button type="submit">Login</button>
        </form>
      </div>
    );
  }

  if (isPMPortal && pmAuthorized && pmData) {
const revenueRankLabel = pmData.revenueRank.rank
  ? `#${pmData.revenueRank.rank}`
  : "N/A";

const closingRankLabel = pmData.closingRateRank.rank
  ? `#${pmData.closingRateRank.rank}`
  : "N/A";
const isRevenueLeader = Number(pmData.revenueRank?.rank) === 1;
const isClosingLeader = Number(pmData.closingRateRank?.rank) === 1;
const revenueVsLY = compareNumbers(
  pmData.contractTotal,
  pmData.lyContractTotal
);

const contractsVsLY = compareNumbers(
  pmData.contracts,
  pmData.lyContracts
);

const averageVsLY = compareNumbers(
  pmData.averageContract,
  pmData.lyAverageContract
);

const closingRateVsLY = compareNumbers(
  pmData.closingRate,
  pmData.lyClosingRate,
  true
);

const revenueVsTeam = compareNumbers(
  pmData.contractTotal,
  pmData.teamContractTotal
);

const contractsVsTeam = compareNumbers(
  pmData.contracts,
  pmData.teamContracts
);

const averageVsTeam = compareNumbers(
  pmData.averageContract,
  pmData.teamAverageContract
);

const closingRateVsTeam = compareNumbers(
  pmData.closingRate,
  pmData.teamClosingRate,
  true
);
const quarterRevenueVsTeam = compareNumbers(
  pmData.quarterRevenue,
  pmData.quarterTeamContractTotal
);

const quarterContractsVsTeam = compareNumbers(
  pmData.contracts,
  pmData.teamContracts
);

const quarterClosingRateVsTeam = compareNumbers(
  pmData.quarterClosingRate,
  pmData.quarterTeamClosingRate,
  true
);

const exceededGoalMessages = [
  "Outstanding work bub. You've surpassed your goal for this month. That's fire.",
  "Excellent performance. You've exceeded expectations and continue to set the pace. That's fire.",
  "You've reached your goal and are building momentum. That's fire.",
  "Fantastic job. Your production cleared the goal for this month. That's fire.",
  "You're operating at a very high level and have successfully exceeded your goal. That's fire.",
  "Bub, that's fire. You didn't just hit the goal, you blew right past it. Keep deep driving.",
  "One team one dream. You're showing everybody what championship-level effort looks like.",
  "You're killing it bub. The work you've put in is showing up in a big way.",
  "No excuses, only progress. That's exactly the mindset that creates results like this.",
  "Mission accomplished. The goal was the target, but you're proving there's another level available.",
  "Keep pushing bub. Great performers stay hungry even after they win.",
  "That's fire. You're building momentum that can carry through the entire quarter.",
  "Deep driving every day created this result. Don't let off the gas now.",
  "Bub, you're setting the standard right now and the numbers prove it.",
  "One team one dream. Keep stacking wins and making things happen."
];

const onPaceMessages = [
  "You're on pace to reach your goal and finish the period strong. Let's go!",
  "Solid progress. You're on pace to reach your monthly goal. Keep it up!",
  "You're tracking well toward your goal and maintaining healthy momentum.",
  "Keep doing what you're doing. You're positioned to reach your target.",
  "Performance remains strong and you're currently pacing toward success. That's fire.",
  "Bub, you're right where you need to be. Keep pushing.",
  "One team one dream. Consistency right now is your biggest advantage.",
  "Stay focused and keep stacking wins - you got this!",
  "You got this bub. The finish line is in sight.",
  "No excuses, only progress. Keep moving forward one contract at a time.",
  "Momentum is building. Keep deep driving and trust the process.",
  "You're creating opportunities and that's exactly what winners do.",
  "Stay locked in. A strong finish can turn a good month into a great month.",
  "Keep your foot on the gas bub. You're trending in the right direction.",
  "You're earning the right to succeed through consistent effort. Keep going."
];

const progressMessages = [
  "You're making steady progress toward your goal with some time remaining.",
  "Momentum is building and there is still plenty of opportunity ahead.",
  "Consistent effort now can create a strong finish to the period.",
  "The foundation is there. Continue focusing on quality opportunities.",
  "Progress remains positive and there is room to accelerate, keep pushing.",
  "Bub, good things are happening. Keep pushing.",
  "One team one dream. Every opportunity matters.",
  "That's fire. The momentum is starting to build.",
  "You got this. Stay focused on the next opportunity.",
  "No excuses, only progress. Keep moving forward.",
  "Deep driving today creates results tomorrow.",
  "The scoreboard doesn't tell the whole story. Keep working.",
  "Stay consistent bub. Consistency wins over time.",
  "The goal is still right in front of you. Stay after it.",
  "Progress is progress. Keep stacking positive days."
];

const pushMessages = [
  "Keep pushing bub. A strong finish can still put you back on pace.",
  "There is plenty of opportunity remaining to improve results - you got this.",
  "Focus on the trainings we've had and the numbers will follow.",
  "A few strong weeks can quickly change the trajectory. You got this!",
  "Stay consistent and continue building momentum, there's still time!",
  "Don't quit bub. The month isn't over.",
  "One team one dream. Keep showing up and doing the work.",
  "No excuses, only progress. Focus on the next opportunity.",
  "You got this. One good week can change everything.",
  "Keep deep driving. Results can change fast.",
  "Stay positive and keep moving forward.",
  "The comeback starts with the next appointment.",
  "Bub, champions don't quit when things get hard.",
  "Keep pushing. The scoreboard can change quicker than you think.",
  "Every conversation is a chance to build momentum."
];

const revenueStrongMessages = [
  "Revenue production is outperforming the team average - that's fire.",
  "Your revenue generation continues to lead the way - great job.",
  "You're producing above the team's current pace.",
  "Revenue performance remains a key strength - keep up the great work.",
  "Your production levels are setting a strong example - way to go bub.",
  "Bub, that's fire. You're outperforming the team average.",
  "One team one dream. You're helping raise the standard.",
  "You're killing it. Revenue production continues to impress.",
  "Keep pushing. Strong numbers come from strong habits.",
  "No excuses, only progress. You're proving it every day.",
  "Hard work is paying off in a big way.",
  "You're creating value and it shows in the numbers.",
  "Stay hungry. There's still more out there.",
  "The effort is showing up on the scoreboard.",
  "Keep stacking productive days."
];

const revenueEliteMessages = [
  "Revenue production is significantly outperforming the team average.",
  "You're operating well above the team's current production pace.",
  "Your revenue performance ranks among the strongest on the team - amazing job.",
  "Production continues to separate you from the field - keep pushing the needle.",
  "You're creating substantial value compared with team benchmarks. Great work.",
  "Bub, that's fire. You're not just ahead, you're creating separation.",
  "One team one dream, but you're carrying a big piece of the load right now.",
  "You're killing it. Revenue production like this changes the game.",
  "Keep pushing. Elite performers stay focused even when they're winning.",
  "No excuses, only progress. The work is showing up in a huge way.",
  "You got this. Keep creating opportunities and the results follow.",
  "Deep driving every day is creating elite-level production.",
  "Bub, don't get comfortable. There's another gear in there.",
  "You're setting the pace for everyone else right now.",
  "That's fire. Keep building on the momentum you've created."
];

const revenueNeedsMessages = [
  "Revenue is currently below the team average, but opportunities remain.",
  "Closing a few additional projects could quickly narrow the gap.",
  "There is still time to improve revenue performance this period.",
  "The team benchmark remains within reach.",
  "Focus on pipeline conversion and revenue growth will follow.",
  "Bub, don't quit. Revenue can move fast when you stay locked in.",
  "No excuses, only progress. The next contract is where the comeback starts.",
  "You got this. Keep deep driving and keep creating chances.",
  "One team one dream. Lean into the process and the numbers will follow.",
  "Keep pushing bub. A few strong closes can change the whole picture.",
  "Stay focused on quality appointments and strong follow-up.",
  "The scoreboard is not permanent. Keep working the pipeline.",
  "Deep driving means staying after it when the gap is uncomfortable.",
  "Bub, the opportunity is still there. Go take it one conversation at a time.",
  "Progress starts with activity. Keep moving and don't quit."
];

const avgStrongMessages = [
  "Your average contract value continues to outperform the team.",
  "Customers are trusting you with larger projects than average.",
  "Contract quality remains one of your strengths.",
  "Higher-value projects are driving strong results.",
  "Your average sale size remains above benchmark levels.",
  "Bub, that's fire. You're not just selling, you're building real project value.",
  "One team one dream. Bigger opportunities help lift the whole team.",
  "You're killing it with project scope. Keep deep driving.",
  "Strong average contracts tell me customers trust your process.",
  "Keep pushing. Quality deals create quality months.",
  "No excuses, only progress. You're showing what strong value-building looks like.",
  "You got this. Keep asking the right questions and uncovering the full opportunity.",
  "That's fire bub. Bigger projects come from deeper conversations.",
  "Your average contract strength is giving you a real advantage.",
  "Keep building value and the numbers will keep rewarding you."
];

const avgNeedsMessages = [
  "Increasing average contract value could have a meaningful impact.",
  "Larger project opportunities could accelerate growth.",
  "Focusing on project scope may improve overall production.",
  "Even modest gains in average contract size would create significant upside.",
  "Exploring higher-value opportunities may strengthen results.",
  "Bub, keep deep driving into the full scope. There may be more opportunity there.",
  "No excuses, only progress. Bigger questions can uncover bigger projects.",
  "You got this. Focus on value, not just the first need they mention.",
  "One team one dream. Stronger project scope helps everyone win.",
  "Keep pushing to understand the whole problem, not just the obvious one.",
  "Don't quit on the discovery process. That's where larger opportunities show up.",
  "Bub, the next level may be hiding in better questions.",
  "Build value, solve problems, and the average contract can climb.",
  "Deep driving on scope can turn a small job into a real win.",
  "Stay curious with the customer. There may be more on the table."
];

const closeStrongMessages = [
  "Your closing efficiency remains one of your strongest advantages.",
  "Customers continue responding well to your sales process.",
  "Conversion performance is helping drive strong results.",
  "Your closing rate continues to outperform expectations.",
  "Strong conversion efficiency remains a competitive advantage.",
  "Bub, that's fire. You're turning opportunities into contracts.",
  "You're killing it in the close. Keep trusting the process.",
  "One team one dream. Strong conversion helps raise the whole standard.",
  "Keep pushing. Your follow-up and execution are showing up.",
  "No excuses, only progress. You're closing with confidence right now.",
  "You got this. Keep listening, solving, and asking for the business.",
  "Deep driving through objections is paying off.",
  "That's fire bub. Customers are responding to how you're leading the conversation.",
  "Your close rate says the process is working. Don't get loose now.",
  "Keep stacking wins. Every close builds more confidence."
];

const closeNeedsMessages = [
  "Improving your conversion efficiency could unlock additional growth.",
  "Small gains in closing rate could have a large impact on production.",
  "More effective follow-up may improve conversion results.",
  "There is opportunity to strengthen closing performance.",
  "Improved conversion rates could significantly boost revenue.",
  "Bub, don't quit. Tighten the follow-up and the closes will come.",
  "No excuses, only progress. Every appointment is a chance to sharpen the process.",
  "You got this. Focus on listening, solving, and asking clearly for the business.",
  "Keep deep driving through the customer's real concerns.",
  "One team one dream. Use the training, trust the process, and keep pushing.",
  "The close starts long before the paperwork. Stay engaged from the first conversation.",
  "Bub, small improvements here can change the whole month.",
  "Don't let one no slow you down. Learn from it and move forward.",
  "Keep pushing. Better follow-up can turn maybe into yes.",
  "Stay disciplined in the process. The next close is waiting."
];



const pickRandom = (messages) => {
  if (!messages || !messages.length) return "";

  const key = messages.join("|");

  let hash = 0;

  for (let i = 0; i < key.length; i += 1) {
    hash = (hash * 31 + key.charCodeAt(i)) | 0;
  }

  const index = Math.abs(hash) % messages.length;

  return messages[index];
};

const onTrackMessages = [
  "You're on track to achieve your goal bub. Continue focusing on consistent activity and strong follow-up. The momentum you're building now can create an exceptional finish.",
  "Great progress bub. Keep pushing toward the finish line. Staying disciplined over the next few weeks will put you in a strong position to exceed expectations.",
  "You're in a strong position to finish successfully. Continue working your pipeline and capitalizing on opportunities. The goal is well within reach.",
  "Steady effort is producing solid results bub. Keep executing the fundamentals and trust the process. Consistency is often what separates good months from great ones.",
  "The goal is well within reach. Continue focusing on quality opportunities and maintaining urgency. A strong finish could put you comfortably over target.",
  "Momentum is building in the right direction. Stay committed to the daily activities that drive results. You're putting yourself in a position to succeed.",
  "You're trending toward a successful month. Keep creating opportunities and staying engaged with your prospects. The hard work is showing up in the numbers.",
  "Stay focused and keep stacking wins. Small victories compound quickly and can create significant momentum. You're heading in the right direction bub!",
  "Your consistency is paying off. Continue building on the progress you've already made. Every additional contract strengthens your position.",
  "Keep the pressure on and finish strong. You're tracking toward your goal and have a great opportunity to exceed it if the momentum continues.",
  "Bub, that's fire. You're on track because you've been showing up and doing the work. Keep deep driving and don't let the pace slip.",
  "One team one dream. You're putting yourself in a position to win, and that helps push the whole team forward.",
  "You got this. The goal is not some faraway thing right now, it's right in front of you. Keep pushing.",
  "No excuses, only progress. Stay locked into the daily actions that got you here.",
  "You're killing it with consistency. Now keep stacking days and turn on-track into over-goal.",
  "Bub, stay hungry. Being on pace is good, but finishing ahead is better.",
  "That's fire. Your momentum is real, and now the job is to protect it.",
  "Keep deep driving into every lead, every appointment, and every follow-up. That's how strong finishes happen.",
  "One team one dream. Keep showing what disciplined execution looks like.",
  "Don't coast now bub. You're in position, but the finish still matters.",
  "Keep pushing. A few more wins can make this month special.",
  "The process is working. Trust it, repeat it, and finish with urgency.",
  "You're building a strong month one decision at a time. Stay focused.",
  "Bub, every contract from here is extra pressure on the scoreboard in the best way.",
  "That's fire. Keep the conversations moving and the pipeline active.",
  "You are doing the right things. Now finish the job.",
  "No excuses, only progress. Keep building momentum until the month closes.",
  "Deep driving is the difference between being close and being over the top.",
  "You're on track, bub. Now go make sure there is no doubt.",
  "Keep pushing. The month is lining up for a strong finish if you stay locked in."
];

const closeMessages = [
  "You're closer than you think. Keep pushing. A few additional contracts could completely change the outcome of this month.",
  "A strong finish can put you over the top. Stay focused on the opportunities directly in front of you.",
  "The goal remains within reach. Consistent activity over the next few days can make all the difference.",
  "Keep building momentum. You're not too far away. Continue working your process and trust the results will come.",
  "Stay focused. A few more wins could make the difference. The opportunity to finish strong is still there.",
  "The goal is still achievable. Continue following up and creating opportunities. Success often comes from persistence.",
  "You're in striking distance of your target. Keep your energy high and maintain urgency in every conversation.",
  "Now is the time to finish strong. Every appointment and estimate matters from this point forward.",
  "Stay disciplined and keep moving forward. The finish line is much closer than it appears.",
  "Every contract matters from here. Continue stacking small wins and the larger results will follow.",
  "Bub, you're close. This is where you keep deep driving and refuse to let the month slip away.",
  "One team one dream. You're right there, and now it's about finishing with urgency.",
  "You got this. One strong close can change the whole feel of the month.",
  "No excuses, only progress. Stay on the phone, stay in the follow-up, and stay active.",
  "That's fire territory, bub. You're close enough to make this thing happen.",
  "Keep pushing. The difference between close and done is execution.",
  "Don't quit now. You're too close to start playing careful.",
  "Deep driving on the next few opportunities could be the difference-maker.",
  "Bub, stay locked in. The next yes matters.",
  "The month is asking you to finish. Go answer it.",
  "You are right there. Keep your energy high and your follow-up tight.",
  "One team one dream. Finish strong for yourself and for the team.",
  "Keep stacking small wins. Small wins are how big finishes happen.",
  "No excuses, only progress. The scoreboard can still move.",
  "You got this bub. Work the process and make the next opportunity count.",
  "Keep deep driving. A strong close to the month is still on the table.",
  "That's fire if you finish this thing right. Stay after it.",
  "You're not far off. Keep pushing until the final whistle.",
  "Every conversation matters now. Treat it that way.",
  "Bub, don't coast into close. Attack the finish."
];


const needsPushMessages = [
  "Keep grinding bub. Consistent effort pays off. The month is not defined by where you are today but by how you finish.",
  "The month isn't over yet. Stay focused. One strong week can completely change the trajectory of your results.",
  "Every opportunity matters from this point forward. Keep creating conversations and the numbers will follow.",
  "Small wins add up quickly. Focus on the next appointment, the next estimate, and the next contract.",
  "Keep working the process and results will follow. Consistency is often the difference between average and exceptional performance.",
  "Stay persistent and keep creating opportunities. Success often comes right after the point where others give up.",
  "There's still time to improve your position. Continue executing the fundamentals and trust the process.",
  "Focus on the next opportunity, not the scoreboard. Daily activity creates long-term success.",
  "Progress starts with consistent activity. Continue building momentum one conversation and one contract at a time.",
  "Keep your energy high and stay engaged. Every productive day moves you closer to your goals.",
  "Bub, don't quit. This is where you lean in and keep deep driving.",
  "No excuses, only progress. The next action is what matters now.",
  "You got this. A slow start does not mean a slow finish.",
  "One team one dream. Stay connected, ask for help, and keep moving.",
  "Keep pushing. The work you do today can change what the scoreboard says tomorrow.",
  "Bub, champions respond. They don't make excuses and they don't disappear.",
  "Deep driving through the tough stretch is how you build the comeback.",
  "That's fire when you keep fighting even when the numbers are not where you want them.",
  "The month is still alive. Go create momentum.",
  "Stay focused on activity. Activity creates opportunity, and opportunity creates contracts.",
  "Don't let frustration slow your feet. Keep moving.",
  "Bub, the only bad response is quitting. Keep pushing.",
  "No excuses, only progress. Make the next call and take the next step.",
  "You are capable of more than this snapshot. Go prove it.",
  "The comeback starts with one appointment, one estimate, one close.",
  "Keep deep driving. The process still works when you work it.",
  "One team one dream. We keep moving together.",
  "You got this bub. Reset, refocus, and go win the next conversation.",
  "Don't quit. There is still room to make progress.",
  "Keep pushing until the month is done. That's how winners operate."
];

const futureMonthMessages = [
  "We're not here yet, bub, but I know you'll be ready for whatever comes your way.",
  "The scoreboard is still blank, but champions prepare before the game starts.",
  "Every great month starts at zero. Let's go build something special.",
  "The opportunity is in front of you bub. Trust the process and attack the month.",
  "Success leaves clues, and you've already proven you know how to win.",
  "A strong month begins long before the first contract is signed.",
  "Stay focused, stay hungry, and let the results take care of themselves.",
  "You've built momentum before, and there's no reason this month can't be another great one.",
  "The month is unwritten, but I like our chances.",
  "Future-you is counting on present-you. Let's get after it.",
  "You know the playbook. Now it's time to execute it.",
  "Preparation today becomes production tomorrow.",
  "The month hasn't started yet bub, but the confidence should.",
  "We're not there yet, bub, but I already know you're going to be ready.",
  "The work you do now is what makes future months successful.",
  "Bub, that's fire. A blank scoreboard means nothing but opportunity.",
  "One team one dream. We prepare before the month ever starts.",
  "You got this. Start with belief, then back it up with activity.",
  "No excuses, only progress. The next month is earned before it opens.",
  "Keep pushing now so the future month starts with momentum.",
  "Deep driving starts before the first appointment hits the calendar.",
  "Don't wait for the month to begin before you decide to win it.",
  "Bub, winners don't wait for perfect timing. They prepare early.",
  "That's fire when preparation turns into production.",
  "The opportunity is coming. Be ready to attack it.",
  "One team one dream. The whole team wins when everybody prepares.",
  "You are killing it when you treat future opportunity like current responsibility.",
  "Keep your head right and your habits sharp.",
  "The next month is a fresh field. Go plant something strong.",
  "Bub, the goal starts at zero, but your preparation should not.",
  "You got this. Build the plan now and execute when the time comes.",
  "No excuses, only progress. Future wins are built by today's discipline.",
  "Keep pushing in the quiet moments. That's where separation starts.",
  "Deep driving now gives you a faster start later.",
  "The best months are not accidents. They're prepared for.",
  "Bub, get ready to make noise when this month opens up.",
  "One team one dream. Stay ready so we don't have to get ready.",
  "That's fire. The next opportunity is already getting lined up.",
  "Don't quit on preparation just because the scoreboard hasn't started.",
  "You know how to win. Now get ready to prove it again.",
  "Keep building the pipeline before the month asks for results.",
  "Bub, confidence comes from doing the work early.",
  "No excuses, only progress. Start sharp, stay sharp, finish strong.",
  "The future month is waiting. Let's make sure it meets a prepared version of you.",
  "You got this. When the month opens, attack it like it already belongs to you."
];



const goalMotivation = (() => {
  const pct = pmData.monthlyGoalPercent;
  const firstName = (currentPM?.name || "").split(" ")[0];

  const selectedDate = new Date(pmData.selectedMonth);
  const today = new Date();

  const monthIsPast =
    selectedDate.getFullYear() < today.getFullYear() ||
    (
      selectedDate.getFullYear() === today.getFullYear() &&
      selectedDate.getMonth() < today.getMonth()
    );

  const monthIsFuture =
    selectedDate.getFullYear() > today.getFullYear() ||
    (
      selectedDate.getFullYear() === today.getFullYear() &&
      selectedDate.getMonth() > today.getMonth()
    );

  let msg = "";

  if (monthIsFuture) {
    msg = pickRandom(futureMonthMessages);
  } else if (monthIsPast) {
    if (pct >= 1)
      msg = pickRandom(exceededGoalMessages);
    else if (pct >= 0.9)
      msg = "you finished close to goal. Carry that momentum into the next month.";
    else if (pct >= 0.7)
      msg = "you made progress, but the month finished short of goal. Use it as fuel for the next one.";
    else
      msg = "that month is behind you. Reset, refocus, and attack the next opportunity.";
  } else {
    if (pct >= 1)
      msg = pickRandom(exceededGoalMessages);
    else if (pct >= 0.9)
      msg = pickRandom(onTrackMessages);
    else if (pct >= 0.7)
      msg = pickRandom(closeMessages);
    else
      msg = pickRandom(needsPushMessages);
  }

  return `${firstName}, ${msg.charAt(0).toLowerCase()}${msg.slice(1)}`;
})();
const quarterCrushingMessages = [
"You're projected to finish the quarter well above goal bub. Keep your foot on the gas!",
"Your quarterly pace is exceptional and currently exceeds expectations. Well done!",
"You're creating strong momentum that projects a successful quarter. Keep it up!",
"If this pace continues, you'll comfortably surpass your quarterly target. Keep the drive alive!",
"Outstanding work bub. The quarter is shaping up extremely well. That's fire.",
"You're operating at an elite pace and setting the standard for the quarter.",
"Your current trajectory puts you in position for a remarkable finish bub, keep it up!",
"The consistency you've shown is creating outstanding quarterly results.",
"You're well ahead of target and proving what's possible with disciplined execution.",
"Every indicator suggests you're on track for an exceptional quarter. One team one dream!",
"Your pace continues to outperform expectations across the board. You're killing it, bub!",
"You're building a quarter that others will benchmark against. Keep it up!",
"Keep attacking opportunities bub. The results are speaking for themselves!",
"You're turning your leads into elite-level production. That's fire!",
];

const quarterGoalAchievedMessages = [
"Quarter goal achieved bub. Great job! Now let's see how far above it we can finish.",
"You've already reached your quarterly target. That's fire!",
"Outstanding work. The goal is behind you and now it's time to build separation.",
"Goal accomplished bub, congrats! Keep your foot on the gas and keep driving.",
"You're officially over quarterly goal bub. Keep stacking wins.",
"One team one dream! You've already hit the target and there's still runway left.",
"Your hard work has paid off. Quarter goal achieved and momentum is still building.",
"You've earned the right to celebrate for about five minutes. Then let's go get more.",
"The goal is complete. Everything from here is bonus production bub.",
"That's what winning looks like. Quarterly goal achieved and still climbing.",
"You didn't just pace toward the goal — you got there. Keep pushing.",
"Quarter goal secured. Now let's see what kind of number we can really put up.",
"You're proving what's possible through consistency and execution. That's fire!",
"Goal reached bub. Stay aggressive and keep attacking opportunities.",
"The quarter is already a success. Now let's make it an exceptional one."
];

const quarterCloseMessages = [
"The quarter remains within reach. A strong finish can make all the difference.",
"You're close just a few more contracts could put you over goal.",
"Keep your momentum and focus over the next appointments and great things will happen.",
"The quarterly target remains achievable. Keep the grind alive.",
"Keep pushing bub. The pace is right where it needs to be.",
"You're within range of the goal and still have time to close the gap. Keep pushing!",
"A strong finish is what you need to close out the quarter. You got this!",
"Momentum over the next few weeks will be critical. Keep focused.",
"The quarter remains highly achievable from this position. Keep pushing bub!",
"Stay focused on your leads and the quarter is yours - You got this!",
"You're not far from where you need to be. Keep pushing, you got this!",
"A few more contracts can turn around your quarter. Keep up the hard work!",
"The opportunity to finish strong remains very real, keep the grind alive!",
"Stay disciplined and keep pushing forward. You are close to your quarterly goal!",
"Keep applying pressure, your quarterly goal is in sight!"
];

const quarterNeedsPushMessages = [
"The quarter needs additional momentum, but there is still time. Keep pushing.",
"Focus on pipeline activity and creating opportunities.",
"Several strong contracts can quickly improve the quarterly outlook.",
"Stay disciplined and continue executing the fundamentals.",
"The quarter isn't decided yet. Keep building momentum.",
"There's still time to improve the outcome of the quarter.",
"Focus on your leads and don't let off the gas - you got this!",
"Every conversation creates a new opportunity. Keep grinding.",
"Momentum can shift quickly when activity increases.",
"Continue building the sales pipeline and creating opportunities.",
"Stay engaged bub and keep attacking the next opportunity.",
"Persistence now can create a much stronger finish later - you got this!",
"One productive stretch can change the entire quarter.",
"Keep your energy high and focus on execution. You can do this!",
"The quarter is still being written. Keep pushing forward."
];

const futureQuarterMessages = [
  "We're not in this quarter just yet bub, but I know you'll be ready when it gets here.",
  "This quarter hasn't started yet bub, but preparation starts now.",
  "The scoreboard is still blank for this quarter bub, but the opportunity is coming.",
  "Future quarters are won by the habits you build today bub.",
  "We're not there yet bub, but I like our chances when the time comes.",
  "The quarter is still ahead of us bub. Stay sharp and be ready to attack it.",
  "No numbers just yet bub, just opportunity. Be ready when the quarter opens.",
  "This quarter is still waiting on us, bub. Let's be ready to make it count."
];

const quarterlyMotivation = (() => {
  const pct = pmData.quarterlyGoalPercent;
  const firstName = (currentPM?.name || "").split(" ")[0];

  const selectedQuarterStart = (() => {
    const selectedDate = new Date(pmData.selectedMonth);

    if (Number.isNaN(selectedDate.getTime())) return null;

    const fiscalYearStart =
      selectedDate.getMonth() >= 10
        ? new Date(selectedDate.getFullYear(), 10, 1)
        : new Date(selectedDate.getFullYear() - 1, 10, 1);

    const monthsSinceFiscalStart =
      (selectedDate.getFullYear() - fiscalYearStart.getFullYear()) * 12 +
      selectedDate.getMonth() -
      fiscalYearStart.getMonth();

    const quarterStartOffset = Math.floor(monthsSinceFiscalStart / 3) * 3;

    return new Date(
      fiscalYearStart.getFullYear(),
      fiscalYearStart.getMonth() + quarterStartOffset,
      1
    );
  })();

  const selectedQuarterEnd = selectedQuarterStart
    ? new Date(
        selectedQuarterStart.getFullYear(),
        selectedQuarterStart.getMonth() + 3,
        0
      )
    : null;

  const today = new Date();
  const quarterIsPast = selectedQuarterEnd ? selectedQuarterEnd < today : false;
  const quarterIsFuture = selectedQuarterStart ? selectedQuarterStart > today : false;

  let msg = "";

  if (quarterIsFuture) {
    msg = pickRandom(futureQuarterMessages);
  } else if (quarterIsPast) {
    if (pct >= 1)
      msg =
        "you exceeded your quarterly goal. Outstanding work and a strong finish to the quarter. That's fire.";
    else if (pct >= 0.9)
      msg =
        "you finished just short of your quarterly goal. Carry that momentum into the next quarter.";
    else if (pct >= 0.7)
      msg =
        "you made solid progress during the quarter. Use those lessons to build an even stronger next quarter.";
    else
      msg =
        "that quarter is complete. Take what you learned, reset, and attack the next opportunity.";
  } else {
    if (pct >= 1.2)
      msg = pickRandom(quarterCrushingMessages);
    else if (pct >= 1.0)
      msg = pickRandom(quarterGoalAchievedMessages);
    else if (pct >= 0.85)
      msg = pickRandom(quarterCloseMessages);
    else
      msg = pickRandom(quarterNeedsPushMessages);
  }

  return `${firstName}, ${msg.charAt(0).toLowerCase()}${msg.slice(1)}`;
})();

const referralGoalMetMessages = [
  "the referral goal was met bub! Great work creating opportunities through relationships!",
  "you met the referral goal! That means people trust you enough to send opportunities your way.",
  "the referral goal was achieved!  Relationship-driven business is working strong. Let's go!",
  "you reached the referral target bub. You're doing a great job!",
  "the goal was met on referrals. That kind of trust is earned not given, and that's fire."
];

const referralGoalNotMetMessages = [
  "the referral goal not met for this period. Review completed projects and look for customers who may still be willing to refer.",
  "the referral goal was not met. Use this as a reminder to make referral asks part of every closeout conversation.",
  "the goal was not met on referrals. The opportunity now is to follow up with happy customers and ask who else we can help.",
  "the referral target was missed. A few intentional referral conversations can help turn that around next month.",
  "the referral goal was not met. Reset, refocus, and keep asking satisfied customers for introductions."
];

const referralInProgressMessages = [
  "the referral goal has not been reached yet. Keep asking happy customers who else you can help.",
  "there is still room to build more referral opportunities this month - you got this.",
  "referrals need a little more focus. Every satisfied customer can open another door.",
  "keep planting referral seeds. The best leads often come from people who already trust us.",
  "the referral activity is behind goal, but a few intentional asks can close the gap quickly."
];

const referralFutureMessages = [
  "this referral month has not started yet, but the best referral opportunities are built before the scoreboard opens.",
  "there is a future referral goal ahead. Keep creating great customer experiences now and the asks will feel natural later.",
  "we are not there yet bub, but every satisfied customer can become a future referral source.",
  "a future month is selected. Build the relationship now, ask when the timing is right, and the referrals will follow.",
  "that referral window is still ahead bub. Keep doing the kind of work people want to recommend."
];

const referralStatusLabel = (() => {
  if (pmData.referralStatus === "goalMet") return "Goal Met";
  if (pmData.referralStatus === "goalNotMet") return "Goal Not Met";
  if (pmData.referralStatus === "future") return "Upcoming";
  return "In Progress";
})();

const referralStatusClass =
  pmData.referralStatus === "goalMet"
    ? "positive"
    : pmData.referralStatus === "goalNotMet"
    ? "negative"
    : "warning";

const referralMotivation = (() => {
  const firstName = (currentPM?.name || "").split(" ")[0];

  const messages =
    pmData.referralStatus === "goalMet"
      ? referralGoalMetMessages
      : pmData.referralStatus === "goalNotMet"
      ? referralGoalNotMetMessages
      : pmData.referralStatus === "future"
      ? referralFutureMessages
      : referralInProgressMessages;

  const msg = pickRandom(messages);

return `${firstName}, ${msg.charAt(0).toLowerCase()}${msg.slice(1)}`;
})();

const quarterlyReferralGoalMetMessages = [
  "quarterly referral goal achieved. Consistent relationship building is paying off.",
  "you hit the quarterly referral target. That's the result of trust earned over time.",
  "quarterly referral production is exactly where it needs to be. Great work.",
  "the quarter referral goal has been met. Keep stacking wins through relationships.",
  "strong quarter on referrals. Word-of-mouth momentum is growing."
];

const quarterlyReferralGoalNotMetMessages = [
  "quarterly referral goal was missed. Review every completed project and identify referral opportunities.",
  "the quarter is behind target. A stronger referral process can close that gap next quarter.",
  "quarterly referrals fell short. Focus on asking every satisfied customer for introductions.",
  "the referral target for the quarter was not reached. Consistency in referral conversations matters.",
  "this quarter missed the referral mark. Build a plan to create more referral opportunities."
];

const quarterlyReferralInProgressMessages = [
  "the quarter is still underway. Every referral conversation matters.",
  "quarterly referral production is building. Stay consistent with referral asks.",
  "there is time left in the quarter to create referral momentum.",
  "quarterly referral progress is developing. Keep relationship building at the forefront.",
  "the quarter is not finished yet. Small referral wins compound over time."
];

const quarterlyReferralFutureMessages = [
  "this quarter has not started yet bub. Prepare now so referrals come naturally later.",
  "future quarter selected. Build relationships today that generate referrals tomorrow.",
  "the next quarter is ahead. Strong customer experiences create future opportunities.",
  "that referral quarter is still coming. Stay focused on earning trust now.",
  "future quarter selected. Set the foundation before the scoreboard opens."
];

const quarterlyReferralStatusClass =
  pmData.quarterlyReferralStatus === "goalMet"
    ? "positive"
    : pmData.quarterlyReferralStatus === "goalNotMet"
    ? "negative"
    : "warning";

const quarterlyReferralMotivation = (() => {
  const firstName = (currentPM?.name || "").split(" ")[0];

  const messages =
    pmData.quarterlyReferralStatus === "goalMet"
      ? quarterlyReferralGoalMetMessages
      : pmData.quarterlyReferralStatus === "goalNotMet"
      ? quarterlyReferralGoalNotMetMessages
      : pmData.quarterlyReferralStatus === "future"
      ? quarterlyReferralFutureMessages
      : quarterlyReferralInProgressMessages;

  const msg = pickRandom(messages);

  return `${firstName}, ${msg}`;
})();
const generatePMInsight = ({
  monthlyGoalPercent,
  quarterlyGoalPercent,
  quarterlyProjectedPercent,
  revenueVsTeam,
  averageVsTeam,
  closingRateVsTeam,
  quarterRevenueVsTeam,
  quarterContractsVsTeam,
  quarterClosingRateVsTeam,
  isPastMonth,
  isFutureMonth,
  isCustomMode,
}) => {
  const percentPhrase = (value) =>
    `${Math.abs(value * 100).toFixed(1)}% ${
      value >= 0 ? "above" : "below"
    }`;

  const pointPhrase = (value) =>
    `${Math.abs(value * 100).toFixed(1)} points ${
      value >= 0 ? "above" : "below"
    }`;

  if (isCustomMode) {
    return {
      monthly:
        "Custom date range selected. Performance insight is based on production and team comparison for the selected dates.",
      quarterly: null,
    };
  }

  if (isFutureMonth) {
    return {
      monthly:
        "Future month selected. The scoreboard is still blank, but preparation now creates production later.",
      quarterly:
        "Quarter-to-date performance will update once production begins for the selected period.",
    };
  }

  const monthlyLead = isPastMonth
    ? monthlyGoalPercent >= 1
      ? "Monthly goal closed above target."
      : "Monthly goal closed below target."
    : monthlyGoalPercent >= 1
    ? "Monthly goal has been achieved and production remains strong."
    : monthlyGoalPercent >= 0.9
    ? "Monthly goal is within striking distance and remains well within reach."
    : "Monthly production is still building toward goal pace.";

  const quarterlyLead =
    quarterlyGoalPercent >= 1
      ? "Quarterly goal has already been achieved."
      : quarterlyProjectedPercent >= 1
      ? "Quarter-to-date production is pacing toward the quarterly goal."
      : quarterlyProjectedPercent >= 0.85
      ? "Quarter-to-date production remains within reach with a strong finish."
      : "Quarter-to-date production needs additional momentum to reach goal pace.";

  return {
    monthly: `${monthlyLead} Revenue production is ${percentPhrase(
      revenueVsTeam
    )} the team average this month. Average contract value is ${percentPhrase(
      averageVsTeam
    )} the team benchmark, and closing rate is ${pointPhrase(
      closingRateVsTeam
    )} the team average.`,

    quarterly: `${quarterlyLead} Quarter-to-date revenue production is ${percentPhrase(
      quarterRevenueVsTeam
    )} the team average, while contract volume is ${percentPhrase(
      quarterContractsVsTeam
    )} team benchmark levels. Quarter-to-date closing rate is ${pointPhrase(
      quarterClosingRateVsTeam
    )} the team average.`,
  }; 
};

const isCustomMode = pmDateMode === "custom";
const selectedMonthDate = new Date(pmData.selectedMonth);
const now = new Date();

const isPastMonth =
  selectedMonthDate.getFullYear() < now.getFullYear() ||
  (selectedMonthDate.getFullYear() === now.getFullYear() &&
    selectedMonthDate.getMonth() < now.getMonth());

const isFutureMonth =
  selectedMonthDate.getFullYear() > now.getFullYear() ||
  (selectedMonthDate.getFullYear() === now.getFullYear() &&
    selectedMonthDate.getMonth() > now.getMonth());

const pmInsight = generatePMInsight({
  monthlyGoalPercent: pmData.monthlyGoalPercent || 0,
  quarterlyGoalPercent: pmData.quarterlyGoalPercent || 0,
  quarterlyProjectedPercent: pmData.quarterlyProjectedPercent || 0,
  revenueVsTeam: revenueVsTeam?.rawDifference || 0,
  averageVsTeam: averageVsTeam?.rawDifference || 0,
  closingRateVsTeam: closingRateVsTeam?.rawDifference || 0,
  quarterRevenueVsTeam: quarterRevenueVsTeam?.rawDifference || 0,
  quarterContractsVsTeam: quarterContractsVsTeam?.rawDifference || 0,
  quarterClosingRateVsTeam: quarterClosingRateVsTeam?.rawDifference || 0,
  isPastMonth,
  isFutureMonth,
  isCustomMode,
});

const monthlyGoalClass =
  pmData.monthlyGoalPercent >= 1
    ? "positive"
    : isPastMonth
    ? "negative"
    : "warning";

const selectedQuarterStartForClass = (() => {
  if (!selectedMonthDate || Number.isNaN(selectedMonthDate.getTime())) return null;

  const fiscalYearStart =
    selectedMonthDate.getMonth() >= 10
      ? new Date(selectedMonthDate.getFullYear(), 10, 1)
      : new Date(selectedMonthDate.getFullYear() - 1, 10, 1);

  const monthsSinceFiscalStart =
    (selectedMonthDate.getFullYear() - fiscalYearStart.getFullYear()) * 12 +
    selectedMonthDate.getMonth() -
    fiscalYearStart.getMonth();

  const quarterStartOffset = Math.floor(monthsSinceFiscalStart / 3) * 3;

  return new Date(
    fiscalYearStart.getFullYear(),
    fiscalYearStart.getMonth() + quarterStartOffset,
    1
  );
})();

const selectedQuarterEndForClass = selectedQuarterStartForClass
  ? new Date(
      selectedQuarterStartForClass.getFullYear(),
      selectedQuarterStartForClass.getMonth() + 3,
      0
    )
  : null;

const quarterComplete = selectedQuarterEndForClass
  ? selectedQuarterEndForClass < now
  : false;

const quarterlyGoalClass =
  pmData.quarterlyGoalPercent >= 1
    ? "positive"
    : quarterComplete
    ? "negative"
    : "warning";
    return (
      <div className="page pm-page">
        <header className="header">
          <button className="header-action-button" onClick={pmLogout}>
            Logout
          </button>

<div className="header-top">

<img
  src="/logo.png"
  alt="Logo"
  className="logo"
/>

  <div className="pm-header-center">

    <h1>{currentPM.name} Dashboard</h1>

    <div className="pm-insight-box">

     <div className="pm-insight-copy">
  <div>
  <strong>Month-to-Date Performance Insight</strong>
  <p>{pmInsight.monthly}</p>
</div>

{pmInsight.quarterly && (
  <>
    <div className="pm-insight-divider" />

    <div>
      <strong>Quarter-to-Date Performance Insight</strong>
      <p>{pmInsight.quarterly}</p>
    </div>
  </>
)}
</div>
    </div>

  </div>

</div>
        </header>

        <section className="pm-dashboard">
          <div className="pm-hero-card">
            <img src={currentPM.image} alt={(currentPM?.name || projectManagers[0].name)} />

<div className="pm-hero-content">
  <span>Project Manager</span>
  <h2>{currentPM?.name || projectManagers[0].name}</h2>

  {!isCustomMode && (
<div className="pm-rank-row">
  <div className={`pm-rank-card ${isRevenueLeader ? "rank-leader" : ""}`}>
    <small>Revenue Rank</small>
    <strong>{revenueRankLabel}</strong>

    {isRevenueLeader && (
      <div className="top-dawg-label">
        <b>TOP DAWG</b>
      </div>
    )}
  </div>

  <div className={`pm-rank-card ${isClosingLeader ? "rank-leader" : ""}`}>
    <small>Closing Rate Rank</small>
    <strong>{closingRankLabel}</strong>

    {isClosingLeader && (
      <div className="top-dawg-label">
        TOP DAWG
      </div>
    )}
  </div>
</div>
  )}
</div>

<div className="pm-date-controls">
  <label className="pm-month-selector">
    Month
    <select
      value={pmData.selectedMonth}
      onChange={(event) => {
        setSelectedPMMonth(event.target.value);
        setPmDateMode("month");
      }}
    >
      {pmData.monthOptions.map((month) => (
        <option key={month} value={month}>
          {month}
        </option>
      ))}
    </select>
  </label>

  <div className="pm-custom-date-row">
    <label>
      Start Date
      <input
        type="date"
        value={pmStartDate}
        onChange={(event) => {
          setPmStartDate(event.target.value);
          setPmDateMode("custom");
        }}
      />
    </label>

    <label>
      End Date
      <input
        type="date"
        value={pmEndDate}
        onChange={(event) => {
          setPmEndDate(event.target.value);
          setPmDateMode("custom");
        }}
      />
    </label>
  </div>

<button
  type="button"
  className="pm-fiscal-button"
  onClick={() => {
    setPmDateMode("fiscalYTD");
    setPmStartDate("2025-11-01");
    setPmEndDate(new Date().toISOString().slice(0, 10));
  }}
>
  Fiscal YTD
</button>
</div>
          </div>

          <div className="pm-section-card">
            <h2>
  {pmDateMode === "fiscalYTD"
    ? "Fiscal YTD Performance"
    : pmDateMode === "custom"
    ? "Custom Date Performance"
    : `${pmData.selectedMonth} Performance`}
</h2>

            <div className="pm-metric-grid">
<PMMetricCard
  label="Contract Total"
  value={money(pmData.contractTotal)}
  comparisonLabel={
    pmDateMode === "fiscalYTD"
      ? "vs 11/1/2024 - 6/19/2025"
      : isCustomMode
      ? null
      : `vs ${pmData.lastYearMonth}`
  }
  comparisonValue={isCustomMode ? null : money(pmData.lyContractTotal)}
  difference={isCustomMode ? null : revenueVsLY}
/>

<PMMetricCard
  label="Contracts"
  value={Math.round(pmData.contracts || 0)}
  comparisonLabel={
    pmDateMode === "fiscalYTD"
      ? "vs 11/1/2024 - 6/19/2025"
      : isCustomMode
      ? null
      : `vs ${pmData.lastYearMonth}`
  }
  comparisonValue={isCustomMode ? null : Math.round(pmData.lyContracts || 0)}
  difference={isCustomMode ? null : contractsVsLY}
/>

<PMMetricCard
  label="Average Contract"
  value={money(pmData.averageContract)}
  comparisonLabel={
    pmDateMode === "fiscalYTD"
      ? "vs 11/1/2024 - 6/19/2025"
      : isCustomMode
      ? null
      : `vs ${pmData.lastYearMonth}`
  }
  comparisonValue={isCustomMode ? null : money(pmData.lyAverageContract)}
  difference={isCustomMode ? null : averageVsLY}
/>

<PMMetricCard
  label="Closing Rate"
  value={displayPercent(pmData.closingRate, 1)}
  comparisonLabel={
    pmDateMode === "fiscalYTD"
      ? "vs 11/1/2024 - 6/19/2025"
      : isCustomMode
      ? null
      : `vs ${pmData.lastYearMonth}`
  }
  comparisonValue={isCustomMode ? null : displayPercent(pmData.lyClosingRate, 1)}
  difference={isCustomMode ? null : closingRateVsLY}
/>
            </div>
          </div>
          
{pmDateMode !== "fiscalYTD" && !isCustomMode && (
  <>
    <div className="pm-section-card">
      <h2>Month-to-Date Team Comparison</h2>

      <div className="pm-metric-grid">
        <PMMetricCard
          label="Revenue vs Team"
          value={money(pmData.contractTotal)}
          comparisonLabel="Team Avg"
          comparisonValue={money(pmData.teamContractTotal)}
          difference={revenueVsTeam}
        />

        <PMMetricCard
          label="Contracts vs Team"
          value={Math.round(pmData.contracts || 0)}
          comparisonLabel="Team Avg"
          comparisonValue={Math.round(pmData.teamContracts || 0)}
          difference={contractsVsTeam}
        />

        <PMMetricCard
          label="Avg Contract vs Team"
          value={money(pmData.averageContract)}
          comparisonLabel="Team Avg"
          comparisonValue={money(pmData.teamAverageContract)}
          difference={averageVsTeam}
        />

        <PMMetricCard
          label="Closing Rate vs Team"
          value={displayPercent(pmData.closingRate, 1)}
          comparisonLabel="Team Avg"
          comparisonValue={displayPercent(pmData.teamClosingRate, 1)}
          difference={closingRateVsTeam}
        />
      </div>
    </div>

    <div className="pm-section-card">
      <h2>Rolling 90-Day Team Comparison</h2>

      <div className="pm-metric-grid">
        <PMMetricCard
          label="Revenue vs Team"
          value={money(pmData.ninetyDayContractTotal)}
          comparisonLabel="Team Avg"
          comparisonValue={money(pmData.ninetyDayTeamContractTotal)}
          difference={compareNumbers(
            pmData.ninetyDayContractTotal,
            pmData.ninetyDayTeamContractTotal
          )}
        />

        <PMMetricCard
          label="Contracts vs Team"
          value={Math.round(pmData.ninetyDayContracts || 0)}
          comparisonLabel="Team Avg"
          comparisonValue={Math.round(pmData.ninetyDayTeamContracts || 0)}
          difference={compareNumbers(
            pmData.ninetyDayContracts,
            pmData.ninetyDayTeamContracts
          )}
        />

        <PMMetricCard
          label="Avg Contract vs Team"
          value={money(pmData.ninetyDayAverageContract)}
          comparisonLabel="Team Avg"
          comparisonValue={money(pmData.ninetyDayTeamAverageContract)}
          difference={compareNumbers(
            pmData.ninetyDayAverageContract,
            pmData.ninetyDayTeamAverageContract
          )}
        />

        <PMMetricCard
          label="90-Day Closing Rate vs Team"
          value={displayPercent(pmData.ninetyDayClosingRate, 1)}
          comparisonLabel="Team Avg"
          comparisonValue={displayPercent(pmData.ninetyDayTeamClosingRate, 1)}
          difference={compareNumbers(
            pmData.ninetyDayClosingRate,
            pmData.ninetyDayTeamClosingRate,
            true
          )}
        />
      </div>
    </div>
  </>
)}
          {!isCustomMode && (
            <>
              <div className="pm-goal-card">
            <div className="pm-goal-top">
              <span>YTD Revenue Goal Progress - November 1, 2025-October 31, 2026</span>
              <strong>{money(pmData.ytdRevenue)}</strong>
              <p>Goal: {money(pmData.individualGoal)}</p>
            </div>

            <div className="pm-thermometer">
              <div
                className="pm-thermometer-fill"
                style={{
                  width: `${Math.min(pmData.goalPercent * 100, 100)}%`,
                }}
              />
            </div>
            

            <div className="pm-goal-stats">
              <div>
                <span>Complete</span>
                <strong>{displayPercent(pmData.goalPercent, 1)}</strong>
              </div>

              <div>
                <span>Remaining</span>
                <strong>{money(pmData.remainingToGoal)}</strong>
              </div>

              <div>
                <span>YTD Contracts</span>
                <strong>{Math.round(pmData.ytdContracts || 0)}</strong>
              </div>

              <div>
                <span>YTD Avg Contract</span>
                <strong>{money(pmData.ytdAverageContract)}</strong>
              </div>
            </div>
          </div>
<div className="pm-section-card">
  <h2>Goal Pace Breakdown</h2>

  <div className="pm-metric-grid">
<PMMetricCard
  label="Monthly Goal"
  value={money(pmData.monthlyGoal)}
  comparisonLabel="Monthly Actual"
  comparisonValue={money(pmData.contractTotal)}
  difference={{
    label: `${displayPercent(pmData.monthlyGoalPercent, 1)} Complete`,
    className: monthlyGoalClass,
  }}
/>


<PMMetricCard
  label={
    pmData.monthlyRemaining >= 0
      ? "Monthly Remaining"
      : "Over Goal"
  }
  value={money(Math.abs(pmData.monthlyRemaining))}
  customMessage={goalMotivation}
  messageTitle="✨MAGIC MIKE MOMENT✨"
/> 
<PMMetricCard
  label="Quarterly Goal"
  value={money(pmData.quarterlyGoal)}
  comparisonLabel="Quarterly Actual"
  comparisonValue={money(pmData.quarterRevenue)}
  difference={{
    label: `${displayPercent(pmData.quarterlyGoalPercent, 1)} Complete`,
    className: quarterlyGoalClass,
  }}
/>

<PMMetricCard
  label="Quarterly Remaining"
  value={money(pmData.quarterlyRemaining)}
  comparisonLabel={`${pmData.quarterDaysRemaining} days quarterly remaining`}
  comparisonValue=""
  customMessage={quarterlyMotivation}
/>
  </div>
</div>


<div className="pm-section-card">
  <h2>Referral Information</h2>

  <div className="pm-metric-grid">
    <PMMetricCard
      label="Referral Goal"
      value={Math.round(pmData.referralGoal || 0)}
      comparisonLabel="Referral Actual"
      comparisonValue={Math.round(pmData.referralTotal || 0)}
      difference={{
        label: `${displayPercent(pmData.referralPercent, 1)} Complete`,
        className: referralStatusClass,
      }}
    />

    <PMMetricCard
      label={
        pmData.referralDelta >= 0
          ? "Over Referral Goal"
          : "Referral Remaining"
      }
      value={`${pmData.referralDelta >= 0 ? "+" : ""}${Math.round(
        pmData.referralDelta || 0
      )}`}
      customMessage={referralMotivation}
      messageTitle="✨MAGIC MIKE MOMENT✨"
    />

    <PMMetricCard
      label="Quarterly Goal"
      value={Math.round(pmData.quarterlyReferralGoal || 0)}
      comparisonLabel="Quarterly Actual"
      comparisonValue={Math.round(pmData.quarterlyReferralTotal || 0)}
      difference={{
        label: `${displayPercent(pmData.quarterlyReferralPercent, 1)} Complete`,
        className: quarterlyReferralStatusClass,
      }}
    />

    <PMMetricCard
      label={
        pmData.quarterlyReferralDelta >= 0
          ? "Over Quarterly Goal"
          : "Quarterly Remaining"
      }
      value={`${pmData.quarterlyReferralDelta >= 0 ? "+" : ""}${Math.round(
        pmData.quarterlyReferralDelta || 0
      )}`}
      customMessage={quarterlyReferralMotivation}
      messageTitle="✨MAGIC MIKE MOMENT✨"
    />
  </div>
</div>
            </>
          )}

          <div className="pm-commission-card">
            <h2>Commission Calculator</h2>

            <div
              className="pm-sale-input-grid"
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: "16px",
                marginBottom: "18px",
              }}
            >
              <div className="pm-sale-input">
                <label>Potential Sale Amount</label>
                <input
                  type="text"
                  inputMode="decimal"
                  value={formatMoneyInput(pmSaleAmount)}
                  onChange={(event) =>
                    setPmSaleAmount(cleanMoneyInput(event.target.value))
                  }
                />
              </div>

              <div className="pm-sale-input">
                <label>Job Cost (Labor and Materials)</label>
                <input
                  type="text"
                  inputMode="decimal"
                  value={formatMoneyInput(pmJobCost)}
                  onChange={(event) =>
                    setPmJobCost(cleanMoneyInput(event.target.value))
                  }
                />
              </div>
            </div>


<div className="pm-commission-summary">
  <div>
    <span>Commission Rate</span>
    <strong>{displayPercent(pmData.commissionRate, 0)}</strong>
  </div>

<div
  ref={commissionCardRef}
  className={`pm-commission-reward ${
    Number(cleanMoneyInput(pmJobCost)) > 0 ? "active" : ""
  }`}
>
  <span>Your Estimated Commission</span>
  <strong>{money(pmData.commission)}</strong>
</div>
</div>
          </div>

          <div className="pm-section-card top-dawg-leaderboard-card">
            <h2>Top Dawg Hall of Fame</h2>

            <p className="top-dawg-leaderboard-note">
              Month-end winners only. Current open month is excluded.
            </p>

            <div className="top-dawg-leaderboard-grid">
              <div className="top-dawg-board">
                <h3>Revenue Top Dawgs</h3>

                {pmData.topDawgLeaderboards?.revenue?.length ? (
                  pmData.topDawgLeaderboards.revenue.map((item) => (
                    <div className="top-dawg-board-row" key={`revenue-${item.month}`}>
                      <span>{item.month}</span>
                      <strong>{renderTopDawgNames(item.names)}</strong>
                      <b>{money(item.value)}</b>
                    </div>
                  ))
                ) : (
                  <p className="top-dawg-empty">No closed-month revenue winners yet.</p>
                )}
              </div>

              <div className="top-dawg-board">
                <h3>Closing Rate Top Dawgs</h3>

                {pmData.topDawgLeaderboards?.closingRate?.length ? (
                  pmData.topDawgLeaderboards.closingRate.map((item) => (
                    <div className="top-dawg-board-row" key={`closing-${item.month}`}>
                      <span>{item.month}</span>
                      <strong>{renderTopDawgNames(item.names)}</strong>
                      <b>{displayPercent(item.value, 1)}</b>
                    </div>
                  ))
                ) : (
                  <p className="top-dawg-empty">No closed-month closing-rate winners yet.</p>
                )}
              </div>
            </div>
          </div>
                    <div className="pm-section-card best-of-best-card">
            <h2>All Time GOATs</h2>

            <div className="best-of-best-grid">
              <div className="best-of-best-item">
  <span>Revenue GOAT</span>

  {pmData.bestEverTopDawgs?.revenue ? (
    <>
      {revenueGoatPM && (
        <img
          className="goat-photo"
          src={revenueGoatPM.image}
          alt={revenueGoatPM.name}
        />
      )}

<strong>
  {pmData.bestEverTopDawgs.revenue.names.join(", ")}
</strong>

      <b>{money(pmData.bestEverTopDawgs.revenue.value)}</b>

      <small>{pmData.bestEverTopDawgs.revenue.month}</small>
    </>
  ) : (
    <p>No revenue winner yet.</p>
  )}
</div>

              <div className="best-of-best-item">
  <span>Closing Rate GOAT</span>

  {pmData.bestEverTopDawgs?.closingRate ? (
    <>
      {closingRateGoatPM && (
        <img
          className="goat-photo"
          src={closingRateGoatPM.image}
          alt={closingRateGoatPM.name}
        />
      )}

<strong>
  {pmData.bestEverTopDawgs.closingRate.names.join(", ")}
</strong>

      <b>
        {displayPercent(
          pmData.bestEverTopDawgs.closingRate.value,
          1
        )}
      </b>

      <small>
        {pmData.bestEverTopDawgs.closingRate.month}
      </small>
    </>
  ) : (
    <p>No closing-rate winner yet.</p>
  )}
</div>
            </div>
            </div>
        </section>
      </div>
    );
  }

  if (!authorized) {
    return (
      <div className="login-screen">
        <form className="login-box" onSubmit={handleLogin}>
          <img src="/logo.png" alt="Logo" className="login-logo" />

          <h1>Dashboard Login</h1>
          <p>Enter the password to access the forecast dashboard.</p>

          <input
            type="text"
            placeholder="Password"
            value={password}
            autoFocus
            onChange={(event) => {
              setPassword(event.target.value);
              setLoginError("");
            }}
          />

          {loginError && <div className="login-error">{loginError}</div>}

          <button type="submit">Login</button>
        </form>
      </div>
    );
  }

  if (screen === "home") {
    return (
      <div className="page">
        <Header />

        <section className="home-grid">
          <button className="home-card" onClick={() => setScreen("leads")}>
            <span>Lead Forecast Revenue & Profit Margin</span>
            <strong>Forecast Revenue from Lead Counts</strong>
            <p>Uses Average Revenue Per Project × Margin x Close Rate.</p>
          </button>

          <button
            className="home-card projects-card"
            onClick={() => setScreen("projects")}
          >
            <span>Project Revenue & Profit Margin</span>
            <strong>Revenue & Profit from Historical Sales</strong>
            <p>Uses Historical Sales Data from Acculynx Sales Report</p>
          </button>

          <button
            className="home-card marketing-card"
            onClick={() => setScreen("marketing")}
          >
            <span>Marketing Forecast</span>
            <strong>Predict Revenue from Marketing Spend</strong>
            <p>
              Enter total spend, apply historical channel split, then override by
              channel.
            </p>
          </button>
        </section>
      </div>
    );
  }

  return (
    <div className="page">
      <Header />

      {screen === "marketing" && (
        <section className="calculator-section">
          <div className="marketing-total-spend-panel">
            <label>
              Total Marketing Spend
              <input
  type="text"
  inputMode="decimal"
  value={formatMoneyInput(totalMarketingSpend)}
  onChange={(event) =>
    setTotalMarketingSpend(cleanMoneyInput(event.target.value))
  }
/>
            </label>

            <button onClick={applyHistoricalSpendSplit}>
              Update Numbers
            </button>

          </div>
<div className="marketing-total-section">
            <h2>Total Marketing Forecast</h2>

            <div className="marketing-scenario-grid">
              {marketingForecast.totals.map((scenario) => (
                <div
                  className={`marketing-scenario-card ${scenario.key}`}
                  key={scenario.key}
                >
                  <h3>{scenario.label}</h3>

                  <div className="scenario-main-number">
                    <span>Forecast Revenue</span>
                    <strong>{money(scenario.revenue)}</strong>
                  </div>

                  <div className="scenario-row">
                    <span>Total Spend</span>
                    <strong>{money(scenario.spend)}</strong>
                  </div>

                  <div className="scenario-row">
                    <span>Leads</span>
                    <strong>{Math.round(scenario.leads)}</strong>
                  </div>

                  <div className="scenario-row">
                    <span>Appointments</span>
                    <strong>{Math.round(scenario.appointments)}</strong>
                  </div>

                  <div className="scenario-row">
                    <span>Return / Spend</span>
                    <strong>{scenario.returnPerSpend.toFixed(2)}x</strong>
                  </div>

                  <div className="scenario-row true">
                    <span>Project Profit</span>
                    <strong>{money(scenario.trueProjectProfit)}</strong>
                  </div>

                  <div className="scenario-row company">
                    <span>Company Profit</span>
                    <strong>{money(scenario.companyProfit)}</strong>
                  </div>
                </div>
              ))}
            </div>
          </div>
          <div className="marketing-channel-grid">
            {marketingForecast.channelForecasts.map((channel) => (
              <div className="marketing-channel-card" key={channel.key}>
                <h3>{channel.label}</h3>

                <label>
                  Forecast Spend
                  <input
  type="text"
  inputMode="decimal"
  value={formatMoneyInput(marketingSpendByChannel[channel.key])}
onChange={(event) => {
  const nextValue = cleanMoneyInput(event.target.value);

  setMarketingSpendByChannel((current) => {
    const updated = {
      ...current,
      [channel.key]: nextValue,
    };

    const newTotal = Object.values(updated).reduce(
      (sum, value) => sum + Number(value || 0),
      0
    );

    setTotalMarketingSpend(String(newTotal.toFixed(2)));

    return updated;
  });
}}
/>
                </label>

                <div className="scenario-row">
                  <span>Historical Spend Split</span>
                  <strong>{(channel.splitPercent * 100).toFixed(2)}%</strong>
                </div>

                <div className="scenario-row">
                  <span>Avg Revenue / Appointment</span>
                  <strong>{money(channel.averageRevenuePerAppointment)}</strong>
                </div>

                <div className="scenario-row">
                  <span>Appointments / Spend</span>
                  <strong>{channel.appointmentsPerSpend.toFixed(6)}</strong>
                </div>

                <div className="scenario-row">
                  <span>Cost / Lead</span>
                  <strong>
                    {channel.costPerLead ? money(channel.costPerLead) : "N/A"}
                  </strong>
                </div>

                <div className="mini-scenario-list">
                  {channel.scenarios.map((scenario) => (
                    <div
                      className={`mini-scenario ${scenario.key}`}
                      key={scenario.key}
                    >
                      <h4>{scenario.label}</h4>

                      <p>
                        Revenue: <strong>{money(scenario.revenue)}</strong>
                      </p>

                      <p>
                        Leads:{" "}
                        <strong>
                          {scenario.leads === null
                            ? "N/A"
                            : Math.round(scenario.leads)}
                        </strong>
                      </p>

                      <p>
                        Appointments:{" "}
                        <strong>{Math.round(scenario.appointments)}</strong>
                      </p>

                      <p>
                        ROAS:{" "}
                        <strong>{scenario.returnPerSpend.toFixed(2)}x</strong>
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>

          <div className="historical-marketing-section">
  <h2>Historical Spend & Revenue</h2>

  <div className="historical-marketing-grid">
    {historicalMarketingData.map((channel) => (
      <div className="historical-marketing-card" key={channel.channel}>
        <h3>{channel.channel}</h3>

        {channel.rows.map((row) => (
<>
{row.month === "Oct 2025" && (
  <div className="historical-header-row">
    <span>Month</span>
    <span>Spend</span>
    <span>Revenue</span>
  </div>
)}

  <div className="historical-row" key={row.month}>
    <span>{row.month}</span>

    <strong className="historical-spend">
      {money(row.spend)}
    </strong>

    <strong className="historical-revenue">
      {money(row.revenue)}
    </strong>
  </div>
</>
        ))}
      </div>
    ))}
  </div>
</div>
        </section>
      )}

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
              <button onClick={clearLeads}>Clear Leads</button>
              <button onClick={loadLeadFile}>Reload Leads File</button>
            </div>
          </div>

          <SummaryTotals
            revenue={leadTotals.revenue}
            trueProjectProfit={leadTotals.trueProjectProfit}
            companyProfit={leadTotals.companyProfit}
            companyExpense={leadTotals.companyExpense}
            revenueLabel="Lead Revenue"
          />

          <div className="lead-control-grid">
            <div className="unknown-leads-panel">
              <h3>Appointments - Unknown Work Type</h3>
              <p>Using {money(UNKNOWN_RPP)} avg revenue/project</p>
              <p>True Margin: {percent(UNKNOWN_MARGIN + COMPANY_EXPENSE_RATE)}</p>
              <p>Company Margin: {percent(UNKNOWN_MARGIN)}</p>
              <strong>{unknownLeadCount}</strong>
            </div>

            <div className="close-rate-panel">
              <h3>Appointment Close Rate</h3>
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

          <div className="lead-kpi-row">
            <div className="unknown-leads-panel lead-kpi-card">
              <h3>Total Appointments</h3>
              <strong>{leadTotals.totalLeads}</strong>
            </div>

            <div className="unknown-leads-panel lead-kpi-card">
              <h3>Current Prospects</h3>
              <strong>{currentProspectCount}</strong>
            </div>
          </div>

          <div className="grid">
            {categories.map((category) => (
              <div className={`card ${category.className}`} key={category.title}>
                <h3>{category.title}</h3>

                {[...category.items]
                  .sort(
                    (a, b) =>
                      WORK_TYPE_ORDER.indexOf(a.label) -
                      WORK_TYPE_ORDER.indexOf(b.label)
                  )
                  .map((item) => {
                    const key = `${category.title}-${item.label}`;
                    const quantity = Number(leads[key] || 0);
                    const revenue = quantity * activeCloseRate * item.rpp;
                    const breakdown = getProfitBreakdown(revenue, item.margin);

                    return (
                      <ForecastRow
                        key={key}
                        item={item}
                        keyName={key}
                        quantity={leads[key]}
                        revenue={revenue}
                        breakdown={breakdown}
                        type="leads"
                        onLeadQuantityChange={handleLeadQuantityChange}
                      />
                    );
                  })}

                <CategoryLeadTotal category={category} />
              </div>
            ))}
          </div>
        </section>
      )}



      {screen === "projects" && (
        <section className="calculator-section">
          <div className="project-top-row">
            <div className="data-panel project-date-panel">
              <div className="date-controls">
                <label>
                  Start Date
                  <input
                    type="date"
                    value={projectStartDate}
                    onChange={(event) => setProjectStartDate(event.target.value)}
                  />
                </label>

                <label>
                  End Date
                  <input
                    type="date"
                    value={projectEndDate}
                    onChange={(event) => setProjectEndDate(event.target.value)}
                  />
                </label>

                <button onClick={clearProjects}>Clear Projects</button>
                <button onClick={loadSalesFile}>Reload Sales File</button>
              </div>
            </div>

            <div className="unknown-leads-panel project-count-panel">
              <h3>Total Projects</h3>
              <strong>{projectTotals.projectCount}</strong>

              <p style={{ marginTop: "8px", fontSize: "14px", color: "#666" }}>
                {projectTotals.zeroProjectCount} projects at $0
              </p>
            </div>
          </div>

          <SummaryTotals
            revenue={projectTotals.revenue}
            trueProjectProfit={projectTotals.trueProjectProfit}
            companyProfit={projectTotals.companyProfit}
            companyExpense={projectTotals.companyExpense}
            revenueLabel="Project Revenue"
          />

          <div className="grid">
            {categories.map((category) => (
              <div className={`card ${category.className}`} key={category.title}>
                <h3>{category.title}</h3>

                {[...category.items]
                  .sort(
                    (a, b) =>
                      WORK_TYPE_ORDER.indexOf(a.label) -
                      WORK_TYPE_ORDER.indexOf(b.label)
                  )
                  .map((item) => {
                    const key = `${category.title}-${item.label}`;
                    const quantity = Number(projectData.projectCounts[key] || 0);
                    const revenue = Number(projectData.projectRevenue[key] || 0);
                    const breakdown = getProfitBreakdown(revenue, item.margin);

                    return (
                      <ForecastRow
                        key={key}
                        item={item}
                        keyName={key}
                        quantity={quantity}
                        revenue={revenue}
                        breakdown={breakdown}
                        type="projects"
                        onLeadQuantityChange={handleLeadQuantityChange}
                      />
                    );
                  })}

                <CategoryProjectTotal category={category} />
              </div>
            ))}
          </div>
        </section>
      )}
    </div>
  );
}