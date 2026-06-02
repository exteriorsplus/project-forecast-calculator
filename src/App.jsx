import { useCallback, useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
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
const PM_COMMISSION_RATE = 0.2;

const projectManagers = [
  {
    name: "Jamie Jenkins",
    slug: "jamiejenkins",
    image: "/pm/jamiejenkins.jpg",
    password: "jamie123",
    activeGoal: true,
  },
  {
    name: "Megan Rice",
    slug: "meganrice",
    image: "/pm/meganrice.jpg",
    password: "megan123",
    activeGoal: true,
  },
  {
    name: "Dani Cole",
    slug: "danicole",
    image: "/pm/danicole.jpg",
    password: "dani123",
    activeGoal: true,
  },
  {
    name: "John Fincher",
    slug: "johnfincher",
    image: "/pm/johnfincher.jpg",
    password: "john123",
    activeGoal: true,
  },
  {
    name: "Andrew Painter",
    slug: "andrewpainter",
    image: "/pm/andrewpainter.jpg",
    password: "andrew123",
    activeGoal: true,
  },
  {
    name: "George Anim",
    slug: "georgeanim",
    image: "/pm/georgeanim.jpg",
    password: "george123",
    activeGoal: true,
  },
  {
    name: "William Dye",
    slug: "williamdye",
    image: "/pm/williamdye.jpg",
    password: "william123",
    activeGoal: true,
  },
{
  name: "Mike Harr",
  slug: "mikeharr",
  image: "/pm/mikeharr.jpg",
  password: "mike123",
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

const WORK_TYPE_ORDER = ["Retail", "Insurance", "Repair", "Service"];

const money = (value) =>
  Number(value || 0).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
  });

  const formatMoneyInput = (value) => {
  const number = String(value || "").replace(/[^\d.]/g, "");

  if (!number) return "";

  return `$${Number(number).toLocaleString("en-US")}`;
};

const cleanMoneyInput = (value) => {
  return String(value || "").replace(/[^\d.]/g, "");
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

  if (isRate) {
    const pointDifference = currentNumber - comparisonNumber;

    return {
      label: `${pointDifference >= 0 ? "+" : ""}${(pointDifference * 100).toFixed(
        1
      )} pts`,
      className: pointDifference >= 0 ? "positive" : "negative",
      rawDifference: pointDifference,
    };
  }

  const percentDifference = (currentNumber - comparisonNumber) / comparisonNumber;

  return {
    label: `${percentDifference >= 0 ? "+" : ""}${(
      percentDifference * 100
    ).toFixed(1)}%`,
    className: percentDifference >= 0 ? "positive" : "negative",
    rawDifference: percentDifference,
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

function dateOnly(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function normalizeTrade(value) {
  const text = String(value || "").trim().toLowerCase();

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

  const currentPM = getCurrentProjectManager();
  const isPMPortal = Boolean(currentPM);

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
      setDataStatus("No sales.xlsx file found yet.");
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
    .slice(teamSectionStart, teamSectionStart + 8)
    .find(
      (row) =>
        String(row?.[0] || "").trim().toLowerCase() ===
        metric.toLowerCase()
    );
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
    const metricRow = pmRows
  .slice(startIndex + 1, startIndex + 10)
      .find(
        (row) =>
          String(row?.[0] || "").trim().toLowerCase() === metric.toLowerCase()
      );

    if (!metricRow) return 0;

    const columnIndex = headerRow.findIndex(
      (cell) => formatPMMonth(cell) === monthLabel
    );

    if (columnIndex < 0) return 0;

    return Number(metricRow[columnIndex] || 0);
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

    return Number(row[columnIndex] || 0);
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

  const getPMCustomSalesData = (pmName) => {
    if (!pmStartDate || !pmEndDate) {
      return {
        contractTotal: 0,
        contracts: 0,
        averageContract: 0,
      };
    }

    const start = dateOnly(parseInputDate(pmStartDate));
    const end = dateOnly(parseInputDate(pmEndDate));

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

      contractTotal += amount;
      contracts += 1;
    });

    return {
      contractTotal,
      contracts,
      averageContract: contracts > 0 ? contractTotal / contracts : 0,
    };
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

const ytdRevenue = ytdMonths.reduce(
  (sum, month) =>
    sum + getPMMetric(pmName, "Contract Total", month),
  0
);

const ytdContracts = ytdMonths.reduce(
  (sum, month) =>
    sum + getPMMetric(pmName, "Contracts", month),
  0
);

const ytdAverageContract =
  ytdContracts > 0 ? ytdRevenue / ytdContracts : 0;

const ytdClosingRate = 0;

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

const customSalesData = getPMCustomSalesData(pmName);

const contractTotal =
  pmDateMode === "custom"
    ? customSalesData.contractTotal
    : activeMonths.reduce(
        (sum, month) =>
          sum + getPMMetric(pmName, "Contract Total", month),
        0
      );

const contracts =
  pmDateMode === "custom"
    ? customSalesData.contracts
    : activeMonths.reduce(
        (sum, month) =>
          sum + getPMMetric(pmName, "Contracts", month),
        0
      );

const averageContract =
  pmDateMode === "custom"
    ? customSalesData.averageContract
    : contracts > 0
    ? contractTotal / contracts
    : 0;

const closingRate =
  pmDateMode === "custom"
    ? (
        activeMonths.reduce(
          (sum, month) =>
            sum + getPMMetric(pmName, "Closing Rate", month),
          0
        ) / activeMonths.length
      )
    : getPMMetric(pmName, "Closing Rate", selectedMonth);


    const lyContractTotal = getPMMetric(
      (currentPM?.name || projectManagers[0].name),
      "Contract Total",
      lastYearMonth
    );
    const lyContracts = getPMMetric((currentPM?.name || projectManagers[0].name), "Contracts", lastYearMonth);
    const lyAverageContract = getPMMetric(
      (currentPM?.name || projectManagers[0].name),
      "Average Contract",
      lastYearMonth
    );
    const lyClosingRate = getPMMetric(
      (currentPM?.name || projectManagers[0].name),
      "Closing Rate",
      lastYearMonth
    );

const teamContractTotal = getTeamMetric(
  "Contract Total Average",
  selectedMonth
);

const teamContracts = getTeamMetric(
  "Contracts",
  selectedMonth
);

const teamAverageContract = getTeamMetric(
  "Average Contract",
  selectedMonth
);

const teamClosingRate = getTeamMetric(
  "Closing Rate",
  selectedMonth
);

    const revenueRank = getRankForMetric(
      "Contract Total",
      selectedMonth,
      (currentPM?.name || projectManagers[0].name)
    );
    const closingRateRank = getRankForMetric(
      "Closing Rate",
      selectedMonth,
      (currentPM?.name || projectManagers[0].name)
    );

    const saleAmount = Number(pmSaleAmount || 0);
    const projectedYTDRevenue = ytdRevenue + saleAmount;
    const commission = saleAmount * PM_COMMISSION_RATE;
const individualGoal =
  PM_GOALS[currentPM?.name || projectManagers[0].name] || 0;

const goalPercent = individualGoal > 0 ? ytdRevenue / individualGoal : 0;
const projectedGoalPercent =
  individualGoal > 0 ? projectedYTDRevenue / individualGoal : 0;
const remainingToGoal = Math.max(individualGoal - ytdRevenue, 0);
const monthlyGoal = individualGoal / 12;

const monthlyPace =
  activeMonths.length > 0 ? contractTotal / activeMonths.length : contractTotal;

const monthlyGoalPercent =
  monthlyGoal > 0 ? monthlyPace / monthlyGoal : 0;

const monthlyRemaining = monthlyGoal - monthlyPace;

const quarterStartIndex =
  fiscalMonthIndex >= 0 ? Math.floor(fiscalMonthIndex / 3) * 3 : 0;

const quarterMonths = fiscalMonths.slice(
  quarterStartIndex,
  quarterStartIndex + 3
);

const quarterRevenue = quarterMonths.reduce(
  (sum, month) =>
    sum + getPMMetric(pmName, "Contract Total", month),
  0
);

const quarterlyGoal = individualGoal / 4;
const quarterlyGoalPercent =
  quarterlyGoal > 0 ? quarterRevenue / quarterlyGoal : 0;
const quarterlyRemaining =
  quarterlyGoal - quarterRevenue;
  
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
      revenueRank,
      closingRateRank,
      saleAmount,
      projectedYTDRevenue,
      commission,
      goalPercent,
      projectedGoalPercent,
      remainingToGoal,
      monthlyGoal,
      monthlyPace,
monthlyGoalPercent,
monthlyRemaining,
quarterRevenue,
quarterlyGoal,
quarterlyGoalPercent,
quarterlyRemaining,
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
}, [pmAuthorized]);

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
        <img src="/logo.png" alt="Logo" className="logo" />
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

const exceededGoalMessages = [
  "Outstanding work. You've already surpassed your goal for this period.",
  "Excellent performance. You've exceeded expectations and continue to set the pace.",
  "You've reached your goal ahead of schedule and are building momentum.",
  "Fantastic job. Your production has already cleared the target for this period.",
  "You're operating at a very high level and have successfully exceeded your goal."
];

const onPaceMessages = [
  "You're on pace to reach your goal and finish the period strong.",
  "Solid progress. Current production trends support achieving your target.",
  "You're tracking well toward your goal and maintaining healthy momentum.",
  "Keep doing what you're doing. You're positioned to reach your target.",
  "Performance remains strong and you're currently pacing toward success."
];

const progressMessages = [
  "You're making steady progress toward your goal with time remaining.",
  "Momentum is building and there is still plenty of opportunity ahead.",
  "Consistent effort now can create a strong finish to the period.",
  "The foundation is there. Continue focusing on quality opportunities.",
  "Progress remains positive and there is room to accelerate."
];

const pushMessages = [
  "Keep pushing. A strong finish can still put you back on pace.",
  "There is plenty of opportunity remaining to improve results.",
  "Focus on controllable activities and the numbers will follow.",
  "A few strong weeks can quickly change the trajectory.",
  "Stay consistent and continue building momentum."
];

const revenueStrongMessages = [
  "Revenue production is outperforming the team average.",
  "Your revenue generation continues to lead the way.",
  "You're producing above the team's current pace.",
  "Revenue performance remains a key strength.",
  "Your production levels are setting a strong example."
];

const revenueEliteMessages = [
  "Revenue production is significantly outperforming the team average.",
  "You're operating well above the team's current production pace.",
  "Your revenue performance ranks among the strongest on the team.",
  "Production continues to separate you from the field.",
  "You're creating substantial value compared with team benchmarks."
];

const revenueNeedsMessages = [
  "Revenue is currently below the team average, but opportunities remain.",
  "Closing a few additional projects could quickly narrow the gap.",
  "There is still time to improve revenue performance this period.",
  "The team benchmark remains within reach.",
  "Focus on pipeline conversion and revenue growth will follow."
];

const avgStrongMessages = [
  "Your average contract value continues to outperform the team.",
  "Customers are trusting you with larger projects than average.",
  "Contract quality remains one of your strengths.",
  "Higher-value projects are driving strong results.",
  "Your average sale size remains above benchmark levels."
];

const avgNeedsMessages = [
  "Increasing average contract value could have a meaningful impact.",
  "Larger project opportunities could accelerate growth.",
  "Focusing on project scope may improve overall production.",
  "Even modest gains in average contract size would create significant upside.",
  "Exploring higher-value opportunities may strengthen results."
];

const closeStrongMessages = [
  "Your closing efficiency remains one of your strongest advantages.",
  "Customers continue responding well to your sales process.",
  "Conversion performance is helping drive strong results.",
  "Your closing rate continues to outperform expectations.",
  "Strong conversion efficiency remains a competitive advantage."
];

const closeNeedsMessages = [
  "Improving conversion efficiency could unlock additional growth.",
  "Small gains in closing rate could have a large impact on production.",
  "More effective follow-up may improve conversion results.",
  "There is opportunity to strengthen closing performance.",
  "Improved conversion rates could significantly boost revenue."
];

const pickRandom = (messages) =>
  messages[Math.floor(Math.random() * messages.length)];

const crushingItMessages = [
  "You're crushing it. Production is well above goal.",
  "Outstanding work. You're significantly ahead of pace.",
  "You're setting the standard for performance this month.",
  "Exceptional production. Keep your foot on the gas.",
  "Your hard work is paying off in a big way.",
  "Momentum is strong and results continue to follow.",
  "You're operating at a very high level right now.",
  "Fantastic job. You're well ahead of target.",
  "This is the kind of month people remember.",
  "Keep doing what you're doing. It's working.",
];

const onTrackMessages = [
  "You're on track to achieve your goal.",
  "Great progress. Keep pushing toward the finish line.",
  "You're in a strong position to finish successfully.",
  "Steady effort is producing solid results.",
  "The goal is well within reach.",
  "Momentum is building in the right direction.",
  "You're trending toward a successful month.",
  "Stay focused and keep stacking wins.",
  "Your consistency is paying off.",
  "Keep the pressure on and finish strong."
];

const closeMessages = [
  "You're closer than you think. Keep pushing.",
  "A strong finish can put you over the top.",
  "The goal remains within reach.",
  "Keep building momentum. You're not far away.",
  "Stay focused. A few more wins could make the difference.",
  "The opportunity is still there.",
  "You're in striking distance of your target.",
  "Now is the time to finish strong.",
  "Stay disciplined and keep moving forward.",
  "Every contract matters from here."
];

const needsPushMessages = [
  "Keep grinding. Consistent effort pays off.",
  "The month isn't over yet. Stay focused.",
  "Every opportunity matters from this point forward.",
  "Small wins add up quickly.",
  "Keep working the process and results will follow.",
  "Stay persistent and keep creating opportunities.",
  "There's still time to improve your position.",
  "Focus on the next opportunity, not the scoreboard.",
  "Progress starts with consistent activity.",
  "Keep your energy high and stay engaged.",
];

const goalMotivation = (() => {
  const pct = pmData.monthlyGoalPercent;

  if (pct >= 1.5)
    return pickRandom(crushingItMessages);

  if (pct >= 0.9)
    return pickRandom(onTrackMessages);

  if (pct >= 0.7)
    return pickRandom(closeMessages);

  return pickRandom(needsPushMessages);
})();

const generatePMInsight = ({
  goalPercent,
  revenueVsTeam,
  averageVsTeam,
  closingRateVsTeam,
}) => {
  const messages = [];

  if (goalPercent >= 1) {
messages.push(
  pickRandom(exceededGoalMessages)
);
  } else if (goalPercent >= 0.8) {
messages.push(
  pickRandom(onPaceMessages)
);
  } else if (goalPercent >= 0.6) {
messages.push(
  pickRandom(progressMessages)
);
  } else {
messages.push(
  pickRandom(pushMessages)
);
  }

if (revenueVsTeam > 0.25) {
  messages.push(
    pickRandom(revenueEliteMessages)
  );
}
else if (revenueVsTeam > 0) {
  messages.push(
    pickRandom(revenueStrongMessages)
  );
}
else {
  messages.push(
    pickRandom(revenueNeedsMessages)
  );
}

if (averageVsTeam > 0.1) {
  messages.push(
    pickRandom(avgStrongMessages)
  );
} else if (averageVsTeam < -0.1) {
  messages.push(
    pickRandom(avgNeedsMessages)
  );
}

if (closingRateVsTeam > 0.05) {
  messages.push(
    pickRandom(closeStrongMessages)
  );
} else if (closingRateVsTeam < -0.05) {
  messages.push(
    pickRandom(closeNeedsMessages)
  );
}

  return messages.slice(0, 3).join(" ");
};

const pmInsight = generatePMInsight({
  goalPercent: pmData.goalPercent || 0,
  revenueVsTeam: revenueVsTeam?.rawDifference || 0,
  averageVsTeam: averageVsTeam?.rawDifference || 0,
  closingRateVsTeam: closingRateVsTeam?.rawDifference || 0,
});
const isCustomMode = pmDateMode === "custom";
const selectedMonthDate = new Date(pmData.selectedMonth);
const now = new Date();

const isPastMonth =
  selectedMonthDate.getFullYear() < now.getFullYear() ||
  (selectedMonthDate.getFullYear() === now.getFullYear() &&
    selectedMonthDate.getMonth() < now.getMonth());

const monthlyGoalClass =
  pmData.monthlyGoalPercent >= 1
    ? "positive"
    : isPastMonth
    ? "negative"
    : "warning";

const selectedMonthIndex = selectedMonthDate.getMonth();

let quarterEndMonth;

if ([10, 11, 0].includes(selectedMonthIndex)) {
  quarterEndMonth = 0; // Nov-Dec-Jan
} else if ([1, 2, 3].includes(selectedMonthIndex)) {
  quarterEndMonth = 3; // Feb-Mar-Apr
} else if ([4, 5, 6].includes(selectedMonthIndex)) {
  quarterEndMonth = 6; // May-Jun-Jul
} else {
  quarterEndMonth = 9; // Aug-Sep-Oct
}

const quarterComplete =
  selectedMonthDate.getFullYear() < now.getFullYear() ||
  (
    selectedMonthDate.getFullYear() === now.getFullYear() &&
    quarterEndMonth < now.getMonth()
  );

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
            <img src="/logo.png" alt="Logo" className="logo" />
            <h1>{currentPM.name} Dashboard</h1>
            <div className="pm-insight-box">
  <div className="pm-insight-title">
    Performance Insight
  </div>

  <p>{pmInsight}</p>
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
      <div>
        <small>Revenue Rank</small>
        <strong>{revenueRankLabel}</strong>
      </div>

      <div>
        <small>Closing Rate Rank</small>
        <strong>{closingRankLabel}</strong>
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
  comparisonLabel={isCustomMode ? null : `vs ${pmData.lastYearMonth}`}
  comparisonValue={isCustomMode ? null : money(pmData.lyContractTotal)}
  difference={isCustomMode ? null : revenueVsLY}
/>

<PMMetricCard
  label="Contracts"
  value={Math.round(pmData.contracts || 0)}
  comparisonLabel={isCustomMode ? null : `vs ${pmData.lastYearMonth}`}
  comparisonValue={isCustomMode ? null : Math.round(pmData.lyContracts || 0)}
  difference={isCustomMode ? null : contractsVsLY}
/>

<PMMetricCard
  label="Average Contract"
  value={money(pmData.averageContract)}
  comparisonLabel={isCustomMode ? null : `vs ${pmData.lastYearMonth}`}
  comparisonValue={isCustomMode ? null : money(pmData.lyAverageContract)}
  difference={isCustomMode ? null : averageVsLY}
/>

<PMMetricCard
  label="Closing Rate"
  value={displayPercent(pmData.closingRate, 1)}
  comparisonLabel={isCustomMode ? null : `vs ${pmData.lastYearMonth}`}
  comparisonValue={isCustomMode ? null : displayPercent(pmData.lyClosingRate, 1)}
  difference={isCustomMode ? null : closingRateVsLY}
/>
            </div>
          </div>

          {!isCustomMode && (
  <div className="pm-section-card">
    <h2>Team Comparison</h2>

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
)}

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
comparisonValue={money(pmData.monthlyPace)}
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
  customText={goalMotivation}
/>
<PMMetricCard
  label="Quarterly Goal"
  value={money(pmData.quarterlyGoal)}
  comparisonLabel="Current Quarter"
  comparisonValue={money(pmData.quarterRevenue)}
  difference={{
    label: `${displayPercent(pmData.quarterlyGoalPercent, 1)} Complete`,
    className: quarterlyGoalClass,
  }}
/>

    <PMMetricCard
      label="Quarterly Remaining"
      value={money(pmData.quarterlyRemaining)}
    />
  </div>
</div>
          <div className="pm-commission-card">
            <h2>Commission Calculator</h2>

            <label>
              Potential Sale Amount
              <input
                type="text"
                inputMode="decimal"
                value={formatMoneyInput(pmSaleAmount)}
                onChange={(event) =>
                  setPmSaleAmount(cleanMoneyInput(event.target.value))
                }
              />
            </label>

            <div className="pm-commission-grid">
              <div>
                <span>Commission Rate</span>
                <strong>20%</strong>
              </div>

              <div>
                <span>Estimated Commission</span>
                <strong>{money(pmData.commission)}</strong>
              </div>

              <div>
                <span>Projected YTD Revenue</span>
                <strong>{money(pmData.projectedYTDRevenue)}</strong>
              </div>

              <div>
                <span>Projected Goal %</span>
                <strong>{displayPercent(pmData.projectedGoalPercent, 1)}</strong>
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