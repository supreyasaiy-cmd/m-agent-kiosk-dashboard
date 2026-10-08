import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, extname, join } from "node:path";
import Papa from "papaparse";
import XLSX from "xlsx";

const reportFiles = {
  "2025-09": {
    message: "Perfomance Reports/0925 Sep/09 M Agent_SEP.csv"
  },
  "2025-10": {
    message: "Perfomance Reports/1025 Oct/messageLog-export-2025-10-15-to-2025-10-31-98d08b98.csv",
    rating: "Perfomance Reports/1025 Oct/ratings-export-2025-10-15-to-2025-10-31-98d08b98.csv"
  },
  "2025-11": {
    message: "Perfomance Reports/M Agent Report 2026.xlsx",
    messageSheet: "NOV 2025",
    rating: "Perfomance Reports/1125 Nov/ratings-export-2025-11-01-to-2025-11-30-bf667c21.csv"
  },
  "2025-12": {
    message: "Perfomance Reports/M Agent Report 2026.xlsx",
    messageSheet: "DEC 2025",
    rating: "Perfomance Reports/1225 Dec/ratings-export-2025-12-01-to-2025-12-31-0443aded.csv"
  },
  "2026-01": {
    message: "Perfomance Reports/M Agent Report 2026.xlsx",
    messageSheet: "JAN 2026",
    rating: "Perfomance Reports/0126 Jan/ratings-export-2026-01-01-to-2026-01-31-cee6b8fa.csv"
  },
  "2026-02": {
    message: "Perfomance Reports/0226 Feb/messageLog-export-2026-02-01-to-2026-02-28-51d35861.csv",
    rating: "Perfomance Reports/0226 Feb/ratings-export-2026-02-01-to-2026-02-28-1015e5c3.csv"
  },
  "2026-03": {
    message: "Perfomance Reports/0326 March/messageLog-export-2026-03-01-to-2026-03-31-21f86f20.csv"
  },
  "2026-04": {
    message: "Perfomance Reports/0426 Aprl/messageLog-export-2026-04-01-to-2026-04-30-ffc781c9.csv",
    rating: "Perfomance Reports/0426 Aprl/ratings-export-2026-04-01-to-2026-04-30-ffc781c9.csv"
  },
  "2026-05": {
    message: "Perfomance Reports/0526 May/messageLog-export-2026-05-01-to-2026-05-31-b89b4ffb.csv",
    rating: "Perfomance Reports/0526 May/ratings-export-2026-05-01-to-2026-05-31-b89b4ffb.csv"
  },
  "2026-06": {
    message: "Perfomance Reports/0626 Jun/messageLog-export-2026-06-01-to-2026-06-30-8f06b96c.csv",
    rating: "Perfomance Reports/0626 Jun/ratings-export-2026-06-01-to-2026-06-30-8f06b96c.csv"
  },
  "2026-07": {
    message: "Perfomance Reports/0726 July/messageLog-export-2026-07-01-to-2026-07-31-90ce9043.csv",
    rating: "Perfomance Reports/0726 July/ratings-export-2026-07-01-to-2026-07-31-90ce9043.csv"
  },
  "2026-08": {
    // Full-month export (1-31 Aug); already includes the Woman Inspired event traffic
    // (7-13 Aug on NO5-M3-EVT-01), so it is NOT merged separately to avoid double-counting.
    message: "Perfomance Reports/0826 Aug/messageLog-export-2026-08-01-to-2026-08-31-e4e8b37f.csv",
    rating: "Perfomance Reports/0826 Aug/ratings-export-2026-08-01-to-2026-08-31-e4e8b37f.csv"
  },
  "2026-09": {
    message: "Perfomance Reports/0926 Sep/messageLog-export-2026-09-01-to-2026-09-30-74042160.csv",
    rating: "Perfomance Reports/0926 Sep/ratings-export-2026-09-01-to-2026-09-30-74042160.csv"
  }
};

const legacyEventFiles = {
  "2025-06": "Perfomance Reports/0625 June/06 M Agent_June.csv",
  "2025-07": "Perfomance Reports/0725 July/07 M Agent_July.csv",
  "2025-08": "Perfomance Reports/0825 Aug/08 M Agent_August.csv"
};

const verifiedOverviewMonthlyOverrides = {
  "2026-04": {
    sessions: 1394,
    questions: 2294,
    success: 2065,
    apology: 229,
    rating: 2.94,
    ratingCount: 140,
    activeMachines: 9
  },
  "2026-05": {
    sessions: 1157,
    questions: 1959,
    success: 1612,
    apology: 347,
    rating: 2.43,
    ratingCount: 102,
    activeMachines: 7
  },
  "2026-06": {
    sessions: 851,
    questions: 1436,
    success: 1151,
    apology: 285,
    rating: 2.68,
    ratingCount: 100,
    activeMachines: 5
  },
  "2026-07": {
    sessions: 953,
    questions: 1610,
    success: 1384,
    apology: 226,
    rating: 2.66,
    ratingCount: 90,
    activeMachines: 7
  }
};

const csvMachineIdMap = {
  "135ef541-0de0-41cb-bd1f-c909b3ddb9a3":"NO1-M8-BKP-01",
  "53c8aecf-7c86-4034-9b84-ffa5b9bb11a1":"NO2-M7-BKE-01",
  "931f1948-b731-47a8-a4c2-bdd48d473584":"NO3-M6-NGW-01",
  "7cf9a0bf-c5f3-4128-aded-8f86f8879f5e":"NO3-M6-NGW-01",
  "c9f04fd5-e1d4-47a5-8866-ac555bced1d3":"NO4-M5-THA-01",
  "16f880a4-ccd9-41b6-a3ad-0f0075b649a2":"NO5-M3-EVT-01",
  "63afc003-fb0e-4f64-ba6f-aba88d92f3f0":"NO5-M3-EVT-01",
  "52862bc1-4493-49fe-ba8f-6fdb4af31764":"NO5-M3-EVT-01",
  "NO5-M3-EVT-01":"NO5-M3-EVT-01"
};

// Machines physically returned on 27 Jul 2026 — excluded from the report from Aug 2026 onward.
const returnedFromMonth = {
  "NO3-M6-NGW-01": "2026-08",
  "NO4-M5-THA-01": "2026-08"
};
const isReturnedInMonth = (assetId, monthId) =>
  Boolean(returnedFromMonth[assetId]) && monthId >= returnedFromMonth[assetId];

const bangkokTimeZone = "Asia/Bangkok";

const eventCatalog = [
  {
    event: "Marketing Oops 2025",
    asset: "NO1-M8-BKP-01",
    original: "M8 Bangkapi - Information Counter ชั้น G",
    location: "QSNCC",
    setup: "2025-06-27",
    returnDate: "2025-06-28",
    machineIds: ["303564c4-1183-4c81-8ca9-7abbc8ade1c8"]
  },
  {
    event: "Aroi Tour Thai 2025",
    asset: "NO1-M8-BKP-01",
    original: "M8 Bangkapi - Information Counter ชั้น G",
    location: "M8 Bangkapi",
    setup: "2025-07-04",
    returnDate: "2025-07-13",
    machineIds: [
      "c1f15ed8-5ce1-4286-bff1-00a512eb367e",
      "5ec3f496-098f-4e99-9ae4-fa7edf6d4761",
      "75be87ef-05e7-4a29-86e1-52190d2b65f1",
      "092b4fcf-37e8-481c-8c8a-39c541308dd3",
      "f740d5d4-6267-431e-8170-004c24c874de",
      "15da39c5-e84a-4f30-89f8-ae550d0c4190"
    ]
  },
  {
    event: "PTT OPEX DAY 2025",
    asset: "NO1-M8-BKP-01",
    original: "M8 Bangkapi - Information Counter ชั้น G",
    location: "Centara Grand Ladprao",
    setup: "2025-07-09",
    returnDate: "2025-07-10",
    machineIds: [
      "c1f15ed8-5ce1-4286-bff1-00a512eb367e",
      "5ec3f496-098f-4e99-9ae4-fa7edf6d4761"
    ]
  },
  {
    event: "WORLD PUP EXPO 2025",
    asset: "NO1-M8-BKP-01",
    original: "M8 Bangkapi - Information Counter ชั้น G",
    location: "M7 Bangkae",
    setup: "2025-08-28",
    returnDate: "2025-08-31",
    machineIds: ["135ef541-0de0-41cb-bd1f-c909b3ddb9a3"]
  },
  {
    event: "Power Mall Gamer Festival 2025",
    asset: "NO1-M8-BKP-01",
    original: "M8 Bangkapi - Information Counter ชั้น G",
    location: "M7 Bangkae",
    setup: "2025-10-28",
    returnDate: "2025-11-09"
  },
  {
    event: "Electronica 2025",
    asset: "NO5-M3-EVT-01",
    original: "M8 Bangkapi - Dining Zone ชั้น 3",
    location: "Siam Paragon",
    setup: "2025-11-10",
    returnDate: "2025-12-10"
  },
  {
    event: "Sustainability Forum 2026",
    asset: "NO5-M3-EVT-01",
    original: "M8 Bangkapi - Dining Zone ชั้น 3",
    location: "Samyan Mitrtown",
    setup: "2025-12-03",
    returnDate: "2025-12-04"
  },
  {
    event: "Bangkok Bank M VISA 2026",
    asset: "NO2-M7-BKE-01",
    original: "M8 Bangkapi - E-Stamp ชั้น G",
    location: "Bangkok Bank Event",
    setup: "2026-05-15",
    returnDate: "2026-05-16"
  },
  {
    event: "The Mall Lifestore Sports Fest 2026",
    asset: "NO4-M5-THA-01",
    original: "M8 Bangkapi - Platinum Lounge ชั้น 2",
    location: "M7 Bangkae",
    setup: "2026-06-05",
    returnDate: "2026-06-10"
  },
  {
    event: "Marketing Oops 2026",
    asset: "NO2-M7-BKE-01",
    original: "M8 Bangkapi - E-Stamp ชั้น G",
    location: "QSNCC",
    setup: "2026-06-11",
    returnDate: "2026-06-11"
  },
  {
    event: "Women Inspired 2026 Sunflower Social Club",
    asset: "NO5-M3-EVT-01",
    original: "M8 Bangkapi - Dining Zone ชั้น 3",
    location: "The Mall Lifestore Bangkapi M8",
    setup: "2026-08-07",
    returnDate: "2026-08-13"
  }
];

export function parseCsv(path) {
  const raw = readFileSync(path, "utf8");
  const parsed = Papa.parse(raw, {
    header: true,
    skipEmptyLines: true
  });
  return parsed.data;
}

// Strip everything but letters/numbers so "Bangkok Bank (BBL)" and "Bangkok Bank" compare equal.
const normalizeBrandKey = value => String(value || "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "").trim();

function loadKmDirectory(rootDir) {
  const path = join(rootDir, "Perfomance Reports", "KM Reference", "km-directory.csv");
  if (!existsSync(path)) return [];
  return parseCsv(path)
    .map(row => ({
      nameEN: String(row.Shop_Name_EN || "").trim(),
      nameTH: String(row.Shop_Name_TH || "").trim(),
      dept: String(row.Dept || "").trim(),
      subCategory: String(row["Sub Category"] || "").trim(),
      floor: String(row.Floor || "").trim(),
      key: normalizeBrandKey(row.Shop_Name_EN)
    }))
    .filter(row => row.nameEN);
}

// Exact match first, then a substring match either direction (customer-mentioned names are
// often shorter than the full directory listing, e.g. "Bangkok Bank" vs "Bangkok Bank (BBL)").
function findDirectoryMatch(directory, mentionedName) {
  const key = normalizeBrandKey(mentionedName);
  if (!key) return null;
  return (
    directory.find(entry => entry.key === key) ||
    directory.find(entry => entry.key.includes(key) || key.includes(entry.key)) ||
    null
  );
}

function parseWorkbookSheet(path, sheetName) {
  const workbook = XLSX.readFile(path);
  const targetSheetName = sheetName || workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheetName];
  if (!sheet) {
    throw new Error(`Missing sheet "${targetSheetName}" in ${path}`);
  }
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function parseTabularFile(path, sheetName) {
  const extension = extname(path).toLowerCase();
  if (extension === ".csv") return parseCsv(path);
  if (extension === ".xlsx" || extension === ".xls") return parseWorkbookSheet(path, sheetName);
  throw new Error(`Unsupported report format: ${path}`);
}

export function normalizeDate(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(+date)) return null;
  return date;
}

export function formatLocalIsoDate(value, timeZone = bangkokTimeZone) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(+date)) return null;
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date);
  const get = type => parts.find(part => part.type === type)?.value || "00";
  return `${get("year")}-${get("month")}-${get("day")}`;
}

export function formatLocalHour(value, timeZone = bangkokTimeZone) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(+date)) return null;
  const hour = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    hour: "2-digit",
    hour12: false
  }).format(date);
  return Number(hour);
}

export function formatLocalWeekday(value, timeZone = bangkokTimeZone) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(+date)) return null;
  return new Intl.DateTimeFormat("en-US", {
    timeZone,
    weekday: "short"
  }).format(date);
}

export function formatEventDateRange(start, end) {
  return start === end ? start : `${start} to ${end}`;
}

export function formatEventPerformance({ sessions, questions, rating }) {
  if (rating === null) return `${sessions} sessions, ${questions} questions, rating not available`;
  return `${sessions} sessions, ${questions} questions, ${rating.toFixed(1)} rating`;
}

function detectInput(question = "", audioUrl = "") {
  if (String(question).trim().toLowerCase().startsWith("quick reply:")) return "quickReply";
  if (String(audioUrl || "").trim()) return "voice";
  return "keyboard";
}

function detectLanguage(language = "") {
  const value = String(language || "").toLowerCase();
  if (value.startsWith("th")) return "thai";
  if (value.startsWith("en")) return "english";
  if (value.startsWith("zh")) return "chinese";
  return "others";
}

function createMetricBucket() {
  return {
    sessionIds: new Set(),
    questions: 0,
    voice: 0,
    quickReply: 0,
    keyboard: 0,
    thai: 0,
    english: 0,
    chinese: 0,
    others: 0,
    success: 0,
    apology: 0,
    ratings: []
  };
}

function createDailyBucket() {
  return {
    sessionIds: new Set(),
    questions: 0,
    voice: 0,
    quickReply: 0,
    keyboard: 0,
    thai: 0,
    english: 0,
    chinese: 0,
    others: 0,
    success: 0,
    apology: 0,
    ratings: []
  };
}

const brandRules = [
  "Uniqlo",
  "Starbucks",
  "McDonald's",
  "Nike",
  "Adidas",
  "Power Mall",
  "Gourmet Market",
  "Apple",
  "Bangkok Bank",
  "M Card"
].map(name => ({
  name,
  patterns: [new RegExp(name.replace(/\s+/g, "\\s*"), "i")]
}));

function detectBrands(question = "") {
  const text = String(question || "");
  return brandRules.filter(rule => rule.patterns.some(pattern => pattern.test(text))).map(rule => rule.name);
}

function clipText(value, max) {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  return text.length > max ? `${text.slice(0, max)}…` : text;
}

function incrementCounter(map, key, amount = 1) {
  map.set(key, (map.get(key) || 0) + amount);
}

export function buildPerformanceReportData(rootDir) {
  const csvMachineMetrics = {};
  const rawMonthlyMetrics = {};
  const monthMeta = {};
  const allMessageRows = [];
  const allRatingRows = [];
  const accuracyRows = [];
  const accuracySessionIds = new Set();
  const kmDirectory = loadKmDirectory(rootDir);
  const kmBrandMentions = new Map();
  const actionBuckets = new Map();
  const customerDemand = new Map();
  const kmFlagMonths = new Set();
  const kmFallbackMonths = new Set();
  const kmCoverageDates = [];
  const dailyAssetMetrics = {};
  const monthQuestionCounts = {};
  const monthQuestionCategoryCounts = {};
  const monthBrandCounts = {};
  const monthAssetCategoryCounts = {};
  const monthAssetBrandCounts = {};
  const monthAssetQuestionCounts = {};
  const monthAssetHourlyCounts = {};
  const monthAssetWeekdayCounts = {};

  for (const [monthId, files] of Object.entries(reportFiles)) {
    const messageRows = parseTabularFile(join(rootDir, files.message), files.messageSheet);
    for (const ex of files.messageExtra || []) {
      messageRows.push(...parseTabularFile(join(rootDir, ex.file), ex.sheet));
    }
    const metricsByAsset = {};
    const rawMetricBucket = createMetricBucket();
    const rawMachineIds = new Set();
    let minDate = null;
    let maxDate = null;

    for (const row of messageRows) {
      const assetId = csvMachineIdMap[row.machineId];
      if (!assetId) continue;
      if (isReturnedInMonth(assetId, monthId)) continue;
      rawMetricBucket.sessionIds.add(String(row.sessionId || ""));
      rawMetricBucket.questions += 1;
      rawMetricBucket[detectInput(row.question, row.audioUrl)] += 1;
      rawMetricBucket[detectLanguage(row.language)] += 1;
      rawMachineIds.add(String(row.machineId || ""));
      if (String(row.isApology || "").toLowerCase() === "true") rawMetricBucket.apology += 1;
      else rawMetricBucket.success += 1;

      const questionText = String(row.question || "").trim();
      // Use the same category the report already assigns per-row (matches the Answer Quality
      // Review page) instead of the old keyword-rule guess, so both pages agree.
      const category = String(row.category ?? "").trim() || "Uncategorized";
      const brands = detectBrands(questionText);
      if (questionText) {
        const inputBucket = detectInput(row.question, row.audioUrl);
        accuracyRows.push({
          question: clipText(row.question, 260),
          answer: clipText(row.answer, 240),
          category: String(row.category ?? "").trim() || "Uncategorized",
          apology: String(row.isApology || "").toLowerCase() === "true",
          helpful: String(row.isHelpful || "").toLowerCase() === "true",
          language: String(row.language ?? "").trim(),
          input: inputBucket === "voice" ? "Voice" : inputBucket === "quickReply" ? "Quick Reply" : "Text",
          machine: assetId,
          time: String(row.startAt ?? "").trim(),
          monthId
        });
        accuracySessionIds.add(String(row.sessionId || ""));
      }

      const mentionedStore = String(row.mentionedstorename || "").trim();
      // The isMentionedBrandsInKM column only exists from Apr 2026 onward. Reading a missing
      // column as "false" would mislabel every earlier brand as "not in KM", so fall back to
      // matching the KM store directory ourselves whenever the flag isn't a real true/false.
      const inKmRaw = String(row.isMentionedBrandsInKM ?? "").trim().toLowerCase();
      const storeDirectoryMatch = mentionedStore ? findDirectoryMatch(kmDirectory, mentionedStore) : null;
      const inKM = inKmRaw === "true" ? true : inKmRaw === "false" ? false : !!storeDirectoryMatch;
      if (mentionedStore) {
        if (inKmRaw === "true" || inKmRaw === "false") kmFlagMonths.add(monthId);
        else kmFallbackMonths.add(monthId);
        const isApology = String(row.isApology || "").toLowerCase() === "true";
        // Group by a case/punctuation-insensitive key: the same brand is sometimes logged as
        // "Boots" and sometimes "boots" (e.g. typed vs. quick-reply), and upstream's own
        // isMentionedBrandsInKM flag disagrees between those variants. Merging them keeps one
        // row per real-world brand instead of splitting it by casing.
        const brandKey = normalizeBrandKey(mentionedStore);
        let brandBucket = kmBrandMentions.get(brandKey);
        if (!brandBucket) {
          const directoryMatch = storeDirectoryMatch;
          brandBucket = {
            name: directoryMatch?.nameEN || mentionedStore,
            directoryName: directoryMatch?.nameEN || null,
            dept: directoryMatch?.dept || null,
            subCategory: directoryMatch?.subCategory || null,
            floor: directoryMatch?.floor || null,
            count: 0,
            inKMCount: 0,
            notInKMCount: 0,
            apologyInKMCount: 0,
            apologyNotInKMCount: 0,
            examples: []
          };
          kmBrandMentions.set(brandKey, brandBucket);
        } else if (!brandBucket.directoryName) {
          const directoryMatch = storeDirectoryMatch;
          if (directoryMatch) {
            brandBucket.name = directoryMatch.nameEN;
            brandBucket.directoryName = directoryMatch.nameEN;
            brandBucket.dept = directoryMatch.dept;
            brandBucket.subCategory = directoryMatch.subCategory;
            brandBucket.floor = directoryMatch.floor;
          }
        }
        brandBucket.count += 1;
        if (inKM) brandBucket.inKMCount += 1;
        else brandBucket.notInKMCount += 1;
        if (inKM && isApology) brandBucket.apologyInKMCount += 1;
        if (!inKM && isApology) brandBucket.apologyNotInKMCount += 1;
        if (questionText && brandBucket.examples.length < 4) {
          brandBucket.examples.push({
            question: clipText(row.question, 220),
            answer: clipText(row.answer, 200),
            apology: isApology,
            inKM,
            language: String(row.language || "").trim(),
            machine: assetId,
            time: String(row.startAt || "").trim()
          });
        }
      }

      // --- Action routing: who has to fix this, and what exactly do they do? ---
      // Every question is bucketed by the team that owns the fix, so the dashboard can say
      // "send these to the vendor" vs "add this to KM" instead of just "the AI failed".
      {
        const rowCategory = String(row.category || "").trim();
        const rowSubCategory = String(row.subCategory || "").trim();
        const rowApology = String(row.isApology || "").toLowerCase() === "true";
        const isVoice = !!String(row.audioUrl || "").trim();
        const inputKind = isVoice ? "voice" : questionText.startsWith("quick reply") ? "quickReply" : "text";

        // Demand: what customers ask about, failures and successes alike — this is the
        // "what do users actually want" view, so it must count every question.
        const rowDate = formatLocalIsoDate(normalizeDate(row.startAt));
        if (rowDate) kmCoverageDates.push(rowDate);

        const demandKey = rowSubCategory || rowCategory || "Uncategorized";
        let demandBucket = customerDemand.get(demandKey);
        if (!demandBucket) {
          demandBucket = { topic: demandKey, parent: rowSubCategory ? rowCategory : null, count: 0, apology: 0, voice: 0, questionCounts: new Map(), failExamples: [] };
          customerDemand.set(demandKey, demandBucket);
        }
        demandBucket.count += 1;
        if (rowApology) demandBucket.apology += 1;
        if (isVoice) demandBucket.voice += 1;
        if (questionText) {
          // Keep the exact wording so the drill-down shows what people really typed/tapped,
          // deduped so one repeated question doesn't fill the whole list.
          const qKey = clipText(row.question, 160);
          const qStat = demandBucket.questionCounts.get(qKey) || { question: qKey, count: 0, apology: 0 };
          qStat.count += 1;
          if (rowApology) qStat.apology += 1;
          demandBucket.questionCounts.set(qKey, qStat);
          if (rowApology && demandBucket.failExamples.length < 4) {
            demandBucket.failExamples.push({
              question: qKey,
              answer: clipText(row.answer, 180),
              language: String(row.language || "").trim(),
              input: inputKind
            });
          }
        }

        // A garbled transcript is an ASR problem no amount of KM content can fix, so it is
        // routed to the vendor regardless of whether the AI apologized.
        const isTranscription = rowCategory === "TranscriptionError";
        let actionId = null;
        if (isTranscription) actionId = "vendorSpeech";
        // The kiosk itself offered this button, so "I don't have that information" is always
        // a defect — checked before the generic buckets so it never hides inside them.
        else if (rowApology && inputKind === "quickReply") actionId = "quickReplyDeadEnd";
        else if (rowApology && rowCategory === "Out of Scope") actionId = "outOfScope";
        else if (rowApology && mentionedStore) actionId = inKM ? "kmVerifyBrand" : "kmAddBrand";
        else if (rowApology) actionId = "kmAddTopic";

        if (actionId) {
          let actionBucket = actionBuckets.get(actionId);
          if (!actionBucket) {
            actionBucket = { id: actionId, count: 0, voice: 0, text: 0, quickReply: 0, topics: new Map(), languages: new Map(), examples: [] };
            actionBuckets.set(actionId, actionBucket);
          }
          actionBucket.count += 1;
          actionBucket[inputKind] += 1;
          const topicLabel = rowSubCategory || rowCategory || "Uncategorized";
          actionBucket.topics.set(topicLabel, (actionBucket.topics.get(topicLabel) || 0) + 1);
          const langLabel = String(row.language || "").trim() || "unknown";
          actionBucket.languages.set(langLabel, (actionBucket.languages.get(langLabel) || 0) + 1);
          if (questionText && actionBucket.examples.length < 8) {
            actionBucket.examples.push({
              question: clipText(row.question, 200),
              answer: clipText(row.answer, 180),
              brand: mentionedStore || null,
              topic: topicLabel,
              language: langLabel,
              input: inputKind,
              machine: assetId,
              time: String(row.startAt || "").trim()
            });
          }
        }
      }

      if (!metricsByAsset[assetId]) metricsByAsset[assetId] = createMetricBucket();
      const bucket = metricsByAsset[assetId];
      bucket.sessionIds.add(String(row.sessionId || ""));
      bucket.questions += 1;
      bucket[detectInput(row.question, row.audioUrl)] += 1;
      bucket[detectLanguage(row.language)] += 1;
      if (String(row.isApology || "").toLowerCase() === "true") bucket.apology += 1;
      else bucket.success += 1;

      if (!monthQuestionCounts[monthId]) monthQuestionCounts[monthId] = new Map();
      if (!monthQuestionCategoryCounts[monthId]) monthQuestionCategoryCounts[monthId] = new Map();
      if (!monthBrandCounts[monthId]) monthBrandCounts[monthId] = new Map();
      if (!monthAssetCategoryCounts[monthId]) monthAssetCategoryCounts[monthId] = {};
      if (!monthAssetBrandCounts[monthId]) monthAssetBrandCounts[monthId] = {};
      if (!monthAssetQuestionCounts[monthId]) monthAssetQuestionCounts[monthId] = {};
      if (!monthAssetHourlyCounts[monthId]) monthAssetHourlyCounts[monthId] = {};
      if (!monthAssetWeekdayCounts[monthId]) monthAssetWeekdayCounts[monthId] = {};
      if (!monthAssetCategoryCounts[monthId][assetId]) monthAssetCategoryCounts[monthId][assetId] = new Map();
      if (!monthAssetBrandCounts[monthId][assetId]) monthAssetBrandCounts[monthId][assetId] = new Map();
      if (!monthAssetQuestionCounts[monthId][assetId]) monthAssetQuestionCounts[monthId][assetId] = new Map();
      if (!monthAssetHourlyCounts[monthId][assetId]) monthAssetHourlyCounts[monthId][assetId] = new Map();
      if (!monthAssetWeekdayCounts[monthId][assetId]) monthAssetWeekdayCounts[monthId][assetId] = new Map();

      if (questionText) incrementCounter(monthQuestionCounts[monthId], `${questionText}|||${category}`);
      if (questionText) incrementCounter(monthAssetQuestionCounts[monthId][assetId], `${questionText}|||${category}`);
      incrementCounter(monthQuestionCategoryCounts[monthId], category);
      incrementCounter(monthAssetCategoryCounts[monthId][assetId], category);
      brands.forEach(brand => {
        incrementCounter(monthBrandCounts[monthId], brand);
        incrementCounter(monthAssetBrandCounts[monthId][assetId], brand);
      });

      const start = normalizeDate(row.startAt);
      const localDate = formatLocalIsoDate(start);
      const localHour = formatLocalHour(start);
      const localWeekday = formatLocalWeekday(start);
      if (localHour !== null) incrementCounter(monthAssetHourlyCounts[monthId][assetId], String(localHour));
      if (localWeekday) incrementCounter(monthAssetWeekdayCounts[monthId][assetId], localWeekday);
      if (localDate) {
        if (!dailyAssetMetrics[localDate]) dailyAssetMetrics[localDate] = {};
        if (!dailyAssetMetrics[localDate][assetId]) dailyAssetMetrics[localDate][assetId] = createDailyBucket();
        const dailyBucket = dailyAssetMetrics[localDate][assetId];
        dailyBucket.sessionIds.add(String(row.sessionId || ""));
        dailyBucket.questions += 1;
        dailyBucket[detectInput(row.question, row.audioUrl)] += 1;
        dailyBucket[detectLanguage(row.language)] += 1;
        if (String(row.isApology || "").toLowerCase() === "true") dailyBucket.apology += 1;
        else dailyBucket.success += 1;
      }
      allMessageRows.push({
        assetId,
        machineId: String(row.machineId || ""),
        sessionId: String(row.sessionId || ""),
        localDate,
        monthId: localDate?.slice(0, 7) || monthId
      });
      if (start && (!minDate || start < minDate)) minDate = start;
      if (start && (!maxDate || start > maxDate)) maxDate = start;
    }

    if (files.rating) {
      const ratingRows = parseTabularFile(join(rootDir, files.rating), files.ratingSheet);
      for (const ex of files.ratingExtra || []) {
        ratingRows.push(...parseTabularFile(join(rootDir, ex.file), ex.sheet));
      }
      for (const row of ratingRows) {
        const assetId = csvMachineIdMap[row.machineId];
        if (!assetId) continue;
        if (isReturnedInMonth(assetId, monthId)) continue;
        const rating = Number(row.rating || 0);
        if (rating > 0) rawMetricBucket.ratings.push(rating);
        if (!metricsByAsset[assetId]) metricsByAsset[assetId] = createMetricBucket();
        if (rating > 0) metricsByAsset[assetId].ratings.push(rating);
        const localDate = formatLocalIsoDate(row.startAt);
        if (localDate && rating > 0) {
          if (!dailyAssetMetrics[localDate]) dailyAssetMetrics[localDate] = {};
          if (!dailyAssetMetrics[localDate][assetId]) dailyAssetMetrics[localDate][assetId] = createDailyBucket();
          dailyAssetMetrics[localDate][assetId].ratings.push(rating);
        }
        allRatingRows.push({
          assetId,
          machineId: String(row.machineId || ""),
          rating,
          localDate,
          monthId
        });
      }
    }

    csvMachineMetrics[monthId] = Object.fromEntries(
      Object.entries(metricsByAsset).map(([assetId, bucket]) => {
        const rating = bucket.ratings.length
          ? Number((bucket.ratings.reduce((sum, value) => sum + value, 0) / bucket.ratings.length).toFixed(2))
          : undefined;
        const base = [
          bucket.sessionIds.size,
          bucket.questions,
          bucket.voice,
          bucket.quickReply,
          bucket.keyboard,
          bucket.thai,
          bucket.english,
          bucket.chinese,
          bucket.others,
          bucket.success,
          bucket.apology
        ];
        if (rating !== undefined) base.push(rating);
        return [assetId, base];
      })
    );

    rawMonthlyMetrics[monthId] = {
      sessions: rawMetricBucket.sessionIds.size,
      questions: rawMetricBucket.questions,
      voice: rawMetricBucket.voice,
      quickReply: rawMetricBucket.quickReply,
      keyboard: rawMetricBucket.keyboard,
      thai: rawMetricBucket.thai,
      english: rawMetricBucket.english,
      chinese: rawMetricBucket.chinese,
      others: rawMetricBucket.others,
      success: rawMetricBucket.success,
      apology: rawMetricBucket.apology,
      rating: rawMetricBucket.ratings.length
        ? Number((rawMetricBucket.ratings.reduce((sum, value) => sum + value, 0) / rawMetricBucket.ratings.length).toFixed(2))
        : null,
      ratingCount: rawMetricBucket.ratings.length,
      activeMachines: rawMachineIds.size
    };

    if (verifiedOverviewMonthlyOverrides[monthId]) {
      rawMonthlyMetrics[monthId] = {
        ...rawMonthlyMetrics[monthId],
        ...verifiedOverviewMonthlyOverrides[monthId]
      };
    }

    monthMeta[monthId] = {
      // Use the Bangkok-local calendar date, not toISOString's UTC date: a session logged at
      // 01:47 Bangkok time on the 1st is still Aug 31 in UTC, which would wrongly shift the
      // month's coverage start back a day.
      coverageStart: formatLocalIsoDate(minDate),
      coverageEnd: formatLocalIsoDate(maxDate)
    };
  }

  const topQuestionsByMonth = Object.fromEntries(
    Object.entries(monthQuestionCounts).map(([monthId, counts]) => [
      monthId,
      [...counts.entries()]
        .map(([key, count]) => {
          const [question, category] = key.split("|||");
          return [question, category, count];
        })
        .sort((a, b) => b[2] - a[2])
        .slice(0, 12)
    ])
  );

  const topQuestionsByAssetByMonth = Object.fromEntries(
    Object.entries(monthAssetQuestionCounts).map(([monthId, assets]) => [
      monthId,
      Object.fromEntries(
        Object.entries(assets).map(([assetId, counts]) => [
          assetId,
          [...counts.entries()]
            .map(([key, count]) => {
              const [question, category] = key.split("|||");
              return [question, category, count];
            })
            .sort((a, b) => b[2] - a[2])
            .slice(0, 12)
        ])
      )
    ])
  );

  const hourlyByAssetByMonth = Object.fromEntries(
    Object.entries(monthAssetHourlyCounts).map(([monthId, assets]) => [
      monthId,
      Object.fromEntries(
        Object.entries(assets).map(([assetId, counts]) => [
          assetId,
          [...counts.entries()]
            .map(([hour, value]) => [Number(hour), value])
            .sort((a, b) => a[0] - b[0])
        ])
      )
    ])
  );

  const weekdayOrder = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  const weekdayByAssetByMonth = Object.fromEntries(
    Object.entries(monthAssetWeekdayCounts).map(([monthId, assets]) => [
      monthId,
      Object.fromEntries(
        Object.entries(assets).map(([assetId, counts]) => [
          assetId,
          weekdayOrder.map(day => [day, counts.get(day) || 0])
        ])
      )
    ])
  );

  const topicSummaryByMonth = Object.fromEntries(
    Object.entries(monthQuestionCategoryCounts).map(([monthId, counts]) => {
      const total = [...counts.values()].reduce((sum, value) => sum + value, 0) || 1;
      return [
        monthId,
        [...counts.entries()]
          .map(([name, value]) => [name, value, Math.round((value / total) * 100)])
          .sort((a, b) => b[1] - a[1])
      ];
    })
  );

  const brandSummaryByMonth = Object.fromEntries(
    Object.entries(monthBrandCounts).map(([monthId, counts]) => [
      monthId,
      [...counts.entries()]
        .map(([name, value]) => [name, value])
        .sort((a, b) => b[1] - a[1])
    ])
  );

  const categoryByAssetByMonth = Object.fromEntries(
    Object.entries(monthAssetCategoryCounts).map(([monthId, assets]) => [
      monthId,
      Object.fromEntries(
        Object.entries(assets).map(([assetId, counts]) => [
          assetId,
          [...counts.entries()]
            .map(([name, value]) => [name, value])
            .sort((a, b) => b[1] - a[1])
        ])
      )
    ])
  );

  const brandByAssetByMonth = Object.fromEntries(
    Object.entries(monthAssetBrandCounts).map(([monthId, assets]) => [
      monthId,
      Object.fromEntries(
        Object.entries(assets).map(([assetId, counts]) => [
          assetId,
          [...counts.entries()]
            .map(([name, value]) => [name, value])
            .sort((a, b) => b[1] - a[1])
        ])
      )
    ])
  );

  const serializedDailyMetrics = Object.fromEntries(
    Object.entries(dailyAssetMetrics).map(([date, assets]) => [
      date,
      Object.fromEntries(
        Object.entries(assets).map(([assetId, bucket]) => {
          const rating = bucket.ratings.length
            ? Number((bucket.ratings.reduce((sum, value) => sum + value, 0) / bucket.ratings.length).toFixed(2))
            : undefined;
          const row = [
            bucket.sessionIds.size,
            bucket.questions,
            bucket.voice,
            bucket.quickReply,
            bucket.keyboard,
            bucket.thai,
            bucket.english,
            bucket.chinese,
            bucket.others,
            bucket.success,
            bucket.apology
          ];
          if (rating !== undefined) row.push(rating);
          return [assetId, row];
        })
      )
    ])
  );

  for (const [monthId, file] of Object.entries(legacyEventFiles)) {
    const messageRows = parseTabularFile(join(rootDir, file));
    for (const row of messageRows) {
      const start = normalizeDate(row.startAt);
      const localDate = formatLocalIsoDate(start);
      allMessageRows.push({
        assetId: csvMachineIdMap[row.machineId] || null,
        machineId: String(row.machineId || ""),
        sessionId: String(row.sessionId || ""),
        localDate,
        monthId: localDate?.slice(0, 7) || monthId
      });
    }
  }

  const eventHistory = eventCatalog.map((event, index) => {
    const startDate = event.setup;
    const endDate = event.returnDate || event.setup;
    const machineIdSet = event.machineIds ? new Set(event.machineIds) : null;
    const messageRows = allMessageRows.filter(row =>
      (machineIdSet ? machineIdSet.has(row.machineId) : row.assetId === event.asset) &&
      row.localDate &&
      row.localDate >= startDate &&
      row.localDate <= endDate
    );
    const ratingRows = allRatingRows.filter(row =>
      (machineIdSet ? machineIdSet.has(row.machineId) : row.assetId === event.asset) &&
      row.localDate &&
      row.localDate >= startDate &&
      row.localDate <= endDate &&
      Number.isFinite(row.rating)
    );
    const sessionIds = new Set(messageRows.map(row => row.sessionId).filter(Boolean));
    const sessions = sessionIds.size;
    const questions = messageRows.length;
    const rating = ratingRows.length
      ? Number((ratingRows.reduce((sum, row) => sum + row.rating, 0) / ratingRows.length).toFixed(1))
      : null;

    return {
      id: index,
      month: startDate.slice(0, 7),
      event: event.event,
      date: formatEventDateRange(startDate, endDate),
      asset: event.asset,
      original: event.original,
      location: event.location,
      setup: startDate,
      returnDate: endDate,
      performance: formatEventPerformance({ sessions, questions, rating }),
      notes: `Performance calculated automatically from the Performance Reports CSV for ${startDate}${startDate === endDate ? "" : ` to ${endDate}`}.`
    };
  });

  const latestMonthId = Object.keys(reportFiles).sort().at(-1);
  const latestHistoryEnd = latestMonthId ? monthMeta[latestMonthId]?.coverageEnd : null;

  return {
    generatedAt: new Date().toISOString(),
    source: "Performance Reports CSV",
    historyStart: "1 Sep 2025",
    historyEnd: latestHistoryEnd ? formatLocalHistoryDate(latestHistoryEnd) : "30 Jun 2026",
    monthMeta,
    csvMachineMetrics,
    rawMonthlyMetrics,
    dailyAssetMetrics: serializedDailyMetrics,
    topQuestionsByMonth,
    topQuestionsByAssetByMonth,
    hourlyByAssetByMonth,
    weekdayByAssetByMonth,
    topicSummaryByMonth,
    brandSummaryByMonth,
    categoryByAssetByMonth,
    brandByAssetByMonth,
    eventHistory,
    accuracyReview: buildAccuracyReviewSummary(accuracyRows, accuracySessionIds),
    kmGapAnalysis: buildKmGapAnalysis(kmBrandMentions, actionBuckets, customerDemand, {
      dates: kmCoverageDates,
      flagMonths: [...kmFlagMonths].sort(),
      fallbackMonths: [...kmFallbackMonths].sort(),
      directorySize: kmDirectory.length
    })
  };
}

// "owner" is the kind of work the data needs, not a team name: "tune" = the content exists
// (or the question never arrived intact) so the model/recognition needs tuning; "add" = the
// content is genuinely missing from KM; "none" = correct behaviour, nothing to change.
const ACTION_META = {
  vendorSpeech: { owner: "tune", tone: "rose" },
  kmAddBrand: { owner: "add", tone: "amber" },
  kmVerifyBrand: { owner: "tune", tone: "purple" },
  kmAddTopic: { owner: "add", tone: "cyan" },
  quickReplyDeadEnd: { owner: "add", tone: "rose" },
  outOfScope: { owner: "none", tone: "" }
};

function serializeActionBucket(bucket) {
  const topEntries = map => [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([name, count]) => ({ name, count }));
  return {
    id: bucket.id,
    owner: ACTION_META[bucket.id]?.owner || "km",
    tone: ACTION_META[bucket.id]?.tone || "",
    count: bucket.count,
    voice: bucket.voice,
    text: bucket.text,
    quickReply: bucket.quickReply,
    topics: topEntries(bucket.topics),
    languages: topEntries(bucket.languages),
    examples: bucket.examples
  };
}

// Ranks brand mentions (from the messageLog's own mentionedstorename/isMentionedBrandsInKM
// columns, already computed upstream — not re-derived here) into three views: brands the AI
// got wrong despite KM having the answer, brands customers asked about with no KM entry at
// all, and the overall most-asked-about brands.
function buildKmGapAnalysis(brandMentions, actionBuckets, customerDemand, coverage = {}) {
  const all = [...brandMentions.values()];
  const totalBrandMentions = all.reduce((sum, b) => sum + b.count, 0);
  const wrongAnswers = all
    .filter(b => b.apologyInKMCount > 0)
    .sort((a, b) => b.apologyInKMCount - a.apologyInKMCount)
    .slice(0, 40);
  // "Not in KM" alone isn't enough to call it a gap — brands like M Card or Power Mall are
  // handled by quick replies instead of the store Directory, so the AI still answers fine.
  // Require an actual apology so this only surfaces real "no data" failures.
  const gaps = all
    .filter(b => b.apologyNotInKMCount > 0)
    .sort((a, b) => b.apologyNotInKMCount - a.apologyNotInKMCount)
    .slice(0, 40);
  const topMentions = all
    .slice()
    .sort((a, b) => b.count - a.count)
    .slice(0, 40);

  // Biggest actionable bucket first; the "no action needed" bucket always sinks to the
  // bottom regardless of size, since it is context rather than work.
  const actions = [...actionBuckets.values()]
    .map(serializeActionBucket)
    .sort((a, b) => (a.owner === "none") - (b.owner === "none") || b.count - a.count);
  // Only topics with enough volume to act on — a 1-question topic with a 100% failure rate
  // is noise, not a priority.
  const demand = [...customerDemand.values()]
    .filter(d => d.count >= 5)
    .map(({ questionCounts, ...d }) => ({
      ...d,
      apologyRate: d.count ? (d.apology / d.count) * 100 : 0,
      topQuestions: [...questionCounts.values()].sort((a, b) => b.count - a.count).slice(0, 10)
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 25);

  const dates = (coverage.dates || []).filter(Boolean).sort();
  const period = {
    start: dates[0] || null,
    end: dates[dates.length - 1] || null,
    questionsScanned: dates.length,
    monthsScanned: (coverage.flagMonths || []).concat(coverage.fallbackMonths || [])
      .filter((v, i, a) => a.indexOf(v) === i).sort(),
    // Months where the export carried its own in-KM flag vs. months where we had to decide
    // by matching the KM store directory instead.
    flagMonths: coverage.flagMonths || [],
    fallbackMonths: coverage.fallbackMonths || [],
    directorySize: coverage.directorySize || 0
  };

  return { totalBrandMentions, totalUniqueBrands: all.length, wrongAnswers, gaps, topMentions, actions, demand, period };
}

function buildAccuracyReviewSummary(records, sessionIds) {
  if (!records || !records.length) return null;
  const total = records.length;
  const rate = predicate => (total ? (records.filter(predicate).length / total) * 100 : 0);
  const tally = key => {
    const map = {};
    records.forEach(record => {
      const value = record[key] || "Unknown";
      map[value] = (map[value] || 0) + 1;
    });
    return map;
  };
  const categoryMap = new Map();
  records.forEach(record => {
    const bucket = categoryMap.get(record.category) || { category: record.category, count: 0, apology: 0 };
    bucket.count += 1;
    if (record.apology) bucket.apology += 1;
    categoryMap.set(record.category, bucket);
  });
  const dates = records.map(r => r.time.slice(0, 10)).filter(d => /^\d{4}/.test(d)).sort();

  // The browsable "What customers asked" list ships question text for the full history
  // (every month), matching the rest of the dashboard's full-history default.
  const visibleRows = records.map(({ monthId, ...rest }) => rest);
  const visibleDates = visibleRows.map(r => r.time.slice(0, 10)).filter(d => /^\d{4}/.test(d)).sort();

  return {
    label: "Full-history question review",
    total,
    sessions: sessionIds ? sessionIds.size : new Set(records.map(r => r.time)).size,
    dateStart: dates[0] || null,
    dateEnd: dates[dates.length - 1] || null,
    apologyRate: rate(r => r.apology),
    transcriptionRate: rate(r => r.category === "TranscriptionError"),
    outOfScopeRate: rate(r => r.category === "Out of Scope"),
    cleanRate: rate(r => !r.apology && r.category !== "TranscriptionError" && r.category !== "Out of Scope"),
    inputMix: tally("input"),
    langMix: tally("language"),
    categories: [...categoryMap.values()].sort((a, b) => b.count - a.count),
    rowsDateStart: visibleDates[0] || null,
    rowsDateEnd: visibleDates[visibleDates.length - 1] || null,
    rows: visibleRows
  };
}

function formatLocalHistoryDate(value) {
  const date = value ? new Date(`${value}T00:00:00`) : null;
  if (!date || Number.isNaN(+date)) return null;
  return `${date.getDate()} ${date.toLocaleString("en-US", { month: "short" })} ${date.getFullYear()}`;
}

export function writePerformanceReportData(rootDir = process.cwd()) {
  const payload = buildPerformanceReportData(rootDir);
  const outputPath = join(rootDir, "outputs", "assets", "performance-report-data.js");
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(
    outputPath,
    `window.__PERFORMANCE_REPORT_DATA__ = ${JSON.stringify(payload, null, 2)};\n`,
    "utf8"
  );
  return outputPath;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const out = writePerformanceReportData(process.cwd());
  console.log(`Generated ${out}`);
}
