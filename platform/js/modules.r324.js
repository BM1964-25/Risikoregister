import { initialState } from "./state.r321.js";

export const riskCategoryOptions = [
  "Markt- und Wirtschaftsrisiken",
  "Standort- und Umweltrisiken",
  "Rechtliche und Genehmigungsrisiken",
  "Finanzielle Risiken",
  "Projekt- und Managementrisiken",
  "Technische und Planungsrisiken",
  "Stakeholder- und Kommunikationsrisiken",
  "Ausführungs- und Baurisiken",
  "Lieferketten- und Beschaffungsrisiken",
  "Sicherheits- und Gesundheitsrisiken",
  "IT- und Datenrisiken",
  "Externe Risiken und höhere Gewalt",
  "Sonstige / Individuell"
];

export const riskPhaseOptions = [
  "Projektinitiierung und Entwicklung",
  "Planung",
  "Genehmigung",
  "Ausschreibung und Vergabe",
  "Beschaffung",
  "Ausführung",
  "Inbetriebnahme und Abnahme",
  "Betrieb und Gewährleistung"
];

export function normalizeRiskPhaseValue(value) {
  const current = String(value || "").trim();
  const normalized = current.toLowerCase();
  if (!normalized) return riskPhaseOptions[0];
  const matched = riskPhaseOptions.find((option) => option.toLowerCase() === normalized);
  return matched || current;
}

function buildRiskPhaseOptions(currentValue) {
  const current = normalizeRiskPhaseValue(currentValue);
  return [...new Set([current, ...riskPhaseOptions])];
}

export function normalizeRiskStatusValue(value) {
  const current = String(value || "").trim().toLowerCase();
  if (!current) return "offen";
  if (current === "beobachtung" || current === "überwachung") return "überwachung";
  if (current === "in beurteilung" || current === "in bewertung") return "in bewertung";
  if (current === "in bearbeitung") return "in bearbeitung";
  if (current === "maßnahme läuft" || current === "massnahme läuft" || current === "massnahme laeuft") return "in bearbeitung";
  if (current === "geschlossen") return "geschlossen";
  return "offen";
}

export function deriveRiskLikelihoodFromPercent(probabilityPercent, fallback = 1) {
  const percent = Number(probabilityPercent);
  if (!Number.isFinite(percent)) {
    return Math.max(1, Math.min(5, Number(fallback) || 1));
  }
  if (percent <= 0) return 1;
  if (percent <= 20) return 1;
  if (percent <= 40) return 2;
  if (percent <= 60) return 3;
  if (percent <= 80) return 4;
  return 5;
}

export function normalizeRiskCategoryValue(value) {
  const current = String(value || "").trim().toLowerCase();
  if (!current) return "Projekt- und Managementrisiken";
  if (current === "operativ") return "Projekt- und Managementrisiken";
  if (current === "technisch") return "Technische und Planungsrisiken";
  if (current === "finanziell") return "Finanzielle Risiken";
  if (current === "rechtlich") return "Rechtliche und Genehmigungsrisiken";
  if (current === "standort") return "Standort- und Umweltrisiken";
  if (current === "lieferkette" || current === "beschaffung") return "Lieferketten- und Beschaffungsrisiken";
  if (current === "sicherheit") return "Sicherheits- und Gesundheitsrisiken";
  if (current === "it" || current === "daten") return "IT- und Datenrisiken";
  if (current === "externe" || current === "höhere gewalt" || current === "hoehere gewalt") return "Externe Risiken und höhere Gewalt";
  return riskCategoryOptions.some((option) => option.toLowerCase() === current)
    ? value
    : "Sonstige / Individuell";
}

function buildRiskCategoryOptions(currentValue) {
  const current = normalizeRiskCategoryValue(currentValue);
  return [...new Set([current, ...riskCategoryOptions])];
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatCurrency(value) {
  return new Intl.NumberFormat("de-DE", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 0
  }).format(value || 0);
}

function formatPercent(value) {
  return new Intl.NumberFormat("de-DE", {
    style: "percent",
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  }).format(value || 0);
}

function formatNumber(value) {
  return new Intl.NumberFormat("de-DE", {
    maximumFractionDigits: 0
  }).format(value || 0);
}

function formatDecimalInput(value, fractionDigits = 2) {
  return new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits
  }).format(Number(value) || 0);
}

function formatCurrencyInput(value) {
  return `${formatNumber(value)} €`;
}

function formatCompactCurrency(value) {
  const absolute = Math.abs(Number(value) || 0);
  if (absolute >= 1000000) {
    return `${new Intl.NumberFormat("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format((Number(value) || 0) / 1000000)}Mio.€`;
  }
  if (absolute >= 1000) {
    return `${new Intl.NumberFormat("de-DE", { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format((Number(value) || 0) / 1000)}T€`;
  }
  return formatCurrency(value);
}

function formatSignedNumber(value, fractionDigits = 1) {
  const amount = Number(value) || 0;
  const sign = amount > 0 ? "+" : amount < 0 ? "−" : "";
  return `${sign}${new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits
  }).format(Math.abs(amount))}`;
}

function formatSignedCurrency(value) {
  const amount = Number(value) || 0;
  return `${amount > 0 ? "+" : amount < 0 ? "−" : ""}${formatCurrency(Math.abs(amount))}`;
}

function formatSignedCompactCurrency(value) {
  const amount = Number(value) || 0;
  return `${amount > 0 ? "+" : amount < 0 ? "−" : ""}${formatCompactCurrency(Math.abs(amount))}`;
}

function formatPercentValue(value) {
  return new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  }).format((Number(value) || 0) * 100);
}

function formatDate(value) {
  if (!value) return "nicht angegeben";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString("de-DE");
}

function splitTitleIntoTwoLines(value) {
  const title = String(value || "").trim();
  const words = title.split(/\s+/).filter(Boolean);
  if (words.length < 5 && title.length < 42) {
    return [title, ""];
  }
  let bestIndex = 1;
  let bestDiff = Infinity;
  for (let i = 1; i < words.length; i += 1) {
    const left = words.slice(0, i).join(" ");
    const right = words.slice(i).join(" ");
    const diff = Math.abs(left.length - right.length);
    if (diff < bestDiff) {
      bestDiff = diff;
      bestIndex = i;
    }
  }
  return [words.slice(0, bestIndex).join(" "), words.slice(bestIndex).join(" ")];
}

function parseRiskDateValue(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date.getTime();
}

function clampEvaTolerancePercent(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 3;
  return Math.max(0, Math.min(10, numeric));
}

function calculateProjectTimeContext(project) {
  const start = project?.startDate ? new Date(project.startDate) : null;
  const end = project?.endDate ? new Date(project.endDate) : null;
  const analysis = project?.analysisDate ? new Date(project.analysisDate) : new Date();

  const hasValidDates = [start, end, analysis].every((date) => date instanceof Date && !Number.isNaN(date.getTime()));
  if (!hasValidDates || !start || !end) {
    return {
      status: "Zeitkontext unvollständig",
      detail: "Für eine belastbare Einordnung fehlen Bau- oder Stichtagsdaten.",
      progressLabel: "nicht ableitbar"
    };
  }

  const totalDuration = end.getTime() - start.getTime();
  if (totalDuration <= 0) {
    return {
      status: "Zeitlogik prüfen",
      detail: "Die geplante Fertigstellung liegt nicht nach dem Baubeginn.",
      progressLabel: "nicht ableitbar"
    };
  }

  const elapsed = analysis.getTime() - start.getTime();
  const progressRatio = Math.min(Math.max(elapsed / totalDuration, 0), 1);
  const progressPercent = Math.round(progressRatio * 100);

  let scheduleStatus = "Vor Baubeginn";
  if (analysis > end) {
    scheduleStatus = "Nach Soll-Termin";
  } else if (analysis >= start) {
    if (progressRatio >= 0.9) scheduleStatus = "Kurz vor Übergabe";
    else if (progressRatio >= 0.55) scheduleStatus = "In Ausführung";
    else scheduleStatus = "Frühe Ausführung";
  }

  return {
    status: scheduleStatus,
    detail: `Betrachtungszeitpunkt ${formatDate(project.analysisDate)} · kalendarischer Projektfortschritt ca. ${progressPercent} %`,
    progressLabel: `${progressPercent} % der geplanten Laufzeit`
  };
}

function getProjectTimeTone(status) {
  const value = String(status || "").toLowerCase();
  if (value.includes("nach soll-termin")) return "critical";
  if (value.includes("kurz vor übergabe")) return "warn";
  if (value.includes("in ausführung") || value.includes("frühe ausführung")) return "active";
  return "neutral";
}

function normalizePertSecurityLabel(label) {
  if (!label) return "";
  return String(label)
    .replace("84 % · Budget+", "84 % · Vorsichtiges Niveau")
    .replace("97,5 % · Konservativ", "97,5 % · Konservatives Niveau");
}

function normalizePertSnapshotLabel(label) {
  return normalizePertSecurityLabel(label);
}

function formatIntegerInput(value, suffix = "") {
  const formatted = formatNumber(Number(value) || 0);
  return suffix ? `${formatted} ${suffix}` : formatted;
}

function mean(values) {
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function percentile(values, ratio) {
  if (!values.length) return 0;
  const ordered = [...values].sort((a, b) => a - b);
  const index = (ordered.length - 1) * ratio;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  if (lower === upper) return ordered[lower];
  const weight = index - lower;
  return ordered[lower] * (1 - weight) + ordered[upper] * weight;
}

function createRng(seed) {
  let state = (Number(seed) || 1) >>> 0;
  return function rng() {
    state = (1664525 * state + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

function sampleNormal(rng) {
  let u = 0;
  let v = 0;
  while (u === 0) u = rng();
  while (v === 0) v = rng();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
}

function sampleGamma(alpha, rng) {
  if (alpha < 1) {
    return sampleGamma(alpha + 1, rng) * Math.pow(rng(), 1 / alpha);
  }
  const d = alpha - 1 / 3;
  const c = 1 / Math.sqrt(9 * d);
  while (true) {
    const x = sampleNormal(rng);
    const v = Math.pow(1 + c * x, 3);
    if (v <= 0) continue;
    const u = rng();
    if (u < 1 - 0.0331 * Math.pow(x, 4)) return d * v;
    if (Math.log(u) < 0.5 * x * x + d * (1 - v + Math.log(v))) return d * v;
  }
}

function sampleBeta(alpha, beta, rng) {
  const x = sampleGamma(alpha, rng);
  const y = sampleGamma(beta, rng);
  return x / (x + y);
}

function samplePert(lowValue, modeValue, highValue, rng) {
  const low = Math.min(lowValue, modeValue, highValue);
  const high = Math.max(lowValue, modeValue, highValue);
  const mode = Math.min(high, Math.max(low, modeValue));
  if (low === high) return low;
  const expected = (low + 4 * mode + high) / 6;
  const alpha = 1 + (4 * (expected - low)) / (high - low);
  const beta = 1 + (4 * (high - expected)) / (high - low);
  return low + sampleBeta(alpha, beta, rng) * (high - low);
}

export function calculatePertResult(state) {
  const items = state.pert.items || [];
  const enriched = items.map((item) => {
    const optimistic = Number(item.optimistic) || 0;
    const mostLikely = Number(item.mostLikely) || 0;
    const pessimistic = Number(item.pessimistic) || 0;
    const expectedValue = (optimistic + 4 * mostLikely + pessimistic) / 6;
    const standardDeviation = (pessimistic - optimistic) / 6;
    return {
      ...item,
      optimistic,
      mostLikely,
      pessimistic,
      expectedValue,
      standardDeviation,
      variance: standardDeviation * standardDeviation
    };
  });

  const expectedValue = enriched.reduce((sum, item) => sum + item.expectedValue, 0);
  const variance = enriched.reduce((sum, item) => sum + item.variance, 0);
  const standardDeviation = Math.sqrt(variance);
  const selectedBudget = Number(state.pert.securityLevel) === 2
    ? expectedValue + 2 * standardDeviation
    : Number(state.pert.securityLevel) === 0
      ? expectedValue
      : expectedValue + standardDeviation;
  const projectBudget = Number(state.project?.budget) || 0;
  const budgetDelta = projectBudget - selectedBudget;

  return {
    items: enriched,
    optimisticTotal: enriched.reduce((sum, item) => sum + item.optimistic, 0),
    mostLikelyTotal: enriched.reduce((sum, item) => sum + item.mostLikely, 0),
    pessimisticTotal: enriched.reduce((sum, item) => sum + item.pessimistic, 0),
    expectedValue,
    standardDeviation,
    budget84: expectedValue + standardDeviation,
    budget975: expectedValue + 2 * standardDeviation,
    selectedBudget,
    budgetDelta
  };
}

function buildPertPhaseRows(state, result) {
  const phaseLabels = ["LPH 1", "LPH 2", "LPH 3", "LPH 4", "LPH 5", "LPH 6", "LPH 7", "LPH 8"];
  const snapshotByPhase = new Map();
  (state.pert.snapshots || []).forEach((snapshot) => {
    if (snapshot?.phase && !snapshotByPhase.has(snapshot.phase)) {
      snapshotByPhase.set(snapshot.phase, snapshot);
    }
  });

  return phaseLabels.map((label, index) => {
    const snapshot = snapshotByPhase.get(label);
    const isCurrentPhase = state.project?.phase === label;
    const optimistic = isCurrentPhase
      ? result.optimisticTotal
      : Number(snapshot?.optimisticTotal) || 0;
    const expectedValue = isCurrentPhase
      ? result.expectedValue
      : Number(snapshot?.expectedValue) || 0;
    const pessimistic = isCurrentPhase
      ? result.pessimisticTotal
      : Number(snapshot?.pessimisticTotal) || 0;
    const previous = phaseLabels
      .slice(0, index)
      .reverse()
      .map((phaseLabel) => {
        const prevSnapshot = snapshotByPhase.get(phaseLabel);
        if (state.project?.phase === phaseLabel) {
          return { expectedValue: result.expectedValue };
        }
        if (prevSnapshot?.expectedValue) {
          return prevSnapshot;
        }
        return null;
      })
      .find(Boolean);
    return {
      label,
      optimistic,
      expectedValue,
      pessimistic,
      comment: state.pert?.lphComments?.[label] || "",
      snapshotLabel: isCurrentPhase
        ? "Aktueller Arbeitsstand"
        : snapshot?.savedAt || "Kein Snapshot",
      source: isCurrentPhase ? "current" : snapshot ? "snapshot" : "empty",
      deltaExpected: previous ? expectedValue - (Number(previous.expectedValue) || 0) : 0
    };
  });
}

function getLatestPertSnapshotsByPhase(state) {
  const latestByPhase = new Map();
  (state.pert.snapshots || []).forEach((snapshot) => {
    if (snapshot?.phase && !latestByPhase.has(snapshot.phase)) {
      latestByPhase.set(snapshot.phase, snapshot);
    }
  });

  return [...latestByPhase.values()]
    .map((snapshot) => ({
      ...snapshot,
      phaseIndex: getPhaseIndex(snapshot.phase)
    }))
    .filter((snapshot) => snapshot.phaseIndex > 0)
    .sort((a, b) => a.phaseIndex - b.phaseIndex);
}

function getPhaseIndex(label) {
  const match = String(label || "").match(/LPH\s+(\d+)/i);
  return match ? Number(match[1]) : 0;
}

function buildPertPhaseCorridorRows(state, result) {
  const phaseLabels = ["LPH 1", "LPH 2", "LPH 3", "LPH 4", "LPH 5", "LPH 6", "LPH 7", "LPH 8"];
  const baseExpected = result.expectedValue || 0;
  const currentPhaseIndex = getPhaseIndex(state.project?.phase) || 1;
  const liveRows = buildPertPhaseRows(state, result);
  const phaseUncertainty = {
    1: 0.30,
    2: 0.25,
    3: 0.20,
    4: 0.16,
    5: 0.12,
    6: 0.09,
    7: 0.07,
    8: 0.05
  };

  return phaseLabels.map((label, index) => {
    const phaseNumber = index + 1;
    const uncertainty = phaseUncertainty[phaseNumber] || 0.05;
    const liveRow = liveRows[index];
    const expectedValue = liveRow?.expectedValue > 0 ? liveRow.expectedValue : baseExpected;
    return {
      label,
      expectedValue,
      upper: expectedValue * (1 + uncertainty),
      lower: expectedValue * (1 - uncertainty),
      isCurrent: phaseNumber === currentPhaseIndex,
      isReached: phaseNumber <= currentPhaseIndex
    };
  });
}

function createNormalDistributionSvg(meanValue, standardDeviation) {
  const width = 572;
  const height = 248;
  const padding = 38;
  const sigma = standardDeviation > 0 ? standardDeviation : Math.max(meanValue * 0.05, 1);
  const minX = meanValue - 3 * sigma;
  const maxX = meanValue + 3 * sigma;
  const points = [];
  let maxY = 0;

  for (let index = 0; index <= 96; index += 1) {
    const ratio = index / 96;
    const xValue = minX + (maxX - minX) * ratio;
    const exponent = -0.5 * Math.pow((xValue - meanValue) / sigma, 2);
    const yValue = Math.exp(exponent);
    maxY = Math.max(maxY, yValue);
    points.push({ xValue, yValue });
  }

  const path = points.map((point, index) => {
    const x = padding + ((point.xValue - minX) / (maxX - minX)) * (width - padding * 2);
    const y = height - padding - (point.yValue / maxY) * (height - padding * 2);
    return `${index === 0 ? "M" : "L"}${x.toFixed(2)} ${y.toFixed(2)}`;
  }).join(" ");

  const axisY = height - padding;
  const meanX = padding + ((meanValue - minX) / (maxX - minX)) * (width - padding * 2);
  const p84Value = meanValue + sigma;
  const p975Value = meanValue + 2 * sigma;
  const p16Value = meanValue - sigma;
  const p25Value = meanValue - 2 * sigma;
  const p16X = padding + ((p16Value - minX) / (maxX - minX)) * (width - padding * 2);
  const p25X = padding + ((p25Value - minX) / (maxX - minX)) * (width - padding * 2);
  const p84X = padding + ((p84Value - minX) / (maxX - minX)) * (width - padding * 2);
  const p975X = padding + ((p975Value - minX) / (maxX - minX)) * (width - padding * 2);
  const tickValues = [minX, p25Value, p16Value, meanValue, p84Value, p975Value, maxX];
  const gridLines = [0.25, 0.5, 0.75].map((ratio) => height - padding - ratio * (height - padding * 2));
  const peakY = height - padding - (Math.exp(-0.5 * Math.pow((meanValue - meanValue) / sigma, 2)) / maxY) * (height - padding * 2);
  const markerLineTop = padding - 10;
  const percentLabelY = markerLineTop - 10;

  const svg = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Normalverteilung">
      <defs>
        <linearGradient id="pertCurveFill" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="rgba(17,55,92,0.10)"></stop>
          <stop offset="48%" stop-color="rgba(17,55,92,0.042)"></stop>
          <stop offset="100%" stop-color="rgba(17,55,92,0.006)"></stop>
        </linearGradient>
      </defs>
      ${gridLines.map((y) => `
        <line x1="${padding}" y1="${y}" x2="${width - padding}" y2="${y}" stroke="#e8eef3" stroke-width="0.72"></line>
      `).join("")}
      <line x1="${padding}" y1="${axisY}" x2="${width - padding}" y2="${axisY}" stroke="#bcc8d1" stroke-width="0.88"></line>
      <line x1="${padding}" y1="${padding}" x2="${padding}" y2="${axisY}" stroke="#e0e8ee" stroke-width="0.75"></line>
      <path d="${path} L ${width - padding} ${axisY} L ${padding} ${axisY} Z" fill="url(#pertCurveFill)" stroke="none"></path>
      <path d="${path}" fill="none" stroke="#143a61" stroke-width="1.28" stroke-linecap="round" stroke-linejoin="round"></path>
      <line x1="${p25X}" y1="${markerLineTop}" x2="${p25X}" y2="${axisY}" stroke="#b46d3b" stroke-dasharray="2.6 6.2" stroke-width="0.9"></line>
      <line x1="${p16X}" y1="${markerLineTop}" x2="${p16X}" y2="${axisY}" stroke="#c47845" stroke-dasharray="2.6 6.2" stroke-width="0.9"></line>
      <line x1="${meanX}" y1="${markerLineTop}" x2="${meanX}" y2="${axisY}" stroke="#143a61" stroke-dasharray="3.1 6.1" stroke-width="1"></line>
      <line x1="${p84X}" y1="${markerLineTop}" x2="${p84X}" y2="${axisY}" stroke="#987431" stroke-dasharray="2.6 6.2" stroke-width="0.9"></line>
      <line x1="${p975X}" y1="${markerLineTop}" x2="${p975X}" y2="${axisY}" stroke="#1f6b5e" stroke-dasharray="2.6 6.2" stroke-width="0.9"></line>
      <circle cx="${meanX}" cy="${peakY}" r="2.7" fill="#143a61"></circle>
      <text x="${p25X}" y="${percentLabelY}" text-anchor="middle" fill="#b46d3b" font-size="10" font-weight="620">2,5 %</text>
      <text x="${p16X}" y="${percentLabelY}" text-anchor="middle" fill="#c47845" font-size="10" font-weight="620">16 %</text>
      <text x="${meanX}" y="${percentLabelY}" text-anchor="middle" fill="#143a61" font-size="10.1" font-weight="680">EW</text>
      <text x="${p84X}" y="${percentLabelY}" text-anchor="middle" fill="#987431" font-size="10" font-weight="620">84 %</text>
      <text x="${p975X}" y="${percentLabelY}" text-anchor="middle" fill="#1f6b5e" font-size="10" font-weight="620">97,5 %</text>
      ${tickValues.map((tick) => {
        const x = padding + ((tick - minX) / (maxX - minX)) * (width - padding * 2);
        return `
          <line x1="${x}" y1="${axisY}" x2="${x}" y2="${axisY + 4.2}" stroke="#9eafbb" stroke-width="0.72"></line>
          <text x="${x}" y="${axisY + 18}" text-anchor="middle" fill="#667686" font-size="6.8">
            <tspan>${formatNumber(tick)}</tspan>
            <tspan font-size="4.8"> €</tspan>
          </text>
        `;
      }).join("")}
    </svg>
  `;

  return {
    svg,
    positions: {
      p25X,
      p16X,
      meanX,
      p84X,
      p975X,
      width
    },
    markers: [
      { label: "EW − 2 SD", sublabel: "2,5 %-Niveau", value: p25Value, color: "#b2662f" },
      { label: "EW − 1 SD", sublabel: "16 %-Niveau", value: p16Value, color: "#c06f34" },
      { label: "EW", value: meanValue, color: "#163f63" },
      { label: "EW + 1 SD", sublabel: "84 %-Niveau", value: p84Value, color: "#1f7a53" },
      { label: "EW + 2 SD", sublabel: "97,5 %-Niveau", value: p975Value, color: "#c78a2d" }
    ]
  };
}

function createPertBarChartSvg(rows, valueKey, color, formatter = (value) => formatNumber(value), options = {}) {
  const items = rows;
  const width = options.width || 400;
  const barHeight = options.barHeight || 18;
  const gap = options.gap || 12;
  const height = 42 + items.length * (barHeight + gap);
  const left = options.left || 118;
  const right = options.right || 112;
  const maxValue = Math.max(...items.map((item) => Number(item[valueKey]) || 0), 1);
  const valueColumnX = width - 6;
  const labelFontSize = options.labelFontSize || 11.5;
  const valueFontSize = options.valueFontSize || 11.5;
  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="PERT Balkendiagramm">
      <defs>
        <linearGradient id="pertBarGradient-${color.replace(/[^a-zA-Z0-9]/g, "")}" x1="0" y1="0" x2="1" y2="0">
          <stop offset="0%" stop-color="${color}"></stop>
          <stop offset="100%" stop-color="${color}CC"></stop>
        </linearGradient>
      </defs>
      ${items.map((item, index) => {
        const y = 16 + index * (barHeight + gap);
        const value = Number(item[valueKey]) || 0;
        const barWidth = ((width - left - right) * value) / maxValue;
        return `
          <text x="0" y="${y + 13}" fill="#304252" font-size="${labelFontSize}" font-weight="700">${item.name}</text>
          <rect x="${left}" y="${y}" width="${width - left - right}" height="${barHeight}" rx="10" fill="#e4ebf1"></rect>
          <rect x="${left}" y="${y}" width="${barWidth}" height="${barHeight}" rx="10" fill="url(#pertBarGradient-${color.replace(/[^a-zA-Z0-9]/g, "")})"></rect>
          <text x="${valueColumnX}" y="${y + 13}" text-anchor="end" fill="#163f63" font-size="${valueFontSize}" font-weight="700">${formatter(value)}</text>
        `;
      }).join("")}
    </svg>
  `;
}

function createMonteCarloDistributionSvg(values, markers, options = {}) {
  const source = (values || []).filter((value) => Number.isFinite(value));
  const width = 520;
  const height = 220;
  const padding = { top: 24, right: 22, bottom: 34, left: 40 };
  if (!source.length) {
    return `
      <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Monte-Carlo-Verteilung">
        <text x="${width / 2}" y="${height / 2}" text-anchor="middle" fill="#667789" font-size="14">Keine Simulationsdaten vorhanden</text>
      </svg>
    `;
  }

  const min = Math.min(...source);
  const max = Math.max(...source);
  const bucketCount = Math.max(8, Math.min(16, Math.round(Math.sqrt(source.length) / 2)));
  const range = Math.max(max - min, 1);
  const bucketSize = range / bucketCount;
  const buckets = Array.from({ length: bucketCount }, (_, index) => ({
    start: min + index * bucketSize,
    end: index === bucketCount - 1 ? max : min + (index + 1) * bucketSize,
    count: 0
  }));

  source.forEach((value) => {
    const rawIndex = Math.floor((value - min) / bucketSize);
    const index = Math.min(bucketCount - 1, Math.max(0, rawIndex));
    buckets[index].count += 1;
  });

  const maxCount = Math.max(...buckets.map((bucket) => bucket.count), 1);
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;
  const barGap = 4;
  const barWidth = (chartWidth - barGap * (bucketCount - 1)) / bucketCount;

  const markerLine = (value, label, color) => {
    const x = padding.left + ((value - min) / range) * chartWidth;
    return `
      <line x1="${x}" y1="${padding.top - 4}" x2="${x}" y2="${height - padding.bottom}" stroke="${color}" stroke-width="1.3" stroke-dasharray="4 5"></line>
      <text x="${x}" y="${padding.top - 8}" text-anchor="middle" fill="${color}" font-size="11" font-weight="700">${label}</text>
    `;
  };

  const formatter = options.formatter || ((value) => formatNumber(value));
  const xTicks = [min, markers.p50, markers.p80, max];

  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="${options.ariaLabel || "Monte-Carlo-Verteilung"}">
      ${[0.25, 0.5, 0.75].map((ratio) => {
        const y = padding.top + chartHeight - ratio * chartHeight;
        return `<line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#e3ebf1" stroke-width="1"></line>`;
      }).join("")}
      <line x1="${padding.left}" y1="${height - padding.bottom}" x2="${width - padding.right}" y2="${height - padding.bottom}" stroke="#bcc8d1" stroke-width="1"></line>
      ${buckets.map((bucket, index) => {
        const x = padding.left + index * (barWidth + barGap);
        const barHeight = (bucket.count / maxCount) * chartHeight;
        const y = padding.top + chartHeight - barHeight;
        return `
          <rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" rx="6" fill="${options.barColor || "#1e4f7b"}" opacity="${0.32 + (bucket.count / maxCount) * 0.58}"></rect>
        `;
      }).join("")}
      ${markerLine(markers.p50, "P50", "#163f63")}
      ${markerLine(markers.p80, "P80", "#8a6831")}
      ${xTicks.map((value) => {
        const x = padding.left + ((value - min) / range) * chartWidth;
        return `
          <line x1="${x}" y1="${height - padding.bottom}" x2="${x}" y2="${height - padding.bottom + 4}" stroke="#9eafbb" stroke-width="0.8"></line>
          <text x="${x}" y="${height - 8}" text-anchor="middle" fill="#667686" font-size="10.5">${formatter(value)}</text>
        `;
      }).join("")}
    </svg>
  `;
}

function createPertTrendSvg(rows) {
  const validRows = rows.filter((row) => row.expectedValue > 0);
  const source = validRows.length ? validRows : rows;
  const width = 520;
  const height = 230;
  const padding = 30;
  const maxValue = Math.max(...source.map((row) => row.pessimistic), 1);
  const gridLines = [0.25, 0.5, 0.75].map((ratio) => height - padding - ratio * (height - padding * 2));
  const expectedPath = source.map((row, index) => {
    const x = padding + (index / Math.max(source.length - 1, 1)) * (width - padding * 2);
    const y = height - padding - (row.expectedValue / maxValue) * (height - padding * 2);
    return `${index === 0 ? "M" : "L"}${x.toFixed(1)} ${y.toFixed(1)}`;
  }).join(" ");
  const optimisticPath = source.map((row, index) => {
    const x = padding + (index / Math.max(source.length - 1, 1)) * (width - padding * 2);
    const y = height - padding - (row.optimistic / maxValue) * (height - padding * 2);
    return `${index === 0 ? "M" : "L"}${x.toFixed(1)} ${y.toFixed(1)}`;
  }).join(" ");
  const pessimisticPath = source.map((row, index) => {
    const x = padding + (index / Math.max(source.length - 1, 1)) * (width - padding * 2);
    const y = height - padding - (row.pessimistic / maxValue) * (height - padding * 2);
    return `${index === 0 ? "M" : "L"}${x.toFixed(1)} ${y.toFixed(1)}`;
  }).join(" ");
  const bandPath = [
    source.map((row, index) => {
      const x = padding + (index / Math.max(source.length - 1, 1)) * (width - padding * 2);
      const y = height - padding - (row.pessimistic / maxValue) * (height - padding * 2);
      return `${index === 0 ? "M" : "L"}${x.toFixed(1)} ${y.toFixed(1)}`;
    }).join(" "),
    source.slice().reverse().map((row, reverseIndex) => {
      const index = source.length - 1 - reverseIndex;
      const x = padding + (index / Math.max(source.length - 1, 1)) * (width - padding * 2);
      const y = height - padding - (row.optimistic / maxValue) * (height - padding * 2);
      return `L${x.toFixed(1)} ${y.toFixed(1)}`;
    }).join(" "),
    "Z"
  ].join(" ");
  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Kostenverlauf">
      <defs>
        <linearGradient id="pertTrendBand" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="rgba(22,63,99,0.18)"></stop>
          <stop offset="100%" stop-color="rgba(22,63,99,0.04)"></stop>
        </linearGradient>
      </defs>
      ${gridLines.map((y) => `
        <line x1="${padding}" y1="${y}" x2="${width - padding}" y2="${y}" stroke="#dde7ef" stroke-width="1"></line>
      `).join("")}
      <line x1="${padding}" y1="${height - padding}" x2="${width - padding}" y2="${height - padding}" stroke="#c5d1db" stroke-width="1.4"></line>
      <line x1="${padding}" y1="${padding}" x2="${padding}" y2="${height - padding}" stroke="#c5d1db" stroke-width="1.4"></line>
      <path d="${bandPath}" fill="url(#pertTrendBand)" stroke="none"></path>
      <path d="${optimisticPath}" fill="none" stroke="#c06f34" stroke-width="2.2" stroke-dasharray="5 4" stroke-linecap="round"></path>
      <path d="${pessimisticPath}" fill="none" stroke="#1f7a53" stroke-width="2.2" stroke-dasharray="5 4" stroke-linecap="round"></path>
      <path d="${expectedPath}" fill="none" stroke="#163f63" stroke-width="3" stroke-linecap="round"></path>
      ${source.map((row, index) => {
        const x = padding + (index / Math.max(source.length - 1, 1)) * (width - padding * 2);
        const y = height - padding - (row.expectedValue / maxValue) * (height - padding * 2);
        return `
          <circle cx="${x}" cy="${y}" r="4" fill="#8a6831"></circle>
          <text x="${x}" y="${height - 10}" text-anchor="middle" fill="#536577" font-size="11">${row.label.replace("LPH ", "")}</text>
        `;
      }).join("")}
    </svg>
  `;
}

function createPertPhaseCorridorSvg(rows) {
  const source = rows.length ? rows : [];
  const width = 960;
  const height = 220;
  const padding = { top: 34, right: 28, bottom: 38, left: 58 };
  const values = source.flatMap((row) => [row.upper, row.expectedValue, row.lower]).filter((value) => value > 0);
  const minValue = values.length ? Math.min(...values) : 0;
  const maxValue = values.length ? Math.max(...values) : 1;
  const valueRange = Math.max(maxValue - minValue, 1);
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;

  const getX = (index) => padding.left + (index / Math.max(source.length - 1, 1)) * chartWidth;
  const getY = (value) => padding.top + (1 - ((value - minValue) / valueRange)) * chartHeight;

  const upperPath = source.map((row, index) => `${index === 0 ? "M" : "L"}${getX(index).toFixed(1)} ${getY(row.upper).toFixed(1)}`).join(" ");
  const lowerPath = source.map((row, index) => `${index === 0 ? "M" : "L"}${getX(index).toFixed(1)} ${getY(row.lower).toFixed(1)}`).join(" ");
  const expectedPath = source
    .filter((row) => row.isReached)
    .map((row, reachedIndex) => {
      const index = source.indexOf(row);
      return `${reachedIndex === 0 ? "M" : "L"}${getX(index).toFixed(1)} ${getY(row.expectedValue).toFixed(1)}`;
    })
    .join(" ");

  const ticks = Array.from({ length: 5 }, (_, index) => minValue + (valueRange * index) / 4);

  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Kostenverlauf nach Leistungsphasen">
      ${ticks.map((tick) => {
        const y = getY(tick);
        return `
          <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#e1e8ef" stroke-width="1"></line>
          <text x="${padding.left - 10}" y="${y + 4}" text-anchor="end" fill="#667789" font-size="12">${(tick / 1000000).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Mio. €</text>
        `;
      }).join("")}
      <line x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${height - padding.bottom}" stroke="#d6e0e8" stroke-width="1.5"></line>
      <line x1="${padding.left}" y1="${height - padding.bottom}" x2="${width - padding.right}" y2="${height - padding.bottom}" stroke="#d6e0e8" stroke-width="1.5"></line>
      <path d="${upperPath}" fill="none" stroke="#d78a71" stroke-width="2" stroke-dasharray="6 5" stroke-linecap="round"></path>
      <path d="${lowerPath}" fill="none" stroke="#7f99e8" stroke-width="2" stroke-dasharray="6 5" stroke-linecap="round"></path>
      ${expectedPath ? `<path d="${expectedPath}" fill="none" stroke="#1c2f4d" stroke-width="3" stroke-linecap="round"></path>` : ""}
      ${source.filter((row) => row.isReached).map((row) => {
        const index = source.indexOf(row);
        const x = getX(index);
        const y = getY(row.expectedValue);
        return `<circle cx="${x}" cy="${y}" r="5.5" fill="#1c2f4d" stroke="#fff" stroke-width="1.5"></circle>`;
      }).join("")}
      ${source.map((row, index) => `
        <text x="${getX(index)}" y="${height - 12}" text-anchor="middle" fill="#6a7b8c" font-size="12" font-weight="700">${row.label}</text>
      `).join("")}
    </svg>
  `;
}

function createPertSnapshotTrendSvg(rows) {
  const source = rows.length ? rows : [];
  const width = 960;
  const height = 240;
  const padding = { top: 28, right: 28, bottom: 40, left: 58 };
  const values = source.flatMap((row) => [row.expectedValue, row.budget84, row.budget975]).filter((value) => value > 0);
  const minValue = values.length ? Math.min(...values) : 0;
  const maxValue = values.length ? Math.max(...values) : 1;
  const paddedMin = Math.max(0, minValue - (maxValue - minValue) * 0.18);
  const paddedMax = maxValue + (maxValue - minValue) * 0.12 || 1;
  const valueRange = Math.max(paddedMax - paddedMin, 1);
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;

  const getX = (index) => padding.left + (index / Math.max(source.length - 1, 1)) * chartWidth;
  const getY = (value) => padding.top + (1 - ((value - paddedMin) / valueRange)) * chartHeight;
  const makePath = (key) => source.map((row, index) => `${index === 0 ? "M" : "L"}${getX(index).toFixed(1)} ${getY(row[key]).toFixed(1)}`).join(" ");
  const ticks = Array.from({ length: 5 }, (_, index) => paddedMin + (valueRange * index) / 4);

  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Kostenverlauf und Trendanalyse">
      ${ticks.map((tick) => {
        const y = getY(tick);
        return `
          <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#e1e8ef" stroke-width="1"></line>
          <text x="${padding.left - 10}" y="${y + 4}" text-anchor="end" fill="#667789" font-size="12">${(tick / 1000000).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Mio. €</text>
        `;
      }).join("")}
      <line x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${height - padding.bottom}" stroke="#d6e0e8" stroke-width="1.5"></line>
      <line x1="${padding.left}" y1="${height - padding.bottom}" x2="${width - padding.right}" y2="${height - padding.bottom}" stroke="#d6e0e8" stroke-width="1.5"></line>
      <path d="${makePath("expectedValue")}" fill="none" stroke="#1c2f4d" stroke-width="3" stroke-linecap="round"></path>
      <path d="${makePath("budget84")}" fill="none" stroke="#4f956c" stroke-width="2" stroke-dasharray="5 4" stroke-linecap="round"></path>
      <path d="${makePath("budget975")}" fill="none" stroke="#c6a03f" stroke-width="2" stroke-dasharray="5 4" stroke-linecap="round"></path>
      ${source.map((row, index) => {
        const x = getX(index);
        return `
          <circle cx="${x}" cy="${getY(row.expectedValue)}" r="5.5" fill="#1c2f4d"></circle>
          <circle cx="${x}" cy="${getY(row.budget84)}" r="4.5" fill="#4f956c"></circle>
          <circle cx="${x}" cy="${getY(row.budget975)}" r="4.5" fill="#c6a03f"></circle>
          <text x="${x}" y="${height - 12}" text-anchor="middle" fill="#6a7b8c" font-size="12" font-weight="700">${row.axisLabel}</text>
        `;
      }).join("")}
    </svg>
  `;
}

function calculatePertAnalytics(state, result) {
  const totalExpectedValue = result.expectedValue || 1;
  const lphRows = buildPertPhaseRows(state, result);
  const phaseCorridorRows = buildPertPhaseCorridorRows(state, result);
  const detailRows = result.items.map((item, originalIndex) => {
    const span = item.pessimistic - item.optimistic;
    const share = item.expectedValue / totalExpectedValue;
    const sensitivityScore = item.variance / Math.max(result.standardDeviation * result.standardDeviation, 1);
    return {
      ...item,
      originalIndex,
      span,
      share,
      sensitivityScore
    };
  }).sort((a, b) => b.expectedValue - a.expectedValue);

  const trendRows = lphRows.map((row) => ({
    ...row,
    value: row.expectedValue,
    delta: row.deltaExpected
  }));

  const snapshotRows = getLatestPertSnapshotsByPhase(state)
    .map((snapshot) => ({
      ...snapshot,
      budget84: Number(snapshot.expectedValue || 0) + Number(snapshot.standardDeviation || 0),
      budget975: Number(snapshot.expectedValue || 0) + 2 * Number(snapshot.standardDeviation || 0)
    }))
    .map((snapshot) => ({
      ...snapshot,
      axisLabel: `${snapshot.phase} · ${formatDate(snapshot.savedAt)}`
    }));

  const topSensitivity = [...detailRows]
    .sort((a, b) => b.sensitivityScore - a.sensitivityScore)
    .slice(0, 5);

  const latestSnapshot = state.pert.snapshots?.[0] || null;
  const scenarioCompare = latestSnapshot ? {
    label: latestSnapshot.label,
    savedAt: latestSnapshot.savedAt,
    expectedValueDelta: result.expectedValue - latestSnapshot.expectedValue,
    selectedBudgetDelta: result.selectedBudget - latestSnapshot.selectedBudget,
    standardDeviationDelta: result.standardDeviation - latestSnapshot.standardDeviation
  } : null;

  const normalDistribution = createNormalDistributionSvg(result.expectedValue, result.standardDeviation);
  return {
    detailRows,
    trendRows,
    phaseCorridorRows,
    snapshotRows,
    topSensitivity,
    scenarioCompare,
    normalDistributionSvg: normalDistribution.svg,
    normalDistributionMarkers: normalDistribution.markers,
    normalDistributionPositions: normalDistribution.positions,
    phaseCorridorSvg: createPertPhaseCorridorSvg(phaseCorridorRows),
    trendChartSvg: createPertSnapshotTrendSvg(snapshotRows),
    costComparisonSvg: createPertBarChartSvg(detailRows, "expectedValue", "#163f63", (value) => `${formatNumber(value)} €`),
    sensitivityChartSvg: createPertBarChartSvg(topSensitivity, "sensitivityScore", "#8a6831", (value) => `${(value * 100).toFixed(1)} %`),
    spanChartSvg: createPertBarChartSvg(detailRows, "span", "#1f5f53", (value) => `${formatNumber(value)} €`)
  };
}

function calculateMonteCarloAnalytics(result) {
  const workRows = result.workPackages
    .map((item, originalIndex) => ({
      ...item,
      originalIndex,
      spanCost: item.pessimisticCost - item.optimisticCost,
      spanDays: item.pessimisticDays - item.optimisticDays,
      costShare: item.meanCost / Math.max(result.meanCost, 1)
    }))
    .sort((a, b) => b.meanCost - a.meanCost);

  const riskRows = result.riskEvents
    .map((risk, originalIndex) => ({
      ...risk,
      originalIndex,
      exposureScore: risk.avgCostImpact + risk.avgDelayImpact * 1000
    }))
    .sort((a, b) => b.exposureScore - a.exposureScore);

  const workCostChartSvg = createPertBarChartSvg(
    workRows,
    "meanCost",
    "#1e4f7b",
    (value) => formatCurrency(value),
    { left: 164, right: 112, labelFontSize: 10.25, valueFontSize: 11 }
  );

  const workDayChartSvg = createPertBarChartSvg(
    workRows,
    "meanDays",
    "#8a6831",
    (value) => `${formatNumber(value)} T`,
    { left: 164, right: 112, labelFontSize: 10.25, valueFontSize: 11 }
  );

  const riskDriverChartSvg = createPertBarChartSvg(
    riskRows,
    "avgCostImpact",
    "#2f7f69",
    (value) => formatCurrency(value),
    { left: 148, right: 112, labelFontSize: 10.5, valueFontSize: 11 }
  );

  const costDistributionSvg = createMonteCarloDistributionSvg(
    result.costRuns,
    { p50: result.p50Cost, p80: result.p80Cost },
    {
      ariaLabel: "Kostenverteilung aus Monte-Carlo-Simulation",
      barColor: "#1e4f7b",
      formatter: (value) => `${formatNumber(value)} €`
    }
  );

  const dayDistributionSvg = createMonteCarloDistributionSvg(
    result.dayRuns,
    { p50: result.p50Days, p80: result.p80Days },
    {
      ariaLabel: "Dauerverteilung aus Monte-Carlo-Simulation",
      barColor: "#8a6831",
      formatter: (value) => `${formatNumber(value)} T`
    }
  );

  return {
    workRows,
    riskRows,
    workCostChartSvg,
    workDayChartSvg,
    riskDriverChartSvg,
    costDistributionSvg,
    dayDistributionSvg
  };
}

export function calculateEarnedValueResult(state) {
  const activities = (state.earnedValue.activities || []).map((activity) => {
    const plannedValue = Number(activity.plannedValue) || 0;
    const earnedValue = Number(activity.earnedValue) || 0;
    const actualCost = Number(activity.actualCost) || 0;
    const progressPercent = Number(activity.progressPercent) || 0;
    return {
      ...activity,
      plannedValue,
      earnedValue,
      actualCost,
      progressPercent
    };
  });

  const pv = activities.reduce((sum, item) => sum + item.plannedValue, 0);
  const ev = activities.reduce((sum, item) => sum + item.earnedValue, 0);
  const ac = activities.reduce((sum, item) => sum + item.actualCost, 0);
  const cv = ev - ac;
  const sv = ev - pv;
  const cpi = ac > 0 ? ev / ac : 0;
  const spi = pv > 0 ? ev / pv : 0;
  const budget = Number(state.project.budget) || 0;
  const eac = cpi > 0 ? budget / cpi : 0;
  const etc = Math.max(0, eac - ac);
  const vac = budget - eac;
  const tcpi = budget - ac > 0 ? (budget - ev) / (budget - ac) : 0;
  const progressPercent = budget > 0 ? (ev / budget) * 100 : 0;

  return {
    activities,
    pv,
    ev,
    ac,
    cv,
    sv,
    cpi,
    spi,
    eac,
    etc,
    vac,
    progressPercent,
    tcpi
  };
}

function getEarnedValueStatus(result, thresholds) {
  const cpiYellow = Number(thresholds?.cpiYellow) || 0.95;
  const spiYellow = Number(thresholds?.spiYellow) || 0.95;
  const cpiGreen = Number(thresholds?.cpiGreen) || 1;
  const spiGreen = Number(thresholds?.spiGreen) || 1;

  if (result.cpi >= cpiGreen && result.spi >= spiGreen) {
    return {
      level: "stabil",
      message: "Kosten und Termin liegen im grünen Bereich."
    };
  }

  if (result.cpi >= cpiYellow && result.spi >= spiYellow) {
    return {
      level: "angespannt",
      message: "Die Leistung ist steuerbar, zeigt aber erste Spannungen in Kosten oder Termin."
    };
  }

  return {
    level: "kritisch",
    message: "Die laufende Projektperformance ist kritisch und sollte kurzfristig nachgesteuert werden."
  };
}

function getEvaStatusLevel(cpi, spi, thresholds) {
  const cpiYellow = Number(thresholds?.cpiYellow) || 0.95;
  const spiYellow = Number(thresholds?.spiYellow) || 0.95;
  const cpiGreen = Number(thresholds?.cpiGreen) || 1;
  const spiGreen = Number(thresholds?.spiGreen) || 1;

  if (cpi >= cpiGreen && spi >= spiGreen) return "stabil";
  if (cpi >= cpiYellow && spi >= spiYellow) return "angespannt";
  return "kritisch";
}

function getEvaCostTolerance(state, budget) {
  const thresholds = state.earnedValue?.thresholds || {};
  const storedAmount = Number(thresholds.cvTolerance) || 180000;
  const percent = Number.isFinite(Number(thresholds.cvTolerancePercent))
    ? clampEvaTolerancePercent(thresholds.cvTolerancePercent)
    : budget > 0
      ? (storedAmount / budget) * 100
      : 3;
  const amount = budget > 0 ? (budget * percent) / 100 : storedAmount;
  return {
    percent,
    amount
  };
}

function getEvaActivityStatus(activity, thresholds, options = {}) {
  const budget = Number(options.budget) || 0;
  const bac = Number(activity?.bac) || 0;
  const cvTolerancePercent = clampEvaTolerancePercent(thresholds?.cvTolerancePercent);
  const globalEacTolerance = Number(thresholds?.eacTolerance) || 180000;
  const localCvTolerance = bac > 0 ? (bac * cvTolerancePercent) / 100 : 0;
  const localEacTolerance = budget > 0 && bac > 0 ? (globalEacTolerance / budget) * bac : globalEacTolerance;
  const statusLevel = getEvaStatusLevel(activity.cpi, activity.spi, thresholds);
  const eacOverrun = Number(activity.eac || 0) - bac;
  const cvCritical = localCvTolerance > 0 && Math.abs(Number(activity.cv) || 0) > localCvTolerance;
  const eacCritical = localEacTolerance > 0 && eacOverrun > localEacTolerance;

  if (statusLevel === "kritisch" || cvCritical || eacCritical) return "Kritisch";
  if (statusLevel === "angespannt") return "Achtung";
  return "OK";
}

function getEvaManagementCards(state, result) {
  const budget = Number(state.project?.budget) || 0;
  const thresholds = state.earnedValue?.thresholds || {};
  const cvTolerance = getEvaCostTolerance(state, budget).amount;
  const eacTolerance = Number(thresholds.eacTolerance) || 180000;
  const tcpiActionLevel = Number(thresholds.tcpiActionLevel) || 1.02;
  const progressPercent = budget > 0 ? (result.ev / budget) * 100 : 0;
  const absCv = Math.abs(result.cv);
  const eacOverrun = result.eac - budget;

  const cpiStatus = result.cpi >= 0.97 ? "STABIL" : result.cpi >= 0.93 ? "ANGESPANNT" : "KRITISCH";
  const cpiAction = result.cpi >= 0.97 ? "EMPFEHLUNG" : "MASSNAHMEN ERFORDERLICH";
  const cpiMessage = result.cpi >= 0.97
    ? `Kosten im Rahmen – CPI ${result.cpi.toFixed(2)}, Abweichung ${formatCurrency(result.cv)} liegt innerhalb der Toleranz (${formatCurrency(cvTolerance)}). Monitoring beibehalten.`
    : `Kosten stehen unter Druck – CPI ${result.cpi.toFixed(2)}, Abweichung ${formatCurrency(result.cv)} überschreitet die Toleranz (${formatCurrency(cvTolerance)}). Kostenursachen aktiv nachsteuern.`;

  const spiStatus = result.spi >= 1 ? "STABIL" : result.spi >= 0.95 ? "ANGESPANNT" : "KRITISCH";
  const spiAction = result.spi >= 1 ? "EMPFEHLUNG" : "MASSNAHMEN ERFORDERLICH";
  const spiMessage = result.spi >= 1
    ? `Terminverlauf entspricht dem Plan (SPI ${result.spi.toFixed(2)}). Kritischen Pfad weiter beobachten.`
    : `Terminverlauf ist angespannt – SPI ${result.spi.toFixed(2)}. Kritische Vorgänge eng führen und Puffer aktivieren.`;

  const eacStatus = budget <= 0
    ? "NEUTRAL"
    : Math.abs(eacOverrun) <= eacTolerance ? "STABIL" : eacOverrun > 0 ? "ANGESPANNT" : "STABIL";
  const eacAction = budget <= 0 ? "HINWEIS" : Math.abs(eacOverrun) <= eacTolerance ? "EMPFEHLUNG" : "MASSNAHMEN ERFORDERLICH";
  const eacMessage = budget <= 0
    ? "Für die EAC-Bewertung ist ein globales Projektbudget erforderlich."
    : Math.abs(eacOverrun) <= eacTolerance
      ? `EAC-Abweichung von ${formatCurrency(eacOverrun)} liegt innerhalb der Toleranz (${formatCurrency(eacTolerance)}). Monitoring beibehalten.`
      : eacOverrun > 0
        ? `EAC-Überschreitung von ${formatCurrency(eacOverrun)} liegt außerhalb der Toleranz (${formatCurrency(eacTolerance)}). Gegensteuerung und Budgetklärung notwendig.`
        : `EAC liegt ${formatCurrency(Math.abs(eacOverrun))} unter Budget und damit innerhalb der steuerbaren Bandbreite.`;

  const tcpiStatus = result.tcpi < 1 ? "STABIL" : result.tcpi < tcpiActionLevel ? "ANGESPANNT" : "KRITISCH";
  const tcpiAction = result.tcpi < tcpiActionLevel ? "EMPFEHLUNG" : "MASSNAHMEN ERFORDERLICH";
  const tcpiMessage = result.tcpi < 1
    ? `TCPI ${result.tcpi.toFixed(2)}: Die Restleistung kann mit dem bisherigen Leistungsniveau erreicht werden.`
    : result.tcpi < tcpiActionLevel
      ? `TCPI ${result.tcpi.toFixed(2)}: Das Restbudget muss leicht effizienter eingesetzt werden als bisher. Enges Monitoring ist ausreichend.`
      : `TCPI ${result.tcpi.toFixed(2)}: Das Restbudget muss um ${formatNumber((result.tcpi - 1) * 100)} % effizienter eingesetzt werden als bisher. Gegensteuerungsmaßnahmen und Kostendisziplin sind notwendig.`;

  const progressStatus = "NEUTRAL";
  const progressAction = "EINORDNUNG";
  const progressMessage = "FGR ist leistungswertgewichtet – budgetintensive Gewerke gewichten stärker. Eine Abweichung zum visuellen Baufortschritt ist methodisch normal.";

  return [
    {
      title: "Kostenstatus (CPI)",
      value: result.cpi.toFixed(2),
      status: cpiStatus,
      action: cpiAction,
      actionShort: cpiAction.includes("ERFORDERLICH") ? "Nachsteuern" : "Monitoring",
      message: cpiMessage,
      tone: cpiStatus === "KRITISCH" ? "kritisch" : cpiStatus === "STABIL" ? "stabil" : "angespannt"
    },
    {
      title: "Terminstatus (SPI)",
      value: result.spi.toFixed(2),
      status: spiStatus,
      action: spiAction,
      actionShort: spiAction.includes("ERFORDERLICH") ? "Termin steuern" : "Monitoring",
      message: spiMessage,
      tone: spiStatus === "KRITISCH" ? "kritisch" : spiStatus === "STABIL" ? "stabil" : "angespannt"
    },
    {
      title: "Kostenprognose (EAC)",
      value: formatCurrency(result.eac),
      status: eacStatus,
      action: eacAction,
      actionShort: eacAction.includes("ERFORDERLICH") ? "Budget klären" : eacAction === "HINWEIS" ? "Prüfen" : "Monitoring",
      message: eacMessage,
      tone: eacStatus === "NEUTRAL" ? "neutral" : eacStatus === "STABIL" ? "stabil" : "angespannt"
    },
    {
      title: "Restleistungsindex (TCPI)",
      value: result.tcpi.toFixed(2),
      status: tcpiStatus,
      action: tcpiAction,
      actionShort: tcpiAction.includes("ERFORDERLICH") ? "Maßnahmen" : "Monitoring",
      message: tcpiMessage,
      tone: tcpiStatus === "KRITISCH" ? "kritisch" : tcpiStatus === "STABIL" ? "stabil" : "angespannt"
    },
    {
      title: "Fertigstellungsgrad (FGR)",
      value: `${formatNumber(progressPercent)}%`,
      status: progressStatus,
      action: progressAction,
      actionShort: "Einordnen",
      message: progressMessage,
      tone: "neutral"
    }
  ];
}

function classifyEvaCostGroup(name) {
  const value = String(name || "").toLowerCase();
  if (value.includes("tga") || value.includes("technik") || value.includes("gebäude")) return "KG 400";
  if (value.includes("außen") || value.includes("landschaft")) return "KG 500";
  return "KG 300";
}

function createEvaToleranceSvg(toleranceRatio, percentLabel) {
  return "";
}

function createEvaTrendSvg(rows, budget) {
  const width = 980;
  const height = 272;
  const left = 52;
  const right = width - 56;
  const top = 36;
  const bottom = height - 48;
  const costValues = rows.flatMap((row) => [row.eac, row.bac || budget]);
  const indexValues = rows.flatMap((row) => [row.cpi, row.spi]);
  const minIndex = Math.min(0.6, ...indexValues, 0.95);
  const maxIndex = Math.max(1.4, ...indexValues, 1.05);
  const maxCost = Math.max(...costValues, budget || 0, 1);
  const minCost = Math.min(...costValues, budget || maxCost, maxCost);
  const costRange = Math.max(maxCost - minCost, 1);

  const xAt = (index) => left + ((right - left) * index) / Math.max(rows.length - 1, 1);
  const yIndexAt = (value) => bottom - ((value - minIndex) / Math.max(maxIndex - minIndex, 0.01)) * (bottom - top);
  const yCostAt = (value) => bottom - ((value - minCost) / costRange) * (bottom - top);

  const buildPoints = (key, mapper) => rows.map((row, index) => `${xAt(index)},${mapper(row[key])}`).join(" ");

  const gridLines = [0, 0.25, 0.5, 0.75, 1].map((ratio) => {
    const y = top + (bottom - top) * ratio;
    const value = (maxIndex - (maxIndex - minIndex) * ratio).toFixed(2).replace(".", ",");
    return `
      <line x1="${left}" y1="${y}" x2="${right}" y2="${y}" stroke="#dbe4ea" stroke-width="1"/>
      <text x="${left - 8}" y="${y + 4}" text-anchor="end" fill="#667789" font-size="10">${value}</text>
    `;
  }).join("");

  const xLabels = rows.map((row, index) => `
    <text x="${xAt(index)}" y="${height - 14}" text-anchor="middle" fill="#4b5d6e" font-size="11">${formatDate(row.statusDate)}</text>
  `).join("");

  const costAxisLabels = [0, 0.25, 0.5, 0.75, 1].map((ratio) => {
    const y = top + (bottom - top) * ratio;
    const value = formatCompactCurrency(maxCost - costRange * ratio);
    return `<text x="${right + 10}" y="${y + 4}" fill="#667789" font-size="10">${value}</text>`;
  }).join("");

  const lastRow = rows[rows.length - 1];
  const lastX = xAt(rows.length - 1);
  const cpiY = yIndexAt(lastRow.cpi);
  const spiY = yIndexAt(lastRow.spi);
  const eacY = yCostAt(lastRow.eac);

  return `
    <svg viewBox="0 0 ${width} ${height}" class="eva-trend-svg" role="img" aria-label="Stichtag-Verlauf und Trendanalyse">
      ${gridLines}
      ${costAxisLabels}
      <line x1="${left}" y1="${bottom}" x2="${right}" y2="${bottom}" stroke="#b9c7d4" stroke-width="1.2"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${bottom}" stroke="#b9c7d4" stroke-width="1.2"/>
      <polyline fill="none" stroke="#b07d22" stroke-width="2.2" points="${buildPoints("cpi", yIndexAt)}"/>
      <polyline fill="none" stroke="#356daa" stroke-width="2.2" points="${buildPoints("spi", yIndexAt)}"/>
      <polyline fill="none" stroke="#c45d4b" stroke-width="2.2" points="${buildPoints("eac", yCostAt)}"/>
      <polyline fill="none" stroke="#c0c4c8" stroke-width="2.2" stroke-dasharray="5 4" points="${rows.map((row, index) => `${xAt(index)},${yCostAt(row.bac || budget)}`).join(" ")}"/>
      ${rows.map((row, index) => `
        <circle cx="${xAt(index)}" cy="${yIndexAt(row.cpi)}" r="4" fill="#b07d22"/>
        <circle cx="${xAt(index)}" cy="${yIndexAt(row.spi)}" r="4" fill="#356daa"/>
        <circle cx="${xAt(index)}" cy="${yCostAt(row.eac)}" r="4.5" fill="#c45d4b"/>
      `).join("")}
      <text x="${left}" y="${top - 10}" fill="#b07d22" font-size="11" font-weight="700">CPI</text>
      <text x="${left + 34}" y="${top - 10}" fill="#356daa" font-size="11" font-weight="700">SPI</text>
      <text x="${right - 70}" y="${top - 10}" fill="#c45d4b" font-size="11" font-weight="700">EAC</text>
      <text x="${right - 48}" y="${top - 10}" fill="#9da8b4" font-size="11" font-weight="700">Budget (BAC)</text>
      <text x="${left - 30}" y="${top - 2}" fill="#667789" font-size="10" transform="rotate(-90 ${left - 30} ${top - 2})">Index (CPI / SPI)</text>
      <text x="${right + 24}" y="${top + 6}" fill="#667789" font-size="10" transform="rotate(90 ${right + 24} ${top + 6})">Kosten (EAC / Budget)</text>
      <text x="${lastX + 10}" y="${cpiY - 8}" fill="#b07d22" font-size="10" font-weight="700">${lastRow.cpi.toFixed(2).replace(".", ",")}</text>
      <text x="${lastX + 10}" y="${spiY + 4}" fill="#356daa" font-size="10" font-weight="700">${lastRow.spi.toFixed(2).replace(".", ",")}</text>
      <text x="${lastX + 10}" y="${eacY - 6}" fill="#c45d4b" font-size="10" font-weight="700">${formatCompactCurrency(lastRow.eac)}</text>
      ${xLabels}
    </svg>
  `;
}

function createEvaSCurveSvg(project, result) {
  const width = 470;
  const height = 236;
  const left = 48;
  const right = width - 22;
  const top = 18;
  const bottom = height - 30;
  const budget = Number(project?.budget) || Math.max(result.ac, result.ev, result.pv, 1);
  const startDate = new Date(project?.startDate || project?.analysisDate || Date.now());
  const endDate = new Date(project?.endDate || project?.analysisDate || Date.now());
  const statusDate = new Date(project?.analysisDate || Date.now());
  const totalDuration = Math.max(1, endDate.getTime() - startDate.getTime());
  const elapsedDuration = Math.max(0, Math.min(totalDuration, statusDate.getTime() - startDate.getTime()));
  const statusRatio = totalDuration > 0 ? elapsedDuration / totalDuration : Math.max(0.08, Math.min(0.92, (result.pv || result.ev || result.ac) / Math.max(budget, 1)));
  const safeStatusRatio = Math.max(0.08, Math.min(0.92, statusRatio));
  const pvRatio = Math.max(0, Math.min(1, result.pv / Math.max(budget, 1)));
  const evRatio = Math.max(0, Math.min(1, result.ev / Math.max(budget, 1)));
  const acRatio = Math.max(0, Math.min(1, result.ac / Math.max(budget, 1)));
  const xAt = (ratio) => left + (right - left) * ratio;
  const yAt = (value) => bottom - (value / Math.max(budget, 1)) * (bottom - top);
  const buildPolyline = (pointBuilder, steps = 16) => Array.from({ length: steps + 1 }, (_, index) => {
    const ratio = index / steps;
    const valueRatio = pointBuilder(ratio);
    return `${xAt(ratio)},${yAt(budget * valueRatio)}`;
  }).join(" ");
  const plannedRatioAt = (ratio) => {
    if (ratio <= safeStatusRatio) {
      const localRatio = safeStatusRatio > 0 ? ratio / safeStatusRatio : 0;
      return pvRatio * Math.pow(localRatio, 1.18);
    }
    const remainingRatio = (ratio - safeStatusRatio) / Math.max(1 - safeStatusRatio, 0.0001);
    return pvRatio + (1 - pvRatio) * Math.pow(remainingRatio, 0.9);
  };
  const buildActualPolyline = (targetRatio, curveStrength) => {
    const points = [`${xAt(0)},${yAt(0)}`];
    [0.24, 0.5, 0.76, 1].forEach((step) => {
      const ratio = safeStatusRatio * step;
      const valueRatio = targetRatio * Math.pow(step, curveStrength);
      points.push(`${xAt(ratio)},${yAt(budget * valueRatio)}`);
    });
    return points.join(" ");
  };

  const pvPoints = buildPolyline(plannedRatioAt);
  const evPoints = buildActualPolyline(evRatio, 1.08);
  const acPoints = buildActualPolyline(acRatio, 1.02);
  const statusX = xAt(safeStatusRatio);
  const formatAxisDate = (dateLike) => {
    if (!dateLike) return "";
    const date = new Date(dateLike);
    return new Intl.DateTimeFormat("de-DE", { month: "short", year: "2-digit" }).format(date);
  };
  const addDays = (date, ratio) => new Date(startDate.getTime() + totalDuration * ratio);
  const timeTicks = [0, 0.25, 0.5, 0.75, 1].map((ratio) => ({
    ratio,
    label: formatAxisDate(addDays(startDate, ratio))
  }));

  return `
    <svg viewBox="0 0 ${width} ${height}" class="eva-scurve-svg" role="img" aria-label="S-Kurve">
      ${[0, 0.25, 0.5, 0.75, 1].map((ratio) => {
        const y = top + (bottom - top) * ratio;
        return `<line x1="${left}" y1="${y}" x2="${right}" y2="${y}" stroke="#e1e8ee" stroke-width="1"/><text x="${left - 8}" y="${y + 4}" text-anchor="end" fill="#667789" font-size="10">${formatCompactCurrency(budget * (1 - ratio))}</text>`;
      }).join("")}
      ${timeTicks.map((tick) => `<line x1="${xAt(tick.ratio)}" y1="${bottom}" x2="${xAt(tick.ratio)}" y2="${bottom + 5}" stroke="#b9c7d4" stroke-width="1"/>`).join("")}
      <line x1="${left}" y1="${bottom}" x2="${right}" y2="${bottom}" stroke="#b9c7d4" stroke-width="1.2"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${bottom}" stroke="#b9c7d4" stroke-width="1.2"/>
      <line x1="${statusX}" y1="${top}" x2="${statusX}" y2="${bottom}" stroke="#d2dce5" stroke-width="1.1" stroke-dasharray="4 4"/>
      <polyline fill="none" stroke="#9ea7b3" stroke-width="2.2" stroke-dasharray="4 3" points="${pvPoints}"/>
      <polyline fill="none" stroke="#1f3c63" stroke-width="2.4" points="${evPoints}"/>
      <polyline fill="none" stroke="#c45d4b" stroke-width="2.2" points="${acPoints}"/>
      <circle cx="${xAt(1)}" cy="${yAt(budget)}" r="3.6" fill="#9ea7b3"/>
      <circle cx="${statusX}" cy="${yAt(result.ev)}" r="4" fill="#1f3c63"/>
      <circle cx="${statusX}" cy="${yAt(result.ac)}" r="4" fill="#c45d4b"/>
      <text x="${xAt(1) - 6}" y="${yAt(budget) - 8}" text-anchor="end" fill="#7f8a97" font-size="10" font-weight="700">Planwert gesamt (PV) ${formatCompactCurrency(budget)}</text>
      <text x="${statusX + 10}" y="${yAt(result.ev) - 8}" fill="#1f3c63" font-size="10" font-weight="700">${formatCompactCurrency(result.ev)}</text>
      <text x="${statusX + 10}" y="${yAt(result.ac) + 14}" fill="#c45d4b" font-size="10" font-weight="700">${formatCompactCurrency(result.ac)}</text>
      ${timeTicks.map((tick) => `<text x="${xAt(tick.ratio)}" y="${height - 10}" text-anchor="middle" fill="#667789" font-size="10">${tick.label}</text>`).join("")}
      <text x="${left}" y="${height - 24}" fill="#667789" font-size="10">Projektzeitachse</text>
      <text x="${statusX + 6}" y="${top + 12}" fill="#667789" font-size="10">Stichtag</text>
    </svg>
  `;
}

function createEvaProgressComparisonSvg(activities) {
  const width = 470;
  const height = Math.max(206, activities.length * 48 + 42);
  const left = 92;
  const right = width - 40;
  const top = 20;
  const barArea = right - left;
  const rowGap = 44;
  return `
    <svg viewBox="0 0 ${width} ${height}" class="eva-progress-svg" role="img" aria-label="Fertigstellungsgrad Plan vs. Ist">
      ${activities.map((activity, index) => {
        const y = top + index * rowGap;
        const planProgress = Math.max(0, Math.min(100, activity.plannedValue > 0 ? (activity.earnedValue / Math.max(activity.plannedValue, activity.earnedValue)) * 100 : 0));
        const actualProgress = Math.max(0, Math.min(100, activity.fgrPercent || 0));
        return `
          <text x="${left - 12}" y="${y + 18}" text-anchor="end" fill="#4b5d6e" font-size="11">${activity.name}</text>
          <rect x="${left}" y="${y}" width="${barArea * (planProgress / 100)}" height="16" rx="7" fill="#c9ced6"/>
          <rect x="${left}" y="${y + 20}" width="${barArea * (actualProgress / 100)}" height="16" rx="7" fill="#5e9b74"/>
          <text x="${Math.min(right + 8, left + barArea * (planProgress / 100) + 8)}" y="${y + 12}" fill="#667789" font-size="10">${formatNumber(planProgress)}%</text>
          <text x="${Math.min(right + 8, left + barArea * (actualProgress / 100) + 8)}" y="${y + 32}" fill="#3e7a55" font-size="10" font-weight="700">${formatNumber(actualProgress)}%</text>
        `;
      }).join("")}
    </svg>
  `;
}

function createEvaDonutSvg(groups) {
  const total = groups.reduce((sum, group) => sum + group.pv, 0) || 1;
  const radius = 78;
  const centerX = 110;
  const centerY = 110;
  let current = -Math.PI / 2;
  const colors = ["#1f3c63", "#9aa6b8", "#cf8d7e", "#5e9b74"];
  const arcs = groups.map((group, index) => {
    const angle = (group.pv / total) * Math.PI * 2;
    const next = current + angle;
    const x1 = centerX + Math.cos(current) * radius;
    const y1 = centerY + Math.sin(current) * radius;
    const x2 = centerX + Math.cos(next) * radius;
    const y2 = centerY + Math.sin(next) * radius;
    const largeArc = angle > Math.PI ? 1 : 0;
    const path = `M ${centerX} ${centerY} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
    current = next;
    return `<path d="${path}" fill="${colors[index % colors.length]}"/>`;
  }).join("");
  return `
    <svg viewBox="0 0 320 220" class="eva-donut-svg" role="img" aria-label="Budget nach Kostengruppe">
      ${arcs}
      <circle cx="${centerX}" cy="${centerY}" r="44" fill="#ffffff"/>
      <text x="${centerX}" y="${centerY - 6}" text-anchor="middle" fill="#667789" font-size="10">Planwert gesamt (PV)</text>
      <text x="${centerX}" y="${centerY + 14}" text-anchor="middle" fill="#1f3c63" font-size="14" font-weight="700">${formatCompactCurrency(total)}</text>
      ${groups.map((group, index) => `<text x="238" y="${52 + index * 24}" fill="#4b5d6e" font-size="11">${group.code}</text><circle cx="224" cy="${48 + index * 24}" r="5" fill="${["#1f3c63", "#9aa6b8", "#cf8d7e", "#5e9b74"][index % 4]}"/>`).join("")}
    </svg>
  `;
}

function createEvaCostGroupBarsSvg(groups) {
  const width = 430;
  const height = 220;
  const left = 42;
  const right = width - 20;
  const bottom = height - 30;
  const top = 22;
  const maxValue = Math.max(...groups.flatMap((group) => [group.pv, group.ev, group.ac]), 1);
  const groupWidth = (right - left) / Math.max(groups.length, 1);
  const barWidth = Math.min(30, (groupWidth - 22) / 3);
  const colors = { pv: "#9aa6b8", ev: "#1f3c63", ac: "#cf8d7e" };
  return `
    <svg viewBox="0 0 ${width} ${height}" class="eva-cost-group-svg" role="img" aria-label="PV EV AC je Kostengruppe">
      ${[0, 0.25, 0.5, 0.75, 1].map((ratio) => {
        const y = top + (bottom - top) * ratio;
        return `<line x1="${left}" y1="${y}" x2="${right}" y2="${y}" stroke="#e1e8ee" stroke-width="1"/><text x="${left - 8}" y="${y + 4}" text-anchor="end" fill="#667789" font-size="10">${formatCompactCurrency(maxValue * (1 - ratio))}</text>`;
      }).join("")}
      ${groups.map((group, index) => {
        const baseX = left + index * groupWidth + 20;
        const items = [["pv", group.pv], ["ev", group.ev], ["ac", group.ac]];
        return items.map(([key, value], itemIndex) => {
          const barHeight = (value / maxValue) * (bottom - top);
          const x = baseX + itemIndex * (barWidth + 8);
          return `<rect x="${x}" y="${bottom - barHeight}" width="${barWidth}" height="${barHeight}" rx="6" fill="${colors[key]}"/>`;
        }).join("") + `<text x="${left + index * groupWidth + groupWidth / 2}" y="${height - 10}" text-anchor="middle" fill="#4b5d6e" font-size="11">${group.code}</text>`;
      }).join("")}
    </svg>
  `;
}

function buildEvaAnalytics(state, result) {
  const budget = Number(state.project?.budget) || 0;
  const thresholds = state.earnedValue?.thresholds || {};
  const snapshotRowsRaw = [...(state.earnedValue?.snapshots || [])]
    .sort((a, b) => String(a.statusDate || a.savedAt).localeCompare(String(b.statusDate || b.savedAt)));
  const snapshotRows = snapshotRowsRaw.map((row, index) => {
    const previous = index > 0 ? snapshotRowsRaw[index - 1] : null;
    return {
      ...row,
      deltaCpi: previous ? row.cpi - previous.cpi : null,
      deltaSpi: previous ? row.spi - previous.spi : null,
      deltaEac: previous ? row.eac - previous.eac : null,
      deltaFgr: previous ? row.fgr - previous.fgr : null
    };
  });

  const activities = result.activities.map((activity) => {
    const cv = activity.earnedValue - activity.actualCost;
    const sv = activity.earnedValue - activity.plannedValue;
    const cpi = activity.actualCost > 0 ? activity.earnedValue / activity.actualCost : 0;
    const spi = activity.plannedValue > 0 ? activity.earnedValue / activity.plannedValue : 0;
    const bac = Math.max(activity.plannedValue, activity.earnedValue, activity.actualCost);
    const fgrPercent = bac > 0 ? Math.max(0, Math.min(100, (activity.earnedValue / bac) * 100)) : 0;
    const eac = cpi > 0 ? bac / cpi : 0;
    const status = getEvaActivityStatus({ cv, cpi, spi, bac, eac }, thresholds, { budget });
    return {
      ...activity,
      cv,
      sv,
      cpi,
      spi,
      bac,
      fgrPercent,
      eac,
      status,
      costGroup: classifyEvaCostGroup(activity.name)
    };
  });

  const costGroupsMap = new Map();
  activities.forEach((activity) => {
    const code = activity.costGroup;
    const current = costGroupsMap.get(code) || { code, pv: 0, ev: 0, ac: 0 };
    current.pv += activity.plannedValue;
    current.ev += activity.earnedValue;
    current.ac += activity.actualCost;
    costGroupsMap.set(code, current);
  });
  const costGroups = [...costGroupsMap.values()];

  const cvTolerance = getEvaCostTolerance(state, budget).amount;
  const latestSnapshot = snapshotRows.length ? snapshotRows[snapshotRows.length - 1] : null;
  const currentDeltas = latestSnapshot ? {
    cpi: result.cpi - latestSnapshot.cpi,
    spi: result.spi - latestSnapshot.spi,
    eac: result.eac - latestSnapshot.eac,
    fgr: result.progressPercent - latestSnapshot.fgr
  } : null;

  return {
    snapshotRows,
    currentDeltas,
    activities,
    costGroups,
    trendSvg: snapshotRows.length ? createEvaTrendSvg(snapshotRows, budget) : "",
    sCurveSvg: createEvaSCurveSvg(state.project, result),
    progressSvg: createEvaProgressComparisonSvg(activities),
    donutSvg: createEvaDonutSvg(costGroups),
    costGroupSvg: createEvaCostGroupBarsSvg(costGroups)
  };
}

function buildMonteCarloConfig(state, options = {}) {
  const monteCarloSource = options.monteCarloOverride || state.monteCarlo || {};
  const measureSettings = monteCarloSource.measureSettings || {};
  const applyMeasures = Boolean(options.applyMeasures);
  const budgetFactor = applyMeasures ? 1 + (Number(measureSettings.budgetPercent) || 0) / 100 : 1;
  const dayFactor = applyMeasures ? 1 + (Number(measureSettings.daysPercent) || 0) / 100 : 1;
  const riskFactor = applyMeasures ? Math.max(0, 1 - (Number(measureSettings.riskReductionPercent) || 0) / 100) : 1;
  const impactFactor = applyMeasures ? Math.max(0, 1 - (Number(measureSettings.impactReductionPercent) || 0) / 100) : 1;

  const workPackages = (monteCarloSource.workPackages || []).map((item) => {
    const optimisticCost = Number(item.optimisticCost) || 0;
    const mostLikelyCost = Number(item.mostLikelyCost) || 0;
    const pessimisticCost = Number(item.pessimisticCost) || 0;
    const optimisticDays = Number(item.optimisticDays) || 0;
    const mostLikelyDays = Number(item.mostLikelyDays) || 0;
    const pessimisticDays = Number(item.pessimisticDays) || 0;
    const meanCost = (optimisticCost + 4 * mostLikelyCost + pessimisticCost) / 6;
    const meanDays = (optimisticDays + 4 * mostLikelyDays + pessimisticDays) / 6;
    return {
      ...item,
      optimisticCost,
      mostLikelyCost,
      pessimisticCost,
      optimisticDays,
      mostLikelyDays,
      pessimisticDays,
      meanCost,
      meanDays
    };
  });

  const riskEvents = (monteCarloSource.riskEvents || []).map((risk) => {
    const probability = (Number(risk.probability) || 0) * riskFactor;
    const minCostImpact = (Number(risk.minCostImpact) || 0) * impactFactor;
    const maxCostImpact = (Number(risk.maxCostImpact) || 0) * impactFactor;
    const minDelayDays = (Number(risk.minDelayDays) || 0) * impactFactor;
    const maxDelayDays = (Number(risk.maxDelayDays) || 0) * impactFactor;
    const avgCostImpact = ((minCostImpact + maxCostImpact) / 2) * probability;
    const avgDelayImpact = ((minDelayDays + maxDelayDays) / 2) * probability;
    return {
      ...risk,
      probability,
      minCostImpact,
      maxCostImpact,
      minDelayDays,
      maxDelayDays,
      avgCostImpact,
      avgDelayImpact
    };
  });
  return {
    iterations: Math.max(1000, Number(monteCarloSource.iterations) || 10000),
    seed: monteCarloSource.seed || 42,
    targetBudget: (Number(monteCarloSource.targetBudget) || 0) * budgetFactor,
    targetDays: (Number(monteCarloSource.targetDays) || 0) * dayFactor,
    workPackages,
    riskEvents,
    measureSettings
  };
}

export function calculateMonteCarloResult(state, options = {}) {
  const config = buildMonteCarloConfig(state, options);
  const workPackages = config.workPackages;
  const riskEvents = config.riskEvents;
  const iterations = config.iterations;
  const rng = createRng(config.seed);
  const costRuns = [];
  const dayRuns = [];
  let budgetHits = 0;
  let scheduleHits = 0;
  let combinedHits = 0;
  let totalCostOverrun = 0;
  let totalDayOverrun = 0;
  const driverStats = riskEvents.map((risk) => ({
    name: risk.name,
    totalCostImpact: 0,
    totalDelayImpact: 0,
    hits: 0
  }));

  const targetBudget = config.targetBudget;
  const targetDays = config.targetDays;

  for (let index = 0; index < iterations; index += 1) {
    let totalCost = 0;
    let totalDays = 0;

    workPackages.forEach((item) => {
      totalCost += samplePert(item.optimisticCost, item.mostLikelyCost, item.pessimisticCost, rng);
      totalDays += samplePert(item.optimisticDays, item.mostLikelyDays, item.pessimisticDays, rng);
    });

    riskEvents.forEach((risk, riskIndex) => {
      if (rng() <= risk.probability) {
        const costImpact = risk.minCostImpact + rng() * (risk.maxCostImpact - risk.minCostImpact);
        const dayImpact = risk.minDelayDays + rng() * (risk.maxDelayDays - risk.minDelayDays);
        totalCost += costImpact;
        totalDays += dayImpact;
        driverStats[riskIndex].totalCostImpact += costImpact;
        driverStats[riskIndex].totalDelayImpact += dayImpact;
        driverStats[riskIndex].hits += 1;
      }
    });

    costRuns.push(totalCost);
    dayRuns.push(totalDays);

    const withinBudget = targetBudget > 0 ? totalCost <= targetBudget : false;
    const withinSchedule = targetDays > 0 ? totalDays <= targetDays : false;

    if (withinBudget) budgetHits += 1;
    if (withinSchedule) scheduleHits += 1;
    if (withinBudget && withinSchedule) combinedHits += 1;

    totalCostOverrun += Math.max(0, totalCost - targetBudget);
    totalDayOverrun += Math.max(0, totalDays - targetDays);
  }

  const meanCost = mean(costRuns);
  const meanDays = mean(dayRuns);
  const medianCost = percentile(costRuns, 0.5);
  const medianDays = percentile(dayRuns, 0.5);
  const p10Cost = percentile(costRuns, 0.1);
  const p50Cost = percentile(costRuns, 0.5);
  const p80Cost = percentile(costRuns, 0.8);
  const p90Cost = percentile(costRuns, 0.9);
  const p10Days = percentile(dayRuns, 0.1);
  const p50Days = percentile(dayRuns, 0.5);
  const p80Days = percentile(dayRuns, 0.8);
  const p90Days = percentile(dayRuns, 0.9);
  const budgetSuccessRate = targetBudget > 0 ? budgetHits / iterations : 0;
  const scheduleSuccessRate = targetDays > 0 ? scheduleHits / iterations : 0;
  const combinedSuccessRate = targetBudget > 0 && targetDays > 0 ? combinedHits / iterations : 0;
  const avgCostOverrun = iterations > 0 ? totalCostOverrun / iterations : 0;
  const avgDayOverrun = iterations > 0 ? totalDayOverrun / iterations : 0;
  const reserveBudgetGap = Math.max(0, p80Cost - targetBudget);
  const reserveDayGap = Math.max(0, p80Days - targetDays);
  const enrichedDriverStats = driverStats
    .map((item) => ({
      name: item.name,
      avgCostImpact: item.totalCostImpact / iterations,
      avgDelayImpact: item.totalDelayImpact / iterations,
      hitRate: item.hits / iterations
    }))
    .sort((a, b) => (b.avgCostImpact + b.avgDelayImpact * 1000) - (a.avgCostImpact + a.avgDelayImpact * 1000));

  return {
    iterations,
    targetBudget,
    targetDays,
    workPackages,
    riskEvents,
    costRuns,
    dayRuns,
    meanCost,
    meanDays,
    medianCost,
    medianDays,
    p10Cost,
    p50Cost,
    p80Cost,
    p90Cost,
    p10Days,
    p50Days,
    p80Days,
    p90Days,
    budgetSuccessRate,
    scheduleSuccessRate,
    combinedSuccessRate,
    avgCostOverrun,
    avgDayOverrun,
    reserveBudgetGap,
    reserveDayGap,
    driverStats: enrichedDriverStats
  };
}

function calculateRiskRegisterResult(state) {
  const view = state.ui?.riskRegisterView || {};
  const editSortBy = ["newest", "id"].includes(view.editSortBy) ? view.editSortBy : "newest";
  const sourceRisks = Array.isArray(state.riskRegister?.risks) && state.riskRegister.risks.length
    ? state.riskRegister.risks
    : initialState.riskRegister.risks;
  const risks = sourceRisks.map((risk) => {
    const financialImpact = Number(risk.financialImpact) || 0;
    const probabilityPercent = Number(risk.probabilityPercent) || 0;
    const likelihood = deriveRiskLikelihoodFromPercent(probabilityPercent, risk.likelihood || 1);
    const impact = Number(risk.impact) || 0;
    const expectedDamage = financialImpact * (probabilityPercent / 100);
    const qualitativeRiskValue = likelihood * impact;
    return {
      ...risk,
      financialImpact,
      probabilityPercent,
      likelihood,
      impact,
      expectedDamage,
      qualitativeRiskValue
    };
  });
  const editRisks = [...risks].sort((a, b) => {
    const riskNumber = (risk) => {
      const match = String(risk.id || "").match(/^R-(\d{4})$/i);
      return match ? Number(match[1]) || 0 : 0;
    };
    if (editSortBy === "id") {
      return riskNumber(a) - riskNumber(b) || String(a.id || "").localeCompare(String(b.id || ""), "de");
    }
    const aNo = riskNumber(a);
    const bNo = riskNumber(b);
    if (aNo !== bNo) return bNo - aNo;
    const aTs = Date.parse(a.createdAt || "");
    const bTs = Date.parse(b.createdAt || "");
    if (Number.isFinite(aTs) && Number.isFinite(bTs) && aTs !== bTs) return bTs - aTs;
    return String(b.id || "").localeCompare(String(a.id || ""), "de");
  });
  const totalExpectedDamage = risks.reduce((sum, risk) => sum + risk.expectedDamage, 0);
  const criticalCount = risks.filter((risk) => risk.qualitativeRiskValue >= 13).length;
  const activeCount = risks.filter((risk) => normalizeRiskStatusValue(risk.status) !== "geschlossen").length;
  const closedCount = risks.filter((risk) => String(risk.status || "").toLowerCase() === "geschlossen").length;
  const today = new Date().toISOString().slice(0, 10);
  const todayTs = parseRiskDateValue(today);
  const overdueCount = risks.filter((risk) => {
    const dueTs = parseRiskDateValue(risk.dueDate);
    return dueTs !== null && todayTs !== null && dueTs < todayTs && String(risk.status || "").toLowerCase() !== "geschlossen";
  }).length;
  const ownerOptions = [...new Set(risks.map((risk) => String(risk.owner || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, "de"));
  const categoryOptions = [...new Set(risks.map((risk) => String(risk.category || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, "de"));
  const searchTerm = String(view.search || "").trim().toLowerCase();
  const selectedStatus = String(view.status || "alle").toLowerCase();
  const selectedOwner = String(view.owner || "alle").trim().toLowerCase();
  const selectedCategory = String(view.category || "alle").trim().toLowerCase();
  const criticalOnly = Boolean(view.criticalOnly);
  const dueFrom = String(view.dueFrom || "").trim();
  const dueTo = String(view.dueTo || "").trim();
  const dueFromTs = parseRiskDateValue(dueFrom);
  const dueToTs = parseRiskDateValue(dueTo);

  const filteredRisks = risks
    .filter((risk) => {
      if (selectedStatus !== "alle" && String(risk.status || "").toLowerCase() !== selectedStatus) return false;
      if (selectedOwner !== "alle" && String(risk.owner || "").trim().toLowerCase() !== selectedOwner) return false;
      if (selectedCategory !== "alle" && String(risk.category || "").trim().toLowerCase() !== selectedCategory) return false;
      const riskDueTs = parseRiskDateValue(risk.dueDate);
      if ((dueFromTs !== null || dueToTs !== null) && riskDueTs === null) return false;
      if (dueFromTs !== null && riskDueTs !== null && riskDueTs < dueFromTs) return false;
      if (dueToTs !== null && riskDueTs !== null && riskDueTs > dueToTs) return false;
      if (criticalOnly && risk.qualitativeRiskValue < 13) return false;
      if (!searchTerm) return true;
      return [
        risk.id,
        risk.title,
        risk.category,
        risk.phase,
        risk.area,
        risk.owner,
        risk.description,
        risk.measures
      ].some((value) => String(value || "").toLowerCase().includes(searchTerm));
    })
    .sort((a, b) => {
      const aDueTs = parseRiskDateValue(a.dueDate);
      const bDueTs = parseRiskDateValue(b.dueDate);
      const overdueDelta = Number(Boolean(bDueTs !== null && todayTs !== null && bDueTs < todayTs && String(b.status || "").toLowerCase() !== "geschlossen")) - Number(Boolean(aDueTs !== null && todayTs !== null && aDueTs < todayTs && String(a.status || "").toLowerCase() !== "geschlossen"));
      if (overdueDelta !== 0) return overdueDelta;
      const scoreDelta = b.qualitativeRiskValue - a.qualitativeRiskValue;
      if (scoreDelta !== 0) return scoreDelta;
      return b.expectedDamage - a.expectedDamage;
    });

  const ownerStats = ownerOptions.map((owner) => {
    const ownerRisks = risks.filter((risk) => String(risk.owner || "").trim() === owner);
    return {
      owner,
      activeCount: ownerRisks.filter((risk) => String(risk.status || "").toLowerCase() !== "geschlossen").length,
      criticalCount: ownerRisks.filter((risk) => risk.qualitativeRiskValue >= 13).length,
      overdueCount: ownerRisks.filter((risk) => {
        const dueTs = parseRiskDateValue(risk.dueDate);
        return dueTs !== null && todayTs !== null && dueTs < todayTs && String(risk.status || "").toLowerCase() !== "geschlossen";
      }).length
    };
  });

  return {
    risks,
    filteredRisks,
    totalExpectedDamage,
    criticalCount,
    activeCount,
    closedCount,
    overdueCount,
    ownerOptions,
    categoryOptions,
    ownerStats,
    editRisks,
    filterSummary: {
      search: view.search || "",
      status: view.status || "alle",
      owner: view.owner || "alle",
      category: view.category || "alle",
      dueFrom: view.dueFrom || "",
      dueTo: view.dueTo || "",
      criticalOnly
    }
  };
}

function getRiskLevel(score) {
  if (score >= 0.65) return "stabil";
  if (score >= 0.4) return "angespannt";
  return "kritisch";
}

export function buildManagementReportData(state) {
  const pert = calculatePertResult(state);
  const monteCarlo = calculateMonteCarloResult(state);
  const earnedValue = calculateEarnedValueResult(state);
  const riskRegister = calculateRiskRegisterResult(state);
  const level = getRiskLevel(monteCarlo.combinedSuccessRate);
  const timeContext = calculateProjectTimeContext(state.project);
  const location = state.project?.location || {};
  const addressPrefix = [location.street, location.houseNumber].filter(Boolean).join(" ");
  const projectAddress = [addressPrefix, [location.postalCode, location.city].filter(Boolean).join(" ")].filter(Boolean).join(", ");

  const topline = {
    projectName: state.project.name,
    type: state.project.type,
    bauart: state.project.bauart,
    phase: state.project.phase,
    budget: state.project.budget,
    client: state.project.client,
    projectLead: state.project.projectLead,
    projectAddress: projectAddress || "nicht angegeben",
    startDate: formatDate(state.project.startDate),
    endDate: formatDate(state.project.endDate),
    reportDate: formatDate(state.project.analysisDate || new Date().toISOString()),
    analysisDate: formatDate(state.project.analysisDate || new Date().toISOString()),
    scheduleStatus: timeContext.status,
    level
  };

  const summaryCards = [
    {
      label: "Zeitliche Einordnung",
      value: timeContext.status,
      detail: `${formatDate(state.project.analysisDate)}\n${timeContext.progressLabel}`,
      tone: getProjectTimeTone(timeContext.status)
    },
    {
      label: "PERT-Erwartungswert",
      value: formatCurrency(pert.expectedValue),
      detail: `Budget 84 %: ${formatCurrency(pert.budget84)}`,
      linkModule: "pert",
      linkTarget: "pert-expected-value-card",
      linkLabel: "PERT-Wert im PERT-Modul anzeigen"
    },
    {
      label: "Monte-Carlo Zielerreichung",
      value: formatPercent(monteCarlo.combinedSuccessRate),
      detail: `P80 Kosten: ${formatCurrency(monteCarlo.p80Cost)}`,
      linkModule: "monteCarlo",
      linkTarget: "monte-p80-costs",
      linkLabel: "P80 Kosten im Monte-Carlo-Block anzeigen"
    },
    {
      label: "Leistungswert-Analyse (EVA) · Kosten- und Terminleistungsindex",
      value: `${earnedValue.cpi.toFixed(2)} / ${earnedValue.spi.toFixed(2)}`,
      detail: `CPI ${earnedValue.cpi.toFixed(2)} = Kostenleistung\nSPI ${earnedValue.spi.toFixed(2)} = Terminleistung\nEAC: ${formatCurrency(earnedValue.eac)}`,
      linkModule: "earnedValue",
      linkTarget: "eva-kpi-overview",
      linkLabel: "CPI / SPI / EAC im EVA-Modul anzeigen"
    },
    {
      label: "Erwarteter Risikoschaden",
      value: formatCurrency(riskRegister.totalExpectedDamage),
      detail: `${riskRegister.criticalCount} kritische Risiken\nSumme der erwarteten Schäden aus dem Register`,
      linkModule: "riskRegister",
      linkTarget: "risk-register-summary",
      linkLabel: "Risikoschaden im Register anzeigen"
    }
  ];

  const topRisks = [...riskRegister.risks]
    .sort((a, b) => b.qualitativeRiskValue - a.qualitativeRiskValue || b.expectedDamage - a.expectedDamage)
    .slice(0, 3)
    .map((risk) => ({
      id: risk.id,
      title: risk.title,
      owner: risk.owner || "nicht zugewiesen",
      status: risk.status || "offen",
      expectedDamage: formatCurrency(risk.expectedDamage),
      score: `${risk.qualitativeRiskValue} / 25`
    }));

  const focusPoints = [];
  if (pert.budget84 > state.project.budget) {
    focusPoints.push(`Die PERT-Betrachtung liegt mit dem 84 %-Budget von ${formatCurrency(pert.budget84)} über dem aktuellen Projektrahmen.`);
  } else {
    focusPoints.push("Die PERT-Betrachtung liegt innerhalb des aktuellen Projektrahmens und zeigt derzeit keine direkte Budgetüberschreitung im 84 %-Niveau.");
  }

  if (monteCarlo.combinedSuccessRate < 0.4) {
    focusPoints.push(`Die kombinierte Monte-Carlo-Zielerreichung liegt nur bei ${formatPercent(monteCarlo.combinedSuccessRate)} und ist damit derzeit kritisch.`);
  } else if (monteCarlo.combinedSuccessRate < 0.65) {
    focusPoints.push(`Die Monte-Carlo-Zielerreichung liegt bei ${formatPercent(monteCarlo.combinedSuccessRate)} und erfordert aktive Steuerung von Budget, Termin und Reserven.`);
  } else {
    focusPoints.push(`Die Monte-Carlo-Zielerreichung liegt bei ${formatPercent(monteCarlo.combinedSuccessRate)} und ist aktuell vergleichsweise robust.`);
  }

  if (earnedValue.cpi < 0.95 || earnedValue.spi < 0.95) {
    focusPoints.push(`Die EVA-Indikatoren mit Kostenleistungsindex (CPI) ${earnedValue.cpi.toFixed(2)} und Terminleistungsindex (SPI) ${earnedValue.spi.toFixed(2)} zeigen eine laufende Kosten- und/oder Terminspannung.`);
  } else {
    focusPoints.push(`Die EVA-Indikatoren mit Kostenleistungsindex (CPI) ${earnedValue.cpi.toFixed(2)} und Terminleistungsindex (SPI) ${earnedValue.spi.toFixed(2)} liegen im steuerbaren Bereich.`);
  }

  if (riskRegister.criticalCount > 0 || riskRegister.overdueCount > 0) {
    focusPoints.push(`Im Risikoregister sind ${riskRegister.criticalCount} kritische und ${riskRegister.overdueCount} überfällige Risiken sichtbar. Maßnahmen- und Terminverfolgung sollten priorisiert werden.`);
  } else {
    focusPoints.push("Das Risikoregister zeigt aktuell keine kritischen oder überfälligen Risiken mit unmittelbarem Eskalationsbedarf.");
  }

  focusPoints.push(`Der Betrachtungszeitpunkt liegt in der Einordnung "${timeContext.status}". ${timeContext.detail}.`);

  const nextSteps = [];
  if (monteCarlo.p80Cost > state.project.budget) {
    nextSteps.push(`Budgetreserve überprüfen und Projektrahmen gegen das P80-Niveau von ${formatCurrency(monteCarlo.p80Cost)} spiegeln.`);
  }
  if (earnedValue.cpi < 1) {
    nextSteps.push("Kostenabweichungen der laufenden Arbeitspakete in EVA gezielt nachsteuern und Verantwortlichkeiten schärfen.");
  }
  if (riskRegister.criticalCount > 0) {
    nextSteps.push("Kritische Risiken mit Maßnahmen, Termin und Owner in der nächsten Steuerungsrunde einzeln durchgehen.");
  }
  if (!nextSteps.length) {
    nextSteps.push("Monatliche Fortschreibung von PERT, EVA, Monte Carlo und Risikoregister beibehalten und Berichtswesen weiter vereinheitlichen.");
  }

  return {
    topline,
    summaryCards,
    focusPoints,
    topRisks,
    nextSteps,
    timeContext,
    raw: {
      pert,
      monteCarlo,
      earnedValue,
      riskRegister
    }
  };
}

export function renderManagementReportText(state) {
  const report = buildManagementReportData(state);
  return [
    `Management-Zusammenfassung für ${report.topline.projectName}`,
    `Berichtsdatum: ${report.topline.reportDate}`,
    `Betrachtungszeitpunkt: ${report.topline.analysisDate}`,
    `Zeitliche Einordnung: ${report.topline.scheduleStatus}`,
    `Projektart: ${report.topline.type}`,
    `Bauart: ${report.topline.bauart}`,
    `Leistungsphase: ${report.topline.phase}`,
    `Projektstandort: ${report.topline.projectAddress}`,
    `Auftraggeber: ${report.topline.client}`,
    `Projektleitung: ${report.topline.projectLead}`,
    `Baubeginn: ${report.topline.startDate}`,
    `Geplante Fertigstellung: ${report.topline.endDate}`,
    "",
    "Executive Summary",
    ...report.focusPoints.map((item) => `- ${item}`),
    "",
    "Risikobericht",
    `- Erwarteter Gesamtschaden: ${formatCurrency(report.raw.riskRegister.totalExpectedDamage)}`,
    `- Kritische Risiken: ${report.raw.riskRegister.criticalCount}`,
    `- Aktive Risiken: ${report.raw.riskRegister.activeCount}`,
    `- Überfällige Risiken: ${report.raw.riskRegister.overdueCount}`,
    ...(report.topRisks.length
      ? report.topRisks.map((risk) => `- ${risk.id} ${risk.title} | ${risk.score} | ${risk.expectedDamage} | Verantwortlich: ${risk.owner} | Status: ${risk.status}`)
      : ["- Keine priorisierten Risiken im Register."]),
    "",
    "Top-Risiken",
    ...(report.topRisks.length
      ? report.topRisks.map((risk) => `- ${risk.id} ${risk.title} | ${risk.score} | ${risk.expectedDamage} | Verantwortlich: ${risk.owner} | Status: ${risk.status}`)
      : ["- Keine priorisierten Risiken im Register."]),
    "",
    "Nächste Schritte",
    ...report.nextSteps.map((item) => `- ${item}`)
  ].join("\n");
}

export function renderRiskReportText(state) {
  const report = buildManagementReportData(state);
  const sections = getActiveReportSections(state);
  const selectedModules = [];
  if (sections.includePert) selectedModules.push("PERT");
  if (sections.includeMonteCarlo) selectedModules.push("Monte Carlo");
  if (sections.includeEarnedValue) selectedModules.push("EVA");
  if (sections.includeRiskRegister) selectedModules.push("Risikoregister");
  const lines = [
    `Risikobericht für ${report.topline.projectName}`,
    `Berichtsdatum: ${report.topline.reportDate}`,
    `Betrachtungszeitpunkt: ${report.topline.analysisDate}`,
    `Zeitliche Einordnung: ${report.topline.scheduleStatus}`,
    `Ausgewählte Module: ${selectedModules.length ? selectedModules.join(" · ") : "keine"}`,
    "",
    "Lagebild",
    `- Aktive Risiken: ${report.raw.riskRegister.activeCount}`,
    `- Kritische Risiken: ${report.raw.riskRegister.criticalCount}`,
    `- Überfällige Risiken: ${report.raw.riskRegister.overdueCount}`,
    `- Erwarteter Gesamtschaden: ${formatCurrency(report.raw.riskRegister.totalExpectedDamage)}`,
  ];

  if (sections.includePert) {
    lines.push(
      "",
      "PERT",
      `- Erwartungswert: ${formatCurrency(report.raw.pert.expectedValue)}`,
      `- 84 %-Budget: ${formatCurrency(report.raw.pert.budget84)}`,
      `- 97,5 %-Budget: ${formatCurrency(report.raw.pert.budget975)}`
    );
  }

  if (sections.includeMonteCarlo) {
    lines.push(
      "",
      "Monte Carlo",
      `- Erwartete Kosten: ${formatCurrency(report.raw.monteCarlo.meanCost)}`,
      `- Erwartete Dauer: ${formatNumber(report.raw.monteCarlo.meanDays)} Tage`,
      `- P80 Kosten: ${formatCurrency(report.raw.monteCarlo.p80Cost)}`,
      `- P80 Dauer: ${formatNumber(report.raw.monteCarlo.p80Days)} Tage`,
      `- Zielerreichung gesamt: ${formatPercent(report.raw.monteCarlo.combinedSuccessRate)}`
    );
  }

  if (sections.includeEarnedValue) {
    lines.push(
      "",
      "Leistungswert-Analyse (EVA)",
      `- Planwert (PV): ${formatCurrency(report.raw.earnedValue.pv)}`,
      `- Leistungswert (EV): ${formatCurrency(report.raw.earnedValue.ev)}`,
      `- Istkosten (AC): ${formatCurrency(report.raw.earnedValue.ac)}`,
      `- Kostenleistungsindex (CPI): ${report.raw.earnedValue.cpi.toFixed(2)}`,
      `- Terminleistungsindex (SPI): ${report.raw.earnedValue.spi.toFixed(2)}`,
      `- Kostenprognose (EAC): ${formatCurrency(report.raw.earnedValue.eac)}`
    );
  }

  if (sections.includeRiskRegister) {
    lines.push(
      "",
      "Risikoregister",
      `- Kritische Risiken: ${report.raw.riskRegister.criticalCount}`,
      `- Überfällige Risiken: ${report.raw.riskRegister.overdueCount}`,
      `- Erwarteter Gesamtschaden: ${formatCurrency(report.raw.riskRegister.totalExpectedDamage)}`,
      "",
      "Top-Risiken",
      ...(report.topRisks.length
        ? report.topRisks.map((risk) => `- ${risk.id} ${risk.title} | ${risk.score} | ${risk.expectedDamage} | Verantwortlich: ${risk.owner} | Status: ${risk.status}`)
        : ["- Derzeit keine priorisierten Risiken im Register."]),
      "",
      "Maßnahmen",
      ...(report.raw.riskRegister.activeCount
        ? [
            `- Kritische Risiken sofort priorisieren und mit Verantwortlichen abstimmen.`,
            `- Überfällige Risiken mit Zieltermin und Maßnahmenplanung nachschärfen.`,
            `- Markierte Risiken bei Bedarf in Monte Carlo übernehmen.`
          ]
        : ["- Für den aktuellen Stand sind keine operativen Maßnahmen erforderlich."]),
      "",
      "Restgefahr",
      ...(report.raw.riskRegister.activeCount
        ? [
            `- Die Restgefahr bleibt sichtbar, solange Risiken offen oder in Bearbeitung sind.`,
            `- Überfällige Risiken erhöhen die Restgefahr zusätzlich.`,
            `- Residual-Risiken werden im Register unter den jeweiligen Einzelrisiken gepflegt.`
          ]
        : ["- Aktuell keine Restgefahr aus aktiven Risiken erkennbar."]),
      "",
      "Nächste Prioritäten",
      ...(report.raw.riskRegister.criticalCount > 0
        ? [
            "Kritische Risiken direkt mit Maßnahmen und Termin klären.",
            "Markierte Risiken bei Bedarf in Monte Carlo übernehmen.",
            "Risikoregister und Berichtsausgabe im nächsten Zyklus aktualisieren."
          ]
        : report.nextSteps.map((item) => `- ${item}`))
    );
  }

  return lines.join("\n");
}

function getActiveReportSections(state) {
  return {
    includePert: state.ui?.reportOptions?.includePert !== false,
    includeMonteCarlo: state.ui?.reportOptions?.includeMonteCarlo !== false,
    includeEarnedValue: state.ui?.reportOptions?.includeEarnedValue !== false,
    includeRiskRegister: state.ui?.reportOptions?.includeRiskRegister !== false
  };
}

export function buildSelectedReportData(state) {
  const report = buildManagementReportData(state);
  const sections = getActiveReportSections(state);
  const selectedModules = [];
  if (sections.includePert) selectedModules.push("PERT");
  if (sections.includeMonteCarlo) selectedModules.push("Monte Carlo");
  if (sections.includeEarnedValue) selectedModules.push("EVA");
  if (sections.includeRiskRegister) selectedModules.push("Risikoregister");

  const payload = {
    meta: {
      app: initialState.meta.app,
      version: initialState.meta.version,
      reportMode: state.ui?.reportMode === "management" ? "management" : "risk",
      reportTitle: state.ui?.reportMode === "management" ? "Managementbericht" : "Risikobericht",
      generatedAt: new Date().toISOString()
    },
    project: {
      name: report.topline.projectName,
      type: report.topline.type,
      bauart: report.topline.bauart,
      phase: report.topline.phase,
      projectAddress: report.topline.projectAddress,
      client: report.topline.client,
      projectLead: report.topline.projectLead,
      reportDate: report.topline.reportDate,
      analysisDate: report.topline.analysisDate,
      startDate: report.topline.startDate,
      endDate: report.topline.endDate,
      scheduleStatus: report.topline.scheduleStatus
    },
    report: {
      timeContext: report.timeContext,
      focusPoints: report.focusPoints,
      nextSteps: report.nextSteps,
      selectedModules,
      text: state.ui?.reportMode === "management" ? renderSelectedReportText(state) : renderRiskReportText(state)
    },
    modules: {}
  };

  if (sections.includePert) {
    payload.modules.pert = {
      expectedValue: report.raw.pert.expectedValue,
      standardDeviation: report.raw.pert.standardDeviation,
      budget84: report.raw.pert.budget84,
      budget975: report.raw.pert.budget975,
      items: report.raw.pert.items
    };
  }

  if (sections.includeMonteCarlo) {
    payload.modules.monteCarlo = {
      meanCost: report.raw.monteCarlo.meanCost,
      meanDays: report.raw.monteCarlo.meanDays,
      p80Cost: report.raw.monteCarlo.p80Cost,
      p80Days: report.raw.monteCarlo.p80Days,
      budgetSuccessRate: report.raw.monteCarlo.budgetSuccessRate,
      scheduleSuccessRate: report.raw.monteCarlo.scheduleSuccessRate,
      combinedSuccessRate: report.raw.monteCarlo.combinedSuccessRate
    };
  }

  if (sections.includeEarnedValue) {
    payload.modules.earnedValue = {
      pv: report.raw.earnedValue.pv,
      ev: report.raw.earnedValue.ev,
      ac: report.raw.earnedValue.ac,
      cv: report.raw.earnedValue.cv,
      sv: report.raw.earnedValue.sv,
      cpi: report.raw.earnedValue.cpi,
      spi: report.raw.earnedValue.spi,
      eac: report.raw.earnedValue.eac,
      tcpi: report.raw.earnedValue.tcpi
    };
  }

  if (sections.includeRiskRegister) {
    payload.modules.riskRegister = {
      totalExpectedDamage: report.raw.riskRegister.totalExpectedDamage,
      criticalCount: report.raw.riskRegister.criticalCount,
      activeCount: report.raw.riskRegister.activeCount,
      overdueCount: report.raw.riskRegister.overdueCount,
      topRisks: report.topRisks
    };
  }

  return payload;
}

export function renderSelectedReportText(state) {
  const report = buildManagementReportData(state);
  const sections = getActiveReportSections(state);
  const lines = [
    `Projektbericht für ${report.topline.projectName}`,
    `Berichtsdatum: ${report.topline.reportDate}`,
    `Betrachtungszeitpunkt: ${report.topline.analysisDate}`,
    `Zeitliche Einordnung: ${report.topline.scheduleStatus}`,
    `Projektart: ${report.topline.type}`,
    `Bauart: ${report.topline.bauart}`,
    `Leistungsphase: ${report.topline.phase}`,
    `Projektstandort: ${report.topline.projectAddress}`,
    `Auftraggeber: ${report.topline.client}`,
    `Projektleitung: ${report.topline.projectLead}`,
    `Baubeginn: ${report.topline.startDate}`,
    `Geplante Fertigstellung: ${report.topline.endDate}`,
    "",
    "Executive Summary",
    ...report.focusPoints.map((item) => `- ${item}`)
  ];

  if (sections.includePert) {
    lines.push(
      "",
      "PERT",
      `- Erwartungswert: ${formatCurrency(report.raw.pert.expectedValue)}`,
      `- Standardabweichung: ${formatCurrency(report.raw.pert.standardDeviation)}`,
      `- Budget 84 %: ${formatCurrency(report.raw.pert.budget84)}`,
      `- Budget 97,5 %: ${formatCurrency(report.raw.pert.budget975)}`
    );
  }

  if (sections.includeMonteCarlo) {
    lines.push(
      "",
      "Monte Carlo",
      `- Erwartete Kosten: ${formatCurrency(report.raw.monteCarlo.meanCost)}`,
      `- Erwartete Dauer: ${formatNumber(report.raw.monteCarlo.meanDays)} Tage`,
      `- P80 Kosten: ${formatCurrency(report.raw.monteCarlo.p80Cost)}`,
      `- P80 Dauer: ${formatNumber(report.raw.monteCarlo.p80Days)} Tage`,
      `- Zielerreichung gesamt: ${formatPercent(report.raw.monteCarlo.combinedSuccessRate)}`
    );
  }

  if (sections.includeEarnedValue) {
    lines.push(
      "",
      "Leistungswert-Analyse (EVA)",
      `- Planwert (PV): ${formatCurrency(report.raw.earnedValue.pv)}`,
      `- Leistungswert (EV): ${formatCurrency(report.raw.earnedValue.ev)}`,
      `- Istkosten (AC): ${formatCurrency(report.raw.earnedValue.ac)}`,
      `- Kostenleistungsindex (CPI): ${report.raw.earnedValue.cpi.toFixed(2)}`,
      `- Terminleistungsindex (SPI): ${report.raw.earnedValue.spi.toFixed(2)}`,
      `- Kostenprognose (EAC): ${formatCurrency(report.raw.earnedValue.eac)}`
    );
  }

  if (sections.includeRiskRegister) {
    lines.push(
      "",
      "Risikoregister",
      `- Erwarteter Gesamtschaden: ${formatCurrency(report.raw.riskRegister.totalExpectedDamage)}`,
      `- Kritische Risiken: ${report.raw.riskRegister.criticalCount}`,
      `- Aktive Risiken: ${report.raw.riskRegister.activeCount}`,
      `- Überfällige Risiken: ${report.raw.riskRegister.overdueCount}`,
      ...(report.topRisks.length
        ? report.topRisks.map((risk) => `- ${risk.id} ${risk.title} | ${risk.score} | ${risk.expectedDamage}`)
        : ["- Keine priorisierten Risiken im Register."])
    );
  }

  lines.push("", "Nächste Schritte", ...report.nextSteps.map((item) => `- ${item}`));
  return lines.join("\n");
}

export const modules = {
  project: {
    key: "project",
    label: "Projekt",
    subtitle: "Gemeinsame Stammdaten für alle Module",
    render(state) {
      const project = state.project;
      const report = state.reportProfile;
      const timeContext = calculateProjectTimeContext(project);
      const addressLine = [project.location?.street, project.location?.houseNumber].filter(Boolean).join(" ");
      return `
        <div class="module-shell pert-module">
          <div class="module-title project-module-title">
            <div class="project-module-intro">
              <h2>Projekt</h2>
              <p><span>Dieses Modul hält die gemeinsamen Stammdaten zentral. PERT-Methode, Monte-Carlo-Simulation,</span><span>Leistungswert-Analyse (EVA) und Risikoregister lesen diesen Kontext gemeinsam, statt ihn mehrfach zu pflegen.</span></p>
            </div>
            <span class="badge">Zentrale Datenbasis</span>
          </div>
          <div class="kpi-grid project-kpi-grid">
            <article class="kpi-card project-name-card">
              <div class="kpi-label">Projektname</div>
              <div class="kpi-value">${project.name}</div>
              <div class="kpi-sub">${project.type}</div>
            </article>
            <article class="kpi-card gold project-phase-card">
              <div class="kpi-label">Leistungsphase</div>
              <div class="kpi-value">${project.phase}</div>
              <div class="kpi-sub">BGF: ${new Intl.NumberFormat("de-DE").format(project.bgf)} m²</div>
            </article>
            <article class="kpi-card green project-budget-card">
              <div class="kpi-label">Projektbudget</div>
              <div class="kpi-value">${formatCurrency(project.budget)}</div>
              <div class="kpi-sub">Kostenbasis: ${project.costBasis}</div>
            </article>
            <article class="kpi-card blue project-time-card">
              <div class="kpi-label">Betrachtungs-<br>zeitpunkt</div>
              <div class="kpi-value">${formatDate(project.analysisDate)}</div>
              <div class="kpi-sub">${timeContext.progressLabel}</div>
            </article>
          </div>
          <div class="card-grid pert-upper-grid">
            <section class="info-card">
              <h3>Projektkontext</h3>
              <ul>
                <li><strong>Bauart:</strong> ${project.bauart}</li>
                <li><strong>Standort:</strong> ${project.location.street} ${project.location.houseNumber}, ${project.location.postalCode} ${project.location.city}</li>
                <li><strong>Auftraggeber:</strong> ${project.client}</li>
                <li><strong>Projektleitung:</strong> ${project.projectLead}</li>
                <li><strong>Baubeginn:</strong> ${formatDate(project.startDate)}</li>
                <li><strong>Geplante Fertigstellung:</strong> ${formatDate(project.endDate)}</li>
                <li><strong>Betrachtungszeitpunkt:</strong> ${formatDate(project.analysisDate)}</li>
              </ul>
            </section>
            <section class="info-card">
              <h3>Zeitliche Einordnung</h3>
              <ul>
                <li><strong>Status:</strong> ${timeContext.status}</li>
                <li><strong>Einordnung:</strong> ${timeContext.detail}</li>
                <li>Diese Datumslogik steht später auch Berichten und KI-Auswertungen gemeinsam zur Verfügung.</li>
              </ul>
            </section>
          </div>
          <section class="info-card">
            <h3>Projektdaten bearbeiten</h3>
            <div class="form-grid project-form-grid">
              <div class="form-field wide">
                <label for="project_name">Projektname</label>
                <input id="project_name" data-project-field="name" type="text" value="${project.name}">
              </div>
              <div class="form-field">
                <label for="project_type">Projekttyp</label>
                <input id="project_type" data-project-field="type" type="text" value="${project.type}">
              </div>
              <div class="form-field">
                <label for="project_bauart">Bauart</label>
                <select id="project_bauart" data-project-field="bauart">
                  <option value="Neubau" ${project.bauart === "Neubau" ? "selected" : ""}>Neubau</option>
                  <option value="Sanierung" ${project.bauart === "Sanierung" ? "selected" : ""}>Sanierung</option>
                  <option value="Umbau" ${project.bauart === "Umbau" ? "selected" : ""}>Umbau</option>
                  <option value="Erweiterung" ${project.bauart === "Erweiterung" ? "selected" : ""}>Erweiterung</option>
                </select>
              </div>
              <div class="form-field">
                <label for="project_phase">Leistungsphase</label>
                <select id="project_phase" data-project-field="phase">
                  ${["LPH 1","LPH 2","LPH 3","LPH 4","LPH 5","LPH 6","LPH 7","LPH 8"].map((phase) => `
                    <option value="${phase}" ${project.phase === phase ? "selected" : ""}>${phase}</option>
                  `).join("")}
                </select>
              </div>
              <div class="form-field wide">
                <div class="project-inline-grid project-inline-grid-address">
                  <div class="form-field">
                    <label for="project_address_line">Projektstandort: Straße / Hausnummer</label>
                    <input id="project_address_line" data-project-address-field="combined" type="text" value="${addressLine}">
                  </div>
                  <div class="form-field">
                    <label for="project_postal_code">PLZ</label>
                    <input id="project_postal_code" data-project-location-field="postalCode" type="text" value="${project.location.postalCode}">
                  </div>
                  <div class="form-field">
                    <label for="project_city">Ort</label>
                    <input id="project_city" data-project-location-field="city" type="text" value="${project.location.city}">
                  </div>
                </div>
              </div>
              <div class="form-field">
                <label for="project_bgf">BGF in m²</label>
                <input id="project_bgf" class="numeric-project-input" data-project-field="bgf" type="text" inputmode="numeric" value="${formatIntegerInput(project.bgf, "m²")}">
              </div>
              <div class="form-field">
                <label for="project_budget">Projektbudget in EUR</label>
                <input id="project_budget" class="numeric-project-input" data-project-field="budget" type="text" inputmode="numeric" value="${formatCurrencyInput(project.budget)}">
              </div>
              <div class="form-field">
                <label for="project_cost_basis">Kostenbasis</label>
                <select id="project_cost_basis" data-project-field="costBasis">
                  <option value="netto" ${project.costBasis === "netto" ? "selected" : ""}>Netto</option>
                  <option value="brutto" ${project.costBasis === "brutto" ? "selected" : ""}>Brutto</option>
                </select>
              </div>
              <div class="form-field">
                <label for="project_client">Auftraggeber</label>
                <input id="project_client" data-project-field="client" type="text" value="${project.client}">
              </div>
              <div class="form-field">
                <label for="project_lead">Projektleitung</label>
                <input id="project_lead" data-project-field="projectLead" type="text" value="${project.projectLead}">
              </div>
              <div class="form-field wide">
                <div class="project-inline-grid project-inline-grid-3">
                  <div class="form-field">
                    <label for="project_start">Projektstart</label>
                    <input id="project_start" data-project-field="startDate" type="date" value="${project.startDate}">
                  </div>
                  <div class="form-field">
                    <label for="project_end">Geplante Fertigstellung</label>
                    <input id="project_end" data-project-field="endDate" type="date" value="${project.endDate}">
                  </div>
                  <div class="form-field">
                    <label for="project_analysis_date">Betrachtungszeitpunkt</label>
                    <input id="project_analysis_date" data-project-field="analysisDate" type="date" value="${project.analysisDate || ""}">
                  </div>
                </div>
              </div>
              <div class="form-field wide">
                <label for="project_description">Projektbeschreibung</label>
                <textarea id="project_description" data-project-field="description">${project.description}</textarea>
              </div>
            </div>
            <p class="form-note">Diese Angaben bilden die gemeinsame Projektbasis für PERT-Methode, Monte-Carlo-Simulation, Leistungswert-Analyse (EVA), Risikoregister sowie die spätere KI- und Berichtseinordnung des Projektstands.</p>
          </section>
          <div class="card-grid">
            <section class="info-card">
              <h3>Berichtsprofil</h3>
              <div class="form-grid">
                <div class="form-field">
                  <label for="report_company">Firma / Büro</label>
                  <input id="report_company" data-report-field="company" type="text" value="${report.company}">
                </div>
                <div class="form-field">
                  <label for="report_author">Erstellt von</label>
                  <input id="report_author" data-report-field="author" type="text" value="${report.author}">
                </div>
                <div class="form-field wide">
                  <label for="report_company_address">Firmenanschrift</label>
                  <input id="report_company_address" data-report-field="companyAddress" type="text" value="${report.companyAddress}">
                </div>
                <div class="form-field">
                  <label for="report_client_name">Empfänger / Auftraggeber</label>
                  <input id="report_client_name" data-report-field="clientName" type="text" value="${report.clientName}">
                </div>
                <div class="form-field">
                  <label for="report_project_address">Projektadresse</label>
                  <input id="report_project_address" data-report-field="projectAddress" type="text" value="${report.projectAddress}">
                </div>
                <div class="form-field wide">
                  <label for="report_confidentiality">Vertraulichkeit</label>
                  <select id="report_confidentiality" data-report-field="confidentiality">
                    ${["Intern", "Vertraulich", "Streng vertraulich"].map((option) => `
                      <option value="${option}" ${String(report.confidentiality || "Vertraulich") === option ? "selected" : ""}>${option}</option>
                    `).join("")}
                  </select>
                </div>
                <div class="form-field wide">
                  <label for="report_notes">Notizen</label>
                  <textarea id="report_notes" data-report-field="notes">${report.notes}</textarea>
                </div>
              </div>
            </section>
            <section class="info-card">
              <h3>Zentraler Nutzen</h3>
              <ul>
                <li>Nur noch ein Berichtprofil pro Projekt</li>
                <li>PDF, Word und Excel greifen auf denselben Kontext zu</li>
                <li>Spätere Modulberichte müssen keine eigenen Profilfelder mehr pflegen</li>
              </ul>
            </section>
          </div>
        </div>
      `;
    }
  },
  pert: {
    key: "pert",
    label: "PERT-Methode",
    subtitle: "Frühe Schätzung und Sicherheitsniveaus",
    render(state) {
      const result = calculatePertResult(state);
      const bgf = Number(state.project.bgf) || 0;
      const selectedPertIds = state.ui?.transferSelections?.pertItemIds || [];
      const analytics = calculatePertAnalytics(state, result);
      const latestSnapshot = state.pert.snapshots?.[0] || null;
      const securityLabels = {
        0: "50 % · Erwartungswert",
        1: "84 % · Vorsichtiges Niveau",
        2: "97,5 % · Konservatives Niveau"
      };
      return `
        <div class="module-shell pert-module">
          <div class="module-title">
            <div>
              <h2>PERT-Methode</h2>
              <p>Drei-Punkt-Schätzung, Budgetpfade und Fortschreibung nach Leistungsphasen. Die Gewerke bleiben bewusst von den Monte-Carlo-Arbeitspaketen getrennt.</p>
            </div>
            <span class="badge">Schätzung</span>
          </div>
          <div class="kpi-grid">
            <article class="kpi-card" id="pert-expected-value-card"><div class="kpi-label">Erwartungswert</div><div class="kpi-value">${formatCurrency(result.expectedValue)}</div><div class="kpi-sub">EW aus O / M / P</div></article>
            <article class="kpi-card gold"><div class="kpi-label">Budget 84%</div><div class="kpi-value">${formatCurrency(result.budget84)}</div><div class="kpi-sub">vorsichtiges Niveau</div></article>
            <article class="kpi-card green"><div class="kpi-label">Budget 97,5%</div><div class="kpi-value">${formatCurrency(result.budget975)}</div><div class="kpi-sub">konservatives Niveau</div></article>
            <article class="kpi-card blue"><div class="kpi-label">Aktiver Budgetpfad</div><div class="kpi-value">${formatCurrency(result.selectedBudget)}</div><div class="kpi-sub">${securityLabels[Number(state.pert.securityLevel)] || "84 % · Vorsichtiges Niveau"}</div></article>
          </div>
          <div class="card-grid pert-upper-grid">
            <section class="info-card control-card">
              <h3>Steuerung</h3>
              <div class="signal-band">
                <div class="signal-chip">
                  <span>Aktiver Pfad</span>
                  <strong>${securityLabels[Number(state.pert.securityLevel)] || "84 % · Vorsichtiges Niveau"}</strong>
                </div>
                <div class="signal-chip">
                  <span>Budget je m²</span>
                  <strong>${bgf > 0 ? `${formatNumber(result.selectedBudget / bgf)} €/m²` : "BGF fehlt"}</strong>
                </div>
                <div class="signal-chip">
                  <span>Abgleich</span>
                  <strong>${result.budgetDelta >= 0 ? `+ ${formatCurrency(result.budgetDelta)}` : `- ${formatCurrency(Math.abs(result.budgetDelta))}`}</strong>
                </div>
              </div>
              <div class="form-grid">
                <div class="form-field wide">
                  <label for="pert_security_level">Sicherheitsniveau</label>
                  <select id="pert_security_level" data-pert-field="securityLevel">
                    <option value="0" ${Number(state.pert.securityLevel) === 0 ? "selected" : ""}>50 % · Erwartungswert</option>
                    <option value="1" ${Number(state.pert.securityLevel) === 1 ? "selected" : ""}>84 % · Vorsichtiges Niveau</option>
                    <option value="2" ${Number(state.pert.securityLevel) === 2 ? "selected" : ""}>97,5 % · Konservatives Niveau</option>
                  </select>
                </div>
                <div class="form-field">
                  <label>Aktiver Budgetpfad</label>
                  <div class="form-note form-note-strong">${formatCurrency(result.selectedBudget)}</div>
                </div>
                <div class="form-field">
                  <label>Aktive Leistungsphase</label>
                  <div class="form-note form-note-strong">${state.project?.phase || "LPH 1"}</div>
                </div>
                <div class="form-field">
                  <label>Erfasste Gewerke</label>
                  <div class="form-note form-note-strong">${result.items.length} Positionen in der Schätzung</div>
                </div>
                <div class="form-field">
                  <label>Letzter Snapshot</label>
                  <div class="form-note form-note-strong">${latestSnapshot ? `${latestSnapshot.phase} · ${latestSnapshot.savedAt}` : "Noch kein Snapshot gespeichert"}</div>
                </div>
              </div>
              <div class="control-actions">
                <button class="action-btn" type="button" data-action="add-pert-item">Gewerk hinzufügen</button>
                <button class="action-btn" type="button" data-action="transfer-pert-to-mc">Nach Monte Carlo übernehmen</button>
                <button class="action-btn" type="button" data-action="transfer-selected-pert-to-mc">Markierte nach Monte Carlo anhängen</button>
              </div>
              <p class="form-note" style="margin-top:10px;">Markiert: ${selectedPertIds.length} Positionen</p>
            </section>
            <section class="info-card steering-card">
              <h3>Steuerungsbild</h3>
              <ul>
                <li><strong>Aktive Leistungsphase:</strong> ${state.project?.phase || "LPH 1"}</li>
                <li><strong>Standardabweichung:</strong> ${formatCurrency(result.standardDeviation)}</li>
                <li><strong>Globales Projektbudget:</strong> ${formatCurrency(Number(state.project.budget) || 0)}</li>
                <li>Snapshots werden zentral im Verlaufsblock gespeichert und von dort für Fortschreibung und Trendanalyse genutzt.</li>
              </ul>
              <div class="chart-box compact-chart-box" style="margin-top:14px;">
                <h4>Budgetkorridor</h4>
                <div class="chart-stat-grid budget-corridor-grid">
                  <div><span>EW</span><strong>${formatCurrency(result.expectedValue)}</strong></div>
                  <div><span>84 %</span><strong>${formatCurrency(result.budget84)}</strong></div>
                  <div><span>97,5 %</span><strong>${formatCurrency(result.budget975)}</strong></div>
                </div>
              </div>
            </section>
          </div>
          <div class="card-grid pert-progress-grid">
            <section class="info-card lph-card">
              <h3>LPH-Fortschreibung</h3>
              <p class="lph-intro">Je Leistungsphase wird der zuletzt gespeicherte Snapshot gelesen. Für die aktive LPH wird der laufende Schätzstand direkt verwendet.</p>
              <div class="lph-columns">
                <div class="table-wrap pert-table-wrap">
                  <table class="data-table pert-metrics-table pert-lph-table">
                    <thead>
                      <tr>
                        <th>LPH</th>
                        <th>Snapshot</th>
                        <th>EW</th>
                        <th>Trend</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${(analytics.trendRows.length ? analytics.trendRows : result.lphRows).slice(0, 4).map((row, idx, arr) => `
                        <tr>
                          <td><strong>${row.label}</strong></td>
                          <td>
                            <div class="phase-source-cell">
                              <span class="phase-source-badge ${row.source === "current" ? "live" : row.source === "snapshot" ? "snapshot" : "empty"}">
                                ${row.source === "current" ? "Live-Stand" : row.source === "snapshot" ? "Snapshot" : "Kein Snapshot"}
                              </span>
                              ${row.source === "snapshot" ? `<small>${row.snapshotLabel}</small>` : ""}
                            </div>
                          </td>
                          <td>
                            <div class="lph-metric-cell">
                              <strong>${row.expectedValue > 0 ? formatCurrency(row.expectedValue) : "—"}</strong>
                            </div>
                          </td>
                          <td>
                            <div class="lph-metric-cell">
                              <strong>${row.expectedValue > 0 ? (row.delta === 0 ? "Startwert" : `${row.delta > 0 ? "+" : ""}${formatCurrency(row.delta)}`) : "—"}</strong>
                            </div>
                          </td>
                        </tr>
                        <tr class="lph-comment-row">
                          <td colspan="4">
                            <div class="lph-comment-box">
                              <label>Kommentar</label>
                              <textarea data-pert-lph-field="comment" data-pert-lph-label="${row.label}">${row.comment}</textarea>
                            </div>
                          </td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
                <div class="table-wrap pert-table-wrap">
                  <table class="data-table pert-metrics-table pert-lph-table">
                    <thead>
                      <tr>
                        <th>LPH</th>
                        <th>Snapshot</th>
                        <th>EW</th>
                        <th>Trend</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${(analytics.trendRows.length ? analytics.trendRows : result.lphRows).slice(4, 8).map((row, idx, arr) => `
                        <tr>
                          <td><strong>${row.label}</strong></td>
                          <td>
                            <div class="phase-source-cell">
                              <span class="phase-source-badge ${row.source === "current" ? "live" : row.source === "snapshot" ? "snapshot" : "empty"}">
                                ${row.source === "current" ? "Live-Stand" : row.source === "snapshot" ? "Snapshot" : "Kein Snapshot"}
                              </span>
                              ${row.source === "snapshot" ? `<small>${row.snapshotLabel}</small>` : ""}
                            </div>
                          </td>
                          <td>
                            <div class="lph-metric-cell">
                              <strong>${row.expectedValue > 0 ? formatCurrency(row.expectedValue) : "—"}</strong>
                            </div>
                          </td>
                          <td>
                            <div class="lph-metric-cell">
                              <strong>${row.expectedValue > 0 ? (row.delta === 0 ? "Startwert" : `${row.delta > 0 ? "+" : ""}${formatCurrency(row.delta)}`) : "—"}</strong>
                            </div>
                          </td>
                        </tr>
                        <tr class="lph-comment-row">
                          <td colspan="4">
                            <div class="lph-comment-box">
                              <label>Kommentar</label>
                              <textarea data-pert-lph-field="comment" data-pert-lph-label="${row.label}">${row.comment}</textarea>
                            </div>
                          </td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </div>
          <section class="info-card snapshots-card">
              <h3>Snapshots</h3>
              ${state.pert.snapshots.length ? `
                <div class="analysis-grid">
                  ${state.pert.snapshots.map((snapshot, index) => `
                    <div class="mini-card">
                      <h4>${normalizePertSnapshotLabel(snapshot.label)}</h4>
                      <div class="snapshot-meta-line">
                        <span>${snapshot.phase || "LPH"}</span>
                        <span>${normalizePertSecurityLabel(snapshot.securityLabel) || securityLabels[Number(snapshot.securityLevel)] || "84 % · Vorsichtiges Niveau"}</span>
                        <span>${snapshot.savedAt}</span>
                      </div>
                      <div class="snapshot-metric-grid">
                        <div class="snapshot-metric">
                          <span>Budgetpfad</span>
                          <strong>${formatCurrency(snapshot.selectedBudget)}</strong>
                        </div>
                        <div class="snapshot-metric">
                          <span>Erwartungswert</span>
                          <strong>${formatCurrency(snapshot.expectedValue)}</strong>
                        </div>
                      </div>
                      <div class="snapshot-card-actions">
                        <button class="action-btn" type="button" data-action="restore-pert-snapshot" data-index="${index}">Wiederherstellen</button>
                        <button class="action-btn" type="button" data-action="delete-pert-snapshot" data-index="${index}">Löschen</button>
                      </div>
                    </div>
                  `).join("")}
                </div>
              ` : `
                <p class="form-note">Noch keine Snapshots vorhanden. Nutze sie, um Budgetpfade und Schätzstände pro Fortschreibung zu sichern.</p>
              `}
          </section>
          <section class="info-card cost-items-card" id="pertCostItemsSection">
            <h3>Gewerke / Kostenpositionen</h3>
            <div class="table-wrap pert-table-wrap">
              <table class="data-table pert-edit-table">
                <thead>
                  <tr>
                    <th>Mark.</th>
                    <th>Gewerk</th>
                    <th>Optimistisch</th>
                    <th>Wahrscheinlich</th>
                    <th>Pessimistisch</th>
                    <th>Kommentar</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  ${analytics.detailRows.map((item) => `
                    <tr>
                      <td><input data-transfer-select="pertItemIds" data-transfer-value="${item.id}" type="checkbox" ${selectedPertIds.includes(item.id) ? "checked" : ""}></td>
                      <td><input class="compact-input name pert-name-input" id="pert_name_${item.originalIndex}" data-pert-item-index="${item.originalIndex}" data-pert-item-field="name" type="text" value="${item.name}"></td>
                      <td><input class="compact-input currency-input" id="pert_o_${item.originalIndex}" data-pert-item-index="${item.originalIndex}" data-pert-item-field="optimistic" type="text" inputmode="numeric" value="${formatCurrencyInput(item.optimistic)}"></td>
                      <td><input class="compact-input currency-input" id="pert_m_${item.originalIndex}" data-pert-item-index="${item.originalIndex}" data-pert-item-field="mostLikely" type="text" inputmode="numeric" value="${formatCurrencyInput(item.mostLikely)}"></td>
                      <td><input class="compact-input currency-input" id="pert_p_${item.originalIndex}" data-pert-item-index="${item.originalIndex}" data-pert-item-field="pessimistic" type="text" inputmode="numeric" value="${formatCurrencyInput(item.pessimistic)}"></td>
                      <td><input class="compact-input name pert-comment-input" id="pert_comment_${item.originalIndex}" data-pert-item-index="${item.originalIndex}" data-pert-item-field="comment" type="text" value="${item.comment || ""}"></td>
                      <td><button class="danger-icon-button" type="button" aria-label="Kostenposition entfernen" title="Kostenposition entfernen" data-action="remove-pert-item" data-index="${item.originalIndex}">×</button></td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </section>
          <section class="info-card">
            <h3>Tabellarische Detailauswertung nach Gewerken</h3>
            <div class="table-wrap pert-table-wrap">
              <table class="data-table pert-metrics-table pert-detail-metrics-table">
                <thead>
                  <tr>
                    <th>Gewerk</th>
                    <th>Erwartungswert</th>
                    <th>Spannweite</th>
                    <th>Anteil</th>
                    <th>SD</th>
                  </tr>
                </thead>
                <tbody>
                  ${analytics.detailRows.map((item) => `
                    <tr>
                      <td><strong>${item.name}</strong></td>
                      <td class="numeric-output">${formatCurrency(item.expectedValue)}</td>
                      <td class="numeric-output">${formatCurrency(item.span)}</td>
                      <td class="numeric-output">${formatPercent(item.share)}</td>
                      <td class="numeric-output">${formatCurrency(item.standardDeviation)}</td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </section>
          <section class="chart-box pert-phase-chart-card">
            <h3>Kostenverlauf nach Leistungsphasen – Präzisierung im Projektverlauf</h3>
            <p class="chart-intro">Gestrichelte Linien zeigen die theoretische Unsicherheitsbandbreite von LPH 1 bis LPH 8. Die durchgezogene Linie zeigt den aktuellen Projektverlauf. Aktuell: ${state.project?.phase || "LPH"}.</p>
            <div class="chart-legend-row phase-chart-legend">
              <span class="legend-pill"><span class="legend-line legend-line-solid"></span>Tatsächlicher Projektverlauf (EW)</span>
              <span class="legend-pill"><span class="legend-line legend-line-upper"></span>Obere Bandbreite – Pessimistisch (+SD)</span>
              <span class="legend-pill"><span class="legend-line legend-line-lower"></span>Untere Bandbreite – Optimistisch (−SD)</span>
            </div>
            ${analytics.phaseCorridorSvg}
          </section>
          <section class="chart-box pert-snapshot-chart-card">
            <div class="snapshot-chart-head">
              <div>
                <h3>Kostenverlauf & Trendanalyse – Schnappschüsse über Planungsphasen</h3>
                <p class="chart-intro">Schnappschüsse sichern den aktuellen Schätzstand. Die Trendlinie zeigt, wie sich die Kostenprognose über die Planungsphasen entwickelt.</p>
              </div>
              <button class="action-btn" type="button" data-action="save-pert-snapshot">+ Schnappschuss erfassen</button>
            </div>
            ${analytics.snapshotRows.length ? `
              <div class="snapshot-trend-list">
                ${analytics.snapshotRows.map((snapshot, index) => `
                  <div class="snapshot-trend-row">
                    <div class="snapshot-trend-label">${snapshot.phase} · ${formatDate(snapshot.savedAt)}</div>
                    <div class="snapshot-trend-values">
                      <strong>${(snapshot.expectedValue / 1000000).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Mio. €</strong>
                      <span class="snapshot-level snapshot-level-84">84 %: ${(snapshot.budget84 / 1000000).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Mio. €</span>
                      <span class="snapshot-level snapshot-level-975">97,5 %: ${(snapshot.budget975 / 1000000).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Mio. €</span>
                      <button class="snapshot-delete-inline" type="button" aria-label="Snapshot löschen" data-action="delete-pert-snapshot" data-index="${state.pert.snapshots.findIndex((entry) => entry.savedAt === snapshot.savedAt && entry.phase === snapshot.phase)}">×</button>
                    </div>
                  </div>
                `).join("")}
              </div>
              <div class="chart-legend-row snapshot-chart-legend">
                <span class="legend-pill"><span class="legend-dot" style="background:#1c2f4d"></span>Erwartungswert (EW)</span>
                <span class="legend-pill"><span class="legend-dot" style="background:#4f956c"></span>84%-Niveau</span>
                <span class="legend-pill"><span class="legend-dot" style="background:#c6a03f"></span>97,5%-Niveau</span>
              </div>
              ${analytics.trendChartSvg}
            ` : `
              <p class="form-note">Noch keine Schnappschüsse vorhanden. Speichere einen Snapshot, um hier den Verlauf über die Planungsphasen zu sehen.</p>
            `}
          </section>
          <div class="pert-feature-stack">
            <section class="info-card pert-feature-card pert-feature-cost">
              <div class="pert-feature-head">
                <div>
                  <h3>Kostenvergleich nach Gewerken</h3>
                  <p>Vergleich der erwarteten Kostenbeiträge je Gewerk.</p>
                </div>
                <span class="badge">Kostenstruktur</span>
              </div>
              <div class="pert-feature-layout">
                <div class="chart-box compact-chart-box pert-feature-chart">
                  ${analytics.costComparisonSvg}
                </div>
                <div class="table-wrap pert-table-wrap pert-side-table">
                  <table class="data-table">
                    <thead>
                      <tr>
                        <th>Gewerk</th>
                        <th>EW</th>
                        <th>Anteil</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${analytics.detailRows.map((item) => `
                        <tr>
                          <td><strong>${item.name}</strong></td>
                          <td class="numeric-output">${formatCurrency(item.expectedValue)}</td>
                          <td class="numeric-output">${formatPercent(item.share)}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <section class="info-card pert-feature-card pert-feature-sensitivity">
              <div class="pert-feature-head">
                <div>
                  <h3>Sensitivitätsanalyse</h3>
                  <p>Zeigt, welche Positionen den größten Beitrag zur Gesamtunsicherheit leisten.</p>
                </div>
                <span class="badge">Unsicherheit</span>
              </div>
              <div class="pert-feature-layout">
                <div class="chart-box compact-chart-box pert-feature-chart">
                  ${analytics.sensitivityChartSvg}
                </div>
                <div class="table-wrap pert-table-wrap pert-side-table">
                  <table class="data-table">
                    <thead>
                      <tr>
                        <th>Gewerk</th>
                        <th>Varianzanteil</th>
                        <th>Spannweite</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${analytics.topSensitivity.map((item) => `
                        <tr>
                          <td><strong>${item.name}</strong></td>
                          <td class="numeric-output">${(item.sensitivityScore * 100).toFixed(1)} %</td>
                          <td class="numeric-output">${formatCurrency(item.span)}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <section class="info-card pert-feature-card pert-feature-span">
              <div class="pert-feature-head">
                <div>
                  <h3>Spannweite je Kostenposition</h3>
                  <p>Differenz zwischen optimistischer und pessimistischer Annahme je Gewerk.</p>
                </div>
                <span class="badge">Bandbreite</span>
              </div>
              <div class="pert-feature-layout">
                <div class="chart-box compact-chart-box pert-feature-chart">
                  ${analytics.spanChartSvg}
                </div>
                <div class="table-wrap pert-table-wrap pert-side-table">
                  <table class="data-table">
                    <thead>
                      <tr>
                        <th>Gewerk</th>
                        <th>Optimistisch</th>
                        <th>Pessimistisch</th>
                        <th>Spannweite</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${analytics.detailRows.map((item) => `
                        <tr>
                          <td><strong>${item.name}</strong></td>
                          <td class="numeric-output">${formatCurrency(item.optimistic)}</td>
                          <td class="numeric-output">${formatCurrency(item.pessimistic)}</td>
                          <td class="numeric-output">${formatCurrency(item.span)}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <section class="info-card pert-feature-card pert-feature-compare">
              <div class="pert-feature-head">
                <div>
                  <h3>Szenariovergleich</h3>
                  <p>Vergleich des aktuellen Schätzstands mit dem letzten gespeicherten Referenzpunkt.</p>
                </div>
                <span class="badge">Vergleich</span>
              </div>
              ${analytics.scenarioCompare ? `
                <div class="analysis-grid">
                  <div class="mini-card">
                    <h4>Referenz</h4>
                    <p><strong>${analytics.scenarioCompare.label}</strong><br>${analytics.scenarioCompare.savedAt}</p>
                  </div>
                  <div class="mini-card">
                    <h4>Erwartungswert</h4>
                    <p>${analytics.scenarioCompare.expectedValueDelta > 0 ? "+" : ""}${formatCurrency(analytics.scenarioCompare.expectedValueDelta)}</p>
                  </div>
                  <div class="mini-card">
                    <h4>Budgetpfad</h4>
                    <p>${analytics.scenarioCompare.selectedBudgetDelta > 0 ? "+" : ""}${formatCurrency(analytics.scenarioCompare.selectedBudgetDelta)}</p>
                  </div>
                  <div class="mini-card">
                    <h4>Standardabweichung</h4>
                    <p>${analytics.scenarioCompare.standardDeviationDelta > 0 ? "+" : ""}${formatCurrency(analytics.scenarioCompare.standardDeviationDelta)}</p>
                  </div>
                </div>
              ` : `<p class="form-note">Lege zuerst einen Snapshot an, um den aktuellen Stand gegen einen gespeicherten Schätzstand zu vergleichen.</p>`}
            </section>
          </div>
          <section class="chart-box pert-distribution-card">
            <h3>Normalverteilung der Schätzung</h3>
            <div class="distribution-top-row">
              ${analytics.normalDistributionMarkers.map((marker, index) => {
                const keys = ["p25X", "p16X", "meanX", "p84X", "p975X"];
                const positionKey = keys[index];
                const leftPercent = ((analytics.normalDistributionPositions[positionKey] / analytics.normalDistributionPositions.width) * 100).toFixed(2);
                return `
                  <div class="distribution-marker anchored" style="left:${leftPercent}%;">
                    <strong style="color:${marker.color}">${formatCurrency(marker.value)}</strong>
                    <span>${marker.label}</span>
                    <small>${marker.sublabel || (marker.label === "EW" ? "Erwartungswert" : "")}</small>
                  </div>
                `;
              }).join("")}
            </div>
            ${analytics.normalDistributionSvg}
          </section>
        </div>
      `;
    }
  },
  monteCarlo: {
    key: "monteCarlo",
    label: "Monte-Carlo-Simulation",
    subtitle: "Simulation, Szenarien und Maßnahmen",
    render(state) {
      const result = calculateMonteCarloResult(state);
      const measureResult = calculateMonteCarloResult(state, { applyMeasures: true });
      const analytics = calculateMonteCarloAnalytics(result);
      const riskStatus = getRiskLevel(result.combinedSuccessRate);
      const budgetGapToProject = state.project.budget > 0 ? state.project.budget - result.p80Cost : null;
      const targetGapToProject = state.project.budget > 0 ? state.project.budget - state.monteCarlo.targetBudget : null;
      const baselineSummary = state.monteCarlo.baselineScenario?.summary || null;
      const baselineCompare = baselineSummary ? {
        costDelta: result.meanCost - baselineSummary.meanCost,
        daysDelta: result.meanDays - baselineSummary.meanDays,
        successDelta: result.combinedSuccessRate - baselineSummary.combinedSuccessRate,
        p80Delta: result.p80Cost - baselineSummary.p80Cost
      } : null;
      const measureCompare = {
        costDelta: measureResult.meanCost - result.meanCost,
        daysDelta: measureResult.meanDays - result.meanDays,
        successDelta: measureResult.combinedSuccessRate - result.combinedSuccessRate,
        p80Delta: measureResult.p80Cost - result.p80Cost
      };
      const workWithoutDuration = result.workPackages.filter((item) => item.optimisticDays === 0 && item.mostLikelyDays === 0 && item.pessimisticDays === 0).length;
      const workWithoutCost = result.workPackages.filter((item) => item.optimisticCost === 0 && item.mostLikelyCost === 0 && item.pessimisticCost === 0).length;
      const risksWithoutImpact = result.riskEvents.filter((item) => item.maxCostImpact === 0 && item.maxDelayDays === 0).length;
      const plausibilityChecks = [
        {
          label: "Arbeitspakete ohne Terminwerte",
          value: workWithoutDuration,
          tone: workWithoutDuration === 0 ? "ok" : "warn",
          detail: workWithoutDuration === 0 ? "alle Pakete haben Terminannahmen" : "diese Pakete schwächen die Terminsimulation"
        },
        {
          label: "Arbeitspakete ohne Kostenwerte",
          value: workWithoutCost,
          tone: workWithoutCost === 0 ? "ok" : "warn",
          detail: workWithoutCost === 0 ? "alle Pakete haben Kostenannahmen" : "diese Pakete schwächen die Kostensimulation"
        },
        {
          label: "Risiken ohne Wirkung",
          value: risksWithoutImpact,
          tone: risksWithoutImpact === 0 ? "ok" : "warn",
          detail: risksWithoutImpact === 0 ? "alle Risiken haben einen Impact" : "ohne Wirkung bleiben sie im Modell stumm"
        },
        {
          label: "Simulationsgüte",
          value: result.iterations,
          tone: result.iterations >= 10000 ? "ok" : "warn",
          detail: result.iterations >= 10000 ? "Iterationenzahl ist belastbar" : "mehr Läufe würden die Verteilung stabilisieren"
        }
      ];
      const summaryLines = [
        `Die aktuelle Zielerreichung liegt bei ${formatPercent(result.combinedSuccessRate)} für Budget und Termin gemeinsam.`,
        `Das P80-Niveau liegt bei ${formatCurrency(result.p80Cost)} und ${formatNumber(result.p80Days)} Tagen.`
      ];
      if (budgetGapToProject === null) {
        summaryLines.push("Für das globale Projektbudget ist noch kein belastbarer Abgleich möglich.");
      } else if (budgetGapToProject < 0) {
        summaryLines.push("Das P80-Niveau liegt aktuell über dem globalen Projektbudget.");
      }
      if (workWithoutDuration > 0) {
        summaryLines.push(`${workWithoutDuration} Arbeitspakete haben noch keine Terminwerte und schwächen die Terminseite des Modells.`);
      }
      if (workWithoutCost > 0) {
        summaryLines.push(`${workWithoutCost} Arbeitspakete haben noch keine Kostenwerte und schwächen die Kostenseite des Modells.`);
      }
      if (risksWithoutImpact > 0) {
        summaryLines.push(`${risksWithoutImpact} Risiken haben aktuell keine Wirkung und tragen noch nicht zur Verteilung bei.`);
      }
      return `
        <div class="module-shell monte-module">
          <div class="module-title">
            <div>
              <h2>Monte-Carlo-Simulation</h2>
              <p>Das Modul simuliert Kosten- und Terminunsicherheit als Verteilung. Arbeitspakete und Risiken verdichten sich zu einer gemeinsamen Sicht auf Budget- und Terminziel.</p>
            </div>
            <span class="badge">Simulation</span>
          </div>
          <div class="kpi-grid">
            <article class="kpi-card"><div class="kpi-label">Erwartete Kosten</div><div class="kpi-value">${formatCurrency(result.meanCost)}</div><div class="kpi-sub">P50: ${formatCurrency(result.p50Cost)} · P80: ${formatCurrency(result.p80Cost)}</div></article>
            <article class="kpi-card gold"><div class="kpi-label">Erwartete Dauer</div><div class="kpi-value">${Math.round(result.meanDays)} T</div><div class="kpi-sub">P50: ${Math.round(result.p50Days)} · P80: ${Math.round(result.p80Days)} Tage</div></article>
            <article class="kpi-card green"><div class="kpi-label">Budgettreffer</div><div class="kpi-value">${formatPercent(result.budgetSuccessRate)}</div><div class="kpi-sub">gegen Simulationsbudget ${formatCurrency(state.monteCarlo.targetBudget)}</div></article>
            <article class="kpi-card blue"><div class="kpi-label">Zielerreichung gesamt</div><div class="kpi-value">${formatPercent(result.combinedSuccessRate)}</div><div class="kpi-sub">Budget und Termin gemeinsam · ${riskStatus}</div></article>
          </div>
          <div class="card-grid monte-upper-grid">
            <section class="info-card control-card">
              <h3>Steuerung</h3>
              <div class="monte-control-grid">
                <div class="monte-control-field">
                  <label for="mc_iterations">Iterationen</label>
                  <input class="compact-input integer-input" id="mc_iterations" data-mc-field="iterations" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.iterations)}">
                </div>
                <div class="monte-control-field">
                  <label for="mc_seed">Seed</label>
                  <input class="compact-input integer-input" id="mc_seed" data-mc-field="seed" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.seed)}">
                </div>
                <div class="monte-control-field">
                  <label for="mc_target_budget">Zielbudget</label>
                  <input class="compact-input currency-input" id="mc_target_budget" data-mc-field="targetBudget" type="text" inputmode="numeric" value="${formatCurrencyInput(state.monteCarlo.targetBudget)}">
                </div>
                <div class="monte-control-field">
                  <label for="mc_target_days">Zieldauer</label>
                  <input class="compact-input integer-input" id="mc_target_days" data-mc-field="targetDays" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.targetDays)}">
                </div>
              </div>
              <div class="monte-control-metrics">
                <div class="monte-control-stat">
                  <span>Median</span>
                  <strong>${formatCurrency(result.medianCost)}</strong>
                  <small>Kostenverteilung P50</small>
                </div>
                <div class="monte-control-stat">
                  <span>Dauer Median</span>
                  <strong>${formatNumber(result.medianDays)} Tage</strong>
                  <small>Terminverteilung P50</small>
                </div>
                <div class="monte-control-stat">
                  <span>Globales Projektbudget</span>
                  <strong>${state.project.budget ? formatCurrency(state.project.budget) : "—"}</strong>
                  <small>globaler Projektrahmen</small>
                </div>
                <div class="monte-control-stat">
                  <span>Abgleich Simulationsbudget</span>
                  <strong>${targetGapToProject === null ? "—" : `${targetGapToProject >= 0 ? "+" : ""}${formatCurrency(targetGapToProject)}`}</strong>
                  <small>globaler Rahmen gegen Simulationsbudget</small>
                </div>
              </div>
            </section>
            <section class="info-card steering-card">
              <h3>Steuerungsbild</h3>
              <div class="monte-steering-grid">
                <div class="monte-steering-item">
                  <span>Berechnete Läufe</span>
                  <strong>${formatNumber(result.iterations)}</strong>
                </div>
                <div class="monte-steering-item">
                  <span>Median Kosten / Dauer</span>
                  <strong>${formatCurrency(result.medianCost)} · ${formatNumber(result.medianDays)} Tage</strong>
                </div>
                <div class="monte-steering-item">
                  <span>Reserve bis P80</span>
                  <strong>${formatCurrency(result.reserveBudgetGap)} · ${formatNumber(result.reserveDayGap)} Tage</strong>
                </div>
                <div class="monte-steering-item">
                  <span>Projektbudget vs. P80</span>
                  <strong>${budgetGapToProject === null ? "Kein Projektbudget" : `${budgetGapToProject >= 0 ? "+" : ""}${formatCurrency(budgetGapToProject)}`}</strong>
                </div>
              </div>
              <p class="form-note monte-budget-note" style="margin-top:12px;"><span>Projektbudget = globaler Rahmen</span><span>Zielbudget = Simulationswert</span></p>
            </section>
          </div>
          <div class="analysis-grid">
            <section class="mini-card" id="monte-p80-costs">
              <h4>Verteilung und Reserven</h4>
              <ul>
                <li><strong>Kosten P10 / P50 / P80 / P90:</strong> ${formatCurrency(result.p10Cost)} · ${formatCurrency(result.p50Cost)} · ${formatCurrency(result.p80Cost)} · ${formatCurrency(result.p90Cost)}</li>
                <li><strong>Dauer P10 / P50 / P80 / P90:</strong> ${formatNumber(result.p10Days)} · ${formatNumber(result.p50Days)} · ${formatNumber(result.p80Days)} · ${formatNumber(result.p90Days)} Tage</li>
                <li><strong>Ø Überschreitung:</strong> ${formatCurrency(result.avgCostOverrun)} und ${formatNumber(result.avgDayOverrun)} Tage</li>
              </ul>
            </section>
            <section class="mini-card">
              <h4>Risikotreiber</h4>
              <ul>
                ${(result.driverStats.length
                  ? result.driverStats.slice(0, 4).map((item) => `
                    <li><strong>${item.name}:</strong> ${formatCurrency(item.avgCostImpact)} · ${formatNumber(item.avgDelayImpact)} Tage · Treffer ${formatPercent(item.hitRate)}</li>
                  `).join("")
                  : "<li>Keine Risikoereignisse im Modell.</li>")}
              </ul>
            </section>
          </div>
          <div class="card-grid monte-upper-grid">
            <section class="info-card">
              <h3>Welche Werte kannst du eingeben?</h3>
              <div class="monte-explain-grid">
                <div class="monte-explain-item">
                  <strong>Optimistisch</strong>
                  <p>günstiger Wert für Kosten oder Dauer, wenn der Ablauf sehr gut gelingt.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Wahrscheinlich</strong>
                  <p>realistischster Wert, der unter normalen Projektbedingungen erwartet wird.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Pessimistisch</strong>
                  <p>ungünstiger Wert, wenn Störungen, Reibung oder Verzögerungen eintreten.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Kostenentwicklung</strong>
                  <p>bildet die Spannweite zwischen optimistisch und pessimistisch über das jeweilige Arbeitspaket oder Risiko ab.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Eintrittswahrscheinlichkeit</strong>
                  <p>Eintrittswahrscheinlichkeit eines Risikoereignisses in Prozent.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Terminwirkung</strong>
                  <p>zusätzliche Verzögerung in Tagen, die ein Risiko bei Eintritt auslösen kann.</p>
                </div>
              </div>
              <div class="form-note monte-fixed-note">
                <div class="monte-fixed-note-line">Alle Eingaben wirken direkt auf die Simulation</div>
                <div class="monte-fixed-note-line">und damit auf Kostenverteilung, Dauerverteilung und Zielerreichung.</div>
              </div>
            </section>
            <section class="info-card">
              <h3>Leseschlüssel</h3>
              <div class="monte-explain-grid">
                <div class="monte-explain-item">
                  <strong>P10</strong>
                  <p>früher günstiger Bereich. Nur 10 % der Läufe liegen darunter.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>P50</strong>
                  <p>Median der Simulation. 50 % der Läufe liegen darunter.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>P80</strong>
                  <p>vorsichtiger Steuerungswert. 80 % der Läufe liegen darunter.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>P90</strong>
                  <p>sehr defensiver Wert für hohe Vorsicht und Reserven.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Budgettreffer</strong>
                  <p>Anteil der Läufe, die das Zielbudget einhalten.</p>
                </div>
                <div class="monte-explain-item">
                  <strong>Termintreffer</strong>
                  <p>Anteil der Läufe, die die Zieldauer einhalten.</p>
                </div>
              </div>
            </section>
          </div>
          <section class="info-card monte-coverage-card">
            <h3>Zielabdeckung für Budget und Termin</h3>
            <p class="chart-intro">Wie gut decken die aktuellen Zielwerte die simulierten Ergebnisse ab? Dargestellt wird die Trefferquote für Budget und Termin auf Basis der aktuellen Monte-Carlo-Läufe.</p>
            <div class="monte-coverage-grid">
              <div class="monte-coverage-item">
                <div class="monte-coverage-head">
                  <span>Simulationsbudget</span>
                  <strong>${formatPercent(result.budgetSuccessRate)}</strong>
                </div>
                <div class="monte-coverage-bar">
                  <span style="width:${Math.max(0, Math.min(100, result.budgetSuccessRate * 100))}%;"></span>
                </div>
                <small>aktuelles Simulationsbudget ${formatCurrency(state.monteCarlo.targetBudget)}</small>
              </div>
              <div class="monte-coverage-item">
                <div class="monte-coverage-head">
                  <span>Terminziel</span>
                  <strong>${formatPercent(result.scheduleSuccessRate)}</strong>
                </div>
                <div class="monte-coverage-bar">
                  <span style="width:${Math.max(0, Math.min(100, result.scheduleSuccessRate * 100))}%;"></span>
                </div>
                <small>Zieldauer ${formatNumber(state.monteCarlo.targetDays)} Tage</small>
              </div>
            </div>
          </section>
          <div class="card-grid monte-upper-grid monte-secondary-grid">
            <section class="info-card">
              <h3>Plausibilitätsprüfung</h3>
              <div class="monte-explain-grid">
                ${plausibilityChecks.map((item) => `
                  <div class="monte-explain-item monte-check-item ${item.tone === "ok" ? "is-ok" : "is-warn"}">
                    <strong>${item.label}</strong>
                    <p><span class="monte-check-value">${formatNumber(item.value)}</span> · ${item.detail}</p>
                  </div>
                `).join("")}
              </div>
            </section>
            <section class="info-card">
              <h3>Zusammenfassung</h3>
              <ul class="monte-summary-list">
                ${summaryLines.map((line) => `<li>${line}</li>`).join("")}
              </ul>
            </section>
          </div>
          <div class="card-grid monte-upper-grid">
            <section class="info-card">
              <h3>Maßnahmenlogik</h3>
              <p class="chart-intro">Mit diesen Stellhebeln kannst du Zielwerte bewusst lockern oder Risiken methodisch dämpfen. Die Auswirkungen werden direkt auf P80, Durchschnittsdauer und Zielerreichung sichtbar.</p>
              <div class="form-grid">
                <div class="form-field">
                  <label for="mc_measure_budget">Budgetziel lockern in&nbsp;%</label>
                  <input id="mc_measure_budget" class="compact-input integer-input" data-mc-measure-field="budgetPercent" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.measureSettings?.budgetPercent ?? 0)}">
                </div>
                <div class="form-field">
                  <label for="mc_measure_days">Terminziel lockern in&nbsp;%</label>
                  <input id="mc_measure_days" class="compact-input integer-input" data-mc-measure-field="daysPercent" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.measureSettings?.daysPercent ?? 0)}">
                </div>
                <div class="form-field">
                  <label for="mc_measure_risk">Eintritte senken in&nbsp;%</label>
                  <input id="mc_measure_risk" class="compact-input integer-input" data-mc-measure-field="riskReductionPercent" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.measureSettings?.riskReductionPercent ?? 0)}">
                </div>
                <div class="form-field">
                  <label for="mc_measure_impact">Risikoauswirkungen senken in&nbsp;%</label>
                  <input id="mc_measure_impact" class="compact-input integer-input" data-mc-measure-field="impactReductionPercent" type="text" inputmode="numeric" value="${formatIntegerInput(state.monteCarlo.measureSettings?.impactReductionPercent ?? 0)}">
                </div>
              </div>
              <div class="analysis-grid monte-measure-impact-grid" style="margin-top:14px;">
                <div class="mini-card">
                  <h4>Wirkung auf<br>P80 Kosten</h4>
                  <p>${measureCompare.p80Delta > 0 ? "+" : ""}${formatCurrency(measureCompare.p80Delta)}</p>
                </div>
                <div class="mini-card">
                  <h4>Wirkung auf<br>Durchschnittsdauer</h4>
                  <p>${measureCompare.daysDelta > 0 ? "+" : ""}${formatNumber(measureCompare.daysDelta)} Tage</p>
                </div>
                <div class="mini-card">
                  <h4>Wirkung auf<br>Zielerreichung</h4>
                  <p>${measureCompare.successDelta > 0 ? "+" : ""}${formatPercent(measureCompare.successDelta)}</p>
                </div>
              </div>
            </section>
            <section class="info-card">
              <h3>Baseline und Szenarien</h3>
              <p class="chart-intro">Sichere einen belastbaren Referenzstand als Baseline und halte Varianten als Szenarien fest. So lassen sich Modellstände später gezielt vergleichen und wiederherstellen.</p>
              <div class="control-actions" style="margin-bottom:12px;">
                <button class="action-btn" type="button" data-action="save-mc-baseline">Baseline speichern</button>
                <button class="action-btn" type="button" data-action="save-mc-scenario">Szenario sichern</button>
              </div>
              <div class="monte-baseline-panel ${state.monteCarlo.baselineScenario ? "is-available" : "is-empty"}">
                <div class="monte-baseline-head">
                  <span class="badge">${state.monteCarlo.baselineScenario ? "Baseline aktiv" : "Keine Baseline"}</span>
                  <strong>${state.monteCarlo.baselineScenario ? state.monteCarlo.baselineScenario.label : "Noch keine Baseline gespeichert"}</strong>
                  <small>${state.monteCarlo.baselineScenario ? state.monteCarlo.baselineScenario.savedAt : "Sichere den aktuellen Modellstand als belastbaren Referenzpunkt."}</small>
                </div>
                <div class="monte-baseline-metrics">
                  <div><span>Gesicherte Szenarien</span><strong>${formatNumber((state.monteCarlo.storedScenarios || []).length)}</strong></div>
                  <div><span>Maßnahmenrahmen</span><strong>${formatNumber(state.monteCarlo.measureSettings?.budgetPercent || 0)} % Budget · ${formatNumber(state.monteCarlo.measureSettings?.daysPercent || 0)} % Termin</strong></div>
                </div>
                ${state.monteCarlo.baselineScenario ? `
                  <div class="control-actions" style="margin-top:12px;">
                    <button class="action-btn" type="button" data-action="restore-mc-baseline">Baseline wiederherstellen</button>
                  </div>
                ` : ""}
              </div>
            </section>
          </div>
          ${baselineCompare ? `
            <div class="analysis-grid monte-baseline-compare-grid">
              <div class="mini-card baseline-highlight">
                <h4>Abweichung zur Baseline</h4>
                <div class="monte-scenario-metrics">
                  <div><span>Kosten</span><strong>${baselineCompare.costDelta > 0 ? "+" : ""}${formatCurrency(baselineCompare.costDelta)}</strong></div>
                  <div><span>Dauer</span><strong>${baselineCompare.daysDelta > 0 ? "+" : ""}${formatNumber(baselineCompare.daysDelta)} Tage</strong></div>
                </div>
              </div>
              <div class="mini-card baseline-highlight">
                <h4>Risikoniveau zur Baseline</h4>
                <div class="monte-scenario-metrics">
                  <div><span>P80</span><strong>${baselineCompare.p80Delta > 0 ? "+" : ""}${formatCurrency(baselineCompare.p80Delta)}</strong></div>
                  <div><span>Zielerreichung</span><strong>${baselineCompare.successDelta > 0 ? "+" : ""}${formatPercent(baselineCompare.successDelta)}</strong></div>
                </div>
              </div>
            </div>
          ` : ""}
          ${(state.monteCarlo.storedScenarios || []).length ? `
            <section class="info-card">
              <h3>Gesicherte Szenarien</h3>
              <p class="chart-intro">Die Szenarien speichern Zielwerte, Iterationen und die zentralen Simulationsergebnisse des jeweiligen Modellstands.</p>
              <div class="analysis-grid monte-scenario-grid">
                ${state.monteCarlo.storedScenarios.map((scenario, index) => `
                  <div class="mini-card monte-scenario-card">
                    <div class="monte-scenario-card-head">
                      <span class="badge">Szenario</span>
                      <h4>${scenario.label}</h4>
                      <p>${scenario.savedAt}</p>
                    </div>
                    <div class="monte-scenario-metrics">
                      <div><span>Zielbudget</span><strong>${formatCurrency(scenario.targetBudget)}</strong></div>
                      <div><span>Zieldauer</span><strong>${formatNumber(scenario.targetDays)} Tage</strong></div>
                      <div><span>Iterationen</span><strong>${formatNumber(scenario.iterations)}</strong></div>
                      <div><span>P80 Kosten</span><strong>${scenario.summary ? formatCurrency(scenario.summary.p80Cost) : "—"}</strong></div>
                      <div><span>Zielerreichung</span><strong>${scenario.summary ? formatPercent(scenario.summary.combinedSuccessRate) : "—"}</strong></div>
                    </div>
                    <div style="margin-top:12px;display:flex;justify-content:flex-end;gap:10px;flex-wrap:wrap;">
                      <button class="action-btn" type="button" data-action="restore-mc-scenario" data-index="${index}">Wiederherstellen</button>
                      <button class="action-btn" type="button" data-action="delete-mc-scenario" data-index="${index}">Löschen</button>
                    </div>
                  </div>
                `).join("")}
              </div>
            </section>
          ` : ""}
          <div class="card-grid monte-distribution-grid">
            <section class="chart-box monte-distribution-card">
              <h3>Kostenverteilung der Simulation</h3>
              <p class="chart-intro">Histogramm der Simulationsläufe für die Kosten. Markiert sind P50 und P80 als zentrale Steuerungswerte.</p>
              ${analytics.costDistributionSvg}
            </section>
            <section class="chart-box monte-distribution-card">
              <h3>Dauerverteilung der Simulation</h3>
              <p class="chart-intro">Histogramm der Simulationsläufe für die Projektdauer. Die Verteilung macht Streuung und Pufferbedarf sichtbar.</p>
              ${analytics.dayDistributionSvg}
            </section>
          </div>
          <div class="monte-feature-stack">
            <section class="info-card monte-feature-card monte-feature-cost">
              <div class="monte-feature-head">
                <div>
                  <h3>Kostenstruktur der Arbeitspakete</h3>
                  <p>Vergleich der erwarteten Kosten je Arbeitspaket. So wird sichtbar, welche Pakete die Simulation kostenseitig am stärksten prägen.</p>
                </div>
                <span class="badge">Kosten</span>
              </div>
              <div class="monte-feature-layout">
                <div class="chart-box compact-chart-box monte-feature-chart">
                  ${analytics.workCostChartSvg}
                </div>
                <div class="table-wrap monte-side-table">
                  <table class="data-table">
                    <thead>
                      <tr>
                        <th>Arbeitspaket</th>
                        <th>EW Kosten</th>
                        <th>Spannweite</th>
                        <th>Anteil</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${analytics.workRows.map((item) => `
                        <tr>
                          <td><strong>${item.name}</strong></td>
                          <td class="numeric-output">${formatCurrency(item.meanCost)}</td>
                          <td class="numeric-output">${formatCurrency(item.spanCost)}</td>
                          <td class="numeric-output">${formatPercent(item.costShare)}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <section class="info-card monte-feature-card monte-feature-duration">
              <div class="monte-feature-head">
                <div>
                  <h3>Terminstruktur der Arbeitspakete</h3>
                  <p>Die erwartete Dauer je Paket zeigt, wo die Terminseite des Modells am stärksten belastet wird.</p>
                </div>
                <span class="badge">Dauer</span>
              </div>
              <div class="monte-feature-layout">
                <div class="chart-box compact-chart-box monte-feature-chart">
                  ${analytics.workDayChartSvg}
                </div>
                <div class="table-wrap monte-side-table">
                  <table class="data-table">
                    <thead>
                      <tr>
                        <th>Arbeitspaket</th>
                        <th>EW Dauer</th>
                        <th>Spannweite</th>
                        <th>T-Anteil</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${analytics.workRows.map((item) => `
                        <tr>
                          <td><strong>${item.name}</strong></td>
                          <td class="numeric-output">${formatNumber(item.meanDays)} Tage</td>
                          <td class="numeric-output">${formatNumber(item.spanDays)} Tage</td>
                          <td class="numeric-output">${formatPercent(item.meanDays / Math.max(result.meanDays, 1))}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <section class="info-card monte-feature-card monte-feature-risk">
              <div class="monte-feature-head">
                <div>
                  <h3>Risikotreiber im Modell</h3>
                  <p>Die stärksten Risikoereignisse nach durchschnittlichem Kosteneinfluss. So bleibt sichtbar, welche Ereignisse Kosten und Termine des Modells wirklich treiben.</p>
                </div>
                <span class="badge">Treiber</span>
              </div>
              <div class="monte-feature-layout">
                <div class="chart-box compact-chart-box monte-feature-chart">
                  ${analytics.riskDriverChartSvg}
                </div>
                <div class="table-wrap monte-side-table">
                  <table class="data-table">
                    <thead>
                      <tr>
                        <th>Risiko</th>
                        <th>Ø Kosten</th>
                        <th>Ø Tage</th>
                        <th>Treffer</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${analytics.riskRows.map((risk) => `
                        <tr>
                          <td><strong>${risk.name}</strong></td>
                          <td class="numeric-output">${formatCurrency(risk.avgCostImpact)}</td>
                          <td class="numeric-output">${formatNumber(risk.avgDelayImpact)} Tage</td>
                          <td class="numeric-output">${formatPercent(risk.hitRate)}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </div>
          <section class="info-card">
            <div class="monte-section-head">
              <div>
                <h3>Arbeitspakete</h3>
                <p>Pflege hier die Drei-Punkt-Annahmen für Kosten und Dauer. Jedes Arbeitspaket fließt direkt in die Simulationsläufe ein.</p>
              </div>
              <span class="badge">${formatNumber(analytics.workRows.length)} Pakete</span>
            </div>
            <div class="monte-edit-stack">
              ${analytics.workRows.map((item) => `
                <details class="monte-edit-card monte-collapsible-card work-card">
                  <summary class="monte-card-summary">
                    <div class="monte-card-summary-main">
                      <div class="monte-card-summary-topline">
                        <span class="monte-card-toggle" aria-hidden="true"></span>
                        <div class="monte-card-summary-title">
                          <strong>${item.name}</strong>
                          <span>Kosten- und Terminannahmen</span>
                        </div>
                      </div>
                      <div class="monte-card-summary-metrics">
                        <span class="monte-card-metric"><small>EW Kosten</small><strong>${formatCurrency(item.meanCost)}</strong></span>
                        <span class="monte-card-metric"><small>EW Dauer</small><strong>${formatNumber(item.meanDays)} Tage</strong></span>
                        <span class="monte-card-metric"><small>Spannweite</small><strong>${formatCurrency(item.spanCost)}</strong></span>
                        <span class="monte-card-metric"><small>Anteil</small><strong>${formatPercent(item.costShare)}</strong></span>
                      </div>
                    </div>
                    <div class="monte-card-summary-actions">
                      <button class="danger-icon-btn" type="button" aria-label="Arbeitspaket entfernen" data-action="remove-mc-work" data-index="${item.originalIndex}">×</button>
                    </div>
                  </summary>
                  <div class="monte-card-content">
                    <div class="monte-edit-card-head">
                      <div class="monte-edit-title">
                        <label for="mc_wp_name_${item.originalIndex}">Arbeitspaket</label>
                        <input class="compact-input name" id="mc_wp_name_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="name" type="text" value="${item.name}">
                      </div>
                    </div>
                    <div class="monte-edit-group-stack">
                      <section class="monte-edit-group">
                        <div class="monte-edit-group-head">
                          <h4>Kostenannahmen</h4>
                          <p>Drei-Punkt-Kostenwerte.</p>
                        </div>
                        <div class="monte-edit-grid">
                          <div class="form-field">
                            <label for="mc_wp_oc_${item.originalIndex}">Optimistisch in €</label>
                            <input class="compact-input currency-input" id="mc_wp_oc_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="optimisticCost" type="text" inputmode="numeric" value="${formatCurrencyInput(item.optimisticCost)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_wp_mc_${item.originalIndex}">Wahrscheinlich in €</label>
                            <input class="compact-input currency-input" id="mc_wp_mc_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="mostLikelyCost" type="text" inputmode="numeric" value="${formatCurrencyInput(item.mostLikelyCost)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_wp_pc_${item.originalIndex}">Pessimistisch in €</label>
                            <input class="compact-input currency-input" id="mc_wp_pc_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="pessimisticCost" type="text" inputmode="numeric" value="${formatCurrencyInput(item.pessimisticCost)}">
                          </div>
                        </div>
                      </section>
                      <section class="monte-edit-group">
                        <div class="monte-edit-group-head">
                          <h4>Terminannahmen</h4>
                          <p>Drei-Punkt-Dauerwerte.</p>
                        </div>
                        <div class="monte-edit-grid">
                          <div class="form-field">
                            <label for="mc_wp_od_${item.originalIndex}">Optimistisch in Tagen</label>
                            <input class="compact-input integer-input" id="mc_wp_od_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="optimisticDays" type="text" inputmode="numeric" value="${formatIntegerInput(item.optimisticDays)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_wp_md_${item.originalIndex}">Wahrscheinlich in Tagen</label>
                            <input class="compact-input integer-input" id="mc_wp_md_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="mostLikelyDays" type="text" inputmode="numeric" value="${formatIntegerInput(item.mostLikelyDays)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_wp_pd_${item.originalIndex}">Pessimistisch in Tagen</label>
                            <input class="compact-input integer-input" id="mc_wp_pd_${item.originalIndex}" data-mc-work-index="${item.originalIndex}" data-mc-work-field="pessimisticDays" type="text" inputmode="numeric" value="${formatIntegerInput(item.pessimisticDays)}">
                          </div>
                        </div>
                      </section>
                    </div>
                    <div class="monte-edit-metrics monte-work-result-row">
                      <div><span>Erwartungswert Kosten</span><strong>${formatCurrency(item.meanCost)}</strong></div>
                      <div><span>Spannweite Kosten</span><strong>${formatCurrency(item.spanCost)}</strong></div>
                      <div><span>Erwartungswert Dauer</span><strong>${formatNumber(item.meanDays)} Tage</strong></div>
                      <div><span>Spannweite Dauer</span><strong>${formatNumber(item.spanDays)} Tage</strong></div>
                    </div>
                  </div>
                </details>
              `).join("")}
            </div>
            <div style="margin-top:14px;">
              <button class="action-btn add-with-icon" type="button" data-action="add-mc-work"><span class="add-icon" aria-hidden="true">+</span><span>Arbeitspaket hinzufügen</span></button>
            </div>
          </section>
          <section class="info-card">
            <div class="monte-section-head">
              <div>
                <h3>Risikoereignisse</h3>
                <p>Ergänze Eintrittswahrscheinlichkeit sowie Kosten- und Terminwirkung. Risiken werden probabilistisch aktiviert und mit den Arbeitspaketen verrechnet.</p>
              </div>
              <span class="badge">${formatNumber(analytics.riskRows.length)} Risiken</span>
            </div>
            <div class="monte-edit-stack">
              ${analytics.riskRows.map((risk) => `
                <details class="monte-edit-card monte-collapsible-card risk-card">
                  <summary class="monte-card-summary">
                    <div class="monte-card-summary-main">
                      <div class="monte-card-summary-topline">
                        <span class="monte-card-toggle" aria-hidden="true"></span>
                        <div class="monte-card-summary-title">
                          <strong>${risk.name}</strong>
                          <span>Eintritt, Kosten- und Terminwirkung</span>
                        </div>
                      </div>
                      <div class="monte-card-summary-metrics">
                        <span class="monte-card-metric"><small>Ø Kosten</small><strong>${formatCurrency(risk.avgCostImpact)}</strong></span>
                        <span class="monte-card-metric"><small>Ø Tage</small><strong>${formatNumber(risk.avgDelayImpact)} Tage</strong></span>
                        <span class="monte-card-metric"><small>Treffer</small><strong>${formatPercent(risk.hitRate)}</strong></span>
                        <span class="monte-card-metric"><small>Eintritt</small><strong>${formatPercent(risk.probability)}</strong></span>
                      </div>
                    </div>
                    <div class="monte-card-summary-actions">
                      <button class="danger-icon-btn" type="button" aria-label="Risiko entfernen" data-action="remove-mc-risk" data-index="${risk.originalIndex}">×</button>
                    </div>
                  </summary>
                  <div class="monte-card-content">
                    <div class="monte-edit-card-head">
                      <div class="monte-edit-title">
                        <label for="mc_risk_name_${risk.originalIndex}">Risiko</label>
                        <input class="compact-input name" id="mc_risk_name_${risk.originalIndex}" data-mc-risk-index="${risk.originalIndex}" data-mc-risk-field="name" type="text" value="${risk.name}">
                      </div>
                    </div>
                    <div class="monte-edit-group-stack">
                      <section class="monte-edit-group">
                        <div class="monte-edit-group-head">
                          <h4>Eintritt und Kostenwirkung</h4>
                          <p>Eintritt und Kostenwirkung.</p>
                        </div>
                        <div class="monte-edit-grid monte-risk-grid">
                          <div class="form-field">
                            <label for="mc_risk_probability_${risk.originalIndex}">Eintritt in %</label>
                            <input class="compact-input integer-input" id="mc_risk_probability_${risk.originalIndex}" data-mc-risk-index="${risk.originalIndex}" data-mc-risk-field="probability" type="text" inputmode="decimal" value="${formatPercentValue(risk.probability)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_risk_min_cost_${risk.originalIndex}">Kosten minimum in €</label>
                            <input class="compact-input currency-input" id="mc_risk_min_cost_${risk.originalIndex}" data-mc-risk-index="${risk.originalIndex}" data-mc-risk-field="minCostImpact" type="text" inputmode="numeric" value="${formatCurrencyInput(risk.minCostImpact)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_risk_max_cost_${risk.originalIndex}">Kosten maximum in €</label>
                            <input class="compact-input currency-input" id="mc_risk_max_cost_${risk.originalIndex}" data-mc-risk-index="${risk.originalIndex}" data-mc-risk-field="maxCostImpact" type="text" inputmode="numeric" value="${formatCurrencyInput(risk.maxCostImpact)}">
                          </div>
                        </div>
                      </section>
                      <section class="monte-edit-group">
                        <div class="monte-edit-group-head">
                          <h4>Terminwirkung</h4>
                          <p>Zusätzliche Tage bei Eintritt.</p>
                        </div>
                        <div class="monte-edit-grid monte-risk-grid">
                          <div class="form-field">
                            <label for="mc_risk_min_days_${risk.originalIndex}">Termin minimum in Tagen</label>
                            <input class="compact-input integer-input" id="mc_risk_min_days_${risk.originalIndex}" data-mc-risk-index="${risk.originalIndex}" data-mc-risk-field="minDelayDays" type="text" inputmode="numeric" value="${formatIntegerInput(risk.minDelayDays)}">
                          </div>
                          <div class="form-field">
                            <label for="mc_risk_max_days_${risk.originalIndex}">Termin maximum in Tagen</label>
                            <input class="compact-input integer-input" id="mc_risk_max_days_${risk.originalIndex}" data-mc-risk-index="${risk.originalIndex}" data-mc-risk-field="maxDelayDays" type="text" inputmode="numeric" value="${formatIntegerInput(risk.maxDelayDays)}">
                          </div>
                        </div>
                      </section>
                    </div>
                    <div class="monte-edit-metrics monte-risk-result-row">
                      <div><span>Erwartete Kostenwirkung</span><strong>${formatCurrency(risk.avgCostImpact)}</strong></div>
                      <div><span>Trefferquote</span><strong>${formatPercent(risk.hitRate)}</strong></div>
                      <div><span>Erwartete Terminwirkung</span><strong>${formatNumber(risk.avgDelayImpact)} Tage</strong></div>
                      <div><span>Eintrittswahrscheinlichkeit</span><strong>${formatPercent(risk.probability)}</strong></div>
                    </div>
                  </div>
                </details>
              `).join("")}
            </div>
            <div style="margin-top:14px;">
              <button class="action-btn add-with-icon" type="button" data-action="add-mc-risk"><span class="add-icon" aria-hidden="true">+</span><span>Risiko hinzufügen</span></button>
            </div>
          </section>
        </div>
      `;
    }
  },
  earnedValue: {
    key: "earnedValue",
    label: "Leistungswert-Analyse (EVA)",
    subtitle: "Projektperformance im laufenden Projekt",
    render(state) {
      const result = calculateEarnedValueResult(state);
      const analytics = buildEvaAnalytics(state, result);
      const status = getEarnedValueStatus(result, state.earnedValue.thresholds);
      const costTolerance = getEvaCostTolerance(state, Number(state.project?.budget) || 0);
      const evaManagementCards = getEvaManagementCards(state, result);
      const evaManagementRows = [...evaManagementCards].sort((a, b) => {
        const order = { kritisch: 0, angespannt: 1, stabil: 2, neutral: 3 };
        return (order[a.tone] ?? 9) - (order[b.tone] ?? 9);
      });
      const evaDetailView = state.ui?.evaDetailView === "full" ? "full" : "compact";
      const activities = result.activities.map((activity) => {
        const cv = activity.earnedValue - activity.actualCost;
        const sv = activity.earnedValue - activity.plannedValue;
        const cpi = activity.actualCost > 0 ? activity.earnedValue / activity.actualCost : 0;
        const spi = activity.plannedValue > 0 ? activity.earnedValue / activity.plannedValue : 0;
        return {
          ...activity,
          cv,
          sv,
          cpi,
          spi
        };
      });
      const evaRecommendations = [];
      const budget = Number(state.project?.budget) || 0;
      const eacTolerance = Number(state.earnedValue?.thresholds?.eacTolerance) || 180000;
      const tcpiActionLevel = Number(state.earnedValue?.thresholds?.tcpiActionLevel) || 1.02;
      const eacOverrun = result.eac - budget;

      const hasCostPressure = result.cpi < 1 || Math.abs(result.cv) > costTolerance.amount;
      const hasSchedulePressure = result.spi < 1 || result.sv < 0;
      const hasEacPressure = budget > 0 && eacOverrun > eacTolerance;
      const hasTcpiPressure = result.tcpi >= tcpiActionLevel;
      const needsSnapshotSetup = !analytics.snapshotRows.length;

      if (hasCostPressure) {
        evaRecommendations.push({
          area: "Kostensteuerung",
          priority: result.cpi < 0.95 ? "hoch" : "mittel",
          text: "Kostenabweichungen der laufenden Aktivitäten gezielt nachsteuern, Kostentreiber je Vorgang offenlegen und die nächsten Gegenmaßnahmen frühzeitig ableiten. So bleibt sichtbar, wo die größten Hebel liegen und welche Vorgänge zuerst adressiert werden sollten."
        });
      }

      if (hasSchedulePressure) {
        evaRecommendations.push({
          area: "Terminsteuerung",
          priority: result.spi < 0.95 ? "hoch" : "mittel",
          text: "Terminverzug auf kritischen Vorgängen enger verfolgen, Puffer aktivieren und den kritischen Pfad fortschreiben, damit Folgeverzüge früh sichtbar bleiben. So bleibt der Bauablauf besser steuerbar und die Terminlage klar nachvollziehbar."
        });
      }

      if (hasEacPressure) {
        evaRecommendations.push({
          area: "Kostenprognose",
          priority: "hoch",
          text: "Kostenprognose gegen Projektbudget klären und Budgetmaßnahmen mit belastbarer Nachkalkulation absichern, damit die Endkosten früh steuerbar bleiben. Das schafft früh Klarheit über die Budgetwirkung und die notwendige Gegensteuerung."
        });
      }

      if (hasTcpiPressure) {
        evaRecommendations.push({
          area: "Restbudget",
          priority: "mittel",
          text: "Restleistungsdruck aktiv steuern und Maßnahmen für ein effizienteres Restbudget verbindlich hinterlegen."
        });
      }

      if (needsSnapshotSetup) {
        evaRecommendations.push({
          area: "Stichtagssicherung",
          priority: "mittel",
          text: "Einen ersten EVA-Stichtag sichern, damit Trendanalyse und Delta-Vergleiche aufgebaut werden können."
        });
      }

      if (hasCostPressure || hasSchedulePressure || hasEacPressure || hasTcpiPressure) {
        evaRecommendations.push({
          area: "Systemabgleich",
          priority: "niedrig",
          text: "Leistungswert-Analyse (EVA) mit Monte-Carlo-Simulation und Risikoregister spiegeln, damit Kosten-, Termin- und Risikosicht konsistent bleiben."
        });
      }

      if (!evaRecommendations.length) {
        evaRecommendations.push({
          area: "Monitoring",
          priority: "niedrig",
          text: "Die aktuelle Lage ist stabil. Kennzahlen und Stichtage weiter mit der bestehenden Taktung fortschreiben."
        });
      }

      evaRecommendations.sort((a, b) => {
        const order = { hoch: 0, mittel: 1, niedrig: 2 };
        return (order[a.priority] ?? 9) - (order[b.priority] ?? 9);
      });
      const managementFocusItems = evaRecommendations.slice(0, 3);
      const evaDetailColumns = {
        pv: state.ui?.evaDetailColumns?.pv !== false,
        ac: state.ui?.evaDetailColumns?.ac !== false,
        ev: state.ui?.evaDetailColumns?.ev !== false,
        cpi: state.ui?.evaDetailColumns?.cpi !== false,
        spi: state.ui?.evaDetailColumns?.spi !== false,
        fgr: state.ui?.evaDetailColumns?.fgr !== false,
        cv: state.ui?.evaDetailColumns?.cv !== false,
        sv: state.ui?.evaDetailColumns?.sv !== false,
        eac: state.ui?.evaDetailColumns?.eac !== false,
        status: state.ui?.evaDetailColumns?.status !== false
      };
      const activityToneClasses = ["eva-tone-blue", "eva-tone-green", "eva-tone-gold", "eva-tone-rose", "eva-tone-plum", "eva-tone-slate"];
      const evaDetailColumnOptions = [
        { key: "pv", label: "Planwert (PV)" },
        { key: "ac", label: "Istkosten (AC)" },
        { key: "ev", label: "Leistungswert (EV)" },
        { key: "cpi", label: "Kostenleistungsindex (CPI)" },
        { key: "spi", label: "Terminleistungsindex (SPI)" },
        { key: "fgr", label: "Fertigstellungsgrad (FGR)" },
        { key: "cv", label: "Kostenabweichung (CV)" },
        { key: "sv", label: "Terminabweichung (SV)" },
        { key: "eac", label: "Kostenprognose (EAC)" },
        { key: "status", label: "Status" }
      ];
      const evaVisibleDetailColumns = evaDetailView === "compact"
        ? ["pv", "ac", "ev", "cpi", "spi", "status"]
        : evaDetailColumnOptions.filter((column) => evaDetailColumns[column.key]).map((column) => column.key);
      const evaDetailMinWidth = evaDetailView === "full"
        ? 180 + (evaVisibleDetailColumns.length * 132)
        : null;
      const openActivityIds = new Set(state.ui?.evaActivityOpenIds || []);
      const evaIndicatorCards = [
        { code: "PV", value: formatCurrency(result.pv), title: "Planwert (PV)", formula: "PV = geplanter Leistungswert", hint: "Budgetierter Sollwert zum Stichtag." },
        { code: "EV", value: formatCurrency(result.ev), title: "Leistungswert (EV)", formula: "EV = FGR × BAC", hint: "Fertigstellungsgrad × Gesamtbudget – der Wert der geleisteten Arbeit." },
        { code: "AC", value: formatCurrency(result.ac), title: "Istkosten (AC)", formula: "AC = tatsächliche Kosten", hint: "Bisher angefallene Kosten zum Stichtag." },
        { code: "CV", value: formatCurrency(result.cv), title: "Kostenabweichung (CV)", formula: "CV = EV − AC", hint: "Positiv = unter Budget · negativ = über Budget." },
        { code: "SV", value: formatCurrency(result.sv), title: "Terminabweichung (SV)", formula: "SV = EV − PV", hint: "Positiv = vor Plan · negativ = hinter Plan." },
        { code: "CPI", value: result.cpi.toFixed(2), title: "Kostenleistungsindex (CPI)", formula: "CPI = EV / AC", hint: ">1 unter Budget · =1 plangemäß · <1 über Budget" },
        { code: "SPI", value: result.spi.toFixed(2), title: "Terminleistungsindex (SPI)", formula: "SPI = EV / PV", hint: ">1 Terminvorsprung · =1 plangemäß · <1 Terminverzug" },
        { code: "EAC", value: formatCurrency(result.eac), title: "Kostenprognose (EAC)", formula: "EAC = BAC / CPI", hint: "Hochrechnung der Gesamtkosten bei Fortschreibung des CPI-Trends." },
        { code: "ETC", value: formatCurrency(result.etc), title: "Restkosten (ETC)", formula: "ETC = EAC − AC", hint: "Noch benötigtes Budget bis Projektabschluss." },
        { code: "VAC", value: formatCurrency(result.vac), title: "Abschlussabweichung (VAC)", formula: "VAC = BAC − EAC", hint: "Prognostizierte Über-/Unterschreitung des Gesamtbudgets am Ende." },
        { code: "TCPI", value: result.tcpi.toFixed(2), title: "Restleistungsindex (TCPI)", formula: "TCPI = (BAC − EV) / (BAC − AC)", hint: "Benötigte Effizienz für das Restbudget. ≤ 1,0 = erreichbar · > 1,2 = unrealistisch." }
      ];
      return `
        <div class="module-shell">
          <div class="module-title">
            <div>
              <h2>Leistungswert-Analyse (EVA)</h2>
              <p>EVA ist das laufende Controlling-Modul. Es nutzt die gemeinsamen Projektdaten, pflegt Aktivitäten und Schwellenwerte aber bewusst separat.</p>
            </div>
            <span class="badge">Controlling</span>
          </div>
          <div class="eva-kpi-stack">
            <div class="kpi-grid" id="eva-kpi-overview">
              <article class="kpi-card" id="eva-pv-card"><div class="kpi-label">PV</div><div class="kpi-value">${formatCurrency(result.pv)}</div><div class="kpi-sub">Planwert (PV)</div></article>
              <article class="kpi-card gold"><div class="kpi-label">EV</div><div class="kpi-value">${formatCurrency(result.ev)}</div><div class="kpi-sub">Leistungswert (EV)</div></article>
              <article class="kpi-card green"><div class="kpi-label">CPI</div><div class="kpi-value">${result.cpi.toFixed(2)}</div><div class="kpi-sub">Kostenleistungsindex (CPI)</div></article>
              <article class="kpi-card blue"><div class="kpi-label">SPI</div><div class="kpi-value">${result.spi.toFixed(2)}</div><div class="kpi-sub">Terminleistungsindex (SPI) · ${status.level}</div></article>
            </div>
            <div class="kpi-grid eva-kpi-secondary">
              <article class="kpi-card"><div class="kpi-label">CV</div><div class="kpi-value">${formatCurrency(result.cv)}</div><div class="kpi-sub">Kostenabweichung (CV)</div></article>
              <article class="kpi-card gold"><div class="kpi-label">SV</div><div class="kpi-value">${formatCurrency(result.sv)}</div><div class="kpi-sub">Terminabweichung (SV)</div></article>
              <article class="kpi-card green"><div class="kpi-label">EAC</div><div class="kpi-value">${formatCurrency(result.eac)}</div><div class="kpi-sub">Kostenprognose (EAC)</div></article>
              <article class="kpi-card blue"><div class="kpi-label">TCPI</div><div class="kpi-value">${result.tcpi.toFixed(2)}</div><div class="kpi-sub">Restleistungsindex (TCPI)</div></article>
            </div>
          </div>
          <details class="info-card eva-fold-card eva-fold-primary eva-fold-overview" open>
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                  <h3>Steuerung, Steuerungsbild und Steuerungsfokus</h3>
                  <p>Hier legst du fest, wie die EVA gelesen wird, siehst den aktuellen Projektzustand und erkennst, wo als Nächstes gehandelt werden sollte.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">Übersicht</span>
              </div>
            </summary>
            <div class="eva-fold-body">
          <div class="card-grid eva-upper-grid">
            <section class="info-card control-card">
              <h3>Steuerung</h3>
              <p>Hier stellst du die Bewertungsregeln der EVA ein. Die Werte entscheiden, wann ein Projekt noch im Rahmen liegt und wann es kritisch wird.</p>
              <div class="eva-control-grid">
                <div class="eva-control-field">
                  <label for="eva_status_date">Statusdatum</label>
                  <input id="eva_status_date" data-eva-field="statusDate" type="date" value="${state.earnedValue.statusDate}">
                  <div class="eva-control-hint">Basis für Stichtag und Trend</div>
                </div>
                <div class="eva-control-field">
                  <label for="eva_spi_yellow">Terminleistungsindex Schwelle (SPI)</label>
                  <div class="eva-stepper-field">
                    <input id="eva_spi_yellow" data-eva-threshold-field="spiYellow" data-eva-threshold-kind="decimal" type="text" inputmode="decimal" value="${formatDecimalInput(state.earnedValue.thresholds.spiYellow, 2)}">
                    <div class="eva-stepper-buttons">
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="spiYellow" data-step="0.01" data-direction="up" aria-label="SPI-Warnschwelle erhöhen">▲</button>
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="spiYellow" data-step="0.01" data-direction="down" aria-label="SPI-Warnschwelle verringern">▼</button>
                    </div>
                  </div>
                  <div class="eva-control-hint">Warnung bei Werten unter dem Grenzwert</div>
                </div>
                <div class="eva-control-field">
                  <label for="eva_cpi_yellow">Kostenleistungsindex Schwelle (CPI)</label>
                  <div class="eva-stepper-field">
                    <input id="eva_cpi_yellow" data-eva-threshold-field="cpiYellow" data-eva-threshold-kind="decimal" type="text" inputmode="decimal" value="${formatDecimalInput(state.earnedValue.thresholds.cpiYellow, 2)}">
                    <div class="eva-stepper-buttons">
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="cpiYellow" data-step="0.01" data-direction="up" aria-label="CPI-Warnschwelle erhöhen">▲</button>
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="cpiYellow" data-step="0.01" data-direction="down" aria-label="CPI-Warnschwelle verringern">▼</button>
                    </div>
                  </div>
                  <div class="eva-control-hint">Warnung bei Werten unter dem Grenzwert</div>
                </div>
                <div class="eva-control-field">
                  <label for="eva_eac_tolerance">Kostenprognose Toleranz (EAC) in €</label>
                  <div class="eva-stepper-field">
                    <input id="eva_eac_tolerance" data-eva-threshold-field="eacTolerance" data-eva-threshold-kind="integer" type="text" inputmode="numeric" value="${formatNumber(state.earnedValue.thresholds.eacTolerance || 180000)}">
                    <div class="eva-stepper-buttons">
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="eacTolerance" data-step="1000" data-direction="up" aria-label="EAC-Toleranz erhöhen">▲</button>
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="eacTolerance" data-step="1000" data-direction="down" aria-label="EAC-Toleranz verringern">▼</button>
                    </div>
                  </div>
                  <div class="eva-control-hint">Zulässige Kostenabweichung</div>
                </div>
                <div class="eva-control-field">
                  <label for="eva_tcpi_action">Restleistungsindex Schwelle (TCPI)</label>
                  <div class="eva-stepper-field">
                    <input id="eva_tcpi_action" data-eva-threshold-field="tcpiActionLevel" data-eva-threshold-kind="decimal" type="text" inputmode="decimal" value="${formatDecimalInput(state.earnedValue.thresholds.tcpiActionLevel || 1.02, 2)}">
                    <div class="eva-stepper-buttons">
                      <button type="button" class="eva-stepper-btn" data-action="eva-threshold-step" data-field="tcpiActionLevel" data-step="0.01" data-direction="up" aria-label="TCPI-Maßnahmenschwelle erhöhen">▲</button>
                      <button type="button" class="eva-stepper-btn" data-field="tcpiActionLevel" data-action="eva-threshold-step" data-step="0.01" data-direction="down" aria-label="TCPI-Maßnahmenschwelle verringern">▼</button>
                    </div>
                  </div>
                  <div class="eva-control-hint">Maßnahmen ab diesem Wert</div>
                </div>
                <div class="eva-control-field eva-control-field-jump" data-action="jump-eva-activities">
                  <label>Aktivitäten<br>und Arbeitspakete</label>
                  <button type="button" class="form-note eva-count-note eva-jump-note" data-action="jump-eva-activities">${activities.length} Pakete öffnen</button>
                  <div class="eva-control-hint">Zum Öffnen anklicken.<br>Direkt zu den Arbeitspaketen springen.</div>
                </div>
              </div>
              <div class="eva-tolerance-card">
                <div class="eva-tolerance-head">
                  <span>Kostentoleranz</span>
                  <strong id="eva_tolerance_amount_display">= ${formatCurrency(costTolerance.amount)}</strong>
                </div>
                <div id="eva_tolerance_percent_display" class="eva-tolerance-percent">${new Intl.NumberFormat("de-DE", { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(costTolerance.percent)} %</div>
                <input id="eva_cv_tolerance_percent" class="eva-tolerance-range" type="range" min="0" max="10" step="0.1" value="${costTolerance.percent.toFixed(1)}">
                <div class="eva-tolerance-ticks" aria-hidden="true">
                  ${Array.from({ length: 11 }, (_, index) => `<span class="eva-tolerance-tick${index === 0 || index === 10 ? " edge" : ""}"></span>`).join("")}
                </div>
                <div class="eva-tolerance-scale" aria-hidden="true">
                  ${Array.from({ length: 11 }, (_, index) => `<span>${index},0 %</span>`).join("")}
                </div>
                <small>Der Regler legt fest, wie streng Kostenabweichungen bewertet werden. ${formatCurrency(costTolerance.amount)} ist die aktuelle Toleranzgrenze. Höher = toleranter, niedriger = strenger.</small>
              </div>
            </section>
            <section class="info-card steering-card steering-card-compact">
              <h3>Steuerungsbild</h3>
              <p>Das Steuerungsbild zeigt die aktuelle Lage zum gewählten Statusdatum auf einen Blick. Es fasst die wichtigsten Kennzahlen zusammen und macht sofort sichtbar, ob Kosten, Termin und Prognose noch akzeptabel sind.</p>
              <div class="eva-steering-grid">
                <div class="eva-steering-item eva-status-card eva-${status.level}">
                  <span>Status</span>
                  <strong>${status.level}</strong>
                  <small>${status.message}</small>
                </div>
                <div class="eva-steering-item">
                  <span>Kosten- und Terminleistung</span>
                    <div class="eva-steering-metric-stack">
                      <div class="eva-steering-metric">
                        <strong>CPI: ${result.cpi.toFixed(2)}</strong>
                        <small>Kostenleistungsindex: zeigt, wie wirtschaftlich bisher gearbeitet wurde. Unter 1,00 bedeutet: teurer als geplant.</small>
                      </div>
                      <div class="eva-steering-metric">
                        <strong>SPI: ${result.spi.toFixed(2)}</strong>
                        <small>Terminleistungsindex: zeigt, ob der Fortschritt zum Zeitplan passt. Unter 1,00 bedeutet: Rückstand.</small>
                      </div>
                    </div>
                  </div>
                  <div class="eva-steering-item">
                    <span>Abweichungen</span>
                    <div class="eva-steering-metric-stack">
                      <div class="eva-steering-metric">
                        <strong>CV: ${formatCurrency(result.cv)}</strong>
                        <small>Kostenabweichung: Abstand zwischen Leistung und Kosten. Negativ bedeutet: Kosten höher als Leistung.</small>
                      </div>
                      <div class="eva-steering-metric">
                        <strong>SV: ${formatCurrency(result.sv)}</strong>
                        <small>Terminabweichung: Abstand zwischen Leistung und Plan. Negativ bedeutet: Projekt liegt zurück.</small>
                      </div>
                    </div>
                  </div>
                  <div class="eva-steering-item">
                    <span>Prognose und Restleistung</span>
                    <div class="eva-steering-metric-stack">
                      <div class="eva-steering-metric">
                        <strong>EAC: ${formatCurrency(result.eac)}</strong>
                        <small>Kostenprognose: erwartete Gesamtkosten am Projektende bei gleicher Entwicklung.</small>
                      </div>
                      <div class="eva-steering-metric">
                        <strong>TCPI: ${result.tcpi.toFixed(2)}</strong>
                        <small>Restleistungsindex: zeigt, wie effizient das Restbudget eingesetzt werden muss.</small>
                      </div>
                    </div>
                  </div>
                </div>
              </section>
              <section class="info-card eva-focus-card eva-focus-card-wide">
                <h3>Steuerungsfokus</h3>
                <p>Der Steuerungsfokus zeigt die nächsten Prioritäten. Er macht klar, wo jetzt zuerst gehandelt werden sollte und welche Themen die Projektsituation am stärksten beeinflussen.</p>
              <ul class="eva-focus-list">
                ${managementFocusItems.map((item, index) => `
                  <li class="eva-focus-item eva-focus-${item.priority}">
                    <div class="eva-focus-head">
                      <span>${item.area}</span>
                      <em class="eva-focus-priority eva-focus-priority-${item.priority}">${item.priority}</em>
                    </div>
                    <small>${item.text}</small>
                  </li>
                `).join("")}
              </ul>
            </section>
          </div>
            </div>
          </details>
          <details class="info-card eva-fold-card eva-trend-card eva-fold-primary" open>
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                  <h3>Stichtag-Verlauf & Trendanalyse</h3>
                  <p>Entwicklung von Kosten- und Terminleistung über mehrere EVA-Stichtage.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">${analytics.snapshotRows.length} Stichtag${analytics.snapshotRows.length === 1 ? "" : "e"}</span>
              </div>
            </summary>
            <div class="eva-fold-body">
              <div class="eva-trend-head">
                <p class="chart-intro">Der Verlauf zeigt die Entwicklung von Kosten- und Terminleistung je Stichtag. Kostenleistungsindex (CPI) und Terminleistungsindex (SPI) laufen auf der linken Achse, Kostenprognose (EAC) und Gesamtbudget (BAC) auf der rechten Kostenskala.</p>
                <div class="eva-trend-actions">
                  <button class="action-btn" type="button" data-action="clear-eva-snapshots">Verlauf löschen</button>
                  <button class="action-btn" type="button" data-action="save-eva-snapshot">+ Stichtag speichern</button>
                </div>
              </div>
              ${analytics.snapshotRows.length ? `
              <div class="analysis-grid eva-delta-grid">
                <div class="mini-card eva-delta-card">
                  <span>Seit letztem Stichtag · Kostenleistungsindex (CPI)</span>
                  <strong class="${(analytics.currentDeltas?.cpi || 0) >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${formatSignedNumber(analytics.currentDeltas?.cpi || 0, 2)}</strong>
                </div>
                <div class="mini-card eva-delta-card">
                  <span>Seit letztem Stichtag · Terminleistungsindex (SPI)</span>
                  <strong class="${(analytics.currentDeltas?.spi || 0) >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${formatSignedNumber(analytics.currentDeltas?.spi || 0, 2)}</strong>
                </div>
                <div class="mini-card eva-delta-card">
                  <span>Seit letztem Stichtag · Kostenprognose (EAC)</span>
                  <strong class="${(analytics.currentDeltas?.eac || 0) <= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${formatSignedCurrency(analytics.currentDeltas?.eac || 0)}</strong>
                </div>
                <div class="mini-card eva-delta-card">
                  <span>Seit letztem Stichtag · Fertigstellungsgrad (FGR)</span>
                  <strong class="${(analytics.currentDeltas?.fgr || 0) >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${formatSignedNumber(analytics.currentDeltas?.fgr || 0, 1)} %</strong>
                </div>
              </div>
              <div class="chart-box compact-chart-box eva-trend-chart-box">
                ${analytics.trendSvg}
              </div>
              <div class="table-wrap eva-trend-table">
                <table class="data-table">
                  <thead>
                    <tr>
                      <th>Stichtag</th>
                      <th>Leistungswert (EV)</th>
                      <th>Istkosten (AC)</th>
                      <th>Kostenleistungsindex (CPI)</th>
                      <th>Δ Kostenleistungsindex (CPI)</th>
                      <th>Terminleistungsindex (SPI)</th>
                      <th>Δ Terminleistungsindex (SPI)</th>
                      <th>Kostenprognose (EAC)</th>
                      <th>Δ Kostenprognose (EAC)</th>
                      <th>Fertigstellungsgrad (FGR)</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    ${analytics.snapshotRows.map((row, index) => `
                      <tr>
                        <td><strong>${formatDate(row.statusDate)}</strong></td>
                        <td class="numeric-output">${formatCompactCurrency(row.ev)}</td>
                        <td class="numeric-output">${formatCompactCurrency(row.ac)}</td>
                        <td class="numeric-output ${row.cpi >= 1 ? "eva-cell-positive" : "eva-cell-negative"}">${row.cpi.toFixed(2).replace(".", ",")}</td>
                        <td class="numeric-output ${row.deltaCpi === null ? "" : row.deltaCpi >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${row.deltaCpi === null ? "—" : formatSignedNumber(row.deltaCpi, 2)}</td>
                        <td class="numeric-output ${row.spi >= 1 ? "eva-cell-positive" : "eva-cell-negative"}">${row.spi.toFixed(2).replace(".", ",")}</td>
                        <td class="numeric-output ${row.deltaSpi === null ? "" : row.deltaSpi >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${row.deltaSpi === null ? "—" : formatSignedNumber(row.deltaSpi, 2)}</td>
                        <td class="numeric-output ${row.eac <= ((row.bac || 0) || Number(state.project?.budget) || 0) ? "eva-cell-positive" : "eva-cell-negative"}">${formatCompactCurrency(row.eac)}</td>
                        <td class="numeric-output ${row.deltaEac === null ? "" : row.deltaEac <= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${row.deltaEac === null ? "—" : formatSignedCompactCurrency(row.deltaEac)}</td>
                        <td class="numeric-output">${formatNumber(row.fgr)}%</td>
                        <td class="numeric-output"><button class="snapshot-delete-inline" type="button" aria-label="Stichtag löschen" data-action="delete-eva-snapshot" data-index="${index}">×</button></td>
                      </tr>
                    `).join("")}
                  </tbody>
                </table>
              </div>
              ` : `
              <div class="eva-trend-empty">
                <div class="eva-trend-empty-copy">
                  <div class="eva-trend-empty-badge">Startpunkt fehlt</div>
                  <h4>Noch kein EVA-Stichtag gespeichert</h4>
                  <p>Speichere den aktuellen Stand, um Kostenleistungsindex (CPI), Terminleistungsindex (SPI), Kostenprognose (EAC) und Fertigstellungsgrad über mehrere Stichtage hinweg vergleichbar zu machen.</p>
                </div>
                <div class="eva-trend-empty-preview">
                  <div class="eva-trend-preview-item">
                    <span>Danach sichtbar</span>
                    <strong>Trendlinie für Kostenleistungsindex (CPI), Terminleistungsindex (SPI) und Kostenprognose (EAC)</strong>
                  </div>
                  <div class="eva-trend-preview-item">
                    <span>Danach sichtbar</span>
                    <strong>Delta-Vergleich zum letzten Stichtag für Leistungsindizes und Kostenprognose</strong>
                  </div>
                  <div class="eva-trend-preview-item">
                    <span>Danach sichtbar</span>
                    <strong>Historie als Tabelle und Managementspur</strong>
                  </div>
                </div>
              </div>
              `}
            </div>
          </details>
          <details class="info-card eva-fold-card eva-fold-management" open>
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                  <h3>Managementsicht der EVA</h3>
                  <p>Status, Einordnung und Managementblick auf die wichtigsten EVA-Kennzahlen.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">Management</span>
              </div>
            </summary>
            <div class="eva-fold-body">
              <p class="chart-intro">Die Managementsicht interpretiert die Kernkennzahlen mit Toleranzen und übersetzt sie in Status, Einordnung und Handlungsempfehlung.</p>
              <div class="eva-management-legend" aria-label="Legende der Managementsicht">
                <span class="eva-management-pill eva-management-pill-kritisch">rot = kritisch</span>
                <span class="eva-management-pill eva-management-pill-angespannt">gelb = angespannt</span>
                <span class="eva-management-pill eva-management-pill-stabil">grün = stabil</span>
                <span class="eva-management-pill eva-management-pill-neutral">grau = neutral</span>
              </div>
              <div class="table-wrap eva-management-table-wrap">
                <table class="data-table eva-management-table">
                  <thead>
                    <tr>
                      <th>Kategorie</th>
                      <th>Wert</th>
                      <th>Status</th>
                      <th>Handlung</th>
                      <th>Einordnung</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${evaManagementRows.map((card) => `
                      <tr class="eva-management-row eva-management-${card.tone}">
                        <td><strong>${card.title}</strong></td>
                        <td class="numeric-output">${card.value}</td>
                        <td><span class="eva-management-pill eva-management-pill-${card.tone}">${card.status}</span></td>
                        <td><strong>${card.actionShort}</strong></td>
                        <td>${card.message}</td>
                      </tr>
                    `).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          </details>
          <details class="info-card eva-fold-card eva-action-summary eva-fold-actions" open>
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                  <h3>Handlungsempfehlungen</h3>
                  <p>Vertiefte Maßnahmen auf Basis der aktuellen EVA-Kennzahlen.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">Vertiefung</span>
              </div>
            </summary>
            <div class="eva-fold-body">
              <p class="chart-intro">Die Empfehlungen leiten sich direkt aus dem aktuellen Bild von Kostenleistungsindex (CPI), Terminleistungsindex (SPI), Kostenabweichung (CV) und Terminabweichung (SV) ab und vertiefen die kompakten Handlungshinweise aus der Managementtabelle.</p>
              <div class="table-wrap eva-recommendation-table-wrap">
                <table class="data-table eva-recommendation-table">
                  <thead>
                    <tr>
                      <th>Nr.</th>
                      <th>Priorität</th>
                      <th>Bereich</th>
                      <th>Empfehlung</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${evaRecommendations.map((item, index) => `
                      <tr class="eva-recommendation-row eva-recommendation-${item.priority}">
                        <td class="numeric-output">${index + 1}</td>
                        <td><span class="eva-recommendation-pill eva-recommendation-pill-${item.priority}">${item.priority}</span></td>
                        <td><strong>${item.area}</strong></td>
                        <td>${item.text}</td>
                      </tr>
                    `).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          </details>
          <details class="info-card eva-fold-card eva-primary-analysis eva-fold-charts" open>
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                  <h3>Kerncharts der EVA</h3>
                  <p>S-Kurve und Fortschrittsvergleich als zentrale grafische Managementsicht.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">Kerncharts</span>
              </div>
            </summary>
            <div class="eva-fold-body">
              <p class="chart-intro">Diese beiden Grafiken tragen die tägliche Managementsicht: S-Kurve für Kosten und Leistung, Fortschrittsvergleich für Plan und Ist.</p>
              <div class="card-grid eva-visual-grid eva-primary-analysis-grid">
              <section class="chart-box eva-visual-card">
                <div class="eva-visual-card-head">
                  <div>
                    <h3>S-Kurve: Planwert (PV) · Leistungswert (EV) · Istkosten (AC)</h3>
                    <p>Die S-Kurve ordnet Planwert, Leistungswert und Istkosten in eine gemeinsame Fortschrittslogik ein.</p>
                  </div>
                  <div class="eva-legend">
                    <span class="eva-legend-item eva-pv">Planwert (PV)</span>
                    <span class="eva-legend-item eva-ev">Leistungswert (EV)</span>
                    <span class="eva-legend-item eva-ac">Istkosten (AC)</span>
                  </div>
                </div>
                ${analytics.sCurveSvg}
              </section>
              <section class="chart-box eva-visual-card">
                <div class="eva-visual-card-head">
                  <div>
                    <h3>Fertigstellungsgrad: Plan vs. Ist</h3>
                    <p>Der Vergleich zeigt den abgeleiteten Fertigstellungsgrad (FGR) aus der EVA gegen den zusätzlich gepflegten visuellen Baufortschritt.</p>
                  </div>
                  <div class="eva-legend">
                    <span class="eva-legend-item eva-plan">Plan-Fertigstellungsgrad (FGR)</span>
                    <span class="eva-legend-item eva-ist">Ist-Fertigstellungsgrad (FGR)</span>
                  </div>
                </div>
                ${analytics.progressSvg}
              </section>
              </div>
            </div>
          </details>
          <details class="info-card eva-fold-card eva-secondary-analysis eva-fold-secondary">
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                <h3>DIN 276 – Näherung nach Kostengruppen</h3>
                <p>Vertiefende Kostengruppen-Sicht auf Basis einer heuristischen Zuordnung.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">Vertiefung</span>
              </div>
            </summary>
            <div class="card-grid eva-din-grid">
              <div class="chart-box compact-chart-box eva-din-box">
                <div class="eva-chart-meta">
                  <h4>Budget nach Kostengruppe</h4>
                  <p>Verteilung des Planwerts auf die erkannten Kostengruppen.</p>
                </div>
                ${analytics.donutSvg}
              </div>
              <div class="chart-box compact-chart-box eva-din-box">
                <div class="eva-chart-meta">
                  <h4>Planwert / Leistungswert / Istkosten je Kostengruppe</h4>
                  <div class="eva-legend">
                    <span class="eva-legend-item eva-pv">Planwert (PV)</span>
                    <span class="eva-legend-item eva-ev">Leistungswert (EV)</span>
                    <span class="eva-legend-item eva-ac">Istkosten (AC)</span>
                  </div>
                </div>
                ${analytics.costGroupSvg}
              </div>
            </div>
          </details>
          <details class="info-card eva-fold-card eva-fold-table" id="eva-detail-panel">
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                <h3>Detailauswertung nach Vorgängen (${formatDate(state.earnedValue.statusDate)})</h3>
                <p>Budget, Istkosten, Leistungswert und Status je Vorgang.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">Detailtabelle</span>
              </div>
            </summary>
            <div class="eva-table-view-toggle" role="group" aria-label="Ansicht der Detailauswertung">
              <button class="action-btn ${evaDetailView === "compact" ? "active" : ""}" type="button" data-action="set-eva-detail-view" data-view="compact">Kompaktsicht</button>
              <button class="action-btn ${evaDetailView === "full" ? "active" : ""}" type="button" data-action="set-eva-detail-view" data-view="full">Vollsicht</button>
            </div>
            ${evaDetailView === "full" ? `
            <div class="eva-detail-column-toggle" role="group" aria-label="Spalten in der Vollsicht">
              ${evaDetailColumnOptions.map((column) => `
              <button class="action-btn ${evaDetailColumns[column.key] ? "active" : ""}" type="button" data-action="toggle-eva-detail-column" data-field="${column.key}">${column.label}</button>
              `).join("")}
            </div>
            ` : ""}
            <div class="table-wrap eva-detail-table-wrap ${evaDetailView === "full" ? "eva-detail-table-scroll" : ""}">
              <table class="data-table eva-detail-table eva-detail-table-${evaDetailView}"${evaDetailMinWidth ? ` style="min-width:${evaDetailMinWidth}px"` : ""}>
                <thead>
                  <tr>
                    <th><span class="eva-th-two-line"><span>Vorgang</span><span></span></span></th>
                    ${(evaDetailView === "compact" || evaDetailColumns.pv) ? `<th><span class="eva-th-two-line"><span>Planwert</span><span>(PV)</span></span></th>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.ac) ? `<th><span class="eva-th-two-line"><span>Istkosten</span><span>(AC)</span></span></th>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.ev) ? `<th><span class="eva-th-two-line"><span>${evaDetailView === "compact" ? "Leistungs-" : "Leistungswert"}</span><span>${evaDetailView === "compact" ? "wert (EV)" : "(EV)"}</span></span></th>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.cpi) ? `<th><span class="eva-th-two-line"><span>${evaDetailView === "compact" ? "Kostenleistungs-" : "Kostenleistungsindex"}</span><span>${evaDetailView === "compact" ? "index (CPI)" : "(CPI)"}</span></span></th>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.spi) ? `<th><span class="eva-th-two-line"><span>${evaDetailView === "compact" ? "Terminleistungs-" : "Terminleistungsindex"}</span><span>${evaDetailView === "compact" ? "index (SPI)" : "(SPI)"}</span></span></th>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.fgr) ? `<th><span class="eva-th-two-line"><span>Fertigstellungsgrad</span><span>(FGR)</span></span></th>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.cv) ? `<th><span class="eva-th-two-line"><span>Kostenabweichung</span><span>(CV)</span></span></th>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.sv) ? `<th><span class="eva-th-two-line"><span>Terminabweichung</span><span>(SV)</span></span></th>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.eac) ? `<th><span class="eva-th-two-line"><span>Kostenprognose</span><span>(EAC)</span></span></th>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.status) ? `<th><span class="eva-th-two-line"><span>Status</span><span></span></span></th>` : ""}
                  </tr>
                </thead>
                <tbody>
                  ${analytics.activities.map((activity) => `
                    <tr>
                      <td><strong>${activity.name}</strong></td>
                      ${(evaDetailView === "compact" || evaDetailColumns.pv) ? `<td class="numeric-output">${formatCurrency(activity.plannedValue)}</td>` : ""}
                      ${(evaDetailView === "compact" || evaDetailColumns.ac) ? `<td class="numeric-output">${formatCurrency(activity.actualCost)}</td>` : ""}
                      ${(evaDetailView === "compact" || evaDetailColumns.ev) ? `<td class="numeric-output">${formatCurrency(activity.earnedValue)}</td>` : ""}
                      ${(evaDetailView === "compact" || evaDetailColumns.cpi) ? `<td class="numeric-output ${activity.cpi >= 1 ? "eva-cell-positive" : "eva-cell-negative"}">${activity.cpi ? activity.cpi.toFixed(2).replace(".", ",") : "—"}</td>` : ""}
                      ${(evaDetailView === "compact" || evaDetailColumns.spi) ? `<td class="numeric-output ${activity.spi >= 1 ? "eva-cell-positive" : "eva-cell-negative"}">${activity.spi ? activity.spi.toFixed(2).replace(".", ",") : "—"}</td>` : ""}
                      ${(evaDetailView === "full" && evaDetailColumns.fgr) ? `<td class="numeric-output">${formatNumber(activity.fgrPercent)}%</td>` : ""}
                      ${(evaDetailView === "full" && evaDetailColumns.cv) ? `<td class="numeric-output ${activity.cv >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${formatCurrency(activity.cv)}</td>` : ""}
                      ${(evaDetailView === "full" && evaDetailColumns.sv) ? `<td class="numeric-output ${activity.sv >= 0 ? "eva-cell-positive" : "eva-cell-negative"}">${formatCurrency(activity.sv)}</td>` : ""}
                      ${(evaDetailView === "full" && evaDetailColumns.eac) ? `<td class="numeric-output ${activity.eac <= activity.bac ? "eva-cell-positive" : "eva-cell-negative"}">${activity.eac ? formatCurrency(activity.eac) : "—"}</td>` : ""}
                      ${(evaDetailView === "compact" || evaDetailColumns.status) ? `<td><span class="eva-row-status eva-${String(activity.status).toLowerCase()}">${activity.status}</span></td>` : ""}
                    </tr>
                  `).join("")}
                  <tr class="eva-total-row">
                    <td style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);color:#163f63;font-weight:800;"><strong>Gesamt (Projekt)</strong></td>
                    ${(evaDetailView === "compact" || evaDetailColumns.pv) ? `<td class="numeric-output" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatCurrency(result.pv)}</td>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.ac) ? `<td class="numeric-output" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatCurrency(result.ac)}</td>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.ev) ? `<td class="numeric-output" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatCurrency(result.ev)}</td>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.cpi) ? `<td class="numeric-output ${result.cpi >= 1 ? "eva-cell-positive" : "eva-cell-negative"}" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${result.cpi.toFixed(2).replace(".", ",")}</td>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.spi) ? `<td class="numeric-output ${result.spi >= 1 ? "eva-cell-positive" : "eva-cell-negative"}" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${result.spi.toFixed(2).replace(".", ",")}</td>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.fgr) ? `<td class="numeric-output" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatNumber(result.progressPercent)}%</td>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.cv) ? `<td class="numeric-output ${result.cv >= 0 ? "eva-cell-positive" : "eva-cell-negative"}" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatCurrency(result.cv)}</td>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.sv) ? `<td class="numeric-output ${result.sv >= 0 ? "eva-cell-positive" : "eva-cell-negative"}" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatCurrency(result.sv)}</td>` : ""}
                    ${(evaDetailView === "full" && evaDetailColumns.eac) ? `<td class="numeric-output ${result.eac <= ((Number(state.project?.budget) || result.pv || result.ev || result.ac)) ? "eva-cell-positive" : "eva-cell-negative"}" style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);">${formatCompactCurrency(result.eac)}</td>` : ""}
                    ${(evaDetailView === "compact" || evaDetailColumns.status) ? `<td style="background:linear-gradient(180deg, #d4e0e9, #c8d5df);"><span class="eva-row-status eva-${String(status.level).toLowerCase()}">${status.level}</span></td>` : ""}
                  </tr>
                </tbody>
              </table>
            </div>
          </details>
          <details class="info-card eva-fold-card eva-fold-activities" id="eva-activities-panel"${state.ui?.evaActivitiesPanelOpen ? " open" : ""}>
            <summary class="eva-fold-summary">
              <div class="eva-fold-main">
                <span class="eva-fold-toggle" aria-hidden="true"></span>
                <div>
                <h3>Aktivitäten / Arbeitspakete</h3>
                <p>Planwert, Leistungswert und Istkosten plus ergänzender visueller Fortschritt je Aktivität.</p>
                </div>
              </div>
              <div class="eva-fold-meta">
                <span class="badge">${activities.length} Pakete</span>
              </div>
            </summary>
            <div class="eva-edit-stack">
              ${activities.map((activity, index) => `
                <details class="monte-edit-card monte-collapsible-card eva-card ${activityToneClasses[index % activityToneClasses.length]}" data-eva-activity-id="${activity.id}"${openActivityIds.has(String(activity.id)) ? " open" : ""}>
                  <summary class="monte-card-summary">
                    <div class="monte-card-summary-main">
                      <div class="monte-card-summary-topline">
                        <span class="monte-card-toggle" aria-hidden="true"></span>
                        <div class="monte-card-summary-title">
                          <strong>${activity.name}</strong>
                          <span>Planwert, Leistungswert, Istkosten und visueller Fortschritt</span>
                        </div>
                      </div>
                      <div class="monte-card-summary-metrics">
                        <span class="monte-card-metric"><small>Planwert (PV)</small><strong>${formatCurrency(activity.plannedValue)}</strong></span>
                        <span class="monte-card-metric"><small>Leistungswert (EV)</small><strong>${formatCurrency(activity.earnedValue)}</strong></span>
                        <span class="monte-card-metric"><small>Istkosten (AC)</small><strong>${formatCurrency(activity.actualCost)}</strong></span>
                        <span class="monte-card-metric"><small>Fertigstellungsgrad (FGR)</small><strong>${formatNumber(activity.fgrPercent)} %</strong></span>
                      </div>
                    </div>
                    <div class="monte-card-summary-actions">
                      <button class="danger-icon-btn" type="button" aria-label="Aktivität entfernen" data-action="remove-eva-activity" data-index="${index}">×</button>
                    </div>
                  </summary>
                  <div class="monte-card-content">
                    <div class="monte-edit-card-head">
                      <div class="monte-edit-title">
                        <label for="eva_name_${index}">Aktivität</label>
                        <input class="compact-input name" id="eva_name_${index}" data-eva-activity-index="${index}" data-eva-activity-field="name" type="text" value="${activity.name}">
                      </div>
                    </div>
                    <div class="monte-edit-group-stack">
                      <section class="monte-edit-group">
                        <div class="monte-edit-group-head">
                          <h4>Kosten- und Leistungswerte</h4>
                          <p>Diese Werte gehen direkt in die EVA ein.</p>
                        </div>
                        <div class="monte-edit-grid eva-edit-grid">
                          <div class="form-field">
                            <label for="eva_pv_${index}">Planwert (PV) in €</label>
                            <input class="compact-input currency-input" id="eva_pv_${index}" data-eva-activity-index="${index}" data-eva-activity-field="plannedValue" type="text" inputmode="numeric" value="${formatCurrencyInput(activity.plannedValue)}">
                          </div>
                          <div class="form-field">
                            <label for="eva_ev_${index}">Leistungswert (EV) in €</label>
                            <input class="compact-input currency-input" id="eva_ev_${index}" data-eva-activity-index="${index}" data-eva-activity-field="earnedValue" type="text" inputmode="numeric" value="${formatCurrencyInput(activity.earnedValue)}">
                          </div>
                          <div class="form-field">
                            <label for="eva_ac_${index}">Istkosten (AC) in €</label>
                            <input class="compact-input currency-input" id="eva_ac_${index}" data-eva-activity-index="${index}" data-eva-activity-field="actualCost" type="text" inputmode="numeric" value="${formatCurrencyInput(activity.actualCost)}">
                          </div>
                        </div>
                      </section>
                      <section class="monte-edit-group">
                        <div class="monte-edit-group-head">
                          <h4>Leistungsstand</h4>
                          <p>Nur ergänzende Beobachtung, nicht führend für Leistungswert (EV), Kostenleistungsindex (CPI) oder Terminleistungsindex (SPI).</p>
                        </div>
                        <div class="monte-edit-grid eva-progress-grid">
                          <div class="form-field">
                            <label for="eva_progress_${index}"><span>VISUELLER</span><span>BAUFORTSCHRITT IN %</span></label>
                            <div class="eva-stepper-field eva-activity-stepper">
                              <input class="compact-input integer-input" id="eva_progress_${index}" data-eva-activity-index="${index}" data-eva-activity-field="progressPercent" type="text" inputmode="numeric" value="${formatIntegerInput(activity.progressPercent)}">
                              <div class="eva-stepper-buttons">
                                <button type="button" class="eva-stepper-btn" data-action="eva-activity-step" data-field="progressPercent" data-step="1" data-direction="up" data-index="${index}" aria-label="Baufortschritt erhöhen">▲</button>
                                <button type="button" class="eva-stepper-btn" data-action="eva-activity-step" data-field="progressPercent" data-step="1" data-direction="down" data-index="${index}" aria-label="Baufortschritt verringern">▼</button>
                              </div>
                            </div>
                          </div>
                          <div class="form-field">
                            <label><span>ABGELEITETER FERTIG-</span><span>STELLUNGSGRAD (FGR) IN %</span></label>
                            <div class="form-note form-note-strong">${formatNumber(activity.fgrPercent)}%</div>
                          </div>
                        </div>
                      </section>
                    </div>
                    <div class="monte-edit-metrics eva-result-row">
                      <div><span class="metric-label"><span>Kostenabweichung</span><span>(CV)</span></span><strong>${formatCurrency(activity.cv)}</strong></div>
                      <div><span class="metric-label"><span>Terminabweichung</span><span>(SV)</span></span><strong>${formatCurrency(activity.sv)}</strong></div>
                      <div><span class="metric-label"><span>Kostenleistungsindex</span><span>(CPI)</span></span><strong>${activity.cpi.toFixed(2)}</strong></div>
                      <div><span class="metric-label"><span>Terminleistungsindex</span><span>(SPI)</span></span><strong>${activity.spi.toFixed(2)}</strong></div>
                    </div>
                  </div>
                </details>
              `).join("")}
            </div>
            <div style="margin-top:14px;">
              <button class="action-btn add-with-icon" type="button" data-action="add-eva-activity"><span class="add-icon" aria-hidden="true">+</span><span>Aktivität hinzufügen</span></button>
            </div>
          </details>
          <div class="card-grid">
            <details class="info-card wide eva-knowledge-card eva-fold-card eva-fold-knowledge">
              <summary class="eva-fold-summary">
                <div class="eva-fold-main">
                  <span class="eva-fold-toggle" aria-hidden="true"></span>
                  <div>
                  <h3 class="eva-knowledge-title">EVA-Kennzahlen und Leseschlüssel</h3>
                  <p>Aktuelle Werte, Formeln und Leselogik des Kennzahlensatzes.</p>
                  </div>
                </div>
                <div class="eva-fold-meta">
                  <span class="badge">${evaIndicatorCards.length} Kennzahlen</span>
                </div>
              </summary>
              <div class="analysis-grid eva-indicator-grid">
                ${evaIndicatorCards.map((card) => `
                  <div class="mini-card eva-indicator-card">
                    <span class="eva-indicator-code">${card.code}</span>
                    <p>${card.value}</p>
                    <strong>${card.title}</strong>
                    <small>${card.formula}</small>
                    <small>${card.hint}</small>
                  </div>
                `).join("")}
              </div>
            </details>
          </div>
        </div>
      `;
    }
  },
  riskRegister: {
    key: "riskRegister",
    label: "Risikoregister",
    subtitle: "Operative Risiken, Maßnahmen und Status",
    render(state) {
      const result = calculateRiskRegisterResult(state);
      const selectedRiskIds = state.ui?.transferSelections?.riskIds || [];
      const riskView = state.ui?.riskRegisterView || {};
      const riskUndoCount = Array.isArray(state.ui?.riskRegisterUndoStack) ? state.ui.riskRegisterUndoStack.length : 0;
      const riskRedoCount = Array.isArray(state.ui?.riskRegisterRedoStack) ? state.ui.riskRegisterRedoStack.length : 0;
      const classifyRiskTone = (risk) => {
        const todayTs = parseRiskDateValue(new Date().toISOString().slice(0, 10));
        const dueTs = parseRiskDateValue(risk.dueDate);
        const isOverdue = Boolean(dueTs !== null && todayTs !== null && dueTs < todayTs && String(risk.status || "").toLowerCase() !== "geschlossen");
        if (isOverdue || risk.qualitativeRiskValue >= 13) return { key: "critical", label: "Kritisch" };
        if (risk.qualitativeRiskValue >= 9) return { key: "warn", label: "Erhöht" };
        return { key: "neutral", label: "Beobachten" };
      };
      const classifyStatusTone = (status) => {
        const value = normalizeRiskStatusValue(status);
        if (value === "geschlossen") return { key: "closed", label: "Geschlossen" };
        if (value === "überwachung") return { key: "watch", label: "Überwachung" };
        if (value === "in bearbeitung") return { key: "progress", label: "In Bearbeitung" };
        if (value === "in bewertung") return { key: "action", label: "In Bewertung" };
        return { key: "open", label: "Offen" };
      };
      const statusFilterOptions = [
        { value: "alle", label: "Alle" },
        { value: "offen", label: "Offen" },
        { value: "in bewertung", label: "In Bewertung" },
        { value: "in bearbeitung", label: "In Bearbeitung" },
        { value: "überwachung", label: "Überwachung" },
        { value: "geschlossen", label: "Geschlossen" }
      ];
      const statusOptions = statusFilterOptions.filter((status) => status.value !== "alle");
      const visibleColumns = state.ui?.riskRegisterView?.visibleColumns || ["select", "priority", "status", "value", "category", "phase", "impact", "owner", "dueDate", "measures"];
      const matrixSelection = riskView.matrixSelection || null;
      const columnDefs = [
        { key: "select", label: "Markieren" },
        { key: "priority", label: "Priorität", render: (risk) => classifyRiskTone(risk).label },
        { key: "status", label: "Status", render: (risk) => {
          const value = normalizeRiskStatusValue(risk.status);
          const match = statusOptions.find((status) => status.value === value);
          return match ? match.label : "Offen";
        } },
        { key: "value", label: "Risikowert", render: (risk) => formatCurrency(risk.expectedDamage), numeric: true },
        { key: "category", label: "Kategorie", render: (risk) => risk.category || "—" },
        { key: "phase", label: "Projektphase / Zuordnung", render: (risk) => risk.phase || "—" },
        { key: "impact", label: "Auswirkung", render: (risk) => `${Math.round(Number(risk.probabilityPercent) || 0)}% / ${formatCurrency(risk.financialImpact)}`, numeric: true },
        { key: "owner", label: "Verantwortlichkeit", render: (risk) => risk.owner || "nicht zugewiesen" },
        { key: "dueDate", label: "Ziel-Termin", render: (risk) => (risk.dueDate ? formatDate(risk.dueDate) : "—") },
        { key: "measures", label: "Maßnahmenplanung", render: (risk) => {
          const measures = String(risk.measures || risk.description || "").trim();
          return measures ? (measures.length > 112 ? `${measures.slice(0, 109)}...` : measures) : "Keine Maßnahmen gepflegt";
        } }
      ];
      const tableColumns = [
        ...(visibleColumns.includes("select") ? [columnDefs[0]] : []),
        { key: "id", label: "Risiko-ID" },
        { key: "title", label: "Risiko" },
        ...columnDefs.filter((column) => column.key !== "select" && visibleColumns.includes(column.key))
      ];
      const tableColumnWidths = {
        select: 84,
        id: 94,
        title: 248,
        category: 122,
        phase: 170,
        impact: 138,
        value: 112,
        priority: 116,
        status: 132,
        owner: 168,
        dueDate: 118,
        measures: 238
      };
      const sortBy = ["priority", "value", "id", "dueDate", "category"].includes(riskView.sortBy) ? riskView.sortBy : "priority";
      const matrixCounts = Array.from({ length: 5 }, () => Array(5).fill(0));
      const hasRiskFilters = Boolean(
        String(riskView.search || "").trim() ||
        String(riskView.status || "alle").trim().toLowerCase() !== "alle" ||
        String(riskView.owner || "alle").trim().toLowerCase() !== "alle" ||
        String(riskView.category || "alle").trim().toLowerCase() !== "alle" ||
        riskView.criticalOnly === true ||
        String(riskView.dueFrom || "").trim() ||
        String(riskView.dueTo || "").trim() ||
        matrixSelection
      );
      const visibleRiskRows = hasRiskFilters ? result.filteredRisks : result.risks;
      visibleRiskRows.forEach((risk) => {
        const likelihood = Math.max(1, Math.min(5, Number(risk.likelihood) || 1));
        const impact = Math.max(1, Math.min(5, Number(risk.impact) || 1));
        matrixCounts[5 - impact][likelihood - 1] += 1;
      });
      const qualitativeSummary = {
        critical: visibleRiskRows.filter((risk) => classifyRiskTone(risk).key === "critical").length,
        warn: visibleRiskRows.filter((risk) => classifyRiskTone(risk).key === "warn").length,
        neutral: visibleRiskRows.filter((risk) => classifyRiskTone(risk).key === "neutral").length
      };
      const selectedMatrixRisks = matrixSelection
        ? visibleRiskRows.filter((risk) => Number(risk.likelihood) === matrixSelection.likelihood && Number(risk.impact) === matrixSelection.impact)
        : [];
      const topLimit = [0, 5, 10, 20].includes(Number(riskView.topLimit)) ? Number(riskView.topLimit) : 5;
      const riskTableWidth = tableColumns.reduce((sum, column) => sum + (tableColumnWidths[column.key] || 132), 0);
      const filteredRiskValueTotal = visibleRiskRows.reduce((sum, risk) => sum + (Number(risk.expectedDamage) || 0), 0);
      const sortedRisks = [...visibleRiskRows].sort((a, b) => {
        if (sortBy === "value") {
          return b.expectedDamage - a.expectedDamage || b.qualitativeRiskValue - a.qualitativeRiskValue;
        }
        if (sortBy === "id") {
          return String(a.id || "").localeCompare(String(b.id || ""), "de");
        }
        if (sortBy === "category") {
          return String(a.category || "").localeCompare(String(b.category || ""), "de") || b.qualitativeRiskValue - a.qualitativeRiskValue;
        }
        if (sortBy === "dueDate") {
          const aDueTs = parseRiskDateValue(a.dueDate);
          const bDueTs = parseRiskDateValue(b.dueDate);
          return (aDueTs ?? Number.POSITIVE_INFINITY) - (bDueTs ?? Number.POSITIVE_INFINITY) || b.qualitativeRiskValue - a.qualitativeRiskValue;
        }
        const todayTs = parseRiskDateValue(new Date().toISOString().slice(0, 10));
        const aDueTs = parseRiskDateValue(a.dueDate);
        const bDueTs = parseRiskDateValue(b.dueDate);
        const aOverdue = Number(Boolean(aDueTs !== null && todayTs !== null && aDueTs < todayTs && String(a.status || "").toLowerCase() !== "geschlossen"));
        const bOverdue = Number(Boolean(bDueTs !== null && todayTs !== null && bDueTs < todayTs && String(b.status || "").toLowerCase() !== "geschlossen"));
        return bOverdue - aOverdue || b.qualitativeRiskValue - a.qualitativeRiskValue || b.expectedDamage - a.expectedDamage;
      });
      const topRisks = topLimit === 0 ? sortedRisks : sortedRisks.slice(0, topLimit);
      const topMetricMax = Math.max(...topRisks.map((risk) => sortBy === "value" ? risk.expectedDamage : risk.qualitativeRiskValue), 1);
      return `
        <div class="module-shell risk-register-shell">
          <div class="module-title risk-register-title">
            <div>
              <h2>Risikoregister</h2>
              <p>Im Risikoregister bündelst du Risiken, Verantwortlichkeiten und Maßnahmen in einer operativen Sicht. Es ergänzt Monte Carlo, ersetzt es aber nicht, und bleibt die führende Pflegeebene für dein tägliches Risikomanagement.</p>
            </div>
            <span class="badge">Operative Risiken</span>
          </div>
            <div class="kpi-grid risk-register-kpi-grid" id="risk-register-summary">
            <article class="kpi-card"><div class="kpi-label">Erwarteter Schaden</div><div class="kpi-value">${formatCurrency(result.totalExpectedDamage)}</div><div class="kpi-sub">Summe aller erwarteten Wirkungen</div></article>
            <article class="kpi-card gold"><div class="kpi-label">Aktive Risiken</div><div class="kpi-value">${result.activeCount}</div><div class="kpi-sub">offen, in Bewertung, in Bearbeitung oder Überwachung</div></article>
            <article class="kpi-card green"><div class="kpi-label">Kritische Risiken</div><div class="kpi-value">${result.criticalCount}</div><div class="kpi-sub">hohe Priorität für Steuerung</div></article>
            <article class="kpi-card blue time-critical"><div class="kpi-label">Überfällige Risiken</div><div class="kpi-value">${result.overdueCount}</div><div class="kpi-sub">Termin bereits überschritten</div></article>
          </div>
          <details class="info-card risk-register-card risk-fold-card" open>
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>Registerlogik und operative Sicht</strong>
                    <span>${result.risks.length} Risiken · ${visibleRiskRows.length} in aktueller Sicht · ${qualitativeSummary.critical} kritisch</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">Überblick</span>
              </div>
            </summary>
            <div class="risk-register-overview-grid">
              <div class="risk-register-overview-stack">
                <section class="info-card risk-register-card">
                  <h3>Registerlogik</h3>
                  <p class="form-note" style="margin-bottom:10px;">Das Register ist deine führende Pflegeebene für einzelne Risiken. Es hält Risiko, Bewertung, Maßnahmen und Verantwortung zusammen und liefert dir die Basis für Priorisierung, Reporting und Übergabe an Monte Carlo.</p>
                  <ul>
                    <li>Jedes Risiko führst du mit ID, Status und Verantwortlichem.</li>
                    <li>5x5-Score, Priorität und erwarteter Schaden bleiben für dich live sichtbar.</li>
                    <li>Bei Bedarf übergibst du Registerrisiken an Monte Carlo.</li>
                  </ul>
                </section>
                <section class="info-card risk-register-card">
                  <h3>Qualitative Einordnung</h3>
                  <div class="risk-tone-summary">
                    <div class="risk-tone-line"><strong>Kritisch</strong><span>ab 13 / 25 oder überfällig</span></div>
                    <div class="risk-tone-line"><strong>Erhöht</strong><span>9 bis 12 / 25</span></div>
                    <div class="risk-tone-line"><strong>Beobachten</strong><span>unter 9 / 25</span></div>
                    <div class="risk-tone-line"><strong>Aktuell sichtbar</strong><span>${visibleRiskRows.length} Risiken · ${qualitativeSummary.critical} kritisch</span></div>
                  </div>
                </section>
              </div>
            <section class="info-card risk-register-card risk-register-overview-focus" id="risk-register-operational-focus">
                <div class="risk-operational-head">
                  <span class="badge">Operativer Fokus</span>
                  <h3>Operative Steuerung und Monte-Carlo-Übernahme</h3>
                </div>
                <p class="form-note" style="margin-top:0;">Hier siehst du die aktuellen Steuerungsgrößen und entscheidest, welche Risiken in Monte Carlo übernommen werden.</p>
                <ul>
                  <li><strong>Kritisch:</strong> ${result.criticalCount} Risiken</li>
                  <li><strong>Überfällig:</strong> ${result.overdueCount} Risiken</li>
                  <li><strong>Geschlossen:</strong> ${result.closedCount} Risiken</li>
                  <li><strong>Erwarteter Schaden:</strong> ${formatCurrency(result.totalExpectedDamage)}</li>
                </ul>
                <p class="form-note" style="margin-top:14px;">Du wählst Risiken in der Tabelle <strong>Alle Risiken</strong> über die Spalte <strong>Markieren</strong> aus. Die Auswahl steuert, welche Risiken in die Monte-Carlo-Simulation übernommen werden.</p>
                <button class="action-btn" type="button" data-action="transfer-all-risks-to-mc">Alle Risiken in Monte-Carlo-Simulation übernehmen</button>
                <p class="form-note" style="margin-top:10px;">Mit diesem Button übernimmst du alle Registerrisiken als Risikoereignisse in Monte Carlo.</p>
                <button class="action-btn" type="button" data-action="transfer-selected-risks-to-mc" style="margin-top:10px;">Markierte Risiken in Monte-Carlo-Simulation übernehmen</button>
                <p class="form-note" style="margin-top:10px;">Mit diesem Button übernimmst du nur die aktuell markierten Risiken aus der Tabelle. Markiert: ${selectedRiskIds.length} Risiken.</p>
              </section>
            </div>
          </details>
          <details class="info-card risk-fold-card ai-workshop-card" open style="margin-top:18px;">
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>KI für Risiken und Maßnahmen</strong>
                    <span>Freitext auswerten · Risiken strukturieren · Maßnahmen ableiten</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">KI-Startstufe</span>
              </div>
            </summary>
            <div class="risk-fold-body">
            <p class="form-note">Hier hilfst du dir bei der Risikoerfassung und bei der Nacharbeit bestehender Risiken.<br>Freitext wird strukturiert, Maßnahmen werden vorgeschlagen und das Rest-Risiko wird fachlich verdichtet.</p>
            <div class="report-export-grid" style="align-items:start;">
            <div class="form-field wide">
                <label for="aiWorkshopFreeText">Freitext für Risiko- und Maßnahmenanalyse</label>
                <textarea id="aiWorkshopFreeText" data-ai-workshop-field="freeText" placeholder="Beschreibe das Risiko, die Maßnahme oder den Sachverhalt in Ruhe und klar. Je ausführlicher du formulierst, desto besser kann die KI ein Ergebnis erstellen.">${escapeHtml(state.ui?.aiWorkshop?.freeText || "")}</textarea>
              </div>
            </div>
            <div class="ai-workshop-grid">
              <button class="action-btn ${state.ui?.aiWorkshop?.activeTask === "free-text-risks" ? "active" : ""}" type="button" data-action="run-ai-workshop" data-ai-workshop-task="free-text-risks">Freitext in Risiken übersetzen</button>
              <button class="action-btn ${state.ui?.aiWorkshop?.activeTask === "measures-residual" ? "active" : ""}" type="button" data-action="run-ai-workshop" data-ai-workshop-task="measures-residual">Maßnahmenvorschläge mit Rest-Risiko erzeugen</button>
            </div>
            <p class="form-note">Die KI wertet zuerst den Freitext aus und zeigt dir einen Vorschlag.<br>Der Übernehmen-Button passt sich daran an.</p>
            <div class="code-box ai-workshop-output ${state.ui?.aiWorkshop?.busy ? "is-loading" : ""} ${state.ui?.aiWorkshop?.resultTone === "success" ? "tone-success" : state.ui?.aiWorkshop?.resultTone === "danger" ? "tone-danger" : "tone-neutral"}">
              <strong>${escapeHtml(state.ui?.aiWorkshop?.resultTitle || "KI bereit")}</strong>
              ${(() => {
                const resultData = state.ui?.aiWorkshop?.resultData;
                if (resultData?.resultType === "risk-suggestion" && Array.isArray(resultData.items) && resultData.items.length) {
                  return `
                    <div class="ai-workshop-result-list">
                      ${resultData.items.map((item, index) => `
                        <article class="ai-workshop-result-card">
                          <div class="ai-workshop-result-card-head">
                            <strong>${escapeHtml(item.title || `Vorschlag ${index + 1}`)}</strong>
                            <span class="badge">Risikovorschlag</span>
                          </div>
                          <div class="ai-workshop-result-card-grid">
                            <div class="ai-workshop-result-card-column">
                              <div><span>Kategorie</span><strong>${escapeHtml(item.category || "—")}</strong></div>
                              <div><span>Phase</span><strong>${escapeHtml(item.phase || "—")}</strong></div>
                              <div><span>Verantwortlich</span><strong>${escapeHtml(item.owner || "—")}</strong></div>
                              <div><span>Eintrittswahrscheinlichkeit (1-5)</span><strong>${escapeHtml(String(deriveRiskLikelihoodFromPercent(item.probabilityPercent, item.likelihood) ?? "—"))}</strong></div>
                            </div>
                            <div class="ai-workshop-result-card-column">
                              <div><span>Schaden in Euro</span><strong>${formatCurrency(Number(item.financialImpact) || 0)}</strong></div>
                              <div><span>Eintrittswahrscheinlichkeit in %</span><strong>${escapeHtml(String(item.probabilityPercent ?? "—"))}%</strong></div>
                              <div><span>Erwarteter Schaden</span><strong>${formatCurrency((Number(item.financialImpact) || 0) * (Math.max(0, Math.min(100, Number(item.probabilityPercent) || 0)) / 100))}</strong></div>
                              <div><span>Auswirkung</span><strong>${escapeHtml(String(item.impact ?? "—"))}</strong></div>
                            </div>
                          </div>
                          <div class="ai-workshop-result-description">
                            <span>Beschreibung</span>
                            <p>${escapeHtml(item.description || "—")}</p>
                          </div>
                          <div class="ai-workshop-result-foot">
                            <div class="ai-workshop-result-foot-head">Maßnahmen und Rest-Risiko</div>
                            <div class="ai-workshop-result-foot-grid">
                              <div class="ai-workshop-result-foot-item">
                                <span>Maßnahmen</span>
                                <p>${escapeHtml(item.measures || "—")}</p>
                              </div>
                              <div class="ai-workshop-result-foot-item">
                                <span>Qualitatives Rest-Risiko</span>
                                <p>${escapeHtml(item.residualRisk || "—")}</p>
                              </div>
                            </div>
                          </div>
                        </article>
                      `).join("")}
                    </div>
                  `;
                }
                if (resultData?.resultType === "risk-measures" && Array.isArray(resultData.items) && resultData.items.length) {
                  return `
                    <div class="ai-workshop-result-list">
                      ${resultData.items.map((item, index) => `
                        <article class="ai-workshop-result-card">
                          <div class="ai-workshop-result-card-head">
                            <strong>${escapeHtml(item.riskId || `Ziel ${index + 1}`)}</strong>
                            <span class="badge">Maßnahmenentwurf</span>
                          </div>
                          <div class="ai-workshop-result-card-grid">
                            <div><span>Verantwortlich</span><strong>${escapeHtml(item.owner || "—")}</strong></div>
                            <div><span>Fällig am</span><strong>${escapeHtml(item.dueDate || "—")}</strong></div>
                          </div>
                          <div class="ai-workshop-result-foot">
                            <div class="ai-workshop-result-foot-head">Maßnahmen und Rest-Risiko</div>
                            <div class="ai-workshop-result-foot-grid">
                              <div class="ai-workshop-result-foot-item">
                                <span>Maßnahmen</span>
                                <p>${escapeHtml(item.measures || "—")}</p>
                              </div>
                              <div class="ai-workshop-result-foot-item">
                                <span>Qualitatives Rest-Risiko</span>
                                <p>${escapeHtml(item.residualRisk || "—")}</p>
                              </div>
                            </div>
                          </div>
                        </article>
                      `).join("")}
                    </div>
                  `;
                }
                return `<div style="margin-top:8px; white-space:pre-wrap;">${escapeHtml(state.ui?.aiWorkshop?.resultText || "Wähle eine Funktion für die KI-Startstufe.")}</div>`;
              })()}
            </div>
            <div class="form-field" style="margin-top:12px;">
              <label for="aiWorkshopSelectedRisk">Übernahmeziel</label>
              <select id="aiWorkshopSelectedRisk" data-ai-workshop-field="selectedRiskId">
                <option value="">Neues Risiko anlegen</option>
                ${result.risks.map((risk) => `
                  <option value="${escapeHtml(risk.id)}" ${String(state.ui?.aiWorkshop?.selectedRiskId || "") === String(risk.id) ? "selected" : ""}>${escapeHtml(risk.id)} · ${escapeHtml(risk.title)}</option>
                `).join("")}
              </select>
            </div>
            <p class="form-note">Wählst du ein bestehendes Risiko, wird der Vorschlag dort übernommen. Ohne Auswahl legt die KI ein neues Risiko an.</p>
            <div class="ai-workshop-result-actions" style="margin-top:12px;">
              <button class="action-btn primary" type="button" data-action="apply-ai-workshop-result" ${state.ui?.aiWorkshop?.activeTask === "measures-residual" && !String(state.ui?.aiWorkshop?.selectedRiskId || "").trim() ? "disabled" : ""}>${state.ui?.aiWorkshop?.activeTask === "measures-residual" ? (String(state.ui?.aiWorkshop?.selectedRiskId || "").trim() ? "Bestehendes Risiko aktualisieren" : "Zielrisiko wählen") : (String(state.ui?.aiWorkshop?.selectedRiskId || "").trim() ? "Bestehendes Risiko aktualisieren" : "Neues Risiko anlegen")}</button>
              <button class="action-btn" type="button" data-action="clear-ai-workshop-result">Vorschlag verwerfen</button>
            </div>
            <p class="form-note" style="margin-top:8px;">Mit dem Übernehmen-Button speicherst du den geprüften Vorschlag im Risikoregister.</p>
            </div>
          </details>
          <details class="info-card risk-register-card risk-fold-card risk-fold-edit" open>
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>Risiken erfassen und bearbeiten</strong>
                    <span>Einträge pflegen · Werte, Status und Maßnahmen ändern</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">Bearbeiten</span>
              </div>
            </summary>
            <div class="risk-fold-body">
              <div class="risk-register-history-actions" style="margin-bottom:14px;">
                <button class="action-btn risk-add-btn" type="button" data-action="add-risk-register-item"><span class="risk-add-icon" aria-hidden="true">+</span><span>Risiko hinzufügen (manuell ohne KI-Unterstützung)</span></button>
                <button class="action-btn" type="button" data-action="undo-risk-register-change" ${riskUndoCount > 0 ? "" : "disabled"}><span class="risk-history-icon" aria-hidden="true">◀</span><span>Letzten Schritt rückgängig</span></button>
                <button class="action-btn" type="button" data-action="redo-risk-register-change" ${riskRedoCount > 0 ? "" : "disabled"}><span class="risk-history-icon" aria-hidden="true">▶</span><span>Letzten Schritt wiederherstellen</span></button>
              </div>
              <div class="risk-register-toolbar" style="margin-bottom:12px;">
                <div class="risk-register-toolbar-grid" style="grid-template-columns:minmax(360px, 1fr);">
                  <div class="form-field">
                    <label>Sortierung</label>
                    <div class="risk-sort-actions" style="margin-top:8px;display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:10px;">
                      <button class="risk-chart-chip ${String(result.ui?.riskRegisterView?.editSortBy || "newest") === "newest" ? "active" : ""}" type="button" data-action="set-risk-register-edit-sort" data-sort-by="newest" style="width:100%;">Neueste zuerst</button>
                      <button class="risk-chart-chip ${String(result.ui?.riskRegisterView?.editSortBy || "") === "id" ? "active" : ""}" type="button" data-action="set-risk-register-edit-sort" data-sort-by="id" style="width:100%;">Risiko-ID aufsteigend</button>
                    </div>
                  </div>
                </div>
              </div>
              <div class="risk-edit-list">
              ${result.editRisks.map((risk, index) => `
                <details class="risk-edit-card risk-tone-${classifyRiskTone(risk).key} risk-fold-card risk-edit-fold" ${index === 0 ? "open" : ""}>
                  <summary class="risk-edit-summary">
                    <div class="risk-edit-summary-main">
                      <span class="risk-fold-toggle" aria-hidden="true"></span>
                      <div class="risk-edit-summary-text">
                        <span class="risk-edit-id">${risk.id}</span>
                        <strong>${risk.title}</strong>
                        <span class="risk-edit-summary-value">${risk.qualitativeRiskValue} / 25 · ${formatCurrency(risk.expectedDamage)}</span>
                      </div>
                    </div>
                    <span class="risk-tone-pill risk-tone-${classifyRiskTone(risk).key}">${classifyRiskTone(risk).label}</span>
                  </summary>
                  <div class="risk-edit-grid">
                    <div class="form-field">
                      <label for="risk_id_${index}">Risiko-ID</label>
                      <input id="risk_id_${index}" data-risk-index="${index}" data-risk-field="id" type="text" value="${risk.id}">
                    </div>
                    <div class="form-field half">
                      <label for="risk_title_${index}">Risikobezeichnung</label>
                      <input id="risk_title_${index}" data-risk-index="${index}" data-risk-field="title" type="text" value="${risk.title}">
                    </div>
                    <div class="form-field">
                      <label for="risk_phase_${index}">Projektphase / Zuordnung</label>
                      <select id="risk_phase_${index}" data-risk-index="${index}" data-risk-field="phase">
                        ${buildRiskPhaseOptions(risk.phase).map((phase) => `
                          <option value="${escapeHtml(phase)}" ${String(risk.phase || "").trim().toLowerCase() === String(phase).trim().toLowerCase() ? "selected" : ""}>${escapeHtml(phase)}</option>
                        `).join("")}
                      </select>
                    </div>
                    <div class="form-field wide">
                      <label for="risk_description_${index}">Risikobeschreibung</label>
                      <textarea class="risk-description-textarea" id="risk_description_${index}" data-risk-index="${index}" data-risk-field="description" rows="2">${risk.description}</textarea>
                    </div>
                    <div class="form-field">
                      <label for="risk_category_${index}">Kategorie</label>
                      <select id="risk_category_${index}" data-risk-index="${index}" data-risk-field="category">
                        ${buildRiskCategoryOptions(risk.category).map((category) => `
                          <option value="${escapeHtml(category)}" ${String(risk.category || "").trim().toLowerCase() === String(category).trim().toLowerCase() ? "selected" : ""}>${escapeHtml(category)}</option>
                        `).join("")}
                      </select>
                    </div>
                    <div class="form-field">
                      <label for="risk_owner_${index}">Verantwortlich</label>
                      <input id="risk_owner_${index}" data-risk-index="${index}" data-risk-field="owner" type="text" list="risk_owner_options_${index}" value="${risk.owner}">
                      <datalist id="risk_owner_options_${index}">
                        ${result.ownerOptions.map((owner) => `
                          <option value="${escapeHtml(owner)}"></option>
                        `).join("")}
                      </datalist>
                    </div>
                    <div class="form-field">
                      <label for="risk_status_${index}">Status</label>
                      <select id="risk_status_${index}" data-risk-index="${index}" data-risk-field="status">
                        ${statusOptions.map((status) => `
                          <option value="${status.value}" ${normalizeRiskStatusValue(risk.status) === status.value ? "selected" : ""}>${status.label}</option>
                        `).join("")}
                      </select>
                    </div>
                    <div class="form-field">
                      <label for="risk_due_${index}">Fällig am</label>
                      <input id="risk_due_${index}" data-risk-index="${index}" data-risk-field="dueDate" type="date" value="${risk.dueDate}">
                    </div>
                    <div class="risk-input-grid wide">
                      <div class="form-field">
                        <label for="risk_financial_${index}">Schaden in Euro</label>
                        <div class="eva-stepper-field risk-stepper-field">
                          <input class="compact-input currency-input" id="risk_financial_${index}" data-risk-index="${index}" data-risk-field="financialImpact" type="text" inputmode="numeric" value="${formatCurrencyInput(risk.financialImpact)}">
                          <div class="eva-stepper-buttons">
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="financialImpact" data-step="1000" data-direction="up" data-index="${index}" aria-label="Schaden erhöhen">▲</button>
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="financialImpact" data-step="1000" data-direction="down" data-index="${index}" aria-label="Schaden verringern">▼</button>
                          </div>
                        </div>
                      </div>
                      <div class="form-field">
                        <label for="risk_probability_${index}">Eintrittswahrscheinlichkeit in %</label>
                        <div class="eva-stepper-field risk-stepper-field">
                          <input class="compact-input integer-input" id="risk_probability_${index}" data-risk-index="${index}" data-risk-field="probabilityPercent" type="text" inputmode="numeric" value="${formatIntegerInput(risk.probabilityPercent)}">
                          <div class="eva-stepper-buttons">
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="probabilityPercent" data-step="1" data-direction="up" data-index="${index}" aria-label="Eintrittswahrscheinlichkeit erhöhen">▲</button>
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="probabilityPercent" data-step="1" data-direction="down" data-index="${index}" aria-label="Eintrittswahrscheinlichkeit verringern">▼</button>
                          </div>
                        </div>
                      </div>
                      <div class="form-field">
                        <label>Erwarteter Schaden</label>
                        <div class="risk-value-box risk-value-large">${formatCurrency(risk.expectedDamage)}</div>
                      </div>
                      <div class="risk-grid-spacer" aria-hidden="true"></div>
                    </div>
                    <div class="risk-output-grid wide">
                      <div class="form-field">
                        <label for="risk_likelihood_${index}">Eintrittswahrscheinlichkeit (1-5)</label>
                        <div class="risk-value-box risk-value-large">${deriveRiskLikelihoodFromPercent(risk.probabilityPercent, risk.likelihood)}</div>
                      </div>
                      <div class="form-field">
                        <label for="risk_impact_${index}">Auswirkung (1-5)</label>
                        <div class="eva-stepper-field risk-stepper-field">
                          <input class="compact-input integer-input" id="risk_impact_${index}" data-risk-index="${index}" data-risk-field="impact" type="text" inputmode="numeric" value="${formatIntegerInput(risk.impact)}">
                          <div class="eva-stepper-buttons">
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="impact" data-step="1" data-direction="up" data-index="${index}" aria-label="Auswirkung erhöhen">▲</button>
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="impact" data-step="1" data-direction="down" data-index="${index}" aria-label="Auswirkung verringern">▼</button>
                          </div>
                        </div>
                      </div>
                      <div class="form-field">
                        <label>Risikowert 5×5</label>
                        <div class="risk-value-box risk-value-large">${risk.qualitativeRiskValue} / 25</div>
                      </div>
                      <div class="risk-grid-spacer" aria-hidden="true"></div>
                    </div>
                    <div class="form-field wide">
                      <label for="risk_measures_${index}">Maßnahmenplanung</label>
                      <textarea id="risk_measures_${index}" data-risk-index="${index}" data-risk-field="measures">${risk.measures}</textarea>
                    </div>
                    <div class="risk-residual-stack">
                      <div class="form-field">
                        <label for="risk_residual_${index}">Rest-Risiko</label>
                        <textarea id="risk_residual_${index}" data-risk-index="${index}" data-risk-field="residualRisk" rows="3">${risk.residualRisk}</textarea>
                      </div>
                      <div class="risk-residual-actions">
                        <button class="risk-remove-btn" type="button" data-action="remove-risk-register-item" data-index="${index}">Risiko entfernen</button>
                      </div>
                    </div>
                  </div>
                </details>
              `).join("")}
            </div>
            </div>
          </details>
          <details class="info-card risk-register-card risk-fold-card risk-fold-table" open>
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>Alle Risiken</strong>
                    <span>Sortierte Liste · zuerst überfällige Risiken, dann 5x5-Wert und erwarteter Schaden</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">Sortierte Liste</span>
              </div>
            </summary>
            <div class="risk-fold-body">
              <div class="risk-register-toolbar">
                <div class="risk-register-toolbar-grid">
                  <div class="form-field">
                    <label for="risk_view_search">Suche</label>
                    <input id="risk_view_search" data-risk-ui-field="search" type="text" value="${riskView.search || ""}" placeholder="Risiko, Owner, Beschreibung">
                  </div>
                  <div class="form-field">
                    <label for="risk_view_sort">Sortierung</label>
                    <div class="risk-select-shell">
                      <select id="risk_view_sort" data-risk-ui-field="sortBy">
                        <option value="priority" ${sortBy === "priority" ? "selected" : ""}>Priorität</option>
                        <option value="value" ${sortBy === "value" ? "selected" : ""}>Risikowert</option>
                        <option value="id" ${sortBy === "id" ? "selected" : ""}>Risiko-ID</option>
                        <option value="category" ${sortBy === "category" ? "selected" : ""}>Kategorie</option>
                        <option value="dueDate" ${sortBy === "dueDate" ? "selected" : ""}>Ziel-Termin</option>
                      </select>
                    </div>
                  </div>
                  <div class="form-field">
                    <label for="risk_view_category">Kategorie</label>
                    <div class="risk-select-shell">
                      <select id="risk_view_category" data-risk-ui-field="category">
                        <option value="alle" ${String(riskView.category || "alle").toLowerCase() === "alle" ? "selected" : ""}>Alle</option>
                        ${result.categoryOptions.map((category) => `
                          <option value="${category}" ${String(riskView.category || "").toLowerCase() === category.toLowerCase() ? "selected" : ""}>${category}</option>
                        `).join("")}
                      </select>
                    </div>
                  </div>
                  <div class="form-field">
                    <label for="risk_view_status">Status</label>
                    <div class="risk-select-shell">
                      <select id="risk_view_status" data-risk-ui-field="status">
                        ${statusFilterOptions.map((status) => `
                          <option value="${status.value}" ${String(riskView.status || "alle").toLowerCase() === status.value ? "selected" : ""}>${status.label}</option>
                        `).join("")}
                      </select>
                    </div>
                  </div>
                  <div class="form-field">
                    <label for="risk_view_due_from">Ziel-Termin von</label>
                    <input id="risk_view_due_from" data-risk-ui-field="dueFrom" type="date" value="${riskView.dueFrom || ""}">
                  </div>
                  <div class="form-field">
                    <label for="risk_view_due_to">Ziel-Termin bis</label>
                    <input id="risk_view_due_to" data-risk-ui-field="dueTo" type="date" value="${riskView.dueTo || ""}">
                  </div>
                </div>
                <div class="risk-register-toolbar-actions">
                  <button class="action-btn" type="button" data-action="reset-risk-register-filters">Filter zurücksetzen</button>
                </div>
              </div>
              <p class="risk-filter-summary form-note">Aktuell ${visibleRiskRows.length} Risiko${visibleRiskRows.length === 1 ? "" : "e"} in Sicht${riskView.dueFrom || riskView.dueTo ? ` · Zeitraum ${riskView.dueFrom ? formatDate(riskView.dueFrom) : "offen"} bis ${riskView.dueTo ? formatDate(riskView.dueTo) : "offen"}` : ""}.</p>
              <div class="risk-column-panel">
                <p class="risk-column-note">Die Spalten der Tabelle lassen sich per Klick ein- oder ausblenden. Ausgeblendete Spalten werden rot markiert und durchgestrichen.</p>
                <div class="risk-column-strip risk-column-strip-inline">
                  <button class="risk-column-pill risk-column-pill-all" type="button" data-action="show-all-risk-register-columns">
                    Alle Spalten einblenden
                  </button>
                  ${columnDefs.map((column) => `
                    <button class="risk-column-pill ${visibleColumns.includes(column.key) ? "active" : ""}" type="button" data-action="toggle-risk-register-column" data-column="${column.key}">
                      ${column.key === "phase" ? "Projektphase" : column.label}
                    </button>
                  `).join("")}
                </div>
                <p class="risk-selection-note">Risiken per Checkbox markieren. Markierte Risiken können in die Monte-Carlo-Simulation übernommen werden. Markiert: ${selectedRiskIds.length} Risiken.</p>
              </div>
              <div class="table-wrap risk-register-table-wrap">
                <table class="data-table risk-register-table" style="width:${riskTableWidth}px; min-width:${riskTableWidth}px; table-layout:fixed;">
                  <colgroup>
                    ${tableColumns.map((column) => `<col style="width:${tableColumnWidths[column.key] || 132}px;">`).join("")}
                  </colgroup>
                  <thead>
                    <tr>
                      ${tableColumns.map((column) => `<th class="${column.key === "select" ? "risk-select-head" : ""}">${column.label}</th>`).join("")}
                    </tr>
                  </thead>
                  <tbody>
                    ${sortedRisks.map((risk) => {
                      const riskTone = classifyRiskTone(risk);
                      const isSelected = selectedRiskIds.includes(risk.id);
                      return `
                        <tr>
                          ${tableColumns.map((column) => {
                            if (column.key === "select") {
                              return `
                                <td class="risk-select-cell">
                                  <label class="risk-select-check">
                                    <input type="checkbox" data-transfer-select="riskIds" data-transfer-value="${risk.id}" ${isSelected ? "checked" : ""}>
                                    <span class="risk-select-check-mark" aria-hidden="true"></span>
                                  </label>
                                </td>
                              `;
                            }
                            if (column.key === "id") {
                              return `<td><strong>${risk.id}</strong></td>`;
                            }
                            if (column.key === "title") {
                              return `
                                <td>
                                  <strong>${risk.title}</strong><br>
                                  <span class="risk-register-muted">${String(risk.measures || risk.description || "").trim() ? (String(risk.measures || risk.description || "").trim().length > 110 ? `${String(risk.measures || risk.description || "").trim().slice(0, 107)}...` : String(risk.measures || risk.description || "").trim()) : "Keine Maßnahmen gepflegt"}</span>
                                </td>
                              `;
                            }
                            const value = column.render(risk);
                            if (column.key === "priority") {
                              return `<td class="risk-priority-cell risk-priority-cell-${riskTone.key}"><span class="risk-tone-pill risk-tone-${riskTone.key}">${value}</span></td>`;
                            }
                            if (column.key === "status") {
                              const statusTone = classifyStatusTone(risk.status);
                              return `<td><span class="risk-status-pill risk-status-${statusTone.key}">${statusTone.label}</span></td>`;
                            }
                            return `<td class="${column.numeric ? "numeric-output" : ""}">${value}</td>`;
                          }).join("")}
                        </tr>
                      `;
                    }).join("")}
                  </tbody>
                  <tfoot>
                    <tr class="risk-register-total-row">
                      ${tableColumns.map((column) => {
                        if (column.key === "id") {
                          return `<td class="risk-register-total-label"><strong>Summe Risikowerte</strong></td>`;
                        }
                        if (column.key === "value") {
                          return `<td class="risk-register-total-value"><span class="risk-register-total-value-inner">${formatCurrency(filteredRiskValueTotal)}</span></td>`;
                        }
                        return `<td class="risk-register-total-spacer"></td>`;
                      }).join("")}
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          </details>
          <details class="info-card risk-register-card risk-fold-card risk-fold-chart" open>
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>Top-Risiken nach Risikowert und Priorität</strong>
                    <span>Aktuelle Sicht · Balken nach Priorität oder Risikowert</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">Top-Risiken</span>
              </div>
            </summary>
            <div class="risk-fold-body">
              <div class="risk-chart-toolbar">
                <div class="risk-chart-toggle">
                  <button class="risk-chart-chip ${sortBy === "value" ? "" : "active"}" type="button" data-action="set-risk-register-sort" data-sort-by="priority">Priorität</button>
                  <button class="risk-chart-chip ${sortBy === "value" ? "active" : ""}" type="button" data-action="set-risk-register-sort" data-sort-by="value">Risikowert</button>
                </div>
                <div class="risk-chart-toggle">
                  ${[[5, "Top 5"], [10, "Top 10"], [20, "Top 20"], [0, "Alle"]].map(([limit, label]) => `
                    <button class="risk-chart-chip ${topLimit === limit ? "active" : ""}" type="button" data-action="set-risk-register-limit" data-limit="${limit}">${label}</button>
                  `).join("")}
                </div>
              </div>
              <div class="risk-chart-list">
                ${topRisks.length ? topRisks.map((risk) => {
                  const riskTone = classifyRiskTone(risk);
                  const metric = sortBy === "value" ? risk.expectedDamage : risk.qualitativeRiskValue;
                  const width = Math.max(12, Math.round((metric / topMetricMax) * 100));
                  const metricLabel = sortBy === "value" ? formatCurrency(risk.expectedDamage) : `${risk.qualitativeRiskValue} von 25`;
                  return `
                    <div class="risk-chart-row risk-tone-${riskTone.key}">
                      <div class="risk-chart-identity">
                        <div class="risk-chart-title">${risk.id}</div>
                        <div class="risk-chart-subtitle">${risk.title}</div>
                        <div class="risk-chart-subtitle">${risk.category || "—"} · ${risk.phase || "—"}</div>
                      </div>
                      <div class="risk-chart-bar">
                        <div class="risk-chart-track">
                          <div class="risk-chart-fill risk-tone-${riskTone.key}" style="width:${width}%"></div>
                        </div>
                      </div>
                      <div class="risk-chart-value">
                        <div class="risk-chart-score">${metricLabel}</div>
                        <div class="risk-chart-meta">Verantwortlich: ${risk.owner || "nicht zugewiesen"}</div>
                      </div>
                      <div class="risk-chart-pill">
                        <span class="risk-tone-pill risk-tone-${riskTone.key}">${riskTone.label}</span>
                      </div>
                    </div>
                  `;
                }).join("") : `<p>Die aktuelle Filtersicht enthält keine Risiken.</p>`}
              </div>
            </div>
          </details>
          <details class="info-card risk-register-card risk-fold-card risk-fold-matrix" open>
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>Qualitative 5x5-Matrix</strong>
                    <span>Verteilung nach Wahrscheinlichkeit und Auswirkung</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">Qualitative Sicht</span>
              </div>
            </summary>
            <div class="risk-fold-body">
              <div class="risk-matrix-legend">
                <span class="risk-tone-pill risk-tone-neutral">Beobachten</span>
                <span class="risk-tone-pill risk-tone-warn">Erhöht</span>
                <span class="risk-tone-pill risk-tone-critical">Kritisch</span>
                <span class="risk-matrix-legend-note">Farblogik nach 5×5-Score und Terminlage</span>
              </div>
              <div class="risk-matrix-shell">
                <div class="risk-matrix-yaxis">Auswirkung</div>
                <div class="risk-matrix-grid">
                  ${[5, 4, 3, 2, 1].map((impact) => `
                    ${[1, 2, 3, 4, 5].map((likelihood) => {
                      const count = matrixCounts[5 - impact][likelihood - 1];
                      const score = likelihood * impact;
                      const tone = score >= 13 ? "critical" : score >= 9 ? "warn" : "neutral";
                      const isSelected = matrixSelection && matrixSelection.likelihood === likelihood && matrixSelection.impact === impact;
                      return `
                        <button class="risk-matrix-cell risk-matrix-button risk-tone-${tone} ${isSelected ? "is-active" : ""}" type="button" data-action="select-risk-register-matrix-cell" data-likelihood="${likelihood}" data-impact="${impact}">
                          <span class="risk-matrix-code">W ${likelihood} / A ${impact}</span>
                          <strong>${count}</strong>
                          <small>${score} Punkte</small>
                        </button>
                      `;
                    }).join("")}
                  `).join("")}
                </div>
              </div>
              <div class="risk-matrix-xaxis">Eintrittswahrscheinlichkeit</div>
              <div class="risk-matrix-selected">
                <div class="risk-matrix-selected-head">
                  <strong>${matrixSelection ? `Ausgewählte Matrixzelle W ${matrixSelection.likelihood} / A ${matrixSelection.impact}` : "Matrixauswahl"}</strong>
                  <span>${matrixSelection ? `${selectedMatrixRisks.length} Risiko${selectedMatrixRisks.length === 1 ? "" : "e"} gefunden` : "Klicke auf ein Feld, um die zugehörigen Risiken anzuzeigen."}</span>
                </div>
                ${matrixSelection ? (
                  selectedMatrixRisks.length
                    ? `<div class="risk-matrix-selected-list">
                        ${selectedMatrixRisks.map((risk) => {
                          const riskTone = classifyRiskTone(risk);
                          return `
                            <article class="risk-matrix-selected-card risk-tone-${riskTone.key}">
                              <div class="risk-matrix-selected-top">
                                <strong>${risk.id} · ${risk.title}</strong>
                                <span class="risk-tone-pill risk-tone-${riskTone.key}">${riskTone.label}</span>
                              </div>
                              <div class="risk-matrix-selected-meta">${risk.phase || "—"} · ${risk.owner || "nicht zugewiesen"}</div>
                              <div class="risk-matrix-selected-body">
                                <span><strong>W:</strong> ${Number(risk.likelihood) || 0}</span>
                                <span><strong>A:</strong> ${Number(risk.impact) || 0}</span>
                                <span><strong>Score:</strong> ${risk.qualitativeRiskValue} / 25</span>
                                <span><strong>Schaden:</strong> ${formatCurrency(risk.expectedDamage)}</span>
                              </div>
                            </article>
                          `;
                        }).join("")}
                      </div>`
                    : `<p class="risk-matrix-empty-note">Für diese Matrixzelle sind aktuell keine Risiken vorhanden.</p>`
                ) : `<p class="risk-matrix-empty-note">Wähle ein Feld in der Matrix, um die passenden Risiken zu sehen.</p>`}
              </div>
            </div>
          </details>
          <template id="risk-edit-legacy-template">
            <summary class="risk-fold-summary">
              <div class="risk-fold-summary-main">
                <div class="risk-fold-summary-topline">
                  <span class="risk-fold-toggle" aria-hidden="true"></span>
                  <div class="risk-fold-summary-title">
                    <strong>Risiken erfassen und bearbeiten</strong>
                    <span>Einträge pflegen · Werte, Status und Maßnahmen ändern</span>
                  </div>
                </div>
              </div>
              <div class="risk-fold-summary-actions">
                <span class="badge">Bearbeiten</span>
              </div>
            </summary>
            <div class="risk-fold-body">
              <div class="risk-register-history-actions" style="margin-bottom:14px;">
                <button class="action-btn risk-add-btn" type="button" data-action="add-risk-register-item"><span class="risk-add-icon" aria-hidden="true">+</span><span>Risiko hinzufügen (manuell ohne KI-Unterstützung)</span></button>
                <button class="action-btn" type="button" data-action="undo-risk-register-change" ${riskUndoCount > 0 ? "" : "disabled"}><span class="risk-history-icon" aria-hidden="true">◀</span><span>Letzten Schritt rückgängig</span></button>
                <button class="action-btn" type="button" data-action="redo-risk-register-change" ${riskRedoCount > 0 ? "" : "disabled"}><span class="risk-history-icon" aria-hidden="true">▶</span><span>Letzten Schritt wiederherstellen</span></button>
              </div>
              <div class="risk-edit-list">
              ${result.editRisks.map((risk, index) => `
                <details class="risk-edit-card risk-tone-${classifyRiskTone(risk).key} risk-fold-card risk-edit-fold" ${index === 0 ? "open" : ""}>
                  <summary class="risk-edit-summary">
                    <div class="risk-edit-summary-main">
                      <span class="risk-fold-toggle" aria-hidden="true"></span>
                      <div class="risk-edit-summary-text">
                        <span class="risk-edit-id">${risk.id}</span>
                        <strong>${risk.title}</strong>
                        <span class="risk-edit-summary-value">${risk.qualitativeRiskValue} / 25 · ${formatCurrency(risk.expectedDamage)}</span>
                      </div>
                    </div>
                    <span class="risk-tone-pill risk-tone-${classifyRiskTone(risk).key}">${classifyRiskTone(risk).label}</span>
                  </summary>
                  <div class="risk-edit-grid">
                    <div class="form-field">
                      <label for="risk_id_${index}">Risiko-ID</label>
                      <input id="risk_id_${index}" data-risk-index="${index}" data-risk-field="id" type="text" value="${risk.id}">
                    </div>
                    <div class="form-field half">
                      <label for="risk_title_${index}">Risikobezeichnung</label>
                      <input id="risk_title_${index}" data-risk-index="${index}" data-risk-field="title" type="text" value="${risk.title}">
                    </div>
                    <div class="form-field">
                      <label for="risk_phase_${index}">Projektphase / Zuordnung</label>
                      <select id="risk_phase_${index}" data-risk-index="${index}" data-risk-field="phase">
                        ${buildRiskPhaseOptions(risk.phase).map((phase) => `
                          <option value="${escapeHtml(phase)}" ${String(risk.phase || "").trim().toLowerCase() === String(phase).trim().toLowerCase() ? "selected" : ""}>${escapeHtml(phase)}</option>
                        `).join("")}
                      </select>
                    </div>
                    <div class="form-field wide">
                      <label for="risk_description_${index}">Risikobeschreibung</label>
                      <textarea class="risk-description-textarea" id="risk_description_${index}" data-risk-index="${index}" data-risk-field="description" rows="2">${risk.description}</textarea>
                    </div>
                    <div class="form-field">
                      <label for="risk_category_${index}">Kategorie</label>
                      <select id="risk_category_${index}" data-risk-index="${index}" data-risk-field="category">
                        ${buildRiskCategoryOptions(risk.category).map((category) => `
                          <option value="${escapeHtml(category)}" ${String(risk.category || "").trim().toLowerCase() === String(category).trim().toLowerCase() ? "selected" : ""}>${escapeHtml(category)}</option>
                        `).join("")}
                      </select>
                    </div>
                    <div class="form-field">
                      <label for="risk_owner_${index}">Verantwortlich</label>
                      <input id="risk_owner_${index}" data-risk-index="${index}" data-risk-field="owner" type="text" list="risk_owner_options_${index}" value="${risk.owner}">
                      <datalist id="risk_owner_options_${index}">
                        ${result.ownerOptions.map((owner) => `
                          <option value="${escapeHtml(owner)}"></option>
                        `).join("")}
                      </datalist>
                    </div>
                    <div class="form-field">
                      <label for="risk_status_${index}">Status</label>
                      <select id="risk_status_${index}" data-risk-index="${index}" data-risk-field="status">
                        ${statusOptions.map((status) => `
                          <option value="${status.value}" ${normalizeRiskStatusValue(risk.status) === status.value ? "selected" : ""}>${status.label}</option>
                        `).join("")}
                      </select>
                    </div>
                    <div class="form-field">
                      <label for="risk_due_${index}">Fällig am</label>
                      <input id="risk_due_${index}" data-risk-index="${index}" data-risk-field="dueDate" type="date" value="${risk.dueDate}">
                    </div>
                    <div class="risk-input-grid wide">
                      <div class="form-field">
                        <label for="risk_financial_${index}">Schaden in Euro</label>
                        <div class="eva-stepper-field risk-stepper-field">
                          <input class="compact-input currency-input" id="risk_financial_${index}" data-risk-index="${index}" data-risk-field="financialImpact" type="text" inputmode="numeric" value="${formatCurrencyInput(risk.financialImpact)}">
                          <div class="eva-stepper-buttons">
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="financialImpact" data-step="1000" data-direction="up" data-index="${index}" aria-label="Schaden erhöhen">▲</button>
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="financialImpact" data-step="1000" data-direction="down" data-index="${index}" aria-label="Schaden verringern">▼</button>
                          </div>
                        </div>
                      </div>
                      <div class="form-field">
                        <label for="risk_probability_${index}">Eintrittswahrscheinlichkeit in %</label>
                        <div class="eva-stepper-field risk-stepper-field">
                          <input class="compact-input integer-input" id="risk_probability_${index}" data-risk-index="${index}" data-risk-field="probabilityPercent" type="text" inputmode="numeric" value="${formatIntegerInput(risk.probabilityPercent)}">
                          <div class="eva-stepper-buttons">
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="probabilityPercent" data-step="1" data-direction="up" data-index="${index}" aria-label="Eintrittswahrscheinlichkeit erhöhen">▲</button>
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="probabilityPercent" data-step="1" data-direction="down" data-index="${index}" aria-label="Eintrittswahrscheinlichkeit verringern">▼</button>
                          </div>
                        </div>
                      </div>
                      <div class="form-field">
                        <label>Erwarteter Schaden</label>
                        <div class="risk-value-box risk-value-large">${formatCurrency(risk.expectedDamage)}</div>
                      </div>
                      <div class="risk-grid-spacer" aria-hidden="true"></div>
                    </div>
                    <div class="risk-output-grid wide">
                      <div class="form-field">
                        <label for="risk_likelihood_${index}">Eintrittswahrscheinlichkeit (1-5)</label>
                        <div class="risk-value-box risk-value-large">${deriveRiskLikelihoodFromPercent(risk.probabilityPercent, risk.likelihood)}</div>
                      </div>
                      <div class="form-field">
                        <label for="risk_impact_${index}">Auswirkung (1-5)</label>
                        <div class="eva-stepper-field risk-stepper-field">
                          <input class="compact-input integer-input" id="risk_impact_${index}" data-risk-index="${index}" data-risk-field="impact" type="text" inputmode="numeric" value="${formatIntegerInput(risk.impact)}">
                          <div class="eva-stepper-buttons">
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="impact" data-step="1" data-direction="up" data-index="${index}" aria-label="Auswirkung erhöhen">▲</button>
                            <button type="button" class="eva-stepper-btn" data-action="risk-step" data-field="impact" data-step="1" data-direction="down" data-index="${index}" aria-label="Auswirkung verringern">▼</button>
                          </div>
                        </div>
                      </div>
                      <div class="form-field">
                        <label>Risikowert 5×5</label>
                        <div class="risk-value-box risk-value-large">${risk.qualitativeRiskValue} / 25</div>
                      </div>
                      <div class="risk-grid-spacer" aria-hidden="true"></div>
                    </div>
                    <div class="form-field wide">
                      <label for="risk_measures_${index}">Maßnahmenplanung</label>
                      <textarea id="risk_measures_${index}" data-risk-index="${index}" data-risk-field="measures">${risk.measures}</textarea>
                    </div>
                    <div class="risk-residual-stack">
                      <div class="form-field">
                        <label for="risk_residual_${index}">Rest-Risiko</label>
                        <textarea id="risk_residual_${index}" data-risk-index="${index}" data-risk-field="residualRisk" rows="3">${risk.residualRisk}</textarea>
                      </div>
                      <div class="risk-residual-actions">
                        <button class="risk-remove-btn" type="button" data-action="remove-risk-register-item" data-index="${index}">Risiko entfernen</button>
                      </div>
                    </div>
                  </div>
              `).join("")}
            </div>
            </div>
          </template>
          </section>
        </div>
      `;
    }
  },
  reports: {
    key: "reports",
    label: "Berichte",
    subtitle: "Gemeinsame Berichtsebene",
    render(state) {
      const report = buildManagementReportData(state);
      const reportOptions = getActiveReportSections(state);
      const reportMode = state.ui?.reportMode === "management" ? "management" : "risk";
      const reportExportName = state.ui?.reportExportName || (reportMode === "management" ? "Managementbericht" : "Risikobericht");
      const reportExportFormat = ["json", "txt", "doc", "pdf"].includes(String(state.ui?.reportExportFormat || "").toLowerCase())
        ? String(state.ui.reportExportFormat).toLowerCase()
        : "txt";
      const timeTone = getProjectTimeTone(report.timeContext.status);
      const reportPreviewText = reportMode === "management" ? renderSelectedReportText(state) : renderRiskReportText(state);
      const reportOutputTitle = reportMode === "management"
        ? "Managementbericht: KI-Ausgabe und Berichtstext"
        : "Risikobericht: KI-Ausgabe und Berichtstext";
      const reportOutputButtonLabel = reportMode === "management"
        ? "Managementbericht exportieren"
        : "Risikobericht exportieren";
      const reportSelectionTitle = reportMode === "management"
        ? "Berichtsauswahl für den Managementbericht"
        : "Berichtsauswahl für den Risikobericht";
      const reportSelectionHint = reportMode === "management"
        ? "Wähle die Module, die in den Managementbericht und in die KI-Ausgabe einfließen sollen."
        : "Wähle die Module, die in den Risikobericht und in die KI-Ausgabe einfließen sollen.";
      return `
        <div class="module-shell">
          <div class="module-title">
            <div>
              <h2>Berichte</h2>
              <p>Hier laufen jetzt die Kernsignale aus PERT, Monte Carlo, Leistungswert-Analyse (EVA) und Risikoregister zusammen.<br>Die Berichtsebene verdichtet die vier Module zu einer gemeinsamen Management-Sicht.</p>
            </div>
            <span class="badge">Gemeinsame Ausgabe</span>
          </div>
          <section class="info-card report-mode-card-panel">
            <h3>Berichtstyp wählen</h3>
            <p class="form-note">Wähle zuerst, ob du den Risikobericht oder den Managementbericht sehen und exportieren möchtest.</p>
            <div class="report-mode-grid">
              <button class="report-mode-card ${reportMode === "risk" ? "is-active" : ""}" type="button" data-action="set-report-mode" data-report-mode="risk">
                <span class="report-mode-check" aria-hidden="true"></span>
                <strong>Risikobericht</strong>
                <span>Risiken, Maßnahmen, Restgefahr und Prioritäten im Fokus.</span>
              </button>
              <button class="report-mode-card ${reportMode === "management" ? "is-active" : ""}" type="button" data-action="set-report-mode" data-report-mode="management">
                <span class="report-mode-check" aria-hidden="true"></span>
                <strong>Managementbericht</strong>
                <span>Gesamtbild aus PERT, Monte Carlo, EVA und Risikoregister.</span>
              </button>
            </div>
          </section>
          <div class="kpi-grid">
            ${report.summaryCards.map((card, index) => `
              <article class="kpi-card ${card.tone === "critical" ? "time-critical" : card.tone === "warn" ? "time-warn" : card.tone === "active" ? "time-active" : index === 1 ? "gold" : index === 2 ? "green" : index === 3 ? "report-eva" : index === 4 ? "report-risk" : ""} ${card.linkModule ? "report-summary-card" : ""}" ${card.linkModule ? `data-action="jump-to-report-target" data-target-module="${card.linkModule}" data-target-id="${card.linkTarget}" role="button" tabindex="0" aria-label="${card.linkLabel}"` : ""}>
                <div class="kpi-label">${card.label}</div>
                <div class="kpi-value">${card.value}</div>
                <div class="kpi-sub">${card.detail}</div>
              </article>
            `).join("")}
          </div>
          <section class="info-card">
            <h3>${reportSelectionTitle}</h3>
            <p class="form-note">Klicke auf eine Kachel, um das Modul für den gewählten Bericht zu aktivieren oder zu deaktivieren.</p>
            <div class="report-option-grid">
              <label class="report-option-card ${reportOptions.includePert ? "is-active" : ""}">
                <input data-report-option="includePert" type="checkbox" ${reportOptions.includePert ? "checked" : ""}>
                <span class="report-option-check" aria-hidden="true"></span>
                <div class="report-option-copy">
                  <strong>PERT-Methode</strong>
                  <span>Fließzeit, Erwartungswert und Bandbreite für den Termin- und Kostenausblick.</span>
                </div>
              </label>
              <label class="report-option-card ${reportOptions.includeMonteCarlo ? "is-active" : ""}">
                <input data-report-option="includeMonteCarlo" type="checkbox" ${reportOptions.includeMonteCarlo ? "checked" : ""}>
                <span class="report-option-check" aria-hidden="true"></span>
                <div class="report-option-copy">
                  <strong>Monte-Carlo-Simulation</strong>
                  <span>Wahrscheinlichkeit, Kostenspanne und Zielerreichung aus der Simulationssicht.</span>
                </div>
              </label>
              <label class="report-option-card ${reportOptions.includeEarnedValue ? "is-active" : ""}">
                <input data-report-option="includeEarnedValue" type="checkbox" ${reportOptions.includeEarnedValue ? "checked" : ""}>
                <span class="report-option-check" aria-hidden="true"></span>
                <div class="report-option-copy">
                  <strong>Leistungswert-Analyse (EVA)</strong>
                  <span>Planwert, Istkosten, Leistungswert und Prognose für Kosten und Termin.</span>
                </div>
              </label>
              <label class="report-option-card ${reportOptions.includeRiskRegister ? "is-active" : ""}">
                <input data-report-option="includeRiskRegister" type="checkbox" ${reportOptions.includeRiskRegister ? "checked" : ""}>
                <span class="report-option-check" aria-hidden="true"></span>
                <div class="report-option-copy">
                  <strong>Risikoregister</strong>
                  <span>Top-Risiken, Maßnahmen, Restgefahr und Prioritäten für den Bericht.</span>
                </div>
              </label>
            </div>
            <p class="form-note">${reportSelectionHint}</p>
          </section>
          <section class="info-card" style="margin-top:18px;">
            <h3>Berichtskontext</h3>
            <p class="form-note">Diese Daten bilden die gemeinsame Basis für Export, Druckansicht und spätere KI-Ausgabe. Sie sorgen dafür, dass alle Berichtsteile denselben Projektkontext verwenden.</p>
            <ul>
              <li><strong>Projekt:</strong> ${report.topline.projectName}</li>
              <li><strong>Projektart / Bauart:</strong> ${report.topline.type} · ${report.topline.bauart}</li>
              <li><strong>Projektstandort:</strong> ${report.topline.projectAddress}</li>
              <li><strong>Auftraggeber:</strong> ${report.topline.client}</li>
              <li><strong>Projektleitung:</strong> ${report.topline.projectLead}</li>
              <li><strong>Berichtsdatum:</strong> ${report.topline.reportDate}</li>
              <li><strong>Betrachtungszeitpunkt:</strong> ${report.topline.analysisDate}</li>
              <li><strong>Verfügbare Module:</strong> PERT · Monte Carlo · EVA · Risikoregister</li>
              <li><strong>Gewählte Module:</strong> ${[
                reportOptions.includePert ? "PERT" : null,
                reportOptions.includeMonteCarlo ? "Monte Carlo" : null,
                reportOptions.includeEarnedValue ? "EVA" : null,
                reportOptions.includeRiskRegister ? "Risikoregister" : null
              ].filter(Boolean).join(" · ") || "keine"}</li>
            </ul>
          </section>
          <section class="info-card report-output-card" style="margin-top:18px;">
            <h3>${reportOutputTitle}</h3>
            <p class="form-note">Hier entsteht später die KI-generierte Berichtsausgabe. Die Vorschau folgt dem gewählten Berichtstyp und den ausgewählten Modulen. Dateiname und Format wählst du unten bei den zentralen Exporten.</p>
            <div class="code-box report-output-box">${reportPreviewText}</div>
          </section>
          <section class="info-card" style="margin-top:18px;">
            <h3>KI für Berichtstexte</h3>
            <p class="form-note">Hier erzeugst du Management- oder Risikoberichte direkt aus den aktuellen Projektdaten. Die Ausgabe bleibt auf Berichtsebene und ergänzt die Vorschau.</p>
            <div class="ai-workshop-grid">
              <button class="action-btn primary ${state.ui?.aiWorkshop?.activeTask === "management-report" ? "active" : ""}" type="button" data-action="run-ai-workshop" data-ai-workshop-task="management-report">Managementbericht erzeugen</button>
              <button class="action-btn primary ${state.ui?.aiWorkshop?.activeTask === "risk-report" ? "active" : ""}" type="button" data-action="run-ai-workshop" data-ai-workshop-task="risk-report">Risikobericht erzeugen</button>
            </div>
            <p class="form-note">Die KI verdichtet hier nur die Berichtsdaten. Risiken werden nicht direkt im Register verändert, sondern als Berichtsausgabe dargestellt.</p>
            <div class="code-box ai-workshop-output ${state.ui?.aiWorkshop?.busy ? "is-loading" : ""} ${state.ui?.aiWorkshop?.resultTone === "success" ? "tone-success" : state.ui?.aiWorkshop?.resultTone === "danger" ? "tone-danger" : "tone-neutral"}">
              <strong>${escapeHtml(state.ui?.aiWorkshop?.resultTitle || "KI bereit")}</strong>
              <div style="margin-top:8px; white-space:pre-wrap;">${escapeHtml(state.ui?.aiWorkshop?.resultText || "Wähle eine Berichtsfunktion für die KI-Ausgabe.")}</div>
            </div>
          </section>
          <section class="info-card">
            <h3>Priorisierte Risiken</h3>
            <p class="form-note">Die priorisierten Risiken werden aus dem aktuellen Risikoregister abgeleitet. Maßgeblich sind der 5×5-Wert, der erwartete Schaden und die aktuelle Terminlage. So siehst du zuerst die Risiken mit dem höchsten Steuerungsbedarf.</p>
            ${report.topRisks.length ? `
              <div class="report-top-risk-grid">
                ${report.topRisks.map((risk) => `
                  ${(() => {
                    const [lineOne, lineTwo] = splitTitleIntoTwoLines(risk.title);
                    return `
                  <div class="form-field risk-tone-card risk-tone-${risk.score.includes("15") ? "critical" : risk.score.includes("12") ? "warn" : "neutral"}" style="padding:12px 12px 11px;border:1px solid var(--line);border-radius:14px;">
                    <label>${risk.id}</label>
                    <div class="risk-top-risk-title">
                      <span class="risk-top-risk-title-line">${lineOne}</span>
                      <span class="risk-top-risk-title-line ${lineTwo ? "" : "is-empty"}">${lineTwo || " "}</span>
                    </div>
                    <ul class="risk-top-risk-list">
                      <li>Bewertung: ${risk.score}</li>
                      <li>Schaden: ${risk.expectedDamage}</li>
                      <li>Verantwortlich: ${risk.owner}</li>
                      <li>Status: ${risk.status}</li>
                    </ul>
                  </div>
                    `;
                  })()}
                `).join("")}
              </div>
            ` : `
              <p>Aktuell liegen keine priorisierten Risiken für die Management-Sicht vor.</p>
            `}
          </section>
          <div class="card-grid report-export-grid-shell" style="grid-template-columns:minmax(0,1fr); align-items:start;">
            <section class="info-card">
              <h3>Reportprofil</h3>
              <p class="form-note">Die Vertraulichkeit legst du hier fest. Sie steuert, ob der Bericht intern, vertraulich oder stärker geschützt verwendet wird.</p>
              <ul>
                <li><strong>Erstellt von:</strong> ${state.reportProfile.author || "nicht gepflegt"}</li>
                <li><strong>Firma:</strong> ${state.reportProfile.company || "nicht gepflegt"}</li>
                <li><strong>Empfänger:</strong> ${state.reportProfile.clientName || "nicht gepflegt"}</li>
                <li><strong>Vertraulichkeit:</strong> ${state.reportProfile.confidentiality || "Vertraulich"}</li>
              </ul>
            </section>
            <section class="info-card">
              <h3>Exporte</h3>
              <div class="export-stack">
                <div class="export-mini-card">
                  <h4>Bericht exportieren</h4>
                  <p class="form-note">Wähle hier den Dateinamen und das Format für den Bericht.</p>
                  <ul class="export-format-help">
                    <li><strong>TXT (.txt):</strong> Für Mail, Protokoll, schnelle Weitergabe und das Einfügen in andere Dokumente.</li>
                    <li><strong>DOC (.doc):</strong> Für die Bearbeitung in Word, redaktionelle Nacharbeit und spätere Kommentierung.</li>
                    <li><strong>PDF (.pdf):</strong> Für Freigabe, Versand, Druckansicht und eine unveränderliche Dokumentfassung.</li>
                    <li><strong>JSON (.json):</strong> Für strukturierte Berichtsdaten, Import, KI-Verarbeitung und technische Weitergabe.</li>
                  </ul>
                  <div class="report-export-grid">
                    <div class="form-field">
                      <label for="reportExportFileName">Dateiname</label>
                      <input id="reportExportFileName" data-report-export-field="fileName" type="text" value="${reportExportName}" placeholder="${reportMode === "management" ? "Managementbericht" : "Risikobericht"}">
                    </div>
                    <div class="form-field">
                      <label for="reportExportFormat">Format</label>
                      <select id="reportExportFormat" data-report-export-field="format">
                        <option value="txt"${reportExportFormat === "txt" ? " selected" : ""}>Text (.txt)</option>
                        <option value="json"${reportExportFormat === "json" ? " selected" : ""}>Berichtsdaten (.json)</option>
                        <option value="doc"${reportExportFormat === "doc" ? " selected" : ""}>Word-Dokument (.doc)</option>
                        <option value="pdf"${reportExportFormat === "pdf" ? " selected" : ""}>PDF (Druckansicht)</option>
                      </select>
                    </div>
                  </div>
                  <p class="form-note">Alle Berichtsexporte werden im Downloads-Ordner gespeichert.</p>
                  <div class="report-output-actions">
                    <button class="action-btn primary" type="button" data-action="export-selected-report">${reportOutputButtonLabel}</button>
                  </div>
                </div>
              </div>
            </section>
          </div>
        </div>
      `;
    }
  }
};
