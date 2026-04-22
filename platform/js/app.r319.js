import { createStore, cloneState } from "./state.r319.js";
import { modules, riskCategoryOptions, normalizeRiskCategoryValue, normalizeRiskStatusValue, deriveRiskLikelihoodFromPercent, calculatePertResult, calculateMonteCarloResult, calculateEarnedValueResult, buildManagementReportData, buildSelectedReportData, renderSelectedReportText, renderRiskReportText } from "./modules.r319.js";

const store = createStore(cloneState());
const AUTOSAVE_KEY = "project_controls_hub_autosave_v3";
const AI_SETTINGS_KEY = "project_controls_hub_ai_settings_v1";
let lastAutosavedAt = null;
let suppressNextAutosave = 0;
let aiSettings = loadAiSettings();
let aiConnectionAbortController = null;

function splitStreetAddress(rawValue) {
  const value = String(rawValue || "").trim().replace(/\s+/g, " ");
  const match = value.match(/^(.*\S)\s+(\d+\w*)$/);
  if (!match) return { street: value, houseNumber: "" };
  return {
    street: match[1],
    houseNumber: match[2]
  };
}

function formatTimestamp(value) {
  if (!value) return "noch nicht gespeichert";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Zeitstempel ungültig";
  return date.toLocaleString("de-DE");
}

function jumpToMonteCarloP80() {
  store.setState((state) => ({
    ...state,
    ui: {
      ...state.ui,
      activeModule: "monteCarlo"
    }
  }));
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      const target = document.getElementById("monte-p80-costs");
      if (!target) return;
      target.scrollIntoView({ behavior: "smooth", block: "start" });
      target.classList.add("jump-anchor-highlight");
      window.setTimeout(() => target.classList.remove("jump-anchor-highlight"), 1600);
    });
  });
}

function jumpToReportTarget(moduleKey, targetId) {
  if (!moduleKey || !targetId) return;
  store.setState((state) => ({
    ...state,
    ui: {
      ...state.ui,
      activeModule: moduleKey
    }
  }));
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      const target = document.getElementById(targetId);
      if (!target) return;
      target.scrollIntoView({ behavior: "smooth", block: "start" });
      target.classList.add("jump-anchor-highlight");
      window.setTimeout(() => target.classList.remove("jump-anchor-highlight"), 1600);
    });
  });
}

function updateStorageStatus(message) {
  const target = document.getElementById("storageStatus");
  if (target) target.textContent = message;
}

function applyAiStatusAppearance(target, message) {
  if (!target) return;
  target.classList.remove("ai-status-success", "ai-status-danger", "ai-status-neutral");
  if (message === "Verbindung ok") {
    target.classList.add("ai-status-success");
  } else if (String(message).includes("fehlgeschlagen") || String(message).includes("zu lange")) {
    target.classList.add("ai-status-danger");
  } else {
    target.classList.add("ai-status-neutral");
  }
}

function updateAiStatus(message) {
  const target = document.getElementById("aiStatus");
  if (target) {
    target.textContent = message;
    applyAiStatusAppearance(target, message);
  }
}

function getDefaultAiSettings() {
  return {
    provider: "anthropic",
    apiKey: "",
    modelProfile: "balanced",
    connected: false,
    testing: false,
    lastSavedAt: null,
    lastTestAt: null,
    lastDisconnectAt: null,
    lastStatus: "Noch keine KI-Verbindung eingerichtet."
  };
}

function normalizeAiSettings(raw = {}) {
  const provider = "anthropic";
  return {
    ...getDefaultAiSettings(),
    provider,
    apiKey: typeof raw.apiKey === "string" ? raw.apiKey : "",
    modelProfile: "balanced",
    connected: raw.connected === true,
    testing: raw.testing === true,
    lastSavedAt: typeof raw.lastSavedAt === "string" ? raw.lastSavedAt : null,
    lastTestAt: typeof raw.lastTestAt === "string" ? raw.lastTestAt : null,
    lastDisconnectAt: typeof raw.lastDisconnectAt === "string" ? raw.lastDisconnectAt : null,
    lastStatus: typeof raw.lastStatus === "string" && raw.lastStatus.trim() ? raw.lastStatus : getDefaultAiSettings().lastStatus
  };
}

function loadAiSettings() {
  try {
    const raw = localStorage.getItem(AI_SETTINGS_KEY);
    if (!raw) return getDefaultAiSettings();
    return normalizeAiSettings(JSON.parse(raw));
  } catch (_error) {
    return getDefaultAiSettings();
  }
}

function persistAiSettings(nextSettings) {
  aiSettings = normalizeAiSettings(nextSettings);
  try {
    localStorage.setItem(AI_SETTINGS_KEY, JSON.stringify(aiSettings));
    return true;
  } catch (_error) {
    return false;
  }
}

function getAiProviderLabel(provider) {
  return "Anthropic";
}

function getAiProfileLabel(profile) {
  return "Ausgewogen";
}

function normalizeAiWorkshopState(workshop = {}) {
  return {
    activeTask: ["management-report", "risk-report", "free-text-risks", "measures-residual"].includes(workshop.activeTask)
      ? workshop.activeTask
      : "management-report",
    freeText: typeof workshop.freeText === "string" ? workshop.freeText : "",
    selectedRiskId: typeof workshop.selectedRiskId === "string" ? workshop.selectedRiskId : "",
    busy: workshop.busy === true,
    resultTitle: typeof workshop.resultTitle === "string" && workshop.resultTitle.trim() ? workshop.resultTitle : "KI bereit",
    resultText: typeof workshop.resultText === "string" && workshop.resultText.trim()
      ? workshop.resultText
      : "Wähle eine Funktion für die KI-Startstufe.",
    resultTone: ["neutral", "success", "danger"].includes(workshop.resultTone) ? workshop.resultTone : "neutral",
    resultData: workshop.resultData && typeof workshop.resultData === "object" ? workshop.resultData : null
  };
}

function normalizeRiskRegisterUndoStack(stack) {
  if (!Array.isArray(stack)) return [];
  return stack
    .filter((entry) => Array.isArray(entry) && entry.length)
    .map((entry) => JSON.parse(JSON.stringify(entry)))
    .slice(-20);
}

function pushRiskRegisterUndoSnapshot(state) {
  const currentRisks = JSON.parse(JSON.stringify(state.riskRegister?.risks || []));
  const currentStack = normalizeRiskRegisterUndoStack(state.ui?.riskRegisterUndoStack);
  const nextStack = [...currentStack, currentRisks].slice(-20);
  store.setState((current) => ({
    ...current,
    ui: {
      ...current.ui,
      riskRegisterUndoStack: nextStack
    }
  }));
}

function applyRiskRegisterUndo() {
  const state = store.getState();
  const stack = normalizeRiskRegisterUndoStack(state.ui?.riskRegisterUndoStack);
  if (!stack.length) return false;
  const previousRisks = stack[stack.length - 1];
  const nextStack = stack.slice(0, -1);
  store.setState((current) => ({
    ...current,
    riskRegister: {
      ...current.riskRegister,
      risks: JSON.parse(JSON.stringify(previousRisks))
    },
    ui: {
      ...current.ui,
      riskRegisterUndoStack: nextStack
    }
  }));
  return true;
}

function getAiStatusLabel(settings = aiSettings) {
  const hasKey = String(settings.apiKey || "").trim().length > 0;
  if (settings.testing) return "Verbindung wird geprüft ...";
  if (settings.connected) return "Verbindung ok";
  if (settings.lastStatus && settings.lastStatus !== getDefaultAiSettings().lastStatus) return settings.lastStatus;
  if (!hasKey) return "Noch keine KI-Verbindung eingerichtet.";
  return "Einstellungen speichern";
}

function renderAiSettingsPanel() {
  const panel = document.getElementById("aiConnectionPanel");
  const apiKeyInput = document.getElementById("aiApiKey");
  const statusTarget = document.getElementById("aiStatus");
  const testButton = document.getElementById("testAiSettingsBtn");
  const disconnectButton = document.getElementById("disconnectAiSettingsBtn");
  if (panel) {
    panel.classList.remove("ai-connected", "ai-disconnected");
    panel.classList.add(aiSettings.connected ? "ai-connected" : "ai-disconnected");
  }
  if (apiKeyInput && document.activeElement !== apiKeyInput) {
    apiKeyInput.value = aiSettings.apiKey || "";
  }
  if (statusTarget) {
    const statusText = getAiStatusLabel(aiSettings);
    statusTarget.textContent = statusText;
    applyAiStatusAppearance(statusTarget, statusText);
  }
  if (testButton) {
    testButton.textContent = aiSettings.testing ? "Verbindung wird geprüft ..." : "Verbindung testen";
    testButton.disabled = aiSettings.testing;
    testButton.classList.toggle("is-loading", aiSettings.testing);
  }
  if (disconnectButton) {
    disconnectButton.style.display = "flex";
    disconnectButton.disabled = !aiSettings.connected;
  }
}

function readAiSettingsFromPanel() {
  const apiKey = String(document.getElementById("aiApiKey")?.value || "");
  return normalizeAiSettings({
    provider: "anthropic",
    modelProfile: "balanced",
    apiKey,
    connected: aiSettings.connected,
    testing: aiSettings.testing,
    lastSavedAt: aiSettings.lastSavedAt,
    lastTestAt: aiSettings.lastTestAt,
    lastDisconnectAt: aiSettings.lastDisconnectAt,
    lastStatus: aiSettings.lastStatus
  });
}

function applyAiSettings(nextAiSettings, statusMessage) {
  const saved = persistAiSettings(nextAiSettings);
  aiSettings = saved ? loadAiSettings() : normalizeAiSettings(nextAiSettings);
  if (statusMessage) updateAiStatus(statusMessage);
  renderAiSettingsPanel();
}

async function startAiConnectionTest() {
  if (aiConnectionAbortController) return;
  const nextAiSettings = readAiSettingsFromPanel();
  const snapshot = normalizeAiSettings({
    ...nextAiSettings,
    testing: true,
    connected: false,
    lastStatus: "Verbindung wird geprüft ..."
  });
  if (!String(snapshot.apiKey || "").trim()) {
    aiSettings = normalizeAiSettings({
      ...snapshot,
      testing: false,
      connected: false,
      lastStatus: "Bitte zuerst einen API-Schlüssel eingeben."
    });
    persistAiSettings(aiSettings);
    renderAiSettingsPanel();
    updateAiStatus(aiSettings.lastStatus);
    return;
  }

  aiSettings = snapshot;
  persistAiSettings(aiSettings);
  renderAiSettingsPanel();
  updateAiStatus(aiSettings.lastStatus);

  if (aiConnectionAbortController) {
    aiConnectionAbortController.abort();
  }
  aiConnectionAbortController = new AbortController();
  const connectionTimeout = window.setTimeout(() => aiConnectionAbortController?.abort(), 12000);

  try {
    const response = await fetch("http://127.0.0.1:8161/api/ai/test", {
      method: "POST",
      headers: {
        "content-type": "application/json"
      },
      body: JSON.stringify({
        provider: "anthropic",
        apiKey: snapshot.apiKey,
        modelProfile: "balanced"
      }),
      signal: aiConnectionAbortController.signal
    });

    if (!response.ok) {
      const errorPayload = await response.json().catch(() => ({}));
      const extra = errorPayload?.providerBody ? `\n${String(errorPayload.providerBody).trim()}` : "";
      throw new Error(`${errorPayload?.error || `HTTP ${response.status}`}${extra}`);
    }

    aiSettings = normalizeAiSettings({
      ...snapshot,
      testing: false,
      connected: true,
      lastTestAt: new Date().toISOString(),
      lastSavedAt: aiSettings.lastSavedAt || new Date().toISOString(),
      lastStatus: "Verbindung ok"
    });
    persistAiSettings(aiSettings);
    renderAiSettingsPanel();
    updateAiStatus(aiSettings.lastStatus);
  } catch (error) {
    const message = error?.name === "AbortError"
      ? "KI-Verbindung hat zu lange gedauert."
      : String(error?.message || "").includes("fetch")
        ? "Lokaler KI-Proxy nicht erreichbar. Bitte ai-proxy-server.js starten."
        : "KI-Verbindung fehlgeschlagen. Bitte API-Schlüssel oder Provider prüfen.";
    aiSettings = normalizeAiSettings({
      ...snapshot,
      testing: false,
      connected: false,
      lastStatus: message
    });
    persistAiSettings(aiSettings);
    renderAiSettingsPanel();
    updateAiStatus(message);
  } finally {
    window.clearTimeout(connectionTimeout);
    aiConnectionAbortController = null;
  }
}

function disconnectAiConnection() {
  if (aiConnectionAbortController) {
    aiConnectionAbortController.abort();
    aiConnectionAbortController = null;
  }
  aiSettings = normalizeAiSettings({
    ...readAiSettingsFromPanel(),
    apiKey: "",
    connected: false,
    testing: false,
    lastDisconnectAt: new Date().toISOString(),
    lastStatus: "Verbindung getrennt"
  });
  persistAiSettings(aiSettings);
  renderAiSettingsPanel();
  updateAiStatus(aiSettings.lastStatus);
}

function buildAiWorkshopPayload(state, task) {
  const baseReport = buildManagementReportData(state);
  const reportMode = task === "management-report" ? "management" : "risk";
  const selectedReport = buildSelectedReportData({
    ...state,
    ui: {
      ...state.ui,
      reportMode
    }
  });
  return {
    task,
    reportMode,
    project: selectedReport.project,
    report: selectedReport.report,
    modules: selectedReport.modules,
    selectedRisk: state.riskRegister?.risks?.find((risk) => risk.id === state.ui?.aiWorkshop?.selectedRiskId) || null,
    context: {
      topline: baseReport.topline,
      timeContext: baseReport.timeContext,
      summaryCards: baseReport.summaryCards,
      focusPoints: baseReport.focusPoints,
      nextSteps: baseReport.nextSteps,
      topRisks: baseReport.topRisks,
      riskRegister: baseReport.raw?.riskRegister,
      pert: baseReport.raw?.pert,
      monteCarlo: baseReport.raw?.monteCarlo,
      earnedValue: baseReport.raw?.earnedValue,
      selectedModules: selectedReport.report?.selectedModules || []
    }
  };
}

function buildAiWorkshopSystemPrompt(task) {
  switch (task) {
    case "management-report":
      return [
        "Du bist ein präziser Projekt- und Managementbericht-Assistent.",
        "Schreibe auf Deutsch, fachlich sauber, kurz und entscheidungsorientiert.",
        "Nutze ausschließlich die gelieferten Daten.",
        "Struktur: Lagebild, Kernaussagen, Risiken, Maßnahmen, Nächste Prioritäten."
      ].join(" ");
    case "risk-report":
      return [
        "Du bist ein operativer Risikobericht-Assistent.",
        "Schreibe auf Deutsch, klar und handlungsorientiert.",
        "Nutze ausschließlich die gelieferten Daten.",
        "Struktur: Lagebild, Top-Risiken, Maßnahmen, Restgefahr, Nächste Prioritäten."
      ].join(" ");
    case "free-text-risks":
      return [
        "Du extrahierst aus Freitext strukturierte Risiken.",
        "Antworte ausschließlich als JSON ohne Markdown und ohne Fließtext außerhalb des JSON.",
        "Form: {\"resultType\":\"risk-suggestion\",\"items\":[{\"title\":\"\",\"description\":\"\",\"category\":\"\",\"phase\":\"\",\"owner\":\"\",\"probabilityPercent\":0,\"likelihood\":0,\"impact\":0,\"financialImpact\":0,\"measures\":\"\",\"residualRisk\":\"\"}]}",
        "Der Vorschlag soll die Werte 'Schaden in Euro', 'Eintrittswahrscheinlichkeit in %' und 'Erwarteter Schaden' untereinander darstellen.",
        "Die Eintrittswahrscheinlichkeit (1-5) und der Erwartete Schaden werden von der App abgeleitet bzw. angezeigt.",
        "Das Feld 'residualRisk' ist ausschließlich qualitativ zu formulieren. Keine Euro-Beträge und keine Scheingenauigkeit.",
        "Wenn mehrere Risiken erkennbar sind, fülle items mit mehreren Einträgen.",
        "Wenn nur ein Risiko erkennbar ist, liefere genau einen Eintrag."
      ].join(" ");
    case "measures-residual":
      return [
        "Du leitest aus den gelieferten Risiken konkrete Maßnahmen und eine plausible Restgefahr ab.",
        "Antworte ausschließlich als JSON ohne Markdown und ohne Fließtext außerhalb des JSON.",
        "Form: {\"resultType\":\"risk-measures\",\"items\":[{\"riskId\":\"\",\"measures\":\"\",\"residualRisk\":\"\",\"owner\":\"\",\"dueDate\":\"\"}]}",
        "Erzeuge genau einen Eintrag für genau ein Zielrisiko.",
        "Verwende ausschließlich das ausgewählte Zielrisiko aus dem Kontext und nenne dessen riskId im Feld riskId.",
        "Das Feld 'residualRisk' ist ausschließlich qualitativ zu formulieren. Keine Euro-Beträge und keine Scheingenauigkeit."
      ].join(" ");
    default:
      return "Du bist ein präziser Projektassistenz für Management- und Risikoberichte. Antworte auf Deutsch, fachlich sauber und knapp.";
  }
}

function buildAiWorkshopUserPrompt(state, task) {
  const payload = buildAiWorkshopPayload(state, task);
  const freeText = String(state.ui?.aiWorkshop?.freeText || "").trim();
  const summary = {
    projekt: payload.project,
    bericht: {
      mode: payload.reportMode,
      selectedModules: payload.report.selectedModules,
      timeContext: payload.context.timeContext,
      topRisks: payload.context.topRisks,
      focusPoints: payload.context.focusPoints,
      nextSteps: payload.context.nextSteps,
      selectedRisk: payload.selectedRisk
    },
    moduleKennzahlen: {
      pert: payload.modules.pert || null,
      monteCarlo: payload.modules.monteCarlo || null,
      earnedValue: payload.modules.earnedValue || null,
      riskRegister: payload.modules.riskRegister || null
    }
  };

  return [
    `Aufgabe: ${task}`,
    "",
    "Erstelle eine fachlich präzise Antwort auf Deutsch.",
    "Nutze die folgenden Daten.",
    "Antworte nur mit dem Ergebnis, ohne Meta-Erklärungen.",
    "",
    JSON.stringify(summary, null, 2),
    freeText ? "" : "",
    freeText ? "Freitext:" : "",
    freeText ? freeText : ""
  ].filter((part) => part !== "").join("\n");
}

function setAiWorkshopState(nextWorkshop) {
  store.setState((state) => ({
    ...state,
    ui: {
      ...state.ui,
      aiWorkshop: normalizeAiWorkshopState(nextWorkshop)
    }
  }));
}

function parseAiWorkshopResult(text, task) {
  const raw = String(text || "").trim();
  if (!raw) return null;
  const jsonStart = raw.indexOf("{");
  const jsonEnd = raw.lastIndexOf("}");
  if (jsonStart === -1 || jsonEnd === -1 || jsonEnd <= jsonStart) return null;
  try {
    const parsed = JSON.parse(raw.slice(jsonStart, jsonEnd + 1));
    if (task === "free-text-risks" && parsed?.resultType === "risk-suggestion") return parsed;
    if (task === "measures-residual" && parsed?.resultType === "risk-measures") return parsed;
    return parsed;
  } catch (_error) {
    return null;
  }
}

function normalizeResidualRiskText(value) {
  const text = String(value || "").trim();
  if (!text) return "Nach Maßnahmen bleibt ein qualitativ reduziertes Rest-Risiko bestehen.";
  const containsCurrency = /(?:\d[\d.\s]*,\d+|\d[\d.\s]*)\s*(?:€|EUR)\b|€|\bEUR\b/i.test(text);
  const cleaned = text
    .replace(/(?:\d[\d.\s]*,\d+|\d[\d.\s]*)\s*(?:€|EUR)\b/gi, "")
    .replace(/€|\bEUR\b/gi, "")
    .replace(/\s{2,}/g, " ")
    .replace(/\s+([,.;:])/g, "$1")
    .trim();
  if (containsCurrency) {
    return cleaned && cleaned.length >= 18
      ? cleaned
      : "Nach Maßnahmen bleibt ein qualitativ reduziertes Rest-Risiko bestehen.";
  }
  return cleaned || "Nach Maßnahmen bleibt ein qualitativ reduziertes Rest-Risiko bestehen.";
}

function sanitizeAiWorkshopResult(result, task) {
  if (!result || !Array.isArray(result.items)) return result;
  if (task !== "free-text-risks" && task !== "measures-residual") return result;
  return {
    ...result,
    items: result.items.map((item) => ({
      ...item,
      residualRisk: normalizeResidualRiskText(item?.residualRisk)
    }))
  };
}

function narrowAiWorkshopMeasuresResult(result, selectedRiskId = "") {
  if (!result || result.resultType !== "risk-measures" || !Array.isArray(result.items) || !result.items.length) return result;
  const targetId = String(selectedRiskId || "").trim().toLowerCase();
  if (!targetId) {
    return {
      ...result,
      items: [result.items[0]]
    };
  }
  const matched = result.items.find((item) => String(item?.riskId || "").trim().toLowerCase() === targetId);
  return {
    ...result,
    items: [matched || result.items[0]]
  };
}

function buildAiWorkshopFallbackResult(state, task, responseText = "") {
  const aiWorkshop = state.ui?.aiWorkshop || {};
  const freeText = String(aiWorkshop.freeText || "").trim();
  const selectedRisk = state.riskRegister?.risks?.find((risk) => String(risk.id || "") === String(aiWorkshop.selectedRiskId || "")) || null;
  const descriptionSource = [freeText, responseText].map((value) => String(value || "").trim()).find(Boolean) || "";
  if (task === "free-text-risks") {
    const selectedProbabilityPercent = Number(selectedRisk?.probabilityPercent);
    const titleSource = selectedRisk?.title || freeText.split(/\r?\n/)[0] || descriptionSource.split(/\r?\n/)[0] || "Neues Risiko";
    return {
      resultType: "risk-suggestion",
      items: [
        {
          title: String(titleSource).trim().slice(0, 120) || "Neues Risiko",
          description: descriptionSource || "Aus dem Freitext wurde ein Risikoentwurf vorbereitet.",
          category: normalizeRiskCategoryValue(selectedRisk?.category || riskCategoryOptions[4] || "Projekt- und Managementrisiken"),
          phase: selectedRisk?.phase || "",
          owner: selectedRisk?.owner || "",
          probabilityPercent: Number.isFinite(selectedProbabilityPercent) ? Math.max(0, Math.min(100, selectedProbabilityPercent)) : 50,
          likelihood: Number.isFinite(selectedProbabilityPercent)
            ? deriveRiskLikelihoodFromPercent(selectedProbabilityPercent, selectedRisk?.likelihood || 3)
            : Math.max(1, Math.min(5, Number(selectedRisk?.likelihood) || 3)),
          impact: Number(selectedRisk?.impact) || 3,
          financialImpact: Number(selectedRisk?.financialImpact) || 0,
          measures: selectedRisk?.measures || "",
          residualRisk: normalizeResidualRiskText(selectedRisk?.residualRisk || "")
        }
      ]
    };
  }
  if (task === "measures-residual") {
    const riskId = String(aiWorkshop.selectedRiskId || selectedRisk?.id || "").trim();
    if (!riskId) return null;
    return {
      resultType: "risk-measures",
      items: [
        {
          riskId,
          measures: descriptionSource || selectedRisk?.measures || "Maßnahmen bitte fachlich prüfen und ergänzen.",
          residualRisk: normalizeResidualRiskText(selectedRisk?.residualRisk || "Rest-Risiko bitte fachlich prüfen."),
          owner: selectedRisk?.owner || "",
          dueDate: selectedRisk?.dueDate || ""
        }
      ]
    };
  }
  return null;
}

function mapSuggestedRiskToRegisterItem(suggestion = {}, fallbackId = "") {
  const hasProbabilityPercent = suggestion.probabilityPercent !== undefined && suggestion.probabilityPercent !== null && String(suggestion.probabilityPercent).trim() !== "";
  const explicitProbabilityPercent = Number(suggestion.probabilityPercent);
  const likelihood = hasProbabilityPercent
    ? deriveRiskLikelihoodFromPercent(explicitProbabilityPercent, Number(suggestion.likelihood) || 1)
    : Math.max(1, Math.min(5, Number(suggestion.likelihood) || Number(suggestion.probability) || 1));
  const probabilityPercent = Math.max(
    0,
    Math.min(
      100,
      hasProbabilityPercent ? explicitProbabilityPercent : (Number.isFinite(likelihood) ? likelihood * 20 : 0)
    )
  );
  return {
    id: fallbackId || nextRiskRegisterId(store.getState().riskRegister.risks),
    title: String(suggestion.title || "Neues Risiko").trim() || "Neues Risiko",
    description: String(suggestion.description || "").trim(),
    phase: String(suggestion.phase || "").trim(),
    category: normalizeRiskCategoryValue(suggestion.category || riskCategoryOptions[4] || "Projekt- und Managementrisiken"),
    area: "",
    financialImpact: Math.max(0, Number(suggestion.financialImpact) || 0),
    probabilityPercent,
    expectedDamage: Math.max(0, (Number(suggestion.financialImpact) || 0) * (probabilityPercent / 100)),
    likelihood,
    impact: Math.max(1, Math.min(5, Number(suggestion.impact) || 3)),
    qualitativeRiskValue: Math.max(1, Math.min(25, likelihood * (Math.max(1, Math.min(5, Number(suggestion.impact) || 3))))),
    owner: String(suggestion.owner || "").trim(),
    measures: String(suggestion.measures || "").trim(),
    dueDate: String(suggestion.dueDate || "").trim(),
    status: "offen",
    residualRisk: normalizeResidualRiskText(suggestion.residualRisk || "")
  };
}

function updateRiskRegisterFromAiSuggestion(suggestion, mode = "append") {
  const state = store.getState();
  const items = Array.isArray(suggestion?.items) ? suggestion.items : [];
  if (!items.length) return false;
  if (mode === "append") {
    const nextRisks = items.map((entry, index) => mapSuggestedRiskToRegisterItem(entry, `R-${String(state.riskRegister.risks.length + index + 1).padStart(4, "0")}`));
    store.setState((current) => ({
      ...current,
      riskRegister: {
        ...current.riskRegister,
        risks: [...nextRisks, ...current.riskRegister.risks]
      }
    }));
    return true;
  }
  const selectedId = String(suggestion.riskId || suggestion.id || "").trim();
  if (!selectedId) return false;
  store.setState((current) => ({
    ...current,
    riskRegister: {
      ...current.riskRegister,
      risks: current.riskRegister.risks.map((risk) => risk.id === selectedId ? {
        ...risk,
        measures: String(suggestion.measures || risk.measures || ""),
        residualRisk: normalizeResidualRiskText(suggestion.residualRisk || risk.residualRisk || ""),
        owner: String(suggestion.owner || risk.owner || ""),
        dueDate: String(suggestion.dueDate || risk.dueDate || "")
      } : risk)
    }
  }));
  return true;
}

function applyAiWorkshopRisksToRegister() {
  const state = store.getState();
  const aiWorkshop = state.ui?.aiWorkshop || {};
  const suggestion = aiWorkshop.resultData || buildAiWorkshopFallbackResult(state, "free-text-risks", aiWorkshop.resultText);
  if (!suggestion?.items?.length) return false;
  pushRiskRegisterUndoSnapshot(state);
  const baseIndex = state.riskRegister.risks.length;
  const mappedRisks = suggestion.items.map((entry, index) => mapSuggestedRiskToRegisterItem(entry, `R-${String(baseIndex + index + 1).padStart(4, "0")}`));
  store.setState((state) => ({
    ...state,
    riskRegister: {
      ...state.riskRegister,
      risks: [...mappedRisks, ...state.riskRegister.risks]
    },
    ui: {
      ...state.ui,
      aiWorkshop: normalizeAiWorkshopState({
        ...state.ui.aiWorkshop,
        resultTitle: "Risiken übernommen",
        resultText: `${mappedRisks.length} Risiko${mappedRisks.length === 1 ? "" : "e"} wurden ins Register übernommen.`,
        resultTone: "success"
      })
    }
  }));
  return true;
}

function applyAiWorkshopMeasuresToRisk() {
  const state = store.getState();
  const aiWorkshop = state.ui?.aiWorkshop || {};
  const suggestion = aiWorkshop.resultData || buildAiWorkshopFallbackResult(state, "measures-residual", aiWorkshop.resultText);
  if (!suggestion?.items?.length) return false;
  const targetRiskId = String(aiWorkshop.selectedRiskId || "").trim();
  if (!targetRiskId) return false;
  const first = suggestion.items[0];
  pushRiskRegisterUndoSnapshot(state);
  store.setState((current) => ({
    ...current,
    riskRegister: {
      ...current.riskRegister,
      risks: current.riskRegister.risks.map((risk) => {
        if (risk.id !== targetRiskId) return risk;
        return {
          ...risk,
          owner: String(first.owner || risk.owner || "").trim(),
          measures: String(first.measures || risk.measures || "").trim(),
          residualRisk: String(first.residualRisk || risk.residualRisk || "").trim(),
          dueDate: String(first.dueDate || risk.dueDate || "").trim()
        };
      })
    },
    ui: {
      ...current.ui,
      aiWorkshop: normalizeAiWorkshopState({
        ...current.ui.aiWorkshop,
        resultTitle: "Maßnahmen übernommen",
        resultText: `Die KI-Vorschläge wurden in ${targetRiskId} übernommen. Bitte die Werte noch fachlich prüfen.`,
        resultTone: "success"
      })
    }
  }));
  return true;
}

function applyAiWorkshopFreeTextToRisk() {
  const state = store.getState();
  const aiWorkshop = state.ui?.aiWorkshop || {};
  const suggestion = aiWorkshop.resultData || buildAiWorkshopFallbackResult(state, "free-text-risks", aiWorkshop.resultText);
  if (!suggestion?.items?.length) return false;
  const targetRiskId = String(aiWorkshop.selectedRiskId || "").trim();
  if (!targetRiskId) return false;
  const first = suggestion.items[0];
  pushRiskRegisterUndoSnapshot(state);
  store.setState((current) => ({
    ...current,
    riskRegister: {
      ...current.riskRegister,
      risks: current.riskRegister.risks.map((risk) => {
        if (risk.id !== targetRiskId) return risk;
        const hasProbabilityPercent = first.probabilityPercent !== undefined && first.probabilityPercent !== null && String(first.probabilityPercent).trim() !== "";
        const nextLikelihood = hasProbabilityPercent
          ? deriveRiskLikelihoodFromPercent(first.probabilityPercent, first.likelihood || risk.likelihood || 1)
          : Math.max(1, Math.min(5, Number(first.likelihood) || risk.likelihood || 1));
        const nextImpact = Math.max(1, Math.min(5, Number(first.impact) || risk.impact || 3));
        const nextFinancialImpact = Math.max(0, Number(first.financialImpact) || risk.financialImpact || 0);
        const nextProbabilityPercent = hasProbabilityPercent
          ? Math.max(0, Math.min(100, Number(first.probabilityPercent)))
          : Math.max(0, Math.min(100, nextLikelihood * 20));
        return {
          ...risk,
          title: String(first.title || risk.title || "").trim(),
          description: String(first.description || risk.description || "").trim(),
          phase: String(first.phase || risk.phase || "").trim(),
          category: normalizeRiskCategoryValue(first.category || risk.category || riskCategoryOptions[4] || "Projekt- und Managementrisiken"),
          financialImpact: nextFinancialImpact,
          probabilityPercent: nextProbabilityPercent,
          expectedDamage: Math.max(0, nextFinancialImpact * (nextProbabilityPercent / 100)),
          likelihood: nextLikelihood,
          impact: nextImpact,
          qualitativeRiskValue: Math.max(1, Math.min(25, nextLikelihood * nextImpact)),
          owner: String(first.owner || risk.owner || "").trim(),
          measures: String(first.measures || risk.measures || "").trim(),
          dueDate: String(first.dueDate || risk.dueDate || "").trim(),
          residualRisk: normalizeResidualRiskText(first.residualRisk || risk.residualRisk || ""),
          status: risk.status || "offen"
        };
      })
    },
    ui: {
      ...current.ui,
      aiWorkshop: normalizeAiWorkshopState({
        ...current.ui.aiWorkshop,
        resultTitle: "Risiko übernommen",
        resultText: `Die KI-Vorschläge wurden in ${targetRiskId} übernommen. Bitte die Werte noch fachlich prüfen.`,
        resultTone: "success"
      })
    }
  }));
  return true;
}

function applyAiWorkshopSelectedResult() {
  const aiWorkshop = store.getState().ui?.aiWorkshop || {};
  if (aiWorkshop.activeTask === "measures-residual") {
    const targetRiskId = String(aiWorkshop.selectedRiskId || "").trim();
    if (!targetRiskId) return false;
    return applyAiWorkshopMeasuresToRisk();
  }
  const targetRiskId = String(aiWorkshop.selectedRiskId || "").trim();
  if (targetRiskId) {
    return applyAiWorkshopFreeTextToRisk();
  }
  return applyAiWorkshopRisksToRegister();
}

async function runAiWorkshopTask(task) {
  const currentState = store.getState();
  const hasKey = String(aiSettings.apiKey || "").trim().length > 0;
  if (!hasKey) {
    setAiWorkshopState({
      ...currentState.ui?.aiWorkshop,
      activeTask: task,
      busy: false,
      resultTitle: "API-Schlüssel fehlt",
      resultText: "Bitte zuerst den Anthropic API-Schlüssel speichern und die Verbindung prüfen.",
      resultTone: "danger"
    });
    return;
  }

  setAiWorkshopState({
    ...currentState.ui?.aiWorkshop,
    activeTask: task,
    busy: true,
    resultTitle: "KI verarbeitet ...",
    resultText: "Die gewählte Funktion wird gerade ausgeführt.",
    resultTone: "neutral"
  });

  try {
    const response = await fetch("http://127.0.0.1:8161/api/ai/generate", {
      method: "POST",
      headers: {
        "content-type": "application/json"
      },
      body: JSON.stringify({
        apiKey: aiSettings.apiKey,
        model: "claude-sonnet-4-20250514",
        maxTokens: task === "management-report" || task === "risk-report" ? 1600 : 1200,
        system: buildAiWorkshopSystemPrompt(task),
        userPrompt: buildAiWorkshopUserPrompt(currentState, task)
      })
    });

    if (!response.ok) {
      const errorPayload = await response.json().catch(() => ({}));
      const providerSnippet = String(errorPayload?.providerBody || "").trim();
      const extra = providerSnippet ? `\n${providerSnippet.slice(0, 500)}` : "";
      throw new Error(`${errorPayload?.error || `HTTP ${response.status}`}${extra}`);
    }

    const payload = await response.json();
    const parsedResult = narrowAiWorkshopMeasuresResult(
      sanitizeAiWorkshopResult(
        parseAiWorkshopResult(payload.text, task) || ((task === "free-text-risks" || task === "measures-residual") ? buildAiWorkshopFallbackResult(currentState, task, payload.text) : null),
        task
      ),
      currentState.ui?.aiWorkshop?.selectedRiskId || ""
    );
    setAiWorkshopState({
      ...currentState.ui?.aiWorkshop,
      activeTask: task,
      busy: false,
      resultTitle: {
        "management-report": "Managementbericht erzeugt",
        "risk-report": "Risikobericht erzeugt",
        "free-text-risks": "Freitext ausgewertet",
        "measures-residual": "Maßnahmen und Rest-Risiko erzeugt"
      }[task] || "KI-Ergebnis",
      resultText: String(payload.text || "").trim() || "Die KI hat keine verwertbare Ausgabe zurückgegeben.",
      resultTone: "success",
      resultData: parsedResult
    });
  } catch (error) {
    const fallbackResult = (task === "free-text-risks" || task === "measures-residual")
      ? narrowAiWorkshopMeasuresResult(
          sanitizeAiWorkshopResult(buildAiWorkshopFallbackResult(currentState, task, String(error?.message || "")), task),
          currentState.ui?.aiWorkshop?.selectedRiskId || ""
        )
      : null;
    setAiWorkshopState({
      ...currentState.ui?.aiWorkshop,
      activeTask: task,
      busy: false,
      resultTitle: fallbackResult ? "KI-Vorschlag vorbereitet" : "KI-Verarbeitung fehlgeschlagen",
      resultText: fallbackResult
        ? "Die KI-Ausgabe war nicht vollständig strukturiert. Ein prüfbarer Vorschlag wurde vorbereitet."
        : String(error?.message || "Unbekannter Fehler"),
      resultTone: fallbackResult ? "success" : "danger",
      resultData: fallbackResult
    });
  }
}

function createNeutralProjectState() {
  const next = cloneState();
  next.project = {
    name: "",
    type: "",
    bauart: "",
    phase: "",
    bgf: 0,
    budget: 0,
    costBasis: "netto",
    currency: "EUR",
    location: {
      street: "",
      houseNumber: "",
      postalCode: "",
      city: "",
      country: ""
    },
    client: "",
    projectLead: "",
    startDate: "",
    endDate: "",
    analysisDate: "",
    description: ""
  };
  next.reportProfile = {
    ...next.reportProfile,
    company: "",
    companyAddress: "",
    author: "",
    clientName: "",
    clientAddress: "",
    projectAddress: "",
    confidentiality: "Vertraulich",
    notes: "",
    logoDataUrl: null
  };
  next.pert = {
    ...next.pert,
    items: [],
    snapshots: [],
    lastResult: {
      expectedValue: 0,
      standardDeviation: 0,
      budget84: 0,
      budget975: 0
    }
  };
  next.monteCarlo = {
    ...next.monteCarlo,
    targetBudget: 0,
    targetDays: 0,
    workPackages: [],
    riskEvents: [],
    baselineScenario: null,
    storedScenarios: [],
    lastResult: {
      meanCost: 0,
      meanDays: 0,
      medianCost: 0,
      medianDays: 0,
      p10Cost: 0,
      p50Cost: 0,
      p80Cost: 0,
      p90Cost: 0,
      p10Days: 0,
      p50Days: 0,
      p80Days: 0,
      p90Days: 0,
      budgetSuccessRate: 0,
      scheduleSuccessRate: 0,
      combinedSuccessRate: 0
    }
  };
  next.earnedValue = {
    ...next.earnedValue,
    statusDate: "",
    activities: [],
    snapshots: [],
    lastResult: {
      pv: 0,
      ev: 0,
      ac: 0,
      cv: 0,
      sv: 0,
      cpi: 0,
      spi: 0,
      eac: 0,
      tcpi: 0
    }
  };
  next.riskRegister = {
    ...next.riskRegister,
    risks: []
  };
  next.ui = {
    ...next.ui,
    activeModule: "project",
    projectExportName: "",
    reportExportName: "",
    reportExportFormat: "txt",
    reportMode: "risk",
    riskRegisterView: {
      search: "",
      status: "alle",
      owner: "alle",
      category: "alle",
      criticalOnly: false,
      visibleColumns: ["select", "priority", "status", "value", "category", "phase", "impact", "owner", "dueDate", "measures"],
      topLimit: 5,
      sortBy: "priority",
      dueFrom: "",
      dueTo: "",
      matrixSelection: null
    },
    evaActivitiesPanelOpen: false,
    evaActivityOpenIds: [],
    reportOptions: {
      includePert: true,
      includeMonteCarlo: true,
      includeEarnedValue: true,
      includeRiskRegister: true
    },
    aiWorkshop: normalizeAiWorkshopState(),
    dirty: false
  };
  return next;
}

function clampEvaTolerancePercent(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 3;
  return Math.max(0, Math.min(10, numeric));
}

function normalizeEarnedValueThresholds(thresholds = {}) {
  return {
    ...thresholds,
    spiGreen: normalizeEvaThresholdValue("spiGreen", thresholds.spiGreen ?? 1),
    spiYellow: normalizeEvaThresholdValue("spiYellow", thresholds.spiYellow ?? 0.95),
    cpiGreen: normalizeEvaThresholdValue("cpiGreen", thresholds.cpiGreen ?? 1),
    cpiYellow: normalizeEvaThresholdValue("cpiYellow", thresholds.cpiYellow ?? 0.95),
    cvTolerancePercent: clampEvaTolerancePercent(thresholds.cvTolerancePercent),
    cvTolerance: normalizeEvaThresholdValue("eacTolerance", thresholds.cvTolerance ?? 180000),
    eacTolerance: normalizeEvaThresholdValue("eacTolerance", thresholds.eacTolerance ?? 180000),
    tcpiActionLevel: normalizeEvaThresholdValue("tcpiActionLevel", thresholds.tcpiActionLevel ?? 1.02)
  };
}

function normalizeEvaActivityOpenIds(openIds, activities = []) {
  const activityIds = new Set((activities || []).map((activity) => String(activity?.id || "")).filter(Boolean));
  const fromIds = Array.isArray(openIds)
    ? openIds.map((id) => String(id || "")).filter((id) => activityIds.has(id))
    : [];
  if (fromIds.length) {
    return [...new Set(fromIds)];
  }
  const firstId = activities[0]?.id ? String(activities[0].id) : "";
  return firstId ? [firstId] : [];
}

function normalizeEvaUiState(ui = {}, activities = []) {
  return {
    ...ui,
    activeModule: ui.activeModule || "project",
    evaDetailView: ui.evaDetailView === "full" ? "full" : "compact",
    evaActivityOpenIds: normalizeEvaActivityOpenIds(
      ui.evaActivityOpenIds ?? ui.evaActivityOpenIndexes,
      activities
    ),
    evaDetailColumns: {
      pv: ui.evaDetailColumns?.pv !== false,
      ac: ui.evaDetailColumns?.ac !== false,
      ev: ui.evaDetailColumns?.ev !== false,
      cpi: ui.evaDetailColumns?.cpi !== false,
      spi: ui.evaDetailColumns?.spi !== false,
      fgr: ui.evaDetailColumns?.fgr !== false,
      cv: ui.evaDetailColumns?.cv !== false,
      sv: ui.evaDetailColumns?.sv !== false,
      eac: ui.evaDetailColumns?.eac !== false,
      status: ui.evaDetailColumns?.status !== false
    },
    evaActivitiesPanelOpen: ui.evaActivitiesPanelOpen === true,
    projectExportName: typeof ui.projectExportName === "string" ? ui.projectExportName : "",
    reportExportName: typeof ui.reportExportName === "string" ? ui.reportExportName : "",
    reportExportFormat: ["json", "txt", "doc", "pdf"].includes(String(ui.reportExportFormat || "").toLowerCase())
      ? String(ui.reportExportFormat).toLowerCase()
      : "txt",
    riskRegisterView: {
      search: ui.riskRegisterView?.search || "",
      status: ui.riskRegisterView?.status || "alle",
      owner: ui.riskRegisterView?.owner || "alle",
      category: ui.riskRegisterView?.category || "alle",
      criticalOnly: ui.riskRegisterView?.criticalOnly === true,
      visibleColumns: normalizeRiskRegisterColumns(ui.riskRegisterView?.visibleColumns),
      topLimit: normalizeRiskRegisterTopLimit(ui.riskRegisterView?.topLimit),
      sortBy: ["priority", "value", "id", "dueDate", "category"].includes(ui.riskRegisterView?.sortBy)
        ? ui.riskRegisterView.sortBy
        : "priority",
      editSortBy: ["newest", "id"].includes(ui.riskRegisterView?.editSortBy)
        ? ui.riskRegisterView.editSortBy
        : "newest",
      dueFrom: ui.riskRegisterView?.dueFrom || "",
      dueTo: ui.riskRegisterView?.dueTo || "",
      matrixSelection: normalizeRiskRegisterMatrixSelection(ui.riskRegisterView?.matrixSelection)
    },
    riskRegisterUndoStack: normalizeRiskRegisterUndoStack(ui.riskRegisterUndoStack),
    reportMode: ui.reportMode === "management" ? "management" : "risk",
    aiWorkshop: normalizeAiWorkshopState(ui.aiWorkshop),
    dirty: false
  };
}

function normalizeRiskRegisterTopLimit(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric === 5 || numeric === 10 || numeric === 20 || numeric === 0) return numeric || 5;
  return 5;
}

function normalizeRiskRegisterColumns(visibleColumns) {
  const defaultColumns = ["select", "priority", "status", "value", "category", "phase", "impact", "owner", "dueDate", "measures"];
  if (!Array.isArray(visibleColumns)) return defaultColumns;
  const filtered = visibleColumns.map((value) => String(value || "")).filter(Boolean);
  return filtered.length ? [...new Set(filtered)] : defaultColumns;
}

function normalizeRiskRegisterMatrixSelection(matrixSelection) {
  if (!matrixSelection || typeof matrixSelection !== "object") return null;
  const likelihood = Number(matrixSelection.likelihood);
  const impact = Number(matrixSelection.impact);
  if (!Number.isInteger(likelihood) || !Number.isInteger(impact)) return null;
  if (likelihood < 1 || likelihood > 5 || impact < 1 || impact > 5) return null;
  return { likelihood, impact };
}

function normalizeImportedProjectState(parsed) {
  if (!isValidProjectState(parsed)) return null;
  const importEditSort = String(parsed.ui?.riskRegisterView?.editSortBy || "newest").toLowerCase() === "id"
    ? "id"
    : "newest";
  const riskRegister = {
    ...parsed.riskRegister,
    risks: Array.isArray(parsed.riskRegister?.risks) && parsed.riskRegister.risks.length
      ? sortRiskRegisterEntriesForEdit(normalizeRiskRegisterEntries(parsed.riskRegister.risks), importEditSort)
      : cloneState().riskRegister.risks
  };
  return {
    ...parsed,
    earnedValue: {
      ...parsed.earnedValue,
      thresholds: normalizeEarnedValueThresholds(parsed.earnedValue?.thresholds)
    },
    riskRegister,
    ui: normalizeEvaUiState(parsed.ui, parsed.earnedValue?.activities || [])
  };
}

function formatTolerancePercent(value) {
  return `${new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  }).format(value)} %`;
}

function formatToleranceCurrency(value) {
  return new Intl.NumberFormat("de-DE", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 0
  }).format(value || 0);
}

function updateEvaTolerancePreview(percent) {
  const clamped = clampEvaTolerancePercent(percent);
  const budget = Number(store.getState().project?.budget) || 0;
  const amount = budget > 0 ? (budget * clamped) / 100 : 0;
  const percentTarget = document.getElementById("eva_tolerance_percent_display");
  const amountTarget = document.getElementById("eva_tolerance_amount_display");
  if (percentTarget) percentTarget.textContent = formatTolerancePercent(clamped);
  if (amountTarget) amountTarget.textContent = `= ${formatToleranceCurrency(amount)}`;
}

function layoutEvaToleranceAxis() {
  const slider = document.getElementById("eva_cv_tolerance_percent");
  const ticks = document.querySelectorAll(".eva-tolerance-tick");
  const labels = document.querySelectorAll(".eva-tolerance-scale span");
  if (!slider || !ticks.length || !labels.length) return;

  const styles = getComputedStyle(document.querySelector(".eva-tolerance-card"));
  const thumbSize = parseFloat(styles.getPropertyValue("--eva-tolerance-thumb-size")) || 22;
  const halfThumb = thumbSize / 2;
  const usableWidth = Math.max(0, slider.clientWidth - thumbSize);

  ticks.forEach((tick, index) => {
    const left = halfThumb + usableWidth * (index / 10);
    tick.style.left = `${left}px`;
  });

  labels.forEach((label, index) => {
    const left = halfThumb + usableWidth * (index / 10);
    label.style.left = `${left}px`;
  });
}

function isValidProjectState(payload) {
  return Boolean(
    payload &&
    typeof payload === "object" &&
    payload.project &&
    payload.reportProfile &&
    payload.pert &&
    payload.monteCarlo &&
    payload.earnedValue &&
    payload.riskRegister &&
    payload.ui
  );
}

function persistAutosave(state) {
  if (suppressNextAutosave > 0) {
    suppressNextAutosave -= 1;
    return;
  }
  try {
    localStorage.setItem(AUTOSAVE_KEY, serializeProject(state));
    lastAutosavedAt = new Date().toISOString();
    updateStorageStatus(`Autosave aktiv · ${formatTimestamp(lastAutosavedAt)}`);
  } catch (_error) {
    updateStorageStatus("Autosave konnte im Browser nicht gespeichert werden.");
  }
}

function restoreAutosave() {
  try {
    const raw = localStorage.getItem(AUTOSAVE_KEY);
    if (!raw) {
      updateStorageStatus("Autosave bereit.");
      return;
    }
    const parsed = JSON.parse(raw);
    if (!isValidProjectState(parsed)) {
      updateStorageStatus("Vorhandener Autosave war unvollständig und wurde ignoriert.");
      return;
    }
    const nextState = normalizeImportedProjectState(parsed);
    if (!nextState) {
      updateStorageStatus("Vorhandener Autosave war unvollständig und wurde ignoriert.");
      return;
    }
    nextState.riskRegister = {
      ...nextState.riskRegister,
      risks: sortRiskRegisterEntriesForEdit(
        normalizeRiskRegisterEntries(nextState.riskRegister?.risks),
        nextState.ui?.riskRegisterView?.editSortBy || "newest"
      )
    };
    store.setState(nextState);
    store.markSaved();
    lastAutosavedAt = nextState.meta?.savedAt || new Date().toISOString();
    updateStorageStatus(`Autosave geladen · ${formatTimestamp(lastAutosavedAt)}`);
  } catch (_error) {
    updateStorageStatus("Autosave konnte nicht geladen werden.");
  }
}

function importProjectFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || ""));
      if (!isValidProjectState(parsed)) {
        updateStorageStatus("Die gewählte Datei ist keine gültige Projektdatei.");
        return;
      }
      const nextState = normalizeImportedProjectState(parsed);
      if (!nextState) {
        updateStorageStatus("Die gewählte Datei ist keine gültige Projektdatei.");
        return;
      }
      store.setState(nextState);
      store.markSaved();
      persistAutosave(store.getState());
      updateStorageStatus(`Projektdatei geladen · ${file.name}`);
    } catch (_error) {
      updateStorageStatus("Die Projektdatei konnte nicht gelesen werden.");
    }
  };
  reader.onerror = () => {
    updateStorageStatus("Die Projektdatei konnte nicht gelesen werden.");
  };
  reader.readAsText(file, "utf-8");
}

function serializeProject(state) {
  return JSON.stringify(state, null, 2);
}

function saveProjectFile(state) {
  const safeName = sanitizeExportFileName(state.ui?.projectExportName || state.project.name || "project-controls-hub", state.project.name || "project-controls-hub");
  const blob = new Blob([serializeProject(state)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeName}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function sanitizeExportFileName(value, fallback) {
  const raw = String(value || "").trim().replace(/\s+/g, " ");
  const withoutExtension = raw.replace(/\.(json|doc|pdf)$/i, "");
  const safe = withoutExtension
    .replace(/[^\wäöüÄÖÜß.-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  return safe || fallback;
}

function getActiveReportMode(state) {
  return state.ui?.reportMode === "management" ? "management" : "risk";
}

function getActiveReportTitle(state) {
  return getActiveReportMode(state) === "management" ? "Managementbericht" : "Risikobericht";
}

function getDefaultReportExportName(state) {
  return getActiveReportTitle(state);
}

function getReportExportBaseName(state) {
  const fallback = getActiveReportTitle(state);
  return sanitizeExportFileName(state.ui?.reportExportName, fallback);
}

function getReportExportFormat(state) {
  const format = String(state.ui?.reportExportFormat || "").toLowerCase();
  if (["json", "txt", "doc", "pdf"].includes(format)) return format;
  return "txt";
}

function getReportExportSelection(state) {
  const fileNameInput = document.getElementById("reportExportFileName");
  const formatSelect = document.getElementById("reportExportFormat");
  return {
    fileName: fileNameInput?.value ?? state.ui?.reportExportName ?? "",
    format: (formatSelect?.value || state.ui?.reportExportFormat || "txt").toLowerCase()
  };
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function buildPrintableReportHtml(title, subtitle, reportText, fileName) {
  return `<!DOCTYPE html>
<html lang="de">
  <head>
    <meta charset="UTF-8">
    <title>${escapeHtml(fileName || title)}</title>
    <style>
      body { font-family: "Source Sans 3", "Segoe UI", sans-serif; margin: 40px; color: #08131d; }
      h1 { font-family: "DM Serif Display", Georgia, serif; font-size: 30px; margin: 0 0 8px; color: #091f33; }
      .meta { color: #425466; margin-bottom: 28px; line-height: 1.6; }
      .copy { white-space: pre-wrap; line-height: 1.7; font-size: 15px; }
    </style>
  </head>
  <body>
    <h1>${escapeHtml(title)}</h1>
    <div class="meta">${escapeHtml(subtitle)}</div>
    <div class="copy">${escapeHtml(reportText)}</div>
  </body>
</html>`;
}

function openPrintableReport(state, reportTitle, reportText, fileName) {
  const printWindow = window.open("", "_blank", "width=960,height=1200");
  if (!printWindow) {
    updateStorageStatus("Druckansicht konnte nicht geöffnet werden.");
    return false;
  }

  const subtitle = `Projekt: ${state.project?.name || "nicht angegeben"} · Berichtsdatum: ${formatTimestamp(state.project?.analysisDate)} · Format: PDF`;
  printWindow.document.write(buildPrintableReportHtml(reportTitle, subtitle, reportText, fileName));
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
  return true;
}

function exportSelectedReportByFormat(state) {
  const reportMode = getActiveReportMode(state);
  const reportTitle = getActiveReportTitle(state);
  const selection = getReportExportSelection(state);
  const baseName = sanitizeExportFileName(selection.fileName, reportTitle);
  const format = ["json", "txt", "doc", "pdf"].includes(selection.format) ? selection.format : getReportExportFormat(state);
  const reportText = reportMode === "management" ? renderSelectedReportText(state) : renderRiskReportText(state);

  if (format === "json") {
    const payload = {
      reportMode,
      reportTitle,
      generatedAt: new Date().toISOString(),
      reportText,
      ...buildSelectedReportData(state)
    };
    downloadFile(
      `${baseName}.json`,
      "application/json;charset=utf-8",
      JSON.stringify(payload, null, 2)
    );
    return;
  }

  if (format === "txt") {
    downloadFile(
      `${baseName}.txt`,
      "text/plain;charset=utf-8",
      reportText
    );
    return;
  }

  if (format === "doc") {
    downloadFile(
      `${baseName}.doc`,
      "application/msword;charset=utf-8",
      buildPrintableReportHtml(reportTitle, `Projekt: ${state.project?.name || "nicht angegeben"} · Berichtsdatum: ${formatTimestamp(state.project?.analysisDate)} · Format: DOC`, reportText, baseName)
    );
    return;
  }

  openPrintableReport(state, reportTitle, reportText, `${baseName}.pdf`);
}

function downloadFile(name, type, content) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportManagementJson(state) {
  const payload = buildSelectedReportData(state);
  const safeName = (state.project.name || "management-report")
    .replace(/[^\wäöüÄÖÜß.-]+/g, "_")
    .replace(/_+/g, "_");
  downloadFile(
    `${safeName}_selected-report.json`,
    "application/json;charset=utf-8",
    JSON.stringify(payload, null, 2)
  );
}

function printManagementReport(state) {
  const reportMode = getActiveReportMode(state);
  const reportTitle = reportMode === "management" ? "Managementbericht" : "Risikobericht";
  const reportText = reportMode === "management" ? renderSelectedReportText(state) : renderRiskReportText(state);
  openPrintableReport(
    state,
    `${reportTitle} · ${state.project?.name || "Projekt"}`,
    reportText,
    getReportExportBaseName(state)
  );
}

function renderProjectMeta(state) {
  const target = document.getElementById("projectMeta");
  if (!target) return;
  target.innerHTML = `
    <div class="meta-line"><strong>${state.project.name}</strong></div>
    <div class="meta-line">${state.project.type} · ${state.project.bauart}</div>
    <div class="meta-line">${state.project.location.city} · ${state.project.phase}</div>
    <div class="meta-line">Budget: <strong>${new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(state.project.budget)}</strong></div>
    <div class="meta-line">Dateistatus: <strong>${state.ui.dirty ? "ungesichert" : "gesichert"}</strong></div>
  `;
  const projectExportField = document.getElementById("projectExportFileName");
  if (projectExportField && document.activeElement !== projectExportField) {
    projectExportField.value = state.ui?.projectExportName || state.project.name || "project-controls-hub";
  }
}

function renderNav(state) {
  const target = document.getElementById("moduleNav");
  if (!target) return;
  target.innerHTML = Object.values(modules).map((module) => `
    <button class="nav-btn ${state.ui.activeModule === module.key ? "active" : ""}" data-module="${module.key}" type="button">
      ${module.label}
    </button>
  `).join("");
}

function renderView(state) {
  const target = document.getElementById("appView");
  const active = modules[state.ui.activeModule];
  if (!target || !active) return;
  target.innerHTML = active.render(state);
}

function captureRiskFieldFocus() {
  const activeElement = document.activeElement;
  if (!activeElement || typeof activeElement.matches !== "function") return null;
  if (!activeElement.matches("[data-risk-ui-field]")) return null;
  return {
    field: activeElement.getAttribute("data-risk-ui-field"),
    selectionStart: typeof activeElement.selectionStart === "number" ? activeElement.selectionStart : null,
    selectionEnd: typeof activeElement.selectionEnd === "number" ? activeElement.selectionEnd : null,
    value: activeElement.value
  };
}

function restoreRiskFieldFocus(snapshot) {
  if (!snapshot?.field) return;
  const selector = `[data-risk-ui-field="${snapshot.field}"]`;
  const target = document.querySelector(selector);
  if (!target || typeof target.focus !== "function") return;
  target.focus({ preventScroll: true });
  if (typeof target.setSelectionRange === "function" && snapshot.selectionStart !== null && snapshot.selectionEnd !== null) {
    const valueLength = String(target.value || "").length;
    const start = Math.min(snapshot.selectionStart, valueLength);
    const end = Math.min(snapshot.selectionEnd, valueLength);
    try {
      target.setSelectionRange(start, end);
    } catch (_error) {
      // Some controls do not support selection ranges; ignore silently.
    }
  }
}

function renderApp(state) {
  const preserveScroll = state.ui?.activeModule === "riskRegister";
  const preserveRiskFocus = captureRiskFieldFocus();
  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  renderNav(state);
  renderProjectMeta(state);
  renderAiSettingsPanel();
  renderView(state);
  const evaToleranceSlider = document.getElementById("eva_cv_tolerance_percent");
  if (evaToleranceSlider) {
    const clamped = clampEvaTolerancePercent(state.earnedValue?.thresholds?.cvTolerancePercent);
    evaToleranceSlider.value = clamped.toFixed(1);
    updateEvaTolerancePreview(clamped);
    layoutEvaToleranceAxis();
    evaToleranceSlider.oninput = () => {
      updateEvaTolerancePreview(clampEvaTolerancePercent(evaToleranceSlider.valueAsNumber));
    };
    evaToleranceSlider.onchange = () => {
      const parsed = clampEvaTolerancePercent(evaToleranceSlider.valueAsNumber);
      evaToleranceSlider.value = parsed.toFixed(1);
      updateEvaTolerancePreview(parsed);
      store.setState((currentState) => ({
        ...currentState,
        earnedValue: {
          ...currentState.earnedValue,
          thresholds: {
            ...currentState.earnedValue.thresholds,
            cvTolerancePercent: parsed
          }
        }
      }));
    };
  }
  if (lastAutosavedAt) {
    updateStorageStatus(`Autosave aktiv · ${formatTimestamp(lastAutosavedAt)}`);
  }
  if (preserveScroll) {
    window.requestAnimationFrame(() => {
      window.scrollTo(scrollX, scrollY);
      restoreRiskFieldFocus(preserveRiskFocus);
    });
  }
}

function parseValue(currentValue, rawValue) {
  if (typeof currentValue === "number") {
    const normalized = String(rawValue ?? "")
      .trim()
      .replace(/\./g, "")
      .replace(",", ".")
      .replace(/[^\d.-]/g, "");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : currentValue;
  }
  return rawValue;
}

function parseFlexibleNumber(rawValue) {
  const normalized = String(rawValue ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getEvaThresholdConfig(field) {
  if (field === "eacTolerance") {
    return { min: 0, max: 999999999, step: 1000, kind: "integer" };
  }
  if (field === "spiGreen" || field === "spiYellow" || field === "cpiGreen" || field === "cpiYellow") {
    return { min: 0, max: 1.2, step: 0.01, kind: "decimal" };
  }
  if (field === "tcpiActionLevel") {
    return { min: 0, max: 1.5, step: 0.01, kind: "decimal" };
  }
  return { min: 0, max: 999999999, step: 1, kind: "integer" };
}

function normalizeEvaThresholdValue(field, rawValue) {
  const config = getEvaThresholdConfig(field);
  const parsed = parseFlexibleNumber(rawValue);
  const clamped = Math.max(config.min, Math.min(config.max, parsed));
  return config.kind === "integer" ? Math.round(clamped) : Number(clamped.toFixed(2));
}

function createPertItem() {
  return {
    id: `pert-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: "Neue Kostenposition",
    optimistic: 0,
    mostLikely: 0,
    pessimistic: 0,
    comment: ""
  };
}

function createEvaActivity() {
  return {
    id: `eva-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: "Neue Aktivität",
    plannedValue: 0,
    earnedValue: 0,
    actualCost: 0,
    progressPercent: 0
  };
}

function createEvaSnapshot(state) {
  const result = calculateEarnedValueResult(state);
  const now = new Date();
  return {
    id: `eva-snapshot-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    label: `${state.project?.phase || "LPH"} · ${state.earnedValue?.statusDate || now.toISOString().slice(0, 10)}`,
    savedAt: now.toLocaleString("de-DE"),
    statusDate: state.earnedValue?.statusDate || now.toISOString().slice(0, 10),
    pv: result.pv,
    ev: result.ev,
    ac: result.ac,
    cv: result.cv,
    sv: result.sv,
    cpi: result.cpi,
    spi: result.spi,
    eac: result.eac,
    etc: result.etc,
    vac: result.vac,
    tcpi: result.tcpi,
    bac: Number(state.project?.budget) || 0,
    fgr: result.progressPercent
  };
}

function createMonteCarloWorkPackage() {
  return {
    id: `mc-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: "Neues Work Package",
    optimisticCost: 0,
    mostLikelyCost: 0,
    pessimisticCost: 0,
    optimisticDays: 0,
    mostLikelyDays: 0,
    pessimisticDays: 0
  };
}

function createMonteCarloRisk() {
  return {
    id: `mc-risk-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: "Neues Risiko",
    probability: 0.1,
    minCostImpact: 0,
    maxCostImpact: 0,
    minDelayDays: 0,
    maxDelayDays: 0
  };
}

function createMonteCarloScenario(state, type = "scenario") {
  const now = new Date();
  const labelPrefix = type === "baseline" ? "Baseline" : "Szenario";
  const summary = calculateMonteCarloResult(state);
  return {
    id: `mc-scenario-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    label: `${labelPrefix} · ${state.project?.phase || "LPH"} · ${now.toLocaleDateString("de-DE")}`,
    savedAt: now.toLocaleString("de-DE"),
    projectPhase: state.project?.phase || "",
    targetBudget: state.monteCarlo.targetBudget,
    targetDays: state.monteCarlo.targetDays,
    iterations: state.monteCarlo.iterations,
    seed: state.monteCarlo.seed,
    summary: {
      meanCost: summary.meanCost,
      meanDays: summary.meanDays,
      p50Cost: summary.p50Cost,
      p80Cost: summary.p80Cost,
      p50Days: summary.p50Days,
      p80Days: summary.p80Days,
      combinedSuccessRate: summary.combinedSuccessRate
    },
    measureSettings: JSON.parse(JSON.stringify(state.monteCarlo.measureSettings || {})),
    workPackages: JSON.parse(JSON.stringify(state.monteCarlo.workPackages || [])),
    riskEvents: JSON.parse(JSON.stringify(state.monteCarlo.riskEvents || []))
  };
}

function createRiskRegisterItem() {
  return {
    id: "R-0001",
    createdAt: new Date().toISOString(),
    title: "Neues Risiko",
    description: "",
    phase: "",
    category: normalizeRiskCategoryValue(riskCategoryOptions[4] || "Projekt- und Managementrisiken"),
    area: "",
    financialImpact: 0,
    probabilityPercent: 0,
    expectedDamage: 0,
    likelihood: 1,
    impact: 3,
    qualitativeRiskValue: 3,
    owner: "",
    measures: "",
    dueDate: "",
    status: "offen",
    residualRisk: ""
  };
}

function normalizeRiskRegisterEntries(risks = []) {
  const baseTime = Date.now();
  return (Array.isArray(risks) ? risks : []).map((risk, index) => ({
    ...risk,
    createdAt: typeof risk?.createdAt === "string" && risk.createdAt.trim()
      ? risk.createdAt
      : new Date(baseTime - ((risks.length - index) * 1000)).toISOString()
  }));
}

function sortRiskRegisterEntriesForEdit(risks = [], sortBy = "newest") {
  const items = [...normalizeRiskRegisterEntries(risks)];
  const riskNumber = (risk) => {
    const match = String(risk.id || "").match(/^R-(\d{4})$/i);
    return match ? Number(match[1]) || 0 : 0;
  };
  if (sortBy === "id") {
    return items.sort((a, b) => riskNumber(a) - riskNumber(b) || String(a.id || "").localeCompare(String(b.id || ""), "de"));
  }
  return items.sort((a, b) => {
    const aNo = riskNumber(a);
    const bNo = riskNumber(b);
    if (aNo !== bNo) return bNo - aNo;
    const aTs = Date.parse(a.createdAt || "");
    const bTs = Date.parse(b.createdAt || "");
    if (Number.isFinite(aTs) && Number.isFinite(bTs) && aTs !== bTs) return bTs - aTs;
    return String(b.id || "").localeCompare(String(a.id || ""), "de");
  });
}

function nextRiskRegisterId(risks = []) {
  const maxNumber = risks.reduce((max, risk) => {
    const match = String(risk.id || "").match(/^R-(\d{4})$/i);
    if (!match) return max;
    return Math.max(max, Number(match[1]) || 0);
  }, 0);
  return `R-${String(maxNumber + 1).padStart(4, "0")}`;
}

function createMonteCarloWorkPackageFromPertItem(item) {
  return {
    id: `mc-from-pert-${item.id || Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: item.name || "PERT-Position",
    optimisticCost: Number(item.optimistic) || 0,
    mostLikelyCost: Number(item.mostLikely) || 0,
    pessimisticCost: Number(item.pessimistic) || 0,
    optimisticDays: 0,
    mostLikelyDays: 0,
    pessimisticDays: 0
  };
}

function createMonteCarloRiskFromRegister(risk) {
  const probability = Math.max(0, Math.min(1, (Number(risk.probabilityPercent) || 0) / 100));
  const financialImpact = Number(risk.financialImpact) || 0;
  const likelihood = deriveRiskLikelihoodFromPercent(risk.probabilityPercent, risk.likelihood || 1);
  return {
    id: `mc-from-risk-${risk.id || Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: risk.title || risk.id || "Register-Risiko",
    probability,
    minCostImpact: Math.round(financialImpact * 0.5),
    maxCostImpact: financialImpact,
    minDelayDays: Math.max(0, likelihood || 0),
    maxDelayDays: Math.max(0, (Number(risk.impact) || 0) * 3)
  };
}

function createPertSnapshot(state) {
  const result = calculatePertResult(state);
  const securityLabels = {
    0: "50 % · Erwartungswert",
    1: "84 % · Vorsichtiges Niveau",
    2: "97,5 % · Konservatives Niveau"
  };
  const projectName = state.project?.name || "Projekt";
  const phase = state.project?.phase || "LPH 1";
  const securityLabel = securityLabels[Number(state.pert.securityLevel)] || "84 % · Vorsichtiges Niveau";
  return {
    id: `pert-snapshot-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    label: `${projectName} · ${phase} · ${securityLabel}`,
    savedAt: new Date().toLocaleString("de-DE"),
    phase,
    projectName,
    securityLevel: state.pert.securityLevel,
    securityLabel,
    optimisticTotal: result.optimisticTotal,
    mostLikelyTotal: result.mostLikelyTotal,
    pessimisticTotal: result.pessimisticTotal,
    expectedValue: result.expectedValue,
    standardDeviation: result.standardDeviation,
    selectedBudget: result.selectedBudget,
    items: JSON.parse(JSON.stringify(state.pert.items)),
    lphComments: JSON.parse(JSON.stringify(state.pert.lphComments || {}))
  };
}

function toggleSelection(list, value, checked) {
  const next = new Set(list || []);
  if (checked) next.add(value);
  else next.delete(value);
  return [...next];
}

function bindEvents() {
  document.addEventListener("toggle", (event) => {
    const details = event.target;
    if (!(details instanceof HTMLElement)) return;
    if (details.matches(".risk-edit-fold")) {
      if (details.open) {
        document.querySelectorAll(".risk-edit-list .risk-edit-fold[open]").forEach((item) => {
          if (item !== details) {
            item.open = false;
          }
        });
      }
      return;
    }
    if (!details.matches(".eva-card[data-eva-activity-id], #eva-activities-panel")) return;
    if (details.id === "eva-activities-panel") {
      const current = Boolean(store.getState().ui?.evaActivitiesPanelOpen);
      if (current === details.open) return;
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          evaActivitiesPanelOpen: details.open
        }
      }));
      return;
    }
    const id = details.getAttribute("data-eva-activity-id");
    if (!id) return;
    const current = new Set(store.getState().ui?.evaActivityOpenIds || []);
    const next = new Set(current);
    if (details.open) next.add(id);
    else next.delete(id);
    const changed = current.size !== next.size || [...current].some((value) => !next.has(value));
    if (!changed) return;
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        evaActivityOpenIds: [...next]
      }
    }));
  }, true);

  document.addEventListener("click", (event) => {
    const moduleButton = event.target.closest("[data-module]");
    if (moduleButton) {
      const moduleKey = moduleButton.getAttribute("data-module");
      store.setState((state) => ({
        ...state,
        ui: { ...state.ui, activeModule: moduleKey }
      }));
      return;
    }

    if (event.target.id === "saveProjectBtn") {
      saveProjectFile(store.getState());
      store.markSaved();
      persistAutosave(store.getState());
      updateStorageStatus(`Projektdatei exportiert · ${formatTimestamp(store.getState().ui.lastSavedAt)}`);
      return;
    }

    if (event.target.id === "loadProjectBtn") {
      document.getElementById("projectFileInput")?.click();
      return;
    }

    if (event.target.id === "loadDemoBtn") {
      const next = cloneState();
      next.ui.activeModule = "project";
      store.setState(next);
      store.markSaved();
      persistAutosave(store.getState());
      updateStorageStatus("Demo-Datei geladen und als Autosave gespeichert.");
      return;
    }

    if (event.target.id === "resetProjectBtn") {
      const resetProject = window.confirm("Projektstand zurücksetzen?");
      if (!resetProject) return;
      const next = createNeutralProjectState();
      store.setState(next);
      store.markSaved();
      updateStorageStatus("Projektstand zurückgesetzt. Alle Felder sind geleert.");
      return;
    }

    if (event.target.id === "saveAiSettingsBtn") {
      const nextAiSettings = normalizeAiSettings({
        ...readAiSettingsFromPanel(),
        connected: aiSettings.connected,
        testing: false,
        lastStatus: aiSettings.lastStatus
      });
      nextAiSettings.lastSavedAt = new Date().toISOString();
      nextAiSettings.lastStatus = "API-Schlüssel gespeichert";
      applyAiSettings(nextAiSettings, nextAiSettings.lastStatus);
      return;
    }

    if (event.target.id === "testAiSettingsBtn") {
      startAiConnectionTest();
      return;
    }

    if (event.target.id === "disconnectAiSettingsBtn") {
      disconnectAiConnection();
      return;
    }

    const actionTarget = event.target.closest("[data-action]");
    if (actionTarget && (actionTarget.dataset.action === "remove-mc-work" || actionTarget.dataset.action === "remove-mc-risk")) {
      event.preventDefault();
      event.stopPropagation();
    }
    if (actionTarget?.dataset.action === "export-management-json") {
      exportManagementJson(store.getState());
      updateStorageStatus("Management-Zusammenfassung exportiert.");
      return;
    }

    if (actionTarget?.dataset.action === "export-project-json") {
      saveProjectFile(store.getState());
      store.markSaved();
      persistAutosave(store.getState());
      updateStorageStatus(`Projektdatei exportiert · ${formatTimestamp(store.getState().ui.lastSavedAt)}`);
      return;
    }

    if (actionTarget?.dataset.action === "export-selected-report") {
      exportSelectedReportByFormat(store.getState());
      updateStorageStatus("Bericht exportiert.");
      return;
    }

    if (actionTarget?.dataset.action === "run-ai-workshop") {
      runAiWorkshopTask(String(actionTarget.dataset.aiWorkshopTask || "management-report"));
      return;
    }

    if (actionTarget?.dataset.action === "apply-ai-workshop-result") {
      const aiWorkshop = store.getState().ui?.aiWorkshop || {};
      if (aiWorkshop.activeTask === "measures-residual" && !String(aiWorkshop.selectedRiskId || "").trim()) {
        updateStorageStatus("Bitte zuerst ein Zielrisiko wählen.");
        return;
      }
      if (!applyAiWorkshopSelectedResult()) {
        updateStorageStatus(aiWorkshop.activeTask === "measures-residual"
          ? "Die KI-Vorschläge konnten nicht übernommen werden."
          : "Keine KI-Vorschläge zum Übernehmen vorhanden.");
        return;
      }
      updateStorageStatus(aiWorkshop.activeTask === "measures-residual"
        ? "KI-Maßnahmen im Zielrisiko übernommen."
        : "KI-Vorschläge als neue Risiken übernommen.");
      return;
    }

    if (actionTarget?.dataset.action === "apply-ai-workshop-risks") {
      if (!applyAiWorkshopRisksToRegister()) {
        updateStorageStatus("Keine KI-Vorschläge zum Übernehmen vorhanden.");
      } else {
        updateStorageStatus("KI-Vorschläge als neue Risiken übernommen.");
      }
      return;
    }

    if (actionTarget?.dataset.action === "apply-ai-workshop-measures") {
      if (!applyAiWorkshopMeasuresToRisk()) {
        updateStorageStatus("Bitte zuerst ein Zielrisiko wählen und einen KI-Vorschlag erzeugen.");
      } else {
        updateStorageStatus("KI-Maßnahmen im Zielrisiko übernommen.");
      }
      return;
    }

    if (actionTarget?.dataset.action === "clear-ai-workshop-result") {
      setAiWorkshopState({
        ...store.getState().ui?.aiWorkshop,
        busy: false,
        resultTitle: "KI bereit",
        resultText: "Wähle eine Funktion für die KI-Startstufe.",
        resultTone: "neutral",
        resultData: null
      });
      updateStorageStatus("KI-Vorschlag verworfen.");
      return;
    }

    if (actionTarget?.dataset.action === "print-management-report") {
      printManagementReport(store.getState());
      return;
    }

    if (actionTarget?.dataset.action === "set-report-mode") {
      const reportMode = actionTarget.dataset.reportMode === "management" ? "management" : "risk";
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          reportMode,
          reportExportName: state.ui?.reportExportName
            ? state.ui.reportExportName
            : getDefaultReportExportName({ ui: { reportMode } })
        }
      }));
      updateStorageStatus(reportMode === "management" ? "Managementbericht gewählt." : "Risikobericht gewählt.");
      return;
    }

    if (actionTarget?.dataset.action === "jump-to-monte-p80") {
      jumpToMonteCarloP80();
      updateStorageStatus("Monte-Carlo-P80 geöffnet.");
      return;
    }

    if (actionTarget?.dataset.action === "jump-to-report-target") {
      jumpToReportTarget(actionTarget.dataset.targetModule, actionTarget.dataset.targetId);
      updateStorageStatus("Berichtsverweis geöffnet.");
      return;
    }

    if (actionTarget?.dataset.action === "add-pert-item") {
      const newItem = createPertItem();
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          items: [...state.pert.items, newItem]
        }
      }));
      requestAnimationFrame(() => {
        const section = document.getElementById("pertCostItemsSection");
        section?.scrollIntoView({ behavior: "smooth", block: "start" });
        requestAnimationFrame(() => {
          const input = document.getElementById(`pert_name_${store.getState().pert.items.length - 1}`);
          input?.focus();
          input?.select?.();
        });
      });
      return;
    }

    if (actionTarget?.dataset.action === "remove-pert-item") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          items: state.pert.items.filter((_, itemIndex) => itemIndex !== index)
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "save-pert-snapshot") {
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          snapshots: [createPertSnapshot(state), ...(state.pert.snapshots || [])].slice(0, 12)
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "restore-pert-snapshot") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => {
        const snapshot = state.pert.snapshots?.[index];
        if (!snapshot) return state;
        return {
          ...state,
          project: {
            ...state.project,
            phase: snapshot.phase || state.project.phase
          },
          pert: {
            ...state.pert,
            securityLevel: snapshot.securityLevel,
            items: JSON.parse(JSON.stringify(snapshot.items || [])),
            lphComments: JSON.parse(JSON.stringify(snapshot.lphComments || {}))
          }
        };
      });
      return;
    }

    if (actionTarget?.dataset.action === "delete-pert-snapshot") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          snapshots: state.pert.snapshots.filter((_, snapshotIndex) => snapshotIndex !== index)
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "transfer-pert-to-mc") {
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          workPackages: state.pert.items.map((item) => createMonteCarloWorkPackageFromPertItem(item))
        },
        ui: {
          ...state.ui,
          activeModule: "monteCarlo"
        }
      }));
      updateStorageStatus("PERT-Positionen wurden in Monte Carlo als Arbeitspakete übernommen.");
      return;
    }

    if (actionTarget?.dataset.action === "transfer-selected-pert-to-mc") {
      store.setState((state) => {
        const selectedIds = state.ui.transferSelections?.pertItemIds || [];
        const selectedItems = state.pert.items.filter((item) => selectedIds.includes(item.id));
        if (!selectedItems.length) return state;
        return {
          ...state,
          monteCarlo: {
            ...state.monteCarlo,
            workPackages: [...state.monteCarlo.workPackages, ...selectedItems.map((item) => createMonteCarloWorkPackageFromPertItem(item))]
          },
          ui: {
            ...state.ui,
            activeModule: "monteCarlo",
            transferSelections: {
              ...state.ui.transferSelections,
              pertItemIds: []
            }
          }
        };
      });
      updateStorageStatus("Markierte PERT-Positionen wurden an Monte Carlo angehängt.");
      return;
    }

    if (actionTarget?.dataset.action === "add-eva-activity") {
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          activities: [...state.earnedValue.activities, createEvaActivity()]
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "remove-eva-activity") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          activities: state.earnedValue.activities.filter((_, itemIndex) => itemIndex !== index)
        },
        ui: {
          ...state.ui,
          evaActivityOpenIds: normalizeEvaActivityOpenIds(
            (state.ui?.evaActivityOpenIds || []).filter((id) => {
              const activity = state.earnedValue.activities.find((item) => String(item.id) === String(id));
              return Boolean(activity);
            }),
            state.earnedValue.activities.filter((_, itemIndex) => itemIndex !== index)
          )
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "save-eva-snapshot") {
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          snapshots: [createEvaSnapshot(state), ...(state.earnedValue.snapshots || [])]
        }
      }));
      updateStorageStatus("EVA-Stichtag wurde gespeichert.");
      return;
    }

    if (actionTarget?.dataset.action === "delete-eva-snapshot") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          snapshots: (state.earnedValue.snapshots || []).filter((_, itemIndex) => itemIndex !== index)
        }
      }));
      updateStorageStatus("EVA-Stichtag wurde entfernt.");
      return;
    }

    if (actionTarget?.dataset.action === "clear-eva-snapshots") {
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          snapshots: []
        }
      }));
      updateStorageStatus("EVA-Verlauf wurde gelöscht.");
      return;
    }

    if (actionTarget?.dataset.action === "add-mc-work") {
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          workPackages: [...state.monteCarlo.workPackages, createMonteCarloWorkPackage()]
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "remove-mc-work") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          workPackages: state.monteCarlo.workPackages.filter((_, itemIndex) => itemIndex !== index)
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "add-mc-risk") {
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          riskEvents: [...state.monteCarlo.riskEvents, createMonteCarloRisk()]
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "save-mc-baseline") {
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          baselineScenario: createMonteCarloScenario(state, "baseline")
        }
      }));
      updateStorageStatus("Monte-Carlo-Baseline wurde gespeichert.");
      return;
    }

    if (actionTarget?.dataset.action === "save-mc-scenario") {
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          storedScenarios: [createMonteCarloScenario(state, "scenario"), ...(state.monteCarlo.storedScenarios || [])].slice(0, 8)
        }
      }));
      updateStorageStatus("Monte-Carlo-Szenario wurde gespeichert.");
      return;
    }

    if (actionTarget?.dataset.action === "restore-mc-baseline") {
      store.setState((state) => {
        const baseline = state.monteCarlo.baselineScenario;
        if (!baseline) return state;
        return {
          ...state,
          project: {
            ...state.project,
            phase: baseline.projectPhase || state.project.phase
          },
          monteCarlo: {
            ...state.monteCarlo,
            targetBudget: baseline.targetBudget,
            targetDays: baseline.targetDays,
            iterations: baseline.iterations,
            seed: baseline.seed,
            measureSettings: JSON.parse(JSON.stringify(baseline.measureSettings || {})),
            workPackages: JSON.parse(JSON.stringify(baseline.workPackages || [])),
            riskEvents: JSON.parse(JSON.stringify(baseline.riskEvents || []))
          }
        };
      });
      updateStorageStatus("Monte-Carlo-Baseline wurde wiederhergestellt.");
      return;
    }

    if (actionTarget?.dataset.action === "restore-mc-scenario") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => {
        const scenario = state.monteCarlo.storedScenarios?.[index];
        if (!scenario) return state;
        return {
          ...state,
          project: {
            ...state.project,
            phase: scenario.projectPhase || state.project.phase
          },
          monteCarlo: {
            ...state.monteCarlo,
            targetBudget: scenario.targetBudget,
            targetDays: scenario.targetDays,
            iterations: scenario.iterations,
            seed: scenario.seed,
            measureSettings: JSON.parse(JSON.stringify(scenario.measureSettings || {})),
            workPackages: JSON.parse(JSON.stringify(scenario.workPackages || [])),
            riskEvents: JSON.parse(JSON.stringify(scenario.riskEvents || []))
          }
        };
      });
      updateStorageStatus("Monte-Carlo-Szenario wurde wiederhergestellt.");
      return;
    }

    if (actionTarget?.dataset.action === "delete-mc-scenario") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          storedScenarios: state.monteCarlo.storedScenarios.filter((_, itemIndex) => itemIndex !== index)
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "remove-mc-risk") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          riskEvents: state.monteCarlo.riskEvents.filter((_, itemIndex) => itemIndex !== index)
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "transfer-all-risks-to-mc") {
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          riskEvents: state.riskRegister.risks.map((risk) => createMonteCarloRiskFromRegister(risk))
        },
        ui: {
          ...state.ui,
          activeModule: "monteCarlo"
        }
      }));
      updateStorageStatus("Registerrisiken wurden in die Monte-Carlo-Simulation übernommen.");
      return;
    }

    if (actionTarget?.dataset.action === "transfer-selected-risks-to-mc") {
      store.setState((state) => {
        const selectedIds = state.ui.transferSelections?.riskIds || [];
        const selectedRisks = state.riskRegister.risks.filter((risk) => selectedIds.includes(risk.id));
        if (!selectedRisks.length) return state;
        return {
          ...state,
          monteCarlo: {
            ...state.monteCarlo,
            riskEvents: [...state.monteCarlo.riskEvents, ...selectedRisks.map((risk) => createMonteCarloRiskFromRegister(risk))]
          },
          ui: {
            ...state.ui,
            activeModule: "monteCarlo",
            transferSelections: {
              ...state.ui.transferSelections,
              riskIds: []
            }
          }
        };
      });
      updateStorageStatus("Markierte Registerrisiken wurden in die Monte-Carlo-Simulation übernommen.");
      return;
    }

    if (actionTarget?.dataset.action === "transfer-risk-to-mc") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => {
        const risk = state.riskRegister.risks[index];
        if (!risk) return state;
        return {
          ...state,
          monteCarlo: {
            ...state.monteCarlo,
            riskEvents: [...state.monteCarlo.riskEvents, createMonteCarloRiskFromRegister(risk)]
          },
          ui: {
            ...state.ui,
            activeModule: "monteCarlo"
          }
        };
      });
      updateStorageStatus("Risiko aus dem Register wurde in Monte Carlo übernommen.");
      return;
    }

    if (actionTarget?.dataset.action === "toggle-risk-register-column") {
      const column = String(actionTarget.dataset.column || "");
      if (!column) return;
      store.setState((state) => {
        const current = state.ui?.riskRegisterView?.visibleColumns || [];
        const next = current.includes(column)
          ? current.filter((item) => item !== column)
          : [...current, column];
        return {
          ...state,
          ui: {
            ...state.ui,
            riskRegisterView: {
              ...state.ui.riskRegisterView,
              visibleColumns: next
            }
          }
        };
      });
      window.requestAnimationFrame(() => {
        const tableWrap = document.querySelector(".risk-register-table-wrap");
        if (tableWrap) {
          tableWrap.scrollLeft = 0;
        }
      });
      return;
    }

    if (actionTarget?.dataset.action === "show-all-risk-register-columns") {
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            visibleColumns: ["select", "priority", "status", "value", "category", "phase", "impact", "owner", "dueDate", "measures"]
          }
        }
      }));
      window.requestAnimationFrame(() => {
        const tableWrap = document.querySelector(".risk-register-table-wrap");
        if (tableWrap) {
          tableWrap.scrollLeft = 0;
        }
      });
      return;
    }

    if (actionTarget?.dataset.action === "set-risk-register-sort") {
      const sortBy = ["priority", "value", "id", "category", "dueDate"].includes(actionTarget.dataset.sortBy)
        ? actionTarget.dataset.sortBy
        : "priority";
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            sortBy
          }
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "set-risk-register-edit-sort") {
      const editSortBy = ["newest", "id"].includes(actionTarget.dataset.sortBy)
        ? actionTarget.dataset.sortBy
        : "newest";
      store.setState((state) => ({
        ...state,
        riskRegister: {
          ...state.riskRegister,
          risks: sortRiskRegisterEntriesForEdit(state.riskRegister.risks, editSortBy)
        },
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            editSortBy
          }
        }
      }));
      requestAnimationFrame(() => {
        const editList = document.querySelector(".risk-edit-list");
        if (editList) editList.scrollIntoView({ block: "start", behavior: "auto" });
      });
      return;
    }

    if (actionTarget?.dataset.action === "reset-risk-register-filters") {
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            search: "",
            status: "alle",
            owner: "alle",
            category: "alle",
            criticalOnly: false,
            dueFrom: "",
            dueTo: "",
            sortBy: "id",
            matrixSelection: null
          },
          transferSelections: {
            ...state.ui.transferSelections,
            riskIds: []
          }
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "set-risk-register-limit") {
      const topLimit = Number(actionTarget.dataset.limit || 5);
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            topLimit: topLimit === 0 || topLimit === 5 || topLimit === 10 || topLimit === 20 ? topLimit : 5
          }
        }
      }));
      return;
    }

    if (actionTarget?.dataset.action === "risk-step") {
      const index = Number(actionTarget.dataset.index);
      const field = actionTarget.dataset.field;
      const step = Number(actionTarget.dataset.step || 1);
      const direction = actionTarget.dataset.direction === "down" ? -1 : 1;
      if (!field || !Number.isFinite(index)) return;
      store.setState((state) => {
        pushRiskRegisterUndoSnapshot(state);
        const risk = state.riskRegister.risks[index];
        if (!risk) return state;
        const currentValue = Number(risk[field]) || 0;
        let nextValue = currentValue + (step * direction);
        if (field === "probabilityPercent") {
          nextValue = Math.max(0, Math.min(100, nextValue));
        } else if (field === "likelihood" || field === "impact") {
          nextValue = Math.max(1, Math.min(5, nextValue));
        } else {
          nextValue = Math.max(0, nextValue);
        }
        return {
          ...state,
          riskRegister: {
            ...state.riskRegister,
            risks: state.riskRegister.risks.map((item, itemIndex) => {
              if (itemIndex !== index) return item;
              const nextProbabilityPercent = field === "probabilityPercent"
                ? Math.round(nextValue)
                : Math.max(0, Math.min(100, Number(item.probabilityPercent) || 0));
              const nextLikelihood = deriveRiskLikelihoodFromPercent(nextProbabilityPercent, Number(item.likelihood) || 1);
              const nextImpact = field === "impact"
                ? Math.max(1, Math.min(5, nextValue))
                : Math.max(1, Math.min(5, Number(item.impact) || 1));
              const nextFinancialImpact = field === "financialImpact"
                ? Math.round(nextValue)
                : Math.max(0, Number(item.financialImpact) || 0);
              return {
                ...item,
                financialImpact: nextFinancialImpact,
                probabilityPercent: nextProbabilityPercent,
                likelihood: nextLikelihood,
                expectedDamage: Math.max(0, nextFinancialImpact * (nextProbabilityPercent / 100)),
                impact: nextImpact,
                qualitativeRiskValue: Math.max(1, Math.min(25, nextLikelihood * nextImpact)),
                [field]: field === "financialImpact" || field === "probabilityPercent" ? Math.round(nextValue) : nextValue
              };
            })
          }
        };
      });
      return;
    }

    if (actionTarget?.dataset.action === "select-risk-register-matrix-cell") {
      const likelihood = Number(actionTarget.dataset.likelihood || 0);
      const impact = Number(actionTarget.dataset.impact || 0);
      if (!Number.isInteger(likelihood) || !Number.isInteger(impact)) return;
      store.setState((state) => {
        const current = state.ui?.riskRegisterView?.matrixSelection || null;
        const sameCell = current && current.likelihood === likelihood && current.impact === impact;
        return {
          ...state,
          ui: {
            ...state.ui,
            riskRegisterView: {
              ...state.ui.riskRegisterView,
              matrixSelection: sameCell ? null : { likelihood, impact }
            }
          }
        };
      });
      return;
    }

    if (actionTarget?.dataset.action === "add-risk-register-item") {
      store.setState((state) => {
        pushRiskRegisterUndoSnapshot(state);
        const editSortBy = ["newest", "id"].includes(state.ui?.riskRegisterView?.editSortBy)
          ? state.ui.riskRegisterView.editSortBy
          : "newest";
        const nextItem = { ...createRiskRegisterItem(), id: nextRiskRegisterId(state.riskRegister.risks) };
        const nextRisks = editSortBy === "id"
          ? sortRiskRegisterEntriesForEdit([nextItem, ...state.riskRegister.risks], "id")
          : [nextItem, ...state.riskRegister.risks];
        return {
          ...state,
          riskRegister: {
            ...state.riskRegister,
            risks: nextRisks
          }
        };
      });
      return;
    }

    if (actionTarget?.dataset.action === "remove-risk-register-item") {
      const index = Number(actionTarget.dataset.index);
      store.setState((state) => {
        pushRiskRegisterUndoSnapshot(state);
        return {
          ...state,
          riskRegister: {
            ...state.riskRegister,
            risks: state.riskRegister.risks.filter((_, itemIndex) => itemIndex !== index)
          }
        };
      });
    }

    if (actionTarget?.dataset.action === "undo-risk-register-change") {
      if (applyRiskRegisterUndo()) {
        renderApp();
      }
      return;
    }

    if (actionTarget?.dataset.action === "jump-eva-activities") {
      event.preventDefault();
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          evaActivitiesPanelOpen: true
        }
      }));
      requestAnimationFrame(() => {
        const activitiesPanel = document.getElementById("eva-activities-panel");
        if (activitiesPanel) {
          activitiesPanel.open = true;
          activitiesPanel.scrollIntoView({ behavior: "smooth", block: "start" });
        }
      });
      return;
    }

    if (actionTarget?.dataset.action === "set-eva-detail-view") {
      event.preventDefault();
      event.stopPropagation();
      const view = actionTarget.dataset.view === "full" ? "full" : "compact";
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          evaDetailView: view
        }
      }));
      requestAnimationFrame(() => {
        const panel = document.getElementById("eva-detail-panel");
        if (panel) {
          panel.open = true;
          panel.scrollIntoView({ behavior: "auto", block: "start" });
        }
      });
      return;
    }

    if (actionTarget?.dataset.action === "toggle-eva-detail-column") {
      event.preventDefault();
      event.stopPropagation();
      const field = actionTarget.dataset.field;
      if (!field) return;
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          evaDetailColumns: {
            ...state.ui.evaDetailColumns,
            [field]: !(state.ui.evaDetailColumns?.[field] !== false)
          }
        }
      }));
      requestAnimationFrame(() => {
        const panel = document.getElementById("eva-detail-panel");
        if (panel) {
          panel.open = true;
          panel.scrollIntoView({ behavior: "auto", block: "start" });
        }
      });
      return;
    }

    if (actionTarget?.dataset.action === "eva-threshold-step") {
      const field = actionTarget.dataset.field;
      const direction = actionTarget.dataset.direction === "down" ? -1 : 1;
      const config = getEvaThresholdConfig(field);
      store.setState((state) => {
        const current = Number(state.earnedValue.thresholds?.[field]) || 0;
        const stepped = current + direction * config.step;
        const normalized = normalizeEvaThresholdValue(field, stepped);
        return {
          ...state,
          earnedValue: {
            ...state.earnedValue,
            thresholds: {
              ...state.earnedValue.thresholds,
              [field]: normalized
            }
          }
        };
      });
      return;
    }

    if (actionTarget?.dataset.action === "eva-activity-step") {
      event.preventDefault();
      const index = Number(actionTarget.dataset.index);
      const field = actionTarget.dataset.field;
      const direction = actionTarget.dataset.direction === "down" ? -1 : 1;
      const step = Number(actionTarget.dataset.step) || 1;
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          activities: state.earnedValue.activities.map((activity, activityIndex) => {
            if (activityIndex !== index) return activity;
            const current = Number(activity[field]) || 0;
            const next = Math.max(0, Math.min(100, current + direction * step));
            return {
              ...activity,
              [field]: next
            };
          })
        }
      }));
      return;
    }

  });

  document.addEventListener("keydown", (event) => {
    if (!["Enter", " "].includes(event.key)) return;
    const actionTarget = event.target.closest?.("[data-action]");
    if (!actionTarget) return;
    const action = actionTarget.dataset.action;
    if (!action) return;
    event.preventDefault();
    if (action === "jump-to-report-target") {
      jumpToReportTarget(actionTarget.dataset.targetModule, actionTarget.dataset.targetId);
      updateStorageStatus("Berichtsverweis geöffnet.");
      return;
    }
    if (action === "jump-to-monte-p80") {
      jumpToMonteCarloP80();
      updateStorageStatus("Monte-Carlo-P80 geöffnet.");
      return;
    }
    actionTarget.click();
  });

  document.addEventListener("change", (event) => {
    if (event.target.id === "eva_cv_tolerance_percent") {
      return;
    }

    if (event.target.id === "projectFileInput") {
      importProjectFile(event.target.files?.[0]);
      event.target.value = "";
      return;
    }

    const projectField = event.target.closest("[data-project-field]");
    if (projectField) {
      const field = projectField.getAttribute("data-project-field");
      const value = parseValue(store.getState().project[field], projectField.value);
      store.setState((state) => ({
        ...state,
        project: {
          ...state.project,
          [field]: value
        }
      }));
      return;
    }

    const locationField = event.target.closest("[data-project-location-field]");
    if (locationField) {
      const field = locationField.getAttribute("data-project-location-field");
      store.setState((state) => ({
        ...state,
        project: {
          ...state.project,
          location: {
            ...state.project.location,
            [field]: locationField.value
          }
        }
      }));
      return;
    }

    const addressField = event.target.closest("[data-project-address-field]");
    if (addressField) {
      const nextAddress = splitStreetAddress(addressField.value);
      store.setState((state) => ({
        ...state,
        project: {
          ...state.project,
          location: {
            ...state.project.location,
            street: nextAddress.street,
            houseNumber: nextAddress.houseNumber
          }
        }
      }));
      return;
    }

    const reportField = event.target.closest("[data-report-field]");
    if (reportField) {
      const field = reportField.getAttribute("data-report-field");
      store.setState((state) => ({
        ...state,
        reportProfile: {
          ...state.reportProfile,
          [field]: reportField.value
        }
      }));
      return;
    }

    const reportOptionField = event.target.closest("[data-report-option]");
    if (reportOptionField) {
      const field = reportOptionField.getAttribute("data-report-option");
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          reportOptions: {
            ...state.ui.reportOptions,
            [field]: reportOptionField.checked
          }
        }
      }));
      return;
    }

    const reportExportField = event.target.closest("[data-report-export-field]");
    if (reportExportField) {
      const field = reportExportField.getAttribute("data-report-export-field");
      if (!field) return;
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          reportExportName: field === "fileName" ? reportExportField.value : state.ui.reportExportName,
          reportExportFormat: field === "format" ? reportExportField.value : state.ui.reportExportFormat
        }
      }));
      return;
    }

    const projectExportField = event.target.closest("[data-project-export-field]");
    if (projectExportField) {
      const field = projectExportField.getAttribute("data-project-export-field");
      if (!field) return;
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          projectExportName: field === "fileName" ? projectExportField.value : state.ui.projectExportName
        }
      }));
      return;
    }

    const aiSettingField = event.target.closest("[data-ai-setting-field]");
    if (aiSettingField) {
      aiSettings = persistAiSettings({
        ...readAiSettingsFromPanel(),
        [aiSettingField.getAttribute("data-ai-setting-field")]: aiSettingField.value,
        connected: false,
        testing: false,
        lastStatus: "KI-Einstellungen geändert. Bitte Verbindung testen."
      }) ? loadAiSettings() : normalizeAiSettings({
        ...readAiSettingsFromPanel(),
        [aiSettingField.getAttribute("data-ai-setting-field")]: aiSettingField.value,
        connected: false,
        testing: false,
        lastStatus: "KI-Einstellungen geändert. Bitte Verbindung testen."
      });
      renderAiSettingsPanel();
      return;
    }

    const riskUiField = event.target.closest("[data-risk-ui-field]");
    if (riskUiField) {
      const field = riskUiField.getAttribute("data-risk-ui-field");
      const nextValue = riskUiField.type === "checkbox" ? riskUiField.checked : riskUiField.value;
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            ...(field === "dueFrom" || field === "dueTo"
              ? { sortBy: nextValue ? "dueDate" : state.ui.riskRegisterView.sortBy }
              : {}),
            [field]: nextValue
          }
        }
      }));
      return;
    }

    const riskEditSortField = event.target.closest("[data-risk-edit-sort-field]");
    if (riskEditSortField) {
      const sortBy = riskEditSortField.value;
      store.setState((state) => ({
        ...state,
        riskRegister: {
          ...state.riskRegister,
          risks: sortRiskRegisterEntriesForEdit(state.riskRegister.risks, sortBy)
        },
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            editSortBy: ["newest", "id"].includes(sortBy) ? sortBy : "newest"
          }
        }
      }));
      return;
    }

    const transferSelect = event.target.closest("[data-transfer-select]");
    if (transferSelect) {
      const group = transferSelect.getAttribute("data-transfer-select");
      const value = transferSelect.getAttribute("data-transfer-value");
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          transferSelections: {
            ...state.ui.transferSelections,
            [group]: toggleSelection(state.ui.transferSelections?.[group], value, transferSelect.checked)
          }
        }
      }));
      return;
    }

    const pertField = event.target.closest("[data-pert-field]");
    if (pertField) {
      const field = pertField.getAttribute("data-pert-field");
      const currentValue = store.getState().pert[field];
      const nextValue = parseValue(currentValue, pertField.value);
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          [field]: nextValue
        }
      }));
      return;
    }

    const pertItemField = event.target.closest("[data-pert-item-field]");
    if (pertItemField) {
      const index = Number(pertItemField.getAttribute("data-pert-item-index"));
      const field = pertItemField.getAttribute("data-pert-item-field");
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          items: state.pert.items.map((item, itemIndex) => {
            if (itemIndex !== index) return item;
            const nextValue = parseValue(item[field], pertItemField.value);
            return {
              ...item,
              [field]: nextValue
            };
          })
        }
      }));
      return;
    }

    const pertLphField = event.target.closest("[data-pert-lph-field]");
    if (pertLphField) {
      const field = pertLphField.getAttribute("data-pert-lph-field");
      const label = pertLphField.getAttribute("data-pert-lph-label");
      store.setState((state) => ({
        ...state,
        pert: {
          ...state.pert,
          lphComments: field === "comment"
            ? {
                ...state.pert.lphComments,
                [label]: pertLphField.value
              }
            : state.pert.lphComments
        }
      }));
      return;
    }

    const evaField = event.target.closest("[data-eva-field]");
    if (evaField) {
      const field = evaField.getAttribute("data-eva-field");
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          [field]: evaField.value
        }
      }));
      return;
    }

    const evaThresholdField = event.target.closest("[data-eva-threshold-field]");
    if (evaThresholdField) {
      const field = evaThresholdField.getAttribute("data-eva-threshold-field");
      const rawValue = evaThresholdField.value;
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          thresholds: {
            ...state.earnedValue.thresholds,
            [field]: (() => {
              const decimalFields = new Set(["spiYellow", "cpiYellow", "tcpiActionLevel"]);
              const parsed = field === "cvTolerancePercent" || evaThresholdField.type === "range"
                ? Number(rawValue)
                : decimalFields.has(field)
                  ? normalizeEvaThresholdValue(field, rawValue)
                  : field === "eacTolerance"
                    ? normalizeEvaThresholdValue(field, rawValue)
                    : parseValue(state.earnedValue.thresholds[field], rawValue);
              return Number.isFinite(parsed) ? parsed : state.earnedValue.thresholds[field];
            })()
          }
        }
      }));
      return;
    }

    const evaActivityField = event.target.closest("[data-eva-activity-field]");
    if (evaActivityField) {
      const index = Number(evaActivityField.getAttribute("data-eva-activity-index"));
      const field = evaActivityField.getAttribute("data-eva-activity-field");
      store.setState((state) => ({
        ...state,
        earnedValue: {
          ...state.earnedValue,
          activities: state.earnedValue.activities.map((activity, activityIndex) => {
            if (activityIndex !== index) return activity;
            const nextValue = parseValue(activity[field], evaActivityField.value);
            return {
              ...activity,
              [field]: nextValue
            };
          })
        }
      }));
      return;
    }

    const mcField = event.target.closest("[data-mc-field]");
    if (mcField) {
      const field = mcField.getAttribute("data-mc-field");
      const currentValue = store.getState().monteCarlo[field];
      const nextValue = parseValue(currentValue, mcField.value);
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          [field]: nextValue
        }
      }));
      return;
    }

    const mcMeasureField = event.target.closest("[data-mc-measure-field]");
    if (mcMeasureField) {
      const field = mcMeasureField.getAttribute("data-mc-measure-field");
      const currentValue = store.getState().monteCarlo.measureSettings?.[field];
      const value = parseValue(currentValue, mcMeasureField.value);
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          measureSettings: {
            ...state.monteCarlo.measureSettings,
            [field]: value
          }
        }
      }));
      return;
    }

    const mcWorkField = event.target.closest("[data-mc-work-field]");
    if (mcWorkField) {
      const index = Number(mcWorkField.getAttribute("data-mc-work-index"));
      const field = mcWorkField.getAttribute("data-mc-work-field");
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          workPackages: state.monteCarlo.workPackages.map((item, itemIndex) => {
            if (itemIndex !== index) return item;
            const nextValue = parseValue(item[field], mcWorkField.value);
            return {
              ...item,
              [field]: nextValue
            };
          })
        }
      }));
      return;
    }

    const mcRiskField = event.target.closest("[data-mc-risk-field]");
    if (mcRiskField) {
      const index = Number(mcRiskField.getAttribute("data-mc-risk-index"));
      const field = mcRiskField.getAttribute("data-mc-risk-field");
      store.setState((state) => ({
        ...state,
        monteCarlo: {
          ...state.monteCarlo,
          riskEvents: state.monteCarlo.riskEvents.map((item, itemIndex) => {
            if (itemIndex !== index) return item;
            const nextValue = field === "probability"
              ? Math.max(0, Math.min(1, parseFlexibleNumber(mcRiskField.value) / 100))
              : parseValue(item[field], mcRiskField.value);
            return {
              ...item,
              [field]: nextValue
            };
          })
        }
      }));
      return;
    }

    const riskField = event.target.closest("[data-risk-field]");
    if (riskField) {
      const index = Number(riskField.getAttribute("data-risk-index"));
      const field = riskField.getAttribute("data-risk-field");
      store.setState((state) => {
        pushRiskRegisterUndoSnapshot(state);
        return {
          ...state,
          riskRegister: {
            ...state.riskRegister,
            risks: state.riskRegister.risks.map((risk, riskIndex) => {
              if (riskIndex !== index) return risk;
              const nextValue = parseValue(risk[field], riskField.value);
              return {
                ...risk,
                [field]: nextValue
              };
            })
          }
        };
      });
    }
  });

  function handleAiWorkshopFieldEvent(event) {
    const aiWorkshopField = event.target.closest("[data-ai-workshop-field]");
    if (!aiWorkshopField) return false;
    if (aiWorkshopField.getAttribute("data-ai-workshop-field") === "freeText") {
      const nextValue = aiWorkshopField.value;
      const state = store.getState();
      state.ui = {
        ...state.ui,
        aiWorkshop: normalizeAiWorkshopState({
          ...state.ui?.aiWorkshop,
          freeText: nextValue
        })
      };
      persistAutosave(state);
      return true;
    }
    if (aiWorkshopField.getAttribute("data-ai-workshop-field") === "selectedRiskId") {
      const nextValue = aiWorkshopField.value;
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          aiWorkshop: normalizeAiWorkshopState({
            ...state.ui?.aiWorkshop,
            selectedRiskId: nextValue
          })
        }
      }));
      return true;
    }
    return false;
  }

  document.addEventListener("input", (event) => {
    if (handleAiWorkshopFieldEvent(event)) return;

    const aiSettingField = event.target.closest("[data-ai-setting-field]");
    if (aiSettingField && aiSettingField.getAttribute("data-ai-setting-field") === "apiKey") {
      const saved = persistAiSettings({
        ...readAiSettingsFromPanel(),
        apiKey: aiSettingField.value,
        connected: false,
        testing: false,
        lastStatus: "Einstellungen speichern"
      });
      aiSettings = saved ? loadAiSettings() : normalizeAiSettings({
        ...readAiSettingsFromPanel(),
        apiKey: aiSettingField.value,
        connected: false,
        testing: false,
        lastStatus: "Einstellungen speichern"
      });
      if (!saved) {
        updateAiStatus("KI-Einstellungen konnten im Browser nicht gespeichert werden.");
      }
      renderAiSettingsPanel();
      return;
    }

    const riskEditSortField = event.target.closest("[data-risk-edit-sort-field]");
    if (riskEditSortField) {
      const sortBy = ["newest", "id"].includes(riskEditSortField.value)
        ? riskEditSortField.value
        : "newest";
      store.setState((state) => ({
        ...state,
        riskRegister: {
          ...state.riskRegister,
          risks: sortRiskRegisterEntriesForEdit(state.riskRegister.risks, sortBy)
        },
        ui: {
          ...state.ui,
          riskRegisterView: {
            ...state.ui.riskRegisterView,
            editSortBy: sortBy
          }
        }
      }));
      return;
    }

    const riskUiField = event.target.closest("[data-risk-ui-field]");
    if (!riskUiField) return;
    if (riskUiField.type === "checkbox" || riskUiField.tagName === "SELECT") return;
    const field = riskUiField.getAttribute("data-risk-ui-field");
    if (!field) return;
    const nextValue = riskUiField.value;
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        riskRegisterView: {
          ...state.ui.riskRegisterView,
          [field]: nextValue
        }
        }
      }));
    });

  document.addEventListener("change", (event) => {
    handleAiWorkshopFieldEvent(event);
  });

}

store.subscribe(renderApp);
store.subscribe(persistAutosave);
bindEvents();
renderApp(store.getState());
restoreAutosave();
window.addEventListener("resize", layoutEvaToleranceAxis);
