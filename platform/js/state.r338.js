export const initialState = {
  meta: {
    app: "Project Controls Hub",
    version: "1.0.0",
    createdAt: new Date().toISOString(),
    savedAt: new Date().toISOString(),
  },
  project: {
    name: "Neubau Verwaltungsgebäude Nordpark",
    type: "Büro- und Verwaltungsbau",
    bauart: "Neubau",
    phase: "LPH 2",
    bgf: 3000,
    budget: 18500000,
    costBasis: "netto",
    currency: "EUR",
    location: {
      street: "Bauhofstraße",
      houseNumber: "18",
      postalCode: "40221",
      city: "Düsseldorf",
      country: "DE"
    },
    client: "Nordpark Immobilien GmbH",
    projectLead: "Dipl.-Ing. Jana Richter",
    startDate: "2026-04-01",
    endDate: "2027-08-31",
    analysisDate: "2026-04-04",
    description: "Modulare Project-Controls-Plattform für Schätzung, Simulation, Controlling und Risiken."
  },
  reportProfile: {
    company: "BuiltSmart Hub",
    companyAddress: "",
    author: "Bernhard Metzger",
    clientName: "Nordpark Immobilien GmbH",
    clientAddress: "",
    projectAddress: "Bauhofstraße 18, 40221 Düsseldorf",
    confidentiality: "Vertraulich",
    notes: "",
    logoDataUrl: null
  },
  pert: {
    securityLevel: 1,
    items: [
      { id: "pert-1", name: "Rohbau", optimistic: 3600000, mostLikely: 4000000, pessimistic: 4800000, comment: "Vorplanung" },
      { id: "pert-2", name: "TGA", optimistic: 950000, mostLikely: 1100000, pessimistic: 1450000, comment: "Erste Grobschätzung" },
      { id: "pert-3", name: "Ausbau & Fassade", optimistic: 800000, mostLikely: 950000, pessimistic: 1200000, comment: "" }
    ],
    lphComments: {
      "LPH 2": "Aktueller Schätzstand auf Basis der Vorplanung.",
      "LPH 3": "Fortschreibung nach vertiefter Planung vorgesehen."
    },
    snapshots: [],
    lastResult: {
      expectedValue: 6050000,
      standardDeviation: 216000,
      budget84: 6266000,
      budget975: 6482000
    }
  },
  monteCarlo: {
    iterations: 10000,
    seed: 42,
    targetBudget: 1400000,
    targetDays: 280,
    workPackages: [
      { id: "mc-1", name: "Rohbau", optimisticCost: 420000, mostLikelyCost: 500000, pessimisticCost: 650000, optimisticDays: 70, mostLikelyDays: 90, pessimisticDays: 125 },
      { id: "mc-2", name: "TGA", optimisticCost: 180000, mostLikelyCost: 240000, pessimisticCost: 320000, optimisticDays: 35, mostLikelyDays: 50, pessimisticDays: 80 }
    ],
    riskEvents: [
      { id: "mc-r1", name: "Schlechtwetterphase", probability: 0.3, minCostImpact: 8000, maxCostImpact: 35000, minDelayDays: 5, maxDelayDays: 20 },
      { id: "mc-r2", name: "Materialpreissteigerung", probability: 0.25, minCostImpact: 20000, maxCostImpact: 90000, minDelayDays: 0, maxDelayDays: 7 }
    ],
    baselineScenario: null,
    storedScenarios: [],
    measureSettings: {
      budgetPercent: 10,
      daysPercent: 10,
      riskReductionPercent: 15,
      impactReductionPercent: 10
    },
    lastResult: {
      meanCost: 1512000,
      meanDays: 297,
      medianCost: 1489000,
      medianDays: 292,
      p10Cost: 1320000,
      p50Cost: 1489000,
      p80Cost: 1635000,
      p90Cost: 1718000,
      p10Days: 268,
      p50Days: 292,
      p80Days: 316,
      p90Days: 329,
      budgetSuccessRate: 0.34,
      scheduleSuccessRate: 0.39,
      combinedSuccessRate: 0.22
    }
  },
  earnedValue: {
    statusDate: "2026-04-04",
    activities: [
      { id: "eva-1", name: "Rohbau", plannedValue: 500000, earnedValue: 460000, actualCost: 520000, progressPercent: 92 },
      { id: "eva-2", name: "TGA", plannedValue: 220000, earnedValue: 180000, actualCost: 210000, progressPercent: 78 }
    ],
    snapshots: [],
    thresholds: {
      spiGreen: 1.0,
      spiYellow: 0.95,
      cpiGreen: 1.0,
      cpiYellow: 0.95,
      cvTolerancePercent: 3,
      cvTolerance: 180000,
      eacTolerance: 180000,
      tcpiActionLevel: 1.02
    },
    lastResult: {
      pv: 720000,
      ev: 640000,
      ac: 730000,
      cv: -90000,
      sv: -80000,
      cpi: 0.88,
      spi: 0.89,
      eac: 20500000,
      tcpi: 1.08
    }
  },
  riskRegister: {
    risks: [
      {
        id: "R-0005",
        createdAt: "2026-04-05T08:00:00.000Z",
        title: "Sicherheitsmangel bei Arbeiten am Baugerüst",
        description: "Unzureichend gesicherte Arbeitsbereiche können zu Unterbrechungen und Nachforderungen führen.",
        phase: "Compliance / Fassade / Arbeitssicherheit",
        category: "Compliance",
        area: "Arbeitssicherheit / Gerüst",
        financialImpact: 400000,
        probabilityPercent: 15,
        expectedDamage: 60000,
        likelihood: 3,
        impact: 3,
        qualitativeRiskValue: 9,
        owner: "HSE-Koordination",
        measures: "Tägliche Sicherheitsbegehungen und Nachschulungen der Gewerke durchführen",
        dueDate: "2026-04-05",
        status: "Maßnahme läuft",
        residualRisk: "Sicherheitsrisiko bleibt reduziert"
      },
      {
        id: "R-0004",
        createdAt: "2026-04-04T08:00:00.000Z",
        title: "Witterungsbedingte Unterbrechung der Dachabdichtung",
        description: "Anhaltender Regen oder Sturm können Außenarbeiten und Taktung verzögern.",
        phase: "Operativ / Ausbau / Dach",
        category: "Operativ",
        area: "Dach / Witterung",
        financialImpact: 90000,
        probabilityPercent: 45,
        expectedDamage: 40500,
        likelihood: 3,
        impact: 4,
        qualitativeRiskValue: 12,
        owner: "Bauleitung Ausbau",
        measures: "Wetterfenster täglich bewerten, Schutzabdeckungen bereithalten, Puffer einplanen",
        dueDate: "2026-04-15",
        status: "In Beobachtung",
        residualRisk: "Wetterrisiko bleibt mittel"
      },
      {
        id: "R-0003",
        createdAt: "2026-04-03T08:00:00.000Z",
        title: "Unerwartete Bodenverhältnisse im Baugrubenaushub",
        description: "Abweichende Bodenklassen oder Altlasten können Tiefbau und Gründung deutlich beeinflussen.",
        phase: "Technisch / Baugrube / Gründung",
        category: "Technisch",
        area: "Baugrube / Geotechnik",
        financialImpact: 320000,
        probabilityPercent: 25,
        expectedDamage: 80000,
        likelihood: 3,
        impact: 3,
        qualitativeRiskValue: 9,
        owner: "Projektleitung Tiefbau",
        measures: "Zusätzliche Sondierungen veranlassen, Bodengutachter einbinden, Reserven prüfen",
        dueDate: "2026-04-07",
        status: "Offen",
        residualRisk: "Nachgründung möglich"
      },
      {
        id: "R-0002",
        createdAt: "2026-04-02T08:00:00.000Z",
        title: "Kostensteigerung bei Bewehrungsstahl",
        description: "Steigende Stahlpreise können die freigegebenen Budgets in der Rohbauvergabe belasten.",
        phase: "Finanziell / Vergabe / Rohbau",
        category: "Finanziell",
        area: "Rohbau / Materialkosten",
        financialImpact: 250000,
        probabilityPercent: 30,
        expectedDamage: 75000,
        likelihood: 5,
        impact: 3,
        qualitativeRiskValue: 15,
        owner: "Einkauf / Projektkaufmann",
        measures: "Preisklauseln prüfen, Vergabepakete vorziehen, Marktanalyse aktualisieren",
        dueDate: "2026-04-12",
        status: "In Beobachtung",
        residualRisk: "Kostenrisiko bleibt erhöht"
      },
      {
        id: "R-0001",
        createdAt: "2026-04-01T08:00:00.000Z",
        title: "Lieferverzug bei Betonfertigteilen",
        description: "Die termingerechte Montage der Fertigteile ist gefährdet und kann Folgegewerke verzögern.",
        phase: "Operativ / Rohbau / Terminsteuerung",
        category: "Operativ",
        area: "Rohbau / Lieferkette",
        financialImpact: 180000,
        probabilityPercent: 35,
        expectedDamage: 63000,
        likelihood: 5,
        impact: 3,
        qualitativeRiskValue: 15,
        owner: "Projektleitung Rohbau",
        measures: "Alternativlieferanten anfragen, Montagefolgen takten, Lieferstatus täglich prüfen",
        dueDate: "2026-04-09",
        status: "Maßnahme läuft",
        residualRisk: "Terminrisiko bleibt mittel"
      }
    ],
    lastResult: {
      totalExpectedDamage: 318500,
      criticalCount: 2,
      activeCount: 5,
      closedCount: 0,
      overdueCount: 1
    }
  },
  ui: {
    activeModule: "project",
    dirty: false,
    lastSavedAt: null,
    evaDetailView: "compact",
    evaDetailColumns: {
      pv: true,
      ac: true,
      ev: true,
      cpi: true,
      spi: true,
      fgr: true,
      cv: true,
      sv: true,
      eac: true,
      status: true
    },
    evaActivitiesPanelOpen: false,
    evaActivityOpenIds: ["eva-1"],
    reportOptions: {
      includePert: true,
      includeMonteCarlo: true,
      includeEarnedValue: true,
      includeRiskRegister: true
    },
    aiWorkshop: {
      activeTask: "management-report",
      freeText: "",
      selectedRiskId: "",
      busy: false,
      resultTitle: "KI bereit",
      resultText: "Wähle eine Funktion für die KI-Startstufe.",
      resultTone: "neutral",
      resultData: null
    },
    projectExportName: "Neubau Verwaltungsgebäude Nordpark",
    reportExportName: "",
    reportExportFormat: "txt",
    reportMode: "risk",
    riskRegisterView: {
      search: "",
      status: "alle",
      owner: "alle",
      criticalOnly: false,
      visibleColumns: ["select", "category", "phase", "impact", "value", "priority", "status", "owner", "dueDate", "measures"],
      topLimit: 5,
      topSortBy: "priority",
      sortBy: "priority",
      editSortBy: "newest",
      matrixSelection: null
    },
    transferSelections: {
      pertItemIds: [],
      riskIds: []
    },
    riskRegisterRedoStack: []
  }
};

export function cloneState() {
  return JSON.parse(JSON.stringify(initialState));
}

export function createStore(initial) {
  let state = initial;
  const listeners = new Set();

  function getState() {
    return state;
  }

  function setState(updater) {
    state = typeof updater === "function" ? updater(state) : updater;
    state.ui.dirty = true;
    listeners.forEach((listener) => listener(state));
  }

  function subscribe(listener) {
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  function markSaved() {
    state = {
      ...state,
      meta: { ...state.meta, savedAt: new Date().toISOString() },
      ui: { ...state.ui, dirty: false, lastSavedAt: new Date().toISOString() }
    };
    listeners.forEach((listener) => listener(state));
  }

  return { getState, setState, subscribe, markSaved };
}
