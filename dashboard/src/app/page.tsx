'use client';

import { useCallback, useMemo, useRef, useState } from 'react';
import { FiDownload, FiUpload, FiTrash2, FiFileText, FiRefreshCcw } from 'react-icons/fi';

type FixedUnitRow = {
  id: string;
  unitId: string;
  bands: string;
};

type FloorRuleRow = {
  id: string;
  minFloor: string;
  maxFloor: string;
  minBand: string;
};

type UtilitiesState = {
  cooking: string;
  heat: string;
  hot_water: string;
};

type AnalysisResponse = {
  project_summary?: Record<string, any>;
  analysis_notes?: string[];
  scenarios?: Record<string, any>;
  download_link?: string;
};

const DEFAULT_UTILITIES: UtilitiesState = {
  cooking: 'na',
  heat: 'na',
  hot_water: 'na',
};

const STEP_CONFIG = [
  { title: 'Upload Units', subtitle: 'Provide the project rent roll (.xlsx, .xlsm, or .xlsb).' },
  { title: 'Rent Workbook', subtitle: 'Attach the annual rent calculator if you need updated allowances.' },
  { title: 'Utilities', subtitle: 'Tell us who pays for cooking, heat, and hot water.' },
  { title: 'Overrides & Settings', subtitle: 'Fine tune solver preferences, overrides, and export settings.' },
  { title: 'Review & Run', subtitle: 'Verify inputs, run the solver, and download reports.' },
];

const COOKING_OPTIONS = [
  { value: 'electric', label: 'Electric Stove' },
  { value: 'gas', label: 'Gas Stove' },
  { value: 'na', label: 'N/A or owner pays' },
];

const HEAT_OPTIONS = [
  { value: 'electric_ccashp', label: 'Electric Heat - ccASHP' },
  { value: 'electric_other', label: 'Electric Heat - Other' },
  { value: 'gas', label: 'Gas Heat' },
  { value: 'oil', label: 'Oil Heat' },
  { value: 'na', label: 'N/A or owner pays' },
];

const HOT_WATER_OPTIONS = [
  { value: 'electric_heat_pump', label: 'Electric Hot Water - Heat Pump' },
  { value: 'electric_other', label: 'Electric Hot Water - Other' },
  { value: 'gas', label: 'Gas Hot Water' },
  { value: 'oil', label: 'Oil Hot Water' },
  { value: 'na', label: 'N/A or owner pays' },
];

const SCENARIO_TITLES: Record<string, string> = {
  scenario_absolute_best: 'Absolute Best (S1)',
  scenario_client_oriented: 'Client Oriented (S2)',
  scenario_best_3_band: 'Best 3-Band (S3)',
  scenario_best_2_band: 'Best 2-Band (S4)',
  scenario_alternative: 'Alternative (S5)',
};

export default function DashboardPage() {
  const [currentStep, setCurrentStep] = useState(0);
  const [unitFile, setUnitFile] = useState<File | null>(null);
  const [rentFile, setRentFile] = useState<File | null>(null);
  const [utilities, setUtilities] = useState<UtilitiesState>(DEFAULT_UTILITIES);
  const [bandWhitelist, setBandWhitelist] = useState('');
  const [fixedUnits, setFixedUnits] = useState<FixedUnitRow[]>([]);
  const [floorRules, setFloorRules] = useState<FloorRuleRow[]>([]);
  const [premiumWeights, setPremiumWeights] = useState({
    floor: 0.45,
    net_sf: 0.3,
    bedrooms: 0.15,
    balcony: 0.1,
  });
  const [notes, setNotes] = useState('');
  const [preferXlsb, setPreferXlsb] = useState(true);
  const [analysis, setAnalysis] = useState<AnalysisResponse | null>(null);
  const [selectedScenarioKey, setSelectedScenarioKey] = useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const settingsInputRef = useRef<HTMLInputElement | null>(null);

  const stepDisabled = (index: number) => {
    if (index === 0) return false;
    if (!unitFile) return true;
    return false;
  };

  const canGoNext = currentStep < STEP_CONFIG.length - 1;
  const canGoPrev = currentStep > 0;

  const handleStepChange = (index: number) => {
    if (stepDisabled(index)) return;
    setCurrentStep(index);
  };

  const handleFileSelection = (fileList: FileList | null, setter: (file: File | null) => void) => {
    if (!fileList || fileList.length === 0) {
      setter(null);
      return;
    }
    setter(fileList[0]);
  };

  const addFixedUnitRow = () => {
    setFixedUnits((rows) => [
      ...rows,
      { id: crypto.randomUUID(), unitId: '', bands: '' },
    ]);
  };

  const updateFixedUnitRow = (id: string, field: keyof FixedUnitRow, value: string) => {
    setFixedUnits((rows) =>
      rows.map((row) => (row.id === id ? { ...row, [field]: value } : row))
    );
  };

  const removeFixedUnitRow = (id: string) => {
    setFixedUnits((rows) => rows.filter((row) => row.id !== id));
  };

  const addFloorRuleRow = () => {
    setFloorRules((rows) => [
      ...rows,
      { id: crypto.randomUUID(), minFloor: '', maxFloor: '', minBand: '' },
    ]);
  };

  const updateFloorRuleRow = (id: string, field: keyof FloorRuleRow, value: string) => {
    setFloorRules((rows) =>
      rows.map((row) => (row.id === id ? { ...row, [field]: value } : row))
    );
  };

  const removeFloorRuleRow = (id: string) => {
    setFloorRules((rows) => rows.filter((row) => row.id !== id));
  };

  const exportSettings = () => {
    const settings = {
      utilities,
      bandWhitelist,
      fixedUnits,
      floorRules,
      premiumWeights,
      notes,
      preferXlsb,
    };
    const blob = new Blob([JSON.stringify(settings, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'ami_optix_settings.json';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const importSettings = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result as string);
        if (parsed.utilities) setUtilities({ ...DEFAULT_UTILITIES, ...parsed.utilities });
        if (typeof parsed.bandWhitelist === 'string') setBandWhitelist(parsed.bandWhitelist);
        if (Array.isArray(parsed.fixedUnits)) setFixedUnits(parsed.fixedUnits);
        if (Array.isArray(parsed.floorRules)) setFloorRules(parsed.floorRules);
        if (parsed.premiumWeights) setPremiumWeights({ ...premiumWeights, ...parsed.premiumWeights });
        if (typeof parsed.notes === 'string') setNotes(parsed.notes);
        if (typeof parsed.preferXlsb === 'boolean') setPreferXlsb(parsed.preferXlsb);
      } catch (err) {
        setErrorMessage('Unable to import settings: invalid JSON file.');
      }
    };
    reader.readAsText(file);
    event.target.value = '';
  };

  const buildOverridesPayload = useCallback(() => {
    const whitelist = bandWhitelist
      .split(',')
      .map((item) => item.trim())
      .filter(Boolean)
      .map((item) => Number(item))
      .filter((num) => !Number.isNaN(num));

    const fixedUnitsPayload = fixedUnits
      .map((row) => {
        const bands = row.bands
          .split(',')
          .map((item) => item.trim())
          .filter(Boolean)
          .map((item) => Number(item))
          .filter((num) => !Number.isNaN(num));
        if (!row.unitId || bands.length === 0) return null;
        return {
          unitId: row.unitId.trim(),
          bands,
        };
      })
      .filter(Boolean);

    const floorRulesPayload = floorRules
      .map((row) => {
        const minBand = Number(row.minBand);
        if (Number.isNaN(minBand)) return null;
        const minFloor = row.minFloor ? Number(row.minFloor) : null;
        const maxFloor = row.maxFloor ? Number(row.maxFloor) : null;
        const payload: Record<string, number> = { minBand };
        if (minFloor !== null && !Number.isNaN(minFloor)) payload.minFloor = minFloor;
        if (maxFloor !== null && !Number.isNaN(maxFloor)) payload.maxFloor = maxFloor;
        return payload;
      })
      .filter(Boolean);

    const notesPayload = notes
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    return {
      bandWhitelist: whitelist.length ? whitelist : undefined,
      fixedUnits: fixedUnitsPayload.length ? fixedUnitsPayload : undefined,
      floorMinimums: floorRulesPayload.length ? floorRulesPayload : undefined,
      premiumWeights: premiumWeights,
      notes: notesPayload.length ? notesPayload : undefined,
    };
  }, [bandWhitelist, fixedUnits, floorRules, notes, premiumWeights]);

  const runAnalysis = async () => {
    if (!unitFile) {
      setErrorMessage('Please upload the unit schedule before running the solver.');
      return;
    }
    setErrorMessage(null);
    setIsSubmitting(true);
    setAnalysis(null);
    setSelectedScenarioKey(null);
    try {
      const formData = new FormData();
      formData.append('file', unitFile);
      if (rentFile) {
        formData.append('rentCalculator', rentFile);
      }
      formData.append('utilities', JSON.stringify(utilities));
      const overridesPayload = buildOverridesPayload();
      formData.append('overrides', JSON.stringify(overridesPayload));
      if (preferXlsb) {
        formData.append('preferXlsb', 'true');
      }

      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData,
      });
      const data = await response.json();
      if (!response.ok) {
        setErrorMessage(data.error || 'Solver returned an error.');
        return;
      }
      setAnalysis(data);
      const scenarioKey = Object.keys(data.scenarios || {}).find((key) => data.scenarios?.[key]);
      if (scenarioKey) {
        setSelectedScenarioKey(scenarioKey);
      }
      setCurrentStep(STEP_CONFIG.length - 1);
    } catch (err) {
      setErrorMessage('Request failed. Please check your network connection and try again.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const totalPremiumWeight = useMemo(() => {
    return (
      Number(premiumWeights.floor || 0) +
      Number(premiumWeights.net_sf || 0) +
      Number(premiumWeights.bedrooms || 0) +
      Number(premiumWeights.balcony || 0)
    );
  }, [premiumWeights]);

  const projectSummaryEntries = useMemo(() => {
    if (!analysis?.project_summary) return [];
    const summary = analysis.project_summary;
    return [
      { label: 'Affordable Units', value: summary.total_affordable_units },
      { label: 'Affordable SF', value: summary.total_affordable_sf },
      { label: '40% Units', value: summary.forty_percent_units },
      { label: '40% Share', value: summary.forty_percent_share ? `${(summary.forty_percent_share * 100).toFixed(2)}%` : '—' },
      { label: 'Total Monthly Rent', value: summary.total_monthly_rent ? `$${summary.total_monthly_rent.toLocaleString()}` : '—' },
      { label: 'Total Annual Rent', value: summary.total_annual_rent ? `$${summary.total_annual_rent.toLocaleString()}` : '—' },
    ];
  }, [analysis]);

  const currentScenario = useMemo(() => {
    if (!analysis?.scenarios || !selectedScenarioKey) return null;
    return analysis.scenarios[selectedScenarioKey];
  }, [analysis, selectedScenarioKey]);

  const handleDownloadReports = () => {
    if (!analysis?.download_link) return;
    window.open(analysis.download_link, '_blank');
  };

  return (
    <div className="min-h-screen bg-slate-100">
      <header className="bg-slate-900 text-white py-6 shadow-md">
        <div className="mx-auto max-w-6xl px-4">
          <h1 className="text-3xl font-semibold">NYC AMI Calculator Dashboard</h1>
          <p className="mt-2 text-sm text-slate-300">
            Step through the inputs, run the solver, review rent totals, and export the updated workbooks.
          </p>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-4 py-8">
        <nav className="grid gap-3 sm:grid-cols-5">
          {STEP_CONFIG.map((step, index) => {
            const active = index === currentStep;
            const completed = index < currentStep;
            const disabled = stepDisabled(index);
            return (
              <button
                key={step.title}
                onClick={() => handleStepChange(index)}
                disabled={disabled}
                className={`rounded-md border px-3 py-4 text-left transition-colors ${
                  active
                    ? 'border-slate-900 bg-white shadow-md'
                    : completed
                    ? 'border-emerald-500 bg-emerald-50 text-emerald-700'
                    : 'border-slate-200 bg-slate-50 text-slate-600'
                } ${disabled ? 'cursor-not-allowed opacity-50' : 'hover:border-slate-400 hover:shadow-sm'}`}
              >
                <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                  Step {index + 1}
                </span>
                <div className="mt-1 text-base font-medium">{step.title}</div>
                <div className="mt-1 text-xs text-slate-500">{step.subtitle}</div>
              </button>
            );
          })}
        </nav>

        <section className="mt-8 rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
          {currentStep === 0 && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-semibold">Upload Unit Schedule</h2>
                <p className="mt-1 text-sm text-slate-600">
                  Accepted formats: .xlsx, .xlsm, .xlsb, or .csv. We recommend the latest rent roll export from the client.
                </p>
              </div>
              <label className="flex cursor-pointer items-center justify-between rounded-md border border-dashed border-slate-300 bg-slate-50 p-6 hover:border-slate-400">
                <div>
                  <div className="text-lg font-medium text-slate-700">
                    {unitFile ? unitFile.name : 'Choose a file'}
                  </div>
                  <div className="text-xs text-slate-500">
                    Drag & drop or click to browse your files
                  </div>
                </div>
                <span className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700">
                  Browse
                </span>
                <input
                  type="file"
                  accept=".xlsx,.xlsm,.xlsb,.csv"
                  className="hidden"
                  onChange={(event) => handleFileSelection(event.target.files, setUnitFile)}
                />
              </label>
              {unitFile && (
                <div className="flex items-center gap-3 rounded-md bg-emerald-50 p-3 text-sm text-emerald-700">
                  <FiFileText className="text-lg" />
                  <span>Ready to analyze: {unitFile.name}</span>
                  <button
                    className="ml-auto text-xs font-medium underline"
                    onClick={() => setUnitFile(null)}
                  >
                    Remove
                  </button>
                </div>
              )}
            </div>
          )}

          {currentStep === 1 && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-semibold">Rent Calculator (Optional)</h2>
                <p className="mt-1 text-sm text-slate-600">
                  Upload the client-provided rent workbook (e.g., 2025 AMI Rent Calculator). If skipped, we'll use the default copy bundled with the app.
                </p>
              </div>
              <label className="flex cursor-pointer items-center justify-between rounded-md border border-dashed border-slate-300 bg-slate-50 p-6 hover:border-slate-400">
                <div>
                  <div className="text-lg font-medium text-slate-700">
                    {rentFile ? rentFile.name : 'Choose a workbook (optional)'}
                  </div>
                  <div className="text-xs text-slate-500">
                    Accepted formats: .xlsx, .xlsm, .xlsb
                  </div>
                </div>
                <span className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700">
                  Browse
                </span>
                <input
                  type="file"
                  accept=".xlsx,.xlsm,.xlsb"
                  className="hidden"
                  onChange={(event) => handleFileSelection(event.target.files, setRentFile)}
                />
              </label>
              {rentFile ? (
                <div className="flex items-center gap-3 rounded-md bg-blue-50 p-3 text-sm text-blue-700">
                  <FiFileText className="text-lg" />
                  <span>Custom rent workbook attached: {rentFile.name}</span>
                  <button
                    className="ml-auto text-xs font-medium underline"
                    onClick={() => setRentFile(null)}
                  >
                    Remove
                  </button>
                </div>
              ) : (
                <div className="rounded-md bg-slate-50 p-3 text-sm text-slate-600">
                  No workbook uploaded. The server will use the bundled 2025 calculator if needed.
                </div>
              )}
            </div>
          )}

          {currentStep === 2 && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-semibold">Utility Responsibilities</h2>
                <p className="mt-1 text-sm text-slate-600">
                  These answers populate row 17 of the client's rent calculator and determine the net rents (column H). If everything is owner-paid, choose N/A for each category.
                </p>
              </div>
              <div className="grid gap-4 sm:grid-cols-3">
                <div className="rounded-md border border-slate-200 p-4">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                    Cooking
                  </label>
                  <select
                    value={utilities.cooking}
                    onChange={(event) =>
                      setUtilities((prev) => ({ ...prev, cooking: event.target.value }))
                    }
                    className="mt-2 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                  >
                    {COOKING_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="rounded-md border border-slate-200 p-4">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                    Heat
                  </label>
                  <select
                    value={utilities.heat}
                    onChange={(event) =>
                      setUtilities((prev) => ({ ...prev, heat: event.target.value }))
                    }
                    className="mt-2 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                  >
                    {HEAT_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="rounded-md border border-slate-200 p-4">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                    Hot Water
                  </label>
                  <select
                    value={utilities.hot_water}
                    onChange={(event) =>
                      setUtilities((prev) => ({ ...prev, hot_water: event.target.value }))
                    }
                    className="mt-2 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                  >
                    {HOT_WATER_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          )}

          {currentStep === 3 && (
            <div className="space-y-8">
              <div className="grid gap-6 lg:grid-cols-2">
                <div className="space-y-4">
                  <div>
                    <h2 className="text-xl font-semibold">Band Whitelist</h2>
                    <p className="mt-1 text-sm text-slate-600">
                      Limit the solver to a specific set of AMI bands (comma separated). Leave blank to use the global defaults.
                    </p>
                    <input
                      type="text"
                      value={bandWhitelist}
                      onChange={(event) => setBandWhitelist(event.target.value)}
                      placeholder="e.g. 40, 60, 80"
                      className="mt-2 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                    />
                  </div>

                  <div>
                    <div className="flex items-center justify-between">
                      <h2 className="text-xl font-semibold">Pinned Units</h2>
                      <button
                        className="flex items-center gap-2 rounded-md border border-slate-300 px-3 py-1 text-sm font-medium text-slate-700 hover:bg-slate-50"
                        onClick={addFixedUnitRow}
                      >
                        <FiUpload className="text-sm" />
                        Add Unit
                      </button>
                    </div>
                    <p className="mt-1 text-sm text-slate-600">
                      Force specific units into target bands. Provide a unit ID and one or more allowed AMIs (comma separated).
                    </p>
                    <div className="mt-3 space-y-3">
                      {fixedUnits.length === 0 && (
                        <div className="rounded-md border border-dashed border-slate-300 p-3 text-sm text-slate-500">
                          No overrides yet. Click "Add Unit" to pin a unit.
                        </div>
                      )}
                      {fixedUnits.map((row) => (
                        <div
                          key={row.id}
                          className="grid gap-3 rounded-md border border-slate-200 p-3 sm:grid-cols-[1.5fr,1.5fr,auto]"
                        >
                          <input
                            type="text"
                            value={row.unitId}
                            onChange={(event) => updateFixedUnitRow(row.id, 'unitId', event.target.value)}
                            placeholder="Unit ID"
                            className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                          />
                          <input
                            type="text"
                            value={row.bands}
                            onChange={(event) => updateFixedUnitRow(row.id, 'bands', event.target.value)}
                            placeholder="Bands (e.g. 40, 60)"
                            className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                          />
                          <button
                            className="flex items-center justify-center rounded-md border border-red-200 bg-red-50 px-3 py-2 text-sm font-medium text-red-600 hover:bg-red-100"
                            onClick={() => removeFixedUnitRow(row.id)}
                          >
                            <FiTrash2 className="text-base" />
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="space-y-4">
                  <div>
                    <div className="flex items-center justify-between">
                      <h2 className="text-xl font-semibold">Floor Minimums</h2>
                      <button
                        className="flex items-center gap-2 rounded-md border border-slate-300 px-3 py-1 text-sm font-medium text-slate-700 hover:bg-slate-50"
                        onClick={addFloorRuleRow}
                      >
                        <FiUpload className="text-sm" />
                        Add Rule
                      </button>
                    </div>
                    <p className="mt-1 text-sm text-slate-600">
                      Require specific floors to stay above a minimum band. You can leave max floor blank for a single floor.
                    </p>
                    <div className="mt-3 space-y-3">
                      {floorRules.length === 0 && (
                        <div className="rounded-md border border-dashed border-slate-300 p-3 text-sm text-slate-500">
                          No floor constraints yet. Click "Add Rule" to get started.
                        </div>
                      )}
                      {floorRules.map((row) => (
                        <div
                          key={row.id}
                          className="grid gap-3 rounded-md border border-slate-200 p-3 sm:grid-cols-[repeat(3,1fr),auto]"
                        >
                          <input
                            type="number"
                            value={row.minFloor}
                            onChange={(event) => updateFloorRuleRow(row.id, 'minFloor', event.target.value)}
                            placeholder="Min floor"
                            className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                          />
                          <input
                            type="number"
                            value={row.maxFloor}
                            onChange={(event) => updateFloorRuleRow(row.id, 'maxFloor', event.target.value)}
                            placeholder="Max floor"
                            className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                          />
                          <input
                            type="number"
                            value={row.minBand}
                            onChange={(event) => updateFloorRuleRow(row.id, 'minBand', event.target.value)}
                            placeholder="Min AMI"
                            className="rounded-md border border-slate-300 px-3 py-2 text-sm"
                          />
                          <button
                            className="flex items-center justify-center rounded-md border border-red-200 bg-red-50 px-3 py-2 text-sm font-medium text-red-600 hover:bg-red-100"
                            onClick={() => removeFloorRuleRow(row.id)}
                          >
                            <FiTrash2 className="text-base" />
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div>
                    <h2 className="text-xl font-semibold">Premium Weights</h2>
                    <p className="mt-1 text-sm text-slate-600">
                      Adjust the tie-break recipe for the solver. Values do not need to add up to exactly 1.0—the backend normalizes them before use.
                    </p>
                    <div className="mt-3 grid gap-3 sm:grid-cols-2">
                      <div>
                        <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                          Floor
                        </label>
                        <input
                          type="number"
                          step="0.01"
                          value={premiumWeights.floor}
                          onChange={(event) =>
                            setPremiumWeights((prev) => ({ ...prev, floor: Number(event.target.value) }))
                          }
                          className="mt-1 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                          Net SF
                        </label>
                        <input
                          type="number"
                          step="0.01"
                          value={premiumWeights.net_sf}
                          onChange={(event) =>
                            setPremiumWeights((prev) => ({ ...prev, net_sf: Number(event.target.value) }))
                          }
                          className="mt-1 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                          Bedrooms
                        </label>
                        <input
                          type="number"
                          step="0.01"
                          value={premiumWeights.bedrooms}
                          onChange={(event) =>
                            setPremiumWeights((prev) => ({ ...prev, bedrooms: Number(event.target.value) }))
                          }
                          className="mt-1 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                          Balcony
                        </label>
                        <input
                          type="number"
                          step="0.01"
                          value={premiumWeights.balcony}
                          onChange={(event) =>
                            setPremiumWeights((prev) => ({ ...prev, balcony: Number(event.target.value) }))
                          }
                          className="mt-1 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                        />
                      </div>
                    </div>
                    <div className="mt-2 text-xs text-slate-500">
                      Current total: {totalPremiumWeight.toFixed(2)}
                    </div>
                  </div>

                  <div>
                    <h2 className="text-xl font-semibold">Notes to Solver</h2>
                    <p className="mt-1 text-sm text-slate-600">
                      Provide context or special instructions. One note per line—these flow into the solver diagnostics.
                    </p>
                    <textarea
                      value={notes}
                      onChange={(event) => setNotes(event.target.value)}
                      rows={4}
                      className="mt-2 w-full rounded-md border border-slate-300 px-3 py-2 text-sm"
                      placeholder="Example: Floors 10-12 must stay market-rate."
                    />
                  </div>

                  <div className="flex items-center gap-3 rounded-md border border-slate-200 bg-slate-50 p-4">
                    <input
                      id="prefer-xlsb"
                      type="checkbox"
                      checked={preferXlsb}
                      onChange={(event) => setPreferXlsb(event.target.checked)}
                      className="h-5 w-5 rounded border border-slate-300 text-slate-900 focus:ring-slate-900"
                    />
                    <label htmlFor="prefer-xlsb" className="text-sm font-medium text-slate-700">
                      Return downloads in .xlsb format when possible
                    </label>
                  </div>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-3 rounded-md border border-slate-200 bg-slate-50 p-4">
                <button
                  onClick={exportSettings}
                  className="flex items-center gap-2 rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100"
                >
                  <FiDownload className="text-base" />
                  Save Settings to JSON
                </button>
                <button
                  onClick={() => settingsInputRef.current?.click()}
                  className="flex items-center gap-2 rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100"
                >
                  <FiUpload className="text-base" />
                  Load Settings from JSON
                </button>
                <input
                  ref={settingsInputRef}
                  type="file"
                  accept="application/json"
                  className="hidden"
                  onChange={importSettings}
                />
                <span className="text-xs text-slate-500">
                  Settings include utilities, overrides, notes, premium weights, and export preference.
                </span>
              </div>
            </div>
          )}

          {currentStep === 4 && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-semibold">Review & Run</h2>
                <p className="mt-1 text-sm text-slate-600">
                  Double-check your selections, then launch the solver. Results include rent totals and 40% AMI compliance.
                </p>
              </div>

              <div className="grid gap-4 sm:grid-cols-2">
                <div className="rounded-md border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700">
                  <h3 className="text-sm font-semibold uppercase tracking-wide text-slate-500">
                    Uploads
                  </h3>
                  <ul className="mt-2 space-y-1">
                    <li>
                      <span className="font-medium text-slate-600">Unit schedule:</span>{' '}
                      {unitFile ? unitFile.name : '—'}
                    </li>
                    <li>
                      <span className="font-medium text-slate-600">Rent workbook:</span>{' '}
                      {rentFile ? rentFile.name : 'Bundled default'}
                    </li>
                    <li>
                      <span className="font-medium text-slate-600">Prefer XLSB:</span>{' '}
                      {preferXlsb ? 'Yes' : 'No'}
                    </li>
                  </ul>
                </div>
                <div className="rounded-md border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700">
                  <h3 className="text-sm font-semibold uppercase tracking-wide text-slate-500">
                    Utilities
                  </h3>
                  <ul className="mt-2 space-y-1">
                    <li>
                      <span className="font-medium text-slate-600">Cooking:</span>{' '}
                      {COOKING_OPTIONS.find((item) => item.value === utilities.cooking)?.label ?? utilities.cooking}
                    </li>
                    <li>
                      <span className="font-medium text-slate-600">Heat:</span>{' '}
                      {HEAT_OPTIONS.find((item) => item.value === utilities.heat)?.label ?? utilities.heat}
                    </li>
                    <li>
                      <span className="font-medium text-slate-600">Hot Water:</span>{' '}
                      {HOT_WATER_OPTIONS.find((item) => item.value === utilities.hot_water)?.label ?? utilities.hot_water}
                    </li>
                  </ul>
                </div>
              </div>

              <div className="rounded-md border border-dashed border-slate-300 bg-white p-4 text-sm text-slate-600">
                <div className="flex flex-wrap items-center gap-3">
                  <button
                    onClick={runAnalysis}
                    disabled={isSubmitting || !unitFile}
                    className="flex items-center gap-2 rounded-md bg-slate-900 px-5 py-2 text-sm font-semibold text-white hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
                  >
                    {isSubmitting ? (
                      <>
                        <FiRefreshCcw className="animate-spin text-base" />
                        Running...
                      </>
                    ) : (
                      <>
                        <FiUpload className="text-base" />
                        Run Solver
                      </>
                    )}
                  </button>
                  <span className="text-xs text-slate-500">
                    The solver enforces the 20–21% low-band window, applies overrides, and returns rent totals per unit.
                  </span>
                </div>
                {!unitFile && (
                  <div className="mt-2 text-xs text-red-600">
                    Upload the unit schedule before running the analysis.
                  </div>
                )}
              </div>

              {errorMessage && (
                <div className="rounded-md border border-red-200 bg-red-50 p-4 text-sm text-red-700">
                  {errorMessage}
                </div>
              )}

              {analysis && (
                <div className="space-y-6">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <h3 className="text-lg font-semibold text-slate-800">Project Summary</h3>
                    {analysis.download_link && (
                      <button
                        onClick={handleDownloadReports}
                        className="flex items-center gap-2 rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100"
                      >
                        <FiDownload className="text-base" />
                        Download Reports
                      </button>
                    )}
                  </div>
                  <div className="grid gap-3 sm:grid-cols-3">
                    {projectSummaryEntries.map((entry) => (
                      <div key={entry.label} className="rounded-md border border-slate-200 bg-slate-50 p-4">
                        <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                          {entry.label}
                        </div>
                        <div className="mt-1 text-lg font-semibold text-slate-800">
                          {entry.value ?? '—'}
                        </div>
                      </div>
                    ))}
                  </div>

                  {analysis.analysis_notes && analysis.analysis_notes.length > 0 && (
                    <div className="rounded-md border border-slate-200 bg-white p-4">
                      <h4 className="text-sm font-semibold uppercase tracking-wide text-slate-500">
                        Analysis Notes
                      </h4>
                      <ul className="mt-2 list-disc space-y-1 pl-5 text-sm text-slate-700">
                        {analysis.analysis_notes.map((note, index) => (
                          <li key={index}>{note}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  <div className="space-y-3">
                    <h3 className="text-lg font-semibold text-slate-800">Scenarios</h3>
                    <div className="flex flex-wrap gap-2">
                      {analysis.scenarios &&
                        Object.entries(analysis.scenarios).map(([key, scenario]) => {
                          if (!scenario) return null;
                          const label = SCENARIO_TITLES[key] ?? key;
                          const active = key === selectedScenarioKey;
                          return (
                            <button
                              key={key}
                              onClick={() => setSelectedScenarioKey(key)}
                              className={`rounded-md border px-4 py-2 text-sm font-medium ${
                                active
                                  ? 'border-slate-900 bg-slate-900 text-white'
                                  : 'border-slate-300 bg-white text-slate-700 hover:border-slate-400'
                              }`}
                            >
                              {label}
                            </button>
                          );
                        })}
                    </div>

                    {currentScenario ? (
                      <div className="rounded-md border border-slate-200 bg-white p-4">
                        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                          <SummaryCell label="WAAMI" value={`${(currentScenario.waami * 100).toFixed(2)}%`} />
                          <SummaryCell
                            label="Total Units"
                            value={currentScenario.metrics?.total_units ?? '—'}
                          />
                          <SummaryCell
                            label="Total Monthly Rent"
                            value={
                              currentScenario.metrics?.total_monthly_rent
                                ? `$${currentScenario.metrics.total_monthly_rent.toLocaleString()}`
                                : '—'
                            }
                          />
                          <SummaryCell
                            label="Total Annual Rent"
                            value={
                              currentScenario.metrics?.total_annual_rent
                                ? `$${currentScenario.metrics.total_annual_rent.toLocaleString()}`
                                : '—'
                            }
                          />
                          <SummaryCell
                            label="40% Units"
                            value={currentScenario.metrics?.low_band_units ?? '—'}
                          />
                          <SummaryCell
                            label="40% Share"
                            value={
                              currentScenario.metrics?.low_band_share !== undefined
                                ? `${(currentScenario.metrics.low_band_share * 100).toFixed(2)}%`
                                : '—'
                            }
                          />
                          <SummaryCell
                            label="Band Mix"
                            value={
                              currentScenario.metrics?.band_mix
                                ?.map(
                                  (entry: any) =>
                                    `${entry.band}% (${entry.units} units · ${(entry.share_of_sf * 100).toFixed(1)}% SF)`
                                )
                                .join('; ') || '—'
                            }
                            span
                          />
                        </div>

                        <div className="mt-6 overflow-auto rounded-md border border-slate-200">
                          <table className="min-w-full divide-y divide-slate-200 text-sm">
                            <thead className="bg-slate-50">
                              <tr className="text-left text-xs font-semibold uppercase tracking-wide text-slate-500">
                                <th className="px-3 py-2">Unit</th>
                                <th className="px-3 py-2">Bedrooms</th>
                                <th className="px-3 py-2">Net SF</th>
                                <th className="px-3 py-2">Floor</th>
                                <th className="px-3 py-2">Assigned AMI</th>
                                <th className="px-3 py-2">Monthly Rent</th>
                                <th className="px-3 py-2">Annual Rent</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-200 bg-white">
                              {currentScenario.assignments?.map((unit: any, index: number) => (
                                <tr key={`${unit.unit_id}-${index}`} className="text-slate-700">
                                  <td className="px-3 py-2 font-medium">{unit.unit_id}</td>
                                  <td className="px-3 py-2">{unit.bedrooms}</td>
                                  <td className="px-3 py-2">{unit.net_sf}</td>
                                  <td className="px-3 py-2">{unit.floor ?? '—'}</td>
                                  <td className="px-3 py-2">{`${(unit.assigned_ami * 100).toFixed(0)}%`}</td>
                                  <td className="px-3 py-2">
                                    {unit.monthly_rent !== undefined
                                      ? `$${unit.monthly_rent.toLocaleString()}`
                                      : '—'}
                                  </td>
                                  <td className="px-3 py-2">
                                    {unit.annual_rent !== undefined
                                      ? `$${unit.annual_rent.toLocaleString()}`
                                      : '—'}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    ) : (
                      <div className="rounded-md border border-dashed border-slate-300 p-4 text-sm text-slate-600">
                        Select a scenario to view the detailed assignments.
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          )}
        </section>

        <div className="mt-6 flex items-center justify-between">
          <button
            onClick={() => setCurrentStep((prev) => Math.max(prev - 1, 0))}
            disabled={!canGoPrev}
            className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100 disabled:cursor-not-allowed disabled:bg-slate-200"
          >
            Previous
          </button>
          <button
            onClick={() => setCurrentStep((prev) => Math.min(prev + 1, STEP_CONFIG.length - 1))}
            disabled={!canGoNext || (currentStep === 0 && !unitFile)}
            className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
          >
            Next
          </button>
        </div>
      </main>
    </div>
  );
}

function SummaryCell({
  label,
  value,
  span = false,
}: {
  label: string;
  value: string | number;
  span?: boolean;
}) {
  return (
    <div className={`rounded-md border border-slate-200 bg-slate-50 p-4 ${span ? 'sm:col-span-2 lg:col-span-4' : ''}`}>
      <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</div>
      <div className="mt-1 text-lg font-semibold text-slate-800">{value ?? '—'}</div>
    </div>
  );
}
