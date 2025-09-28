"use client";

import { useEffect, useState } from "react";
import type { ChangeEvent, DragEvent } from "react";
import {
  FiUploadCloud,
  FiChevronDown,
  FiSun,
  FiMoon,
  FiDownload
} from "react-icons/fi";

type Assignment = {
  unit_id: string;
  bedrooms: number;
  net_sf: number;
  assigned_ami: number;
};

type Scenario = {
  waami: number;
  bands: number[];
  assignments: Assignment[];
};

type ComplianceItem = {
  status: string;
  check: string;
  details: string;
};

interface AnalysisResult {
  narrative_analysis?: string;
  compliance_report?: ComplianceItem[];
  scenario_absolute_best?: Scenario | null;
  scenario_client_oriented?: Scenario | null;
  scenario_best_2_band?: Scenario | null;
  scenario_alternative?: Scenario | null;
  download_link?: string;
}

// --- Reusable Components ---

const ScenarioCard = ({ title, scenario }: { title: string; scenario?: Scenario | null }) => {
  const [isOpen, setIsOpen] = useState(true);

  if (!scenario) return null;

  const toggle = () => setIsOpen((prev) => !prev);

  return (
    <div className="border border-gray-200 dark:border-gray-700 rounded-lg mb-4 bg-gray-50/50 dark:bg-gray-800/20">
      <button
        className="w-full text-left p-4 bg-gray-100 dark:bg-gray-700/50 hover:bg-gray-200 dark:hover:bg-gray-600/50 flex justify-between items-center rounded-t-lg"
        onClick={toggle}
        type="button"
      >
        <div className="flex items-center">
          <h4 className="font-semibold text-lg text-gray-800 dark:text-gray-200">{title}</h4>
        </div>
        <FiChevronDown
          className={`transform transition-transform text-gray-500 dark:text-gray-400 ${isOpen ? "rotate-180" : ""}`}
        />
      </button>
      {isOpen && (
        <div className="p-4">
          <div className="grid grid-cols-2 gap-4 mb-4 text-sm">
            <div>
              <p className="font-semibold text-gray-500 dark:text-gray-400">WAAMI</p>
              <p className="text-2xl font-bold text-blue-500 dark:text-blue-400">{(scenario.waami * 100).toFixed(4)}%</p>
            </div>
            <div>
              <p className="font-semibold text-gray-500 dark:text-gray-400">Bands Used</p>
              <p className="text-lg font-medium">{scenario.bands.join("%, ")}%</p>
            </div>
          </div>
          <AssignmentsTable assignments={scenario.assignments} />
        </div>
      )}
    </div>
  );
};

const AssignmentsTable = ({ assignments }: { assignments?: Assignment[] }) => {
  const rows = assignments ?? [];

  return (
    <div className="overflow-x-auto max-h-60 border dark:border-gray-700 rounded-md">
      <table className="w-full text-sm text-left">
        <thead className="text-xs text-gray-700 uppercase bg-gray-100 dark:bg-gray-700 dark:text-gray-400">
          <tr>
            <th scope="col" className="px-4 py-2">Unit ID</th>
            <th scope="col" className="px-4 py-2">Bedrooms</th>
            <th scope="col" className="px-4 py-2">Net SF</th>
            <th scope="col" className="px-4 py-2">Assigned AMI</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((unit) => (
            <tr
              key={unit.unit_id}
              className="bg-white border-b dark:bg-gray-800 dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-600/20"
            >
              <td className="px-4 py-2 font-medium">{unit.unit_id}</td>
              <td className="px-4 py-2">{unit.bedrooms}</td>
              <td className="px-4 py-2">{unit.net_sf}</td>
              <td className="px-4 py-2 font-bold text-blue-600 dark:text-blue-400">{(unit.assigned_ami * 100).toFixed(0)}%</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};


const getInitialTheme = () => {
  if (typeof window === "undefined") return false;
  const stored = window.localStorage.getItem("ami-optix-theme");
  if (stored === "dark") return true;
  if (stored === "light") return false;
  return window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
};

// --- Main Page Component ---

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isDarkMode, setIsDarkMode] = useState(getInitialTheme);
  const [isUploading, setIsUploading] = useState(false);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    const root = document.documentElement;
    if (isDarkMode) {
      root.classList.add('dark');
    } else {
      root.classList.remove('dark');
    }
    window.localStorage.setItem('ami-optix-theme', isDarkMode ? 'dark' : 'light');
  }, [isDarkMode]);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (event.target.files?.[0]) {
      setFile(event.target.files[0]);
      setError(null);
      setAnalysisResult(null);
    }
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    if (event.dataTransfer.files?.[0]) {
      setFile(event.dataTransfer.files[0]);
      setError(null);
      setAnalysisResult(null);
    }
  };

  const handleRunAnalysis = async () => {
    if (!file) {
      setError("Please select a file first.");
      return;
    }

    setIsUploading(true);
    setError(null);
    setAnalysisResult(null);

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        body: formData
      });

      const result = (await response.json()) as AnalysisResult & { error?: string };

      if (!response.ok) {
        throw new Error(result.error ?? "An unknown error occurred during analysis.");
      }

      setAnalysisResult(result);
    } catch (err) {
      const message = err instanceof Error ? err.message : "An unexpected error occurred.";
      setError(message);
    } finally {
      setIsUploading(false);
    }
  };

  return (
    <div className={isDarkMode ? "dark" : ""}>
      <main className="bg-gray-50 dark:bg-gray-900 min-h-screen text-gray-900 dark:text-gray-100 transition-colors duration-300 font-sans">
        <div className="container mx-auto p-4 md:p-8">
          <header className="flex justify-between items-center mb-8 pb-4 border-b border-gray-200 dark:border-gray-700">
            <h1 className="text-4xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-blue-500 to-teal-400">
              AMI-Optix
            </h1>
            <button
              onClick={() => setIsDarkMode((prev) => !prev)}
              className="p-2 rounded-full text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700"
              type="button"
            >
              {isDarkMode ? <FiSun size={20} /> : <FiMoon size={20} />}
            </button>
          </header>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
            <div className="lg:col-span-1 bg-white dark:bg-gray-800/50 rounded-xl shadow-lg p-6 border border-gray-200 dark:border-gray-700 self-start">
              <h2 className="text-2xl font-semibold mb-4 text-gray-800 dark:text-gray-200">Controls</h2>

              <div
                className="border-2 border-dashed border-gray-300 dark:border-gray-600 rounded-lg p-8 text-center cursor-pointer hover:border-blue-500 dark:hover:border-blue-400 transition-colors"
                onDrop={handleDrop}
                onDragOver={(event) => event.preventDefault()}
                onClick={() => document.getElementById("file-upload")?.click()}
                role="button"
                tabIndex={0}
                onKeyPress={() => document.getElementById("file-upload")?.click()}
              >
                <FiUploadCloud className="mx-auto text-4xl text-gray-400 mb-2" />
                <p className="text-gray-500 dark:text-gray-400">
                  {file ? <span className="text-blue-500">{file.name}</span> : "Drag & drop file here"}
                </p>
                <input
                  id="file-upload"
                  type="file"
                  className="hidden"
                  onChange={handleFileChange}
                  accept=".xlsx, .csv"
                />
              </div>

              <button
                onClick={handleRunAnalysis}
                disabled={isUploading || !file}
                className="w-full mt-6 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-4 rounded-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed"
                type="button"
              >
                {isUploading ? "Analyzing..." : "Run Analysis"}
              </button>

              {error && <p className="text-red-500 mt-4 text-sm">{error}</p>}

              {analysisResult?.download_link && (
                <a
                  href={analysisResult.download_link}
                  className="w-full mt-4 bg-teal-600 hover:bg-teal-700 text-white font-bold py-3 px-4 rounded-lg transition-all flex items-center justify-center"
                >
                  <FiDownload className="mr-2" />
                  Download Reports (.zip)
                </a>
              )}
            </div>

            <div className="lg:col-span-2 bg-white dark:bg-gray-800/50 rounded-xl shadow-lg p-6 border border-gray-200 dark:border-gray-700">
              <h2 className="text-2xl font-semibold mb-4 text-gray-800 dark:text-gray-200">Analysis</h2>
              {!isUploading && !analysisResult && (
                <div className="text-center py-16 text-gray-500 dark:text-gray-400">
                  <p>Upload a file and run the analysis to see the results.</p>
                </div>
              )}
              {isUploading && (
                <div className="text-center py-16 text-gray-500 dark:text-gray-400">
                  <p>Running optimization... This may take a moment.</p>
                </div>
              )}
              {analysisResult && (
                <div className="space-y-6">
                  <div>
                    <h3 className="text-xl font-semibold mb-2">Internal Summary</h3>
                    <div className="text-gray-600 dark:text-gray-300 bg-gray-100 dark:bg-gray-900/70 p-4 rounded-md text-sm leading-relaxed whitespace-pre-wrap">
                      {analysisResult.narrative_analysis ?? "No summary generated."}
                    </div>
                  </div>

                  <div>
                    <h3 className="text-xl font-semibold mb-2">Scenario Details</h3>
                    <div className="space-y-4">
                      <ScenarioCard title="Absolute Best" scenario={analysisResult.scenario_absolute_best} />
                      <ScenarioCard title="Client Oriented" scenario={analysisResult.scenario_client_oriented} />
                      <ScenarioCard title="Best 2-Band" scenario={analysisResult.scenario_best_2_band} />
                      <ScenarioCard title="Alternative" scenario={analysisResult.scenario_alternative} />
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}
