"use client";

import { useState, useMemo } from 'react';
import { FiUploadCloud, FiChevronDown, FiSun, FiMoon, FiCheckCircle, FiAlertTriangle, FiDownload } from 'react-icons/fi';

// --- Reusable Components ---

const ScenarioCard = ({ title, scenario }: { title: string, scenario: any }) => {
  if (!scenario) return null;

  const [isOpen, setIsOpen] = useState(true);

  return (
    <div className="border border-gray-200 dark:border-gray-700 rounded-lg mb-4 bg-gray-50/50 dark:bg-gray-800/20">
      <button
        className="w-full text-left p-4 bg-gray-100 dark:bg-gray-700/50 hover:bg-gray-200 dark:hover:bg-gray-600/50 flex justify-between items-center rounded-t-lg"
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="flex items-center">
          <h4 className="font-semibold text-lg text-gray-800 dark:text-gray-200">{title}</h4>
        </div>
        <FiChevronDown className={`transform transition-transform text-gray-500 dark:text-gray-400 ${isOpen ? 'rotate-180' : ''}`} />
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
              <p className="text-lg font-medium">{scenario.bands.join('%, ')}%</p>
            </div>
          </div>
          <AssignmentsTable assignments={scenario.assignments} />
        </div>
      )}
    </div>
  );
};

const AssignmentsTable = ({ assignments }: { assignments: any[] }) => {
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
          {assignments.map((unit, index) => (
            <tr key={index} className="bg-white border-b dark:bg-gray-800 dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-600/20">
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

const ComplianceReport = ({ report }: { report: any[] }) => {
  const flagged = report.filter(r => r.status === 'FLAGGED');
  if (flagged.length === 0) {
    return (
      <div className="flex items-center text-green-500">
        <FiCheckCircle className="mr-2" />
        <p>All compliance checks passed.</p>
      </div>
    );
  }
  return (
    <div className="space-y-2">
      {flagged.map((item, i) => (
        <div key={i} className="flex items-start text-yellow-500 text-sm">
          <FiAlertTriangle className="mr-2 mt-1 flex-shrink-0" />
          <span><strong>{item.check}:</strong> {item.details}</span>
        </div>
      ))}
    </div>
  );
};


// --- Main Page Component ---

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);
  const [isDarkMode, setIsDarkMode] = useState(true);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setFile(e.target.files[0]);
      setError(null);
      setAnalysisResult(null);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFile(e.dataTransfer.files[0]);
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
    formData.append('file', file);

    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || 'An unknown error occurred during analysis.');
      }

      setAnalysisResult(result);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsUploading(false);
    }
  };

  return (
    <div className={isDarkMode ? 'dark' : ''}>
      <main className="bg-gray-50 dark:bg-gray-900 min-h-screen text-gray-900 dark:text-gray-100 transition-colors duration-300 font-sans">
        <div className="container mx-auto p-4 md:p-8">

          <header className="flex justify-between items-center mb-8 pb-4 border-b border-gray-200 dark:border-gray-700">
            <h1 className="text-4xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-blue-500 to-teal-400">
              AMI-Optix
            </h1>
            <button onClick={() => setIsDarkMode(!isDarkMode)} className="p-2 rounded-full text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700">
              {isDarkMode ? <FiSun size={20}/> : <FiMoon size={20} />}
            </button>
          </header>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">

            <div className="lg:col-span-1 bg-white dark:bg-gray-800/50 rounded-xl shadow-lg p-6 border border-gray-200 dark:border-gray-700 self-start">
              <h2 className="text-2xl font-semibold mb-4 text-gray-800 dark:text-gray-200">Controls</h2>

              <div
                className="border-2 border-dashed border-gray-300 dark:border-gray-600 rounded-lg p-8 text-center cursor-pointer hover:border-blue-500 dark:hover:border-blue-400 transition-colors"
                onDrop={handleDrop}
                onDragOver={(e) => e.preventDefault()}
                onClick={() => document.getElementById('file-upload')?.click()}
              >
                <FiUploadCloud className="mx-auto text-4xl text-gray-400 mb-2" />
                <p className="text-gray-500 dark:text-gray-400">
                  {file ? <span className="text-blue-500">{file.name}</span> : 'Drag & drop file here'}
                </p>
                <input id="file-upload" type="file" className="hidden" onChange={handleFileChange} accept=".xlsx, .csv" />
              </div>

              <button
                onClick={handleRunAnalysis}
                disabled={isUploading || !file}
                className="w-full mt-6 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-4 rounded-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isUploading ? 'Analyzing...' : 'Run Analysis'}
              </button>

              {error && <p className="text-red-500 mt-4 text-sm">{error}</p>}

              {analysisResult && analysisResult.download_link && (
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
                      {analysisResult.narrative_analysis || "No summary generated."}
                    </div>
                  </div>

                  <div>
                    <h3 className="text-xl font-semibold mb-2">Compliance Report</h3>
                    <div className="bg-gray-100 dark:bg-gray-900/70 p-4 rounded-md">
                      <ComplianceReport report={analysisResult.compliance_report || []} />
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
