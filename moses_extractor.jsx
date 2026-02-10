/**
 * MOSES Value Extraction - Web Interface
 * 
 * Installation:
 * npm install xlsx lucide-react
 * 
 * This component works entirely client-side - no backend required!
 */
import React, { useState } from 'react';
import { Upload, Download, FileText, CheckCircle, AlertCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

export default function MosesExtractor() {
  const [file, setFile] = useState(null);
  const [extractedData, setExtractedData] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState('');

  const handleFileUpload = (e) => {
    const uploadedFile = e.target.files[0];
    if (uploadedFile) {
      setFile(uploadedFile);
      setError('');
      setExtractedData(null);
    }
  };

  const extractMosesData = async () => {
    if (!file) {
      setError('Please upload a file first');
      return;
    }

    setProcessing(true);
    setError('');

    try {
      const text = await file.text();
      const data = parseMosesFile(text);
      setExtractedData(data);
    } catch (err) {
      setError('Error processing file: ' + err.message);
    } finally {
      setProcessing(false);
    }
  };

  const parseMosesFile = (text) => {
    const results = [];
    const lines = text.split('\n');

    // Find all stability summary sections
    let currentDamage = null;
    let currentGM = null;
    let currentDrafts = {};
    let inStabilitySummary = false;
    let inDraftMarks = false;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      // Detect damage case from section headers
      if (line.includes('DAMAGE STABILITY Case-')) {
        const match = line.match(/Case-(\d+)\s*\(Compartment\s+(\w+)\s+Flooded\)/);
        if (match) {
          currentDamage = match[2];
        }
      } else if (line.includes('INTACT TOW CONDITION') || line.includes('Damage = NONE')) {
        currentDamage = 'NONE';
      }

      // Detect draft mark readings section
      if (line.includes('+++ D R A F T   M A R K   R E A D I N G S +++')) {
        inDraftMarks = true;
        currentDrafts = {};
        continue;
      }

      // Parse draft marks
      if (inDraftMarks && line.includes('AFT(P)')) {
        const draftLine = line.replace(/\s+/g, ' ');
        const parts = draftLine.split(' ');

        for (let j = 0; j < parts.length; j += 2) {
          if (parts[j] && parts[j + 1] && parts[j] !== 'Name' && parts[j] !== 'Draft') {
            const name = parts[j];
            const value = parseFloat(parts[j + 1]);
            if (!isNaN(value) && name !== 'MEAN(P)' && name !== 'MEAN(S)') {
              currentDrafts[name] = value;
            }
          }
        }
        inDraftMarks = false;
      }

      // Detect stability summary section
      if (line.includes('+++ S T A B I L I T Y   S U M M A R Y +++')) {
        inStabilitySummary = true;
        continue;
      }

      // Extract GM value from stability summary
      if (inStabilitySummary && line.includes('GM') && line.includes('>=')) {
        const match = line.match(/GM\s+>=\s+[\d.]+\s+M\s+([\d.]+)/);
        if (match) {
          currentGM = parseFloat(match[1]);

          // Store the complete record
          if (currentDamage !== null && Object.keys(currentDrafts).length > 0) {
            results.push({
              damage: currentDamage,
              gm: currentGM,
              drafts: { ...currentDrafts }
            });
          }

          inStabilitySummary = false;
        }
      }
    }

    return results;
  };

  const downloadExcel = () => {
    if (!extractedData) return;

    try {
      // Create workbook and worksheet
      const wb = XLSX.utils.book_new();

      // Prepare data for Excel
      const excelData = extractedData.map(row => ({
        'Damage': row.damage,
        'GM (m)': row.gm.toFixed(2),
        'AFT(P) (m)': row.drafts['AFT(P)']?.toFixed(2) || '-',
        'AFT(S) (m)': row.drafts['AFT(S)']?.toFixed(2) || '-',
        'FWD(P) (m)': row.drafts['FWD(P)']?.toFixed(2) || '-',
        'FWD(S) (m)': row.drafts['FWD(S)']?.toFixed(2) || '-'
      }));

      // Create worksheet from data
      const ws = XLSX.utils.json_to_sheet(excelData);

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(wb, ws, 'MOSES Data Extraction');

      // Generate Excel file and trigger download
      XLSX.writeFile(wb, 'moses_extraction.xlsx');
    } catch (err) {
      setError('Error creating Excel file: ' + err.message);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-8">
      <div className="max-w-4xl mx-auto">
        <div className="bg-white rounded-lg shadow-xl p-8">
          <div className="flex items-center gap-3 mb-8">
            <FileText className="w-10 h-10 text-indigo-600" />
            <h1 className="text-3xl font-bold text-gray-800">MOSES Value Extraction</h1>
          </div>

          {/* File Upload Section */}
          <div className="mb-8">
            <label className="block mb-4">
              <div className="flex items-center gap-2 mb-2">
                <Upload className="w-5 h-5 text-gray-600" />
                <span className="text-lg font-semibold text-gray-700">Upload MOSES File</span>
              </div>
              <input
                type="file"
                accept=".txt,.out"
                onChange={handleFileUpload}
                className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer"
              />
            </label>

            {file && (
              <div className="flex items-center gap-2 text-green-600 mt-2">
                <CheckCircle className="w-5 h-5" />
                <span className="text-sm">File loaded: {file.name}</span>
              </div>
            )}
          </div>

          {/* Extract Button */}
          <button
            onClick={extractMosesData}
            disabled={!file || processing}
            className="w-full bg-indigo-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-indigo-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors mb-6"
          >
            {processing ? 'Processing...' : 'Extract Data'}
          </button>

          {/* Error Message */}
          {error && (
            <div className="flex items-center gap-2 bg-red-50 border border-red-200 text-red-700 p-4 rounded-lg mb-6">
              <AlertCircle className="w-5 h-5" />
              <span>{error}</span>
            </div>
          )}

          {/* Results Table */}
          {extractedData && extractedData.length > 0 && (
            <div className="mb-6">
              <h2 className="text-xl font-bold text-gray-800 mb-4">Extracted Data</h2>
              <div className="overflow-x-auto border rounded-lg">
                <table className="min-w-full divide-y divide-gray-200">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Damage</th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">GM (m)</th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">AFT(P)</th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">AFT(S)</th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">FWD(P)</th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">FWD(S)</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {extractedData.map((row, idx) => (
                      <tr key={idx} className={idx % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                        <td className="px-4 py-3 whitespace-nowrap text-sm font-medium text-gray-900">{row.damage}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">{row.gm.toFixed(2)}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">{row.drafts['AFT(P)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">{row.drafts['AFT(S)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">{row.drafts['FWD(P)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">{row.drafts['FWD(S)']?.toFixed(2) || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <button
                onClick={downloadExcel}
                className="mt-6 w-full bg-green-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-green-700 transition-colors flex items-center justify-center gap-2"
              >
                <Download className="w-5 h-5" />
                Download Excel File
              </button>
            </div>
          )}

          {/* Instructions */}
          <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-6">
            <h3 className="text-lg font-semibold text-blue-900 mb-3">Instructions:</h3>
            <ol className="list-decimal list-inside space-y-2 text-sm text-blue-800">
              <li>Upload your MOSES output file (.txt or .out format)</li>
              <li>Click "Extract Data" to process the file</li>
              <li>Review the extracted damage cases, GM values, and draft marks</li>
              <li>Click "Download Excel File" to save the results</li>
            </ol>
          </div>
        </div>
      </div>
    </div>
  );
}
