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
      if (data.length === 0) {
        throw new Error('No data was extracted. Please verify the file format.');
      }
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

    // Data storage
    const stabilityData = {};
    const draftData = {};

    // State variables
    let currentCase = null;
    let inStabilitySummary = false;
    let inDraftMarks = false;
    let tempHeel = null;
    let tempTrim = null;
    let tempDrafts = {};

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      // Detect damage case
      if (line.includes('DAMAGE STABILITY Case-')) {
        const match = line.match(/Case-(\d+)\s*\(Compartment\s+(\w+)\s+Flooded\)/);
        if (match) {
          currentCase = match[1];
        }
      } else if (line.includes('INTACT TOW CONDITION')) {
        currentCase = 'Intact';
      } else if (line.includes('Damage = NONE') && line.includes('VCG')) {
        currentCase = 'Intact';
      } else if (line.includes('Damage =') && line.includes('VCG')) {
        const match = line.match(/Damage = (\w+)\s/);
        if (match) {
          const damageId = match[1];
          if (damageId !== 'NONE!') {
            const caseMatch = damageId.match(/(\d+)/);
            if (caseMatch) {
              currentCase = caseMatch[1];
            }
          }
        }
      }

      // Detect draft marks
      if (line.includes('+++ D R A F T   M A R K   R E A D I N G S +++')) {
        inDraftMarks = true;
        tempDrafts = {};
        continue;
      }

      // Parse draft marks
      if (inDraftMarks && line.includes('AFT(P)') && line.includes('AFT(S)')) {
        const pattern = /(\w+\([PS]\))\s+([\d.]+)/g;
        let match;
        while ((match = pattern.exec(line)) !== null) {
          const name = match[1];
          if (name !== 'MEAN(P)' && name !== 'MEAN(S)') {
            tempDrafts[name] = parseFloat(match[2]);
          }
        }

        if (currentCase) {
          draftData[currentCase] = { ...tempDrafts };
        }

        inDraftMarks = false;
      }

      // Detect stability summary
      if (line.includes('+++ S T A B I L I T Y   S U M M A R Y +++')) {
        inStabilitySummary = true;
        tempHeel = null;
        tempTrim = null;
        continue;
      }

      if (inStabilitySummary) {
        // Extract Roll (Heel)
        if (line.includes('Roll') && line.includes('Deg')) {
          const rollMatch = line.match(/Roll\s*=\s*([-\d.]+)\s*Deg/);
          if (rollMatch) tempHeel = parseFloat(rollMatch[1]);
        }

        // Extract Pitch (Trim)
        if (line.includes('Pitch') && line.includes('Deg')) {
          const pitchMatch = line.match(/Pitch\s*=\s*([-\d.]+)\s*Deg/);
          if (pitchMatch) tempTrim = parseFloat(pitchMatch[1]);
        }

        // Extract Area Ratio
        if (line.includes('Area Ratio') && line.includes('Passes')) {
          const match = line.match(/Area Ratio\s+>=\s+([\d.]+)\s+([\d.]+)\s+Passes/);
          if (match && currentCase) {
            stabilityData[currentCase] = {
              heel: tempHeel !== null ? tempHeel : 0.0,
              trim: tempTrim !== null ? tempTrim : 0.0,
              areaRatioRequired: parseFloat(match[1]),
              areaRatioActual: parseFloat(match[2])
            };
            inStabilitySummary = false;
          }
        }
      }
    }

    // Combine data
    // Sort keys to maintain order (Intact first, then numeric cases)
    const sortedCases = Object.keys(stabilityData).sort((a, b) => {
      if (a === 'Intact') return -1;
      if (b === 'Intact') return 1;
      return parseInt(a) - parseInt(b);
    });

    for (const caseId of sortedCases) {
      if (draftData[caseId]) {
        results.push({
          case: caseId,
          drafts: draftData[caseId],
          heel: stabilityData[caseId].heel,
          trim: stabilityData[caseId].trim,
          areaRatioActual: stabilityData[caseId].areaRatioActual,
          areaRatioRequired: stabilityData[caseId].areaRatioRequired,
          remarks: 'Pass'
        });
      }
    }

    return results;
  };

  const downloadExcel = () => {
    if (!extractedData) return;

    try {
      // Define headers
      const excelHeader = [
        ['', 'Draft Mark (m)', '', '', '', 'Heel', 'Trim', 'Wind Area Ratio', '', ''],
        ['Case No.', 'Aft Port', 'Aft Stbd', 'Fwd Port', 'Fwd Stbd', '(deg)', '(deg)', 'Actual', 'Required', 'Remarks']
      ];

      // Map data
      const excelData = extractedData.map(row => [
        row.case,
        parseFloat(row.drafts['AFT(P)']?.toFixed(2) || 0),
        parseFloat(row.drafts['AFT(S)']?.toFixed(2) || 0),
        parseFloat(row.drafts['FWD(P)']?.toFixed(2) || 0),
        parseFloat(row.drafts['FWD(S)']?.toFixed(2) || 0),
        parseFloat(row.heel.toFixed(2)),
        parseFloat(row.trim.toFixed(2)),
        parseFloat(row.areaRatioActual.toFixed(2)),
        parseFloat(row.areaRatioRequired.toFixed(1)),
        row.remarks
      ]);

      // Combine for worksheet
      const wsData = [...excelHeader, ...excelData];

      // Create workbook and worksheet
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      // Merge cells for headers
      if (!ws['!merges']) ws['!merges'] = [];
      ws['!merges'].push(
        { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Case No
        { s: { r: 0, c: 1 }, e: { r: 0, c: 4 } }, // Draft Marks
        { s: { r: 0, c: 5 }, e: { r: 1, c: 5 } }, // Heel
        { s: { r: 0, c: 6 }, e: { r: 1, c: 6 } }, // Trim
        { s: { r: 0, c: 7 }, e: { r: 0, c: 8 } }, // Wind Area Ratio
        { s: { r: 0, c: 9 }, e: { r: 1, c: 9 } }  // Remarks
      );

      // Set column widths
      ws['!cols'] = [
        { wch: 10 }, // Case
        { wch: 10 }, // AFT P
        { wch: 10 }, // AFT S
        { wch: 10 }, // FWD P
        { wch: 10 }, // FWD S
        { wch: 8 },  // Heel
        { wch: 8 },  // Trim
        { wch: 12 }, // Actual
        { wch: 10 }, // Required
        { wch: 10 }  // Remarks
      ];

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(wb, ws, 'Damage Stability Results');

      // Generate Excel file and trigger download
      XLSX.writeFile(wb, 'Damage_Stability_Results.xlsx');
    } catch (err) {
      setError('Error creating Excel file: ' + err.message);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-8">
      <div className="max-w-6xl mx-auto">
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
                      <th rowSpan="2" className="px-4 py-3 text-center text-xs font-bold text-gray-500 uppercase tracking-wider border-b">Case No.</th>
                      <th colSpan="4" className="px-4 py-3 text-center text-xs font-bold text-gray-500 uppercase tracking-wider border-b">Draft Mark (m)</th>
                      <th rowSpan="2" className="px-4 py-3 text-center text-xs font-bold text-gray-500 uppercase tracking-wider border-b">Heel<br />(deg)</th>
                      <th rowSpan="2" className="px-4 py-3 text-center text-xs font-bold text-gray-500 uppercase tracking-wider border-b">Trim<br />(deg)</th>
                      <th colSpan="2" className="px-4 py-3 text-center text-xs font-bold text-gray-500 uppercase tracking-wider border-b">Wind Area Ratio</th>
                      <th rowSpan="2" className="px-4 py-3 text-center text-xs font-bold text-gray-500 uppercase tracking-wider border-b">Remarks</th>
                    </tr>
                    <tr>
                      <th className="px-4 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-100">Aft Port</th>
                      <th className="px-4 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-100">Aft Stbd</th>
                      <th className="px-4 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-100">Fwd Port</th>
                      <th className="px-4 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-100">Fwd Stbd</th>
                      <th className="px-4 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-100">Actual</th>
                      <th className="px-4 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-100">Required</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {extractedData.map((row, idx) => (
                      <tr key={idx} className={idx % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                        <td className="px-4 py-3 whitespace-nowrap text-sm font-bold text-center text-gray-900">{row.case}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.drafts['AFT(P)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.drafts['AFT(S)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.drafts['FWD(P)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.drafts['FWD(S)']?.toFixed(2) || '-'}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.heel.toFixed(2)}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.trim.toFixed(2)}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.areaRatioActual.toFixed(2)}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-gray-900">{row.areaRatioRequired.toFixed(1)}</td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-center text-green-600 font-bold">{row.remarks}</td>
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
              <li>Review the extracted damage cases, stability data, and draft marks</li>
              <li>Click "Download Excel File" to save the results</li>
            </ol>
          </div>
        </div>
      </div>
    </div>
  );
}
