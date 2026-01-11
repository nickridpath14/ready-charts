import React, { useState, useRef, useEffect } from 'react';
import { 
  LineChart, Line, BarChart, Bar, PieChart, Pie, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell 
} from 'recharts';
import { 
  Download, Plus, Trash2, BarChart2, PieChart as PieIcon, 
  Activity, Layers, FileImage, FileText, Settings, Palette, Upload,
  Calendar, Zap, ArrowRightLeft, AlignLeft
} from 'lucide-react';

// --- Utility: Script Loader for External Libs ---
const useScript = (url) => {
  useEffect(() => {
    const script = document.createElement('script');
    script.src = url;
    script.async = true;
    document.body.appendChild(script);
    return () => {
      document.body.removeChild(script);
    };
  }, [url]);
};

// --- Utility: Color Generator ---
// Converts a hex color to an RGBA array for shade generation
const hexToRgb = (hex) => {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : { r: 129, g: 140, b: 248 }; // Default indigo-400
};

// Generates a palette of shades based on the primary color
const generateShades = (baseColor, count) => {
  const rgb = hexToRgb(baseColor);
  const shades = [];
  for (let i = 0; i < count; i++) {
    // Decrease opacity/saturation to create distinct steps
    // We use RGBA to handle transparency for a modern look
    const opacity = Math.max(0.4, 1 - (i * (0.6 / count))); // Adjusted opacity range for better visibility
    shades.push(`rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${opacity})`);
  }
  return shades;
};

// --- Main Application ---
export default function ReadyCharts() {
  // Load Libraries for Export and Excel Import
  useScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
  useScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  useScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

  // --- State ---
  const [chartType, setChartType] = useState('bar');
  const [barLayout, setBarLayout] = useState('horizontal'); // 'horizontal' or 'vertical'
  const [chartTitle, setChartTitle] = useState('Project Timeline');
  const [xAxisLabel, setXAxisLabel] = useState('Date');
  const [yAxisLabel, setYAxisLabel] = useState('Metric Value');
  const [primaryColor, setPrimaryColor] = useState('#818cf8'); // Indigo-400
  const [curveType, setCurveType] = useState('monotone'); // monotone, linear, step
  const [data, setData] = useState([
    { name: '2023-01', value: 4000 },
    { name: '2023-02', value: 3000 },
    { name: '2023-03', value: 2000 },
    { name: '2023-04', value: 2780 },
    { name: '2023-05', value: 1890 },
    { name: '2023-06', value: 2390 },
    { name: '2023-07', value: 3490 },
  ]);

  const chartRef = useRef(null);
  const fileInputRef = useRef(null);
  const [isExporting, setIsExporting] = useState(false);

  // --- Handlers ---
  const handleDataChange = (index, field, val) => {
    const newData = [...data];
    newData[index][field] = field === 'value' ? Number(val) : val;
    setData(newData);
  };

  const addDataPoint = () => {
    setData([...data, { name: 'New', value: 0 }]);
  };

  const removeDataPoint = (index) => {
    if (data.length > 1) {
      setData(data.filter((_, i) => i !== index));
    }
  };

  const sortDataByDate = () => {
    const sorted = [...data].sort((a, b) => {
      // Try to parse as date
      const dateA = new Date(a.name);
      const dateB = new Date(b.name);
      
      // If invalid date, fall back to string comparison
      if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) {
        return a.name.localeCompare(b.name);
      }
      return dateA - dateB;
    });
    setData(sorted);
  };

  // --- File Import Handler ---
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const bstr = evt.target.result;
      let newData = [];

      try {
        if (file.name.endsWith('.csv')) {
          // Simple CSV Parser
          const lines = bstr.split(/\r\n|\n/);
          // Skip header if it exists, assume first line is header
          const startIdx = 1; 
          for (let i = startIdx; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            const parts = line.split(',');
            if (parts.length >= 2) {
              newData.push({
                name: parts[0].trim(),
                value: Number(parts[1].trim()) || 0
              });
            }
          }
        } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
          // Excel Parser using SheetJS
          if (window.XLSX) {
            const workbook = window.XLSX.read(bstr, { type: 'binary' });
            const wsname = workbook.SheetNames[0];
            const ws = workbook.Sheets[wsname];
            const jsonData = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
            
            // Assume Row 0 is header, start from Row 1
            for (let i = 1; i < jsonData.length; i++) {
              const row = jsonData[i];
              if (row && row.length >= 2) {
                newData.push({
                  name: String(row[0]), // First column as Label
                  value: Number(row[1]) || 0 // Second column as Value
                });
              }
            }
          } else {
            alert('Excel parser is still loading, please try again in a moment.');
            return;
          }
        }

        if (newData.length > 0) {
          setData(newData);
        } else {
          alert('Could not parse data. Ensure file has 2 columns: Label and Value.');
        }
      } catch (err) {
        console.error("Import error:", err);
        alert("Failed to parse file.");
      }
    };

    if (file.name.endsWith('.csv')) {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
    
    // Reset input
    e.target.value = null;
  };

  const triggerFileUpload = () => {
    fileInputRef.current?.click();
  };

  // --- Export Functions ---
  const exportToPNG = async () => {
    if (!window.html2canvas || !chartRef.current) return;
    setIsExporting(true);
    try {
      await new Promise(resolve => setTimeout(resolve, 100));
      // Use dark background for export
      const canvas = await window.html2canvas(chartRef.current, { backgroundColor: '#0f172a', scale: 2 });
      const image = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.href = image;
      link.download = `${chartTitle.replace(/\s+/g, '_')}_chart.png`;
      link.click();
    } catch (err) {
      console.error("Export failed", err);
    }
    setIsExporting(false);
  };

  const exportToPDF = async () => {
    if (!window.html2canvas || !window.jspdf || !chartRef.current) return;
    setIsExporting(true);
    try {
      await new Promise(resolve => setTimeout(resolve, 100));
      const canvas = await window.html2canvas(chartRef.current, { backgroundColor: '#0f172a', scale: 2 });
      const imgData = canvas.toDataURL('image/png');
      
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'px',
        format: [canvas.width / 2, canvas.height / 2]
      });
      
      pdf.addImage(imgData, 'PNG', 0, 0, canvas.width / 2, canvas.height / 2);
      pdf.save(`${chartTitle.replace(/\s+/g, '_')}_report.pdf`);
    } catch (err) {
      console.error("PDF Export failed", err);
    }
    setIsExporting(false);
  };

  // --- Render Helpers ---
  const renderChart = () => {
    // Dynamic palette for Pie Charts based on theme
    const PIE_COLORS = generateShades(primaryColor, data.length);

    // Increased margins to accommodate labels
    const CommonProps = {
      data: data,
      margin: { top: 20, right: 30, left: 30, bottom: 20 }
    };
    
    // Dark mode axis styles
    const axisStyle = { stroke: '#94a3b8' };
    const gridStyle = { stroke: '#334155', opacity: 0.5 };
    const tooltipStyle = { backgroundColor: '#1e293b', borderColor: '#334155', color: '#f1f5f9', borderRadius: '8px' };
    
    // Label Props
    const xLabelProps = { value: xAxisLabel, position: 'bottom', offset: 0, fill: '#64748b', fontSize: 12 };
    const yLabelProps = { value: yAxisLabel, angle: -90, position: 'insideLeft', offset: 10, fill: '#64748b', fontSize: 12, style: { textAnchor: 'middle' } };

    switch (chartType) {
      case 'line':
        return (
          <LineChart {...CommonProps}>
            <CartesianGrid strokeDasharray="3 3" {...gridStyle} />
            <XAxis dataKey="name" {...axisStyle} tick={{fill: '#94a3b8'}} label={xLabelProps} />
            <YAxis {...axisStyle} tick={{fill: '#94a3b8'}} label={yLabelProps} />
            <Tooltip contentStyle={tooltipStyle} />
            <Legend wrapperStyle={{ color: '#cbd5e1', paddingTop: '10px' }} />
            <Line type={curveType} dataKey="value" stroke={primaryColor} strokeWidth={3} activeDot={{ r: 8 }} dot={{ r: 4, fill: primaryColor, strokeWidth: 0 }} />
          </LineChart>
        );
      case 'bar':
        // Handle Vertical vs Horizontal Layout
        if (barLayout === 'vertical') {
          return (
            <BarChart {...CommonProps} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" {...gridStyle} />
              {/* Swap Axis Types for Vertical Layout */}
              <XAxis type="number" {...axisStyle} tick={{fill: '#94a3b8'}} label={{...xLabelProps, value: yAxisLabel}} />
              <YAxis dataKey="name" type="category" width={80} {...axisStyle} tick={{fill: '#94a3b8'}} label={{...yLabelProps, value: xAxisLabel, angle: -90, position: 'insideLeft'}} />
              <Tooltip cursor={{fill: '#334155', opacity: 0.4}} contentStyle={tooltipStyle} />
              <Legend wrapperStyle={{ color: '#cbd5e1', paddingTop: '10px' }} />
              <Bar dataKey="value" fill={primaryColor} radius={[0, 4, 4, 0]}>
                {data.map((entry, index) => (
                   <Cell key={`cell-${index}`} fill={index % 2 === 0 ? primaryColor : primaryColor + 'cc'} />
                ))}
              </Bar>
            </BarChart>
          );
        }
        // Default Horizontal Bar
        return (
          <BarChart {...CommonProps}>
            <CartesianGrid strokeDasharray="3 3" {...gridStyle} />
            <XAxis dataKey="name" {...axisStyle} tick={{fill: '#94a3b8'}} label={xLabelProps} />
            <YAxis {...axisStyle} tick={{fill: '#94a3b8'}} label={yLabelProps} />
            <Tooltip cursor={{fill: '#334155', opacity: 0.4}} contentStyle={tooltipStyle} />
            <Legend wrapperStyle={{ color: '#cbd5e1', paddingTop: '10px' }} />
            <Bar dataKey="value" fill={primaryColor} radius={[4, 4, 0, 0]}>
              {data.map((entry, index) => (
                <Cell key={`cell-${index}`} fill={index % 2 === 0 ? primaryColor : primaryColor + 'cc'} />
              ))}
            </Bar>
          </BarChart>
        );
      case 'area':
        return (
          <AreaChart {...CommonProps}>
            <CartesianGrid strokeDasharray="3 3" {...gridStyle} />
            <XAxis dataKey="name" {...axisStyle} tick={{fill: '#94a3b8'}} label={xLabelProps} />
            <YAxis {...axisStyle} tick={{fill: '#94a3b8'}} label={yLabelProps} />
            <Tooltip contentStyle={tooltipStyle} />
            <Legend wrapperStyle={{ color: '#cbd5e1', paddingTop: '10px' }} />
            <Area type={curveType} dataKey="value" stroke={primaryColor} fill={primaryColor} fillOpacity={0.3} />
          </AreaChart>
        );
      case 'pie':
        return (
          <PieChart>
            <Pie
              key={primaryColor} // Forces re-render on theme change
              data={data}
              cx="50%"
              cy="50%"
              labelLine={true} // Add lines connecting slices to labels
              label={({ cx, cy, midAngle, innerRadius, outerRadius, percent, index }) => {
                const RADIAN = Math.PI / 180;
                // Move labels outside the radius
                const radius = outerRadius * 1.25; 
                const x = cx + radius * Math.cos(-midAngle * RADIAN);
                const y = cy + radius * Math.sin(-midAngle * RADIAN);
                
                // Hide labels for very small slices
                if (percent < 0.05) return null;

                const name = data[index]?.name || '';

                return (
                  <text 
                    x={x} 
                    y={y} 
                    fill="#94a3b8" 
                    textAnchor={x > cx ? 'start' : 'end'} 
                    dominantBaseline="central"
                    className="text-[10px] font-bold label-fade-in"
                  >
                    {`${name}: ${(percent * 100).toFixed(0)}%`}
                  </text>
                );
              }}
              // Reduce outerRadius to give more room for external labels
              outerRadius={105} 
              fill="#8884d8"
              dataKey="value"
            >
              {data.map((entry, index) => (
                <Cell key={`cell-${index}`} fill={PIE_COLORS[index % PIE_COLORS.length]} />
              ))}
            </Pie>
            <Tooltip contentStyle={tooltipStyle} />
            <Legend wrapperStyle={{ color: '#cbd5e1' }} />
          </PieChart>
        );
      default:
        return null;
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans">
      {/* Header */}
      <header className="bg-slate-900 border-b border-slate-800 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            {/* New Unique Logo */}
            <div className="bg-slate-800 p-1.5 rounded-xl border border-slate-700 shadow-lg shadow-indigo-500/10 group hover:border-indigo-500/50 transition-colors">
              <svg width="32" height="32" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
                 <rect x="4" y="4" width="24" height="24" rx="6" className="fill-slate-900" />
                 {/* Stylized R made of bars */}
                 <path d="M10 24V10H15C18.3137 10 21 12.6863 21 16C21 17.6569 20.3284 19.1569 19.2426 20.2426" stroke="#818cf8" strokeWidth="3" strokeLinecap="round" />
                 <path d="M19 20L22 24" stroke="#c084fc" strokeWidth="3" strokeLinecap="round" />
                 <path d="M10 16H15" stroke="#818cf8" strokeWidth="3" />
                 <circle cx="26" cy="6" r="2" className="fill-indigo-400 animate-pulse" />
              </svg>
            </div>
            <h1 className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-indigo-400 to-purple-400 tracking-wide">
              ReadyCharts
            </h1>
          </div>
          <div className="flex items-center space-x-4">
            <button 
              onClick={exportToPNG}
              disabled={isExporting}
              className="flex items-center space-x-2 px-4 py-2 text-sm font-medium text-slate-300 bg-slate-800 border border-slate-700 rounded-md hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-colors"
            >
              <FileImage className="w-4 h-4" />
              <span>PNG</span>
            </button>
            <button 
              onClick={exportToPDF}
              disabled={isExporting}
              className="flex items-center space-x-2 px-4 py-2 text-sm font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-colors shadow-sm"
            >
              <FileText className="w-4 h-4" />
              <span>PDF</span>
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          
          {/* Sidebar: Controls */}
          <div className="lg:col-span-4 space-y-6">
            
            {/* Chart Config Card */}
            <div className="bg-slate-900 rounded-xl shadow-lg border border-slate-800 p-6">
              <div className="flex items-center space-x-2 mb-4 text-indigo-400">
                <Settings className="w-5 h-5" />
                <h2 className="font-semibold text-lg">Configuration</h2>
              </div>
              
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-slate-400 mb-1">Chart Title</label>
                  <input 
                    type="text" 
                    value={chartTitle}
                    onChange={(e) => setChartTitle(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-800 border border-slate-700 text-slate-100 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm"
                  />
                </div>

                {/* Axis Labels Config */}
                {(chartType !== 'pie') && (
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-sm font-medium text-slate-400 mb-1">X-Axis Label</label>
                      <input 
                        type="text" 
                        value={xAxisLabel}
                        onChange={(e) => setXAxisLabel(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-800 border border-slate-700 text-slate-100 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-slate-400 mb-1">Y-Axis Label</label>
                      <input 
                        type="text" 
                        value={yAxisLabel}
                        onChange={(e) => setYAxisLabel(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-800 border border-slate-700 text-slate-100 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm"
                      />
                    </div>
                  </div>
                )}

                <div>
                  <label className="block text-sm font-medium text-slate-400 mb-2">Chart Type</label>
                  <div className="grid grid-cols-2 gap-2">
                    {[
                      { id: 'line', icon: Activity, label: 'Line' },
                      { id: 'bar', icon: BarChart2, label: 'Bar' },
                      { id: 'area', icon: Layers, label: 'Area' },
                      { id: 'pie', icon: PieIcon, label: 'Pie' },
                    ].map((type) => (
                      <button
                        key={type.id}
                        onClick={() => setChartType(type.id)}
                        className={`flex items-center justify-center space-x-2 px-3 py-2 rounded-md text-sm font-medium transition-all ${
                          chartType === type.id 
                            ? 'bg-indigo-500/10 text-indigo-400 border border-indigo-500/50' 
                            : 'bg-slate-800 text-slate-400 border border-slate-700 hover:bg-slate-700'
                        }`}
                      >
                        <type.icon className="w-4 h-4" />
                        <span>{type.label}</span>
                      </button>
                    ))}
                  </div>
                </div>

                {/* Customizations specific to types */}
                
                {/* 1. Bar Chart: Layout Toggle */}
                {chartType === 'bar' && (
                  <div>
                     <label className="block text-sm font-medium text-slate-400 mb-2">Orientation</label>
                     <div className="flex bg-slate-800 p-1 rounded-md border border-slate-700">
                        <button
                          onClick={() => setBarLayout('horizontal')}
                          className={`flex-1 flex items-center justify-center space-x-2 py-1 text-xs font-medium rounded transition-colors ${barLayout === 'horizontal' ? 'bg-slate-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                        >
                          <AlignLeft className="w-3 h-3 rotate-90" />
                          <span>Vertical</span>
                        </button>
                        <button
                          onClick={() => setBarLayout('vertical')}
                          className={`flex-1 flex items-center justify-center space-x-2 py-1 text-xs font-medium rounded transition-colors ${barLayout === 'vertical' ? 'bg-slate-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                        >
                          <AlignLeft className="w-3 h-3" />
                          <span>Horizontal</span>
                        </button>
                     </div>
                  </div>
                )}

                {/* 2. Timeline Customization: Curve Type */}
                {(chartType === 'line' || chartType === 'area') && (
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Line Style</label>
                    <div className="flex bg-slate-800 p-1 rounded-md border border-slate-700">
                      {[
                        { id: 'monotone', label: 'Smooth' },
                        { id: 'linear', label: 'Straight' },
                        { id: 'step', label: 'Step' },
                      ].map((style) => (
                        <button
                          key={style.id}
                          onClick={() => setCurveType(style.id)}
                          className={`flex-1 py-1 text-xs font-medium rounded transition-colors ${
                            curveType === style.id 
                              ? 'bg-slate-600 text-white shadow' 
                              : 'text-slate-400 hover:text-slate-200'
                          }`}
                        >
                          {style.label}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                <div>
                  <label className="block text-sm font-medium text-slate-400 mb-2">Theme Color</label>
                  <div className="flex space-x-2">
                    {['#818cf8', '#f87171', '#34d399', '#fbbf24', '#f472b6', '#60a5fa'].map((color) => (
                      <button
                        key={color}
                        onClick={() => setPrimaryColor(color)}
                        className={`w-8 h-8 rounded-full border-2 transition-transform hover:scale-110 ${primaryColor === color ? 'border-white ring-2 ring-slate-700' : 'border-transparent'}`}
                        style={{ backgroundColor: color }}
                        aria-label={`Select color ${color}`}
                      />
                    ))}
                  </div>
                  {chartType === 'pie' && (
                    <p className="text-[10px] text-slate-500 mt-1">Pie chart shades will be generated from this color.</p>
                  )}
                </div>
              </div>
            </div>

            {/* Data Editor Card */}
            <div className="bg-slate-900 rounded-xl shadow-lg border border-slate-800 p-6">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center space-x-2 text-indigo-400">
                  <BarChart2 className="w-5 h-5" />
                  <h2 className="font-semibold text-lg">Data Points</h2>
                </div>
                <div className="flex space-x-2">
                   {/* Hidden File Input */}
                   <input 
                    type="file" 
                    ref={fileInputRef} 
                    onChange={handleFileUpload} 
                    className="hidden" 
                    accept=".csv, .xlsx, .xls"
                  />
                  <button 
                    onClick={sortDataByDate}
                    className="p-1.5 bg-slate-800 text-slate-400 border border-slate-700 rounded-md hover:bg-slate-700 hover:text-indigo-400 transition-colors"
                    title="Sort Chronologically"
                  >
                    <Calendar className="w-4 h-4" />
                  </button>
                  <button 
                    onClick={triggerFileUpload}
                    className="p-1.5 bg-slate-800 text-slate-400 border border-slate-700 rounded-md hover:bg-slate-700 hover:text-indigo-400 transition-colors"
                    title="Import CSV/XLSX"
                  >
                    <Upload className="w-4 h-4" />
                  </button>
                  <button 
                    onClick={addDataPoint}
                    className="p-1.5 bg-indigo-500/10 text-indigo-400 border border-indigo-500/50 rounded-md hover:bg-indigo-500/20 transition-colors"
                    title="Add Data Point"
                  >
                    <Plus className="w-4 h-4" />
                  </button>
                </div>
              </div>

              <div className="space-y-2 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
                {data.map((item, index) => (
                  <div key={index} className="flex space-x-2 items-center group">
                    <input
                      type="text"
                      value={item.name}
                      onChange={(e) => handleDataChange(index, 'name', e.target.value)}
                      className="w-1/3 px-2 py-1.5 text-sm bg-slate-800 border border-slate-700 text-slate-200 rounded focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 placeholder-slate-500"
                      placeholder="Label"
                    />
                    <input
                      type="number"
                      value={item.value}
                      onChange={(e) => handleDataChange(index, 'value', e.target.value)}
                      className="w-1/3 px-2 py-1.5 text-sm bg-slate-800 border border-slate-700 text-slate-200 rounded focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 placeholder-slate-500"
                      placeholder="Value"
                    />
                    <button
                      onClick={() => removeDataPoint(index)}
                      className="p-1.5 text-slate-500 hover:text-red-400 rounded opacity-0 group-hover:opacity-100 transition-opacity"
                      disabled={data.length <= 1}
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ))}
              </div>
              <p className="text-xs text-slate-500 mt-4 text-center">
                Import CSV/Excel or enter dates (YYYY-MM-DD) to sort.
              </p>
            </div>

          </div>

          {/* Main Content: Chart Preview */}
          <div className="lg:col-span-8">
            <div className="bg-slate-900 rounded-xl shadow-lg border border-slate-800 p-8 min-h-[600px] flex flex-col">
              <div ref={chartRef} className="flex-1 flex flex-col bg-slate-900 p-4 rounded-lg">
                <div className="mb-8 text-center">
                  <h2 className="text-2xl font-bold text-slate-100 tracking-tight">{chartTitle}</h2>
                </div>
                
                <div className="w-full h-[500px]">
                  <ResponsiveContainer width="100%" height="100%">
                    {renderChart()}
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          </div>

        </div>
      </main>
      
      <style jsx global>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: #1e293b;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #475569;
          border-radius: 3px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #64748b;
        }
        
        .label-fade-in {
          animation: fadeIn 0.5s ease-out forwards;
          opacity: 0;
        }
        
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(5px); }
          to { opacity: 1; transform: translateY(0); }
        }
      `}</style>
    </div>
  );
}