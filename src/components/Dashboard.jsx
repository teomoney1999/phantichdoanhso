import React, { useState, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line, AreaChart, Area, ComposedChart, PieChart, Pie, Cell } from 'recharts';
import { TrendingUp, DollarSign, Activity, PieChart as PieChartIcon, LayoutDashboard, FileText, Users, Settings, Menu } from 'lucide-react';
import * as XLSX from 'xlsx';
import logoImg from '../assets/logo.jpg';

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#14b8a6', '#f97316', '#6366f1', '#06b6d4', '#d946ef'];

export default function Dashboard() {
  const [summaryData, setSummaryData] = useState([]);
  const [costData, setCostData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  useEffect(() => {
    const fetchAllData = async () => {
      try {
        // 1. Fetch Summary Data
        const summaryRes = await fetch('/data/Tong_hop_Kinh_doanh_T1_T7_2026.xlsx');
        const summaryBuffer = await summaryRes.arrayBuffer();
        const summaryWorkbook = XLSX.read(summaryBuffer, { type: 'array' });
        const summarySheet = summaryWorkbook.Sheets[summaryWorkbook.SheetNames[0]];
        const summaryJsonRaw = XLSX.utils.sheet_to_json(summarySheet, { header: 1 });
        
        // Headers are at index 2 (row 3)
        const summaryHeaders = summaryJsonRaw[2];
        const summaryRows = summaryJsonRaw.slice(3).filter(row => {
          if (!row || row.length === 0 || !row[0]) return false;
          const firstCol = String(row[0]).toLowerCase();
          return !firstCol.includes('tổng');
        });
        
        const summaryParsed = summaryRows.map(row => {
          let obj = {};
          summaryHeaders.forEach((header, index) => {
            obj[header] = row[index] || 0;
          });
          return obj;
        });

        // Calculate Moving Average Trendline for Revenue (DOANH THU)
        const windowSize = 3; // 3-month moving average
        summaryParsed.forEach((item, i) => {
          let sum = 0;
          let count = 0;
          
          // Look back up to 'windowSize' months including the current month
          for (let j = Math.max(0, i - windowSize + 1); j <= i; j++) {
            sum += Number(summaryParsed[j]["DOANH THU"]) || 0;
            count++;
          }
          
          item.trendRevenue = count > 0 ? sum / count : 0;
        });

        // Calculate Profit Margin & Cost-to-Revenue Ratio
        summaryParsed.forEach((item) => {
          const revenue = Number(item["DOANH THU"]) || 0;
          const profit = Number(item["LỢI NHUẬN"]) || 0;
          const cost = Number(item["TỔNG CHI"]) || 0;
          
          item.profitMargin = revenue > 0 ? (profit / revenue) * 100 : 0;
          item.costToRevenueRatio = revenue > 0 ? (cost / revenue) * 100 : 0;
        });

        // 2. Fetch Cost Detail Data
        const costRes = await fetch('/data/Tong_hop_Chi_phi_Theo_Chu_de_v2.xlsx');
        const costBuffer = await costRes.arrayBuffer();
        const costWorkbook = XLSX.read(costBuffer, { type: 'array' });
        const costSheet = costWorkbook.Sheets[costWorkbook.SheetNames[0]];
        const costJsonRaw = XLSX.utils.sheet_to_json(costSheet, { header: 1 });

        const costHeaders = costJsonRaw[2];
        const costRows = costJsonRaw.slice(3).filter(row => {
          if (!row || row.length === 0 || !row[0]) return false;
          const firstCol = String(row[0]).toLowerCase();
          return !firstCol.includes('tổng');
        });
        
        const costParsed = costRows.map(row => {
          let obj = {};
          costHeaders.forEach((header, index) => {
            obj[header] = row[index] || 0;
          });
          return obj;
        });

        setSummaryData(summaryParsed);
        setCostData(costParsed);
        setLoading(false);
      } catch (error) {
        console.error("Error loading data:", error);
        setLoading(false);
      }
    };
    
    fetchAllData();
  }, []);

  if (loading) {
    return <div className="flex h-screen items-center justify-center text-xl text-gray-500">Đang tải dữ liệu từ thư mục /data...</div>;
  }

  if (summaryData.length === 0) {
    return <div className="flex h-screen items-center justify-center text-xl text-gray-500">Không có dữ liệu.</div>;
  }

  // Calculate totals
  const totalRevenue = summaryData.reduce((sum, item) => sum + (Number(item["DOANH THU"]) || 0), 0);
  const totalCost = summaryData.reduce((sum, item) => sum + (Number(item["TỔNG CHI"]) || 0), 0);
  const totalProfit = summaryData.reduce((sum, item) => sum + (Number(item["LỢI NHUẬN"]) || 0), 0);

  const formatCurrency = (value) => {
    if (value === undefined || value === null) return '0 ₫';
    return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(value);
  };

  const formatCompact = (value) => {
    if (value >= 1000000) {
      return (value / 1000000).toFixed(0) + 'Tr';
    }
    if (value >= 1000) {
      return (value / 1000).toFixed(0) + 'K';
    }
    return value;
  };

  // 1. Process aggregated cost data for Pie Chart
  const costTotalsByCategory = {};
  costData.forEach(monthData => {
    Object.keys(monthData).forEach(key => {
      if (key.toLowerCase() !== 'tháng' && !key.toLowerCase().includes('tổng')) {
        const val = Number(monthData[key]) || 0;
        costTotalsByCategory[key] = (costTotalsByCategory[key] || 0) + val;
      }
    });
  });
  
  const aggregatedCostData = Object.keys(costTotalsByCategory)
    .map(key => ({ name: key, value: costTotalsByCategory[key] }))
    .filter(item => item.value > 0)
    .sort((a, b) => b.value - a.value);

  // 2. Extract cost categories for the Stacked Bar Chart
  const costCategories = costData.length > 0 
    ? Object.keys(costData[0]).filter(k => k !== 'THÁNG' && !k.toLowerCase().includes('tổng'))
    : [];

  return (
    <div className="flex min-h-screen bg-slate-50">
      
      {/* Sidebar */}
      <aside className={`bg-white border-r border-slate-200 flex flex-col sticky top-0 h-screen transition-all duration-300 ${isSidebarOpen ? 'w-64' : 'w-20'} overflow-hidden`}>
        <div className={`border-b border-slate-100 flex justify-center items-center transition-all duration-300 ${isSidebarOpen ? 'p-6 h-52' : 'p-4 h-24'}`}>
          <img src={logoImg} alt="Logo" className={`object-cover rounded-full transition-all duration-300 ${isSidebarOpen ? 'w-40 h-40 min-w-[160px]' : 'w-10 h-10 min-w-[40px]'}`} />
        </div>
        <nav className="flex-1 p-4 space-y-2 overflow-y-auto overflow-x-hidden">
          <a href="#" className={`flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center px-0'} py-3 bg-blue-50 text-blue-700 rounded-xl font-medium transition-colors`} title="Dashboard">
            <LayoutDashboard className="w-5 h-5 min-w-[20px]" />
            {isSidebarOpen && <span className="whitespace-nowrap">Dashboard</span>}
          </a>
          <a href="#" className={`flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center px-0'} py-3 text-slate-600 hover:bg-slate-50 hover:text-slate-900 rounded-xl font-medium transition-colors`} title="Hiệu Quả">
            <TrendingUp className="w-5 h-5 min-w-[20px]" />
            {isSidebarOpen && <span className="whitespace-nowrap">Hiệu Quả</span>}
          </a>
          <a href="#" className={`flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center px-0'} py-3 text-slate-600 hover:bg-slate-50 hover:text-slate-900 rounded-xl font-medium transition-colors`} title="Báo Cáo">
            <FileText className="w-5 h-5 min-w-[20px]" />
            {isSidebarOpen && <span className="whitespace-nowrap">Báo Cáo</span>}
          </a>
          <a href="#" className={`flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center px-0'} py-3 text-slate-600 hover:bg-slate-50 hover:text-slate-900 rounded-xl font-medium transition-colors`} title="Nhân Sự">
            <Users className="w-5 h-5 min-w-[20px]" />
            {isSidebarOpen && <span className="whitespace-nowrap">Nhân Sự</span>}
          </a>
        </nav>
        <div className="p-4 border-t border-slate-100 flex flex-col space-y-2 overflow-x-hidden">
          <a href="#" className={`flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center px-0'} py-3 text-slate-600 hover:bg-slate-50 hover:text-slate-900 rounded-xl font-medium transition-colors`} title="Cài Đặt">
            <Settings className="w-5 h-5 min-w-[20px]" />
            {isSidebarOpen && <span className="whitespace-nowrap">Cài Đặt</span>}
          </a>
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className={`flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center px-0'} py-3 text-slate-600 hover:bg-slate-50 hover:text-slate-900 rounded-xl font-medium transition-colors w-full`}
            title="Toggle Sidebar"
          >
            <Menu className="w-5 h-5 min-w-[20px]" />
            {isSidebarOpen && <span className="whitespace-nowrap">Thu Gọn</span>}
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto">
        <div className="p-4 md:p-8 max-w-7xl mx-auto space-y-8">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-3xl font-bold text-slate-800 tracking-tight">Dashboard Tổng Hợp</h1>
              <p className="text-slate-500 mt-1">Dữ liệu phân tích từ T1-T7/2026</p>
            </div>
          </div>
          
          {/* KPI Cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 md:gap-6">
            <div className="bg-white p-4 md:p-6 rounded-2xl shadow-sm border border-slate-100 hover:shadow-md transition-shadow flex items-center md:items-start md:flex-col gap-4 md:gap-0">
              <div className="md:mb-4">
                <div className="p-3 bg-blue-50 text-blue-600 rounded-xl">
                  <DollarSign size={24} />
                </div>
              </div>
              <div className="flex-1 overflow-hidden">
                <p className="text-sm text-slate-500 font-medium mb-1">Tổng Doanh Thu</p>
                <p className="text-2xl xl:text-3xl font-bold text-slate-800 tracking-tight truncate" title={formatCurrency(totalRevenue)}>{formatCurrency(totalRevenue)}</p>
              </div>
            </div>
            
            <div className="bg-white p-4 md:p-6 rounded-2xl shadow-sm border border-slate-100 hover:shadow-md transition-shadow flex items-center md:items-start md:flex-col gap-4 md:gap-0">
              <div className="md:mb-4">
                <div className="p-3 bg-rose-50 text-rose-600 rounded-xl">
                  <Activity size={24} />
                </div>
              </div>
              <div className="flex-1 overflow-hidden">
                <p className="text-sm text-slate-500 font-medium mb-1">Tổng Chi Phí</p>
                <p className="text-2xl xl:text-3xl font-bold text-slate-800 tracking-tight truncate" title={formatCurrency(totalCost)}>{formatCurrency(totalCost)}</p>
              </div>
            </div>
            
            <div className="bg-white p-4 md:p-6 rounded-2xl shadow-sm border border-slate-100 hover:shadow-md transition-shadow flex items-center md:items-start md:flex-col gap-4 md:gap-0">
              <div className="md:mb-4">
                <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl">
                  <TrendingUp size={24} />
                </div>
              </div>
              <div className="flex-1 overflow-hidden">
                <p className="text-sm text-slate-500 font-medium mb-1">Tổng Lợi Nhuận</p>
                <p className="text-2xl xl:text-3xl font-bold text-slate-800 tracking-tight truncate" title={formatCurrency(totalProfit)}>{formatCurrency(totalProfit)}</p>
              </div>
            </div>
          </div>

          {/* Main Charts */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
              <h2 className="text-lg font-semibold mb-6 text-slate-800">Hiệu Quả Kinh Doanh Theo Tháng</h2>
              <div className="h-[350px]">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={summaryData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="THÁNG" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} dy={10} />
                    <YAxis yAxisId="left" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} tickFormatter={formatCompact} />
                    <Tooltip 
                      formatter={(value, name) => {
                        if (name === "Xu Hướng Doanh Thu") return [formatCurrency(value), name];
                        return [formatCurrency(value), name];
                      }}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                    />
                    <Legend wrapperStyle={{ paddingTop: '20px' }} />
                    <Bar yAxisId="left" dataKey="DOANH THU" name="Doanh Thu" fill="#3b82f6" radius={[4, 4, 0, 0]} maxBarSize={40} />
                    <Bar yAxisId="left" dataKey="TỔNG CHI" name="Chi Phí" fill="#fb7185" radius={[4, 4, 0, 0]} maxBarSize={40} />
                    <Line yAxisId="left" type="monotone" dataKey="trendRevenue" name="Xu Hướng Doanh Thu" stroke="#8b5cf6" strokeWidth={3} dot={{ r: 5, strokeWidth: 2, fill: "#fff", stroke: "#8b5cf6" }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
              <h2 className="text-lg font-semibold mb-6 text-slate-800">Lợi Nhuận Ròng & Biên Lợi Nhuận</h2>
              <div className="h-[350px]">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={summaryData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="THÁNG" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} dy={10} />
                    
                    {/* Left Axis for Profit (VND) */}
                    <YAxis yAxisId="left" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} tickFormatter={formatCompact} />
                    
                    {/* Right Axis for Margin (%) */}
                    <YAxis yAxisId="right" orientation="right" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} tickFormatter={(v) => `${(v).toFixed(0)}%`} />
                    
                    <Tooltip 
                      formatter={(value, name) => {
                        if (name === "Biên Lợi Nhuận") return [`${value.toFixed(1)}%`, name];
                        return [formatCurrency(value), name];
                      }}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                    />
                    <Legend wrapperStyle={{ paddingTop: '20px' }} />
                    
                    <Bar yAxisId="left" dataKey="LỢI NHUẬN" name="Lợi Nhuận Ròng" fill="#10b981" radius={[4, 4, 0, 0]} maxBarSize={50} />
                    <Line yAxisId="right" type="monotone" dataKey="profitMargin" name="Biên Lợi Nhuận" stroke="#f59e0b" strokeWidth={3} dot={{ r: 5, strokeWidth: 2, fill: "#fff", stroke: "#f59e0b" }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
          
          {/* Cost Charts Section */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
              <h2 className="text-lg font-semibold mb-6 text-slate-800">Cơ Cấu Chi Phí (Tổng Hợp)</h2>
              <div className="h-[350px]">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={aggregatedCostData}
                      cx="50%"
                      cy="45%"
                      innerRadius="40%"
                      outerRadius="75%"
                      paddingAngle={2}
                      dataKey="value"
                      label={({ percent }) => (percent * 100).toFixed(0) > 0 ? `${(percent * 100).toFixed(0)}%` : ''}
                    >
                      {aggregatedCostData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip 
                      formatter={(value, name) => [formatCurrency(value), name]}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                    />
                    <Legend 
                      layout="horizontal" 
                      verticalAlign="bottom" 
                      align="center"
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
              <h2 className="text-lg font-semibold mb-6 text-slate-800">Chi Phí Từng Tháng</h2>
              <div className="h-[350px]">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={summaryData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="THÁNG" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} dy={10} />
                    
                    {/* Left Axis for Cost (VND) */}
                    <YAxis yAxisId="left" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} tickFormatter={formatCompact} />
                    
                    {/* Right Axis for Ratio (%) */}
                    <YAxis yAxisId="right" orientation="right" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} tickFormatter={(v) => `${(v).toFixed(0)}%`} />
                    
                    <Tooltip 
                      formatter={(value, name) => {
                        if (name === "Tỷ Lệ Chi Phí / Doanh Thu") return [`${value.toFixed(1)}%`, name];
                        return [formatCurrency(value), name];
                      }}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                    />
                    <Legend wrapperStyle={{ paddingTop: '20px' }} />
                    <Bar yAxisId="left" dataKey="TỔNG CHI" name="Tổng Chi Phí" fill="#fb7185" maxBarSize={50} radius={[4, 4, 0, 0]} />
                    <Line yAxisId="right" type="monotone" dataKey="costToRevenueRatio" name="Tỷ Lệ Chi Phí / Doanh Thu" stroke="#f97316" strokeWidth={3} dot={{ r: 5, strokeWidth: 2, fill: "#fff", stroke: "#f97316" }} />
                    <Line yAxisId="right" type="monotone" dataKey="costToRevenueRatio" name="Tỷ Lệ Chi Phí / Doanh Thu" stroke="#6366f1" strokeWidth={3} dot={{ r: 5, strokeWidth: 2, fill: "#fff", stroke: "#6366f1" }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

        </div>
      </main>
    </div>
  );
}
