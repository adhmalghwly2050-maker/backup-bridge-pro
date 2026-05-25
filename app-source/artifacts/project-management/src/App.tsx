import { useState } from "react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell,
  ResponsiveContainer,
  LineChart,
  Line,
} from "recharts";
import { Printer, FileSpreadsheet, TrendingUp, CheckCircle, Clock, AlertTriangle, FolderOpen, DollarSign } from "lucide-react";

const queryClient = new QueryClient();

const projects = [
  { id: 1, name: "مشروع تطوير البنية التحتية", status: "مكتمل", progress: 100, budget: 500000, spent: 480000, startDate: "2024-01-01", endDate: "2024-06-30", manager: "أحمد محمد", department: "الهندسة" },
  { id: 2, name: "مشروع تحديث نظام المعلومات", status: "قيد التنفيذ", progress: 65, budget: 300000, spent: 195000, startDate: "2024-03-01", endDate: "2024-12-31", manager: "فاطمة علي", department: "تقنية المعلومات" },
  { id: 3, name: "مشروع توسعة المستودعات", status: "قيد التنفيذ", progress: 40, budget: 750000, spent: 300000, startDate: "2024-05-01", endDate: "2025-03-31", manager: "خالد عبدالله", department: "العمليات" },
  { id: 4, name: "مشروع تدريب الكوادر البشرية", status: "متأخر", progress: 30, budget: 150000, spent: 120000, startDate: "2024-02-01", endDate: "2024-08-31", manager: "نورة سالم", department: "الموارد البشرية" },
  { id: 5, name: "مشروع تطوير المنتجات الجديدة", status: "قيد التنفيذ", progress: 55, budget: 400000, spent: 220000, startDate: "2024-04-01", endDate: "2025-01-31", manager: "محمد إبراهيم", department: "البحث والتطوير" },
  { id: 6, name: "مشروع الاستدامة البيئية", status: "مكتمل", progress: 100, budget: 200000, spent: 185000, startDate: "2023-10-01", endDate: "2024-04-30", manager: "سارة يوسف", department: "البيئة" },
  { id: 7, name: "مشروع تطوير العلاقات التجارية", status: "قيد التنفيذ", progress: 70, budget: 250000, spent: 175000, startDate: "2024-01-15", endDate: "2024-11-30", manager: "عمر حسن", department: "التسويق" },
  { id: 8, name: "مشروع أتمتة المصانع", status: "متأخر", progress: 25, budget: 600000, spent: 350000, startDate: "2024-02-15", endDate: "2024-10-31", manager: "رانيا محمود", department: "الإنتاج" },
];

const statusColors: Record<string, string> = {
  "مكتمل": "#22c55e",
  "قيد التنفيذ": "#3b82f6",
  "متأخر": "#ef4444",
  "معلق": "#f59e0b",
};

const statusBg: Record<string, string> = {
  "مكتمل": "bg-green-100 text-green-800",
  "قيد التنفيذ": "bg-blue-100 text-blue-800",
  "متأخر": "bg-red-100 text-red-800",
  "معلق": "bg-yellow-100 text-yellow-800",
};

const pieData = [
  { name: "مكتمل", value: projects.filter(p => p.status === "مكتمل").length },
  { name: "قيد التنفيذ", value: projects.filter(p => p.status === "قيد التنفيذ").length },
  { name: "متأخر", value: projects.filter(p => p.status === "متأخر").length },
];

const PIE_COLORS = ["#22c55e", "#3b82f6", "#ef4444"];

const monthlyData = [
  { month: "يناير", مكتمل: 2, قيدالتنفيذ: 5, متأخر: 1 },
  { month: "فبراير", مكتمل: 3, قيدالتنفيذ: 4, متأخر: 2 },
  { month: "مارس", مكتمل: 1, قيدالتنفيذ: 6, متأخر: 1 },
  { month: "أبريل", مكتمل: 4, قيدالتنفيذ: 3, متأخر: 2 },
  { month: "مايو", مكتمل: 2, قيدالتنفيذ: 5, متأخر: 0 },
  { month: "يونيو", مكتمل: 3, قيدالتنفيذ: 4, متأخر: 1 },
];

const budgetData = projects.map(p => ({
  name: p.name.substring(0, 15) + "...",
  الميزانية: p.budget,
  المصروف: p.spent,
}));

function StatCard({ icon: Icon, label, value, color }: { icon: any; label: string; value: string | number; color: string }) {
  return (
    <div className={`bg-white rounded-xl shadow p-5 flex items-center gap-4 border-r-4 ${color}`}>
      <div className="p-3 rounded-full bg-gray-50">
        <Icon className="w-6 h-6 text-gray-600" />
      </div>
      <div>
        <p className="text-sm text-gray-500">{label}</p>
        <p className="text-2xl font-bold text-gray-800">{value}</p>
      </div>
    </div>
  );
}

function ProjectManagementApp() {
  const [activeTab, setActiveTab] = useState<"dashboard" | "projects">("dashboard");
  const [filterStatus, setFilterStatus] = useState("الكل");

  const totalBudget = projects.reduce((s, p) => s + p.budget, 0);
  const totalSpent = projects.reduce((s, p) => s + p.spent, 0);
  const completed = projects.filter(p => p.status === "مكتمل").length;
  const delayed = projects.filter(p => p.status === "متأخر").length;

  const filteredProjects = filterStatus === "الكل" ? projects : projects.filter(p => p.status === filterStatus);

  const handlePrint = () => {
    const printContent = `
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <title>تقرير المشاريع</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: white; color: #1f2937; direction: rtl; font-size: 13px; }
    .header { background: linear-gradient(to left, #1d4ed8, #1e3a8a); color: white; padding: 20px 24px; display: flex; justify-content: space-between; align-items: center; }
    .header h1 { font-size: 20px; font-weight: bold; }
    .header p { font-size: 12px; color: #bfdbfe; margin-top: 2px; }
    .header .date { font-size: 12px; color: #bfdbfe; text-align: left; }
    .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; padding: 16px 24px; }
    .stat-card { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 10px; padding: 14px 16px; display: flex; align-items: center; gap: 12px; }
    .stat-icon { width: 40px; height: 40px; border-radius: 50%; background: #e5e7eb; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0; }
    .stat-label { font-size: 11px; color: #6b7280; }
    .stat-value { font-size: 22px; font-weight: bold; color: #111827; }
    .budget-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; padding: 0 24px 16px; }
    .section-title { font-size: 14px; font-weight: 600; color: #374151; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 2px solid #e5e7eb; }
    .section { padding: 0 24px 20px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    thead tr { background: #1d4ed8; color: white; }
    th { padding: 9px 10px; text-align: right; font-weight: 600; }
    tbody tr:nth-child(even) { background: #f9fafb; }
    tbody tr:nth-child(odd) { background: white; }
    td { padding: 8px 10px; color: #374151; border-bottom: 1px solid #f3f4f6; }
    .badge { display: inline-block; padding: 2px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-blue { background: #dbeafe; color: #1e40af; }
    .badge-red { background: #fee2e2; color: #991b1b; }
    .progress-bar { background: #e5e7eb; border-radius: 10px; height: 8px; width: 80px; display: inline-block; vertical-align: middle; overflow: hidden; }
    .progress-fill { height: 8px; border-radius: 10px; display: block; }
    .footer { text-align: center; color: #9ca3af; font-size: 11px; padding: 16px; border-top: 1px solid #e5e7eb; margin-top: 10px; }
    @media print {
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      @page { margin: 10mm; size: A4 landscape; }
    }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>📁 نظام تحليل بيانات المشاريع</h1>
      <p>لوحة التحكم والمتابعة</p>
    </div>
    <div class="date">تاريخ التقرير: ${new Date().toLocaleDateString("ar-SA")}</div>
  </div>

  <div class="stats-grid">
    <div class="stat-card" style="border-right: 4px solid #3b82f6;">
      <div class="stat-icon">📁</div>
      <div><div class="stat-label">إجمالي المشاريع</div><div class="stat-value">${projects.length}</div></div>
    </div>
    <div class="stat-card" style="border-right: 4px solid #22c55e;">
      <div class="stat-icon">✅</div>
      <div><div class="stat-label">مشاريع مكتملة</div><div class="stat-value">${completed}</div></div>
    </div>
    <div class="stat-card" style="border-right: 4px solid #ef4444;">
      <div class="stat-icon">⚠️</div>
      <div><div class="stat-label">مشاريع متأخرة</div><div class="stat-value">${delayed}</div></div>
    </div>
    <div class="stat-card" style="border-right: 4px solid #f59e0b;">
      <div class="stat-icon">🕐</div>
      <div><div class="stat-label">قيد التنفيذ</div><div class="stat-value">${projects.filter(p => p.status === "قيد التنفيذ").length}</div></div>
    </div>
  </div>

  <div class="budget-grid">
    <div class="stat-card" style="border-right: 4px solid #a855f7;">
      <div class="stat-icon">💰</div>
      <div><div class="stat-label">إجمالي الميزانية</div><div class="stat-value">${(totalBudget / 1000000).toFixed(2)} م ر.س</div></div>
    </div>
    <div class="stat-card" style="border-right: 4px solid #f97316;">
      <div class="stat-icon">📈</div>
      <div><div class="stat-label">إجمالي المصروف</div><div class="stat-value">${(totalSpent / 1000000).toFixed(2)} م ر.س</div></div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">تفاصيل المشاريع</div>
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>اسم المشروع</th>
          <th>الحالة</th>
          <th>الإنجاز</th>
          <th>الميزانية (ر.س)</th>
          <th>المصروف (ر.س)</th>
          <th>تاريخ البدء</th>
          <th>تاريخ الانتهاء</th>
          <th>المدير</th>
          <th>القسم</th>
        </tr>
      </thead>
      <tbody>
        ${projects.map((p, i) => {
          const badgeClass = p.status === "مكتمل" ? "badge-green" : p.status === "متأخر" ? "badge-red" : "badge-blue";
          const fillColor = p.status === "مكتمل" ? "#22c55e" : p.status === "متأخر" ? "#ef4444" : "#3b82f6";
          return `<tr>
            <td>${p.id}</td>
            <td><strong>${p.name}</strong></td>
            <td><span class="badge ${badgeClass}">${p.status}</span></td>
            <td>
              <span class="progress-bar"><span class="progress-fill" style="width:${p.progress}%;background:${fillColor};"></span></span>
              <span style="font-size:11px;margin-right:6px;">${p.progress}%</span>
            </td>
            <td>${p.budget.toLocaleString()}</td>
            <td>${p.spent.toLocaleString()}</td>
            <td>${p.startDate}</td>
            <td>${p.endDate}</td>
            <td>${p.manager}</td>
            <td>${p.department}</td>
          </tr>`;
        }).join("")}
      </tbody>
    </table>
  </div>

  <div class="footer">
    تم إنشاء هذا التقرير من نظام تحليل بيانات المشاريع — ${new Date().toLocaleDateString("ar-SA")}
  </div>

  <script>
    window.onload = function() { window.print(); }
  </script>
</body>
</html>`;

    const printWindow = window.open("", "_blank", "width=1100,height=700");
    if (printWindow) {
      printWindow.document.write(printContent);
      printWindow.document.close();
    }
  };

  const handleExportExcel = () => {
    const wsData = [
      ["رقم المشروع", "اسم المشروع", "الحالة", "نسبة الإنجاز (%)", "الميزانية (ر.س)", "المصروف (ر.س)", "تاريخ البدء", "تاريخ الانتهاء", "مدير المشروع", "القسم"],
      ...projects.map(p => [
        p.id,
        p.name,
        p.status,
        p.progress,
        p.budget,
        p.spent,
        p.startDate,
        p.endDate,
        p.manager,
        p.department,
      ]),
    ];

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws["!cols"] = [
      { wch: 10 }, { wch: 35 }, { wch: 15 }, { wch: 18 },
      { wch: 18 }, { wch: 18 }, { wch: 15 }, { wch: 15 },
      { wch: 20 }, { wch: 20 },
    ];

    const summaryData = [
      ["ملخص المشاريع"],
      ["إجمالي المشاريع", projects.length],
      ["المشاريع المكتملة", completed],
      ["المشاريع قيد التنفيذ", projects.filter(p => p.status === "قيد التنفيذ").length],
      ["المشاريع المتأخرة", delayed],
      ["إجمالي الميزانية (ر.س)", totalBudget],
      ["إجمالي المصروف (ر.س)", totalSpent],
      ["نسبة الصرف", `${Math.round((totalSpent / totalBudget) * 100)}%`],
    ];

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary["!cols"] = [{ wch: 25 }, { wch: 20 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "بيانات المشاريع");
    XLSX.utils.book_append_sheet(wb, wsSummary, "الملخص");

    XLSX.writeFile(wb, "تقرير_المشاريع.xlsx");
  };

  return (
    <div className="min-h-screen bg-gray-100" dir="rtl">
      {/* Print styles */}
      <style>{`
        @media print {
          .no-print { display: none !important; }
          body { background: white; }
          .print-section { page-break-inside: avoid; }
        }
      `}</style>

      {/* Header */}
      <header className="bg-gradient-to-l from-blue-700 to-blue-900 text-white shadow-lg no-print">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <FolderOpen className="w-8 h-8" />
            <div>
              <h1 className="text-xl font-bold">نظام تحليل بيانات المشاريع</h1>
              <p className="text-blue-200 text-xs">لوحة التحكم والمتابعة</p>
            </div>
          </div>
          <div className="flex gap-2">
            <button
              onClick={handleExportExcel}
              className="flex items-center gap-2 bg-green-500 hover:bg-green-600 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors"
            >
              <FileSpreadsheet className="w-4 h-4" />
              تصدير Excel
            </button>
            <button
              onClick={handlePrint}
              className="flex items-center gap-2 bg-white text-blue-800 hover:bg-blue-50 px-4 py-2 rounded-lg text-sm font-medium transition-colors"
            >
              <Printer className="w-4 h-4" />
              طباعة
            </button>
          </div>
        </div>
      </header>

      {/* Tabs */}
      <div className="bg-white shadow-sm no-print">
        <div className="max-w-7xl mx-auto px-4">
          <div className="flex gap-1 pt-3">
            {[
              { key: "dashboard", label: "لوحة التحكم" },
              { key: "projects", label: "جدول المشاريع" },
            ].map(tab => (
              <button
                key={tab.key}
                onClick={() => setActiveTab(tab.key as any)}
                className={`px-6 py-2 text-sm font-medium rounded-t-lg border-b-2 transition-colors ${
                  activeTab === tab.key
                    ? "border-blue-600 text-blue-600 bg-blue-50"
                    : "border-transparent text-gray-500 hover:text-gray-700"
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      <main className="max-w-7xl mx-auto px-4 py-6">
        {activeTab === "dashboard" && (
          <div className="space-y-6">
            {/* Stats */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 print-section">
              <StatCard icon={FolderOpen} label="إجمالي المشاريع" value={projects.length} color="border-blue-500" />
              <StatCard icon={CheckCircle} label="مشاريع مكتملة" value={completed} color="border-green-500" />
              <StatCard icon={AlertTriangle} label="مشاريع متأخرة" value={delayed} color="border-red-500" />
              <StatCard icon={Clock} label="قيد التنفيذ" value={projects.filter(p => p.status === "قيد التنفيذ").length} color="border-yellow-500" />
            </div>

            <div className="grid grid-cols-2 gap-4 print-section">
              <StatCard icon={DollarSign} label="إجمالي الميزانية" value={`${(totalBudget / 1000000).toFixed(2)} م ر.س`} color="border-purple-500" />
              <StatCard icon={TrendingUp} label="إجمالي المصروف" value={`${(totalSpent / 1000000).toFixed(2)} م ر.س`} color="border-orange-500" />
            </div>

            {/* Charts */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Pie Chart */}
              <div className="bg-white rounded-xl shadow p-5 print-section">
                <h3 className="text-base font-semibold text-gray-700 mb-4">توزيع حالات المشاريع</h3>
                <ResponsiveContainer width="100%" height={260}>
                  <PieChart>
                    <Pie data={pieData} cx="50%" cy="50%" outerRadius={90} dataKey="value" label={({ name, value }) => `${name}: ${value}`}>
                      {pieData.map((_, i) => (
                        <Cell key={i} fill={PIE_COLORS[i]} />
                      ))}
                    </Pie>
                    <Tooltip />
                    <Legend />
                  </PieChart>
                </ResponsiveContainer>
              </div>

              {/* Monthly Bar */}
              <div className="bg-white rounded-xl shadow p-5 print-section">
                <h3 className="text-base font-semibold text-gray-700 mb-4">الأداء الشهري للمشاريع</h3>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={monthlyData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                    <YAxis tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="مكتمل" fill="#22c55e" />
                    <Bar dataKey="قيدالتنفيذ" name="قيد التنفيذ" fill="#3b82f6" />
                    <Bar dataKey="متأخر" fill="#ef4444" />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Budget Chart */}
              <div className="bg-white rounded-xl shadow p-5 md:col-span-2 print-section">
                <h3 className="text-base font-semibold text-gray-700 mb-4">مقارنة الميزانية والمصروف لكل مشروع</h3>
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={budgetData} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis type="number" tick={{ fontSize: 11 }} tickFormatter={v => `${(v / 1000).toFixed(0)}k`} />
                    <YAxis type="category" dataKey="name" tick={{ fontSize: 10 }} width={110} />
                    <Tooltip formatter={v => `${v.toLocaleString()} ر.س`} />
                    <Legend />
                    <Bar dataKey="الميزانية" fill="#93c5fd" />
                    <Bar dataKey="المصروف" fill="#3b82f6" />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Progress Line */}
              <div className="bg-white rounded-xl shadow p-5 md:col-span-2 print-section">
                <h3 className="text-base font-semibold text-gray-700 mb-4">نسبة إنجاز المشاريع</h3>
                <div className="space-y-3">
                  {projects.map(p => (
                    <div key={p.id} className="flex items-center gap-3">
                      <span className="text-xs text-gray-600 w-52 truncate shrink-0 text-right">{p.name}</span>
                      <div className="flex-1 bg-gray-200 rounded-full h-4 overflow-hidden">
                        <div
                          className="h-4 rounded-full transition-all"
                          style={{ width: `${p.progress}%`, backgroundColor: statusColors[p.status] }}
                        />
                      </div>
                      <span className="text-xs font-bold text-gray-700 w-10 text-left shrink-0">{p.progress}%</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "projects" && (
          <div className="space-y-4">
            {/* Filter */}
            <div className="bg-white rounded-xl shadow p-4 no-print flex items-center gap-3 flex-wrap">
              <span className="text-sm font-medium text-gray-600">تصفية حسب الحالة:</span>
              {["الكل", "مكتمل", "قيد التنفيذ", "متأخر"].map(s => (
                <button
                  key={s}
                  onClick={() => setFilterStatus(s)}
                  className={`px-4 py-1.5 rounded-full text-sm font-medium transition-colors ${
                    filterStatus === s ? "bg-blue-600 text-white" : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                  }`}
                >
                  {s}
                </button>
              ))}
            </div>

            {/* Table */}
            <div className="bg-white rounded-xl shadow overflow-hidden print-section">
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="bg-blue-700 text-white">
                    <tr>
                      <th className="px-4 py-3 text-right font-semibold">#</th>
                      <th className="px-4 py-3 text-right font-semibold">اسم المشروع</th>
                      <th className="px-4 py-3 text-right font-semibold">الحالة</th>
                      <th className="px-4 py-3 text-right font-semibold">الإنجاز</th>
                      <th className="px-4 py-3 text-right font-semibold">الميزانية</th>
                      <th className="px-4 py-3 text-right font-semibold">المصروف</th>
                      <th className="px-4 py-3 text-right font-semibold">تاريخ البدء</th>
                      <th className="px-4 py-3 text-right font-semibold">تاريخ الانتهاء</th>
                      <th className="px-4 py-3 text-right font-semibold">المدير</th>
                      <th className="px-4 py-3 text-right font-semibold">القسم</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredProjects.map((p, i) => (
                      <tr key={p.id} className={i % 2 === 0 ? "bg-white" : "bg-gray-50"}>
                        <td className="px-4 py-3 text-gray-500">{p.id}</td>
                        <td className="px-4 py-3 font-medium text-gray-800">{p.name}</td>
                        <td className="px-4 py-3">
                          <span className={`px-2 py-1 rounded-full text-xs font-semibold ${statusBg[p.status]}`}>
                            {p.status}
                          </span>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-2">
                            <div className="w-20 bg-gray-200 rounded-full h-2">
                              <div
                                className="h-2 rounded-full"
                                style={{ width: `${p.progress}%`, backgroundColor: statusColors[p.status] }}
                              />
                            </div>
                            <span className="text-xs text-gray-600">{p.progress}%</span>
                          </div>
                        </td>
                        <td className="px-4 py-3 text-gray-700">{p.budget.toLocaleString()} ر.س</td>
                        <td className="px-4 py-3 text-gray-700">{p.spent.toLocaleString()} ر.س</td>
                        <td className="px-4 py-3 text-gray-600">{p.startDate}</td>
                        <td className="px-4 py-3 text-gray-600">{p.endDate}</td>
                        <td className="px-4 py-3 text-gray-700">{p.manager}</td>
                        <td className="px-4 py-3 text-gray-600">{p.department}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Print footer */}
      <div className="hidden print:block text-center text-gray-400 text-xs mt-8 pb-4">
        تم إنشاء هذا التقرير من نظام تحليل بيانات المشاريع — {new Date().toLocaleDateString("ar-SA")}
      </div>
    </div>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <ProjectManagementApp />
    </QueryClientProvider>
  );
}

export default App;
