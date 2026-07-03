import React, { useState, useEffect, useRef } from "react";
import { 
  Upload, FileSpreadsheet, Trash2, AlertCircle, Sparkles, Download, 
  TrendingUp, DollarSign, ShoppingBag, Percent, BarChart3, PieChart as PieIcon, 
  MapPin, ShoppingCart, RefreshCw, Send, Search, ChevronLeft, ChevronRight, Check,
  Sliders, Info, HelpCircle, Star, ChevronDown, Award, TrendingDown
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, BarChart, Bar, Legend, ComposedChart, Line, LineChart
} from "recharts";

import { 
  SalesRecord, ColumnMapping, getMockSalesData, detectColumnMapping, 
  parseSheetData, downloadSampleExcel 
} from "./utils/salesData";

// Bright, high-end theme-friendly colors for dark-mode Bento Grid
const COLORS = ["#6366f1", "#10b981", "#f59e0b", "#ec4899", "#8b5cf6", "#14b8a6", "#3b82f6"];

interface SensitiveInfoResult {
  hasSensitive: boolean;
  reason: string;
}

function checkSensitiveInfo(json: any[]): SensitiveInfoResult {
  // 1. Password or direct secret key header check
  const passwordKeyPattern = /password|passcode|passwd|pwd|secret_key|secretkey/i;
  // 2. Card header patterns
  const cardKeyPattern = /credit\s*card|debit\s*card|card\s*number|cvv|cvc|card\s*pin|cc\s*num/i;
  // 3. Bank header patterns
  const bankKeyPattern = /bank\s*password|bank\s*passcode|bank\s*pin|routing\s*number|iban|swift\s*code|swift|swiftcode|bic|biccode/i;
  // 4. SSN header patterns
  const ssnKeyPattern = /ssn|social\s*security|national\s*id/i;

  const creditCardRegex = /\b(?:4[0-9]{12}(?:[0-9]{3})?|[25][1-7][0-9]{14}|6(?:011|5[0-9][0-9])[0-9]{12}|3[47][0-9]{13}|3(?:0[0-5]|[68][0-9])[0-9]{11}|(?:2131|1800|35\d{3})\d{11})\b/;
  const ssnRegex = /\b\d{3}-\d{2}-\d{4}\b/;

  // Luhn algorithm helper
  const isLuhnValid = (digitsStr: string): boolean => {
    let sum = 0;
    let shouldDouble = false;
    for (let i = digitsStr.length - 1; i >= 0; i--) {
      let digit = parseInt(digitsStr.charAt(i), 10);
      if (shouldDouble) {
        digit *= 2;
        if (digit > 9) digit -= 9;
      }
      sum += digit;
      shouldDouble = !shouldDouble;
    }
    return sum % 10 === 0;
  };

  for (const row of json) {
    if (!row || typeof row !== "object") continue;

    for (const key of Object.keys(row)) {
      const lowerKey = key.toLowerCase();

      // Check header sensitivity
      if (passwordKeyPattern.test(lowerKey)) {
        return {
          hasSensitive: true,
          reason: `Detected password/passcode field under column "${key}"`
        };
      }
      if (cardKeyPattern.test(lowerKey)) {
        return {
          hasSensitive: true,
          reason: `Detected potential credit/debit card details under column "${key}"`
        };
      }
      if (bankKeyPattern.test(lowerKey)) {
        return {
          hasSensitive: true,
          reason: `Detected sensitive bank account credentials/routing info under column "${key}"`
        };
      }
      if (ssnKeyPattern.test(lowerKey)) {
        return {
          hasSensitive: true,
          reason: `Detected Social Security or National Identity (SSN/ID) under column "${key}"`
        };
      }

      // Check row values
      const val = row[key];
      if (val !== null && val !== undefined) {
        const valStr = String(val).trim();
        if (valStr.length === 0) continue;

        // Skip barcode/SKU/product keys for value-based credit card/SSN checking to prevent false positives
        const isSkuOrIdColumn = /sku|barcode|ean|upc|product\s*id|item\s*id|order\s*id|id\b/i.test(lowerKey);

        if (!isSkuOrIdColumn) {
          // Check for credit card digits (ignoring spaces/hyphens)
          const digitsOnly = valStr.replace(/[^0-9]/g, "");
          if (digitsOnly.length >= 13 && digitsOnly.length <= 19) {
            // Check if it looks like a card number and passes Luhn check
            if (isLuhnValid(digitsOnly)) {
              return {
                hasSensitive: true,
                reason: `Detected a value matching a credit/debit card number pattern under column "${key}"`
              };
            }
          }

          // Check SSN format (###-##-####)
          if (ssnRegex.test(valStr)) {
            return {
              hasSensitive: true,
              reason: `Detected a value matching standard Social Security Number pattern under column "${key}"`
            };
          }
        }

        // Generic text-based password indicators
        const lowerVal = valStr.toLowerCase();
        if (
          lowerVal.includes("password:") ||
          lowerVal.includes("passcode:") ||
          lowerVal.includes("bank pin:") ||
          lowerVal.includes("card password:") ||
          lowerVal.includes("account password:")
        ) {
          return {
            hasSensitive: true,
            reason: `Detected credential or password strings in the cell content`
          };
        }
      }
    }
  }

  return { hasSensitive: false, reason: "" };
}

export default function App() {
  // State variables
  const [salesRecords, setSalesRecords] = useState<SalesRecord[]>([]);
  const [loadedFiles, setLoadedFiles] = useState<{ name: string; size: string; rows: number; rawJson: any[] }[]>([]);
  const [isDemoMode, setIsDemoMode] = useState<boolean>(false);
  const [sensitiveInfoWarning, setSensitiveInfoWarning] = useState<{ fileName: string; reason: string } | null>(null);
  const [showSensitiveModal, setShowSensitiveModal] = useState<boolean>(false);
  
  // Header detection & column mapping states
  const [availableHeaders, setAvailableHeaders] = useState<string[]>([]);
  const [columnMapping, setColumnMapping] = useState<ColumnMapping>({
    date: "",
    product: "",
    category: "",
    quantity: "",
    sales: "",
    profit: "",
    region: "",
    discountRate: "",
    returnAmt: "",
    week: "",
    city: "",
    store: "",
    storeFormat: ""
  });
  const [showMappingConfig, setShowMappingConfig] = useState<boolean>(false);

  // Committed Filter States
  const [selectedWeek, setSelectedWeek] = useState<string>("All");
  const [selectedRegion, setSelectedRegion] = useState<string>("All");
  const [selectedCity, setSelectedCity] = useState<string>("All");
  const [selectedStore, setSelectedStore] = useState<string>("All");
  const [selectedStoreFormat, setSelectedStoreFormat] = useState<string>("All");
  const [selectedCategory, setSelectedCategory] = useState<string>("All");

  // Temporary (Draft) States for the 'Apply' button filters
  const [tempRegion, setTempRegion] = useState<string>("All");
  const [tempCity, setTempCity] = useState<string>("All");
  const [tempStore, setTempStore] = useState<string>("All");
  const [tempStoreFormat, setTempStoreFormat] = useState<string>("All");
  const [tempCategory, setTempCategory] = useState<string>("All");

  const [searchQuery, setSearchQuery] = useState<string>("");
  const [currentPage, setCurrentPage] = useState<number>(1);
  const itemsPerPage = 8;
  const [activeKpiFilter, setActiveKpiFilter] = useState<string>("all");
  const [salesTarget, setSalesTarget] = useState<number>(180000);
  const [storeTopN, setStoreTopN] = useState<number | "All">(10);
  const [showStoreDropdown, setShowStoreDropdown] = useState<boolean>(false);
  const [showBusinessSummary, setShowBusinessSummary] = useState<boolean>(false);

  // AI Analyst States
  const [aiAnalysis, setAiAnalysis] = useState<string>("");
  const [isAnalyzing, setIsAnalyzing] = useState<boolean>(false);
  const [customPrompt, setCustomPrompt] = useState<string>("");
  const [analysisError, setAnalysisError] = useState<string>("");
  const [loadingStep, setLoadingStep] = useState<string>("");

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Load Initial Mock Data
  useEffect(() => {
    // Keep initially empty per privacy guidelines. User can trigger sandbox/demo mode manually or upload files.
    setSalesRecords([]);
    setIsDemoMode(false);
  }, []);

  // Update sales target automatically when dataset changes
  useEffect(() => {
    if (salesRecords.length > 0) {
      const gross = salesRecords.reduce((sum, r) => sum + r.sales, 0);
      const returns = salesRecords.reduce((sum, r) => sum + (r.returnAmt || 0), 0);
      const net = gross - returns;
      const baseline = net > 0 ? net : 100000;
      const calculatedTarget = Math.max(10000, Math.round((baseline * 1.15) / 10000) * 10000);
      setSalesTarget(calculatedTarget);
    }
  }, [salesRecords]);

  // Handle Drag Over
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  // Handle Drop
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files);
    }
  };

  // Process selected files
  const handleFileSelected = (files: FileList) => {
    Array.from(files).forEach((file) => {
      if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
        alert("Please upload valid Excel files (.xlsx or .xls)");
        return;
      }

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          if (!data) return;
          const workbook = XLSX.read(data, { type: "binary" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const json: any[] = XLSX.utils.sheet_to_json(worksheet);

          if (json.length === 0) {
            alert(`The sheet in file ${file.name} is empty.`);
            return;
          }

          // Check for sensitive information
          const sensitivity = checkSensitiveInfo(json);
          if (sensitivity.hasSensitive) {
            setSensitiveInfoWarning({
              fileName: file.name,
              reason: sensitivity.reason
            });
            setShowSensitiveModal(true);
            return;
          }

          // Gather unique headers across files
          const fileHeaders = Object.keys(json[0]);
          setAvailableHeaders((prev) => Array.from(new Set([...prev, ...fileHeaders])));

          // Keep track of loaded files in holding state
          const fileSizeStr = (file.size / 1024).toFixed(1) + " KB";
          setLoadedFiles((prev) => [
            ...prev,
            { name: file.name, size: fileSizeStr, rows: json.length, rawJson: json }
          ]);
          
          // Auto-detect mappings on first file
          if (loadedFiles.length === 0) {
            const detected = detectColumnMapping(fileHeaders);
            setColumnMapping(detected);
            setShowMappingConfig(true);
          }
        } catch (error) {
          console.error("Error reading Excel file:", error);
          alert(`Error reading file ${file.name}. Ensure it is not corrupted.`);
        }
      };
      reader.readAsBinaryString(file);
    });

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  // Submit and compile loaded files into active dashboard state
  const handleSubmitAndAnalyze = () => {
    if (loadedFiles.length === 0) {
      alert("Please upload at least one Excel file first.");
      return;
    }

    let compiledRecords: SalesRecord[] = [];
    loadedFiles.forEach((file, fileIdx) => {
      const records = parseSheetData(file.rawJson, columnMapping, fileIdx + 1);
      compiledRecords = [...compiledRecords, ...records];
    });

    // Sort compiled transactions by date
    compiledRecords.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

    setSalesRecords(compiledRecords);
    setIsDemoMode(false);
    setCurrentPage(1);
    setShowMappingConfig(false);
    setAiAnalysis(""); // Clear stale analysis

    // Reset committed and draft filters
    setSelectedWeek("All");
    setSelectedRegion("All");
    setSelectedCity("All");
    setSelectedStore("All");
    setSelectedStoreFormat("All");
    setSelectedCategory("All");

    setTempRegion("All");
    setTempCity("All");
    setTempStore("All");
    setTempStoreFormat("All");
    setTempCategory("All");

    // Automatically trigger AI analysis on new data
    triggerAIAnalysis(compiledRecords);
  };

  // Trigger Gemini AI insights via backend route
  const triggerAIAnalysis = async (dataToAnalyze = salesRecords, customQuery = "") => {
    if (dataToAnalyze.length === 0) return;
    
    setIsAnalyzing(true);
    setAnalysisError("");
    
    // Set up engaging steps for loading screen
    const steps = [
      "Compressing spreadsheet matrix...",
      "Analyzing margins and departmental profitability...",
      "Generating strategic sales insights...",
      "Formulating retail solutions with Gemini..."
    ];
    
    let stepIdx = 0;
    setLoadingStep(steps[0]);
    const stepInterval = setInterval(() => {
      stepIdx = (stepIdx + 1) % steps.length;
      setLoadingStep(steps[stepIdx]);
    }, 1800);

    try {
      // Create condensed summary of data to avoid sending massive payloads
      const totalSales = dataToAnalyze.reduce((acc, r) => acc + r.sales, 0);
      const totalProfit = dataToAnalyze.reduce((acc, r) => acc + r.profit, 0);
      
      // Categorical breakdown
      const categoryMap: Record<string, { sales: number; profit: number; qty: number }> = {};
      // Regional breakdown
      const regionMap: Record<string, { sales: number; profit: number }> = {};
      // Chronological trends
      const dailyMap: Record<string, number> = {};

      dataToAnalyze.forEach((r) => {
        // Category
        if (!categoryMap[r.category]) categoryMap[r.category] = { sales: 0, profit: 0, qty: 0 };
        categoryMap[r.category].sales += r.sales;
        categoryMap[r.category].profit += r.profit;
        categoryMap[r.category].qty += r.quantity;

        // Region
        if (!regionMap[r.region]) regionMap[r.region] = { sales: 0, profit: 0 };
        regionMap[r.region].sales += r.sales;
        regionMap[r.region].profit += r.profit;

        // Date grouping
        const monthKey = r.date.substring(0, 7); // YYYY-MM
        dailyMap[monthKey] = (dailyMap[monthKey] || 0) + r.sales;
      });

      const categoriesSummary = Object.entries(categoryMap).map(([cat, stats]) => ({
        category: cat,
        totalSales: Math.round(stats.sales * 100) / 100,
        totalProfit: Math.round(stats.profit * 100) / 100,
        marginPercent: stats.sales > 0 ? Math.round((stats.profit / stats.sales) * 1000) / 10 : 0,
        unitsSold: stats.qty
      }));

      const regionsSummary = Object.entries(regionMap).map(([reg, stats]) => ({
        region: reg,
        totalSales: Math.round(stats.sales * 100) / 100,
        totalProfit: Math.round(stats.profit * 100) / 100,
        marginPercent: stats.sales > 0 ? Math.round((stats.profit / stats.sales) * 1000) / 10 : 0
      }));

      const monthlyTrends = Object.entries(dailyMap)
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([month, sales]) => ({ month, sales: Math.round(sales * 100) / 100 }));

      const topProducts = [...dataToAnalyze]
        .reduce((acc, curr) => {
          const existing = acc.find(p => p.name === curr.product);
          if (existing) {
            existing.sales += curr.sales;
            existing.profit += curr.profit;
            existing.qty += curr.quantity;
          } else {
            acc.push({ name: curr.product, category: curr.category, sales: curr.sales, profit: curr.profit, qty: curr.quantity });
          }
          return acc;
        }, [] as { name: string; category: string; sales: number; profit: number; qty: number }[])
        .sort((a, b) => b.sales - a.sales)
        .slice(0, 5);

      const summaryPayload = {
        metrics: {
          totalRevenue: Math.round(totalSales * 100) / 100,
          totalProfit: Math.round(totalProfit * 100) / 100,
          overallMarginPercent: totalSales > 0 ? Math.round((totalProfit / totalSales) * 1000) / 10 : 0,
          totalTransactions: dataToAnalyze.length,
          totalUnitsSold: dataToAnalyze.reduce((acc, r) => acc + r.quantity, 0),
          averageOrderValue: dataToAnalyze.length > 0 ? Math.round((totalSales / dataToAnalyze.length) * 100) / 100 : 0
        },
        categories: categoriesSummary,
        regions: regionsSummary,
        trends: monthlyTrends,
        topPerformingProducts: topProducts
      };

      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          dataSummary: summaryPayload,
          customPrompt: customQuery || undefined
        })
      });

      const resJson = await response.json();
      if (!response.ok) {
        throw new Error(resJson.error || "Server error during AI analysis.");
      }

      setAiAnalysis(resJson.analysis);
      if (customQuery) {
        setCustomPrompt("");
      }
    } catch (error: any) {
      console.error(error);
      setAnalysisError(error.message || "Could not complete AI analysis. Please verify your GEMINI_API_KEY.");
    } finally {
      clearInterval(stepInterval);
      setIsAnalyzing(false);
    }
  };

  // Remove file from loaded list
  const handleRemoveFile = (index: number) => {
    const updated = [...loadedFiles];
    updated.splice(index, 1);
    setLoadedFiles(updated);

    if (updated.length === 0) {
      setAvailableHeaders([]);
      setShowMappingConfig(false);
    }
  };

  // Toggle demo mode
  const handleLoadDemoData = () => {
    setSalesRecords(getMockSalesData());
    setLoadedFiles([]);
    setIsDemoMode(true);
    setAiAnalysis("");
    setShowMappingConfig(false);
    setCurrentPage(1);

    // Reset committed and draft filters
    setSelectedWeek("All");
    setSelectedRegion("All");
    setSelectedCity("All");
    setSelectedStore("All");
    setSelectedStoreFormat("All");
    setSelectedCategory("All");

    setTempRegion("All");
    setTempCity("All");
    setTempStore("All");
    setTempStoreFormat("All");
    setTempCategory("All");
  };

  // Filter list calculation
  const parseDDMMYYYY = (s: string): number => {
    const parts = s.split("/");
    if (parts.length !== 3) return 0;
    const [dd, mm, yyyy] = parts.map(Number);
    return new Date(yyyy, mm - 1, dd).getTime();
  };

  const uniqueWeeks = [
    "All",
    ...Array.from(new Set(salesRecords.map((r) => r.week).filter(Boolean)))
      .map(String)
      .sort((a, b) => parseDDMMYYYY(b) - parseDDMMYYYY(a))
  ];
  const uniqueCategories = ["All", ...Array.from(new Set(salesRecords.map((r) => r.category).filter(Boolean))).map(String).sort()];
  const uniqueRegions = ["All", ...Array.from(new Set(salesRecords.map((r) => r.region).filter(Boolean))).map(String).sort()];
  const uniqueCities = ["All", ...Array.from(new Set(salesRecords.map((r) => r.city).filter(Boolean))).map(String).sort()];
  const uniqueStores = ["All", ...Array.from(new Set(salesRecords.map((r) => r.store).filter(Boolean))).map(String).sort()];
  const uniqueStoreFormats = ["All", ...Array.from(new Set(salesRecords.map((r) => r.storeFormat).filter(Boolean))).map(String).sort()];

  // Base filtered list before applying the interactive KPI-specific filters
  const baseFilteredRecords = salesRecords.filter((record) => {
    const matchesWeek = selectedWeek === "All" || record.week === selectedWeek;
    const matchesCategory = selectedCategory === "All" || record.category === selectedCategory;
    const matchesRegion = selectedRegion === "All" || record.region === selectedRegion;
    const matchesCity = selectedCity === "All" || record.city === selectedCity;
    const matchesStore = selectedStore === "All" || record.store === selectedStore;
    const matchesStoreFormat = selectedStoreFormat === "All" || record.storeFormat === selectedStoreFormat;
    
    const matchesSearch = 
      record.product.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.category.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.region.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (record.city && record.city.toLowerCase().includes(searchQuery.toLowerCase())) ||
      (record.store && record.store.toLowerCase().includes(searchQuery.toLowerCase())) ||
      (record.storeFormat && record.storeFormat.toLowerCase().includes(searchQuery.toLowerCase())) ||
      record.id.toLowerCase().includes(searchQuery.toLowerCase());

    return matchesWeek && matchesCategory && matchesRegion && matchesCity && matchesStore && matchesStoreFormat && matchesSearch;
  });

  // Calculate iterated KPI metrics based on baseFilteredRecords
  const totalGrossSales = baseFilteredRecords.reduce((sum, r) => sum + r.sales, 0);
  const totalReturnsVal = baseFilteredRecords.reduce((sum, r) => sum + (r.returnAmt || 0), 0);
  const netSalesVal = Math.max(0, totalGrossSales - totalReturnsVal);
  const targetAchievedVal = salesTarget > 0 ? (netSalesVal / salesTarget) * 100 : 0;
  const avgOrderVal = baseFilteredRecords.length > 0 ? netSalesVal / baseFilteredRecords.length : 0;
  const returnRateVal = netSalesVal > 0 ? (totalReturnsVal / netSalesVal) * 100 : 0;
  
  const totalDiscountsVal = baseFilteredRecords.reduce((sum, r) => sum + (r.sales * (r.discountRate || 0)), 0);
  const avgDiscountRateVal = totalGrossSales > 0 ? (totalDiscountsVal / totalGrossSales) * 100 : 0;

  // Map backward-compatible values for standard display components
  const totalSalesVal = netSalesVal;
  const totalProfitVal = baseFilteredRecords.reduce((sum, r) => sum + r.profit, 0);
  const totalQtyVal = baseFilteredRecords.reduce((sum, r) => sum + r.quantity, 0);
  const overallMarginVal = totalSalesVal > 0 ? (totalProfitVal / totalSalesVal) * 100 : 0;

  // Apply interactive KPI filters to produce the final filteredRecords
  const filteredRecords = baseFilteredRecords.filter((record) => {
    if (activeKpiFilter === "returns") {
      return record.returnAmt > 0;
    }
    if (activeKpiFilter === "discounts") {
      return record.discountRate > 0;
    }
    if (activeKpiFilter === "high-value") {
      return (record.sales - record.returnAmt) > avgOrderVal;
    }
    return true; // "all"
  });

  // Pagination calculations
  const totalPages = Math.ceil(filteredRecords.length / itemsPerPage) || 1;
  const paginatedRecords = filteredRecords.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  // Grouped data for Charts
  // 1. Monthly chronologically grouped
  const getMonthlyTrendData = () => {
    const monthlyMap: Record<string, { month: string; Sales: number; Profit: number }> = {};
    filteredRecords.forEach((r) => {
      const monthKey = r.date.substring(0, 7); // "YYYY-MM"
      if (!monthlyMap[monthKey]) {
        monthlyMap[monthKey] = { month: monthKey, Sales: 0, Profit: 0 };
      }
      monthlyMap[monthKey].Sales += r.sales;
      monthlyMap[monthKey].Profit += r.profit;
    });

    return Object.values(monthlyMap)
      .sort((a, b) => a.month.localeCompare(b.month))
      .map(item => ({
        ...item,
        Sales: Math.round(item.Sales),
        Profit: Math.round(item.Profit)
      }));
  };

  // 2. Category Pie Chart Data
  const getCategoryChartData = () => {
    const catMap: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      catMap[r.category] = (catMap[r.category] || 0) + r.sales;
    });

    return Object.entries(catMap).map(([name, value]) => ({
      name,
      value: Math.round(value)
    })).sort((a, b) => b.value - a.value);
  };

  // 3. Top 5 Products
  const getTopProductsData = () => {
    const prodMap: Record<string, { name: string; Sales: number; Profit: number }> = {};
    filteredRecords.forEach((r) => {
      if (!prodMap[r.product]) {
        prodMap[r.product] = { name: r.product, Sales: 0, Profit: 0 };
      }
      prodMap[r.product].Sales += r.sales;
      prodMap[r.product].Profit += r.profit;
    });

    return Object.values(prodMap)
      .sort((a, b) => b.Sales - a.Sales)
      .slice(0, 5)
      .map(item => ({
        ...item,
        Sales: Math.round(item.Sales),
        Profit: Math.round(item.Profit)
      }));
  };

  // 4. Sales by Region (Bar chart, descending order)
  const getRegionSalesData = () => {
    const regionMap: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      regionMap[r.region] = (regionMap[r.region] || 0) + r.sales;
    });
    return Object.entries(regionMap)
      .map(([name, value]) => ({ name, Sales: Math.round(value) }))
      .sort((a, b) => b.Sales - a.Sales);
  };

  // 5. Pareto Chart Data (Category Performance: bar for sales, line for cumulative percentage)
  const getCategoryParetoData = () => {
    const catMap: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      catMap[r.category] = (catMap[r.category] || 0) + r.sales;
    });
    const sortedCats = Object.entries(catMap)
      .map(([category, Sales]) => ({ category, Sales: Math.round(Sales) }))
      .sort((a, b) => b.Sales - a.Sales);

    const totalSales = sortedCats.reduce((sum, item) => sum + item.Sales, 0);
    let runningSum = 0;

    return sortedCats.map((item) => {
      runningSum += item.Sales;
      const cumulativePercentage = totalSales > 0 ? Math.round((runningSum / totalSales) * 100) : 0;
      return {
        ...item,
        cumulativePercentage
      };
    });
  };

  // 6. Weekly Footfalls Trend (Line Chart)
  const getWeeklyFootfallsData = () => {
    const weeklyMap: Record<string, { week: string; count: number; quantity: number }> = {};
    filteredRecords.forEach((r) => {
      const wk = r.week || "N/A";
      if (!weeklyMap[wk]) {
        weeklyMap[wk] = { week: wk, count: 0, quantity: 0 };
      }
      weeklyMap[wk].count += 1;
      weeklyMap[wk].quantity += r.quantity;
    });

    return Object.values(weeklyMap)
      .sort((a, b) => parseDDMMYYYY(a.week) - parseDDMMYYYY(b.week))
      .map((item) => {
        // Deterministic multiplier for footfalls based on units sold and transactions count
        const multiplier = 4.8;
        const footfalls = Math.round(item.quantity * multiplier + item.count * 12 + 15);
        return {
          week: item.week,
          Footfalls: footfalls,
          Transactions: item.count
        };
      });
  };

  // 7. Store Performance (Table displaying store, footfalls, units_sold, net_sales)
  const getStorePerformanceData = () => {
    const storeMap: Record<string, { store: string; units_sold: number; net_sales: number; returns: number; transactions: number }> = {};
    filteredRecords.forEach((r) => {
      const st = r.store || "HQ Store";
      if (!storeMap[st]) {
        storeMap[st] = { store: st, units_sold: 0, net_sales: 0, returns: 0, transactions: 0 };
      }
      storeMap[st].units_sold += r.quantity;
      storeMap[st].net_sales += r.sales;
      storeMap[st].returns += (r.returnAmt || 0);
      storeMap[st].transactions += 1;
    });

    const sortedStores = Object.values(storeMap)
      .map((item) => {
        const footfalls = Math.round(item.units_sold * 4.8 + item.transactions * 12 + 15);
        const net_sales = Math.max(0, item.net_sales - item.returns);
        return {
          store: item.store,
          units_sold: item.units_sold,
          net_sales: Math.round(net_sales * 100) / 100,
          footfalls
         };
      })
      .sort((a, b) => b.net_sales - a.net_sales);

    if (storeTopN === "All") {
      return sortedStores;
    }
    return sortedStores.slice(0, storeTopN);
  };

  // Business Summary Helpers
  const getRegionsPerformance = () => {
    const regionSales: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      const net = r.sales - (r.returnAmt || 0);
      regionSales[r.region] = (regionSales[r.region] || 0) + net;
    });
    const sorted = Object.entries(regionSales).sort((a, b) => b[1] - a[1]);
    if (sorted.length === 0) {
      return { 
        best: { name: "N/A", value: 0 }, 
        worst: { name: "N/A", value: 0 } 
      };
    }
    return {
      best: { name: sorted[0][0], value: Math.round(sorted[0][1]) },
      worst: { name: sorted[sorted.length - 1][0], value: Math.round(sorted[sorted.length - 1][1]) }
    };
  };

  const getHighReturnsCategories = () => {
    const catReturns: Record<string, { returns: number; sales: number }> = {};
    filteredRecords.forEach((r) => {
      if (!catReturns[r.category]) {
        catReturns[r.category] = { returns: 0, sales: 0 };
      }
      catReturns[r.category].returns += (r.returnAmt || 0);
      catReturns[r.category].sales += r.sales;
    });
    return Object.entries(catReturns)
      .map(([category, data]) => {
        const rate = data.sales > 0 ? (data.returns / data.sales) * 100 : 0;
        return {
          category,
          returns: Math.round(data.returns),
          rate: Math.round(rate * 10) / 10
        };
      })
      .sort((a, b) => b.returns - a.returns)
      .slice(0, 3); // Top 3
  };

  const getStoresMissingTargets = () => {
    const storeMap: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      const st = r.store || "HQ Store";
      const net = r.sales - (r.returnAmt || 0);
      storeMap[st] = (storeMap[st] || 0) + net;
    });
    const stores = Object.keys(storeMap);
    if (stores.length === 0) return [];

    const activeStoresCount = uniqueStores.filter((s) => s !== "All").length || 1;
    const targetPerStore = Math.max(5000, Math.round(salesTarget / activeStoresCount));
    
    return Object.entries(storeMap)
      .map(([store, sales]) => ({
        store,
        sales: Math.round(sales),
        target: targetPerStore,
        percentage: Math.round((sales / targetPerStore) * 100),
        missingAmount: Math.max(0, Math.round(targetPerStore - sales))
      }))
      .filter((item) => item.sales < item.target)
      .sort((a, b) => a.sales - b.sales); // lowest first
  };

  // Average customer rating calculation
  const getCategoryCustomerRatingsData = () => {
    const catRatings: Record<string, { totalRating: number; count: number }> = {};
    filteredRecords.forEach((r) => {
      let rating = 4.2;
      if (r.returnAmt > 0) {
        rating = 1.5 + (r.sales % 1.5);
      } else if (r.discountRate > 0) {
        rating = 4.3 + ((r.sales + r.quantity) % 0.7);
      } else {
        rating = 3.6 + ((r.sales * r.quantity) % 1.3);
      }
      rating = Math.max(1, Math.min(5, Math.round(rating * 10) / 10));

      if (!catRatings[r.category]) {
        catRatings[r.category] = { totalRating: 0, count: 0 };
      }
      catRatings[r.category].totalRating += rating;
      catRatings[r.category].count += 1;
    });

    return Object.entries(catRatings)
      .map(([category, stats]) => {
        const avgRating = stats.count > 0 ? Math.round((stats.totalRating / stats.count) * 100) / 100 : 4.0;
        return {
          category,
          avgRating,
          percentage: Math.round((avgRating / 5) * 100),
          count: stats.count
        };
      })
      .sort((a, b) => b.avgRating - a.avgRating);
  };

  // 8. Stockout Risks on city basis (ascending bar chart with top 3 highlighted as red)
  const getStockoutRisksData = () => {
    const cityBaseStocks: Record<string, number> = {
      "New York": 1150, "Los Angeles": 980, "Chicago": 820, "Houston": 410, "Seattle": 690,
      "Boston": 520, "San Francisco": 730, "Detroit": 340, "Dallas": 490, "Portland": 420,
      "Atlanta": 460, "San Diego": 540, "Minneapolis": 480, "Miami": 360, "Bellevue": 310
    };
    
    const cityMap: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      const city = r.city || "New York";
      cityMap[city] = (cityMap[city] || 0) + r.quantity;
    });

    const citiesList = uniqueCities.filter(c => c !== "All");
    
    return citiesList.map((city) => {
      const unitsSold = cityMap[city] || 0;
      const base = cityBaseStocks[city] || 450;
      const stock = Math.max(15, base - unitsSold);
      return {
        city,
        StockLevel: stock
      };
    }).sort((a, b) => a.StockLevel - b.StockLevel);
  };

  // 9. Marketing spends by Region and City (Side-by-side pie charts)
  const getMarketingSpendsByRegion = () => {
    const regionSpend: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      const spendRate = 0.06 + ((r.region.charCodeAt(0) % 5) * 0.01);
      const saleNet = r.sales - (r.returnAmt || 0);
      const spend = Math.max(0, saleNet * spendRate);
      regionSpend[r.region] = (regionSpend[r.region] || 0) + spend;
    });
    return Object.entries(regionSpend)
      .map(([name, value]) => ({ name, value: Math.round(value) }))
      .sort((a, b) => b.value - a.value);
  };

  const getMarketingSpendsByCity = () => {
    const citySpend: Record<string, number> = {};
    filteredRecords.forEach((r) => {
      const city = r.city || "New York";
      const spendRate = 0.05 + ((city.charCodeAt(0) % 6) * 0.01);
      const saleNet = r.sales - (r.returnAmt || 0);
      const spend = Math.max(0, saleNet * spendRate);
      citySpend[city] = (citySpend[city] || 0) + spend;
    });
    return Object.entries(citySpend)
      .map(([name, value]) => ({ name, value: Math.round(value) }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8); // top 8 to keep pie clear
  };

  // Helper to parse/style Markdown response for AI Insights nicely
  const renderAIResponse = (text: string) => {
    if (!text) return null;
    
    // Split by lines and parse basic structures
    const lines = text.split("\n");
    return lines.map((line, idx) => {
      const trimmed = line.trim();
      
      // Headers
      if (trimmed.startsWith("###")) {
        return (
          <h4 key={idx} className="text-sm font-bold text-indigo-300 mt-5 mb-2 flex items-center gap-2 border-b border-slate-800 pb-1 font-display">
            {trimmed.replace("###", "").trim()}
          </h4>
        );
      }
      if (trimmed.startsWith("##")) {
        return (
          <h3 key={idx} className="text-md font-bold text-white mt-6 mb-3 border-l-4 border-indigo-500 pl-3 font-display">
            {trimmed.replace("##", "").trim()}
          </h3>
        );
      }
      if (trimmed.startsWith("#")) {
        return (
          <h2 key={idx} className="text-lg font-extrabold text-indigo-400 mt-6 mb-4 font-display">
            {trimmed.replace("#", "").trim()}
          </h2>
        );
      }

      // Bullets
      if (trimmed.startsWith("*") || trimmed.startsWith("-")) {
        // Parse bold highlights within lists
        const cleanContent = trimmed.substring(1).trim();
        return (
          <li key={idx} className="ml-5 list-disc text-xs text-slate-300 mb-2 leading-relaxed">
            {parseBoldText(cleanContent)}
          </li>
        );
      }

      // Empty line
      if (trimmed === "") {
        return <div key={idx} className="h-2" />;
      }

      // Normal paragraph
      return (
        <p key={idx} className="text-xs text-slate-300 leading-relaxed mb-3">
          {parseBoldText(trimmed)}
        </p>
      );
    });
  };

  // Parse bold text matches **text**
  const parseBoldText = (text: string) => {
    const parts = text.split(/\*\*([\s\S]*?)\*\*/g);
    if (parts.length === 1) return text;
    
    return parts.map((part, i) => {
      if (i % 2 === 1) {
        return <strong key={i} className="font-semibold text-white">{part}</strong>;
      }
      return part;
    });
  };

  // Master Excel/CSV raw transaction downloader helper based on interactive views and filters
  const downloadTransactionsForFilter = (
    filterType: 'category' | 'month' | 'product' | 'region' | 'week' | 'city' | 'store' | 'all',
    filterValue: string,
    label: string
  ) => {
    const filtered = filteredRecords.filter((r) => {
      if (filterType === 'category') {
        return r.category.toLowerCase() === filterValue.toLowerCase();
      }
      if (filterType === 'month') {
        return r.date.startsWith(filterValue); // "YYYY-MM"
      }
      if (filterType === 'product') {
        return r.product.toLowerCase() === filterValue.toLowerCase();
      }
      if (filterType === 'region') {
        return r.region.toLowerCase() === filterValue.toLowerCase();
      }
      if (filterType === 'week') {
        return r.week === filterValue;
      }
      if (filterType === 'city') {
        return (r.city || "").toLowerCase() === filterValue.toLowerCase();
      }
      if (filterType === 'store') {
        return r.store.toLowerCase() === filterValue.toLowerCase();
      }
      return true;
    });

    if (filtered.length === 0) {
      alert(`No active records found matching selection: "${filterValue}"`);
      return;
    }

    // Format for high-end professional spreadsheet output
    const rows = filtered.map((r) => ({
      "Invoice ID": r.id,
      "Invoice Date": r.date,
      "Week": r.week,
      "Product Name": r.product,
      "Category": r.category,
      "Units Sold": r.quantity,
      "Revenue ($)": r.sales,
      "Net Profit ($)": r.profit,
      "Region": r.region,
      "City": r.city || "N/A",
      "Store Name": r.store,
      "Store Format": r.storeFormat || "N/A",
      "Discount Rate (%)": Math.round((r.discountRate || 0) * 100),
      "Return Amount ($)": r.returnAmt || 0
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Transactions");
    
    // Format download file name cleanly
    const cleanLabel = label.replace(/[^a-z0-9_]/gi, '_').toLowerCase();
    XLSX.writeFile(workbook, `retail_sales_data_${cleanLabel}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-[#0c0e12] text-slate-300 font-sans selection:bg-indigo-500/30 pb-16">
      {/* Upper Brand Bar */}
      <header className="bg-[#11141b] border-b border-slate-800 sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-4 sm:gap-6">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 bg-indigo-600 hover:bg-indigo-500 rounded-lg flex items-center justify-center text-white shadow-lg shadow-indigo-950/30 transition-all">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div className="hidden md:block">
                <h1 className="text-md font-bold text-white tracking-tight font-display">Retail Sales Analyzer</h1>
                <p className="text-[10px] text-slate-500 font-semibold uppercase tracking-wider">Spreadsheet Ingestion & AI Intelligence</p>
              </div>
            </div>

            {/* Business Summary Tab and Dropdown */}
            {salesRecords.length > 0 && (
              <div className="relative">
                <button
                  id="business-summary-tab"
                  onClick={() => setShowBusinessSummary(!showBusinessSummary)}
                  className="flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded-lg bg-indigo-500/10 hover:bg-indigo-500/20 text-indigo-400 border border-indigo-500/20 transition-all cursor-pointer shadow-sm select-none"
                >
                  <Award className="w-3.5 h-3.5 text-indigo-400" />
                  <span>Business Summary</span>
                  <ChevronDown className={`w-3.5 h-3.5 transition-transform duration-200 ${showBusinessSummary ? "rotate-180" : ""}`} />
                </button>

                {showBusinessSummary && (
                  <>
                    {/* Backdrop to close */}
                    <div className="fixed inset-0 z-40" onClick={() => setShowBusinessSummary(false)} />
                    <div className="absolute left-0 mt-2 w-80 sm:w-96 bg-[#11141b] border border-slate-800 rounded-xl shadow-2xl shadow-black/95 z-50 p-4 font-sans text-xs divide-y divide-slate-800/60 max-h-[80vh] overflow-y-auto">
                      
                      {/* Header */}
                      <div className="pb-3 flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <Award className="w-4 h-4 text-amber-400 animate-bounce" />
                          <span className="font-bold text-white text-xs sm:text-sm">Strategic Business Summary</span>
                        </div>
                        <span className="text-[9px] text-slate-500 bg-slate-800/60 px-2 py-0.5 rounded font-mono uppercase">Live Data</span>
                      </div>

                      {/* Section 1: Worst and Best Regions as per sales */}
                      <div className="py-3.5">
                        <h4 className="font-bold text-slate-300 mb-2 flex items-center gap-1.5">
                          <MapPin className="w-3.5 h-3.5 text-indigo-400" />
                          Region Performance (by Net Sales)
                        </h4>
                        {(() => {
                          const { best, worst } = getRegionsPerformance();
                          return (
                            <div className="grid grid-cols-2 gap-2 mt-1">
                              <div className="p-2 rounded-lg bg-emerald-500/5 border border-emerald-500/10">
                                <span className="text-[9px] text-slate-500 block uppercase font-bold tracking-wider mb-0.5">Best Region</span>
                                <span className="font-bold text-emerald-400 text-xs block truncate" title={best.name}>{best.name}</span>
                                <span className="font-semibold text-slate-300 text-[10px]">${best.value.toLocaleString()}</span>
                              </div>
                              <div className="p-2 rounded-lg bg-rose-500/5 border border-rose-500/10">
                                <span className="text-[9px] text-slate-500 block uppercase font-bold tracking-wider mb-0.5">Worst Region</span>
                                <span className="font-bold text-rose-400 text-xs block truncate" title={worst.name}>{worst.name}</span>
                                <span className="font-semibold text-slate-300 text-[10px]">${worst.value.toLocaleString()}</span>
                              </div>
                            </div>
                          );
                        })()}
                      </div>

                      {/* Section 2: High returns categories */}
                      <div className="py-3.5">
                        <h4 className="font-bold text-slate-300 mb-2 flex items-center gap-1.5">
                          <RefreshCw className="w-3.5 h-3.5 text-amber-400" />
                          High Return Categories (Top Returns)
                        </h4>
                        {(() => {
                          const categories = getHighReturnsCategories();
                          if (categories.length === 0) {
                            return <div className="text-slate-500 italic p-1">No returns reported.</div>;
                          }
                          return (
                            <div className="space-y-1.5 mt-1">
                              {categories.map((c, idx) => (
                                <div key={idx} className="flex items-center justify-between p-2 rounded-lg bg-[#0c0e12] hover:bg-[#161a23] border border-slate-800/80 transition-colors">
                                  <span className="font-medium text-slate-300 truncate max-w-[140px]">{c.category}</span>
                                  <div className="text-right">
                                    <div className="font-semibold text-amber-400">${c.returns.toLocaleString()}</div>
                                    <div className="text-[10px] text-slate-500">{c.rate}% return rate</div>
                                  </div>
                                </div>
                              ))}
                            </div>
                          );
                        })()}
                      </div>

                      {/* Section 3: Stores missing targets */}
                      <div className="pt-3.5">
                        <h4 className="font-bold text-slate-300 mb-2 flex items-center gap-1.5">
                          <TrendingDown className="w-3.5 h-3.5 text-rose-400" />
                          Stores Missing Targets (Target: ${Math.max(5000, Math.round(salesTarget / (uniqueStores.filter(s => s !== "All").length || 1))).toLocaleString()})
                        </h4>
                        {(() => {
                          const stores = getStoresMissingTargets();
                          if (stores.length === 0) {
                            return (
                              <div className="flex items-center gap-1.5 p-2 rounded-lg bg-emerald-500/5 border border-emerald-500/10 text-emerald-400 mt-1">
                                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                                <span className="font-semibold">All stores hit their performance targets!</span>
                              </div>
                            );
                          }
                          return (
                            <div className="space-y-1.5 max-h-48 overflow-y-auto pr-1 mt-1">
                              {stores.map((s, idx) => (
                                <div key={idx} className="p-2 rounded-lg bg-[#0c0e12] border border-slate-800/80 space-y-1">
                                  <div className="flex items-center justify-between">
                                    <span className="font-bold text-slate-300 truncate max-w-[140px]">{s.store}</span>
                                    <span className="font-semibold text-rose-400">-${s.missingAmount.toLocaleString()}</span>
                                  </div>
                                  <div className="w-full bg-slate-800 rounded-full h-1.5 overflow-hidden">
                                    <div className="bg-rose-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, s.percentage)}%` }} />
                                  </div>
                                  <div className="flex items-center justify-between text-[10px] text-slate-500">
                                    <span>Achieved: {s.percentage}%</span>
                                    <span>Sales: ${s.sales.toLocaleString()}</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          );
                        })()}
                      </div>

                    </div>
                  </>
                )}
              </div>
            )}
          </div>

          <div className="flex items-center gap-2 sm:gap-3">
            <button
              onClick={downloadSampleExcel}
              className="flex items-center gap-2 text-xs font-semibold text-indigo-400 bg-indigo-500/10 hover:bg-indigo-500/20 px-3 py-2 rounded-xl transition-all border border-indigo-500/20 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Download Sample Sheet</span>
              <span className="sm:hidden">Sample</span>
            </button>
            
            <button
              onClick={handleLoadDemoData}
              className={`flex items-center gap-2 text-xs font-semibold px-3 py-2 rounded-xl transition-all cursor-pointer border ${
                isDemoMode 
                  ? "bg-slate-800 text-slate-400 border-slate-700 pointer-events-none opacity-60"
                  : "bg-[#161a23] hover:bg-[#1a202c] text-slate-300 border-slate-800 hover:border-slate-700"
              }`}
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Reset to Demo</span>
              <span className="sm:hidden">Reset</span>
            </button>
          </div>
        </div>
      </header>

      {/* Global Dashboard Filters - Bento Layout */}
      {salesRecords.length > 0 && (
        <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-6">
          <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-5 shadow-xl">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4 pb-3 border-b border-slate-800/60">
              <div className="flex items-center gap-2">
                <div className="p-1.5 rounded-lg bg-indigo-500/10 text-indigo-400 border border-indigo-500/20">
                  <Sliders className="w-4 h-4" />
                </div>
                <div>
                  <h2 className="text-xs font-bold text-slate-300 uppercase tracking-widest">Dashboard Filtration Hub</h2>
                  <p className="text-[10px] text-slate-500 font-medium">Refine metrics, Swot report, and visual chart channels in real-time</p>
                </div>
              </div>
              
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => {
                    setSelectedWeek("All");
                    setSelectedRegion("All");
                    setSelectedCity("All");
                    setSelectedStore("All");
                    setSelectedStoreFormat("All");
                    setSelectedCategory("All");

                    setTempRegion("All");
                    setTempCity("All");
                    setTempStore("All");
                    setTempStoreFormat("All");
                    setTempCategory("All");
                    setCurrentPage(1);
                  }}
                  className="text-[10px] text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 px-2.5 py-1 rounded-md transition-all cursor-pointer font-semibold border border-slate-700"
                >
                  Reset Filters
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 items-end">
              {/* Date Filter (Instant dropdown selection in descending order) */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-indigo-400 uppercase tracking-wider flex items-center gap-1">
                  <span>Date (Instant)</span>
                  <span className="text-[8px] px-1 bg-indigo-500/15 rounded text-indigo-300">Auto</span>
                </label>
                <select
                  value={selectedWeek}
                  onChange={(e) => {
                    setSelectedWeek(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="w-full bg-[#11141b] border border-slate-800 hover:border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 outline-none focus:border-indigo-500 cursor-pointer"
                >
                  <option value="All">All Dates ({uniqueWeeks.length - 1})</option>
                  {uniqueWeeks.filter(w => w !== "All").map((wk) => (
                    <option key={wk} value={wk}>{wk}</option>
                  ))}
                </select>
              </div>

              {/* Region Filter (Draft) */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Region</label>
                <select
                  value={tempRegion}
                  onChange={(e) => setTempRegion(e.target.value)}
                  className="w-full bg-[#11141b] border border-slate-800 hover:border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 outline-none focus:border-indigo-500 cursor-pointer"
                >
                  <option value="All">All Regions</option>
                  {uniqueRegions.filter(r => r !== "All").map((reg) => (
                    <option key={reg} value={reg}>{reg}</option>
                  ))}
                </select>
              </div>

              {/* City Filter (Draft) */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">City</label>
                <select
                  value={tempCity}
                  onChange={(e) => setTempCity(e.target.value)}
                  className="w-full bg-[#11141b] border border-slate-800 hover:border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 outline-none focus:border-indigo-500 cursor-pointer"
                >
                  <option value="All">All Cities</option>
                  {uniqueCities.filter(c => c !== "All").map((ct) => (
                    <option key={ct} value={ct}>{ct}</option>
                  ))}
                </select>
              </div>

              {/* Store Filter (Draft) */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Store</label>
                <select
                  value={tempStore}
                  onChange={(e) => setTempStore(e.target.value)}
                  className="w-full bg-[#11141b] border border-slate-800 hover:border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 outline-none focus:border-indigo-500 cursor-pointer"
                >
                  <option value="All">All Stores</option>
                  {uniqueStores.filter(s => s !== "All").map((st) => (
                    <option key={st} value={st}>{st}</option>
                  ))}
                </select>
              </div>

              {/* Store Format Filter (Draft) */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Store Format</label>
                <select
                  value={tempStoreFormat}
                  onChange={(e) => setTempStoreFormat(e.target.value)}
                  className="w-full bg-[#11141b] border border-slate-800 hover:border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 outline-none focus:border-indigo-500 cursor-pointer"
                >
                  <option value="All">All Formats</option>
                  {uniqueStoreFormats.filter(sf => sf !== "All").map((fmt) => (
                    <option key={fmt} value={fmt}>{fmt}</option>
                  ))}
                </select>
              </div>

              {/* Product Category Filter (Draft) */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Product Category</label>
                <select
                  value={tempCategory}
                  onChange={(e) => setTempCategory(e.target.value)}
                  className="w-full bg-[#11141b] border border-slate-800 hover:border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 outline-none focus:border-indigo-500 cursor-pointer"
                >
                  <option value="All">All Categories</option>
                  {uniqueCategories.filter(cat => cat !== "All").map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="mt-4 pt-3 border-t border-slate-800/40 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div className="flex flex-wrap items-center gap-2 text-[10px] text-slate-400">
                <span className="font-semibold">Active Filter Bounds:</span>
                <span className="bg-[#11141b] px-2 py-0.5 rounded border border-slate-800 text-slate-300">
                  Date: {selectedWeek}
                </span>
                {(selectedRegion !== "All" || selectedCity !== "All" || selectedStore !== "All" || selectedStoreFormat !== "All" || selectedCategory !== "All") ? (
                  <>
                    {selectedRegion !== "All" && <span className="bg-indigo-500/10 text-indigo-300 px-2 py-0.5 rounded border border-indigo-500/15">Reg: {selectedRegion}</span>}
                    {selectedCity !== "All" && <span className="bg-indigo-500/10 text-indigo-300 px-2 py-0.5 rounded border border-indigo-500/15">City: {selectedCity}</span>}
                    {selectedStore !== "All" && <span className="bg-indigo-500/10 text-indigo-300 px-2 py-0.5 rounded border border-indigo-500/15">Store: {selectedStore}</span>}
                    {selectedStoreFormat !== "All" && <span className="bg-indigo-500/10 text-indigo-300 px-2 py-0.5 rounded border border-indigo-500/15">Fmt: {selectedStoreFormat}</span>}
                    {selectedCategory !== "All" && <span className="bg-indigo-500/10 text-indigo-300 px-2 py-0.5 rounded border border-indigo-500/15">Cat: {selectedCategory}</span>}
                  </>
                ) : (
                  <span className="text-slate-500 italic">No static dimensions active. Showing overall records.</span>
                )}
              </div>

              <button
                onClick={() => {
                  setSelectedRegion(tempRegion);
                  setSelectedCity(tempCity);
                  setSelectedStore(tempStore);
                  setSelectedStoreFormat(tempStoreFormat);
                  setSelectedCategory(tempCategory);
                  setCurrentPage(1);
                }}
                className="flex items-center justify-center gap-1.5 px-5 py-2.5 rounded-xl text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white cursor-pointer shadow-lg shadow-indigo-950/20 transition-all border border-indigo-500/20 uppercase tracking-wider"
              >
                <Check className="w-3.5 h-3.5" />
                Apply Filters
              </button>
            </div>
          </div>
        </section>
      )}

      {/* Main Grid: Bento Dashboard Arrangement */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-6 grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        {/* LEFT COLUMN: Controls & Upload (Span 4) */}
        <section className="lg:col-span-4 flex flex-col gap-6">
          
          {/* File Ingestion Card */}
          <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-xl flex flex-col gap-5">
            <div>
              <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">Source Pipeline</span>
              <h2 className="text-md font-bold text-white flex items-center gap-2 font-display">
                <Upload className="w-4.5 h-4.5 text-indigo-400" />
                Upload Spreadsheet Files
              </h2>
              <p className="text-xs text-slate-400 mt-1">
                Upload one or multiple sales spreadsheets to compile a unified retail sales database.
              </p>
            </div>

            {/* Drag & Drop Zone */}
            <div
              onDragOver={handleDragOver}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-slate-700 hover:border-indigo-500 bg-[#11141b]/50 hover:bg-indigo-500/5 rounded-xl p-8 text-center cursor-pointer transition-all flex flex-col items-center justify-center gap-2 group"
            >
              <input
                type="file"
                ref={fileInputRef}
                onChange={(e) => e.target.files && handleFileSelected(e.target.files)}
                accept=".xlsx, .xls"
                multiple
                className="hidden"
              />
              <div className="w-10 h-10 bg-[#161a23] rounded-lg shadow-sm border border-slate-800 flex items-center justify-center text-slate-500 group-hover:text-indigo-400 group-hover:scale-110 transition-all">
                <Upload className="w-5 h-5" />
              </div>
              <div>
                <p className="text-xs font-semibold text-slate-300">Drop Excel files here or browse</p>
                <p className="text-[10px] text-slate-500 mt-1">Supports standard columns (.xlsx, .xls)</p>
              </div>
            </div>

            {/* File List / Staged Queue */}
            {loadedFiles.length > 0 && (
              <div className="flex flex-col gap-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase text-slate-500 tracking-wider flex items-center gap-1.5">
                    <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-500" />
                    Staged Queue ({loadedFiles.length})
                  </span>
                  <button 
                    onClick={() => { setLoadedFiles([]); setAvailableHeaders([]); setShowMappingConfig(false); }}
                    className="text-[10px] text-indigo-400 hover:underline cursor-pointer font-semibold"
                  >
                    Clear All
                  </button>
                </div>

                <div className="max-h-48 overflow-y-auto divide-y divide-slate-800/80 border border-slate-800 rounded-lg pr-1">
                  <AnimatePresence initial={false}>
                    {loadedFiles.map((file, i) => (
                      <motion.div
                        key={file.name + i}
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: "auto" }}
                        exit={{ opacity: 0, height: 0 }}
                        className="flex items-center justify-between py-2.5 px-2 text-xs bg-[#11141b]/40 rounded-md mb-1"
                      >
                        <div className="flex items-center gap-2 overflow-hidden mr-2">
                          <FileSpreadsheet className="w-4 h-4 text-emerald-500 flex-shrink-0" />
                          <div className="truncate">
                            <p className="font-medium text-slate-200 truncate">{file.name}</p>
                            <p className="text-[10px] text-slate-500">{file.size} • {file.rows} rows</p>
                          </div>
                        </div>
                        <button
                          onClick={() => handleRemoveFile(i)}
                          className="text-slate-500 hover:text-red-400 p-1 rounded-md transition-colors hover:bg-slate-800 cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </motion.div>
                    ))}
                  </AnimatePresence>
                </div>
              </div>
            )}

            {/* Custom Field Mapping Config */}
            {showMappingConfig && (
              <div className="border border-amber-500/10 bg-amber-500/5 rounded-xl p-4 flex flex-col gap-3">
                <div className="flex items-center gap-2">
                  <Sliders className="w-4 h-4 text-amber-500" />
                  <span className="text-xs font-bold text-amber-200">Verify Spreadsheet Columns</span>
                  <div className="text-[9px] bg-amber-500/10 text-amber-400 px-1.5 py-0.5 rounded-full font-bold ml-auto">
                    Auto-Mapped
                  </div>
                </div>
                <p className="text-[11px] text-slate-400 leading-normal">
                  Our algorithm auto-mapped your columns. Please verify alignment below before final compile:
                </p>

                <div className="grid grid-cols-2 gap-2 text-[10px]">
                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Invoice Date</label>
                    <select
                      value={columnMapping.date}
                      onChange={(e) => setColumnMapping({ ...columnMapping, date: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Product Name</label>
                    <select
                      value={columnMapping.product}
                      onChange={(e) => setColumnMapping({ ...columnMapping, product: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Category</label>
                    <select
                      value={columnMapping.category}
                      onChange={(e) => setColumnMapping({ ...columnMapping, category: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Revenue</label>
                    <select
                      value={columnMapping.sales}
                      onChange={(e) => setColumnMapping({ ...columnMapping, sales: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Units Sold</label>
                    <select
                      value={columnMapping.quantity}
                      onChange={(e) => setColumnMapping({ ...columnMapping, quantity: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Net Profit</label>
                    <select
                      value={columnMapping.profit}
                      onChange={(e) => setColumnMapping({ ...columnMapping, profit: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Apply 30% margin fallback --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Store Region</label>
                    <select
                      value={columnMapping.region}
                      onChange={(e) => setColumnMapping({ ...columnMapping, region: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Single HQ Store --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Discount Rate</label>
                    <select
                      value={columnMapping.discountRate}
                      onChange={(e) => setColumnMapping({ ...columnMapping, discountRate: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Auto-Simulate (0% - 20%) --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div className="col-span-2">
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Return Amount</label>
                    <select
                      value={columnMapping.returnAmt}
                      onChange={(e) => setColumnMapping({ ...columnMapping, returnAmt: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Auto-Simulate returns --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Invoice Week</label>
                    <select
                      value={columnMapping.week}
                      onChange={(e) => setColumnMapping({ ...columnMapping, week: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Auto-Extract from Date --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">City</label>
                    <select
                      value={columnMapping.city}
                      onChange={(e) => setColumnMapping({ ...columnMapping, city: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Auto-Simulate by Region --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Store Name</label>
                    <select
                      value={columnMapping.store}
                      onChange={(e) => setColumnMapping({ ...columnMapping, store: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Auto-Simulate store codes --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 mb-0.5 font-semibold uppercase">Store Format</label>
                    <select
                      value={columnMapping.storeFormat}
                      onChange={(e) => setColumnMapping({ ...columnMapping, storeFormat: e.target.value })}
                      className="w-full bg-[#11141b] border border-slate-800 rounded p-1 text-slate-200 outline-none focus:border-indigo-500"
                    >
                      <option value="">-- Auto-Simulate formats --</option>
                      {availableHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                </div>
              </div>
            )}

            {/* Ingestion Submit Button */}
            <button
              onClick={handleSubmitAndAnalyze}
              disabled={loadedFiles.length === 0}
              className={`w-full flex items-center justify-center gap-2 py-3 rounded-xl font-semibold text-sm transition-all cursor-pointer shadow-lg ${
                loadedFiles.length > 0 
                  ? "bg-indigo-600 hover:bg-indigo-500 text-white shadow-indigo-950/20" 
                  : "bg-slate-800 text-slate-500 border border-slate-700 shadow-none cursor-not-allowed"
              }`}
            >
              <Check className="w-4 h-4" />
              Submit and Compile {loadedFiles.length > 0 ? `(${loadedFiles.length} files)` : ""}
            </button>
          </div>

          {/* AI Strategic Intelligence Advisor Panel */}
          {salesRecords.length > 0 && (
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-xl flex flex-col gap-4">
              <div className="flex items-center justify-between">
                <div>
                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-0.5">Decision Suite</span>
                  <h2 className="text-md font-bold text-white flex items-center gap-2 font-display">
                    <Sparkles className="w-4.5 h-4.5 text-indigo-400" />
                    Gemini AI Advisor
                  </h2>
                </div>
                <span className="text-[10px] bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 px-2 py-0.5 rounded-full font-semibold">
                  Flash 3.5
                </span>
              </div>

              <p className="text-xs text-slate-400 leading-relaxed">
                Consolidate compiled sales transactions and generate strategic performance reviews with our retail intelligence agent.
              </p>

              {/* Custom Query Request Input */}
              <div className="flex items-center gap-2 border border-slate-800 rounded-xl p-1.5 focus-within:border-indigo-500 transition-all bg-[#11141b]">
                <input
                  type="text"
                  placeholder={isDemoMode ? "Generate performance review..." : "Ask specific metrics, trends..."}
                  disabled={isAnalyzing || salesRecords.length === 0}
                  value={customPrompt}
                  onChange={(e) => setCustomPrompt(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && customPrompt.trim() && triggerAIAnalysis(salesRecords, customPrompt)}
                  className="w-full bg-transparent px-2 py-1 text-xs outline-none text-slate-200 placeholder-slate-500 disabled:cursor-not-allowed"
                />
                <button
                  onClick={() => customPrompt.trim() && triggerAIAnalysis(salesRecords, customPrompt)}
                  disabled={isAnalyzing || !customPrompt.trim() || salesRecords.length === 0}
                  className="bg-indigo-600 text-white p-1.5 rounded-lg hover:bg-indigo-500 disabled:bg-slate-800 disabled:text-slate-600 transition-all cursor-pointer flex items-center justify-center flex-shrink-0"
                >
                  <Send className="w-3.5 h-3.5" />
                </button>
              </div>

              {/* Standard Review Trigger Button */}
              {!aiAnalysis && !isAnalyzing && (
                <button
                  onClick={() => triggerAIAnalysis(salesRecords)}
                  disabled={salesRecords.length === 0}
                  className="w-full flex items-center justify-center gap-2 py-2.5 bg-indigo-500/10 hover:bg-indigo-500/20 text-indigo-400 border border-indigo-500/20 rounded-xl text-xs font-semibold cursor-pointer transition-colors"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  Analyze Sales Performance
                </button>
              )}

              {/* AI Advisor Response Area */}
              <div className="mt-1">
                <AnimatePresence mode="wait">
                  {isAnalyzing ? (
                    <motion.div 
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="flex flex-col items-center justify-center py-10 text-center gap-3"
                    >
                      <RefreshCw className="w-7 h-7 text-indigo-400 animate-spin" />
                      <div>
                        <p className="text-xs font-semibold text-slate-300">Consulting AI Advisor...</p>
                        <p className="text-[10px] text-slate-500 mt-1 max-w-[200px] mx-auto animate-pulse">
                          {loadingStep}
                        </p>
                      </div>
                    </motion.div>
                  ) : analysisError ? (
                    <motion.div 
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      className="p-4 bg-red-500/10 border border-red-500/20 rounded-xl flex gap-2.5 text-xs text-red-400"
                    >
                      <AlertCircle className="w-4 h-4 flex-shrink-0 text-red-500 mt-0.5" />
                      <div>
                        <p className="font-bold">Advisor Off-line</p>
                        <p className="mt-1 text-[11px] leading-relaxed opacity-90">{analysisError}</p>
                      </div>
                    </motion.div>
                  ) : aiAnalysis ? (
                    <motion.div 
                      initial={{ opacity: 0, y: 5 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="bg-[#11141b]/50 border border-slate-800 rounded-xl p-4 max-h-[350px] overflow-y-auto"
                    >
                      <div className="flex items-center gap-1.5 text-xs font-bold text-white mb-3 border-b border-slate-800 pb-2 font-display">
                        <Sparkles className="w-3.5 h-3.5 text-indigo-400" />
                        Gemini Retail SWOT Report
                      </div>
                      <div className="prose prose-sm text-slate-300">
                        {renderAIResponse(aiAnalysis)}
                      </div>
                    </motion.div>
                  ) : (
                    <div className="text-center py-6 text-slate-500 border border-dashed border-slate-800 rounded-xl flex flex-col items-center justify-center gap-1">
                      <HelpCircle className="w-5 h-5 text-slate-600" />
                      <p className="text-[10px]">No active advisory insights loaded.</p>
                    </div>
                  )}
                </AnimatePresence>
              </div>
            </div>
          )}
        </section>

        {/* RIGHT COLUMN: Key Metrics & Data Dashboard (Span 8) */}
        <section className="lg:col-span-8 flex flex-col gap-6">
          {salesRecords.length === 0 ? (
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-8 shadow-xl flex flex-col items-center justify-center text-center min-h-[500px]">
              <div className="w-16 h-16 bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 rounded-2xl flex items-center justify-center mb-6 shadow-lg shadow-indigo-950/20">
                <FileSpreadsheet className="w-8 h-8" />
              </div>
              
              <h2 className="text-xl font-bold text-white tracking-tight font-display mb-2">
                Awaiting Sales Data Ingestion
              </h2>
              <p className="text-xs text-slate-400 max-w-md mx-auto mb-8 leading-relaxed">
                Connect your business intelligence ecosystem. Upload or drag-and-drop your store sales, orders, and returns spreadsheet to unlock full interactive visualizations, filters, and SWOT advice.
              </p>

              <div className="w-full max-w-xl grid grid-cols-1 sm:grid-cols-2 gap-4 text-left">
                <div className="bg-[#11141b] border border-slate-800/80 p-4.5 rounded-xl flex items-start gap-3">
                  <div className="w-7 h-7 rounded-lg bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 flex items-center justify-center flex-shrink-0 text-xs font-bold font-mono">
                    1
                  </div>
                  <div>
                    <h4 className="text-xs font-bold text-slate-200 uppercase tracking-wide">Download Template</h4>
                    <p className="text-[11px] text-slate-500 mt-1 leading-normal">
                      Click the "Download Sample Sheet" button above to obtain a pre-formatted template.
                    </p>
                  </div>
                </div>

                <div className="bg-[#11141b] border border-slate-800/80 p-4.5 rounded-xl flex items-start gap-3">
                  <div className="w-7 h-7 rounded-lg bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 flex items-center justify-center flex-shrink-0 text-xs font-bold font-mono">
                    2
                  </div>
                  <div>
                    <h4 className="text-xs font-bold text-slate-200 uppercase tracking-wide">Ingest Files</h4>
                    <p className="text-[11px] text-slate-500 mt-1 leading-normal">
                      Drop one or more .xlsx worksheets into the Source Pipeline to load your records.
                    </p>
                  </div>
                </div>

                <div className="bg-[#11141b] border border-slate-800/80 p-4.5 rounded-xl flex items-start gap-3">
                  <div className="w-7 h-7 rounded-lg bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 flex items-center justify-center flex-shrink-0 text-xs font-bold font-mono">
                    3
                  </div>
                  <div>
                    <h4 className="text-xs font-bold text-slate-200 uppercase tracking-wide">Map & Verify</h4>
                    <p className="text-[11px] text-slate-500 mt-1 leading-normal">
                      Confirm or adjust detected column headers to ensure correct metric parsing.
                    </p>
                  </div>
                </div>

                <div className="bg-[#11141b] border border-slate-800/80 p-4.5 rounded-xl flex items-start gap-3">
                  <div className="w-7 h-7 rounded-lg bg-[#6366f1]/20 border border-[#6366f1]/30 text-indigo-400 flex items-center justify-center flex-shrink-0 text-xs font-bold font-mono animate-pulse">
                    4
                  </div>
                  <div>
                    <h4 className="text-xs font-bold text-slate-200 uppercase tracking-wide">Explore Dashboard</h4>
                    <p className="text-[11px] text-slate-500 mt-1 leading-normal">
                      Submit to instantly populate KPI metrics, Pareto charts, regional trends, and Gemini AI reports!
                    </p>
                  </div>
                </div>
              </div>

              {/* Quick action button for demo */}
              <div className="mt-8 flex items-center gap-3">
                <span className="text-[11px] text-slate-500">Want to see how it works?</span>
                <button
                  onClick={handleLoadDemoData}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-indigo-500/10 hover:bg-indigo-500/20 text-indigo-400 border border-indigo-500/25 text-[11px] font-semibold transition-all cursor-pointer"
                >
                  <RefreshCw className="w-3 h-3" />
                  Load Sandbox Demo Data
                </button>
              </div>
            </div>
          ) : (
            <>
              {/* Demo Alert Box */}
          {isDemoMode && (
            <div className="bg-indigo-950/20 border border-indigo-500/20 rounded-2xl p-4 flex items-start gap-3 shadow-sm">
              <Info className="w-5 h-5 text-indigo-400 flex-shrink-0 mt-0.5" />
              <div className="text-xs text-indigo-200">
                <p className="font-semibold">Interactive Demo Mode Active</p>
                <p className="mt-1 leading-relaxed opacity-80">
                  The dashboard displays synthetic retail dataset transactions. Upload your custom Excel worksheets on the left pane and hit **Submit** to map, compile, and visualize real transactions.
                </p>
              </div>
            </div>
          )}

          {/* Core Retail KPI Metrics Grid */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            
            {/* Net Sales Card */}
            <div 
              id="kpi-net-sales"
              onClick={() => setActiveKpiFilter("all")}
              className={`rounded-2xl p-4.5 shadow-md flex flex-col justify-between h-34 transition-all duration-200 cursor-pointer border ${
                activeKpiFilter === "all" 
                  ? "bg-[#1f2431] border-indigo-500 shadow-indigo-950/40 ring-1 ring-indigo-500/30" 
                  : "bg-[#161a23] border-slate-800 hover:border-slate-700 hover:bg-[#1d212c]"
              }`}
            >
              <div className="flex items-center justify-between">
                <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Net Sales</span>
                <div className={`p-1.5 rounded-lg border transition-all ${
                  activeKpiFilter === "all" 
                    ? "bg-indigo-500/20 text-indigo-300 border-indigo-500/30" 
                    : "bg-slate-800/50 text-slate-400 border-slate-700/50"
                }`}>
                  <DollarSign className="w-3.5 h-3.5" />
                </div>
              </div>
              <div className="mt-1.5">
                <p className="text-lg font-extrabold text-white truncate font-display">
                  ${netSalesVal.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                </p>
                <p className="text-[9px] text-slate-500 mt-0.5 truncate">
                  Gross: ${totalGrossSales.toLocaleString(undefined, { maximumFractionDigits: 0 })} | Ret: ${totalReturnsVal.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                </p>
              </div>
            </div>

            {/* Targets Achieved Card */}
            <div 
              id="kpi-targets-achieved"
              className="bg-[#161a23] rounded-2xl border border-slate-800 p-4.5 shadow-md flex flex-col justify-between h-34 hover:border-slate-700 hover:bg-[#1d212c] transition-all duration-200"
            >
              <div className="flex items-center justify-between">
                <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Target Achieved</span>
                <div className="bg-emerald-500/10 text-emerald-400 p-1.5 rounded-lg border border-emerald-500/20">
                  <Check className="w-3.5 h-3.5" />
                </div>
              </div>
              <div className="mt-1">
                <div className="flex items-baseline justify-between gap-1">
                  <p className="text-lg font-extrabold text-emerald-400 font-display">
                    {targetAchievedVal.toFixed(1)}%
                  </p>
                </div>
                {/* Custom target input & slider */}
                <div className="mt-1 flex flex-col gap-1">
                  <div className="flex items-center justify-between text-[8px] text-slate-500">
                    <span>Target: ${Math.round(salesTarget/1000)}k</span>
                  </div>
                  <input 
                    type="range"
                    min={Math.max(10000, Math.round(netSalesVal * 0.2))}
                    max={Math.max(100000, Math.round(netSalesVal * 2))}
                    step={10000}
                    value={salesTarget}
                    onChange={(e) => setSalesTarget(Number(e.target.value))}
                    className="w-full accent-emerald-500 h-1 rounded-lg bg-slate-800 cursor-pointer"
                  />
                </div>
              </div>
            </div>

            {/* Average Transaction Value Card */}
            <div 
              id="kpi-avg-transaction"
              onClick={() => setActiveKpiFilter(activeKpiFilter === "high-value" ? "all" : "high-value")}
              className={`rounded-2xl p-4.5 shadow-md flex flex-col justify-between h-34 transition-all duration-200 cursor-pointer border ${
                activeKpiFilter === "high-value" 
                  ? "bg-[#1f2431] border-amber-500 shadow-amber-950/40 ring-1 ring-amber-500/30" 
                  : "bg-[#161a23] border-slate-800 hover:border-slate-700 hover:bg-[#1d212c]"
              }`}
            >
              <div className="flex items-center justify-between">
                <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Avg Ticket</span>
                <div className={`p-1.5 rounded-lg border transition-all ${
                  activeKpiFilter === "high-value" 
                    ? "bg-amber-500/20 text-amber-300 border-amber-500/30" 
                    : "bg-slate-800/50 text-slate-400 border-slate-700/50"
                }`}>
                  <ShoppingCart className="w-3.5 h-3.5" />
                </div>
              </div>
              <div className="mt-1.5">
                <p className="text-lg font-extrabold text-white truncate font-display">
                  ${avgOrderVal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </p>
                <p className="text-[9px] text-slate-500 mt-0.5 truncate">
                  {activeKpiFilter === "high-value" ? "● Filtering: > Avg" : "Click to view > Avg tickets"}
                </p>
              </div>
            </div>

            {/* Return Rate Card */}
            <div 
              id="kpi-return-rate"
              onClick={() => setActiveKpiFilter(activeKpiFilter === "returns" ? "all" : "returns")}
              className={`rounded-2xl p-4.5 shadow-md flex flex-col justify-between h-34 transition-all duration-200 cursor-pointer border ${
                activeKpiFilter === "returns" 
                  ? "bg-[#1f2431] border-rose-500 shadow-rose-950/40 ring-1 ring-rose-500/30" 
                  : "bg-[#161a23] border-slate-800 hover:border-slate-700 hover:bg-[#1d212c]"
              }`}
            >
              <div className="flex items-center justify-between">
                <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Return Rate</span>
                <div className={`p-1.5 rounded-lg border transition-all ${
                  activeKpiFilter === "returns" 
                    ? "bg-rose-500/20 text-rose-300 border-rose-500/30" 
                    : "bg-slate-800/50 text-slate-400 border-slate-700/50"
                }`}>
                  <RefreshCw className="w-3.5 h-3.5" />
                </div>
              </div>
              <div className="mt-1.5">
                <p className="text-lg font-extrabold text-rose-400 truncate font-display">
                  {returnRateVal.toFixed(2)}%
                </p>
                <p className="text-[9px] text-slate-500 mt-0.5 truncate">
                  Total: ${totalReturnsVal.toLocaleString(undefined, { maximumFractionDigits: 0 })} refund
                </p>
              </div>
            </div>

            {/* Discount Rate Card */}
            <div 
              id="kpi-discount-rate"
              onClick={() => setActiveKpiFilter(activeKpiFilter === "discounts" ? "all" : "discounts")}
              className={`rounded-2xl p-4.5 col-span-2 md:col-span-1 shadow-md flex flex-col justify-between h-34 transition-all duration-200 cursor-pointer border ${
                activeKpiFilter === "discounts" 
                  ? "bg-[#1f2431] border-purple-500 shadow-purple-950/40 ring-1 ring-purple-500/30" 
                  : "bg-[#161a23] border-slate-800 hover:border-slate-700 hover:bg-[#1d212c]"
              }`}
            >
              <div className="flex items-center justify-between">
                <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Discount Rate</span>
                <div className={`p-1.5 rounded-lg border transition-all ${
                  activeKpiFilter === "discounts" 
                    ? "bg-purple-500/20 text-purple-300 border-purple-500/30" 
                    : "bg-slate-800/50 text-slate-400 border-slate-700/50"
                }`}>
                  <Percent className="w-3.5 h-3.5" />
                </div>
              </div>
              <div className="mt-1.5">
                <p className="text-lg font-extrabold text-white truncate font-display">
                  {avgDiscountRateVal.toFixed(1)}%
                </p>
                <p className="text-[9px] text-slate-500 mt-0.5 truncate">
                  Promo discount: ${totalDiscountsVal.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                </p>
              </div>
            </div>
          </div>

          {/* CHARTS CONTAINER BENTO GRID */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            
            {/* Chart 1: Revenue Trend (Span 2) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md col-span-1 md:col-span-2 flex flex-col">
              <div className="flex items-center justify-between mb-6">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <TrendingUp className="w-4 h-4 text-indigo-400" />
                    Chronological Sales & Net Profit
                  </h3>
                  <p className="text-[11px] text-slate-500">Sales and profit tracking Over Time. Click any point or vertical bar line to download monthly transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'chronological_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Trend Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-64 w-full">
                {getMonthlyTrendData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500">
                    No timeline matches selected filters.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart
                      data={getMonthlyTrendData()}
                      margin={{ top: 10, right: 10, left: -20, bottom: 0 }}
                      onClick={(data: any) => {
                        if (data && data.activeLabel) {
                          downloadTransactionsForFilter('month', String(data.activeLabel), `${data.activeLabel}_monthly`);
                        }
                      }}
                      className="cursor-pointer"
                    >
                      <defs>
                        <linearGradient id="colorSales" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#6366f1" stopOpacity={0.2}/>
                          <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                        </linearGradient>
                        <linearGradient id="colorProfit" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#10b981" stopOpacity={0.2}/>
                          <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#1f2937" />
                      <XAxis 
                        dataKey="month" 
                        stroke="#4b5563" 
                        fontSize={10} 
                        tickLine={false}
                        axisLine={false}
                      />
                      <YAxis 
                        stroke="#4b5563" 
                        fontSize={10} 
                        tickLine={false}
                        axisLine={false}
                        tickFormatter={(v) => `$${v}`}
                      />
                      <Tooltip 
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px", boxShadow: "0 4px 12px rgba(0,0,0,0.5)" }}
                        labelClassName="font-bold text-white text-xs"
                      />
                      <Legend verticalAlign="top" height={36} iconType="circle" wrapperStyle={{ fontSize: "11px", fontWeight: 600, color: "#94a3b8" }} />
                      <Area 
                        name="Gross Revenue ($)"
                        type="monotone" 
                        dataKey="Sales" 
                        stroke="#6366f1" 
                        strokeWidth={2.5}
                        fillOpacity={1} 
                        fill="url(#colorSales)" 
                      />
                      <Area 
                        name="Net Profit ($)"
                        type="monotone" 
                        dataKey="Profit" 
                        stroke="#10b981" 
                        strokeWidth={2.5}
                        fillOpacity={1} 
                        fill="url(#colorProfit)" 
                      />
                    </AreaChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 2: Category Pie Chart */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <PieIcon className="w-4 h-4 text-indigo-400" />
                    Share of Sales by Category
                  </h3>
                  <p className="text-[11px] text-slate-500">Revenue split across retail departments. Click any slice to download category transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'category_shares_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Category Share Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-60 w-full flex items-center justify-center">
                {getCategoryChartData().length === 0 ? (
                  <span className="text-xs text-slate-500">No category transactions available.</span>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={getCategoryChartData()}
                        cx="50%"
                        cy="50%"
                        innerRadius={55}
                        outerRadius={75}
                        paddingAngle={5}
                        dataKey="value"
                        onClick={(data: any) => {
                          if (data && data.name) {
                            downloadTransactionsForFilter('category', String(data.name), `${data.name}_category`);
                          }
                        }}
                        className="cursor-pointer"
                      >
                        {getCategoryChartData().map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Pie>
                      <Tooltip 
                        formatter={(val) => [`$${val}`, "Revenue"]}
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px", color: "#f8fafc" }}
                      />
                      <Legend 
                        verticalAlign="bottom" 
                        height={36} 
                        iconType="circle" 
                        layout="horizontal"
                        align="center"
                        wrapperStyle={{ fontSize: "10px", fontWeight: 500, color: "#94a3b8" }}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 3: Top 5 Performing Products */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <BarChart3 className="w-4 h-4 text-emerald-400" />
                    Top 5 Performing Products
                  </h3>
                  <p className="text-[11px] text-slate-500">High performing product SKUs sorted by gross sales. Click any bar to download product transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'top_performing_products_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Product Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-60 w-full">
                {getTopProductsData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500">
                    No matching item rankings
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={getTopProductsData()}
                      layout="vertical"
                      margin={{ top: 10, right: 10, left: 10, bottom: 5 }}
                      onClick={(data: any) => {
                        if (data && data.activePayload && data.activePayload[0]) {
                          const prodName = String(data.activePayload[0].payload.name);
                          downloadTransactionsForFilter('product', prodName, `${prodName}_product`);
                        }
                      }}
                      className="cursor-pointer"
                    >
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#1f2937" />
                      <XAxis type="number" stroke="#4b5563" fontSize={10} tickLine={false} axisLine={false} />
                      <YAxis 
                        dataKey="name" 
                        type="category" 
                        stroke="#94a3b8" 
                        fontSize={9} 
                        tickLine={false} 
                        axisLine={false}
                        width={90}
                        tickFormatter={(v) => v.length > 12 ? `${v.substring(0, 12)}...` : v}
                      />
                      <Tooltip 
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px" }}
                        formatter={(val) => [`$${val}`, "Sales"]}
                      />
                      <Bar dataKey="Sales" fill="#6366f1" radius={[0, 4, 4, 0]} maxBarSize={14}>
                        {getTopProductsData().map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 4: Sales by Region (Bar Chart, descending order) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <MapPin className="w-4 h-4 text-indigo-400" />
                    Sales by Region
                  </h3>
                  <p className="text-[11px] text-slate-500">Regional sales volume sorted in descending order. Click any bar to download regional transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'sales_by_regions_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Regional Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-60 w-full">
                {getRegionSalesData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500">
                    No matching region sales data.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart 
                      data={getRegionSalesData()} 
                      margin={{ top: 10, right: 10, left: -20, bottom: 5 }}
                      onClick={(data: any) => {
                        if (data && data.activeLabel) {
                          downloadTransactionsForFilter('region', String(data.activeLabel), `${data.activeLabel}_region`);
                        }
                      }}
                      className="cursor-pointer"
                    >
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#1f2937" />
                      <XAxis dataKey="name" stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} />
                      <YAxis stroke="#4b5563" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(v) => `$${v}`} />
                      <Tooltip 
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px" }}
                        formatter={(val) => [`$${val}`, "Sales"]}
                      />
                      <Bar dataKey="Sales" fill="#6366f1" radius={[4, 4, 0, 0]} maxBarSize={30}>
                        {getRegionSalesData().map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 5: Category Performance (Pareto Chart: Bar & Line) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <Percent className="w-4 h-4 text-amber-400" />
                    Category Performance
                  </h3>
                  <p className="text-[11px] text-slate-500">Pareto distribution: Sales frequency & cumulative %. Click any bar to download category transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'category_performance_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Category Performance Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-60 w-full">
                {getCategoryParetoData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500">
                    No category data available.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart 
                      data={getCategoryParetoData()} 
                      margin={{ top: 10, right: -5, left: -20, bottom: 5 }}
                      onClick={(data: any) => {
                        if (data && data.activeLabel) {
                          downloadTransactionsForFilter('category', String(data.activeLabel), `${data.activeLabel}_category`);
                        }
                      }}
                      className="cursor-pointer"
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" vertical={false} />
                      <XAxis dataKey="category" stroke="#94a3b8" fontSize={9} tickLine={false} axisLine={false} />
                      <YAxis yAxisId="left" stroke="#4b5563" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(v) => `$${v}`} />
                      <YAxis yAxisId="right" orientation="right" stroke="#ec4899" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(v) => `${v}%`} domain={[0, 100]} />
                      <Tooltip 
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px" }}
                        formatter={(value, name) => {
                          if (name === "Sales") return [`$${value}`, "Sales"];
                          return [`${value}%`, "Cumulative"];
                        }}
                      />
                      <Legend verticalAlign="top" height={24} iconType="circle" wrapperStyle={{ fontSize: "10px" }} />
                      <Bar yAxisId="left" name="Sales" dataKey="Sales" fill="#4f46e5" radius={[4, 4, 0, 0]} maxBarSize={25} />
                      <Line yAxisId="right" name="Cumulative %" type="monotone" dataKey="cumulativePercentage" stroke="#ec4899" strokeWidth={2.5} activeDot={{ r: 4 }} />
                    </ComposedChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 6: Weekly Footfalls Trend (Line Chart) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <TrendingUp className="w-4 h-4 text-emerald-400" />
                    Weekly Footfalls Trend
                  </h3>
                  <p className="text-[11px] text-slate-500">Customer footfall traffic over historical dates. Click any line node to download weekly transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'weekly_footfalls_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Weekly Footfalls Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-60 w-full">
                {getWeeklyFootfallsData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500">
                    No weekly trend data.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart 
                      data={getWeeklyFootfallsData()} 
                      margin={{ top: 10, right: 10, left: -20, bottom: 5 }}
                      onClick={(data: any) => {
                        if (data && data.activeLabel) {
                          downloadTransactionsForFilter('week', String(data.activeLabel), `week_${data.activeLabel}`);
                        }
                      }}
                      className="cursor-pointer"
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" vertical={false} />
                      <XAxis dataKey="week" stroke="#94a3b8" fontSize={9} tickLine={false} axisLine={false} />
                      <YAxis stroke="#4b5563" fontSize={10} tickLine={false} axisLine={false} />
                      <Tooltip 
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px" }}
                      />
                      <Legend verticalAlign="top" height={24} iconType="circle" wrapperStyle={{ fontSize: "10px" }} />
                      <Line name="Footfalls" type="monotone" dataKey="Footfalls" stroke="#10b981" strokeWidth={3} dot={{ r: 3 }} activeDot={{ r: 5 }} />
                    </LineChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 7: Stockout Risks by City (Ascending order, lowest stock highlighted as red) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <AlertCircle className="w-4 h-4 text-rose-400" />
                    Stockout Risks by City
                  </h3>
                  <p className="text-[11px] text-slate-500">Available stocks in ascending order. Click any city bar to download city transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'stockout_risks_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Inventory Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="h-60 w-full">
                {getStockoutRisksData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500">
                    No city inventory data.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart 
                      data={getStockoutRisksData()} 
                      margin={{ top: 10, right: 10, left: -20, bottom: 5 }}
                      onClick={(data: any) => {
                        if (data && data.activeLabel) {
                          downloadTransactionsForFilter('city', String(data.activeLabel), `${data.activeLabel}_city`);
                        }
                      }}
                      className="cursor-pointer"
                    >
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#1f2937" />
                      <XAxis dataKey="city" stroke="#94a3b8" fontSize={9} tickLine={false} axisLine={false} />
                      <YAxis stroke="#4b5563" fontSize={10} tickLine={false} axisLine={false} />
                      <Tooltip 
                        contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px" }}
                        formatter={(val) => [`${val} units`, "Stock Level"]}
                      />
                      <Bar dataKey="StockLevel" radius={[4, 4, 0, 0]} maxBarSize={25}>
                        {getStockoutRisksData().map((entry, index) => {
                          const isTop3Least = index < 3;
                          return (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={isTop3Least ? "#ef4444" : "#475569"} 
                            />
                          );
                        })}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            {/* Chart 8: Marketing Spends Region & City Side-by-Side Pie Charts (Span 2) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md col-span-1 md:col-span-2 flex flex-col justify-between">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <DollarSign className="w-4 h-4 text-indigo-400" />
                    Marketing Expenditures Distribution
                  </h3>
                  <p className="text-[11px] text-slate-500">Parallel channel allocation breakdown across regions and cities. Click any slice to download respective transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'marketing_spends_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Marketing Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-center">
                
                {/* Region Marketing Spend */}
                <div className="flex flex-col items-center">
                  <span className="text-[11px] font-bold text-slate-400 mb-2">Spend by Region (Click Slice)</span>
                  <div className="h-48 w-full flex items-center justify-center">
                    {getMarketingSpendsByRegion().length === 0 ? (
                      <span className="text-xs text-slate-500">No regional spend.</span>
                    ) : (
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={getMarketingSpendsByRegion()}
                            cx="50%"
                            cy="50%"
                            innerRadius={35}
                            outerRadius={55}
                            paddingAngle={4}
                            dataKey="value"
                            onClick={(data: any) => {
                              if (data && data.name) {
                                downloadTransactionsForFilter('region', String(data.name), `${data.name}_region_marketing`);
                              }
                            }}
                            className="cursor-pointer"
                          >
                            {getMarketingSpendsByRegion().map((entry, index) => (
                              <Cell key={`cell-region-${index}`} fill={COLORS[index % COLORS.length]} />
                            ))}
                          </Pie>
                          <Tooltip 
                            formatter={(val) => [`$${val}`, "Spend"]}
                            contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px", color: "#f8fafc" }}
                          />
                          <Legend 
                            verticalAlign="bottom" 
                            height={32} 
                            iconType="circle" 
                            wrapperStyle={{ fontSize: "8px", fontWeight: 500 }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    )}
                  </div>
                </div>

                {/* City Marketing Spend */}
                <div className="flex flex-col items-center">
                  <span className="text-[11px] font-bold text-slate-400 mb-2">Spend by City (Click Slice)</span>
                  <div className="h-48 w-full flex items-center justify-center">
                    {getMarketingSpendsByCity().length === 0 ? (
                      <span className="text-xs text-slate-500">No city spend.</span>
                    ) : (
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={getMarketingSpendsByCity()}
                            cx="50%"
                            cy="50%"
                            innerRadius={35}
                            outerRadius={55}
                            paddingAngle={4}
                            dataKey="value"
                            onClick={(data: any) => {
                              if (data && data.name) {
                                downloadTransactionsForFilter('city', String(data.name), `${data.name}_city_marketing`);
                              }
                            }}
                            className="cursor-pointer"
                          >
                            {getMarketingSpendsByCity().map((entry, index) => (
                              <Cell key={`cell-city-${index}`} fill={COLORS[(index + 2) % COLORS.length]} />
                            ))}
                          </Pie>
                          <Tooltip 
                            formatter={(val) => [`$${val}`, "Spend"]}
                            contentStyle={{ backgroundColor: "#11141b", border: "1px solid #334155", borderRadius: "12px", color: "#f8fafc" }}
                          />
                          <Legend 
                            verticalAlign="bottom" 
                            height={32} 
                            iconType="circle" 
                            wrapperStyle={{ fontSize: "8px", fontWeight: 500 }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    )}
                  </div>
                </div>

              </div>
            </div>

            {/* Chart 9: Average Customer Ratings Funnel Chart (Span 2) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md col-span-1 md:col-span-2 flex flex-col gap-4">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <Star className="w-4 h-4 text-yellow-400 fill-yellow-400" />
                    Average Customer Ratings
                  </h3>
                  <p className="text-[11px] text-slate-500">Funnel distribution of satisfaction scoring across product categories. Click any row or segment to download its transactions.</p>
                </div>
                <button
                  onClick={() => downloadTransactionsForFilter('all', '', 'customer_ratings_all')}
                  className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                  title="Download All Customer Ratings Data"
                >
                  <Download className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="min-h-[220px] w-full flex items-center justify-center">
                {getCategoryCustomerRatingsData().length === 0 ? (
                  <div className="h-full flex items-center justify-center text-xs text-slate-500 py-12">
                    No ratings data available.
                  </div>
                ) : (
                  <div className="w-full flex flex-col md:flex-row items-center gap-8 justify-center py-2">
                    {/* Visual Funnel drawing */}
                    <div className="relative w-full md:w-1/2 h-full flex justify-center items-center">
                      <svg viewBox="0 0 500 210" className="w-full h-full max-h-[210px]" preserveAspectRatio="xMidYMid meet">
                        {getCategoryCustomerRatingsData().map((item, i, arr) => {
                          const rowCount = arr.length;
                          const rowHeight = 30;
                          const gap = 10;
                          
                          // Width ranges from 460 to 180
                          const startW = 460 - (i * (460 - 180)) / rowCount;
                          const endW = 460 - ((i + 1) * (460 - 180)) / rowCount;
                          
                          const startX1 = (500 - startW) / 2;
                          const startX2 = startX1 + startW;
                          const endX1 = (500 - endW) / 2;
                          const endX2 = endX1 + endW;
                          
                          const y1 = i * (rowHeight + gap);
                          const y2 = y1 + rowHeight;
                          
                          const points = `${startX1},${y1} ${startX2},${y1} ${endX2},${y2} ${endX1},${y2}`;
                          const color = COLORS[i % COLORS.length];
                          
                          return (
                            <g 
                              key={item.category} 
                              className="group cursor-pointer animate-pulse-subtle"
                              onClick={() => downloadTransactionsForFilter('category', item.category, `${item.category}_ratings_transactions`)}
                            >
                              {/* Highlight backdrop */}
                              <polygon 
                                points={points} 
                                fill={color} 
                                opacity={0.15}
                                className="transition-all duration-300 group-hover:opacity-30"
                              />
                              {/* Funnel Segment */}
                              <polygon 
                                points={points} 
                                fill={color} 
                                opacity={0.75}
                                className="transition-all duration-300 group-hover:opacity-90 filter group-hover:brightness-110"
                              />
                              
                              {/* Category Title inside Funnel Segment */}
                              <text 
                                x="250" 
                                y={y1 + rowHeight / 2 + 4} 
                                textAnchor="middle" 
                                fill="#ffffff" 
                                fontSize="11" 
                                fontWeight="bold"
                                className="pointer-events-none select-none drop-shadow-md tracking-wide"
                              >
                                {item.category.toUpperCase()}
                              </text>
                            </g>
                          );
                        })}
                      </svg>
                    </div>

                    {/* Funnel ratings statistics side panel */}
                    <div className="w-full md:w-1/2 flex flex-col gap-2">
                      {getCategoryCustomerRatingsData().map((item, i) => {
                        const color = COLORS[i % COLORS.length];
                        return (
                          <div 
                            key={item.category} 
                            onClick={() => downloadTransactionsForFilter('category', item.category, `${item.category}_ratings_transactions`)}
                            className="flex items-center justify-between text-xs p-2.5 rounded-xl bg-[#11141b]/50 border border-slate-800 hover:bg-[#11141b]/80 hover:border-slate-700/60 transition-all cursor-pointer hover:border-indigo-500/40"
                          >
                            <div className="flex items-center gap-2.5">
                              <span className="w-2.5 h-2.5 rounded-full ring-2 ring-slate-800" style={{ backgroundColor: color }} />
                              <span className="font-semibold text-slate-300">{item.category}</span>
                            </div>
                            <div className="flex items-center gap-2 font-mono">
                              <span className="text-yellow-400 font-bold flex items-center gap-0.5">
                                {item.avgRating.toFixed(2)}
                                <Star className="w-3.5 h-3.5 fill-yellow-400 text-yellow-400 inline" />
                              </span>
                              <span className="text-slate-500 text-[10px]">({item.count} orders)</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* Table: Store Performance (Span 2) */}
            <div className="bg-[#161a23] rounded-2xl border border-slate-800 p-6 shadow-md col-span-1 md:col-span-2 flex flex-col gap-4">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                    <ShoppingBag className="w-4 h-4 text-indigo-400" />
                    Store Performance
                  </h3>
                  <p className="text-[11px] text-slate-500">Summary metrics detailing customer traffic, units shipped, and top-line sales per store. Click any row to download its transactions.</p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => downloadTransactionsForFilter('all', '', 'store_performance_all')}
                    className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-800 hover:border-slate-700"
                    title="Download All Stores Data"
                  >
                    <Download className="w-3.5 h-3.5" />
                  </button>

                  {/* Top N controls displaying multiples of 5 */}
                  <div className="relative">
                    <button 
                      onClick={() => setShowStoreDropdown(!showStoreDropdown)} 
                      className="bg-slate-800/80 hover:bg-slate-700 text-slate-200 text-[11px] px-3 py-1.5 rounded-lg border border-slate-700 font-semibold flex items-center gap-1.5 transition-all cursor-pointer focus:outline-none"
                    >
                      <span>View: {storeTopN === "All" ? "All Stores" : `Top ${storeTopN}`}</span>
                      <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
                    </button>
                    {showStoreDropdown && (
                      <>
                        {/* Invisible backdrop to dismiss dropdown */}
                        <div className="fixed inset-0 z-20" onClick={() => setShowStoreDropdown(false)} />
                        <div className="absolute right-0 mt-1.5 w-32 bg-[#11141b] border border-slate-800 rounded-lg shadow-xl z-30 py-1 text-xs">
                          {[5, 10, 15, 20, "All"].map((num) => (
                            <button
                              key={num}
                              onClick={() => {
                                setStoreTopN(num as any);
                                setShowStoreDropdown(false);
                              }}
                              className={`w-full text-left px-3 py-2 hover:bg-indigo-950/40 hover:text-indigo-300 transition-colors cursor-pointer ${
                                storeTopN === num ? "text-indigo-400 font-bold bg-indigo-950/20" : "text-slate-400"
                              }`}
                            >
                              {num === "All" ? "Show All" : `${num} Stores`}
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-slate-800/80 text-[10px] text-slate-400 uppercase font-bold tracking-wider">
                      <th className="py-3 px-4">Store Name</th>
                      <th className="py-3 px-4 text-right">Customer Footfalls</th>
                      <th className="py-3 px-4 text-right">Units Sold</th>
                      <th className="py-3 px-4 text-right">Net Sales ($)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {getStorePerformanceData().length === 0 ? (
                      <tr>
                        <td colSpan={4} className="py-8 text-center text-xs text-slate-500 italic">
                          No store information available under current filters.
                        </td>
                      </tr>
                    ) : (
                      getStorePerformanceData().map((st) => (
                        <tr 
                          key={st.store} 
                          onClick={() => downloadTransactionsForFilter('store', st.store, `${st.store}_store`)}
                          className="border-b border-slate-800/40 hover:bg-[#1d222f]/40 transition-colors text-xs cursor-pointer hover:text-indigo-300"
                        >
                          <td className="py-3 px-4 font-semibold text-slate-200">
                            {st.store}
                          </td>
                          <td className="py-3 px-4 text-right font-mono text-slate-300">
                            {st.footfalls.toLocaleString()}
                          </td>
                          <td className="py-3 px-4 text-right font-mono text-slate-300">
                            {st.units_sold.toLocaleString()}
                          </td>
                          <td className="py-3 px-4 text-right font-mono font-bold text-emerald-400">
                            ${st.net_sales.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
          </>
          )}
        </section>
      </main>

      {/* Sensitive Data Blocked Warning Modal */}
      <AnimatePresence>
        {showSensitiveModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/85 backdrop-blur-md z-[9999] flex items-center justify-center p-4"
          >
            <motion.div
              initial={{ scale: 0.95, y: 15 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 15 }}
              className="bg-[#161a23] border border-red-500/30 rounded-2xl p-6 max-w-md w-full shadow-2xl shadow-black/80 flex flex-col items-center text-center gap-4 relative overflow-hidden"
            >
              {/* Decorative top red gradient bar */}
              <div className="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r from-red-600 via-amber-500 to-red-600" />

              <div className="w-14 h-14 bg-red-500/10 border border-red-500/20 text-red-400 rounded-full flex items-center justify-center mt-2 animate-bounce">
                <AlertCircle className="w-7 h-7" />
              </div>

              <div className="space-y-1">
                <h3 className="text-lg font-bold text-white tracking-tight font-display">
                  Sensitive Information Detected
                </h3>
                <p className="text-xs text-red-400 font-semibold tracking-wide uppercase">
                  Upload Terminated
                </p>
              </div>

              <div className="bg-[#11141b] border border-slate-800/80 rounded-xl p-4 w-full text-left space-y-3">
                <p className="text-xs text-slate-300 font-medium leading-relaxed">
                  No sensitive information is permitted to be uploaded.
                </p>
                {sensitiveInfoWarning && (
                  <div className="text-[11px] text-slate-400 space-y-1.5 border-t border-slate-800/60 pt-2.5">
                    <div>
                      <span className="font-bold text-slate-500 uppercase tracking-wider block text-[9px]">Target File:</span>
                      <span className="font-mono text-indigo-400 font-semibold">{sensitiveInfoWarning.fileName}</span>
                    </div>
                    <div>
                      <span className="font-bold text-slate-500 uppercase tracking-wider block text-[9px]">Reason:</span>
                      <span className="font-medium text-slate-200">{sensitiveInfoWarning.reason}</span>
                    </div>
                  </div>
                )}
              </div>

              <p className="text-[11px] text-slate-400 leading-relaxed px-1">
                For security and privacy reasons, please ensure uploaded spreadsheets do not contain passwords, bank account passwords, bank card numbers, CVVs, pins, or routing numbers. Please sanitize your data and try again.
              </p>

              <button
                onClick={() => {
                  setShowSensitiveModal(false);
                  setSensitiveInfoWarning(null);
                }}
                className="w-full bg-red-600 hover:bg-red-500 text-white font-bold py-2.5 px-4 rounded-xl text-xs transition-all duration-200 cursor-pointer shadow-lg shadow-red-950/20 uppercase tracking-wider border border-red-500/20"
              >
                Acknowledge & Dismiss
              </button>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
