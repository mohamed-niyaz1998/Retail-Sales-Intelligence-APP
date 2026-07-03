import * as XLSX from "xlsx";

export interface SalesRecord {
  id: string;
  date: string; // YYYY-MM-DD
  week: string; // e.g. "2026-W18"
  product: string;
  category: string;
  quantity: number;
  sales: number;
  profit: number;
  region: string;
  city: string;
  store: string;
  storeFormat: string;
  discountRate: number; // Decimal (0.0 to 1.0)
  returnAmt: number; // Return refund amount
}

export interface ColumnMapping {
  date: string;
  product: string;
  category: string;
  quantity: string;
  sales: string;
  profit: string;
  region: string;
  discountRate: string;
  returnAmt: string;
  week: string;
  city: string;
  store: string;
  storeFormat: string;
}

export function getISOWeekString(dateStr: string): string {
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return "N/A";
  
  const date = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const day = date.getUTCDay(); // 0 is Sunday, 1 is Monday...
  const diff = date.getUTCDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), diff));
  
  const dd = String(monday.getUTCDate()).padStart(2, "0");
  const mm = String(monday.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = monday.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// Generate highly realistic mock data for default dashboard visualization
export function getMockSalesData(): SalesRecord[] {
  const categories = ["Electronics", "Apparel & Fashion", "Home & Kitchen", "Beauty & Care", "Sports & Outdoors"];
  
  const products: Record<string, string[]> = {
    "Electronics": ["Quantum Wireless Headphones", "Apex Mechanical Keyboard", "UltraView 27\" Monitor", "Nomad Power Bank 20k", "AeroFit Smart Watch"],
    "Apparel & Fashion": ["Classic Comfort Hoodie", "Slim-Fit Denim Jeans", "All-Weather Active Jacket", "Breathable Mesh Sneakers", "Elegance Silk Scarf"],
    "Home & Kitchen": ["Culinary Pro Chef Knife", "Smart Brew Espresso Maker", "PureAir HEPA Purifier", "Stoneware Dinner Set (16pc)", "ComfortCloud Memory Pillow"],
    "Beauty & Care": ["HydraGlow Hyaluronic Serum", "Botanical Therapy Shampoo", "Mineral Sunscreen SPF 50+", "Velvet Matte Lipstick Kit", "Sonic Cleansing Brush"],
    "Sports & Outdoors": ["TrailBlazer Ergonomic Backpack", "HydroSip Vacuum Flask", "Carbon Fiber Trekking Poles", "Pro-Grip Yoga Mat", "UltraLight Sleeping Pad"]
  };

  const regions = ["East Coast", "West Coast", "Midwest", "South", "Pacific Northwest"];
  const data: SalesRecord[] = [];
  
  // Create records spread across the last 6 months
  const baseDate = new Date();
  baseDate.setMonth(baseDate.getMonth() - 6);

  let currentId = 1;
  for (let i = 0; i < 240; i++) {
    // Generate dates sequentially with some randomness
    const dateOffset = Math.floor(i * 0.75);
    const orderDate = new Date(baseDate.getTime());
    orderDate.setDate(orderDate.getDate() + dateOffset);
    
    // Prevent future dates
    if (orderDate > new Date()) continue;

    const formattedDate = orderDate.toISOString().split("T")[0];
    const category = categories[Math.floor(Math.random() * categories.length)];
    const categoryProducts = products[category];
    const product = categoryProducts[Math.floor(Math.random() * categoryProducts.length)];
    const region = regions[Math.floor(Math.random() * regions.length)];
    
    // Determine city based on region
    const regionToCities: Record<string, string[]> = {
      "East Coast": ["New York", "Boston", "Atlanta"],
      "West Coast": ["Los Angeles", "San Francisco", "San Diego"],
      "Midwest": ["Chicago", "Detroit", "Minneapolis"],
      "South": ["Houston", "Dallas", "Miami"],
      "Pacific Northwest": ["Seattle", "Portland", "Bellevue"]
    };
    const citiesForRegion = regionToCities[region] || ["Other City"];
    const city = citiesForRegion[Math.floor(Math.random() * citiesForRegion.length)];

    // Determine store and format
    const storeNum = 101 + (i % 4);
    const store = `${city} Store #${storeNum}`;
    const storeFormats = ["Supercenter", "Express", "Hypermarket", "Neighborhood Market"];
    const storeFormat = storeFormats[i % storeFormats.length];

    const week = getISOWeekString(formattedDate);
    
    // Quantity: 1 to 12 items
    const quantity = Math.floor(Math.random() * 8) + 1;
    
    // Standard sales prices
    let unitPrice = 19.99;
    if (category === "Electronics") unitPrice = 79.99 + (Math.random() * 150);
    else if (category === "Apparel & Fashion") unitPrice = 29.99 + (Math.random() * 60);
    else if (category === "Home & Kitchen") unitPrice = 24.99 + (Math.random() * 180);
    else if (category === "Beauty & Care") unitPrice = 15.99 + (Math.random() * 40);
    else if (category === "Sports & Outdoors") unitPrice = 12.99 + (Math.random() * 90);
    
    unitPrice = Math.round(unitPrice * 100) / 100;
    const sales = Math.round(unitPrice * quantity * 100) / 100;
    
    // Profit margin between 15% and 65% depending on category
    let margin = 0.35;
    if (category === "Electronics") margin = 0.22 + (Math.random() * 0.15);
    else if (category === "Beauty & Care") margin = 0.45 + (Math.random() * 0.20);
    else if (category === "Apparel & Fashion") margin = 0.35 + (Math.random() * 0.15);
    else if (category === "Home & Kitchen") margin = 0.28 + (Math.random() * 0.18);
    else if (category === "Sports & Outdoors") margin = 0.25 + (Math.random() * 0.20);
    
    const profit = Math.round(sales * margin * 100) / 100;
    
    // Simulate discount rate and return amount
    const discountRate = Math.random() < 0.35 ? [0.05, 0.10, 0.15, 0.20][Math.floor(Math.random() * 4)] : 0;
    const isReturned = Math.random() < 0.06;
    const returnAmt = isReturned ? Math.round(sales * (0.4 + Math.random() * 0.6) * 100) / 100 : 0;

    data.push({
      id: `DEMO-${String(currentId++).padStart(4, "0")}`,
      date: formattedDate,
      week,
      product,
      category,
      quantity,
      sales,
      profit,
      region,
      city,
      store,
      storeFormat,
      discountRate,
      returnAmt
    });
  }

  // Sort by date ascending
  return data.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
}

// Map spreadsheet columns to standard fields
export function detectColumnMapping(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {
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
  };

  const clean = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");

  headers.forEach((h) => {
    const ch = clean(h);
    
    // Date
    if (!mapping.date && (ch.includes("date") || ch.includes("time") || ch === "dt" || ch.includes("orderday"))) {
      mapping.date = h;
    }
    // Sales / Revenue
    else if (!mapping.sales && (ch.includes("sales") || ch.includes("revenue") || ch.includes("amount") || ch.includes("turnover") || ch.includes("totalrev") || ch === "salesamount" || ch === "price" || ch === "total")) {
      mapping.sales = h;
    }
    // Product
    else if (!mapping.product && (ch.includes("product") || ch.includes("item") || ch.includes("sku") || ch.includes("desc") || ch === "name" || ch === "prod")) {
      mapping.product = h;
    }
    // Category
    else if (!mapping.category && (ch.includes("category") || ch.includes("dept") || ch.includes("department") || ch.includes("class") || ch.includes("segment") || ch.includes("group"))) {
      mapping.category = h;
    }
    // Quantity
    else if (!mapping.quantity && (ch.includes("qty") || ch.includes("quantity") || ch.includes("units") || ch.includes("count") || ch.includes("volume") || ch.includes("sold"))) {
      mapping.quantity = h;
    }
    // Profit
    else if (!mapping.profit && (ch.includes("profit") || ch.includes("margin") || ch.includes("earnings") || ch.includes("net") || ch.includes("markup") || ch === "income")) {
      mapping.profit = h;
    }
    // Region
    else if (!mapping.region && (ch.includes("region") || ch.includes("store") || ch.includes("loc") || ch.includes("city") || ch.includes("state") || ch.includes("branch") || ch.includes("country"))) {
      mapping.region = h;
    }
    // Discount Rate
    else if (!mapping.discountRate && (ch.includes("discount") || ch.includes("disc") || ch.includes("promo") || ch.includes("reduction") || ch.includes("rate"))) {
      mapping.discountRate = h;
    }
    // Return Amount
    else if (!mapping.returnAmt && (ch.includes("return") || ch.includes("refund") || ch.includes("returned") || ch.includes("chargeback") || ch.includes("retamt"))) {
      mapping.returnAmt = h;
    }
    // Week
    else if (!mapping.week && (ch.includes("week") || ch === "wk" || ch === "invoice_week")) {
      mapping.week = h;
    }
    // City
    else if (!mapping.city && (ch.includes("city") || ch.includes("town") || ch.includes("municipality") || ch.includes("metro"))) {
      mapping.city = h;
    }
    // Store
    else if (!mapping.store && ((ch.includes("store") && !ch.includes("format") && !ch.includes("region")) || ch === "shop" || ch === "outlet" || ch === "branch")) {
      mapping.store = h;
    }
    // Store Format
    else if (!mapping.storeFormat && (ch.includes("format") || ch.includes("channel") || ch.includes("storetype") || ch.includes("store_format"))) {
      mapping.storeFormat = h;
    }
  });

  // Fallback defaults if no direct match found
  if (!mapping.date) mapping.date = headers.find(h => clean(h).includes("date")) || headers[0] || "";
  if (!mapping.sales) mapping.sales = headers.find(h => clean(h).match(/rev|sale|amount/i)) || headers[1] || "";
  if (!mapping.product) mapping.product = headers.find(h => clean(h).match(/prod|item|sku|name/i)) || headers[2] || "";
  if (!mapping.category) mapping.category = headers.find(h => clean(h).match(/cat|dept/i)) || headers[3] || "";
  if (!mapping.quantity) mapping.quantity = headers.find(h => clean(h).match(/qty|quant|unit/i)) || headers[4] || "";
  if (!mapping.profit) mapping.profit = headers.find(h => clean(h).match(/prof|margin|earn/i)) || headers[5] || "";
  if (!mapping.region) mapping.region = headers.find(h => clean(h).match(/reg|store|loc|city/i)) || headers[6] || "";
  if (!mapping.discountRate) mapping.discountRate = headers.find(h => clean(h).match(/disc|promo/i)) || "";
  if (!mapping.returnAmt) mapping.returnAmt = headers.find(h => clean(h).match(/ret|ref/i)) || "";
  if (!mapping.week) mapping.week = headers.find(h => clean(h).match(/wk|week/i)) || "";
  if (!mapping.city) mapping.city = headers.find(h => clean(h).match(/city|town/i)) || "";
  if (!mapping.store) mapping.store = headers.find(h => clean(h).match(/store|outlet|shop|branch/i) && !clean(h).match(/format|reg/i)) || "";
  if (!mapping.storeFormat) mapping.storeFormat = headers.find(h => clean(h).match(/format|type|channel/i)) || "";

  return mapping;
}

// Parses raw sheet data to standard SalesRecord format using the provided mapping
export function parseSheetData(rawData: any[], mapping: ColumnMapping, fileIndex: number): SalesRecord[] {
  return rawData.map((row, idx) => {
    const rawDate = row[mapping.date];
    let dateStr = new Date().toISOString().split("T")[0];
    
    if (rawDate) {
      if (typeof rawDate === "number" && rawDate > 20000 && rawDate < 60000) {
        // Excel serial date format
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        const jsDate = new Date(excelEpoch.getTime() + rawDate * 24 * 60 * 60 * 1000);
        dateStr = jsDate.toISOString().split("T")[0];
      } else {
        const parsed = new Date(rawDate);
        if (!isNaN(parsed.getTime())) {
          dateStr = parsed.toISOString().split("T")[0];
        }
      }
    }

    const salesVal = parseFloat(row[mapping.sales]);
    const sales = isNaN(salesVal) ? 0 : Math.round(salesVal * 100) / 100;

    const qtyVal = parseInt(row[mapping.quantity]);
    const quantity = isNaN(qtyVal) ? 1 : qtyVal;

    const profitVal = parseFloat(row[mapping.profit]);
    // If profit column isn't mapped or missing, assume 30% margin as estimation
    const profit = isNaN(profitVal) ? Math.round(sales * 0.3 * 100) / 100 : Math.round(profitVal * 100) / 100;

    // Parse or generate discount rate
    const discountVal = parseFloat(row[mapping.discountRate]);
    const discountRate = !isNaN(discountVal) 
      ? (discountVal > 1 ? discountVal / 100 : discountVal) 
      : (((idx * 7) % 10 < 3) ? (((idx * 5) % 4 + 1) * 0.05) : 0);

    // Parse or generate return amount
    const returnVal = parseFloat(row[mapping.returnAmt]);
    const returnAmt = !isNaN(returnVal) 
      ? Math.round(returnVal * 100) / 100 
      : (((idx * 13) % 20 === 0) ? Math.round(sales * (0.3 + ((idx % 5) * 0.15)) * 100) / 100 : 0);

    const region = String(row[mapping.region] || "HQ Store");

    // Get week from column if mapped, or calculate from dateStr
    const week = mapping.week && row[mapping.week] 
      ? String(row[mapping.week]) 
      : getISOWeekString(dateStr);

    // Get city from column if mapped, or simulate logically based on region
    let city = "New York";
    if (mapping.city && row[mapping.city]) {
      city = String(row[mapping.city]);
    } else {
      const regionToCities: Record<string, string[]> = {
        "East Coast": ["New York", "Boston", "Atlanta"],
        "West Coast": ["Los Angeles", "San Francisco", "San Diego"],
        "Midwest": ["Chicago", "Detroit", "Minneapolis"],
        "South": ["Houston", "Dallas", "Miami"],
        "Pacific Northwest": ["Seattle", "Portland", "Bellevue"]
      };
      const citiesForRegion = regionToCities[region] || ["Other City"];
      city = citiesForRegion[idx % citiesForRegion.length];
    }

    // Get store from column if mapped, or simulate
    let store = `${city} Store #101`;
    if (mapping.store && row[mapping.store]) {
      store = String(row[mapping.store]);
    } else {
      const storeNum = 101 + (idx % 4);
      store = `${city} Store #${storeNum}`;
    }

    // Get storeFormat from column if mapped, or simulate
    let storeFormat = "Supercenter";
    if (mapping.storeFormat && row[mapping.storeFormat]) {
      storeFormat = String(row[mapping.storeFormat]);
    } else {
      const storeFormats = ["Supercenter", "Express", "Hypermarket", "Neighborhood Market"];
      storeFormat = storeFormats[idx % storeFormats.length];
    }

    return {
      id: `FILE-${fileIndex}-${idx + 1}`,
      date: dateStr,
      week,
      product: String(row[mapping.product] || "Unlabeled Product"),
      category: String(row[mapping.category] || "General"),
      quantity,
      sales,
      profit,
      region,
      city,
      store,
      storeFormat,
      discountRate,
      returnAmt
    };
  });
}

// Generates a mock retail sales Excel file for users to download
export function downloadSampleExcel() {
  const products = [
    { name: "Quantum Wireless Headphones", category: "Electronics", price: 149.99, cost: 75.00 },
    { name: "Apex Mechanical Keyboard", category: "Electronics", price: 129.99, cost: 60.00 },
    { name: "UltraView 27\" Monitor", category: "Electronics", price: 299.99, cost: 180.00 },
    { name: "Nomad Power Bank 20k", category: "Electronics", price: 49.99, cost: 20.00 },
    { name: "Classic Comfort Hoodie", category: "Apparel & Fashion", price: 59.99, cost: 22.00 },
    { name: "Slim-Fit Denim Jeans", category: "Apparel & Fashion", price: 69.99, cost: 25.00 },
    { name: "All-Weather Active Jacket", category: "Apparel & Fashion", price: 99.99, cost: 40.00 },
    { name: "Smart Brew Espresso Maker", category: "Home & Kitchen", price: 199.99, cost: 110.00 },
    { name: "Culinary Pro Chef Knife", category: "Home & Kitchen", price: 89.99, cost: 35.00 },
    { name: "HydraGlow Hyaluronic Serum", category: "Beauty & Care", price: 34.99, cost: 8.00 },
    { name: "TrailBlazer Ergonomic Backpack", category: "Sports & Outdoors", price: 79.99, cost: 32.00 },
    { name: "HydroSip Vacuum Flask", category: "Sports & Outdoors", price: 29.99, cost: 10.00 }
  ];

  const regions = ["East Coast", "West Coast", "Midwest", "South", "Pacific Northwest"];
  const rows = [];
  
  // Set dates in the first half of 2026
  const baseDate = new Date("2026-01-01");

  for (let i = 0; i < 150; i++) {
    const randomOffsetDays = Math.floor(Math.random() * 160);
    const orderDate = new Date(baseDate.getTime());
    orderDate.setDate(orderDate.getDate() + randomOffsetDays);
    
    const prod = products[Math.floor(Math.random() * products.length)];
    const qty = Math.floor(Math.random() * 5) + 1;
    const region = regions[Math.floor(Math.random() * regions.length)];
    
    const revenue = Math.round(prod.price * qty * 100) / 100;
    const profit = Math.round((revenue - (prod.cost * qty)) * 100) / 100;

    const discountRate = Math.random() < 0.35 ? [0.05, 0.10, 0.15, 0.20][Math.floor(Math.random() * 4)] : 0;
    const isReturned = Math.random() < 0.06;
    const returnAmt = isReturned ? Math.round(revenue * (0.4 + Math.random() * 0.6) * 100) / 100 : 0;

    rows.push({
      "Invoice Date": orderDate.toISOString().split("T")[0],
      "Product Name": prod.name,
      "Category": prod.category,
      "Units Sold": qty,
      "Revenue": revenue,
      "Net Profit": profit,
      "Store Region": region,
      "Discount Rate": discountRate,
      "Return Amount": returnAmt
    });
  }

  // Sort rows chronologically
  rows.sort((a, b) => new Date(a["Invoice Date"]).getTime() - new Date(b["Invoice Date"]).getTime());

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Transactions");
  
  // Write and trigger download
  XLSX.writeFile(workbook, "Retail_Sales_Data_Sample.xlsx");
}
