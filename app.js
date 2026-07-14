/* ==========================================================================
   TARA APP - CORE APPLICATION LOGIC
   POS, Stock Management, Accounting, and Google Sheets Sync Engine
   ========================================================================== */

// --- GOOGLE APPS SCRIPT CODE TEMPLATE ---
// This code is displayed in the Settings panel for the user to copy-paste.
const GOOGLE_APPS_SCRIPT_CODE = `function doGet(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetsExist(spreadsheet);
  
  var result = {
    products: getSheetData(spreadsheet.getSheetByName("Products")),
    sales: getSheetData(spreadsheet.getSheetByName("Sales")),
    expenses: getSheetData(spreadsheet.getSheetByName("Expenses"))
  };
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetsExist(spreadsheet);
  
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    
    if (action === "test") {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "เชื่อมต่อสำเร็จ!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "push") {
      saveDataToSheet(spreadsheet.getSheetByName("Products"), requestData.data.products, ["id", "name", "category", "price", "cost", "stock", "minStock", "imageCode", "image"]);
      saveDataToSheet(spreadsheet.getSheetByName("Sales"), requestData.data.sales, ["id", "timestamp", "items", "subtotal", "discount", "total", "cashReceived", "change"]);
      saveDataToSheet(spreadsheet.getSheetByName("Expenses"), requestData.data.expenses, ["id", "timestamp", "title", "category", "amount", "date"]);
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "ซิงก์ข้อมูลขึ้นคลาวด์สำเร็จ!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ไม่พบ Action ที่ระบุ" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function ensureSheetsExist(spreadsheet) {
  var sheets = {
    "Products": ["id", "name", "category", "price", "cost", "stock", "minStock", "imageCode", "image"],
    "Sales": ["id", "timestamp", "items", "subtotal", "discount", "total", "cashReceived", "change"],
    "Expenses": ["id", "timestamp", "title", "category", "amount", "date"]
  };
  
  for (var name in sheets) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
      sheet.appendRow(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name].length).setFontWeight("bold").setBackground("#e2e8f0");
    }
  }
}

function getSheetData(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      if (headers[j] === "items" && typeof val === "string") {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[headers[j]] = val;
    }
    data.push(obj);
  }
  return data;
}

function saveDataToSheet(sheet, dataList, fields) {
  sheet.clear();
  sheet.appendRow(fields);
  sheet.getRange(1, 1, 1, fields.length).setFontWeight("bold").setBackground("#e2e8f0");
  
  if (!dataList || dataList.length === 0) return;
  
  var rows = [];
  for (var i = 0; i < dataList.length; i++) {
    var item = dataList[i];
    var row = [];
    for (var j = 0; j < fields.length; j++) {
      var val = item[fields[j]];
      if (typeof val === "object") {
        val = JSON.stringify(val);
      }
      row.push(val);
    }
    rows.push(row);
  }
  sheet.getRange(2, 1, rows.length, fields.length).setValues(rows);
}`;

// --- MOCK DATA ---
const DEMO_PRODUCTS = [
    { id: "P-1001", name: "น้ำดื่มคริสตัล 1.5L (แพ็ก 6)", category: "water", price: 60.00, cost: 42.00, stock: 45, minStock: 10, imageCode: "water-pack", image: null },
    { id: "P-1002", name: "น้ำดื่มตราสิงห์ 600ml (แพ็ก 12)", category: "water", price: 65.00, cost: 46.00, stock: 35, minStock: 12, imageCode: "water-pack", image: null },
    { id: "P-1003", name: "โค้ก ออริจินัล 325ml (กระป๋อง)", category: "soda", price: 15.00, cost: 11.00, stock: 120, minStock: 25, imageCode: "soda-can", image: null },
    { id: "P-1004", name: "เป๊ปซี่ 1.45L (ขวด)", category: "soda", price: 32.00, cost: 23.50, stock: 8, minStock: 10, imageCode: "soda-bottle", image: null },
    { id: "P-1005", name: "น้ำส้มดีโด้ 250ml (ขวดแก้ว)", category: "juice", price: 12.00, cost: 8.00, stock: 65, minStock: 15, imageCode: "soda-bottle", image: null },
    { id: "P-1006", name: "โออิชิ กรีนที ต้นตำรับ 500ml", category: "juice", price: 20.00, cost: 14.00, stock: 40, minStock: 15, imageCode: "soda-bottle", image: null },
    { id: "P-1007", name: "กระทิงแดง 150ml (ขวด)", category: "energy", price: 10.00, cost: 7.20, stock: 80, minStock: 20, imageCode: "energy-bottle", image: null },
    { id: "P-1008", name: "ถังน้ำใส 18.9L (ถังใหญ่ถังเติม)", category: "other", price: 120.00, cost: 75.00, stock: 3, minStock: 5, imageCode: "other-beverage", image: null }
];

const DEMO_EXPENSES = [
    { id: "E-1001", timestamp: Date.now() - 24*60*60*1000*2, title: "ซื้อน้ำแข็งยูนิค 5 กระสอบ", category: "raw_materials", amount: 150.00, date: new Date(Date.now() - 24*60*60*1000*2).toISOString().split('T')[0] },
    { id: "E-1002", timestamp: Date.now() - 24*60*60*1000*1, title: "ค่าขนส่งสินค้าเข้าร้าน", category: "transport", amount: 200.00, date: new Date(Date.now() - 24*60*60*1000*1).toISOString().split('T')[0] }
];

// --- APP STATE ---
const TaraApp = {
    products: [],
    sales: [],
    expenses: [],
    cart: [],
    settings: {
        googleSheetUrl: ""
    },
    
    // Load from LocalStorage
    loadState() {
        try {
            const products = localStorage.getItem("tara_products");
            const sales = localStorage.getItem("tara_sales");
            const expenses = localStorage.getItem("tara_expenses");
            const settings = localStorage.getItem("tara_settings");
            
            this.sales = sales ? JSON.parse(sales) : [];
            this.expenses = expenses ? JSON.parse(expenses) : [];
            this.settings = settings ? JSON.parse(settings) : { googleSheetUrl: "" };
            
            if (products === null) {
                // First time opening the app, auto load demo
                this.loadDemoData();
            } else {
                this.products = JSON.parse(products);
            }
        } catch (e) {
            console.error("Error loading state", e);
            this.showToast("เกิดข้อผิดพลาดในการโหลดข้อมูลเก่า", "error");
        }
    },
    
    // Save to LocalStorage
    saveState() {
        localStorage.setItem("tara_products", JSON.stringify(this.products));
        localStorage.setItem("tara_sales", JSON.stringify(this.sales));
        localStorage.setItem("tara_expenses", JSON.stringify(this.expenses));
        localStorage.setItem("tara_settings", JSON.stringify(this.settings));
    },
    
    loadDemoData() {
        this.products = JSON.parse(JSON.stringify(DEMO_PRODUCTS));
        // Add fake sales over the last 7 days for visual graphs
        this.sales = [];
        const now = Date.now();
        for (let i = 6; i >= 0; i--) {
            const dayTime = now - i * 24 * 60 * 60 * 1000;
            // random sales volume
            const orderCount = Math.floor(Math.random() * 6) + 3; // 3 to 8 sales
            for (let j = 0; j < orderCount; j++) {
                const item = this.products[Math.floor(Math.random() * this.products.length)];
                const qty = Math.floor(Math.random() * 3) + 1;
                const subtotal = item.price * qty;
                const discount = Math.random() > 0.7 ? 10 : 0;
                const total = subtotal - discount;
                
                this.sales.push({
                    id: "TR-" + (100000 + Math.floor(Math.random() * 900000)),
                    timestamp: dayTime - j * 30 * 60 * 1000, // staggered time
                    items: [{ id: item.id, name: item.name, qty: qty, price: item.price, cost: item.cost || (item.price * 0.7) }],
                    subtotal: subtotal,
                    discount: discount,
                    total: total,
                    cashReceived: total + (Math.random() > 0.5 ? 50 : 0),
                    change: (Math.random() > 0.5 ? 50 : 0)
                });
            }
        }
        
        this.expenses = JSON.parse(JSON.stringify(DEMO_EXPENSES));
        this.saveState();
        this.showToast("โหลดข้อมูลจำลองเริ่มต้นเรียบร้อยแล้ว!", "success");
    },
    
    resetData() {
        this.products = [];
        this.sales = [];
        this.expenses = [];
        this.cart = [];
        this.saveState();
        this.showToast("ล้างข้อมูลเรียบร้อยแล้ว", "success");
    },
    
    // Toast notification
    showToast(message, type = "success") {
        const container = document.getElementById("toast-container");
        const toast = document.createElement("div");
        toast.className = `toast toast-${type}`;
        
        toast.innerHTML = `
            <span class="toast-message">${message}</span>
            <button class="toast-close">&times;</button>
        `;
        
        container.appendChild(toast);
        
        // Remove toast on click
        toast.querySelector(".toast-close").addEventListener("click", () => {
            toast.remove();
        });
        
        // Auto remove after 4 seconds
        setTimeout(() => {
            if (toast.parentNode) {
                toast.style.animation = "slideInRight 0.3s ease reverse";
                setTimeout(() => toast.remove(), 300);
            }
        }, 4000);
    }
};

// --- SYSTEM SYNC MANAGER (GOOGLE SHEETS BRIDGE) ---
const SyncManager = {
    async testConnection(url) {
        if (!url) return { success: false, message: "กรุณาระบุ URL" };
        
        try {
            const response = await fetch(url, {
                method: "POST",
                mode: "cors",
                headers: { "Content-Type": "text/plain" },
                body: JSON.stringify({ action: "test" })
            });
            const res = await response.json();
            return { success: res.status === "success", message: res.message || "เชื่อมต่อเรียบร้อย" };
        } catch (e) {
            console.error(e);
            return { success: false, message: "ไม่สามารถเชื่อมต่อได้: ตรวจสอบความถูกต้องของ URL และการตั้งค่า Deploy ของคุณ" };
        }
    },
    
    async pushData() {
        const url = TaraApp.settings.googleSheetUrl;
        if (!url) return;
        
        this.setUIStatus("syncing", "กำลังส่งข้อมูล...");
        try {
            const payload = {
                action: "push",
                data: {
                    products: TaraApp.products,
                    sales: TaraApp.sales,
                    expenses: TaraApp.expenses
                }
            };
            
            const response = await fetch(url, {
                method: "POST",
                mode: "cors",
                headers: { "Content-Type": "text/plain" },
                body: JSON.stringify(payload)
            });
            const res = await response.json();
            
            if (res.status === "success") {
                this.setUIStatus("online", "เชื่อมต่อชีตแล้ว");
                TaraApp.showToast("ซิงก์ข้อมูลขึ้น Google Sheet สำเร็จ!", "success");
            } else {
                throw new Error(res.message);
            }
        } catch (e) {
            console.error(e);
            this.setUIStatus("offline", "การซิงก์ล้มเหลว (Offline)");
            TaraApp.showToast("การซิงก์ข้อมูลขึ้น Google Sheet ล้มเหลว", "error");
        }
    },
    
    async pullData() {
        const url = TaraApp.settings.googleSheetUrl;
        if (!url) {
            TaraApp.showToast("กรุณาตั้งค่า URL Google Sheets ในหน้าตั้งค่าก่อน", "warning");
            return;
        }
        
        this.setUIStatus("syncing", "กำลังดึงข้อมูล...");
        try {
            const response = await fetch(url, {
                method: "GET",
                mode: "cors"
            });
            const res = await response.json();
            
            if (res.status === "success" && res.data) {
                // Validate data shapes
                if (res.data.products) TaraApp.products = res.data.products;
                if (res.data.sales) TaraApp.sales = res.data.sales;
                if (res.data.expenses) TaraApp.expenses = res.data.expenses;
                
                // Parse IDs correctly and types
                TaraApp.products.forEach(p => {
                    p.price = Number(p.price) || 0;
                    p.cost = Number(p.cost) || 0;
                    p.stock = Number(p.stock) || 0;
                    p.minStock = Number(p.minStock) || 0;
                    p.image = p.image || null;
                });
                
                TaraApp.sales.forEach(s => {
                    s.timestamp = Number(s.timestamp) || Date.now();
                    s.subtotal = Number(s.subtotal);
                    s.discount = Number(s.discount);
                    s.total = Number(s.total);
                    s.cashReceived = Number(s.cashReceived);
                    s.change = Number(s.change);
                });

                TaraApp.expenses.forEach(e => {
                    e.timestamp = Number(e.timestamp) || Date.now();
                    e.amount = Number(e.amount);
                });
                
                TaraApp.saveState();
                this.setUIStatus("online", "เชื่อมต่อชีตแล้ว");
                TaraApp.showToast("ดึงข้อมูลจาก Google Sheet สำเร็จ!", "success");
                
                // Refresh App UI
                UI.refreshAll();
            } else {
                throw new Error("โครงสร้างข้อมูลในชีตไม่สมบูรณ์");
            }
        } catch (e) {
            console.error(e);
            this.setUIStatus("offline", "ดึงข้อมูลล้มเหลว (Offline)");
            TaraApp.showToast("ดึงข้อมูลจาก Google Sheet ล้มเหลว", "error");
        }
    },
    
    setUIStatus(status, text) {
        const indicator = document.getElementById("connection-indicator");
        const statusText = document.getElementById("connection-status-text");
        
        indicator.className = `sync-status ${status}`;
        statusText.textContent = text;
    }
};

// --- USER INTERFACE RENDERER ---
const UI = {
    currentTab: "dashboard",
    salesChartInstance: null,
    
    init() {
        this.setupClock();
        this.setupNavigation();
        this.setupEventListeners();
        this.setupPOSFilters();
        
        // Load settings URL and check status
        const urlInput = document.getElementById("google-sheet-url");
        if (TaraApp.settings.googleSheetUrl) {
            urlInput.value = TaraApp.settings.googleSheetUrl;
            SyncManager.setUIStatus("online", "เชื่อมต่อชีตแล้ว");
        } else {
            SyncManager.setUIStatus("offline", "Local Storage Only");
        }
        
        // Load apps script copy content
        document.getElementById("apps-script-textarea").value = GOOGLE_APPS_SCRIPT_CODE;
        
        // Render initial state
        this.refreshAll();
    },
    
    setupClock() {
        const clock = document.getElementById("clock-display");
        const updateTime = () => {
            const now = new Date();
            clock.textContent = now.toLocaleTimeString("th-TH");
        };
        setInterval(updateTime, 1000);
        updateTime();
    },
    
    setupNavigation() {
        const menuItems = document.querySelectorAll(".menu-item");
        const panels = document.querySelectorAll(".tab-panel");
        const pageTitle = document.getElementById("current-page-title");
        const sidebar = document.querySelector(".sidebar");
        
        menuItems.forEach(item => {
            item.addEventListener("click", (e) => {
                e.preventDefault();
                const tab = item.getAttribute("data-tab");
                
                menuItems.forEach(mi => mi.classList.remove("active"));
                item.classList.add("active");
                
                panels.forEach(p => p.classList.remove("active"));
                document.getElementById(`tab-${tab}`).classList.add("active");
                
                this.currentTab = tab;
                
                // Update title
                const titles = {
                    dashboard: "แดชบอร์ดสรุปผล",
                    pos: "ระบบขายหน้าร้าน (POS)",
                    stock: "การจัดการสต็อกและสินค้า",
                    accounting: "การบัญชีและบันทึกรายจ่าย",
                    settings: "ตั้งค่าระบบและการเชื่อมต่อ"
                };
                pageTitle.textContent = titles[tab];
                
                // Redraw chart if dashboard active
                if (tab === "dashboard") {
                    this.renderSalesChart();
                }
                
                // Close mobile sidebar if open
                sidebar.classList.remove("active");
            });
        });
        
        // Link on dashboard to ledger
        document.getElementById("dashboard-to-ledger").addEventListener("click", (e) => {
            e.preventDefault();
            document.getElementById("nav-accounting").click();
        });
        
        // Mobile toggle
        document.getElementById("mobile-nav-toggle").addEventListener("click", () => {
            sidebar.classList.toggle("active");
        });
    },
    
    setupPOSFilters() {
        const filters = document.querySelectorAll("#pos-category-filters .filter-btn");
        filters.forEach(btn => {
            btn.addEventListener("click", () => {
                filters.forEach(f => f.classList.remove("active"));
                btn.classList.add("active");
                this.renderPOSProducts();
            });
        });
        
        document.getElementById("pos-search-input").addEventListener("input", () => {
            this.renderPOSProducts();
        });
    },
    
    setupEventListeners() {
        // --- POS CART CONTROLS ---
        const discountInput = document.getElementById("cart-discount-input");
        discountInput.addEventListener("input", () => {
            this.updateCartSummary();
        });
        
        const cashInput = document.getElementById("payment-cash-received");
        cashInput.addEventListener("input", () => {
            this.calculateChange();
        });
        
        // Quick cash click
        const quickCashBtns = document.querySelectorAll(".quick-cash-btn");
        quickCashBtns.forEach(btn => {
            btn.addEventListener("click", () => {
                const val = btn.getAttribute("data-value");
                if (val) {
                    cashInput.value = Number(val);
                } else if (btn.id === "cash-btn-exact") {
                    const total = this.getCartTotal();
                    cashInput.value = Math.ceil(total);
                }
                this.calculateChange();
            });
        });
        
        // Checkout
        document.getElementById("checkout-submit-btn").addEventListener("click", () => {
            this.processCheckout();
        });
        
        // Clear Cart
        document.getElementById("clear-cart-btn").addEventListener("click", () => {
            TaraApp.cart = [];
            this.renderCart();
            TaraApp.showToast("ล้างตะกร้าสินค้าแล้ว", "warning");
        });
        
        // --- STOCK CONTROLS ---
        // Search
        document.getElementById("stock-search-input").addEventListener("input", () => {
            this.renderStockTable();
        });
        
        // Low Stock Filter toggle
        const filterLowBtn = document.getElementById("stock-filter-low");
        filterLowBtn.addEventListener("click", () => {
            const isActive = filterLowBtn.getAttribute("data-active") === "true";
            filterLowBtn.setAttribute("data-active", !isActive);
            filterLowBtn.classList.toggle("active");
            this.renderStockTable();
        });
        
        // Open Add Modal
        document.getElementById("open-add-product-modal-btn").addEventListener("click", () => {
            this.openProductModal();
        });
        
        // Cancel Add Modal
        document.getElementById("product-modal-close").addEventListener("click", () => this.closeProductModal());
        document.getElementById("product-modal-cancel").addEventListener("click", () => this.closeProductModal());
        
        // Add Modal submit
        document.getElementById("product-modal-form").addEventListener("submit", (e) => {
            e.preventDefault();
            this.handleProductFormSubmit();
        });

        // Product image file upload and resize logic
        const fileInput = document.getElementById("product-image-file");
        if (fileInput) {
            fileInput.addEventListener("change", (e) => {
                const file = e.target.files[0];
                if (!file) return;
                
                const reader = new FileReader();
                reader.onload = (event) => {
                    const img = new Image();
                    img.onload = () => {
                        // Resize using a Canvas to max 120x120 pixels to keep base64 string extremely small
                        const canvas = document.createElement("canvas");
                        const ctx = canvas.getContext("2d");
                        
                        const maxDim = 120;
                        let w = img.width;
                        let h = img.height;
                        
                        if (w > h) {
                            if (w > maxDim) {
                                h = Math.round(h * maxDim / w);
                                w = maxDim;
                            }
                        } else {
                            if (h > maxDim) {
                                w = Math.round(w * maxDim / h);
                                h = maxDim;
                            }
                        }
                        
                        canvas.width = w;
                        canvas.height = h;
                        ctx.drawImage(img, 0, 0, w, h);
                        
                        // Output compressed base64 JPEG
                        const dataUrl = canvas.toDataURL("image/jpeg", 0.65);
                        document.getElementById("product-image-base64").value = dataUrl;
                        
                        // Show preview
                        const preview = document.getElementById("image-upload-preview");
                        preview.innerHTML = `<img src="${dataUrl}" alt="Preview" style="height:100%; object-fit:cover;">`;
                    };
                    img.src = event.target.result;
                };
                reader.readAsDataURL(file);
            });
        }
        
        // --- ACCOUNTING CONTROLS ---
        // Expense submission
        const expenseForm = document.getElementById("accounting-expense-form");
        document.getElementById("expense-date").value = new Date().toISOString().split('T')[0];
        
        expenseForm.addEventListener("submit", (e) => {
            e.preventDefault();
            this.handleAddExpense();
        });
        
        // Ledger Filters
        document.getElementById("ledger-filter-type").addEventListener("change", () => {
            this.renderLedgerTable();
        });
        
        // --- SETTINGS CONTROLS ---
        // Test Connection
        document.getElementById("btn-test-sheet-conn").addEventListener("click", async () => {
            const btn = document.getElementById("btn-test-sheet-conn");
            const originalText = btn.innerHTML;
            btn.disabled = true;
            btn.textContent = "กำลังเชื่อมต่อ...";
            
            const url = document.getElementById("google-sheet-url").value.trim();
            const result = await SyncManager.testConnection(url);
            
            btn.disabled = false;
            btn.innerHTML = originalText;
            
            if (result.success) {
                TaraApp.showToast(result.message, "success");
            } else {
                TaraApp.showToast(result.message, "error");
            }
        });
        
        // Save Sheet URL
        document.getElementById("btn-save-sheet-url").addEventListener("click", () => {
            const url = document.getElementById("google-sheet-url").value.trim();
            TaraApp.settings.googleSheetUrl = url;
            TaraApp.saveState();
            
            if (url) {
                SyncManager.setUIStatus("online", "เชื่อมต่อชีตแล้ว");
                // Auto push on save
                SyncManager.pushData();
            } else {
                SyncManager.setUIStatus("offline", "Local Storage Only");
                TaraApp.showToast("บันทึกการตั้งค่าแล้ว", "success");
            }
        });
        
        // Sync Buttons
        document.getElementById("btn-sync-pull").addEventListener("click", () => {
            SyncManager.pullData();
        });
        document.getElementById("btn-sync-push").addEventListener("click", () => {
            SyncManager.pushData();
        });
        document.getElementById("header-sync-btn").addEventListener("click", () => {
            SyncManager.pushData();
        });
        
        // Demo and resets
        document.getElementById("btn-load-demo-data").addEventListener("click", () => {
            if (confirm("คุณแน่ใจว่าต้องการโหลดข้อมูลตัวอย่างสำเร็จรูป? ข้อมูลที่แก้ไขไปแล้วอาจถูกเขียนทับ")) {
                TaraApp.loadDemoData();
                this.refreshAll();
            }
        });
        
        document.getElementById("btn-reset-app-data").addEventListener("click", () => {
            if (confirm("ระวัง! คุณต้องการลบข้อมูลเครื่องดื่ม สต็อก ยอดขาย และรายจ่ายทั้งหมดในเบราว์เซอร์นี้? การลบนี้ไม่สามารถย้อนกลับได้ (แต่ไม่ส่งผลต่อข้อมูลใน Google Sheets)")) {
                TaraApp.resetData();
                this.refreshAll();
            }
        });
        
        // Apps script copy click
        document.getElementById("btn-copy-script").addEventListener("click", () => {
            const textarea = document.getElementById("apps-script-textarea");
            textarea.select();
            document.execCommand("copy");
            TaraApp.showToast("คัดลอกสคริปต์ลงคลิปบอร์ดแล้ว!", "success");
        });
        
        // Receipt close modal
        document.getElementById("receipt-modal-close").addEventListener("click", () => this.closeReceiptModal());
        document.getElementById("btn-close-receipt").addEventListener("click", () => this.closeReceiptModal());
        document.getElementById("btn-print-receipt").addEventListener("click", () => {
            window.print();
        });
    },
    
    refreshAll() {
        this.renderDashboardStats();
        this.renderSalesChart();
        this.renderPOSProducts();
        this.renderCart();
        this.renderStockTable();
        this.renderLedgerTable();
        this.renderAccountingStats();
    },
    
    // ==========================================================================
    // RENDER: DASHBOARD
    // ==========================================================================
    renderDashboardStats() {
        const todayStr = new Date().toISOString().split('T')[0];
        
        // Filter sales for today
        const todaySales = TaraApp.sales.filter(s => {
            const dateStr = new Date(s.timestamp).toISOString().split('T')[0];
            return dateStr === todayStr;
        });
        
        const todaySalesAmount = todaySales.reduce((acc, curr) => acc + curr.total, 0);
        const todayOrdersCount = todaySales.length;
        
        // Calculate COGS for today's sales
        let todayCOGS = 0;
        todaySales.forEach(sale => {
            sale.items.forEach(item => {
                todayCOGS += (Number(item.cost) || 0) * (Number(item.qty) || 0);
            });
        });
        
        // Filter today expenses
        const todayExpenses = TaraApp.expenses.filter(e => e.date === todayStr);
        const todayExpensesAmount = todayExpenses.reduce((acc, curr) => acc + curr.amount, 0);
        
        const todayNetProfit = todaySalesAmount - todayCOGS - todayExpensesAmount;
        
        // Alert Low Stocks count
        const lowStockProducts = TaraApp.products.filter(p => p.stock <= p.minStock);
        const lowStockCount = lowStockProducts.length;
        
        // Update elements
        document.getElementById("stat-sales-today").textContent = `฿${todaySalesAmount.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        document.getElementById("stat-orders-today").textContent = `${todayOrdersCount} บิล`;
        document.getElementById("stat-profit-today").textContent = `฿${todayNetProfit.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        document.getElementById("stat-low-stock-count").textContent = `${lowStockCount} รายการ`;
        
        const lowStockCard = document.getElementById("stat-low-stock-card");
        if (lowStockCount > 0) {
            lowStockCard.classList.add("danger");
        } else {
            lowStockCard.classList.remove("danger");
        }
        
        // Yesterday comparison logic
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        const yesterdayStr = yesterday.toISOString().split('T')[0];
        
        const yesterdaySales = TaraApp.sales.filter(s => {
            const dateStr = new Date(s.timestamp).toISOString().split('T')[0];
            return dateStr === yesterdayStr;
        });
        const yesterdaySalesAmount = yesterdaySales.reduce((acc, curr) => acc + curr.total, 0);
        
        const diffAmount = todaySalesAmount - yesterdaySalesAmount;
        const trendEl = document.getElementById("stat-sales-diff");
        
        if (diffAmount >= 0) {
            trendEl.className = "stat-trend trend-up";
            trendEl.innerHTML = `<span style="font-weight:bold;">+฿${diffAmount.toLocaleString('th-TH', {maximumFractionDigits: 0})}</span> เทียบเมื่อวาน`;
        } else {
            trendEl.className = "stat-trend trend-down";
            trendEl.innerHTML = `<span style="font-weight:bold;">-฿${Math.abs(diffAmount).toLocaleString('th-TH', {maximumFractionDigits: 0})}</span> เทียบเมื่อวาน`;
        }
        
        // Render low stock alerts list
        const alertList = document.getElementById("low-stock-alert-list");
        alertList.innerHTML = "";
        
        if (lowStockProducts.length === 0) {
            alertList.innerHTML = `
                <div class="empty-state">
                    <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
                    <p>สต็อกสินค้าทุกรายการยังคงปกติ</p>
                </div>
            `;
        } else {
            lowStockProducts.forEach(p => {
                const isOut = p.stock === 0;
                const div = document.createElement("div");
                div.className = `alert-item ${isOut ? 'out-of-stock' : ''}`;
                div.innerHTML = `
                    <div class="alert-item-details">
                        <h4>${p.name}</h4>
                        <span>เกณฑ์ขั้นต่ำ: ${p.minStock} | หมวดหมู่: ${this.getCategoryLabel(p.category)}</span>
                    </div>
                    <div class="alert-action">
                        <span class="alert-stock-count">${isOut ? 'หมดเกลี้ยง' : p.stock + ' ชิ้น'}</span>
                        <button class="btn btn-sm btn-secondary" onclick="UI.quickAddStock('${p.id}', 10)">+10</button>
                    </div>
                `;
                alertList.appendChild(div);
            });
        }
        
        // Render recent sales
        const recentTbody = document.getElementById("recent-sales-tbody");
        recentTbody.innerHTML = "";
        
        // Get today's sales sorted by newest
        const sortedTodaySales = [...todaySales].sort((a, b) => b.timestamp - a.timestamp).slice(0, 5);
        
        if (sortedTodaySales.length === 0) {
            recentTbody.innerHTML = `
                <tr>
                    <td colspan="7" class="text-center" style="color: var(--text-muted); padding: 30px;">ไม่มีรายการขายสำเร็จในวันนี้</td>
                </tr>
            `;
        } else {
            sortedTodaySales.forEach(sale => {
                const tr = document.createElement("tr");
                const timeStr = new Date(sale.timestamp).toLocaleTimeString("th-TH", {hour: '2-digit', minute:'2-digit'});
                const itemsSummary = sale.items.map(item => `${item.name} (${item.qty})`).join(", ");
                
                tr.innerHTML = `
                    <td style="font-family:var(--font-english); font-weight:600;">${sale.id}</td>
                    <td style="font-family:var(--font-english);">${timeStr}</td>
                    <td title="${itemsSummary}" style="max-width:260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${itemsSummary}</td>
                    <td style="font-family:var(--font-english); font-weight:600;">฿${sale.total.toLocaleString('th-TH', {minimumFractionDigits: 2})}</td>
                    <td style="font-family:var(--font-english);">฿${sale.cashReceived.toLocaleString('th-TH', {minimumFractionDigits: 2})}</td>
                    <td style="font-family:var(--font-english); color:var(--accent-green);">฿${sale.change.toLocaleString('th-TH', {minimumFractionDigits: 2})}</td>
                    <td>
                        <button class="btn btn-sm btn-secondary" onclick="UI.showReceipt('${sale.id}')">บิลใบเสร็จ</button>
                    </td>
                `;
                recentTbody.appendChild(tr);
            });
        }
    },
    
    renderSalesChart() {
        const canvas = document.getElementById("salesChart");
        if (!canvas) return;
        
        const ctx = canvas.getContext("2d");
        
        // Setup proper DPI scaling
        const rect = canvas.getBoundingClientRect();
        canvas.width = rect.width * window.devicePixelRatio;
        canvas.height = rect.height * window.devicePixelRatio;
        ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
        
        const width = rect.width;
        const height = rect.height;
        const padding = { top: 30, right: 20, bottom: 40, left: 55 };
        const chartWidth = width - padding.left - padding.right;
        const chartHeight = height - padding.top - padding.bottom;
        
        // Gather sales for the last 7 days
        const last7Days = [];
        const dayLabels = [];
        const now = new Date();
        
        for (let i = 6; i >= 0; i--) {
            const date = new Date(now);
            date.setDate(now.getDate() - i);
            const dateStr = date.toISOString().split('T')[0];
            
            // Calc total sales for this day
            const daySales = TaraApp.sales.filter(s => {
                const sDate = new Date(s.timestamp).toISOString().split('T')[0];
                return sDate === dateStr;
            });
            const sum = daySales.reduce((acc, curr) => acc + curr.total, 0);
            
            last7Days.push(sum);
            
            // Label e.g., "14 ก.ค."
            const formatter = new Intl.DateTimeFormat('th-TH', { day: 'numeric', month: 'short' });
            dayLabels.push(formatter.format(date));
        }
        
        // Find max value
        const maxVal = Math.max(...last7Days, 500); // minimum scale ceiling is 500
        const scaleCeil = Math.ceil(maxVal / 100) * 100;
        
        // Clear canvas
        ctx.clearRect(0, 0, width, height);
        
        // Draw grid lines and labels
        const gridLines = 5;
        ctx.strokeStyle = "rgba(255, 255, 255, 0.05)";
        ctx.lineWidth = 1;
        ctx.fillStyle = "#94a3b8"; // text color
        ctx.font = "10px Inter, Kanit";
        ctx.textAlign = "right";
        ctx.textBaseline = "middle";
        
        for (let i = 0; i <= gridLines; i++) {
            const y = padding.top + (chartHeight / gridLines) * i;
            const val = scaleCeil - (scaleCeil / gridLines) * i;
            
            ctx.beginPath();
            ctx.moveTo(padding.left, y);
            ctx.lineTo(width - padding.right, y);
            ctx.stroke();
            
            ctx.fillText(`฿${val.toLocaleString('th-TH', {maximumFractionDigits: 0})}`, padding.left - 10, y);
        }
        
        // Map points coordinates
        const points = [];
        const stepX = chartWidth / (last7Days.length - 1);
        
        for (let i = 0; i < last7Days.length; i++) {
            const x = padding.left + stepX * i;
            const ratio = last7Days[i] / scaleCeil;
            const y = padding.top + chartHeight * (1 - ratio);
            points.push({ x, y, val: last7Days[i] });
        }
        
        // Draw linear gradient under the line
        if (points.length > 1) {
            const grad = ctx.createLinearGradient(0, padding.top, 0, height - padding.bottom);
            grad.addColorStop(0, "rgba(14, 165, 233, 0.28)");
            grad.addColorStop(1, "rgba(14, 165, 233, 0)");
            
            ctx.fillStyle = grad;
            ctx.beginPath();
            ctx.moveTo(points[0].x, height - padding.bottom);
            
            // Draw smooth curve path
            ctx.lineTo(points[0].x, points[0].y);
            for (let i = 0; i < points.length - 1; i++) {
                const xc = (points[i].x + points[i + 1].x) / 2;
                const yc = (points[i].y + points[i + 1].y) / 2;
                ctx.quadraticCurveTo(points[i].x, points[i].y, xc, yc);
            }
            ctx.lineTo(points[points.length - 1].x, points[points.length - 1].y);
            
            ctx.lineTo(points[points.length - 1].x, height - padding.bottom);
            ctx.closePath();
            ctx.fill();
        }
        
        // Draw path line
        ctx.strokeStyle = "#38bdf8";
        ctx.lineWidth = 3;
        ctx.beginPath();
        ctx.moveTo(points[0].x, points[0].y);
        for (let i = 0; i < points.length - 1; i++) {
            const xc = (points[i].x + points[i + 1].x) / 2;
            const yc = (points[i].y + points[i + 1].y) / 2;
            ctx.quadraticCurveTo(points[i].x, points[i].y, xc, yc);
        }
        ctx.lineTo(points[points.length - 1].x, points[points.length - 1].y);
        ctx.stroke();
        
        // Draw points and values
        ctx.fillStyle = "#ffffff";
        ctx.strokeStyle = "#0ea5e9";
        ctx.lineWidth = 2;
        ctx.textAlign = "center";
        
        points.forEach((pt, idx) => {
            // Circle
            ctx.beginPath();
            ctx.arc(pt.x, pt.y, 4.5, 0, Math.PI * 2);
            ctx.fill();
            ctx.stroke();
            
            // Draw value above point
            if (pt.val > 0) {
                ctx.fillStyle = "#f8fafc";
                ctx.font = "bold 9px Inter";
                ctx.fillText(`฿${pt.val.toLocaleString('th-TH', {maximumFractionDigits: 0})}`, pt.x, pt.y - 10);
            }
        });
        
        // Draw X Axis date labels
        ctx.fillStyle = "#94a3b8";
        ctx.font = "10px Kanit";
        ctx.textAlign = "center";
        ctx.textBaseline = "top";
        
        for (let i = 0; i < dayLabels.length; i++) {
            const x = padding.left + stepX * i;
            ctx.fillText(dayLabels[i], x, height - padding.bottom + 10);
        }
    },
    
    // ==========================================================================
    // RENDER: POS PANEL
    // ==========================================================================
    renderPOSProducts() {
        const grid = document.getElementById("pos-products-grid");
        grid.innerHTML = "";
        
        const searchVal = document.getElementById("pos-search-input").value.toLowerCase();
        const activeCategory = document.querySelector("#pos-category-filters .filter-btn.active").getAttribute("data-category");
        
        const filtered = TaraApp.products.filter(p => {
            const matchesSearch = p.name.toLowerCase().includes(searchVal);
            const matchesCat = activeCategory === "all" || p.category === activeCategory;
            return matchesSearch && matchesCat;
        });
        
        if (filtered.length === 0) {
            grid.innerHTML = `
                <div class="empty-state" style="grid-column: 1 / -1; width:100%; padding:60px 0;">
                    <svg xmlns="http://www.w3.org/2000/svg" width="60" height="60" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
                    <p>ไม่พบรายการเครื่องดื่มที่คุณค้นหา</p>
                </div>
            `;
            return;
        }
        
        filtered.forEach(p => {
            const isOutOfStock = p.stock <= 0;
            const isLowStock = p.stock > 0 && p.stock <= p.minStock;
            
            let stockBadge = `<span class="badge badge-success product-stock-badge">มีสินค้า (${p.stock})</span>`;
            if (isOutOfStock) {
                stockBadge = `<span class="badge badge-danger product-stock-badge">สินค้าหมด</span>`;
            } else if (isLowStock) {
                stockBadge = `<span class="badge badge-warning product-stock-badge">สต็อกต่ำ (${p.stock})</span>`;
            }
            
            let imageDisplay = this.getImageEmoji(p.imageCode);
            if (p.image && p.image.startsWith("data:image")) {
                imageDisplay = `<img src="${p.image}" alt="${p.name}">`;
            }
            
            const card = document.createElement("div");
            card.className = `product-card ${isOutOfStock ? 'out-of-stock-card' : ''}`;
            card.innerHTML = `
                <div class="product-image-container">
                    ${imageDisplay}
                    ${stockBadge}
                </div>
                <div class="product-info-container">
                    <h4 title="${p.name}">${p.name}</h4>
                    <div class="product-price-row">
                        <span class="product-price">฿${p.price.toFixed(2)}</span>
                        <button class="btn-add-cart" onclick="UI.addToCart('${p.id}')" ${isOutOfStock ? 'disabled' : ''} title="เพิ่มลงตะกร้า">
                            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                        </button>
                    </div>
                </div>
            `;
            grid.appendChild(card);
        });
    },
    
    renderCart() {
        const cartList = document.getElementById("cart-items-list");
        cartList.innerHTML = "";
        
        if (TaraApp.cart.length === 0) {
            cartList.innerHTML = `
                <div class="empty-cart-state">
                    <svg xmlns="http://www.w3.org/2000/svg" width="56" height="56" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/></svg>
                    <p>ตะกร้ายังว่างเปล่า<br>เลือกรายการเครื่องดื่มด้านข้างเพื่อขาย</p>
                </div>
            `;
            this.updateCartSummary();
            return;
        }
        
        TaraApp.cart.forEach(item => {
            const div = document.createElement("div");
            div.className = "cart-item";
            
            const totalItemPrice = item.product.price * item.quantity;
            
            div.innerHTML = `
                <div class="cart-item-name">
                    <h5>${item.product.name}</h5>
                    <span>@฿${item.product.price.toFixed(2)}</span>
                </div>
                <div class="cart-item-controls">
                    <button class="qty-btn" onclick="UI.changeCartQty('${item.product.id}', -1)">-</button>
                    <span class="cart-item-qty">${item.quantity}</span>
                    <button class="qty-btn" onclick="UI.changeCartQty('${item.product.id}', 1)">+</button>
                    <span class="cart-item-price">฿${totalItemPrice.toFixed(2)}</span>
                </div>
            `;
            cartList.appendChild(div);
        });
        
        this.updateCartSummary();
    },
    
    addToCart(productId) {
        const product = TaraApp.products.find(p => p.id === productId);
        if (!product) return;
        
        const existing = TaraApp.cart.find(item => item.product.id === productId);
        
        if (existing) {
            if (existing.quantity >= product.stock) {
                TaraApp.showToast(`ไม่สามารถซื้อเพิ่มได้เนื่องจากจำนวนสต็อกจำกัด (มีอยู่ ${product.stock})`, "warning");
                return;
            }
            existing.quantity++;
        } else {
            TaraApp.cart.push({ product, quantity: 1 });
        }
        
        this.renderCart();
    },
    
    changeCartQty(productId, amount) {
        const itemIndex = TaraApp.cart.findIndex(item => item.product.id === productId);
        if (itemIndex === -1) return;
        
        const item = TaraApp.cart[itemIndex];
        const product = TaraApp.products.find(p => p.id === productId);
        
        const newQty = item.quantity + amount;
        
        if (newQty <= 0) {
            TaraApp.cart.splice(itemIndex, 1);
        } else {
            if (newQty > product.stock) {
                TaraApp.showToast(`เกินจำนวนจำกัดในคลังสินค้า (คงเหลือ ${product.stock})`, "warning");
                return;
            }
            item.quantity = newQty;
        }
        
        this.renderCart();
    },
    
    getCartSubtotal() {
        return TaraApp.cart.reduce((acc, curr) => acc + (curr.product.price * curr.quantity), 0);
    },
    
    getCartTotal() {
        const sub = this.getCartSubtotal();
        const discount = Number(document.getElementById("cart-discount-input").value) || 0;
        return Math.max(0, sub - discount);
    },
    
    updateCartSummary() {
        const subtotal = this.getCartSubtotal();
        const discount = Number(document.getElementById("cart-discount-input").value) || 0;
        const total = Math.max(0, subtotal - discount);
        
        document.getElementById("cart-subtotal").textContent = `฿${subtotal.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        document.getElementById("cart-total").textContent = `฿${total.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        
        this.calculateChange();
    },
    
    calculateChange() {
        const total = this.getCartTotal();
        const cashReceivedInput = document.getElementById("payment-cash-received");
        const changeDisplay = document.getElementById("payment-change-value");
        const checkoutBtn = document.getElementById("checkout-submit-btn");
        
        const cashReceived = Number(cashReceivedInput.value) || 0;
        
        if (TaraApp.cart.length === 0) {
            changeDisplay.textContent = "฿0.00";
            checkoutBtn.disabled = true;
            return;
        }
        
        const change = cashReceived - total;
        
        if (change >= 0) {
            changeDisplay.textContent = `฿${change.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
            checkoutBtn.disabled = false;
        } else {
            changeDisplay.textContent = "฿0.00";
            checkoutBtn.disabled = true; // Not enough cash entered
        }
    },
    
    processCheckout() {
        const total = this.getCartTotal();
        const cashInput = document.getElementById("payment-cash-received");
        const cashReceived = Number(cashInput.value) || 0;
        const change = cashReceived - total;
        const discount = Number(document.getElementById("cart-discount-input").value) || 0;
        
        if (cashReceived < total) {
            TaraApp.showToast("ยอดเงินสดที่รับมาไม่เพียงพอ", "error");
            return;
        }
        
        const receiptId = "TR-" + (100000 + Math.floor(Math.random() * 900000));
        const timestamp = Date.now();
        
        // Build items list
        const items = TaraApp.cart.map(item => ({
            id: item.product.id,
            name: item.product.name,
            qty: item.quantity,
            price: item.product.price,
            cost: item.product.cost || 0
        }));
        
        // Create Transaction Sale Record
        const saleRecord = {
            id: receiptId,
            timestamp: timestamp,
            items: items,
            subtotal: this.getCartSubtotal(),
            discount: discount,
            total: total,
            cashReceived: cashReceived,
            change: change
        };
        
        // Save to state
        TaraApp.sales.push(saleRecord);
        
        // Deduct Stock
        TaraApp.cart.forEach(item => {
            const product = TaraApp.products.find(p => p.id === item.product.id);
            if (product) {
                product.stock = Math.max(0, product.stock - item.quantity);
            }
        });
        
        // Save State
        TaraApp.saveState();
        TaraApp.showToast("บันทึกการขายและพิมพ์ใบเสร็จเรียบร้อย", "success");
        
        // Reset Cart and Form
        TaraApp.cart = [];
        cashInput.value = "";
        document.getElementById("cart-discount-input").value = 0;
        
        // Update UIs
        this.refreshAll();
        
        // Show Receipt modal
        this.showReceipt(receiptId);
        
        // Auto push to Google Sheet in background
        SyncManager.pushData();
    },
    
    showReceipt(saleId) {
        const sale = TaraApp.sales.find(s => s.id === saleId);
        if (!sale) return;
        
        document.getElementById("rec-number").textContent = sale.id;
        
        const dateObj = new Date(sale.timestamp);
        document.getElementById("rec-date").textContent = dateObj.toLocaleDateString("th-TH");
        document.getElementById("rec-time").textContent = dateObj.toLocaleTimeString("th-TH", {hour: '2-digit', minute:'2-digit'});
        
        const tbody = document.getElementById("receipt-items-tbody");
        tbody.innerHTML = "";
        
        sale.items.forEach(item => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${item.name}</td>
                <td class="text-center" style="font-family:var(--font-english);">${item.qty}</td>
                <td class="text-right" style="font-family:var(--font-english);">฿${(item.price * item.qty).toFixed(2)}</td>
            `;
            tbody.appendChild(tr);
        });
        
        document.getElementById("rec-subtotal").textContent = `฿${sale.subtotal.toFixed(2)}`;
        
        const discountRow = document.getElementById("rec-discount-row");
        if (sale.discount > 0) {
            discountRow.style.display = "flex";
            document.getElementById("rec-discount").textContent = `-฿${sale.discount.toFixed(2)}`;
        } else {
            discountRow.style.display = "none";
        }
        
        document.getElementById("rec-total").textContent = `฿${sale.total.toFixed(2)}`;
        document.getElementById("rec-cash-received").textContent = `฿${sale.cashReceived.toFixed(2)}`;
        document.getElementById("rec-change").textContent = `฿${sale.change.toFixed(2)}`;
        
        document.getElementById("receipt-modal").classList.add("active");
    },
    
    closeReceiptModal() {
        document.getElementById("receipt-modal").classList.remove("active");
    },
    
    // ==========================================================================
    // RENDER: INVENTORY & STOCKS
    // ==========================================================================
    renderStockTable() {
        const tbody = document.getElementById("stock-master-tbody");
        if (!tbody) return;
        
        tbody.innerHTML = "";
        
        const searchVal = document.getElementById("stock-search-input").value.toLowerCase();
        const isFilterLowActive = document.getElementById("stock-filter-low").getAttribute("data-active") === "true";
        
        const filtered = TaraApp.products.filter(p => {
            const matchesSearch = p.name.toLowerCase().includes(searchVal) || p.id.toLowerCase().includes(searchVal);
            const matchesLow = !isFilterLowActive || (p.stock <= p.minStock);
            return matchesSearch && matchesLow;
        });
        
        if (filtered.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="9" class="text-center" style="color: var(--text-muted); padding:30px 0;">ไม่มีรายการสินค้าที่ตรงตามเงื่อนไข</td>
                </tr>
            `;
            return;
        }
        
        filtered.forEach(p => {
            const tr = document.createElement("tr");
            
            const isOutOfStock = p.stock <= 0;
            const isLowStock = p.stock > 0 && p.stock <= p.minStock;
            
            let statusBadge = `<span class="badge badge-success">เพียงพอ</span>`;
            if (isOutOfStock) {
                statusBadge = `<span class="badge badge-danger">หมดชั่วคราว</span>`;
            } else if (isLowStock) {
                statusBadge = `<span class="badge badge-warning">สต็อกใกล้หมด</span>`;
            }
            
            let imageDisplay = this.getImageEmoji(p.imageCode);
            if (p.image && p.image.startsWith("data:image")) {
                imageDisplay = `<img src="${p.image}" alt="${p.name}" style="width:100%; height:100%; object-fit:cover; border-radius:var(--radius-sm);">`;
            }
            
            tr.innerHTML = `
                <td><div class="product-thumbnail">${imageDisplay}</div></td>
                <td>
                    <div style="font-weight: 500;">${p.name}</div>
                    <div style="font-size: 0.75rem; color: var(--text-muted); font-family:var(--font-english);">${p.id}</div>
                </td>
                <td>${this.getCategoryLabel(p.category)}</td>
                <td style="font-family: var(--font-english); font-weight:600;">฿${p.price.toFixed(2)}</td>
                <td style="font-family: var(--font-english); font-weight: 600;" class="${isLowStock ? 'trend-down' : ''}">${p.stock}</td>
                <td style="font-family: var(--font-english); color:var(--text-secondary);">${p.minStock}</td>
                <td>${statusBadge}</td>
                <td>
                    <div class="quick-stock-controls">
                        <button class="btn btn-sm btn-secondary" onclick="UI.quickAddStock('${p.id}', 5)">+5</button>
                        <button class="btn btn-sm btn-secondary" onclick="UI.quickAddStock('${p.id}', 20)">+20</button>
                        <button class="btn btn-sm btn-secondary" onclick="UI.quickAddStock('${p.id}', 50)">+50</button>
                    </div>
                </td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn-sm btn-secondary" onclick="UI.openProductModal('${p.id}')">แก้ไข</button>
                        <button class="btn btn-sm btn-danger-outline" onclick="UI.deleteProduct('${p.id}')">ลบ</button>
                    </div>
                </td>
            `;
            tbody.appendChild(tr);
        });
    },
    
    quickAddStock(productId, amount) {
        const product = TaraApp.products.find(p => p.id === productId);
        if (product) {
            product.stock += amount;
            TaraApp.saveState();
            TaraApp.showToast(`เพิ่มสต็อก ${product.name} จำนวน +${amount} สำเร็จ`, "success");
            this.refreshAll();
            
            // sync updates
            SyncManager.pushData();
        }
    },
    
    deleteProduct(productId) {
        const product = TaraApp.products.find(p => p.id === productId);
        if (product && confirm(`คุณต้องการลบรายการสินค้า "${product.name}" หรือไม่?`)) {
            TaraApp.products = TaraApp.products.filter(p => p.id !== productId);
            TaraApp.saveState();
            TaraApp.showToast("ลบสินค้าเรียบร้อยแล้ว", "warning");
            this.refreshAll();
            
            SyncManager.pushData();
        }
    },
    
    openProductModal(productId = null) {
        const modal = document.getElementById("product-modal");
        const title = document.getElementById("product-modal-title");
        const form = document.getElementById("product-modal-form");
        
        form.reset();
        
        if (productId) {
            const product = TaraApp.products.find(p => p.id === productId);
            if (product) {
                title.textContent = "แก้ไขรายละเอียดสินค้า";
                document.getElementById("product-id-input").value = product.id;
                document.getElementById("product-name-input").value = product.name;
                document.getElementById("product-category-input").value = product.category;
                document.getElementById("product-cost-input").value = product.cost || 0;
                document.getElementById("product-price-input").value = product.price;
                document.getElementById("product-stock-input").value = product.stock;
                document.getElementById("product-min-stock-input").value = product.minStock;
                document.getElementById("product-image-select").value = product.imageCode;
                document.getElementById("product-image-base64").value = product.image || "";
                
                const preview = document.getElementById("image-upload-preview");
                if (product.image && product.image.startsWith("data:image")) {
                    preview.innerHTML = `<img src="${product.image}" alt="Preview" style="height:100%; object-fit:cover;">`;
                } else {
                    preview.innerHTML = `<span style="font-size: 0.8rem; color: var(--text-muted);">ไม่มีรูปภาพ</span>`;
                }
            }
        } else {
            title.textContent = "เพิ่มรายการสินค้าใหม่";
            document.getElementById("product-id-input").value = "";
            document.getElementById("product-cost-input").value = "";
            document.getElementById("product-price-input").value = "";
            document.getElementById("product-image-base64").value = "";
            document.getElementById("image-upload-preview").innerHTML = `<span style="font-size: 0.8rem; color: var(--text-muted);">ไม่มีรูปภาพ</span>`;
        }
        
        modal.classList.add("active");
    },
    
    closeProductModal() {
        document.getElementById("product-modal").classList.remove("active");
    },
    
    handleProductFormSubmit() {
        const id = document.getElementById("product-id-input").value;
        const name = document.getElementById("product-name-input").value.trim();
        const category = document.getElementById("product-category-input").value;
        const cost = Number(document.getElementById("product-cost-input").value) || 0;
        const price = Number(document.getElementById("product-price-input").value) || 0;
        const stock = Number(document.getElementById("product-stock-input").value) || 0;
        const minStock = Number(document.getElementById("product-min-stock-input").value) || 0;
        const imageCode = document.getElementById("product-image-select").value;
        const image = document.getElementById("product-image-base64").value || null;
        
        if (id) {
            // Edit existing
            const product = TaraApp.products.find(p => p.id === id);
            if (product) {
                product.name = name;
                product.category = category;
                product.cost = cost;
                product.price = price;
                product.stock = stock;
                product.minStock = minStock;
                product.imageCode = imageCode;
                product.image = image;
                TaraApp.showToast("ปรับปรุงรายละเอียดสินค้าแล้ว", "success");
            }
        } else {
            // Create new
            const newId = "P-" + (1000 + TaraApp.products.length + Math.floor(Math.random() * 100));
            TaraApp.products.push({
                id: newId,
                name: name,
                category: category,
                cost: cost,
                price: price,
                stock: stock,
                minStock: minStock,
                imageCode: imageCode,
                image: image
            });
            TaraApp.showToast("เพิ่มสินค้าใหม่สำเร็จ", "success");
        }
        
        TaraApp.saveState();
        this.closeProductModal();
        this.refreshAll();
        
        SyncManager.pushData();
    },
    
    // ==========================================================================
    // RENDER: ACCOUNTING & BUDGETS
    // ==========================================================================
    renderAccountingStats() {
        const totalIncome = TaraApp.sales.reduce((acc, curr) => acc + curr.total, 0);
        
        // Calculate COGS
        let totalCOGS = 0;
        TaraApp.sales.forEach(sale => {
            sale.items.forEach(item => {
                totalCOGS += (Number(item.cost) || 0) * (Number(item.qty) || 0);
            });
        });
        
        const totalExpense = TaraApp.expenses.reduce((acc, curr) => acc + curr.amount, 0);
        const totalCostAndExpense = totalCOGS + totalExpense;
        const netProfit = totalIncome - totalCostAndExpense;
        
        document.getElementById("acc-total-income").textContent = `฿${totalIncome.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        document.getElementById("acc-income-count").textContent = `${TaraApp.sales.length} บิลขาย`;
        
        // Show breakdown in Card 2
        document.getElementById("acc-total-expense").textContent = `฿${totalCostAndExpense.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        document.getElementById("acc-expense-count").textContent = `ต้นทุนสินค้า ฿${totalCOGS.toLocaleString('th-TH', {maximumFractionDigits: 0})} | รายจ่ายร้าน ฿${totalExpense.toLocaleString('th-TH', {maximumFractionDigits: 0})}`;
        
        document.getElementById("acc-total-net-profit").textContent = `฿${netProfit.toLocaleString('th-TH', {minimumFractionDigits: 2})}`;
        
        const netProfitPercentageEl = document.getElementById("acc-net-profit-percentage");
        if (totalIncome > 0) {
            const pct = (netProfit / totalIncome) * 100;
            netProfitPercentageEl.textContent = `คิดเป็น ${pct.toFixed(1)}% ของรายรับทั้งหมด`;
            netProfitPercentageEl.className = pct >= 0 ? "acc-card-desc trend-up" : "acc-card-desc trend-down";
        } else {
            netProfitPercentageEl.textContent = `คิดเป็น 0% ของรายรับทั้งหมด`;
            netProfitPercentageEl.className = "acc-card-desc";
        }
    },
    
    renderLedgerTable() {
        const tbody = document.getElementById("ledger-tbody");
        if (!tbody) return;
        
        tbody.innerHTML = "";
        
        const activeFilter = document.getElementById("ledger-filter-type").value;
        
        // Assemble unified ledger list
        let ledgerList = [];
        
        // Add sales (income)
        if (activeFilter === "all" || activeFilter === "income") {
            TaraApp.sales.forEach(sale => {
                ledgerList.push({
                    id: sale.id,
                    timestamp: sale.timestamp,
                    type: "income",
                    title: `ขายสินค้า บิลเลขที่ ${sale.id}`,
                    category: "ยอดขายหน้าร้าน",
                    amount: sale.total,
                    date: new Date(sale.timestamp).toISOString().split('T')[0]
                });
            });
        }
        
        // Add expenses
        if (activeFilter === "all" || activeFilter === "expense") {
            TaraApp.expenses.forEach(exp => {
                ledgerList.push({
                    id: exp.id,
                    timestamp: exp.timestamp,
                    type: "expense",
                    title: exp.title,
                    category: this.getExpenseCategoryLabel(exp.category),
                    amount: exp.amount,
                    date: exp.date
                });
            });
        }
        
        // Sort ledger list by date desc
        ledgerList.sort((a, b) => b.timestamp - a.timestamp);
        
        if (ledgerList.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" class="text-center" style="color: var(--text-muted); padding:30px 0;">ไม่มีรายการบัญชีในตาราง</td>
                </tr>
            `;
            return;
        }
        
        ledgerList.forEach(item => {
            const tr = document.createElement("tr");
            
            const isIncome = item.type === "income";
            const dateObj = new Date(item.timestamp);
            const dateStr = dateObj.toLocaleDateString("th-TH") + " " + dateObj.toLocaleTimeString("th-TH", {hour: '2-digit', minute:'2-digit'});
            
            tr.innerHTML = `
                <td style="font-family: var(--font-english); font-size:0.8rem; white-space:nowrap;">${dateStr}</td>
                <td>
                    <span class="badge ${isIncome ? 'badge-success' : 'badge-danger'}">
                        ${isIncome ? 'รายรับ' : 'รายจ่าย'}
                    </span>
                </td>
                <td style="font-weight: 500;">${item.title}</td>
                <td>${item.category}</td>
                <td class="${isIncome ? 'td-income' : 'td-expense'}" style="font-family:var(--font-english);">
                    ${isIncome ? '+' : '-'}฿${item.amount.toLocaleString('th-TH', {minimumFractionDigits: 2})}
                </td>
                <td>
                    ${isIncome ? `
                        <button class="btn btn-sm btn-secondary" onclick="UI.showReceipt('${item.id}')">ใบเสร็จ</button>
                    ` : `
                        <button class="btn btn-sm btn-danger-outline" onclick="UI.deleteExpense('${item.id}')">ลบ</button>
                    `}
                </td>
            `;
            tbody.appendChild(tr);
        });
    },
    
    handleAddExpense() {
        const title = document.getElementById("expense-title").value.trim();
        const category = document.getElementById("expense-category").value;
        const amount = Number(document.getElementById("expense-amount").value) || 0;
        const dateInput = document.getElementById("expense-date").value;
        
        if (amount <= 0) {
            TaraApp.showToast("กรุณาระบุจำนวนเงินที่ถูกต้อง", "warning");
            return;
        }
        
        const timestamp = new Date(dateInput).getTime() + (new Date().getHours()*60*60*1000) + (new Date().getMinutes()*60*1000);
        
        const newExpense = {
            id: "E-" + (1000 + TaraApp.expenses.length + Math.floor(Math.random() * 100)),
            timestamp: timestamp,
            title: title,
            category: category,
            amount: amount,
            date: dateInput
        };
        
        TaraApp.expenses.push(newExpense);
        TaraApp.saveState();
        TaraApp.showToast(`บันทึกรายจ่าย "${title}" เรียบร้อย`, "success");
        
        // Reset expense form
        document.getElementById("expense-title").value = "";
        document.getElementById("expense-amount").value = "";
        document.getElementById("expense-date").value = new Date().toISOString().split('T')[0];
        
        // Update UIs
        this.refreshAll();
        
        // Sync
        SyncManager.pushData();
    },
    
    deleteExpense(expenseId) {
        const expense = TaraApp.expenses.find(e => e.id === expenseId);
        if (expense && confirm(`ต้องการลบรายการจ่ายเงิน "${expense.title}" จำนวน ฿${expense.amount.toFixed(2)} หรือไม่?`)) {
            TaraApp.expenses = TaraApp.expenses.filter(e => e.id !== expenseId);
            TaraApp.saveState();
            TaraApp.showToast("ลบรายการรายจ่ายแล้ว", "warning");
            this.refreshAll();
            
            SyncManager.pushData();
        }
    },
    
    // ==========================================================================
    // UTILITY LABELS & HELPER FUNCTIONS
    // ==========================================================================
    getCategoryLabel(cat) {
        const map = {
            water: "น้ำดื่มแพ็ก",
            soda: "น้ำอัดลม",
            juice: "น้ำผลไม้",
            energy: "เครื่องดื่มชูกำลัง",
            other: "อื่น ๆ"
        };
        return map[cat] || cat;
    },
    
    getExpenseCategoryLabel(cat) {
        const map = {
            raw_materials: "วัตถุดิบและน้ำดื่ม",
            utilities: "ค่าน้ำ/ค่าไฟ",
            rent: "ค่าเช่าร้าน",
            salary: "ค่าแรงพนักงาน",
            transport: "ค่าเดินทาง/ขนส่ง",
            other: "อื่น ๆ"
        };
        return map[cat] || cat;
    },
    
    getImageEmoji(imageCode) {
        const map = {
            "water-pack": "💧",
            "water-bottle": "🍼",
            "soda-can": "🥤",
            "soda-bottle": "🍾",
            "juice-box": "🧃",
            "energy-bottle": "⚡",
            "other-beverage": "📦"
        };
        return map[imageCode] || "📦";
    }
};

// --- STARTUP INITIALIZER ---
document.addEventListener("DOMContentLoaded", () => {
    TaraApp.loadState();
    UI.init();
});
