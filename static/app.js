/**
 * XYZ Electronic Appliances - Stock Management Portal Client Logic
 * Handcrafted enterprise inventory engine syncing directly with Excel.
 */

// Application State
let currentUser = null;
let allProducts = [];
let currentFilterStatus = "All";
let searchDebounceTimer = null;
let categoryChartInstance = null;
let brandChartInstance = null;

// Initialization
document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    fetchInitialExcelInfo();
    checkExistingSession();
});

async function fetchInitialExcelInfo() {
    try {
        const res = await fetch("/api/excel-info");
        if (res.ok) {
            const info = await res.json();
            updateExcelStatus(info);
        }
    } catch (e) {}
}

// ----------------- THEME MANAGEMENT -----------------
function initTheme() {
    const savedTheme = localStorage.getItem("xyz_theme") || "light";
    if (savedTheme === "dark") {
        document.documentElement.classList.add("dark");
        updateThemeIcon(true);
    } else {
        document.documentElement.classList.remove("dark");
        updateThemeIcon(false);
    }
}

function toggleTheme() {
    const isDark = document.documentElement.classList.toggle("dark");
    localStorage.setItem("xyz_theme", isDark ? "dark" : "light");
    updateThemeIcon(isDark);
    if (categoryChartInstance && brandChartInstance) {
        renderCharts(allProducts);
    }
}

function updateThemeIcon(isDark) {
    const icon = document.getElementById("themeIcon");
    if (icon) {
        icon.className = isDark ? "fa-solid fa-sun text-amber-400" : "fa-solid fa-moon text-slate-600";
    }
}

// ----------------- AUTHENTICATION -----------------
function hasPermission(permission) {
    if (!currentUser || !currentUser.permissions) return false;
    return currentUser.permissions.includes("all") || currentUser.permissions.includes(permission);
}

function checkExistingSession() {
    const savedUser = localStorage.getItem("xyz_user");
    if (savedUser) {
        try {
            currentUser = JSON.parse(savedUser);
            // Clear legacy session if it has old names or missing permissions
            if (!currentUser || currentUser.name === "Alex Mercer" || !currentUser.permissions) {
                localStorage.removeItem("xyz_user");
                currentUser = null;
                showLoginView();
                return;
            }
            showAppView();
            refreshData();
            return;
        } catch (e) {
            localStorage.removeItem("xyz_user");
        }
    }
    showLoginView();
}

function quickFillLogin(email, password) {
    document.getElementById("loginEmail").value = email;
    document.getElementById("loginPassword").value = password;
    handleLogin(new Event('submit'));
}

async function handleLogin(e) {
    if (e && e.preventDefault) e.preventDefault();
    
    const email = document.getElementById("loginEmail").value.trim();
    const password = document.getElementById("loginPassword").value.trim();
    const submitBtn = document.getElementById("loginSubmitBtn");

    if (!email || !password) {
        showToast("Please enter both email and password.", "error");
        return;
    }

    try {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin"></i> <span>Verifying...</span>`;

        const res = await fetch("/api/auth/login", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ email, password })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Authentication failed");
        }

        const data = await res.json();
        currentUser = data;
        localStorage.setItem("xyz_user", JSON.stringify(currentUser));
        
        showToast(`Welcome back, ${currentUser.name}! Connected to Excel sheet.`, "success");
        showAppView();
        refreshData();
    } catch (err) {
        showToast(err.message, "error");
    } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = `<i class="fa-solid fa-arrow-right-to-bracket"></i> <span>Sign In to Inventory Portal</span>`;
    }
}

async function handleLogout() {
    if (currentUser && currentUser.token) {
        try {
            await fetch("/api/auth/logout", {
                method: "POST",
                headers: { "Authorization": `Bearer ${currentUser.token}` }
            });
        } catch (e) {}
    }
    localStorage.removeItem("xyz_user");
    currentUser = null;
    showToast("Signed out successfully.", "info");
    showLoginView();
}

function showLoginView() {
    document.getElementById("loginView").classList.remove("hidden");
    document.getElementById("appView").classList.add("hidden");
}

function showAppView() {
    document.getElementById("loginView").classList.add("hidden");
    document.getElementById("appView").classList.remove("hidden");
    
    if (currentUser) {
        document.getElementById("userName").textContent = currentUser.name;
        document.getElementById("userRole").textContent = currentUser.role;
        document.getElementById("userAvatar").textContent = currentUser.avatar || "👤";

        // Role-based button visibility
        const addBtn = document.getElementById("navAddBtn");
        if (addBtn) addBtn.style.display = hasPermission("stock_add") ? "inline-flex" : "none";

        const uploadBtn = document.getElementById("navUploadBtn");
        if (uploadBtn) uploadBtn.style.display = hasPermission("excel_upload") ? "inline-flex" : "none";

        const poBtn = document.getElementById("navPoBtn");
        if (poBtn) poBtn.style.display = hasPermission("reorder_po") ? "inline-flex" : "none";

        // Banner for role status
        const banner = document.getElementById("userPermissionBanner");
        if (banner) {
            if (currentUser.role === "Godown Staff") {
                banner.className = "text-xs text-center py-2 px-4 rounded-xl bg-amber-50 dark:bg-amber-950/70 text-amber-800 dark:text-amber-200 border border-amber-300 dark:border-amber-800 shadow-xs block no-print";
                banner.innerHTML = `<i class="fa-solid fa-warehouse mr-1.5 text-amber-600"></i><strong>Godown Staff Mode:</strong> Quick In/Out stock adjustments & barcode printing enabled. (Pricing edits & SKU deletion are restricted)`;
            } else if (currentUser.role === "Store Manager") {
                banner.className = "text-xs text-center py-2 px-4 rounded-xl bg-blue-50 dark:bg-blue-950/70 text-blue-800 dark:text-blue-200 border border-blue-300 dark:border-blue-800 shadow-xs block no-print";
                banner.innerHTML = `<i class="fa-solid fa-store mr-1.5 text-blue-600"></i><strong>Store Manager Mode:</strong> Stock management, product edits, new appliances & Purchase Orders active. (Database overwrite restricted)`;
            } else {
                banner.className = "hidden";
                banner.innerHTML = "";
            }
        }
    }
}

function getAuthHeaders() {
    return {
        "Content-Type": "application/json",
        "Authorization": currentUser ? `Bearer ${currentUser.token}` : ""
    };
}

// ----------------- DATA LOADING & SYNC -----------------
async function refreshData(manual = false) {
    const syncIcon = document.getElementById("syncIcon");
    if (syncIcon) syncIcon.classList.add("fa-spin");

    try {
        const res = await fetch("/api/stock");
        if (!res.ok) throw new Error("Could not load inventory from Excel");
        
        const data = await res.json();
        allProducts = data.products || [];
        
        updateKPIs(data.kpis);
        updateExcelStatus(data.excel_info);
        applyFilters();
        renderCharts(allProducts);
        checkReorderThresholds(allProducts);

        document.getElementById("lastSyncedTime").textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        
        if (manual) {
            showToast("Successfully synchronised with Excel on disk!", "success");
        }
    } catch (err) {
        showToast(err.message, "error");
    } finally {
        if (syncIcon) {
            setTimeout(() => syncIcon.classList.remove("fa-spin"), 400);
        }
    }
}

function updateExcelStatus(info) {
    if (!info) return;
    const navFilename = document.getElementById("navExcelFilename");
    if (navFilename) {
        navFilename.textContent = info.file_name;
        navFilename.title = info.file_path || "";
    }
    const navInfo = document.getElementById("navExcelInfo");
    if (navInfo) {
        navInfo.textContent = `(${info.file_size_kb} KB, ${info.total_rows} items)`;
    }
    const loginPath = document.getElementById("loginExcelPath");
    if (loginPath) {
        loginPath.textContent = info.relative_path || "data/XYZ_Stock_Inventory.xlsx";
        loginPath.title = `Full local path on this PC: ${info.file_path || ""}`;
    }
}

function formatCurrency(val) {
    return new Intl.NumberFormat('en-IN', {
        style: 'currency',
        currency: 'INR',
        maximumFractionDigits: 0
    }).format(val);
}

function updateKPIs(kpis) {
    if (!kpis) return;
    document.getElementById("kpiRetailVal").textContent = formatCurrency(kpis.total_retail_valuation);
    document.getElementById("kpiCostVal").textContent = formatCurrency(kpis.total_cost_valuation);
    document.getElementById("kpiMarginBadge").textContent = `+${kpis.margin_percentage}% Margin`;
    document.getElementById("kpiTotalUnits").textContent = `${kpis.total_units} Units`;
    document.getElementById("kpiTotalSkus").textContent = kpis.total_skus;
    document.getElementById("kpiLowStock").textContent = kpis.low_stock_count;
    document.getElementById("kpiOutOfStock").textContent = kpis.out_of_stock_count;

    document.getElementById("countAll").textContent = kpis.total_skus;
    document.getElementById("countInStock").textContent = kpis.in_stock_count;
    document.getElementById("countLowStock").textContent = kpis.low_stock_count;
    document.getElementById("countOutOfStock").textContent = kpis.out_of_stock_count;
}

function checkReorderThresholds(products) {
    const needingRestock = products.filter(p => p.stock_quantity <= p.min_threshold);
    const badge = document.getElementById("reorderBadge");
    if (needingRestock.length > 0) {
        badge.textContent = needingRestock.length;
        badge.classList.remove("hidden");
    } else {
        badge.classList.add("hidden");
    }
}

// ----------------- TABLE RENDERING & FILTERING -----------------
function debounceSearch() {
    clearTimeout(searchDebounceTimer);
    const val = document.getElementById("searchInput").value;
    document.getElementById("clearSearchBtn").classList.toggle("hidden", !val);
    searchDebounceTimer = setTimeout(applyFilters, 250);
}

function clearSearch() {
    document.getElementById("searchInput").value = "";
    document.getElementById("clearSearchBtn").classList.add("hidden");
    applyFilters();
}

function clearAllFilters() {
    document.getElementById("searchInput").value = "";
    document.getElementById("clearSearchBtn").classList.add("hidden");
    document.getElementById("categoryFilter").value = "All";
    document.getElementById("brandFilter").value = "All";
    document.getElementById("sortBy").value = "sku-asc";
    setFilterStatus("All");
}

function setFilterStatus(status) {
    currentFilterStatus = status;
    const pills = [
        { id: "pillStatusAll", key: "All" },
        { id: "pillStatusInStock", key: "In Stock" },
        { id: "pillStatusLowStock", key: "Low Stock" },
        { id: "pillStatusOutOfStock", key: "Out of Stock" }
    ];

    pills.forEach(p => {
        const el = document.getElementById(p.id);
        if (el) {
            if (p.key === status) {
                el.className = "px-2.5 py-1 rounded-lg font-semibold bg-blue-600 text-white transition";
            } else {
                el.className = "px-2.5 py-1 rounded-lg font-semibold bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-700 transition";
            }
        }
    });

    applyFilters();
}

function applyFilters() {
    const searchVal = document.getElementById("searchInput").value.toLowerCase().trim();
    const categoryVal = document.getElementById("categoryFilter").value;
    const brandVal = document.getElementById("brandFilter").value;
    const sortVal = document.getElementById("sortBy").value;

    let filtered = [...allProducts];

    // Status filter
    if (currentFilterStatus !== "All") {
        filtered = filtered.filter(p => p.status === currentFilterStatus);
    }

    // Category filter
    if (categoryVal !== "All") {
        filtered = filtered.filter(p => p.category.toLowerCase() === categoryVal.toLowerCase());
    }

    // Brand filter
    if (brandVal !== "All") {
        filtered = filtered.filter(p => p.brand.toLowerCase() === brandVal.toLowerCase());
    }

    // Search filter
    if (searchVal) {
        filtered = filtered.filter(p => 
            p.sku.toLowerCase().includes(searchVal) ||
            p.product_name.toLowerCase().includes(searchVal) ||
            p.model_number.toLowerCase().includes(searchVal) ||
            p.brand.toLowerCase().includes(searchVal) ||
            p.warehouse_location.toLowerCase().includes(searchVal)
        );
    }

    // Sorting
    filtered.sort((a, b) => {
        if (sortVal === "sku-asc") return a.sku.localeCompare(b.sku);
        if (sortVal === "name-asc") return a.product_name.localeCompare(b.product_name);
        if (sortVal === "stock-asc") return a.stock_quantity - b.stock_quantity;
        if (sortVal === "stock-desc") return b.stock_quantity - a.stock_quantity;
        if (sortVal === "price-asc") return a.selling_price - b.selling_price;
        if (sortVal === "price-desc") return b.selling_price - a.selling_price;
        return 0;
    });

    renderTable(filtered);
}

function getCategoryIcon(cat) {
    const c = (cat || "").toLowerCase();
    if (c.includes("tv") || c.includes("television")) return "fa-tv text-purple-500";
    if (c.includes("refrigerator") || c.includes("fridge")) return "fa-snowflake text-sky-500";
    if (c.includes("wash")) return "fa-soap text-indigo-500";
    if (c.includes("air cond") || c.includes("ac")) return "fa-wind text-teal-500";
    if (c.includes("kitchen")) return "fa-kitchen-set text-orange-500";
    if (c.includes("clean")) return "fa-broom text-emerald-500";
    return "fa-plug text-blue-500";
}

function renderTable(products) {
    const tbody = document.getElementById("stockTableBody");
    const emptyState = document.getElementById("emptyState");
    const visibleCount = document.getElementById("visibleCount");

    visibleCount.textContent = products.length;

    if (products.length === 0) {
        tbody.innerHTML = "";
        emptyState.classList.remove("hidden");
        return;
    }
    emptyState.classList.add("hidden");

    let html = "";
    products.forEach(p => {
        const catIcon = getCategoryIcon(p.category);
        
        // Stock progress percentage relative to 2x threshold
        const targetCapacity = Math.max(p.min_threshold * 2, 10);
        const stockPct = Math.min(100, Math.round((p.stock_quantity / targetCapacity) * 100));
        
        let statusBadge = "";
        let progressBarColor = "bg-emerald-500";
        if (p.status === "In Stock") {
            statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 dark:bg-emerald-950/80 text-emerald-700 dark:text-emerald-400 border border-emerald-300 dark:border-emerald-800">In Stock</span>`;
            progressBarColor = "bg-emerald-500";
        } else if (p.status === "Low Stock") {
            statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 dark:bg-amber-950/80 text-amber-700 dark:text-amber-400 border border-amber-300 dark:border-amber-800 animate-pulse">Low Stock</span>`;
            progressBarColor = "bg-amber-500";
        } else {
            statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-red-100 dark:bg-red-950/80 text-red-700 dark:text-red-400 border border-red-300 dark:border-red-800">Out of Stock</span>`;
            progressBarColor = "bg-red-500";
        }

        const margin = p.selling_price - p.cost_price;
        const marginPct = p.selling_price > 0 ? Math.round((margin / p.selling_price) * 100) : 0;

        html += `
        <tr class="hover:bg-slate-50/80 dark:hover:bg-slate-800/40 transition group border-b border-slate-100 dark:border-slate-800">
            <!-- SKU & Model -->
            <td class="py-3 px-4">
                <div class="font-mono font-bold text-xs text-blue-600 dark:text-blue-400">${p.sku}</div>
                <div class="text-[10px] text-slate-400 font-mono">${p.model_number || 'N/A'}</div>
            </td>

            <!-- Appliance Name & Details -->
            <td class="py-3 px-4 max-w-xs">
                <div class="font-semibold text-xs text-slate-900 dark:text-slate-100 line-clamp-1" title="${p.product_name}">${p.product_name}</div>
                <div class="flex items-center gap-2 mt-0.5 text-[10px] text-slate-400">
                    <span class="font-bold text-slate-600 dark:text-slate-300">${p.brand}</span>
                    <span>&bull;</span>
                    <span>${p.energy_rating !== 'N/A' ? '⭐ ' + p.energy_rating : 'Warranty: ' + p.warranty}</span>
                </div>
            </td>

            <!-- Category -->
            <td class="py-3 px-4 whitespace-nowrap">
                <span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-md bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 text-[11px] font-medium">
                    <i class="fa-solid ${catIcon} text-xs"></i>
                    <span>${p.category}</span>
                </span>
            </td>

            <!-- Location -->
            <td class="py-3 px-4 whitespace-nowrap">
                <span class="text-xs font-mono px-2 py-0.5 rounded bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300">
                    <i class="fa-solid fa-location-dot text-[10px] text-slate-400 mr-1"></i>${p.warehouse_location}
                </span>
            </td>

            <!-- Stock Quantity + Progress -->
            <td class="py-3 px-4 text-center">
                <div class="inline-block text-center">
                    <span class="text-sm font-extrabold ${p.stock_quantity === 0 ? 'text-red-600' : (p.stock_quantity <= p.min_threshold ? 'text-amber-600' : 'text-slate-800 dark:text-white')}">
                        ${p.stock_quantity}
                    </span>
                    <span class="text-[10px] text-slate-400 ml-0.5">/ min ${p.min_threshold}</span>
                    <div class="w-16 h-1.5 bg-slate-200 dark:bg-slate-700 rounded-full overflow-hidden mt-1 mx-auto">
                        <div class="h-full ${progressBarColor} rounded-full" style="width: ${stockPct}%"></div>
                    </div>
                </div>
            </td>

            <!-- Quick Adjuster (+1 / -1) -->
            <td class="py-3 px-4 text-center no-print">
                <div class="inline-flex items-center gap-1 bg-slate-100 dark:bg-slate-800/80 p-1 rounded-xl border border-slate-200 dark:border-slate-700/60 shadow-xs">
                    <button onclick="quickAdjustStock('${p.sku}', -1)" title="Decrement stock by 1"
                        class="w-6 h-6 flex items-center justify-center rounded-lg bg-white dark:bg-slate-700 hover:bg-red-50 hover:text-red-600 text-slate-700 dark:text-slate-200 text-xs font-bold shadow-xs transition active:scale-90">
                        <i class="fa-solid fa-minus text-[10px]"></i>
                    </button>
                    <button onclick="quickAdjustStock('${p.sku}', 1)" title="Increment stock by 1"
                        class="w-6 h-6 flex items-center justify-center rounded-lg bg-white dark:bg-slate-700 hover:bg-emerald-50 hover:text-emerald-600 text-slate-700 dark:text-slate-200 text-xs font-bold shadow-xs transition active:scale-90">
                        <i class="fa-solid fa-plus text-[10px]"></i>
                    </button>
                    <button onclick="promptCustomAdjust('${p.sku}', ${p.stock_quantity})" title="Custom batch adjustment"
                        class="px-1.5 h-6 flex items-center justify-center rounded-lg bg-white dark:bg-slate-700 hover:bg-blue-50 hover:text-blue-600 text-slate-500 dark:text-slate-300 text-[10px] font-semibold shadow-xs transition">
                        ±
                    </button>
                </div>
            </td>

            <!-- Pricing & Margin -->
            <td class="py-3 px-4 text-right whitespace-nowrap">
                <div class="font-bold text-xs text-slate-900 dark:text-white">${formatCurrency(p.selling_price)}</div>
                <div class="text-[10px] text-slate-400">Cost: ${formatCurrency(p.cost_price)} <span class="text-emerald-600 font-semibold">(+${marginPct}%)</span></div>
            </td>

            <!-- Status Pill -->
            <td class="py-3 px-4 text-center whitespace-nowrap">
                ${statusBadge}
            </td>

            <!-- Actions -->
            <td class="py-3 px-4 text-center whitespace-nowrap no-print">
                <div class="flex items-center justify-center gap-1">
                    ${hasPermission('stock_edit') ? `
                    <button onclick="openEditModal('${p.sku}')" title="Edit appliance details"
                        class="p-1.5 rounded-lg text-slate-500 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-slate-800 transition">
                        <i class="fa-solid fa-pen-to-square text-xs"></i>
                    </button>` : ''}

                    <button onclick="openBarcodeModal('${p.sku}')" title="Print warehouse barcode label"
                        class="p-1.5 rounded-lg text-slate-500 hover:text-purple-600 hover:bg-purple-50 dark:hover:bg-slate-800 transition">
                        <i class="fa-solid fa-barcode text-xs"></i>
                    </button>

                    ${hasPermission('stock_delete') ? `
                    <button onclick="confirmDelete('${p.sku}')" title="Delete from Excel (Admin Only)"
                        class="p-1.5 rounded-lg text-slate-500 hover:text-red-600 hover:bg-red-50 dark:hover:bg-slate-800 transition">
                        <i class="fa-regular fa-trash-can text-xs"></i>
                    </button>` : ''}
                </div>
            </td>
        </tr>
        `;
    });

    tbody.innerHTML = html;
}

// ----------------- STOCK ADJUSTMENTS -----------------
async function quickAdjustStock(sku, delta) {
    if (!currentUser) return showToast("Please login first.", "error");

    try {
        const res = await fetch(`/api/stock/${encodeURIComponent(sku)}/adjust`, {
            method: "POST",
            headers: getAuthHeaders(),
            body: JSON.stringify({ delta, reason: delta > 0 ? "Quick Restock (+1)" : "Quick Sale / Dispatch (-1)" })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Failed to adjust stock in Excel");
        }

        const data = await res.json();
        
        // Optimistically update local array
        const idx = allProducts.findIndex(p => p.sku.toUpperCase() === sku.toUpperCase());
        if (idx !== -1) {
            allProducts[idx] = data.product;
        }

        applyFilters();
        showToast(data.message, "success");

        // Quiet background refresh for KPIs
        fetchKPIsOnly();
    } catch (err) {
        showToast(err.message, "error");
    }
}

async function promptCustomAdjust(sku, currentQty) {
    const p = allProducts.find(item => item.sku === sku);
    const input = prompt(`Enter stock adjustment for [${sku}] ${p ? p.product_name : ''}\nUse +number to add stock (e.g. +10) or -number to reduce (e.g. -5):`, "+5");
    if (!input) return;

    const delta = parseInt(input);
    if (isNaN(delta) || delta === 0) {
        showToast("Please enter a valid non-zero number.", "warning");
        return;
    }

    try {
        const res = await fetch(`/api/stock/${encodeURIComponent(sku)}/adjust`, {
            method: "POST",
            headers: getAuthHeaders(),
            body: JSON.stringify({ delta, reason: `Batch Adjustment by ${currentUser.name}` })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Failed to update stock");
        }

        const data = await res.json();
        showToast(data.message, "success");
        refreshData();
    } catch (err) {
        showToast(err.message, "error");
    }
}

async function fetchKPIsOnly() {
    try {
        const res = await fetch("/api/stock");
        if (res.ok) {
            const data = await res.json();
            updateKPIs(data.kpis);
            checkReorderThresholds(data.products);
        }
    } catch (e) {}
}

// ----------------- EDIT APPLIANCE MODAL -----------------
function openEditModal(sku) {
    const product = allProducts.find(p => p.sku.toUpperCase() === sku.toUpperCase());
    if (!product) return;

    document.getElementById("editSku").value = product.sku;
    document.getElementById("editName").value = product.product_name;
    document.getElementById("editCategory").value = product.category;
    document.getElementById("editBrand").value = product.brand;
    document.getElementById("editModel").value = product.model_number;
    document.getElementById("editLocation").value = product.warehouse_location;
    document.getElementById("editStockQty").value = product.stock_quantity;
    document.getElementById("editMinThresh").value = product.min_threshold;
    document.getElementById("editCostPrice").value = product.cost_price;
    document.getElementById("editSellingPrice").value = product.selling_price;
    document.getElementById("editEnergy").value = product.energy_rating || "N/A";
    document.getElementById("editSupplier").value = product.supplier;

    document.getElementById("editModal").classList.remove("hidden");
}

function closeEditModal() {
    document.getElementById("editModal").classList.add("hidden");
}

async function handleEditSubmit(e) {
    e.preventDefault();
    const sku = document.getElementById("editSku").value;
    const saveBtn = document.getElementById("editSaveBtn");

    const payload = {
        product_name: document.getElementById("editName").value.trim(),
        category: document.getElementById("editCategory").value,
        brand: document.getElementById("editBrand").value.trim(),
        model_number: document.getElementById("editModel").value.trim(),
        warehouse_location: document.getElementById("editLocation").value.trim(),
        stock_quantity: parseInt(document.getElementById("editStockQty").value),
        min_threshold: parseInt(document.getElementById("editMinThresh").value),
        cost_price: parseFloat(document.getElementById("editCostPrice").value),
        selling_price: parseFloat(document.getElementById("editSellingPrice").value),
        energy_rating: document.getElementById("editEnergy").value,
        supplier: document.getElementById("editSupplier").value.trim()
    };

    try {
        saveBtn.disabled = true;
        saveBtn.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin"></i> Saving to Excel...`;

        const res = await fetch(`/api/stock/${encodeURIComponent(sku)}`, {
            method: "PUT",
            headers: getAuthHeaders(),
            body: JSON.stringify(payload)
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Failed to save edits to Excel");
        }

        closeEditModal();
        showToast(`Appliance [${sku}] updated and saved to Excel!`, "success");
        refreshData();
    } catch (err) {
        showToast(err.message, "error");
    } finally {
        saveBtn.disabled = false;
        saveBtn.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> <span>Save to Excel File</span>`;
    }
}

// ----------------- ADD APPLIANCE MODAL -----------------
function openAddModal() {
    suggestNextSku();
    document.getElementById("addModal").classList.remove("hidden");
}

function closeAddModal() {
    document.getElementById("addModal").classList.add("hidden");
}

function suggestNextSku() {
    const cat = document.getElementById("addCategory").value;
    let prefix = "XYZ-GEN";
    if (cat === "Televisions") prefix = "XYZ-TV";
    else if (cat === "Refrigerators") prefix = "XYZ-REF";
    else if (cat === "Washing Machines") prefix = "XYZ-WM";
    else if (cat === "Air Conditioners") prefix = "XYZ-AC";
    else if (cat === "Kitchen Appliances") prefix = "XYZ-KIT";
    else if (cat === "Home Cleaning") prefix = "XYZ-CLN";

    // Find highest existing number for this prefix
    let maxNum = 0;
    allProducts.forEach(p => {
        if (p.sku.startsWith(prefix)) {
            const parts = p.sku.split("-");
            const num = parseInt(parts[parts.length - 1]);
            if (!isNaN(num) && num > maxNum) maxNum = num;
        }
    });

    const nextNum = maxNum > 0 ? maxNum + 1 : 101;
    document.getElementById("addSku").value = `${prefix}-${String(nextNum).padStart(3, '0')}`;
}

async function handleAddSubmit(e) {
    e.preventDefault();
    const saveBtn = document.getElementById("addSaveBtn");

    const payload = {
        sku: document.getElementById("addSku").value.trim().toUpperCase(),
        product_name: document.getElementById("addName").value.trim(),
        category: document.getElementById("addCategory").value,
        brand: document.getElementById("addBrand").value.trim(),
        model_number: document.getElementById("addModel").value.trim(),
        stock_quantity: parseInt(document.getElementById("addStockQty").value),
        min_threshold: parseInt(document.getElementById("addMinThresh").value),
        cost_price: parseFloat(document.getElementById("addCostPrice").value),
        selling_price: parseFloat(document.getElementById("addSellingPrice").value),
        warehouse_location: document.getElementById("addLocation").value.trim(),
        energy_rating: document.getElementById("addEnergy").value,
        supplier: document.getElementById("addSupplier").value.trim(),
        warranty: document.getElementById("addWarranty").value.trim()
    };

    try {
        saveBtn.disabled = true;
        saveBtn.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin"></i> Writing to Excel...`;

        const res = await fetch("/api/stock", {
            method: "POST",
            headers: getAuthHeaders(),
            body: JSON.stringify(payload)
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Could not add new appliance");
        }

        closeAddModal();
        showToast(`Appliance [${payload.sku}] added and saved into Excel spreadsheet!`, "success");
        refreshData();
    } catch (err) {
        showToast(err.message, "error");
    } finally {
        saveBtn.disabled = false;
        saveBtn.innerHTML = `<i class="fa-solid fa-plus"></i> <span>Add & Save to Excel</span>`;
    }
}

// ----------------- DELETE APPLIANCE -----------------
async function confirmDelete(sku) {
    const p = allProducts.find(item => item.sku === sku);
    const confirmed = confirm(`Are you sure you want to permanently delete:\n\n[${sku}] ${p ? p.product_name : ''}\n\nThis will remove the row directly from the Excel sheet.`);
    if (!confirmed) return;

    try {
        const res = await fetch(`/api/stock/${encodeURIComponent(sku)}`, {
            method: "DELETE",
            headers: getAuthHeaders()
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Delete failed");
        }

        showToast(`Appliance SKU ${sku} deleted and updated in Excel.`, "info");
        refreshData();
    } catch (err) {
        showToast(err.message, "error");
    }
}

// ----------------- BARCODE MODAL -----------------
function openBarcodeModal(sku) {
    const p = allProducts.find(item => item.sku === sku);
    if (!p) return;

    document.getElementById("barcodeSku").textContent = p.sku;
    document.getElementById("barcodeProdName").textContent = p.product_name;
    document.getElementById("barcodeLocation").textContent = p.warehouse_location;
    document.getElementById("barcodePrice").textContent = formatCurrency(p.selling_price);

    // Generate pseudo barcode bars
    const barcodeDiv = document.getElementById("svgBarcode");
    barcodeDiv.innerHTML = "";
    
    // Hash string to pseudo pattern
    const pattern = (sku + "1098472").split("").map(c => (c.charCodeAt(0) % 5) + 1);
    for (let i = 0; i < 48; i++) {
        const bar = document.createElement("div");
        const width = pattern[i % pattern.length];
        bar.className = (i % 2 === 0) ? "bg-slate-900 h-full" : "bg-transparent h-full";
        bar.style.width = `${width * 2}px`;
        barcodeDiv.appendChild(bar);
    }

    document.getElementById("barcodeModal").classList.remove("hidden");
}

function closeBarcodeModal() {
    document.getElementById("barcodeModal").classList.add("hidden");
}

// ----------------- REORDER PURCHASE ORDER ASSISTANT -----------------
async function openReorderModal() {
    try {
        const res = await fetch("/api/reorder-recommendations", {
            headers: getAuthHeaders()
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.detail || "Could not compute reorder recommendations");
        }
        
        const data = await res.json();
        document.getElementById("reorderItemCount").textContent = data.total_items_to_reorder;
        document.getElementById("reorderTotalCost").textContent = formatCurrency(data.total_estimated_po_cost);

        const tbody = document.getElementById("reorderTableBody");
        if (data.items.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" class="py-8 text-center text-slate-400">🎉 Great news! All appliances are currently above their minimum safety thresholds.</td></tr>`;
        } else {
            let html = "";
            data.items.forEach(item => {
                const urgencyBadge = item.urgency === "CRITICAL"
                    ? `<span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-red-100 text-red-700">CRITICAL (0 stock)</span>`
                    : `<span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-amber-100 text-amber-700">LOW</span>`;

                html += `
                <tr class="hover:bg-slate-50 dark:hover:bg-slate-800/40">
                    <td class="py-2 px-3 font-mono font-bold text-blue-600">${item.sku}</td>
                    <td class="py-2 px-3">
                        <div class="font-medium">${item.product_name}</div>
                        <div class="text-[10px] text-slate-400">${item.brand} &bull; ${urgencyBadge}</div>
                    </td>
                    <td class="py-2 px-3 text-center font-bold">
                        <span class="${item.current_stock === 0 ? 'text-red-600' : 'text-amber-600'}">${item.current_stock}</span> / ${item.min_threshold}
                    </td>
                    <td class="py-2 px-3 text-center">
                        <span class="px-2 py-0.5 rounded-md bg-blue-100 text-blue-800 font-extrabold text-xs">+${item.suggested_order_qty} Units</span>
                    </td>
                    <td class="py-2 px-3 text-slate-500">${item.supplier}</td>
                    <td class="py-2 px-3 text-right">${formatCurrency(item.unit_cost)}</td>
                    <td class="py-2 px-3 text-right font-bold text-emerald-600">${formatCurrency(item.estimated_po_cost)}</td>
                </tr>
                `;
            });
            tbody.innerHTML = html;
        }

        document.getElementById("reorderModal").classList.remove("hidden");
    } catch (err) {
        showToast(err.message, "error");
    }
}

function closeReorderModal() {
    document.getElementById("reorderModal").classList.add("hidden");
}

function printReorderSheet() {
    const dateEl = document.getElementById("printPoDate");
    if (dateEl) dateEl.textContent = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    const userEl = document.getElementById("printPoUser");
    if (userEl && currentUser) userEl.textContent = `${currentUser.name} (${currentUser.role})`;

    document.body.classList.add("printing-reorder");
    window.print();
    setTimeout(() => {
        document.body.classList.remove("printing-reorder");
    }, 1200);
}

// ----------------- EXCEL UPLOAD MODAL -----------------
function openUploadModal() {
    document.getElementById("uploadModal").classList.remove("hidden");
}

function closeUploadModal() {
    document.getElementById("uploadModal").classList.add("hidden");
}

function updateFileNamePreview(input) {
    const fileLabel = document.getElementById("fileSelectedName");
    if (input.files && input.files[0]) {
        fileLabel.textContent = `Selected: ${input.files[0].name} (${Math.round(input.files[0].size / 1024)} KB)`;
        fileLabel.classList.remove("hidden");
    } else {
        fileLabel.classList.add("hidden");
    }
}

async function handleExcelUpload(e) {
    e.preventDefault();
    const fileInput = document.getElementById("excelFileInput");
    const uploadBtn = document.getElementById("uploadSubmitBtn");

    if (!fileInput.files || fileInput.files.length === 0) {
        showToast("Please choose an Excel (.xlsx) file first.", "warning");
        return;
    }

    const formData = new FormData();
    formData.append("file", fileInput.files[0]);

    try {
        uploadBtn.disabled = true;
        uploadBtn.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin"></i> Processing Spreadsheet...`;

        const res = await fetch("/api/upload-excel", {
            method: "POST",
            headers: { "Authorization": currentUser ? `Bearer ${currentUser.token}` : "" },
            body: formData
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "Failed to upload Excel spreadsheet");
        }

        const data = await res.json();
        closeUploadModal();
        showToast(data.message, "success");
        refreshData();
    } catch (err) {
        showToast(err.message, "error");
    } finally {
        uploadBtn.disabled = false;
        uploadBtn.innerHTML = `<i class="fa-solid fa-upload"></i> <span>Upload & Import</span>`;
    }
}

// ----------------- AUDIT TRAIL DRAWER -----------------
async function toggleAuditDrawer() {
    const drawer = document.getElementById("auditDrawer");
    const isHidden = drawer.classList.contains("hidden");

    if (isHidden) {
        drawer.classList.remove("hidden");
        loadAuditLogs();
    } else {
        drawer.classList.add("hidden");
    }
}

async function loadAuditLogs() {
    const list = document.getElementById("auditList");
    list.innerHTML = `<div class="text-center py-6 text-slate-400"><i class="fa-solid fa-circle-notch fa-spin text-lg"></i></div>`;

    try {
        const res = await fetch("/api/audit-log", {
            headers: getAuthHeaders()
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.detail || "Could not fetch audit history");
        }
        
        const logs = await res.json();
        if (logs.length === 0) {
            list.innerHTML = `<div class="text-center py-8 text-xs text-slate-400">No actions recorded in audit log yet.</div>`;
            return;
        }

        let html = "";
        logs.forEach(log => {
            let icon = "fa-pen-to-square text-blue-500";
            if (log.action === "STOCK_ADJUSTMENT") icon = "fa-boxes-stacked text-amber-500";
            if (log.action === "ADD_APPLIANCE") icon = "fa-plus text-emerald-500";
            if (log.action === "DELETE_APPLIANCE") icon = "fa-trash-can text-red-500";
            if (log.action === "EXCEL_IMPORT") icon = "fa-file-excel text-emerald-600";
            if (log.action === "USER_LOGIN") icon = "fa-arrow-right-to-bracket text-indigo-500";

            html += `
            <div class="p-3 rounded-xl bg-slate-50 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-800 text-xs space-y-1">
                <div class="flex items-center justify-between text-[10px] text-slate-400">
                    <span class="font-semibold text-slate-700 dark:text-slate-300 flex items-center gap-1.5">
                        <i class="fa-solid ${icon}"></i> ${log.action.replace('_', ' ')}
                    </span>
                    <span>${log.timestamp}</span>
                </div>
                <div class="font-medium text-slate-800 dark:text-slate-100">${log.details}</div>
                <div class="text-[10px] text-slate-500 flex items-center justify-between pt-1">
                    <span>By: <strong class="text-slate-700 dark:text-slate-300">${log.user_name}</strong> (${log.user_role})</span>
                    <span class="font-mono text-blue-600 dark:text-blue-400">${log.sku}</span>
                </div>
            </div>
            `;
        });

        list.innerHTML = html;
    } catch (err) {
        list.innerHTML = `<div class="text-center py-4 text-xs text-red-500">${err.message}</div>`;
    }
}

// ----------------- CHARTS & ANALYTICS -----------------
function toggleAnalytics() {
    const panel = document.getElementById("analyticsPanel");
    const chevron = document.getElementById("analyticsChevron");
    const isHidden = panel.classList.contains("hidden");

    if (isHidden) {
        panel.classList.remove("hidden");
        chevron.style.transform = "rotate(180deg)";
        renderCharts(allProducts);
    } else {
        panel.classList.add("hidden");
        chevron.style.transform = "rotate(0deg)";
    }
}

function renderCharts(products) {
    if (!products || products.length === 0) return;
    const isDark = document.documentElement.classList.contains("dark");
    const textColor = isDark ? "#cbd5e1" : "#475569";
    const gridColor = isDark ? "rgba(255,255,255,0.08)" : "rgba(0,0,0,0.06)";

    // 1. Category Chart
    const categories = {};
    products.forEach(p => {
        categories[p.category] = (categories[p.category] || 0) + p.stock_quantity;
    });

    const catCtx = document.getElementById("categoryChart");
    if (catCtx) {
        if (categoryChartInstance) categoryChartInstance.destroy();

        categoryChartInstance = new Chart(catCtx, {
            type: 'doughnut',
            data: {
                labels: Object.keys(categories),
                datasets: [{
                    data: Object.values(categories),
                    backgroundColor: [
                        '#6366f1', '#3b82f6', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'
                    ],
                    borderWidth: 2,
                    borderColor: isDark ? '#0f172a' : '#ffffff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: { boxWidth: 12, font: { size: 10 }, color: textColor }
                    }
                }
            }
        });
    }

    // 2. Brand Chart
    const brands = {};
    products.forEach(p => {
        brands[p.brand] = (brands[p.brand] || 0) + p.stock_quantity;
    });

    const brandLabels = Object.keys(brands).sort((a, b) => brands[b] - brands[a]).slice(0, 7);
    const brandData = brandLabels.map(l => brands[l]);

    const brandCtx = document.getElementById("brandChart");
    if (brandCtx) {
        if (brandChartInstance) brandChartInstance.destroy();

        brandChartInstance = new Chart(brandCtx, {
            type: 'bar',
            data: {
                labels: brandLabels,
                datasets: [{
                    label: 'Units In Stock',
                    data: brandData,
                    backgroundColor: '#3b82f6',
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y',
                scales: {
                    x: {
                        grid: { color: gridColor },
                        ticks: { color: textColor, font: { size: 10 } }
                    },
                    y: {
                        grid: { display: false },
                        ticks: { color: textColor, font: { size: 10 } }
                    }
                },
                plugins: {
                    legend: { display: false }
                }
            }
        });
    }
}

// ----------------- TOAST NOTIFICATIONS -----------------
function showToast(message, type = "info") {
    const container = document.getElementById("toastContainer");
    if (!container) return;

    const toast = document.createElement("div");
    let bg = "bg-slate-900 border-slate-700 text-white";
    let icon = "fa-circle-info text-blue-400";

    if (type === "success") {
        bg = "bg-emerald-950 border-emerald-800 text-emerald-100";
        icon = "fa-circle-check text-emerald-400";
    } else if (type === "error") {
        bg = "bg-red-950 border-red-800 text-red-100";
        icon = "fa-circle-exclamation text-red-400";
    } else if (type === "warning") {
        bg = "bg-amber-950 border-amber-800 text-amber-100";
        icon = "fa-triangle-exclamation text-amber-400";
    }

    toast.className = `flex items-center gap-3 px-4 py-3 rounded-2xl shadow-xl border text-xs font-medium pointer-events-auto transition transform duration-300 translate-y-2 opacity-0 ${bg}`;
    toast.innerHTML = `
        <i class="fa-solid ${icon} text-base flex-shrink-0"></i>
        <div class="flex-1">${message}</div>
        <button onclick="this.parentElement.remove()" class="text-slate-400 hover:text-white ml-2"><i class="fa-solid fa-xmark"></i></button>
    `;

    container.appendChild(toast);

    // Animate in
    requestAnimationFrame(() => {
        toast.classList.remove("translate-y-2", "opacity-0");
    });

    // Auto dismiss after 4 seconds
    setTimeout(() => {
        toast.classList.add("translate-y-2", "opacity-0");
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}
