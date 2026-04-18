/* ================================================================
   DemandFlow — Application Logic
   Modules: Auth, Data Input (Sales + Stock), Forecasting, KPIs,
            Demand Drivers, Inventory Planning, What-If Analysis,
            Dashboard Charts
   ================================================================ */

// ────────────────────────────────────────────
// ░░  GLOBAL STATE
// ────────────────────────────────────────────
const state = {
    salesData: [],        // Array of { sku, productName, date, soldQty }
    stockData: [],        // Array of { sku, productName, qty }
    rawData: [],          // Merged: Array of { sku, productName, date, sales, stock }
    productNameMap: {},   // SKU → Product Name lookup
    forecasts: [],        // Array of { sku, productName, last3, forecast, method }
    adjustedForecasts: [], // after demand driver application
    inventoryStatus: [],  // Array of { sku, productName, avgDemand, stock, reorderPt, status }
    forecastMethod: 'ma', // 'ma' | 'wa'
    promotion: 0,         // 0–100 percent
    seasonality: 1.0,     // 0.5–2.0 multiplier
    leadTime: 7,          // days
    whatIfPct: 0,          // -50 to +50
    currentUser: null,     // { email, role } — set after login
};

// Chart instances
let salesForecastChart = null;
let stockLevelsChart = null;
let whatIfChart = null;

// ────────────────────────────────────────────
// ░░  AUTH / LOGIN SYSTEM
// ────────────────────────────────────────────

// Since the app is hosted statically without a database, user credentials
// must be hardcoded here to be accessible across all devices.
const REGISTERED_USERS = [
    { email: 'admin', password: 'admin123', role: 'admin' },
    // Add additional users here, for example:
    { email: 'user', password: 'user123', role: 'user' }
];

/**
 * Initialise the user store
 */
function initAuthStore() {
    if (!localStorage.getItem('df_users_local')) {
        localStorage.setItem('df_users_local', JSON.stringify([]));
    }
}

/**
 * Get all registered users (combines static and local-only)
 */
function getUsers() {
    const localUsers = JSON.parse(localStorage.getItem('df_users_local') || '[]');
    const allUsers = [...REGISTERED_USERS];
    
    localUsers.forEach(lu => {
        if (!allUsers.find(u => u.email.toLowerCase() === lu.email.toLowerCase())) {
            allUsers.push(lu);
        }
    });
    return allUsers;
}

/**
 * Save users array
 */
function saveUsers(users) {
    // Only save users to local storage that are NOT in the REGISTERED_USERS static array
    const localOnly = users.filter(u => 
        !REGISTERED_USERS.find(ru => ru.email.toLowerCase() === u.email.toLowerCase())
    );
    localStorage.setItem('df_users_local', JSON.stringify(localOnly));
}

/**
 * Authenticate a user by email/username and password
 */
function authenticateUser(email, password) {
    const users = getUsers();
    return users.find(u => u.email.toLowerCase() === email.toLowerCase() && u.password === password) || null;
}

/**
 * Add a new user (admin only)
 */
function addUser(email, password, role = 'user') {
    const users = getUsers();
    if (users.find(u => u.email.toLowerCase() === email.toLowerCase())) {
        return { success: false, message: 'A user with this email already exists.' };
    }
    users.push({ email, password, role });
    saveUsers(users);
    return { success: true, message: `User "${email}" added for THIS device only.` };
}

/**
 * Remove a user (admin only)
 */
function removeUser(email) {
    // Prevent removing static users
    if (REGISTERED_USERS.find(u => u.email.toLowerCase() === email.toLowerCase())) {
        return { success: false, message: 'Cannot remove a global hardcoded user from the web interface.' };
    }

    const users = getUsers();
    const idx = users.findIndex(u => u.email.toLowerCase() === email.toLowerCase());
    if (idx === -1) return { success: false, message: 'User not found.' };
    
    users.splice(idx, 1);
    saveUsers(users);
    return { success: true, message: `User "${email}" removed.` };
}

/**
 * Show login overlay
 */
function showLoginScreen() {
    const overlay = document.getElementById('loginOverlay');
    overlay.classList.add('visible');
    document.getElementById('loginEmail').focus();
}

/**
 * Hide login overlay and boot the app
 */
function hideLoginScreen() {
    const overlay = document.getElementById('loginOverlay');
    overlay.classList.remove('visible');
}

/**
 * Handle login form submit
 */
function handleLogin(e) {
    e.preventDefault();
    const email = document.getElementById('loginEmail').value.trim();
    const password = document.getElementById('loginPassword').value;
    const errorEl = document.getElementById('loginError');

    if (!email || !password) {
        errorEl.textContent = 'Please enter both email/username and password.';
        errorEl.style.display = 'block';
        return;
    }

    const user = authenticateUser(email, password);
    if (!user) {
        errorEl.textContent = 'Invalid credentials. Please try again.';
        errorEl.style.display = 'block';
        return;
    }

    errorEl.style.display = 'none';
    state.currentUser = { email: user.email, role: user.role };
    sessionStorage.setItem('df_session', JSON.stringify(state.currentUser));
    hideLoginScreen();
    updateUserDisplay();
    toast(`Welcome, ${user.email}!`, 'success');
}

/**
 * Logout
 */
function handleLogout() {
    state.currentUser = null;
    sessionStorage.removeItem('df_session');
    showLoginScreen();
}

/**
 * Update user display in topbar
 */
function updateUserDisplay() {
    const avatar = document.getElementById('topbarUserArea');
    if (!state.currentUser || !avatar) return;

    const isAdmin = state.currentUser.role === 'admin';
    avatar.innerHTML = `
        <div class="user-info">
            <span class="user-name">${state.currentUser.email}</span>
            <span class="user-role-badge ${isAdmin ? 'admin' : ''}">${isAdmin ? 'Admin' : 'User'}</span>
        </div>
        ${isAdmin ? '<button class="btn btn-sm btn-outline" id="manageUsersBtn" title="Manage Users"><i data-lucide="users"></i></button>' : ''}
        <button class="btn btn-sm btn-outline" id="logoutBtn" title="Logout"><i data-lucide="log-out"></i></button>
    `;
    lucide.createIcons();

    document.getElementById('logoutBtn').addEventListener('click', handleLogout);
    if (isAdmin) {
        document.getElementById('manageUsersBtn').addEventListener('click', showUserManagement);
    }
}

/**
 * Show user management modal (admin only)
 */
function showUserManagement() {
    const modal = document.getElementById('userMgmtModal');
    modal.classList.add('visible');
    renderUserList();
}

/**
 * Render the list of users in the modal
 */
function renderUserList() {
    const users = getUsers();
    const tbody = document.getElementById('userListBody');

    tbody.innerHTML = users.map(u => `
        <tr>
            <td style="color:var(--text-primary);font-weight:500;">${u.email}</td>
            <td><span class="user-role-badge ${u.role === 'admin' ? 'admin' : ''}">${u.role === 'admin' ? 'Admin' : 'User'}</span></td>
            <td>
                ${u.email.toLowerCase() === 'admin' ? '<span style="color:var(--text-muted);font-size:0.78rem;">Protected</span>' :
                    `<button class="btn-icon-danger" data-remove-user="${u.email}" aria-label="Remove user"><i data-lucide="trash-2"></i></button>`}
            </td>
        </tr>
    `).join('');

    lucide.createIcons();

    // Bind remove buttons
    tbody.querySelectorAll('[data-remove-user]').forEach(btn => {
        btn.addEventListener('click', () => {
            const result = removeUser(btn.dataset.removeUser);
            toast(result.message, result.success ? 'success' : 'error');
            renderUserList();
        });
    });
}

/**
 * Handle adding a new user via the modal form
 */
function handleAddUser(e) {
    e.preventDefault();
    const emailInput = document.getElementById('newUserEmail');
    const passInput = document.getElementById('newUserPassword');
    const email = emailInput.value.trim();
    const password = passInput.value;

    if (!email || !password) {
        toast('Please enter both email and password.', 'error');
        return;
    }
    if (password.length < 4) {
        toast('Password must be at least 4 characters.', 'error');
        return;
    }

    const result = addUser(email, password, 'user');
    toast(result.message, result.success ? 'success' : 'error');
    if (result.success) {
        emailInput.value = '';
        passInput.value = '';
        renderUserList();
    }
}

// ────────────────────────────────────────────
// ░░  INITIALISATION
// ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initAuthStore();

    // Check for existing session
    const session = sessionStorage.getItem('df_session');
    if (session) {
        state.currentUser = JSON.parse(session);
        hideLoginScreen();
        updateUserDisplay();
    } else {
        showLoginScreen();
    }

    // Wire up login form
    document.getElementById('loginForm').addEventListener('submit', handleLogin);

    // Wire up user management modal
    document.getElementById('closeUserMgmt').addEventListener('click', () => {
        document.getElementById('userMgmtModal').classList.remove('visible');
    });
    document.getElementById('addUserForm').addEventListener('submit', handleAddUser);

    // Boot app modules
    initNavigation();
    initDataInput();
    initForecasting();
    initDemandDrivers();
    initInventory();
    initWhatIf();
    initExport();
});

// ────────────────────────────────────────────
// ░░  NAVIGATION
// ────────────────────────────────────────────
function initNavigation() {
    const links = document.querySelectorAll('.nav-link');
    const sections = document.querySelectorAll('.content-section');
    const pageTitle = document.getElementById('pageTitle');
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.getElementById('sidebar');
    const mobileMenuBtn = document.getElementById('mobileMenuBtn');

    links.forEach(link => {
        link.addEventListener('click', e => {
            e.preventDefault();
            const sectionId = link.dataset.section;

            // Update active link
            links.forEach(l => l.classList.remove('active'));
            link.classList.add('active');

            // Show target section
            sections.forEach(s => s.classList.remove('active'));
            document.getElementById(`section-${sectionId}`).classList.add('active');

            // Update page title
            pageTitle.textContent = link.querySelector('span').textContent;

            // Close mobile menu
            sidebar.classList.remove('mobile-open');
        });
    });

    // Sidebar collapse toggle (desktop)
    sidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('collapsed');
    });

    // Mobile menu toggle
    mobileMenuBtn.addEventListener('click', () => {
        sidebar.classList.toggle('mobile-open');
    });
}

// ────────────────────────────────────────────
// ░░  DATA INPUT MODULE
// ────────────────────────────────────────────
function initDataInput() {
    // --- Sales: Load pasted ---
    document.getElementById('loadSalesPastedBtn').addEventListener('click', () => {
        const text = document.getElementById('salesPasteArea').value.trim();
        if (!text) { toast('Paste sales CSV data first.', 'error'); return; }
        parseSalesCSV(text);
    });

    // --- Sales: Fill sample ---
    document.getElementById('fillSampleSalesBtn').addEventListener('click', () => {
        document.getElementById('salesPasteArea').value = generateSampleSalesCSVText();
        toast('Sample sales CSV filled! Click "Load Sales Data" to import.', 'info');
    });

    // --- Sales: Clear ---
    document.getElementById('clearSalesBtn').addEventListener('click', () => {
        state.salesData = [];
        renderSalesTable();
        updateDataStatus();
        mergeData();
        toast('Sales data cleared.', 'info');
    });

    // --- Stock: Load pasted ---
    document.getElementById('loadStockPastedBtn').addEventListener('click', () => {
        const text = document.getElementById('stockPasteArea').value.trim();
        if (!text) { toast('Paste stock CSV data first.', 'error'); return; }
        parseStockCSV(text);
    });

    // --- Stock: Fill sample ---
    document.getElementById('fillSampleStockBtn').addEventListener('click', () => {
        document.getElementById('stockPasteArea').value = generateSampleStockCSVText();
        toast('Sample stock CSV filled! Click "Load Stock Data" to import.', 'info');
    });

    // --- Stock: Clear ---
    document.getElementById('clearStockBtn').addEventListener('click', () => {
        state.stockData = [];
        renderStockTable();
        updateDataStatus();
        mergeData();
        toast('Stock data cleared.', 'info');
    });

    // --- Load sample data (both) ---
    document.getElementById('loadSampleBtn').addEventListener('click', () => {
        state.salesData = generateSampleSalesData();
        state.stockData = generateSampleStockData();
        renderSalesTable();
        renderStockTable();
        updateDataStatus();
        mergeData();
        toast('Sample sales & stock data loaded!', 'success');
    });
}

/* Tab switching removed — Google Sheet tabs no longer present */

// ────────────────────────────────────────────
// ░░  SALES CSV PARSING
// ────────────────────────────────────────────

/**
 * Parse pasted Sales CSV: columns SKU, Product Name, Date, Sold Qty
 */
function parseSalesCSV(text) {
    const firstLine = text.split('\n')[0];
    const delimiter = firstLine.includes('\t') ? '\t' : ',';
    const lines = text.trim().split('\n');
    if (lines.length < 2) { toast('Sales data must have a header row and at least one data row.', 'error'); return; }

    const headers = lines[0].split(delimiter).map(h => h.trim().toLowerCase().replace(/['"]/g, ''));
    const skuIdx = findHeaderIndex(headers, ['sku']);
    const nameIdx = findHeaderIndex(headers, ['product name', 'productname', 'product_name', 'name', 'product']);
    const dateIdx = findHeaderIndex(headers, ['date']);
    const qtyIdx = findHeaderIndex(headers, ['sold qty', 'soldqty', 'sold_qty', 'sales', 'quantity', 'qty']);

    if (skuIdx === -1 || dateIdx === -1 || qtyIdx === -1) {
        toast(`Missing columns. Found: [${headers.join(', ')}]. Need: SKU, Product Name, Date, Sold Qty.`, 'error');
        return;
    }

    state.salesData = [];
    for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(delimiter).map(c => c.trim().replace(/['"]/g, ''));
        if (cols.length < 3) continue;
        const sku = cols[skuIdx];
        const date = cols[dateIdx];
        if (!sku || !date) continue;
        state.salesData.push({
            sku,
            productName: nameIdx !== -1 ? cols[nameIdx] || '' : '',
            date,
            soldQty: parseFloat(cols[qtyIdx]) || 0,
        });
    }

    if (state.salesData.length === 0) {
        toast('No valid sales rows found. Check your format.', 'error');
        return;
    }

    renderSalesTable();
    updateDataStatus();
    mergeData();
    toast(`Loaded ${state.salesData.length} sales rows!`, 'success');
}

// ────────────────────────────────────────────
// ░░  STOCK CSV PARSING
// ────────────────────────────────────────────

/**
 * Parse pasted Stock CSV: columns SKU, Product Name, Qty
 */
function parseStockCSV(text) {
    const firstLine = text.split('\n')[0];
    const delimiter = firstLine.includes('\t') ? '\t' : ',';
    const lines = text.trim().split('\n');
    if (lines.length < 2) { toast('Stock data must have a header row and at least one data row.', 'error'); return; }

    const headers = lines[0].split(delimiter).map(h => h.trim().toLowerCase().replace(/['"]/g, ''));
    const skuIdx = findHeaderIndex(headers, ['sku']);
    const nameIdx = findHeaderIndex(headers, ['product name', 'productname', 'product_name', 'name', 'product']);
    const qtyIdx = findHeaderIndex(headers, ['qty', 'quantity', 'stock', 'stock qty', 'stockqty', 'stock_qty']);

    if (skuIdx === -1 || qtyIdx === -1) {
        toast(`Missing columns. Found: [${headers.join(', ')}]. Need: SKU, Product Name, Qty.`, 'error');
        return;
    }

    state.stockData = [];
    for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(delimiter).map(c => c.trim().replace(/['"]/g, ''));
        if (cols.length < 2) continue;
        const sku = cols[skuIdx];
        if (!sku) continue;
        state.stockData.push({
            sku,
            productName: nameIdx !== -1 ? cols[nameIdx] || '' : '',
            qty: parseFloat(cols[qtyIdx]) || 0,
        });
    }

    if (state.stockData.length === 0) {
        toast('No valid stock rows found. Check your format.', 'error');
        return;
    }

    renderStockTable();
    updateDataStatus();
    mergeData();
    toast(`Loaded ${state.stockData.length} stock rows!`, 'success');
}

/* Google Sheets integration removed — data input is now paste-only */

// ────────────────────────────────────────────
// ░░  DATA TABLE RENDERING
// ────────────────────────────────────────────

/**
 * Render the sales data table
 */
function renderSalesTable() {
    const card = document.getElementById('salesTableCard');
    const tbody = document.getElementById('salesTableBody');
    const rowCount = document.getElementById('salesRowCount');

    if (state.salesData.length === 0) {
        card.style.display = 'none';
        return;
    }
    card.style.display = '';
    rowCount.textContent = `${state.salesData.length} rows`;

    tbody.innerHTML = state.salesData.map((row, i) => `
        <tr>
            <td>${row.sku}</td>
            <td>${row.productName}</td>
            <td>${row.date}</td>
            <td>${row.soldQty}</td>
            <td><button class="btn-icon-danger" data-delete-sales="${i}" aria-label="Delete row"><i data-lucide="x"></i></button></td>
        </tr>
    `).join('');

    lucide.createIcons();

    // Delete row buttons
    tbody.querySelectorAll('[data-delete-sales]').forEach(btn => {
        btn.addEventListener('click', () => {
            const idx = parseInt(btn.dataset.deleteSales);
            state.salesData.splice(idx, 1);
            renderSalesTable();
            updateDataStatus();
            mergeData();
        });
    });
}

/**
 * Render the stock data table
 */
function renderStockTable() {
    const card = document.getElementById('stockTableCard');
    const tbody = document.getElementById('stockTableBody');
    const rowCount = document.getElementById('stockRowCount');

    if (state.stockData.length === 0) {
        card.style.display = 'none';
        return;
    }
    card.style.display = '';
    rowCount.textContent = `${state.stockData.length} rows`;

    tbody.innerHTML = state.stockData.map((row, i) => `
        <tr>
            <td>${row.sku}</td>
            <td>${row.productName}</td>
            <td>${row.qty}</td>
            <td><button class="btn-icon-danger" data-delete-stock="${i}" aria-label="Delete row"><i data-lucide="x"></i></button></td>
        </tr>
    `).join('');

    lucide.createIcons();

    // Delete row buttons
    tbody.querySelectorAll('[data-delete-stock]').forEach(btn => {
        btn.addEventListener('click', () => {
            const idx = parseInt(btn.dataset.deleteStock);
            state.stockData.splice(idx, 1);
            renderStockTable();
            updateDataStatus();
            mergeData();
        });
    });
}

// ────────────────────────────────────────────
// ░░  DATA STATUS & MERGE
// ────────────────────────────────────────────

/**
 * Update the data status chips
 */
function updateDataStatus() {
    const salesChip = document.getElementById('salesStatusChip');
    const stockChip = document.getElementById('stockStatusChip');

    if (state.salesData.length > 0) {
        salesChip.classList.add('loaded');
        salesChip.innerHTML = `<i data-lucide="check-circle"></i><span>Sales Data: ${state.salesData.length} rows loaded</span>`;
    } else {
        salesChip.classList.remove('loaded');
        salesChip.innerHTML = `<i data-lucide="alert-circle"></i><span>Sales Data: Not Loaded</span>`;
    }

    if (state.stockData.length > 0) {
        stockChip.classList.add('loaded');
        stockChip.innerHTML = `<i data-lucide="check-circle"></i><span>Stock Data: ${state.stockData.length} rows loaded</span>`;
    } else {
        stockChip.classList.remove('loaded');
        stockChip.innerHTML = `<i data-lucide="alert-circle"></i><span>Stock Data: Not Loaded</span>`;
    }

    lucide.createIcons();
}

/**
 * Merge sales and stock data into rawData for downstream processing.
 * Each sales row gets the stock from the stock dataset (matched by SKU).
 * Product name is taken from whichever dataset provides it.
 */
function mergeData() {
    // Build stock lookup by SKU
    const stockMap = {};
    state.stockData.forEach(s => {
        stockMap[s.sku] = s.qty;
    });

    // Build product name map from both datasets
    state.productNameMap = {};
    state.stockData.forEach(s => {
        if (s.productName) state.productNameMap[s.sku] = s.productName;
    });
    state.salesData.forEach(s => {
        if (s.productName) state.productNameMap[s.sku] = s.productName;
    });

    // Merge: sales data + stock from stock dataset
    state.rawData = state.salesData.map(s => ({
        sku: s.sku,
        productName: state.productNameMap[s.sku] || s.productName || '',
        date: s.date,
        sales: s.soldQty,
        stock: stockMap[s.sku] !== undefined ? stockMap[s.sku] : 0,
    }));

    renderMergedTable();
    onDataChanged();
}

/**
 * Render the merged data preview table
 */
function renderMergedTable() {
    const card = document.getElementById('mergedDataCard');
    const tbody = document.getElementById('mergedTableBody');
    const rowCount = document.getElementById('mergedRowCount');

    if (state.rawData.length === 0) {
        card.style.display = 'none';
        return;
    }
    card.style.display = '';
    rowCount.textContent = `${state.rawData.length} rows`;

    // Show max 50 rows in preview
    const displayRows = state.rawData.slice(0, 50);
    tbody.innerHTML = displayRows.map(row => `
        <tr>
            <td style="color:var(--text-primary);font-weight:500;">${row.sku}</td>
            <td>${row.productName}</td>
            <td>${row.date}</td>
            <td>${row.sales}</td>
            <td>${row.stock}</td>
        </tr>
    `).join('');

    if (state.rawData.length > 50) {
        tbody.innerHTML += `<tr><td colspan="5" style="text-align:center;color:var(--text-muted);font-style:italic;">
            ... and ${state.rawData.length - 50} more rows
        </td></tr>`;
    }
}

// ────────────────────────────────────────────
// ░░  FORECASTING MODULE
// ────────────────────────────────────────────
function initForecasting() {
    const toggleBtns = document.querySelectorAll('#forecastMethodToggle .toggle-btn');
    const runBtn = document.getElementById('runForecastBtn');

    toggleBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            toggleBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.forecastMethod = btn.dataset.method;
        });
    });

    runBtn.addEventListener('click', () => {
        if (state.rawData.length === 0) { toast('Upload data first!', 'error'); return; }
        runForecast();
        toast('Forecast generated!', 'success');
    });
}

/**
 * Run forecast calculations per SKU
 */
function runForecast() {
    const grouped = groupBySku(state.rawData);
    state.forecasts = [];

    for (const sku in grouped) {
        const sales = grouped[sku].map(r => r.sales);
        const last3 = sales.slice(-3);

        let forecast = 0;
        if (state.forecastMethod === 'ma') {
            forecast = last3.reduce((a, b) => a + b, 0) / last3.length;
        } else {
            const weights = [0.5, 0.3, 0.2];
            const reversed = [...last3].reverse();
            forecast = reversed.reduce((sum, val, idx) => sum + val * (weights[idx] || 0), 0);
        }

        state.forecasts.push({
            sku,
            productName: state.productNameMap[sku] || '',
            last3: last3.join(', '),
            forecast: Math.round(forecast * 100) / 100,
            method: state.forecastMethod === 'ma' ? 'Moving Avg (3)' : 'Weighted Avg',
        });
    }

    renderForecastTable();
    computeKPIs();
    applyDemandDrivers();
    updateDashboard();
}

/**
 * Render forecast results table
 */
function renderForecastTable() {
    const card = document.getElementById('forecastResultsCard');
    const tbody = document.getElementById('forecastTableBody');
    card.style.display = '';

    tbody.innerHTML = state.forecasts.map(f => `
        <tr>
            <td style="color:var(--text-primary);font-weight:500;">${f.sku}</td>
            <td>${f.productName}</td>
            <td>${f.last3}</td>
            <td style="color:var(--accent-hover);font-weight:600;">${f.forecast.toLocaleString()}</td>
            <td><span class="badge">${f.method}</span></td>
        </tr>
    `).join('');
}

// ────────────────────────────────────────────
// ░░  KPI CALCULATIONS
// ────────────────────────────────────────────
function computeKPIs() {
    const grouped = groupBySku(state.rawData);
    let mapeSum = 0, mapeCount = 0, biasSum = 0;

    state.forecasts.forEach(f => {
        const skuData = grouped[f.sku];
        if (!skuData || skuData.length === 0) return;
        const lastActual = skuData[skuData.length - 1].sales;
        if (lastActual !== 0) {
            mapeSum += Math.abs(lastActual - f.forecast) / Math.abs(lastActual);
            mapeCount++;
        }
        biasSum += f.forecast - lastActual;
    });

    const mape = mapeCount > 0 ? (mapeSum / mapeCount * 100) : 0;
    const bias = biasSum;

    // Forecast KPI section
    const kpiRow = document.getElementById('forecastKpiRow');
    kpiRow.style.display = '';
    document.getElementById('forecastMape').textContent = mape.toFixed(1) + '%';
    document.getElementById('forecastBias').textContent = (bias >= 0 ? '+' : '') + bias.toFixed(0);

    // Dashboard KPIs
    const totalSales = state.rawData.reduce((s, r) => s + r.sales, 0);
    const avgForecast = state.forecasts.length > 0
        ? state.forecasts.reduce((s, f) => s + f.forecast, 0) / state.forecasts.length : 0;

    document.getElementById('kpiTotalSales').textContent = totalSales.toLocaleString();
    document.getElementById('kpiAvgForecast').textContent = avgForecast.toFixed(0);
    document.getElementById('kpiMape').textContent = mape.toFixed(1) + '%';
    document.getElementById('kpiBias').textContent = (bias >= 0 ? '+' : '') + bias.toFixed(0);
}

// ────────────────────────────────────────────
// ░░  DEMAND DRIVERS MODULE
// ────────────────────────────────────────────
function initDemandDrivers() {
    const promoSlider = document.getElementById('promotionSlider');
    const promoValue = document.getElementById('promotionValue');
    const seasonSlider = document.getElementById('seasonalitySlider');
    const seasonValue = document.getElementById('seasonalityValue');

    promoSlider.addEventListener('input', () => {
        state.promotion = parseInt(promoSlider.value);
        promoValue.textContent = state.promotion + '%';
        applyDemandDrivers();
    });

    seasonSlider.addEventListener('input', () => {
        state.seasonality = parseFloat(seasonSlider.value);
        seasonValue.textContent = state.seasonality.toFixed(2) + '×';
        applyDemandDrivers();
    });
}

/**
 * Apply promotion & seasonality adjustments
 */
function applyDemandDrivers() {
    if (state.forecasts.length === 0) return;

    state.adjustedForecasts = state.forecasts.map(f => {
        const promoAdjust = f.forecast * (state.promotion / 100);
        const adjusted = (f.forecast + promoAdjust) * state.seasonality;
        return {
            sku: f.sku,
            productName: f.productName || state.productNameMap[f.sku] || '',
            baseForecast: f.forecast,
            promoUplift: Math.round(promoAdjust * 100) / 100,
            seasonality: state.seasonality,
            adjustedForecast: Math.round(adjusted * 100) / 100,
        };
    });

    renderAdjustedForecastTable();
}

function renderAdjustedForecastTable() {
    const card = document.getElementById('adjustedForecastCard');
    const tbody = document.getElementById('adjustedTableBody');
    if (state.adjustedForecasts.length === 0) { card.style.display = 'none'; return; }
    card.style.display = '';

    tbody.innerHTML = state.adjustedForecasts.map(a => `
        <tr>
            <td style="color:var(--text-primary);font-weight:500;">${a.sku}</td>
            <td>${a.productName}</td>
            <td>${a.baseForecast.toLocaleString()}</td>
            <td style="color:var(--green);">+${a.promoUplift.toLocaleString()}</td>
            <td>${a.seasonality.toFixed(2)}×</td>
            <td style="color:var(--accent-hover);font-weight:600;">${a.adjustedForecast.toLocaleString()}</td>
        </tr>
    `).join('');
}

// ────────────────────────────────────────────
// ░░  INVENTORY PLANNING MODULE
// ────────────────────────────────────────────
function initInventory() {
    const calcBtn = document.getElementById('calcInventoryBtn');
    const leadTimeInput = document.getElementById('leadTimeInput');

    leadTimeInput.addEventListener('change', () => {
        state.leadTime = parseInt(leadTimeInput.value) || 7;
    });

    calcBtn.addEventListener('click', () => {
        // Always read the latest value from the input
        state.leadTime = parseInt(leadTimeInput.value) || 7;
        if (state.rawData.length === 0) { toast('Upload data first!', 'error'); return; }
        calculateInventory();
        toast('Inventory calculated!', 'success');
    });
}

/**
 * Calculate reorder points and stock status.
 * Reorder point logic:
 *   - Lead time ≤ 30 days → reorder point = avg demand (no inflation)
 *   - Lead time > 30 days  → reorder point = avg demand × (leadTime / 30)
 *     so it only exceeds avg demand when lead time is long.
 */
function calculateInventory() {
    const grouped = groupBySku(state.rawData);
    state.inventoryStatus = [];

    for (const sku in grouped) {
        const rows = grouped[sku];
        const avgDemand = rows.reduce((s, r) => s + r.sales, 0) / rows.length;
        const currentStock = rows[rows.length - 1].stock;

        // Reorder point should only exceed avg demand when lead time > 30 days
        let reorderPt;
        if (state.leadTime <= 30) {
            reorderPt = Math.round(avgDemand);
        } else {
            reorderPt = Math.round(avgDemand * (state.leadTime / 30));
        }

        let status = 'ok';
        if (currentStock < reorderPt) status = 'low';
        else if (currentStock > reorderPt * 2.5) status = 'high';

        state.inventoryStatus.push({
            sku,
            productName: state.productNameMap[sku] || '',
            avgDemand: Math.round(avgDemand),
            stock: currentStock,
            reorderPt,
            status,
        });
    }

    renderInventoryTable();
    updateDashboardAlerts();
}

function renderInventoryTable() {
    const card = document.getElementById('inventoryTableCard');
    const tbody = document.getElementById('inventoryTableBody');
    card.style.display = '';

    const statusLabel = { ok: 'Adequate', low: 'Low Stock', high: 'High Stock' };
    const statusIcon = { ok: 'check-circle', low: 'alert-triangle', high: 'alert-circle' };

    tbody.innerHTML = state.inventoryStatus.map(inv => `
        <tr>
            <td style="color:var(--text-primary);font-weight:500;">${inv.sku}</td>
            <td>${inv.productName}</td>
            <td>${inv.avgDemand.toLocaleString()}</td>
            <td>${inv.stock.toLocaleString()}</td>
            <td>${inv.reorderPt.toLocaleString()}</td>
            <td><span class="status ${inv.status}"><i data-lucide="${statusIcon[inv.status]}"></i>${statusLabel[inv.status]}</span></td>
        </tr>
    `).join('');

    lucide.createIcons();
}

// ────────────────────────────────────────────
// ░░  WHAT-IF ANALYSIS MODULE
// ────────────────────────────────────────────
function initWhatIf() {
    const slider = document.getElementById('whatIfSlider');
    const valueDisplay = document.getElementById('whatIfValue');

    slider.addEventListener('input', () => {
        state.whatIfPct = parseInt(slider.value);
        const sign = state.whatIfPct >= 0 ? '+' : '';
        valueDisplay.textContent = sign + state.whatIfPct + '%';
        // Color
        valueDisplay.style.color = state.whatIfPct > 0 ? 'var(--green)' : state.whatIfPct < 0 ? 'var(--rose)' : 'var(--accent-hover)';
        updateWhatIf();
    });
}

function updateWhatIf() {
    if (state.forecasts.length === 0) return;

    const factor = 1 + state.whatIfPct / 100;
    const adjustedForecasts = state.forecasts.map(f => ({
        sku: f.sku,
        productName: f.productName || state.productNameMap[f.sku] || '',
        original: f.forecast,
        adjusted: Math.round(f.forecast * factor * 100) / 100,
    }));

    // Chart
    renderWhatIfChart(adjustedForecasts);

    // Table — Status is based on current stock vs adjusted forecast
    const grouped = groupBySku(state.rawData);
    const card = document.getElementById('whatIfTableCard');
    const tbody = document.getElementById('whatIfTableBody');
    card.style.display = '';

    tbody.innerHTML = adjustedForecasts.map(a => {
        const skuData = grouped[a.sku];
        const stock = skuData ? skuData[skuData.length - 1].stock : 0;

        // New Status logic: compare stock directly against adjusted forecast
        let status = 'ok';
        if (stock < a.adjusted) {
            status = 'low';   // Stock is below what the adjusted forecast demands
        } else if (stock > a.adjusted * 2) {
            status = 'high';  // Stock is more than double the adjusted forecast — overstock
        }

        const statusLabel = { ok: 'Adequate', low: 'Low Stock', high: 'Overstock' };
        const statusIcon = { ok: 'check-circle', low: 'alert-triangle', high: 'alert-circle' };
        return `
            <tr>
                <td style="color:var(--text-primary);font-weight:500;">${a.sku}</td>
                <td>${a.productName}</td>
                <td>${a.original.toLocaleString()}</td>
                <td style="color:var(--accent-hover);font-weight:600;">${a.adjusted.toLocaleString()}</td>
                <td>${stock.toLocaleString()}</td>
                <td><span class="status ${status}"><i data-lucide="${statusIcon[status]}"></i>${statusLabel[status]}</span></td>
            </tr>
        `;
    }).join('');

    lucide.createIcons();
}

function renderWhatIfChart(data) {
    const ctx = document.getElementById('whatIfChart').getContext('2d');
    if (whatIfChart) whatIfChart.destroy();

    whatIfChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.map(d => d.sku),
            datasets: [
                {
                    label: 'Original Forecast',
                    data: data.map(d => d.original),
                    backgroundColor: 'rgba(99,102,241,0.35)',
                    borderColor: 'rgba(99,102,241,0.8)',
                    borderWidth: 1,
                    borderRadius: 6,
                },
                {
                    label: 'Adjusted Forecast',
                    data: data.map(d => d.adjusted),
                    backgroundColor: 'rgba(167,139,250,0.35)',
                    borderColor: 'rgba(167,139,250,0.8)',
                    borderWidth: 1,
                    borderRadius: 6,
                },
            ],
        },
        options: chartOptions('Forecast Comparison'),
    });
}

// ────────────────────────────────────────────
// ░░  DASHBOARD CHARTS & ALERTS
// ────────────────────────────────────────────
function updateDashboard() {
    renderSalesForecastChart();
    renderStockLevelsChart();
    populateSkuDropdown();
    updateDashboardAlerts();
}

function populateSkuDropdown() {
    const select = document.getElementById('dashSkuSelect');
    const skus = [...new Set(state.rawData.map(r => r.sku))];
    select.innerHTML = '<option value="__all__">All SKUs</option>' +
        skus.map(s => `<option value="${s}">${s}</option>`).join('');

    select.onchange = () => renderSalesForecastChart(select.value);
}

function renderSalesForecastChart(selectedSku = '__all__') {
    const ctx = document.getElementById('salesForecastChart').getContext('2d');
    if (salesForecastChart) salesForecastChart.destroy();

    const grouped = groupBySku(state.rawData);
    let labels, salesData, forecastData;

    if (selectedSku === '__all__') {
        const skus = Object.keys(grouped);
        labels = skus;
        salesData = skus.map(sku => grouped[sku].reduce((s, r) => s + r.sales, 0));
        forecastData = skus.map(sku => {
            const f = state.forecasts.find(fc => fc.sku === sku);
            return f ? f.forecast * grouped[sku].length : 0;
        });
    } else {
        const rows = grouped[selectedSku] || [];
        labels = rows.map(r => r.date);
        salesData = rows.map(r => r.sales);
        const f = state.forecasts.find(fc => fc.sku === selectedSku);
        forecastData = rows.map(() => f ? f.forecast : 0);
    }

    salesForecastChart = new Chart(ctx, {
        type: selectedSku === '__all__' ? 'bar' : 'line',
        data: {
            labels,
            datasets: [
                {
                    label: 'Sales',
                    data: salesData,
                    backgroundColor: 'rgba(59,130,246,0.3)',
                    borderColor: 'rgba(59,130,246,0.9)',
                    borderWidth: 2,
                    borderRadius: 6,
                    fill: selectedSku !== '__all__',
                    tension: 0.4,
                    pointBackgroundColor: '#3b82f6',
                    pointRadius: selectedSku !== '__all__' ? 4 : 0,
                },
                {
                    label: 'Forecast',
                    data: forecastData,
                    backgroundColor: 'rgba(167,139,250,0.3)',
                    borderColor: 'rgba(167,139,250,0.9)',
                    borderWidth: 2,
                    borderRadius: 6,
                    borderDash: selectedSku !== '__all__' ? [5, 5] : [],
                    fill: false,
                    tension: 0.4,
                    pointBackgroundColor: '#a78bfa',
                    pointRadius: selectedSku !== '__all__' ? 4 : 0,
                },
            ],
        },
        options: chartOptions('Sales vs Forecast'),
    });
}

function renderStockLevelsChart() {
    const ctx = document.getElementById('stockLevelsChart').getContext('2d');
    if (stockLevelsChart) stockLevelsChart.destroy();

    // Use stock data directly if available, otherwise infer from rawData
    let skus, stockDataArr;
    if (state.stockData.length > 0) {
        skus = state.stockData.map(s => s.sku);
        stockDataArr = state.stockData.map(s => s.qty);
    } else {
        const grouped = groupBySku(state.rawData);
        skus = Object.keys(grouped);
        stockDataArr = skus.map(sku => {
            const rows = grouped[sku];
            return rows[rows.length - 1].stock;
        });
    }

    // Color bars based on stock level relative to average
    const avgStock = stockDataArr.reduce((a, b) => a + b, 0) / stockDataArr.length;
    const colors = stockDataArr.map(s =>
        s < avgStock * 0.5 ? 'rgba(244,63,94,0.5)' :
        s > avgStock * 1.8 ? 'rgba(245,158,11,0.5)' :
        'rgba(20,184,166,0.4)'
    );
    const borderColors = stockDataArr.map(s =>
        s < avgStock * 0.5 ? 'rgba(244,63,94,0.9)' :
        s > avgStock * 1.8 ? 'rgba(245,158,11,0.9)' :
        'rgba(20,184,166,0.9)'
    );

    stockLevelsChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: skus,
            datasets: [{
                label: 'Current Stock',
                data: stockDataArr,
                backgroundColor: colors,
                borderColor: borderColors,
                borderWidth: 1,
                borderRadius: 6,
            }],
        },
        options: chartOptions('Stock Levels'),
    });
}

function updateDashboardAlerts() {
    const container = document.getElementById('dashAlerts');
    if (state.inventoryStatus.length === 0) {
        container.innerHTML = '<p class="empty-state"><i data-lucide="inbox"></i> Run Inventory Planning to see alerts.</p>';
        lucide.createIcons();
        return;
    }

    const alerts = state.inventoryStatus.filter(i => i.status !== 'ok');
    if (alerts.length === 0) {
        container.innerHTML = '<p class="empty-state"><i data-lucide="check-circle"></i> All stock levels are adequate.</p>';
        lucide.createIcons();
        return;
    }

    container.innerHTML = '<div class="alert-list">' + alerts.map(a => `
        <div class="alert-item ${a.status}">
            <i data-lucide="${a.status === 'low' ? 'alert-triangle' : 'alert-circle'}"></i>
            <strong>${a.sku}</strong> (${a.productName || '—'}) — ${a.status === 'low' ? 'Stock below reorder point' : 'Overstocked'} (Stock: ${a.stock.toLocaleString()}, Reorder: ${a.reorderPt.toLocaleString()})
        </div>
    `).join('') + '</div>';
    lucide.createIcons();
}

// ────────────────────────────────────────────
// ░░  EXPORT TO CSV
// ────────────────────────────────────────────
function initExport() {
    document.getElementById('exportCsvBtn').addEventListener('click', () => {
        if (state.rawData.length === 0 && state.forecasts.length === 0) {
            toast('No data to export.', 'error');
            return;
        }
        exportToCSV();
        toast('CSV exported!', 'success');
    });
}

function exportToCSV() {
    const grouped = groupBySku(state.rawData);
    const lines = ['SKU,Product Name,Date,Sales,Stock,Forecast,Adjusted Forecast,Inventory Status'];

    for (const sku in grouped) {
        const rows = grouped[sku];
        const f = state.forecasts.find(fc => fc.sku === sku);
        const adj = state.adjustedForecasts.find(a => a.sku === sku);
        const inv = state.inventoryStatus.find(i => i.sku === sku);

        rows.forEach(r => {
            lines.push([
                r.sku,
                `"${r.productName || ''}"`,
                r.date,
                r.sales,
                r.stock,
                f ? f.forecast : '',
                adj ? adj.adjustedForecast : '',
                inv ? inv.status : '',
            ].join(','));
        });
    }

    const blob = new Blob([lines.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'demandflow_export.csv';
    a.click();
    URL.revokeObjectURL(url);
}

// ────────────────────────────────────────────
// ░░  HELPER FUNCTIONS
// ────────────────────────────────────────────

/**
 * Find a header index from a list of possible names
 */
function findHeaderIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const idx = headers.indexOf(name);
        if (idx !== -1) return idx;
    }
    return -1;
}

/**
 * Group rawData rows by SKU, sorted by date within each group
 */
function groupBySku(data) {
    const map = {};
    data.forEach(row => {
        if (!map[row.sku]) map[row.sku] = [];
        map[row.sku].push(row);
    });
    // Sort each group by date
    for (const sku in map) {
        map[sku].sort((a, b) => new Date(a.date) - new Date(b.date));
    }
    return map;
}

/**
 * Called whenever underlying data changes
 */
function onDataChanged() {
    // Reset derived state
    state.forecasts = [];
    state.adjustedForecasts = [];
    state.inventoryStatus = [];
    // Reset dashboard KPIs to placeholder
    document.getElementById('kpiTotalSales').textContent = state.rawData.length > 0
        ? state.rawData.reduce((s, r) => s + r.sales, 0).toLocaleString() : '—';
    document.getElementById('kpiAvgForecast').textContent = '—';
    document.getElementById('kpiMape').textContent = '—';
    document.getElementById('kpiBias').textContent = '—';
    // Hide forecast results
    document.getElementById('forecastResultsCard').style.display = 'none';
    document.getElementById('forecastKpiRow').style.display = 'none';
    document.getElementById('adjustedForecastCard').style.display = 'none';
    document.getElementById('inventoryTableCard').style.display = 'none';
    document.getElementById('whatIfTableCard').style.display = 'none';
    // Destroy charts if they exist
    if (salesForecastChart) { salesForecastChart.destroy(); salesForecastChart = null; }
    if (stockLevelsChart) { stockLevelsChart.destroy(); stockLevelsChart = null; }
    if (whatIfChart) { whatIfChart.destroy(); whatIfChart = null; }
    // Update alerts
    document.getElementById('dashAlerts').innerHTML = '<p class="empty-state"><i data-lucide="inbox"></i> Upload data to see inventory alerts.</p>';
    lucide.createIcons();
}

/**
 * Shared Chart.js options for consistent dark-theme look
 */
function chartOptions(title) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                labels: { color: '#94a3b8', font: { family: 'Inter', size: 12 }, boxWidth: 14, padding: 16 },
            },
            tooltip: {
                backgroundColor: 'rgba(17,24,39,0.95)',
                titleColor: '#f1f5f9',
                bodyColor: '#cbd5e1',
                borderColor: 'rgba(255,255,255,0.1)',
                borderWidth: 1,
                cornerRadius: 8,
                padding: 12,
                titleFont: { family: 'Inter', weight: '600' },
                bodyFont: { family: 'Inter' },
            },
        },
        scales: {
            x: {
                ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
                grid: { color: 'rgba(255,255,255,0.04)' },
            },
            y: {
                ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
                grid: { color: 'rgba(255,255,255,0.04)' },
            },
        },
    };
}

// ────────────────────────────────────────────
// ░░  SAMPLE DATA GENERATORS
// ────────────────────────────────────────────

const SAMPLE_PRODUCTS = {
    'SKU-A100': 'Widget Alpha',
    'SKU-B200': 'Gadget Beta',
    'SKU-C300': 'Super Charger',
    'SKU-D400': 'Power Module',
    'SKU-E500': 'Turbo Flex',
};

/**
 * Generate sample sales data
 */
function generateSampleSalesData() {
    const skus = Object.keys(SAMPLE_PRODUCTS);
    const months = ['2024-01', '2024-02', '2024-03', '2024-04', '2024-05', '2024-06'];
    const data = [];

    const baseSales = { 'SKU-A100': 120, 'SKU-B200': 85, 'SKU-C300': 200, 'SKU-D400': 60, 'SKU-E500': 150 };
    const trend = { 'SKU-A100': 1.05, 'SKU-B200': 0.97, 'SKU-C300': 1.1, 'SKU-D400': 1.0, 'SKU-E500': 1.03 };

    skus.forEach(sku => {
        months.forEach((month, i) => {
            const noise = 1 + (Math.random() - 0.5) * 0.2;
            const soldQty = Math.round(baseSales[sku] * Math.pow(trend[sku], i) * noise);
            data.push({ sku, productName: SAMPLE_PRODUCTS[sku], date: month + '-01', soldQty });
        });
    });

    return data;
}

/**
 * Generate sample stock data
 */
function generateSampleStockData() {
    const skus = Object.keys(SAMPLE_PRODUCTS);
    const baseStock = { 'SKU-A100': 400, 'SKU-B200': 150, 'SKU-C300': 800, 'SKU-D400': 300, 'SKU-E500': 500 };
    return skus.map(sku => ({
        sku,
        productName: SAMPLE_PRODUCTS[sku],
        qty: Math.round(baseStock[sku] * (1 + (Math.random() - 0.5) * 0.3)),
    }));
}

/**
 * Generate sample Sales CSV text for the textarea
 */
function generateSampleSalesCSVText() {
    const data = generateSampleSalesData();
    const lines = ['SKU,Product Name,Date,Sold Qty'];
    data.forEach(r => lines.push(`${r.sku},${r.productName},${r.date},${r.soldQty}`));
    return lines.join('\n');
}

/**
 * Generate sample Stock CSV text for the textarea
 */
function generateSampleStockCSVText() {
    const data = generateSampleStockData();
    const lines = ['SKU,Product Name,Qty'];
    data.forEach(r => lines.push(`${r.sku},${r.productName},${r.qty}`));
    return lines.join('\n');
}

/**
 * Show a toast notification
 */
function toast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const iconMap = { success: 'check-circle', error: 'alert-circle', info: 'info' };
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.innerHTML = `<i data-lucide="${iconMap[type]}"></i><span>${message}</span>`;
    container.appendChild(el);
    lucide.createIcons();
    setTimeout(() => el.remove(), 3000);
}
