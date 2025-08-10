
// ------------------ UI Elements ------------------
const appContainer = document.getElementById('app') as HTMLElement;
const loginSection = document.getElementById('login-section') as HTMLElement;
const mainAppSection = document.getElementById('main-app-section') as HTMLElement;
const mainContent = document.getElementById('main-content') as HTMLElement;
const loginForm = document.getElementById('login-form') as HTMLFormElement;
const logoutButton = document.getElementById('logout-button') as HTMLButtonElement;
const loginMessage = document.getElementById('login-message') as HTMLElement;

// ------------------ Mock User ------------------
const MOCK_USER = {
    username: 'admin',
    password: 'password123'
};

// ------------------ Login Check ------------------
function checkLoginStatus(): void {
    const isLoggedIn = localStorage.getItem('isLoggedIn');
    if (isLoggedIn === 'true') {
        showMainApp();
        loadVendorHtml(); 
    } else {
        showLogin();
    }
}

// ------------------ UI Display ------------------
function showLogin(): void {
    appContainer.classList.add('centered-container');
    loginSection.classList.remove('hidden');
    mainAppSection.classList.add('hidden');
}

function showMainApp(): void {
    appContainer.classList.remove('centered-container');
    loginSection.classList.add('hidden');
    mainAppSection.classList.remove('hidden');
}

// ------------------ Login Handling ------------------
loginForm.addEventListener('submit', (event) => {
    event.preventDefault();

    const usernameInput = document.getElementById('username') as HTMLInputElement;
    const passwordInput = document.getElementById('password') as HTMLInputElement;

    if (usernameInput.value === MOCK_USER.username && passwordInput.value === MOCK_USER.password) {
        localStorage.setItem('isLoggedIn', 'true');
        showMainApp();
        loadVendorHtml();
        loginMessage.textContent = '';
    } else {
        loginMessage.textContent = 'Invalid username or password.';
    }
});

// ------------------ Logout Handling ------------------
logoutButton.addEventListener('click', () => {
    localStorage.removeItem('isLoggedIn');
    showLogin();
});

checkLoginStatus();

// ------------------ Load Vendor HTML Dynamically ------------------
function loadVendorHtml(): void {
    fetch('vendor.html')
        .then(res => res.text())
        .then(html => {
            mainContent.innerHTML = html; 
            setupVendorHandlers(); // bind CRUD events
            fetchVendors();
        })
        .catch(err => console.error('Error loading vendor.html:', err));
}

// ------------------ Vendor CRUD Logic ------------------
let vendorIndexToDelete: number | null = null;

function getVendors(): any[] {
    return JSON.parse(localStorage.getItem("vendors") || "[]");
}

function saveVendors(vendors: any[]): void {
    localStorage.setItem("vendors", JSON.stringify(vendors));
}

function fetchVendors(): void {
    if (!localStorage.getItem("vendors")) {
        saveVendors([
            { "Vendor Name": "Alpha Supplies", "Payment Type": "Weekly", "Assigned Account": "123456" },
            { "Vendor Name": "Beta Logistics", "Payment Type": "Biweekly", "Assigned Account": "789012" },
            { "Vendor Name": "Gamma Services", "Payment Type": "On-Demand", "Assigned Account": "345678" },
            { "Vendor Name": "Delta Solutions", "Payment Type": "Weekly", "Assigned Account": "901234" },
            { "Vendor Name": "Epsilon Inc.", "Payment Type": "Biweekly", "Assigned Account": "567890" }
        ]);
    }

    const vendors = getVendors();
    const tbody = document.getElementById('vendorTableBody') as HTMLElement;
    tbody.innerHTML = '';

    vendors.forEach((vendor, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vendor["Vendor Name"]}</td>
            <td>${vendor["Payment Type"]}</td>
            <td>${vendor["Assigned Account"]}</td>
            <td>
                <button onclick="editVendor(${index})">Edit</button>
                <button onclick="triggerDelete(${index})">Delete</button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// ------------------ CRUD Event Handlers ------------------
function setupVendorHandlers(): void {
    const form = document.getElementById('vendorForm') as HTMLFormElement;
    const cancelBtn = document.getElementById('cancelEditBtn') as HTMLButtonElement;

    const deleteModal = document.getElementById('deleteModal') as HTMLElement;
    const confirmDeleteBtn = document.getElementById('confirmDeleteBtn') as HTMLButtonElement;
    const cancelDeleteBtn = document.getElementById('cancelDeleteBtn') as HTMLButtonElement;

    // Add & Edit
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        const name = (document.getElementById('vendorName') as HTMLInputElement).value;
        const type = (document.getElementById('paymentType') as HTMLInputElement).value;
        const account = (document.getElementById('assignedAccount') as HTMLInputElement).value;
        const indexVal = (document.getElementById('vendorIndex') as HTMLInputElement).value;

        const vendors = getVendors();

        if (indexVal) {
            vendors[parseInt(indexVal)] = { "Vendor Name": name, "Payment Type": type, "Assigned Account": account };
        } else {
            vendors.push({ "Vendor Name": name, "Payment Type": type, "Assigned Account": account });
        }

        saveVendors(vendors);
        fetchVendors();
        form.reset();
        (document.getElementById('vendorIndex') as HTMLInputElement).value = '';
        cancelBtn.classList.add('hidden');
    });

    cancelBtn.addEventListener('click', () => {
        form.reset();
        (document.getElementById('vendorIndex') as HTMLInputElement).value = '';
        cancelBtn.classList.add('hidden');
    });

    // Global edit
    (window as any).editVendor = (index: number) => {
        const vendors = getVendors();
        const vendor = vendors[index];
        (document.getElementById('vendorName') as HTMLInputElement).value = vendor["Vendor Name"];
        (document.getElementById('paymentType') as HTMLInputElement).value = vendor["Payment Type"];
        (document.getElementById('assignedAccount') as HTMLInputElement).value = vendor["Assigned Account"];
        (document.getElementById('vendorIndex') as HTMLInputElement).value = index.toString();
        cancelBtn.classList.remove('hidden');
    };

    // Global delete trigger
    (window as any).triggerDelete = (index: number) => {
        vendorIndexToDelete = index;
        deleteModal.classList.remove('hidden');
    };

    // Confirm delete
    confirmDeleteBtn.addEventListener('click', () => {
        if (vendorIndexToDelete !== null) {
            const vendors = getVendors();
            vendors.splice(vendorIndexToDelete, 1);
            saveVendors(vendors);
            fetchVendors();
            vendorIndexToDelete = null;
        }
        deleteModal.classList.add('hidden');
    });

    // Cancel delete
    cancelDeleteBtn.addEventListener('click', () => {
        vendorIndexToDelete = null;
        deleteModal.classList.add('hidden');
    });
}

// ------------------ Office Init ------------------
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Office is ready. Host: " + info.host);
        checkLoginStatus();
    }
});
