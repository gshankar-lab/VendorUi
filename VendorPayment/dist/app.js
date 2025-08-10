"use strict";
// ------------------ UI Elements ------------------
const appContainer = document.getElementById('app');
const loginSection = document.getElementById('login-section');
const mainAppSection = document.getElementById('main-app-section');
const mainContent = document.getElementById('main-content');
const loginForm = document.getElementById('login-form');
const logoutButton = document.getElementById('logout-button');
const loginMessage = document.getElementById('login-message');
// ------------------ Mock User ------------------
const MOCK_USER = {
    username: 'admin',
    password: 'password123'
};
// ------------------ Login Check ------------------
function checkLoginStatus() {
    const isLoggedIn = localStorage.getItem('isLoggedIn');
    if (isLoggedIn === 'true') {
        showMainApp();
        loadVendorHtml();
    }
    else {
        showLogin();
    }
}
// ------------------ UI Display ------------------
function showLogin() {
    appContainer.classList.add('centered-container');
    loginSection.classList.remove('hidden');
    mainAppSection.classList.add('hidden');
}
function showMainApp() {
    appContainer.classList.remove('centered-container');
    loginSection.classList.add('hidden');
    mainAppSection.classList.remove('hidden');
}
// ------------------ Login Handling ------------------
loginForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    if (usernameInput.value === MOCK_USER.username && passwordInput.value === MOCK_USER.password) {
        localStorage.setItem('isLoggedIn', 'true');
        showMainApp();
        loadVendorHtml();
        loginMessage.textContent = '';
    }
    else {
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
function loadVendorHtml() {
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
let vendorIndexToDelete = null;
function getVendors() {
    return JSON.parse(localStorage.getItem("vendors") || "[]");
}
function saveVendors(vendors) {
    localStorage.setItem("vendors", JSON.stringify(vendors));
}
function fetchVendors() {
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
    const tbody = document.getElementById('vendorTableBody');
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
function setupVendorHandlers() {
    const form = document.getElementById('vendorForm');
    const cancelBtn = document.getElementById('cancelEditBtn');
    const deleteModal = document.getElementById('deleteModal');
    const confirmDeleteBtn = document.getElementById('confirmDeleteBtn');
    const cancelDeleteBtn = document.getElementById('cancelDeleteBtn');
    // Add & Edit
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        const name = document.getElementById('vendorName').value;
        const type = document.getElementById('paymentType').value;
        const account = document.getElementById('assignedAccount').value;
        const indexVal = document.getElementById('vendorIndex').value;
        const vendors = getVendors();
        if (indexVal) {
            vendors[parseInt(indexVal)] = { "Vendor Name": name, "Payment Type": type, "Assigned Account": account };
        }
        else {
            vendors.push({ "Vendor Name": name, "Payment Type": type, "Assigned Account": account });
        }
        saveVendors(vendors);
        fetchVendors();
        form.reset();
        document.getElementById('vendorIndex').value = '';
        cancelBtn.classList.add('hidden');
    });
    cancelBtn.addEventListener('click', () => {
        form.reset();
        document.getElementById('vendorIndex').value = '';
        cancelBtn.classList.add('hidden');
    });
    // Global edit
    window.editVendor = (index) => {
        const vendors = getVendors();
        const vendor = vendors[index];
        document.getElementById('vendorName').value = vendor["Vendor Name"];
        document.getElementById('paymentType').value = vendor["Payment Type"];
        document.getElementById('assignedAccount').value = vendor["Assigned Account"];
        document.getElementById('vendorIndex').value = index.toString();
        cancelBtn.classList.remove('hidden');
    };
    // Global delete trigger
    window.triggerDelete = (index) => {
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
