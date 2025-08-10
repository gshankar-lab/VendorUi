"use strict";
// Define constants for the UI elements
const appContainer = document.getElementById('app');
const loginSection = document.getElementById('login-section');
const mainAppSection = document.getElementById('main-app-section');
const loginForm = document.getElementById('login-form');
const logoutButton = document.getElementById('logout-button');
const loginMessage = document.getElementById('login-message');
const contentArea = document.querySelector('#main-app-section main');
// Mock user data for authentication
const MOCK_USER = {
    username: 'admin',
    password: 'password123'
};
// Check for a logged-in user on app start
function checkLoginStatus() {
    const isLoggedIn = localStorage.getItem('isLoggedIn');
    if (isLoggedIn === 'true') {
        showMainApp();
    }
    else {
        showLogin();
    }
}
// Function to show the login form and hide the main app
function showLogin() {
    appContainer.classList.add('centered-container');
    loginSection.classList.remove('hidden');
    mainAppSection.classList.add('hidden');
}
// Function to show the main app and hide the login form
function showMainApp() {
    appContainer.classList.remove('centered-container');
    loginSection.classList.add('hidden');
    mainAppSection.classList.remove('hidden');
}
// Handle the login form submission
loginForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    const username = usernameInput.value;
    const password = passwordInput.value;
    if (username === MOCK_USER.username && password === MOCK_USER.password) {
        localStorage.setItem('isLoggedIn', 'true');
        showMainApp();
        loginMessage.textContent = '';
    }
    else {
        loginMessage.textContent = 'Invalid username or password.';
    }
});
// Handle the logout button click
logoutButton.addEventListener('click', () => {
    localStorage.removeItem('isLoggedIn');
    showLogin();
});
checkLoginStatus();
// Load vendor.html into main content area
function loadVendorPage() {
    fetch('../html/vendor.html')
        .then(response => response.text())
        .then(html => {
        contentArea.innerHTML = html;
        fetchVendors(); // load vendor data after HTML is injected
    })
        .catch(error => console.error('Error loading vendor.html:', error));
}
// Fetch vendor data from API (or mock)
function fetchVendors() {
    fetch('https://mocki.io/v1/2f7b5f60-334a-4c4a-902e-12bb6a937d33') // replace with your API
        .then(res => res.json())
        .then(data => {
        const tbody = document.getElementById('vendorTableBody');
        tbody.innerHTML = '';
        data.forEach((vendor) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                    <td>${vendor.name}</td>
                    <td>${vendor.paymentType}</td>
                    <td>${vendor.assignedAccount}</td>
                `;
            tbody.appendChild(row);
        });
    })
        .catch(err => console.error('Error fetching vendors:', err));
}
// Office.onReady function ensures Office.js is fully loaded before running our code
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Office is ready. Host: " + info.host);
        checkLoginStatus();
    }
});
