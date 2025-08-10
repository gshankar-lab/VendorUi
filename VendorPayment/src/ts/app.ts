// Define constants for the UI elements
const appContainer = document.getElementById('app') as HTMLElement;
const loginSection = document.getElementById('login-section') as HTMLElement;
const mainAppSection = document.getElementById('main-app-section') as HTMLElement;
const loginForm = document.getElementById('login-form') as HTMLFormElement;
const logoutButton = document.getElementById('logout-button') as HTMLButtonElement;
const loginMessage = document.getElementById('login-message') as HTMLElement;


// Mock user data for authentication
const MOCK_USER = {
    username: 'admin',
    password: 'password123'
};

// Check for a logged-in user on app start
function checkLoginStatus(): void {
    const isLoggedIn = localStorage.getItem('isLoggedIn');
    if (isLoggedIn === 'true') {
        showMainApp();
    } 
    else {
        showLogin();
    }
}

// Function to show the login form and hide the main app
function showLogin(): void {
    appContainer.classList.add('centered-container');
    loginSection.classList.remove('hidden');
    mainAppSection.classList.add('hidden');
}

// Function to show the main app and hide the login form
function showMainApp(): void {
    appContainer.classList.remove('centered-container');
    loginSection.classList.add('hidden');
    mainAppSection.classList.remove('hidden');
}

// Handle the login form submission
loginForm.addEventListener('submit', (event) => {
    event.preventDefault();

    const usernameInput = document.getElementById('username') as HTMLInputElement;
    const passwordInput = document.getElementById('password') as HTMLInputElement;
    const username = usernameInput.value;
    const password = passwordInput.value;

    if (username === MOCK_USER.username && password === MOCK_USER.password) {
        localStorage.setItem('isLoggedIn', 'true');
        showMainApp();
        loginMessage.textContent = '';
    } else {
        loginMessage.textContent = 'Invalid username or password.';
    }
});

// Handle the logout button click
logoutButton.addEventListener('click', () => {
    localStorage.removeItem('isLoggedIn');
    showLogin();
});
checkLoginStatus();
// Office.onReady function ensures Office.js is fully loaded before running our code
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Office is ready. Host: " + info.host);
        checkLoginStatus();
    }
});