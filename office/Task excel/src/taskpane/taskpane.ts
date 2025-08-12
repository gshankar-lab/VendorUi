/* global console, document, Excel, Office */

interface Vendor {
  id: string;
  name: string;
  paymentType: string;
  assignedAccount: string;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";

    // Auth events
    document.getElementById("loginBtn")?.addEventListener("click", login);
    document.getElementById("logoutBtn")?.addEventListener("click", logout);

    // Excel function
    document.getElementById("run")?.addEventListener("click", run);

    // Vendor form save
    document.getElementById("saveVendorBtn")?.addEventListener("click", saveVendor);

    // Check login
    const savedUser = localStorage.getItem("user");
    if (savedUser) {
      showAppSection(savedUser);
      loadVendors();
    } else {
      showLoginSection();
    }
  }
});

/* ---------- AUTH ---------- */
function login() {
  const username = (document.getElementById("username") as HTMLInputElement).value;
  const password = (document.getElementById("password") as HTMLInputElement).value;

  if (username === "admin" && password === "1234") {
    localStorage.setItem("user", username);
    showAppSection(username);
    loadVendors();
  } else {
    alert("Invalid username or password");
  }
}

function logout() {
  localStorage.removeItem("user");
  showLoginSection();
}

function showAppSection(username: string) {
  document.getElementById("login-section")!.style.display = "none";
  document.getElementById("app-section")!.style.display = "block";
  (document.getElementById("user-name") as HTMLElement).innerText = username;
}

function showLoginSection() {
  document.getElementById("login-section")!.style.display = "block";
  document.getElementById("app-section")!.style.display = "none";
}

/* ---------- EXCEL ACTION ---------- */
export async function run() {
  const savedUser = localStorage.getItem("user");
  if (!savedUser) {
    alert("Please log in to use this feature.");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

/* ---------- VENDOR MANAGEMENT ---------- */
function loadVendors() {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem("vendors") || "[]");
  const tbody = document.querySelector("#vendorTable tbody")!;
  tbody.innerHTML = "";

  vendors.forEach((vendor) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${vendor.name}</td>
      <td>${vendor.paymentType}</td>
      <td>${vendor.assignedAccount}</td>
      <td>
        <button class="edit-btn" data-id="${vendor.id}">Edit</button>
        <button class="delete-btn" data-id="${vendor.id}">Delete</button>
      </td>
    `;
    tbody.appendChild(row);
  });

  // Attach edit/delete events
  document.querySelectorAll(".edit-btn").forEach(btn =>
    btn.addEventListener("click", () => editVendor((btn as HTMLElement).dataset.id!))
  );
  document.querySelectorAll(".delete-btn").forEach(btn =>
    btn.addEventListener("click", () => deleteVendor((btn as HTMLElement).dataset.id!))
  );
}

function saveVendor() {
  const id = (document.getElementById("vendorId") as HTMLInputElement).value || crypto.randomUUID();
  const name = (document.getElementById("vendorName") as HTMLInputElement).value.trim();
  const paymentType = (document.getElementById("paymentType") as HTMLSelectElement).value;
  const assignedAccount = (document.getElementById("assignedAccount") as HTMLSelectElement).value;

  if (!name) {
    alert("Vendor Name is required");
    return;
  }

  let vendors: Vendor[] = JSON.parse(localStorage.getItem("vendors") || "[]");
  const existingIndex = vendors.findIndex((v) => v.id === id);

  if (existingIndex >= 0) {
    vendors[existingIndex] = { id, name, paymentType, assignedAccount };
  } else {
    vendors.push({ id, name, paymentType, assignedAccount });
  }

  localStorage.setItem("vendors", JSON.stringify(vendors));
  clearVendorForm();
  loadVendors();
}

function editVendor(id: string) {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem("vendors") || "[]");
  const vendor = vendors.find((v) => v.id === id);
  if (vendor) {
    (document.getElementById("vendorId") as HTMLInputElement).value = vendor.id;
    (document.getElementById("vendorName") as HTMLInputElement).value = vendor.name;
    (document.getElementById("paymentType") as HTMLSelectElement).value = vendor.paymentType;
    (document.getElementById("assignedAccount") as HTMLSelectElement).value = vendor.assignedAccount;
  }
}

function deleteVendor(id: string) {
  let vendors: Vendor[] = JSON.parse(localStorage.getItem("vendors") || "[]");
  vendors = vendors.filter((v) => v.id !== id);
  localStorage.setItem("vendors", JSON.stringify(vendors));
  loadVendors();
}

function clearVendorForm() {
  (document.getElementById("vendorId") as HTMLInputElement).value = "";
  (document.getElementById("vendorName") as HTMLInputElement).value = "";
  (document.getElementById("paymentType") as HTMLSelectElement).value = "Weekly";
  (document.getElementById("assignedAccount") as HTMLSelectElement).value = "Account 1";
}
