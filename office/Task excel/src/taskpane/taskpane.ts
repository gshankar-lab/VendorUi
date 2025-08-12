/* global console, document, Excel, Office */

interface Vendor {
  id: string;
  name: string;
  paymentType: string; // "Weekly" | "Biweekly" | "On-Demand"
  assignedAccount: string; // "Account 1" | "Account 2"
  baseAmount?: number;
  lastPaid?: string | null; // ISO date
  skipNext?: boolean;
}

interface Account {
  id: string;
  name: string;
  balance: number;
}

interface PendingPayment {
  id: string;
  vendorId: string;
  vendorName: string;
  amount: number;
  reason: string;
  createdAt: string;
}

/* ---------- Config ---------- */
const DEFAULT_BASE_AMOUNT = 100; // base amount for vendors
const SECURE_PIN = "1234"; // mock PIN
const PENDING_STORAGE_KEY = "pendingPayments";
const ACCOUNTS_STORAGE_KEY = "accounts";
const STARTING_BALANCE = 200000; // user requested starting balance

/* ---------- Helpers ---------- */
const $ = (id: string) => document.getElementById(id) as HTMLElement | null;
const getId = (): string => {
  try {
    // @ts-ignore
    return crypto?.randomUUID?.() || Math.random().toString(36).slice(2, 10);
  } catch {
    return Math.random().toString(36).slice(2, 10);
  }
};
const nowISO = () => new Date().toISOString();

/* ---------- Storage helpers ---------- */
function getVendors(): Vendor[] {
  return JSON.parse(localStorage.getItem("vendors") || "[]");
}
function setVendors(v: Vendor[]) {
  localStorage.setItem("vendors", JSON.stringify(v));
}
function getPending(): PendingPayment[] {
  return JSON.parse(localStorage.getItem(PENDING_STORAGE_KEY) || "[]");
}
function setPending(p: PendingPayment[]) {
  localStorage.setItem(PENDING_STORAGE_KEY, JSON.stringify(p));
}
function getAccounts(): Account[] {
  const a = JSON.parse(localStorage.getItem(ACCOUNTS_STORAGE_KEY) || "null");
  return a || [];
}
function setAccounts(a: Account[]) {
  localStorage.setItem(ACCOUNTS_STORAGE_KEY, JSON.stringify(a));
}

/* ---------- Ensure default accounts exist ---------- */
function ensureDefaultAccounts() {
  let accounts = getAccounts();
  if (!accounts || accounts.length === 0) {
    accounts = [
      { id: "Account 1", name: "Account 1", balance: STARTING_BALANCE },
      { id: "Account 2", name: "Account 2", balance: STARTING_BALANCE }
    ];
    setAccounts(accounts);
  }
}

/* ---------- Date rule helpers ---------- */
const todayIsFriday = (d = new Date()) => d.getDay() === 5;
function isAlternateFriday(date = new Date()) {
  const f = new Date(date.getFullYear(), 0, 1);
  while (f.getDay() !== 5) f.setDate(f.getDate() + 1);
  const weeks = Math.floor((date.getTime() - f.getTime()) / (7 * 24 * 3600 * 1000));
  return weeks % 2 === 0;
}

/* vendor rule by position */
function vendorRuleByIndex(index: number): "weekly" | "alternate" | "on-demand" {
  const pos = index + 1;
  if (pos >= 1 && pos <= 5) return "weekly";
  if (pos >= 6 && pos <= 10) return "alternate";
  if (pos >= 11 && pos <= 20) return "on-demand";
  return "weekly";
}

/* ---------- Small UI modals (no window.prompt/confirm/alert) ---------- */
function ensureModalContainer() {
  if (!$("modal-root")) {
    const div = document.createElement("div");
    div.id = "modal-root";
    document.body.appendChild(div);
    // basic styling
    const style = document.createElement("style");
    style.innerHTML = `
      #modal-root .modal-backdrop { position: fixed; inset:0; background: rgba(0,0,0,0.35); display:flex; align-items:center; justify-content:center; z-index:9999; }
      #modal-root .modal { background:#fff; padding:16px; border-radius:6px; min-width:300px; box-shadow: 0 6px 24px rgba(0,0,0,0.2); }
      #modal-root .modal h3 { margin:0 0 8px 0; font-size:16px; }
      #modal-root .modal .actions { margin-top:12px; text-align:right; }
      #modal-root .modal button { margin-left:8px; }
      #modal-root input[type="password"], #modal-root input[type="text"] { width:100%; padding:6px; box-sizing:border-box; margin-top:6px; }
    `;
    document.head.appendChild(style);
  }
}

function showPrompt(title: string, placeholder = "", type: "text" | "password" = "text"): Promise<string | null> {
  ensureModalContainer();
  return new Promise((resolve) => {
    const root = $("modal-root")!;
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop";
    const modal = document.createElement("div");
    modal.className = "modal";
    modal.innerHTML = `<h3>${escapeHtml(title)}</h3>`;
    const input = document.createElement("input");
    input.type = type;
    input.placeholder = placeholder;
    modal.appendChild(input);
    const actions = document.createElement("div");
    actions.className = "actions";
    const cancelBtn = document.createElement("button");
    cancelBtn.textContent = "Cancel";
    const okBtn = document.createElement("button");
    okBtn.textContent = "OK";
    actions.appendChild(cancelBtn);
    actions.appendChild(okBtn);
    modal.appendChild(actions);
    backdrop.appendChild(modal);
    root.appendChild(backdrop);
    input.focus();

    function cleanup(val: string | null) {
      root.removeChild(backdrop);
      resolve(val);
    }
    cancelBtn.onclick = () => cleanup(null);
    okBtn.onclick = () => cleanup(input.value || "");
    input.onkeydown = (e) => {
      if (e.key === "Enter") okBtn.click();
      if (e.key === "Escape") cancelBtn.click();
    };
  });
}

function showConfirm(title: string): Promise<boolean> {
  ensureModalContainer();
  return new Promise((resolve) => {
    const root = $("modal-root")!;
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop";
    const modal = document.createElement("div");
    modal.className = "modal";
    modal.innerHTML = `<h3>${escapeHtml(title)}</h3>`;
    const actions = document.createElement("div");
    actions.className = "actions";
    const noBtn = document.createElement("button");
    noBtn.textContent = "No";
    const yesBtn = document.createElement("button");
    yesBtn.textContent = "Yes";
    actions.appendChild(noBtn);
    actions.appendChild(yesBtn);
    modal.appendChild(actions);
    backdrop.appendChild(modal);
    root.appendChild(backdrop);

    function cleanup(val: boolean) {
      root.removeChild(backdrop);
      resolve(val);
    }
    noBtn.onclick = () => cleanup(false);
    yesBtn.onclick = () => cleanup(true);
  });
}

function showAlert(message: string): Promise<void> {
  ensureModalContainer();
  return new Promise((resolve) => {
    const root = $("modal-root")!;
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop";
    const modal = document.createElement("div");
    modal.className = "modal";
    modal.innerHTML = `<h3>${escapeHtml(message)}</h3>`;
    const actions = document.createElement("div");
    actions.className = "actions";
    const okBtn = document.createElement("button");
    okBtn.textContent = "OK";
    actions.appendChild(okBtn);
    modal.appendChild(actions);
    backdrop.appendChild(modal);
    root.appendChild(backdrop);
    okBtn.onclick = () => {
      root.removeChild(backdrop);
      resolve();
    };
  });
}

/* ---------- Initialization ---------- */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $('sideload-msg') && ($('sideload-msg')!.style.display = "none");
    $('app-body') && ($('app-body')!.style.display = "flex");

    ensureDefaultAccounts();
    ensureModalContainer();

    // existing bindings
    $('loginBtn')?.addEventListener('click', login);
    $('logoutBtn')?.addEventListener('click', logout);
    $('run')?.addEventListener('click', run);
    $('saveVendorBtn')?.addEventListener('click', (e) => { e.preventDefault(); saveVendor(); });

    ensureRunScheduledButton();
    ensurePendingContainer();
    attachVendorTableDelegation();

    const savedUser = localStorage.getItem("user");
    if (savedUser) {
      showAppSection(savedUser);
      normalizeVendorsAndRender();
    } else {
      showLoginSection();
    }
  }
});

/* ---------- Auth ---------- */
function login() {
  const username = ((document.getElementById("username") as HTMLInputElement) || { value: "" }).value;
  const password = ((document.getElementById("password") as HTMLInputElement) || { value: "" }).value;
  if (username === "admin" && password === "1234") {
    localStorage.setItem("user", username);
    showAppSection(username);
    normalizeVendorsAndRender();
  } else {
    showAlert("Invalid username or password");
  }
}
function logout() {
  localStorage.removeItem("user");
  showLoginSection();
}
function showAppSection(username: string) {
  $('login-section') && ($('login-section')!.style.display = "none");
  $('app-section') && ($('app-section')!.style.display = "block");
  const n = document.getElementById("user-name");
  if (n) n.innerText = username;
}
function showLoginSection() {
  $('login-section') && ($('login-section')!.style.display = "block");
  $('app-section') && ($('app-section')!.style.display = "none");
}

/* ---------- UI helpers to create missing elements ---------- */
function ensureRunScheduledButton() {
  if (!$("runScheduledBtn")) {
    const runBtn = document.createElement("button");
    runBtn.id = "runScheduledBtn";
    runBtn.textContent = "Run Scheduled Payments";
    runBtn.style.marginLeft = "8px";
    const runArea = document.querySelector("#app-section > h3");
    if (runArea && runArea.parentNode) runArea.parentNode.insertBefore(runBtn, runArea.nextSibling);
    else $('app-section')?.appendChild(runBtn);
    runBtn.addEventListener("click", async () => {
      try {
        await runScheduledPayments();
      } catch (err) {
        console.error("runScheduledPayments failed:", err);
        await showAlert("Scheduled run failed. See console.");
      }
    });
  }
}
function ensurePendingContainer() {
  if (!$("pendingPaymentsContainer")) {
    const div = document.createElement("div");
    div.id = "pendingPaymentsContainer";
    div.style.marginTop = "12px";
    $('app-section')?.appendChild(div);
  }
}

/* ---------- Vendor table delegation ---------- */
function attachVendorTableDelegation() {
  const tbody = document.querySelector("#vendorTable tbody") as HTMLElement | null;
  if (!tbody) return;
  tbody.onclick = (ev) => {
    const target = ev.target as HTMLElement | null;
    if (!target) return;

    const btn = target.closest("button") as HTMLButtonElement | null;
    if (btn) {
      const id = btn.dataset.id!;
      if (btn.classList.contains("edit-btn")) editVendor(id);
      else if (btn.classList.contains("pay-now-btn")) onDemandPay(id).catch(err => console.error("onDemandPay failed:", err));
      else if (btn.classList.contains("delete-btn")) (async () => {
        if (await showConfirm("Delete vendor and its pending payments?")) {
          await deleteVendor(id);
        }
      })();
      return;
    }

    const sel = target.closest("select") as HTMLSelectElement | null;
    if (sel && sel.classList.contains("change-account")) {
      const id = sel.dataset.id!;
      changeAssignedAccount(id, sel.value);
    }
  };
}

/* ---------- Normalize vendors & render ---------- */
function normalizeVendorsAndRender() {
  const vendors = getVendors();
  let changed = false;
  const normalized = vendors.map((v, idx) => {
    const copy = { ...v };
    if (!copy.id) { copy.id = getId(); changed = true; }
    if (!copy.baseAmount) { copy.baseAmount = DEFAULT_BASE_AMOUNT; changed = true; }
    if (typeof copy.skipNext === "undefined") { copy.skipNext = false; changed = true; }
    if (!copy.assignedAccount) {
      const rule = vendorRuleByIndex(idx);
      copy.assignedAccount = (rule === "on-demand") ? "Account 2" : "Account 1";
      changed = true;
    }
    return copy;
  });
  if (changed) setVendors(normalized);
  renderVendorTable();
  renderPendingList();
}

/* ---------- Render Vendor Table ---------- */
function renderVendorTable() {
  const tbody = document.querySelector("#vendorTable tbody") as HTMLTableSectionElement | null;
  if (!tbody) return;
  tbody.innerHTML = "";
  const vendors = getVendors();

  vendors.forEach((v, idx) => {
    const rule = vendorRuleByIndex(idx);
    const base = v.baseAmount ?? DEFAULT_BASE_AMOUNT;
    const paymentTypeLabel = v.paymentType || (rule === "on-demand" ? "On-Demand" : (rule === "alternate" ? "Biweekly" : "Weekly"));

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(v.name)}</td>
      <td>${escapeHtml(paymentTypeLabel)}</td>
      <td>
        <select class="change-account" data-id="${v.id}">
          <option value="Account 1"${v.assignedAccount==="Account 1"?" selected":""}>Account 1</option>
          <option value="Account 2"${v.assignedAccount==="Account 2"?" selected":""}>Account 2</option>
        </select>
      </td>
      <td>
        <button class="edit-btn" data-id="${v.id}">Edit</button>
        <button class="pay-now-btn" data-id="${v.id}">Pay Now</button>
        <button class="delete-btn" data-id="${v.id}">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/* ---------- Save / Edit / Delete Vendor ---------- */
function saveVendor() {
  const id = (document.getElementById("vendorId") as HTMLInputElement).value || getId();
  const name = (document.getElementById("vendorName") as HTMLInputElement).value.trim();
  const paymentType = (document.getElementById("paymentType") as HTMLSelectElement).value;
  const assignedAccount = (document.getElementById("assignedAccount") as HTMLSelectElement).value;

  if (!name) { showAlert("Vendor Name is required"); return; }

  const vendors = getVendors();
  const idx = vendors.findIndex(v => v.id === id);
  const vendor: Vendor = {
    id,
    name,
    paymentType,
    assignedAccount,
    baseAmount: vendors[idx]?.baseAmount ?? DEFAULT_BASE_AMOUNT,
    lastPaid: vendors[idx]?.lastPaid ?? null,
    skipNext: vendors[idx]?.skipNext ?? false
  };
  if (idx >= 0) vendors[idx] = vendor;
  else vendors.push(vendor);
  setVendors(vendors);
  clearVendorForm();
  normalizeVendorsAndRender();
  writeVendorsToSheet().catch(err => console.error("write after save failed:", err));
}

function editVendor(id: string) {
  const vendors = getVendors();
  const v = vendors.find(x => x.id === id);
  if (!v) return;
  (document.getElementById("vendorId") as HTMLInputElement).value = v.id;
  (document.getElementById("vendorName") as HTMLInputElement).value = v.name;
  (document.getElementById("paymentType") as HTMLSelectElement).value = v.paymentType;
  (document.getElementById("assignedAccount") as HTMLSelectElement).value = v.assignedAccount;
}

async function deleteVendor(id: string) {
  try {
    let vendors = getVendors();
    vendors = vendors.filter(v => v.id !== id);
    setVendors(vendors);

    let pending = getPending();
    pending = pending.filter(p => p.vendorId !== id);
    setPending(pending);

    normalizeVendorsAndRender();
    renderPendingList();
    await writeVendorsToSheet();
  } catch (err) {
    console.error("deleteVendor error:", err);
    await showAlert("Failed to delete vendor (see console).");
  }
}

function clearVendorForm() {
  (document.getElementById("vendorId") as HTMLInputElement).value = "";
  (document.getElementById("vendorName") as HTMLInputElement).value = "";
  (document.getElementById("paymentType") as HTMLSelectElement).value = "Weekly";
  (document.getElementById("assignedAccount") as HTMLSelectElement).value = "Account 1";
}

/* ---------- Account change ---------- */
function changeAssignedAccount(id: string, accountId: string) {
  const vendors = getVendors();
  const v = vendors.find(x => x.id === id);
  if (!v) return;
  v.assignedAccount = accountId;
  setVendors(vendors);
  writeVendorsToSheet().catch(err => console.error("write after account change failed:", err));
}

/* ---------- Payment Execution ---------- */
async function runScheduledPayments() {
  const vendors = getVendors();
  const today = new Date();

  if (!todayIsFriday(today)) {
    const ok = await showConfirm("Today is not Friday. Run scheduled logic anyway for testing?");
    if (!ok) return;
  }

  const due: { vendor: Vendor; amount: number }[] = [];
  vendors.forEach((v, idx) => {
    const rule = vendorRuleByIndex(idx);
    if (rule === "on-demand") return;
    if (v.skipNext) { v.skipNext = false; return; }
    if (!todayIsFriday(today)) return;
    if (rule === "weekly") due.push({ vendor: v, amount: v.baseAmount ?? DEFAULT_BASE_AMOUNT });
    else if (rule === "alternate" && isAlternateFriday(today)) due.push({ vendor: v, amount: (v.baseAmount ?? DEFAULT_BASE_AMOUNT) * 2 });
  });

  for (const entry of due) {
    await executePayment(entry.vendor, entry.amount);
  }

  setVendors(vendors);
  normalizeVendorsAndRender();
  renderPendingList();
  await writeVendorsToSheet();
  await showAlert("Scheduled run complete.");
}

async function executePayment(vendor: Vendor, amount: number): Promise<boolean> {
  try {
    const accounts = getAccounts();
    let account = accounts.find(a => a.id === vendor.assignedAccount) || accounts[0];
    if (!account) {
      pushPending(vendor, amount, "No account assigned");
      return false;
    }
    if (account.balance >= amount) {
      account.balance -= amount;
      setAccounts(accounts);
      const vendors = getVendors();
      const idx = vendors.findIndex(v => v.id === vendor.id);
      if (idx >= 0) { vendors[idx].lastPaid = nowISO(); setVendors(vendors); }
      console.log(`Paid ${vendor.name} $${amount} from ${account.id}`);
      return true;
    } else {
      pushPending(vendor, amount, "Insufficient funds");
      return false;
    }
  } catch (err) {
    console.error("executePayment error:", err);
    pushPending(vendor, amount, "Execution error");
    return false;
  }
}

/* push pending */
function pushPending(vendor: Vendor, amount: number, reason: string) {
  const pending = getPending();
  pending.push({
    id: getId(),
    vendorId: vendor.id,
    vendorName: vendor.name,
    amount,
    reason,
    createdAt: nowISO()
  });
  setPending(pending);
}

/* ---------- On-demand payment (manual) ---------- */
async function onDemandPay(vendorId: string) {
  const vendors = getVendors();
  const vendor = vendors.find(v => v.id === vendorId);
  if (!vendor) { await showAlert("Vendor not found"); return; }

  const pin = await showPrompt("Enter secure PIN to confirm on-demand payment:", "", "password");
  if (pin === null) return; // cancelled
  if (pin !== SECURE_PIN) { await showAlert("Invalid PIN"); return; }

  const idx = vendors.findIndex(v => v.id === vendorId);
  const rule = vendorRuleByIndex(idx);
  const amount = rule === "alternate" ? (vendor.baseAmount ?? DEFAULT_BASE_AMOUNT) * 2 : (vendor.baseAmount ?? DEFAULT_BASE_AMOUNT);

  const paid = await executePayment(vendor, amount);

  if (paid) {
    if (rule !== "on-demand") {
      const skip = await showConfirm("Vendor is scheduled. Skip next scheduled payment after this on-demand payment?");
      if (skip) {
        vendor.skipNext = true;
        setVendors(vendors);
      }
    }
    normalizeVendorsAndRender();
    renderPendingList();
    await writeVendorsToSheet();
    await showAlert("On-demand payment executed.");
  } else {
    renderPendingList();
    await showAlert("Payment queued to pending (insufficient funds or error).");
  }
}

/* ---------- Pending list UI & retry ---------- */
function renderPendingList() {
  const container = $('pendingPaymentsContainer');
  if (!container) return;
  const pending = getPending();
  if (!pending || pending.length === 0) {
    container.innerHTML = "<p>No pending payments</p>";
    return;
  }
  let html = `<h4>Pending Payments (${pending.length})</h4>`;
  html += `<table border="1" style="width:100%; border-collapse:collapse"><thead><tr><th>Vendor</th><th>Amount</th><th>Reason</th><th>Created</th><th>Action</th></tr></thead><tbody>`;
  pending.forEach(p => {
    html += `<tr>
      <td>${escapeHtml(p.vendorName)}</td>
      <td>${p.amount}</td>
      <td>${escapeHtml(p.reason)}</td>
      <td>${new Date(p.createdAt).toLocaleString()}</td>
      <td><button class="retry-btn" data-id="${p.id}">Retry</button></td>
    </tr>`;
  });
  html += `</tbody></table>`;
  container.innerHTML = html;

  container.querySelectorAll<HTMLButtonElement>('.retry-btn').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.id!;
      await retryPending(id);
    };
  });
}

async function retryPending(pendingId: string) {
  const pending = getPending();
  const item = pending.find(p => p.id === pendingId);
  if (!item) { await showAlert("Pending not found"); return; }
  const vendors = getVendors();
  const vendor = vendors.find(v => v.id === item.vendorId);
  if (!vendor) { await showAlert("Vendor not found"); return; }
  const success = await executePayment(vendor, item.amount);
  if (success) {
    setPending(pending.filter(p => p.id !== pendingId));
    renderPendingList();
    normalizeVendorsAndRender();
    await writeVendorsToSheet();
    await showAlert("Retry successful.");
  } else {
    await showAlert("Retry failed (still pending).");
  }
}

/* ---------- Excel write function ---------- */
async function writeVendorsToSheet() {
  const vendors = getVendors();
  try {
    if (typeof Excel === 'undefined' || typeof Office === 'undefined' || Office.context?.host !== Office.HostType.Excel) {
      console.warn("Excel API not available or not running in Excel host.");
      return;
    }
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      try { sheet.getRange("A1:C1000").clear(); } catch {}
      if (!vendors || vendors.length === 0) { await context.sync(); return; }

      const header = ['Vendor Name', 'Payment Type', 'Assigned Account'];
      const data = [header, ...vendors.map(v => [v.name, v.paymentType, v.assignedAccount])];
      const targetRange = sheet.getRangeByIndexes(0, 0, data.length, header.length);
      targetRange.values = data;
      try { targetRange.getEntireColumn().format.autofitColumns(); } catch {}
      await context.sync();
    });
  } catch (err) {
    console.error("writeVendorsToSheet error:", err);
  }
}

/* ---------- Utility ---------- */
function escapeHtml(str: string) {
  return (str || "").replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as any)[m]);
}


function run(this: HTMLElement, ev: PointerEvent) {
  throw new Error("Function not implemented.");
}
/* ---------- End ---------- */
