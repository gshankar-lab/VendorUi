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
interface PendingBankPayment {
  type: "scheduled" | "ondemand";
  amount: number;
  date: string;
}

let pendingBankPayments: PendingBankPayment[] = [];

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

        // Basic styling (no innerHTML)
        const style = document.createElement("style");
        style.textContent = `
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


function showPrompt(
    title: string,
    placeholder = "",
    type: "text" | "password" = "text"
): Promise<string | null> {
    ensureModalContainer();

    return new Promise((resolve) => {
        const root = $("modal-root")!;
        const backdrop = document.createElement("div");
        backdrop.className = "modal-backdrop";

        const modal = document.createElement("div");
        modal.className = "modal";

        // Title
        const h3 = document.createElement("h3");
        h3.textContent = title;
        modal.appendChild(h3);

        // Input field
        const input = document.createElement("input");
        input.type = type;
        input.placeholder = placeholder;
        modal.appendChild(input);

        // Actions container
        const actions = document.createElement("div");
        actions.className = "actions";

        // Cancel button
        const cancelBtn = document.createElement("button");
        cancelBtn.textContent = "Cancel";

        // OK button
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

        // Title
        const h3 = document.createElement("h3");
        h3.textContent = title;
        modal.appendChild(h3);

        // Actions container
        const actions = document.createElement("div");
        actions.className = "actions";

        // No button
        const noBtn = document.createElement("button");
        noBtn.textContent = "No";

        // Yes button
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

        // Message
        const h3 = document.createElement("h3");
        h3.textContent = message;
        modal.appendChild(h3);

        // Actions container
        const actions = document.createElement("div");
        actions.className = "actions";

        // OK button
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
    document.getElementById("generateReportBtn")?.addEventListener("click", async () => {
      try {
        await generateCurrentReport();
        await showAlert("Report generated successfully.");
      } catch (err) {
        console.error("Report generation failed:", err);
        await showAlert("Failed to generate report (see console).");
      }
    });
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
    tbody.textContent = ""; // safer than innerHTML for clearing

    const vendors = getVendors();

    vendors.forEach((v, idx) => {
        const rule = vendorRuleByIndex(idx);
        const base = v.baseAmount ?? DEFAULT_BASE_AMOUNT;
        const paymentTypeLabel =
            v.paymentType ||
            (rule === "on-demand"
                ? "On-Demand"
                : rule === "alternate"
                ? "Biweekly"
                : "Weekly");

        const tr = document.createElement("tr");

        // Name cell
        const nameTd = document.createElement("td");
        nameTd.textContent = v.name;
        tr.appendChild(nameTd);

        // Payment type cell
        const typeTd = document.createElement("td");
        typeTd.textContent = paymentTypeLabel;
        tr.appendChild(typeTd);

        // Account select cell
        const accountTd = document.createElement("td");
        const select = document.createElement("select");
        select.className = "change-account";
        select.dataset.id = v.id;

        const account1Option = document.createElement("option");
        account1Option.value = "Account 1";
        account1Option.textContent = "Account 1";
        if (v.assignedAccount === "Account 1") {
            account1Option.selected = true;
        }

        const account2Option = document.createElement("option");
        account2Option.value = "Account 2";
        account2Option.textContent = "Account 2";
        if (v.assignedAccount === "Account 2") {
            account2Option.selected = true;
        }

        select.appendChild(account1Option);
        select.appendChild(account2Option);
        accountTd.appendChild(select);
        tr.appendChild(accountTd);

        // Actions cell
        const actionsTd = document.createElement("td");

        const editBtn = document.createElement("button");
        editBtn.className = "edit-btn";
        editBtn.dataset.id = v.id;
        editBtn.textContent = "Edit";

        const payNowBtn = document.createElement("button");
        payNowBtn.className = "pay-now-btn";
        payNowBtn.dataset.id = v.id;
        payNowBtn.textContent = "Pay Now";

        const deleteBtn = document.createElement("button");
        deleteBtn.className = "delete-btn";
        deleteBtn.dataset.id = v.id;
        deleteBtn.textContent = "Delete";

        actionsTd.appendChild(editBtn);
        actionsTd.appendChild(payNowBtn);
        actionsTd.appendChild(deleteBtn);
        tr.appendChild(actionsTd);

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
    const container = $("pendingPaymentsContainer");
    if (!container) return;

    // Clear container safely
    container.textContent = "";

    const pending = getPending();
    if (!pending || pending.length === 0) {
        const noMsg = document.createElement("p");
        noMsg.textContent = "No pending payments";
        container.appendChild(noMsg);
        return;
    }

    // Heading
    const heading = document.createElement("h4");
    heading.textContent = `Pending Payments (${pending.length})`;
    container.appendChild(heading);

    // Table
    const table = document.createElement("table");
    table.style.width = "100%";
    table.style.borderCollapse = "collapse";
    table.border = "1";

    // Table Head
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["Vendor", "Amount", "Reason", "Created", "Action"].forEach(text => {
        const th = document.createElement("th");
        th.textContent = text;
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    // Table Body
    const tbody = document.createElement("tbody");
    pending.forEach(p => {
        const row = document.createElement("tr");

        const vendorTd = document.createElement("td");
        vendorTd.textContent = p.vendorName;
        row.appendChild(vendorTd);

        const amountTd = document.createElement("td");
        amountTd.textContent = p.amount.toString();
        row.appendChild(amountTd);

        const reasonTd = document.createElement("td");
        reasonTd.textContent = p.reason;
        row.appendChild(reasonTd);

        const createdTd = document.createElement("td");
        createdTd.textContent = new Date(p.createdAt).toLocaleString();
        row.appendChild(createdTd);

        const actionTd = document.createElement("td");
        const retryBtn = document.createElement("button");
        retryBtn.className = "retry-btn";
        retryBtn.dataset.id = p.id;
        retryBtn.textContent = "Retry";
        retryBtn.onclick = async () => {
            await retryPending(p.id);
        };
        actionTd.appendChild(retryBtn);
        row.appendChild(actionTd);

        tbody.appendChild(row);
    });

    table.appendChild(tbody);
    container.appendChild(table);
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
      try { sheet.getRange("A1:C1000").clear(); } catch { }
      if (!vendors || vendors.length === 0) { await context.sync(); return; }

      const header = ['Vendor Name', 'Payment Type', 'Assigned Account'];
      const data = [header, ...vendors.map(v => [v.name, v.paymentType, v.assignedAccount])];
      const targetRange = sheet.getRangeByIndexes(0, 0, data.length, header.length);
      targetRange.values = data;
      try { targetRange.getEntireColumn().format.autofitColumns(); } catch { }
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

/* global Excel, Office */

let account1Balance = 200000;
let account2Balance = 200000;

Office.onReady(() => {
  document.getElementById("simulatePayment").onclick = simulatePayment;
  updateTaskPaneBalances();
});

// Simulate a payment
function simulatePayment() {
  const amountInput = document.getElementById("paymentAmount") as HTMLInputElement;
  const typeInput = document.getElementById("paymentTypeData") as HTMLSelectElement;

  const paymentAmount = parseFloat(amountInput.value);
  const paymentTypeData = typeInput.value as "scheduled" | "ondemand";

  if (isNaN(paymentAmount) || paymentAmount <= 0) {
    showMessage("Please enter a valid payment amount.");
    return;
  }

  if (paymentTypeData !== "scheduled" && paymentTypeData !== "ondemand") {
    showMessage("Please select a valid payment type.");
    return;
  }

  if (paymentTypeData === "scheduled") {
    if (account1Balance >= paymentAmount) {
      account1Balance -= paymentAmount;
      showMessage(`Payment of $${paymentAmount.toFixed(2)} processed from Account 1.`);
    } else {
      pendingBankPayments.push({
        type: paymentTypeData,
        amount: paymentAmount,
        date: new Date().toISOString(),
      });
      showMessage(`Insufficient funds in Account 1. Payment added to pending list.`);
    }
  } else {
    if (account2Balance >= paymentAmount) {
      account2Balance -= paymentAmount;
      showMessage(`Payment of $${paymentAmount.toFixed(2)} processed from Account 2.`);
    } else {
      pendingBankPayments.push({
        type: paymentTypeData,
        amount: paymentAmount,
        date: new Date().toISOString(),
      });
      showMessage(`Insufficient funds in Account 2. Payment added to pending list.`);
    }
  }

  updateTaskPaneBalances();
  updateExcelBalances();
}

// Update balances in the task pane UI
function updateTaskPaneBalances() {
  document.getElementById("account1Balance").innerText = `$${account1Balance.toLocaleString()}`;
  document.getElementById("account2Balance").innerText = `$${account2Balance.toLocaleString()}`;
}

// Write balances to Excel
async function updateExcelBalances() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A1").values = [["Account 1 Balance"]];
    sheet.getRange("B1").values = [["Account 2 Balance"]];
    sheet.getRange("A2").values = [[account1Balance]];
    sheet.getRange("B2").values = [[account2Balance]];
    await context.sync();
  });
}

function processPendingPayments() {
  let processedCount = 0;

  pendingBankPayments = pendingBankPayments.filter(payment => {
    if (payment.type === "scheduled" && account1Balance >= payment.amount) {
      account1Balance -= payment.amount;
      processedCount++;
      return false; // remove from list
    }
    if (payment.type === "ondemand" && account2Balance >= payment.amount) {
      account2Balance -= payment.amount;
      processedCount++;
      return false; // remove from list
    }
    return true; // keep in list
  });

  if (processedCount > 0) {
    showMessage(`Processed ${processedCount} pending payment(s).`);
  } else {
    showMessage("No pending payments could be processed.");
  }

  updateTaskPaneBalances();
  updateExcelBalances();
}

/* ---------- End ---------- */
async function generateCurrentReport() {
  // Get accounts, vendors, and completed payments
  const accounts = getAccounts();
  const vendors = getVendors();
  const timestamp = new Date();

  // Prepare completed payment data (vendor names, payment dates, and amounts)
  const completedPayments = vendors
    .filter(v => v.lastPaid)
    .map(v => [v.name, v.lastPaid, v.baseAmount ?? DEFAULT_BASE_AMOUNT]);

  // Use Excel.run to write the data into Excel
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add("Current Report");

    // Section 1: Account Balances
    const accountHeader = [["Account Name", "Balance"]];
    const accountData = accounts.map(a => [a.name, `$${a.balance.toFixed(2)}`]);
    const accountRange = sheet.getRangeByIndexes(0, 0, accountHeader.length + accountData.length, 2);
    accountRange.values = [...accountHeader, ...accountData];

    // Section 2: Completed Payments (Vendor Name, Payment Date, Amount)
    const paymentStartRow = accountData.length + 3;
    const paymentHeader = [["Vendor Name", "Payment Date", "Amount"]];
    const paymentRange = sheet.getRangeByIndexes(paymentStartRow, 0, paymentHeader.length + completedPayments.length, 3);
    paymentRange.values = [...paymentHeader, ...completedPayments];

    // Section 3: Timestamp of Report Generation
    const timestampRow = paymentStartRow + completedPayments.length + 3;
    const tsCell = sheet.getRange(`A${timestampRow + 1}`);
    tsCell.values = [[`Report generated: ${timestamp.toLocaleString()}`]];

    // Formatting (optional)
    try {
      sheet.getUsedRange().format.autofitColumns();
    } catch { }

    sheet.activate();
    await context.sync();
  });
}

// show message
function showMessage(message: string) {
  console.log(message); // for debugging in console
  const statusDiv = document.getElementById("statusMessage");
  if (statusDiv) {
    statusDiv.textContent = message;
  }
}


