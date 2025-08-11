"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
let accounts = [
    { id: "acc1", name: "Account 1", balance: 200000 },
    { id: "acc2", name: "Account 2", balance: 200000 }
];
let pendingPayments = [];
// ------------------ Modal Helpers ------------------
function showModalMessage(message) {
    const modal = document.getElementById("messageModal");
    const modalText = document.getElementById("modalText");
    if (!modal || !modalText)
        return;
    modalText.textContent = message;
    modal.classList.remove("hidden");
}
function showConfirm(message, onYes) {
    const modal = document.getElementById("confirmModal");
    const confirmText = document.getElementById("confirmText");
    if (!modal || !confirmText)
        return;
    confirmText.textContent = message;
    modal.classList.remove("hidden");
    const yesBtn = document.getElementById("confirmYesBtn");
    const noBtn = document.getElementById("confirmNoBtn");
    if (yesBtn)
        yesBtn.onclick = () => {
            modal.classList.add("hidden");
            onYes();
        };
    if (noBtn)
        noBtn.onclick = () => {
            modal.classList.add("hidden");
        };
}
// ------------------ Balances & Excel Sync ------------------
function updateBalancesUI() {
    const acc1El = document.getElementById("acc1Balance");
    const acc2El = document.getElementById("acc2Balance");
    if (acc1El)
        acc1El.textContent = accounts[0].balance.toLocaleString("en-US", { style: "currency", currency: "USD" });
    if (acc2El)
        acc2El.textContent = accounts[1].balance.toLocaleString("en-US", { style: "currency", currency: "USD" });
    // Update Excel if available
    if (typeof Excel !== "undefined") {
        Excel.run((context) => __awaiter(this, void 0, void 0, function* () {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            // Balances
            sheet.getRange("A1").values = [["Account"]];
            sheet.getRange("B1").values = [["Balance"]];
            sheet.getRange("A2").values = [[accounts[0].name]];
            sheet.getRange("B2").values = [[accounts[0].balance]];
            sheet.getRange("A3").values = [[accounts[1].name]];
            sheet.getRange("B3").values = [[accounts[1].balance]];
            // Pending payments table
            sheet.getRange("D1").values = [["Pending Payments"]];
            sheet.getRange("D2").values = [["Vendor Name"]];
            sheet.getRange("E2").values = [["Status"]];
            if (pendingPayments.length > 0) {
                const vendorNames = pendingPayments.map(v => [v["Vendor Name"]]);
                const statuses = pendingPayments.map(v => [v.status || ""]);
                sheet.getRange(`D3:D${2 + vendorNames.length}`).values = vendorNames;
                sheet.getRange(`E3:E${2 + statuses.length}`).values = statuses;
            }
            else {
                sheet.getRange("D3:E100").clear();
            }
            yield context.sync();
        })).catch(err => console.error("Excel update failed:", err));
    }
}
// ------------------ Payment Logic ------------------
function isPaydayForVendor(index) {
    const today = new Date();
    const isFriday = today.getDay() === 5;
    if (!isFriday)
        return false;
    if (index < 5)
        return true; // Vendors 1–5 weekly
    if (index >= 5 && index < 10) {
        const weekNumber = getWeek(today);
        return weekNumber % 2 === 0; // Vendors 6–10 biweekly
    }
    return false;
}
function getWeek(date) {
    const firstDay = new Date(date.getFullYear(), 0, 1);
    const days = Math.floor((+date - +firstDay) / (24 * 60 * 60 * 1000));
    return Math.ceil((days + firstDay.getDay() + 1) / 7);
}
function processScheduledPayments() {
    const vendors = window.getVendors();
    vendors.forEach((vendor, index) => {
        if (vendor["Payment Type"] !== "On-Demand" && isPaydayForVendor(index)) {
            if (vendor.skipNextScheduled) {
                vendor.skipNextScheduled = false;
                vendor.status = "Skipped (On-demand paid)";
                return;
            }
            const amount = index < 5 ? 100 : 200;
            const account = accounts[0];
            if (account.balance >= amount) {
                account.balance -= amount;
                vendor.status = `Paid $${amount}`;
            }
            else {
                vendor.status = "Pending (Insufficient funds)";
                pendingPayments.push(vendor);
            }
        }
    });
    window.saveVendors(vendors);
    updateBalancesUI();
    renderPayments();
    showModalMessage("Scheduled payments processed.");
}
function triggerOnDemandPayment() {
    showConfirm("Trigger on-demand payments now?", () => {
        const vendors = window.getVendors();
        vendors.forEach((vendor) => {
            const amount = 150;
            const account = accounts[1];
            if (account.balance >= amount) {
                account.balance -= amount;
                vendor.status = `On-demand paid $${amount}`;
                if (vendor["Payment Type"] !== "On-Demand") {
                    showConfirm(`Skip next scheduled payment for ${vendor["Vendor Name"]}?`, () => {
                        vendor.skipNextScheduled = true;
                        window.saveVendors(vendors);
                    });
                }
            }
            else {
                vendor.status = "Pending (Insufficient funds)";
                pendingPayments.push(vendor);
            }
        });
        window.saveVendors(vendors);
        updateBalancesUI();
        renderPayments();
        showModalMessage("On-demand payments processed.");
    });
}
// ------------------ Report Generation ------------------
function generateCurrentReport() {
    if (typeof Excel === "undefined") {
        alert("Excel is not available in this environment.");
        return;
    }
    Excel.run((context) => __awaiter(this, void 0, void 0, function* () {
        const sheet = context.workbook.worksheets.add("Current Report");
        // Report timestamp
        sheet.getRange("A1").values = [["Report Generated"]];
        sheet.getRange("B1").values = [[new Date().toLocaleString()]];
        // Account balances
        sheet.getRange("A3").values = [["Account"]];
        sheet.getRange("B3").values = [["Balance"]];
        sheet.getRange("A4").values = [[accounts[0].name]];
        sheet.getRange("B4").values = [[accounts[0].balance]];
        sheet.getRange("A5").values = [[accounts[1].name]];
        sheet.getRange("B5").values = [[accounts[1].balance]];
        // Completed payments
        const vendors = window.getVendors();
        sheet.getRange("A7").values = [["Vendor Name", "Status", "Payment Date"]];
        const completedRows = vendors
            .filter(v => { var _a; return v.status && v.status.startsWith("Paid") || ((_a = v.status) === null || _a === void 0 ? void 0 : _a.startsWith("On-demand")); })
            .map(v => [v["Vendor Name"], v.status, new Date().toLocaleDateString()]);
        if (completedRows.length > 0) {
            sheet.getRange(`A8:C${7 + completedRows.length}`).values = completedRows;
        }
        yield context.sync();
        showModalMessage("Current report generated in Excel.");
    })).catch(err => console.error("Report generation failed:", err));
}
// ------------------ Rendering ------------------
function renderPayments() {
    const tbody = document.getElementById("paymentsTableBody");
    if (!tbody)
        return;
    tbody.innerHTML = "";
    const vendors = window.getVendors();
    vendors.forEach((vendor) => {
        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${vendor["Vendor Name"]}</td>
            <td>${vendor["Payment Type"]}</td>
            <td>${vendor["Assigned Account"]}</td>
            <td>${vendor.status || "Pending"}</td>
        `;
        tbody.appendChild(row);
    });
    const pendingList = document.getElementById("pendingList");
    if (pendingList) {
        pendingList.innerHTML = pendingPayments.map(v => `<li>${v["Vendor Name"]} - ${v.status}</li>`).join("");
    }
}
// ------------------ Init ------------------
// ------------------ Init ------------------
document.addEventListener("DOMContentLoaded", () => {
    updateBalancesUI();
    renderPayments();
    const processBtn = document.getElementById("processScheduledBtn");
    if (processBtn)
        processBtn.addEventListener("click", processScheduledPayments);
    const onDemandBtn = document.getElementById("triggerOnDemandBtn");
    if (onDemandBtn)
        onDemandBtn.addEventListener("click", triggerOnDemandPayment);
    const reportBtn = document.getElementById("generateReportBtn");
    if (reportBtn) {
        reportBtn.addEventListener("click", () => {
            console.log("tr");
            if (typeof Excel === "undefined") {
                console.error("Excel API not available. Make sure this is running inside Excel.");
                showModalMessage("Excel is not available. Please open this add-in inside Excel.");
                return;
            }
            console.log("Generating current report...");
            generateCurrentReport();
        });
    }
    const modalOk = document.getElementById("modalOkBtn");
    if (modalOk)
        modalOk.addEventListener("click", () => {
            const modal = document.getElementById("messageModal");
            if (modal)
                modal.classList.add("hidden");
        });
    const confirmNo = document.getElementById("confirmNoBtn");
    if (confirmNo)
        confirmNo.addEventListener("click", () => {
            const modal = document.getElementById("confirmModal");
            if (modal)
                modal.classList.add("hidden");
        });
});
