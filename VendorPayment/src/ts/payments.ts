interface Vendor {
    "Vendor Name": string;
    "Payment Type": string;
    "Assigned Account": string;
    status?: string;
    skipNextScheduled?: boolean;
}

interface Account {
    id: string;
    name: string;
    balance: number;
}

let accounts: Account[] = [
    { id: "acc1", name: "Account 1", balance: 1000 },
    { id: "acc2", name: "Account 2", balance: 500 }
];

let pendingPayments: Vendor[] = [];

// ------------------ Modal Helpers ------------------
function showModalMessage(message: string): void {
    const modal = document.getElementById("messageModal")!;
    const modalText = document.getElementById("modalText")!;
    modalText.textContent = message;
    modal.classList.remove("hidden");
}

function showConfirm(message: string, onYes: () => void): void {
    const modal = document.getElementById("confirmModal")!;
    const confirmText = document.getElementById("confirmText")!;
    confirmText.textContent = message;
    modal.classList.remove("hidden");

    document.getElementById("confirmYesBtn")!.onclick = () => {
        modal.classList.add("hidden");
        onYes();
    };
    document.getElementById("confirmNoBtn")!.onclick = () => {
        modal.classList.add("hidden");
    };
}

// ------------------ Balances UI ------------------
function updateBalancesUI(): void {
   document.getElementById("acc1Balance")!.textContent =
    accounts[0].balance.toLocaleString("en-US", { style: "currency", currency: "USD" });
document.getElementById("acc2Balance")!.textContent =
    accounts[1].balance.toLocaleString("en-US", { style: "currency", currency: "USD" });

}

// ------------------ Payment Logic ------------------
function isPaydayForVendor(index: number): boolean {
    const today = new Date();
    const isFriday = today.getDay() === 5;

    if (!isFriday) return false;

    if (index < 5) {
        return true; // Vendors 1-5 weekly
    }
    if (index >= 5 && index < 10) {
        const weekNumber = getWeek(today);
        return weekNumber % 2 === 0; // Vendors 6-10 biweekly
    }
    return false;
}

function getWeek(date: Date): number {
    const firstDay = new Date(date.getFullYear(), 0, 1);
    const days = Math.floor((+date - +firstDay) / (24 * 60 * 60 * 1000));
    return Math.ceil((days + firstDay.getDay() + 1) / 7);
}

function processScheduledPayments(): void {
    const vendors: Vendor[] = (window as any).getVendors();

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
            } else {
                vendor.status = "Pending (Insufficient funds)";
                pendingPayments.push(vendor);
            }
        }
    });

    (window as any).saveVendors(vendors);
    updateBalancesUI();
    renderPayments();
    showModalMessage("Scheduled payments processed.");
}

function triggerOnDemandPayment(): void {
    showConfirm("Trigger on-demand payments now?", () => {
        const vendors: Vendor[] = (window as any).getVendors();

        vendors.forEach((vendor, index) => {
            if (vendor["Payment Type"] === "On-Demand") {
                const amount = 150;
                const account = accounts[1];
                if (account.balance >= amount) {
                    account.balance -= amount;
                    vendor.status = `On-demand paid $${amount}`;
                } else {
                    vendor.status = "Pending (Insufficient funds)";
                    pendingPayments.push(vendor);
                }
            } else {
                // Scheduled vendor paid on-demand
                const amount = 150;
                const account = accounts[1];
                if (account.balance >= amount) {
                    account.balance -= amount;
                    vendor.status = `On-demand paid $${amount}`;
                    showConfirm(`Skip next scheduled payment for ${vendor["Vendor Name"]}?`, () => {
                        vendor.skipNextScheduled = true;
                        (window as any).saveVendors(vendors);
                    });
                } else {
                    vendor.status = "Pending (Insufficient funds)";
                    pendingPayments.push(vendor);
                }
            }
        });

        (window as any).saveVendors(vendors);
        updateBalancesUI();
        renderPayments();
        showModalMessage("On-demand payments processed.");
    });
}

// ------------------ Rendering ------------------
function renderPayments(): void {
    const tbody = document.getElementById("paymentsTableBody")!;
    tbody.innerHTML = "";

    const vendors: Vendor[] = (window as any).getVendors();
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

    const pendingList = document.getElementById("pendingList")!;
    pendingList.innerHTML = pendingPayments.map(v => `<li>${v["Vendor Name"]} - ${v.status}</li>`).join("");
}

// ------------------ Init ------------------
document.addEventListener("DOMContentLoaded", () => {
    updateBalancesUI();
    renderPayments();

    document.getElementById("processScheduledBtn")!
        .addEventListener("click", processScheduledPayments);

    document.getElementById("triggerOnDemandBtn")!
        .addEventListener("click", triggerOnDemandPayment);

    document.getElementById("modalOkBtn")!
        .addEventListener("click", () => {
            document.getElementById("messageModal")!.classList.add("hidden");
        });

    document.getElementById("confirmNoBtn")!
        .addEventListener("click", () => {
            document.getElementById("confirmModal")!.classList.add("hidden");
        });
});
document.addEventListener("DOMContentLoaded", () => {
    if (document.getElementById("paymentsTableBody")) {
        renderPayments();
    }
});