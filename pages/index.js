import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── MOCK DATA ────────────────────────────────────────────────────────────────
const SALESPEOPLE = [
  { id: "sp1", name: "Rajesh Sharma", email: "rajesh@company.com", phone: "9876543210", avatar: "RS", target: 2500000 },
  { id: "sp2", name: "Priya Mehta", email: "priya@company.com", phone: "9876543211", avatar: "PM", target: 2000000 },
  { id: "sp3", name: "Amit Patel", email: "amit@company.com", phone: "9876543212", avatar: "AP", target: 1800000 },
  { id: "sp4", name: "Sunita Rao", email: "sunita@company.com", phone: "9876543213", avatar: "SR", target: 1500000 },
];

const generateInvoices = () => {
  const parties = [
    { name: "Gujarat Traders Pvt Ltd", gst: "24AABCG1234A1Z5", phone: "9998887770", email: "accounts@gujarattraders.com" },
    { name: "Mehta Industries", gst: "24AABHM5678B2Z6", phone: "9998887771", email: "finance@mehtaind.com" },
    { name: "Patel Enterprises", gst: "24AABCP9012C3Z7", phone: "9998887772", email: "billing@patelent.in" },
    { name: "Surat Fabrics Co.", gst: "24AABCS3456D4Z8", phone: "9998887773", email: "accounts@suratfabrics.com" },
    { name: "Rajkot Steel Works", gst: "24AABRR7890E5Z9", phone: "9998887774", email: "rswaccts@gmail.com" },
    { name: "Bhavnagar Chemicals", gst: "24AABCB2345F6Z0", phone: "9998887775", email: "cfo@bhavnagarchemicals.com" },
    { name: "Vadodara Plastics Ltd", gst: "24AABCV6789G7Z1", phone: "9998887776", email: "vplfinance@vpl.com" },
    { name: "Anand Dairy Supplies", gst: "24AABCA0123H8Z2", phone: "9998887777", email: "anandds@gmail.com" },
    { name: "Junagadh Agro Foods", gst: "24AABCJ4567I9Z3", phone: "9998887778", email: "jafaccts@jafoods.in" },
    { name: "Navsari Paper Mills", gst: "24AABCN8901J0Z4", phone: "9998887779", email: "npmill@papermill.com" },
  ];

  const today = new Date();
  const invoices = [];
  let invNum = 1001;

  parties.forEach((party, pi) => {
    const spId = SALESPEOPLE[pi % SALESPEOPLE.length].id;
    const count = 2 + Math.floor(Math.random() * 4);
    for (let i = 0; i < count; i++) {
      const daysOffset = -60 + Math.floor(Math.random() * 90);
      const invDate = new Date(today); invDate.setDate(today.getDate() - 30 - Math.floor(Math.random() * 60));
      const dueDate = new Date(invDate); dueDate.setDate(invDate.getDate() + 30 + Math.floor(Math.random() * 30));
      const amount = 50000 + Math.floor(Math.random() * 450000);
      const isPaid = Math.random() < 0.3;
      const isPartial = !isPaid && Math.random() < 0.2;
      const paid = isPaid ? amount : isPartial ? Math.floor(amount * (0.2 + Math.random() * 0.5)) : 0;
      const overdueDays = dueDate < today && !isPaid ? Math.floor((today - dueDate) / 86400000) : 0;

      invoices.push({
        id: `INV-${invNum++}`,
        party: party.name, gst: party.gst, phone: party.phone, email: party.email,
        salespersonId: spId,
        invoiceDate: invDate.toISOString().split("T")[0],
        dueDate: dueDate.toISOString().split("T")[0],
        amount, paid, outstanding: amount - paid,
        status: isPaid ? "paid" : isPartial ? "partial" : dueDate < today ? "overdue" : "pending",
        overdueDays,
        remindersSent: Math.floor(Math.random() * 4),
        lastReminder: overdueDays > 0 ? new Date(today.getTime() - Math.random() * 7 * 86400000).toISOString().split("T")[0] : null,
        notes: "",
      });
    }
  });
  return invoices;
};

const INITIAL_INVOICES = [];

const REMINDER_TEMPLATES = [
  { id: "rt1", name: "7 Days Before Due", trigger: "before_due", days: 7, channel: ["whatsapp", "email"], active: true },
  { id: "rt2", name: "1 Day Before Due", trigger: "before_due", days: 1, channel: ["whatsapp", "sms"], active: true },
  { id: "rt3", name: "Overdue 1 Day", trigger: "after_due", days: 1, channel: ["whatsapp", "email", "sms"], active: true },
  { id: "rt4", name: "Overdue 7 Days", trigger: "after_due", days: 7, channel: ["whatsapp", "email"], active: true },
  { id: "rt5", name: "Overdue 15 Days", trigger: "after_due", days: 15, channel: ["whatsapp"], active: false },
  { id: "rt6", name: "Overdue 30 Days", trigger: "after_due", days: 30, channel: ["whatsapp", "email", "sms"], active: true },
];

const REMINDER_HISTORY = [
  { id: "rh1", invoiceId: "INV-1001", party: "Gujarat Traders Pvt Ltd", channel: "whatsapp", sentAt: "2024-01-15 10:30", status: "delivered", template: "Overdue 7 Days" },
  { id: "rh2", invoiceId: "INV-1003", party: "Mehta Industries", channel: "email", sentAt: "2024-01-15 10:31", status: "opened", template: "7 Days Before Due" },
  { id: "rh3", invoiceId: "INV-1005", party: "Patel Enterprises", channel: "sms", sentAt: "2024-01-14 14:00", status: "delivered", template: "Overdue 1 Day" },
  { id: "rh4", invoiceId: "INV-1007", party: "Surat Fabrics Co.", channel: "whatsapp", sentAt: "2024-01-14 09:00", status: "read", template: "Overdue 30 Days" },
  { id: "rh5", invoiceId: "INV-1009", party: "Rajkot Steel Works", channel: "email", sentAt: "2024-01-13 11:00", status: "bounced", template: "7 Days Before Due" },
];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt = (n) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(n);
const fmtNum = (n) => new Intl.NumberFormat("en-IN").format(n);
const fmtShort = (n) => n >= 1e7 ? `₹${(n / 1e7).toFixed(2)}Cr` : n >= 1e5 ? `₹${(n / 1e5).toFixed(1)}L` : `₹${fmtNum(n)}`;

const statusConfig = {
  paid: { label: "Paid", color: "#10b981", bg: "#d1fae5", icon: "✓" },
  partial: { label: "Partial", color: "#f59e0b", bg: "#fef3c7", icon: "◑" },
  pending: { label: "Pending", color: "#6366f1", bg: "#e0e7ff", icon: "◷" },
  overdue: { label: "Overdue", color: "#ef4444", bg: "#fee2e2", icon: "!" },
};

const channelIcons = { whatsapp: "💬", email: "✉️", sms: "📱" };
const reminderStatusColors = { delivered: "#10b981", opened: "#6366f1", read: "#3b82f6", bounced: "#ef4444", sent: "#f59e0b" };

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function TallyCRM() {
  const [user, setUser] = useState(null);
  const [page, setPage] = useState("login");
  const [invoices, setInvoices] = useState(INITIAL_INVOICES);
  const [reminderTemplates, setReminderTemplates] = useState(REMINDER_TEMPLATES);
  const [reminderHistory] = useState(REMINDER_HISTORY);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [toast, setToast] = useState(null);
  const [importModal, setImportModal] = useState(false);
  const [selectedInvoice, setSelectedInvoice] = useState(null);
  const [reminderModal, setReminderModal] = useState(null);

  const showToast = useCallback((msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }, []);

  const handleLogin = (role, spId = null) => {
    setUser({ role, spId, name: role === "admin" ? "Admin User" : SALESPEOPLE.find(s => s.id === spId)?.name });
    setPage(role === "admin" ? "admin-dashboard" : "sp-dashboard");
  };

  const handleImport = (data) => {
    const newInvoices = data.map((row, i) => ({
      id: `IMP-${Date.now()}-${i}`,
      party: row["Party Name"] || row["party"] || "Unknown",
      gst: row["GST"] || row["gst"] || "",
      phone: row["Phone"] || row["phone"] || "",
      email: row["Email"] || row["email"] || "",
      salespersonId: SALESPEOPLE[i % SALESPEOPLE.length].id,
      invoiceDate: row["Invoice Date"] || row["invoiceDate"] || new Date().toISOString().split("T")[0],
      dueDate: row["Due Date"] || row["dueDate"] || new Date().toISOString().split("T")[0],
      amount: parseFloat(row["Amount"] || row["amount"] || 0),
      paid: parseFloat(row["Paid"] || row["paid"] || 0),
      outstanding: parseFloat(row["Amount"] || 0) - parseFloat(row["Paid"] || 0),
      status: "pending", overdueDays: 0, remindersSent: 0, lastReminder: null, notes: "",
    }));
    setInvoices(prev => {
      const existing = new Set(prev.map(i => i.id));
      const unique = newInvoices.filter(n => !existing.has(n.id));
      showToast(`Imported ${unique.length} new records (${newInvoices.length - unique.length} duplicates skipped)`);
      return [...prev, ...unique];
    });
    setImportModal(false);
  };

  const sendReminder = (invoice, channels) => {
    setInvoices(prev => prev.map(inv => inv.id === invoice.id
      ? { ...inv, remindersSent: inv.remindersSent + 1, lastReminder: new Date().toISOString().split("T")[0] }
      : inv
    ));
    showToast(`Reminder sent via ${channels.join(", ")} to ${invoice.party}`);
    setReminderModal(null);
  };

  if (!user) return <LoginPage onLogin={handleLogin} />;

  return (
    <div style={{ display: "flex", height: "100vh", background: "#0f1117", fontFamily: "'DM Sans', 'Segoe UI', sans-serif", color: "#e2e8f0", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Space+Grotesk:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #1a1d27; }
        ::-webkit-scrollbar-thumb { background: #2d3148; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: #4c5380; }
        .nav-item { transition: all 0.2s; cursor: pointer; }
        .nav-item:hover { background: rgba(99,102,241,0.15) !important; }
        .nav-item.active { background: rgba(99,102,241,0.25) !important; border-left: 3px solid #6366f1 !important; }
        .card { background: #1a1d27; border: 1px solid #2d3148; border-radius: 12px; transition: box-shadow 0.2s; }
        .card:hover { box-shadow: 0 4px 20px rgba(99,102,241,0.1); }
        .btn-primary { background: linear-gradient(135deg, #6366f1, #8b5cf6); color: white; border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 14px; transition: opacity 0.2s; }
        .btn-primary:hover { opacity: 0.9; }
        .btn-secondary { background: #2d3148; color: #a5b4fc; border: 1px solid #3d4270; padding: 8px 16px; border-radius: 8px; cursor: pointer; font-size: 13px; transition: all 0.2s; }
        .btn-secondary:hover { background: #3d4270; }
        .btn-danger { background: rgba(239,68,68,0.15); color: #f87171; border: 1px solid rgba(239,68,68,0.3); padding: 8px 16px; border-radius: 8px; cursor: pointer; font-size: 13px; }
        .badge { display: inline-flex; align-items: center; gap: 4px; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
        .input { background: #1a1d27; border: 1px solid #2d3148; color: #e2e8f0; padding: 10px 14px; border-radius: 8px; font-size: 14px; outline: none; width: 100%; }
        .input:focus { border-color: #6366f1; }
        .table-row { transition: background 0.15s; }
        .table-row:hover { background: rgba(99,102,241,0.05) !important; }
        .pulse { animation: pulse 2s infinite; }
        @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.5; } }
        .fade-in { animation: fadeIn 0.3s ease; }
        @keyframes fadeIn { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
        .modal-overlay { position:fixed; inset:0; background:rgba(0,0,0,0.7); z-index:1000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(4px); }
        .modal { background:#1a1d27; border:1px solid #2d3148; border-radius:16px; padding:28px; max-width:600px; width:calc(100% - 40px); max-height:85vh; overflow-y:auto; }
        .progress-bar { height: 6px; background: #2d3148; border-radius: 3px; overflow: hidden; }
        .progress-fill { height: 100%; border-radius: 3px; transition: width 0.8s ease; }
        select.input option { background: #1a1d27; }
        .tag { display:inline-flex; align-items:center; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
      `}</style>

      <Sidebar page={page} setPage={setPage} open={sidebarOpen} setOpen={setSidebarOpen} user={user} onLogout={() => { setUser(null); setPage("login"); }} />

      <div style={{ flex: 1, overflow: "auto", display: "flex", flexDirection: "column" }}>
        <TopBar page={page} user={user} sidebarOpen={sidebarOpen} setSidebarOpen={setSidebarOpen} onImport={() => setImportModal(true)} invoices={invoices} showToast={showToast} />

        <main style={{ flex: 1, padding: "24px", overflow: "auto" }} className="fade-in">
          {page === "admin-dashboard" && <AdminDashboard invoices={invoices} salespeople={SALESPEOPLE} setPage={setPage} />}
          {page === "sp-dashboard" && <SPDashboard invoices={invoices} user={user} salespeople={SALESPEOPLE} />}
          {page === "invoices" && <InvoicesPage invoices={invoices} salespeople={SALESPEOPLE} user={user} onView={setSelectedInvoice} onRemind={setReminderModal} showToast={showToast} setInvoices={setInvoices} />}
          {page === "parties" && <PartiesPage invoices={invoices} salespeople={SALESPEOPLE} user={user} onRemind={setReminderModal} />}
          {page === "reminders" && <RemindersPage templates={reminderTemplates} setTemplates={setReminderTemplates} history={reminderHistory} showToast={showToast} />}
          {page === "salespeople" && <SalespeopePage invoices={invoices} salespeople={SALESPEOPLE} />}
          {page === "analytics" && <AnalyticsPage invoices={invoices} salespeople={SALESPEOPLE} />}
          {page === "reports" && <ReportsPage invoices={invoices} salespeople={SALESPEOPLE} showToast={showToast} />}
          {page === "settings" && <SettingsPage showToast={showToast} />}
        </main>
      </div>

      {importModal && <ImportModal onClose={() => setImportModal(false)} onImport={handleImport} />}
      {selectedInvoice && <InvoiceDetailModal invoice={selectedInvoice} salespeople={SALESPEOPLE} onClose={() => setSelectedInvoice(null)} onRemind={setReminderModal} />}
      {reminderModal && <SendReminderModal invoice={reminderModal} onClose={() => setReminderModal(null)} onSend={sendReminder} />}
      {toast && <Toast msg={toast.msg} type={toast.type} />}
    </div>
  );
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState("admin@company.com");
  const [pass, setPass] = useState("admin123");
  const [role, setRole] = useState("admin");

  const handleSubmit = () => {
    if (role === "admin") onLogin("admin");
    else onLogin("salesperson", "sp1");
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f1117 0%, #1a1d27 50%, #0f1117 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Space+Grotesk:wght@700&display=swap');* {box-sizing:border-box;} .inp{background:#1a1d27;border:1px solid #2d3148;color:#e2e8f0;padding:12px 16px;border-radius:8px;font-size:14px;outline:none;width:100%;margin-top:6px;font-family:'DM Sans',sans-serif;}.inp:focus{border-color:#6366f1;} .role-btn{padding:10px 20px;border-radius:8px;border:1px solid #2d3148;background:#1a1d27;color:#94a3b8;cursor:pointer;font-size:13px;font-weight:500;transition:all .2s;} .role-btn.active{background:rgba(99,102,241,0.2);color:#a5b4fc;border-color:#6366f1;}`}</style>
      <div style={{ width: "100%", maxWidth: 420, padding: "0 20px" }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
            <div style={{ width: 48, height: 48, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>📊</div>
            <div>
              <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 24, fontWeight: 700, color: "#fff" }}>TallyCRM</div>
              <div style={{ fontSize: 12, color: "#64748b" }}>Payment Collections Platform</div>
            </div>
          </div>
          <p style={{ color: "#64748b", fontSize: 14 }}>Manage outstanding payments & automate reminders</p>
        </div>

        <div style={{ background: "#1a1d27", border: "1px solid #2d3148", borderRadius: 16, padding: 32 }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 24 }}>
            {["admin", "salesperson"].map(r => (
              <button key={r} className={`role-btn${role === r ? " active" : ""}`} style={{ flex: 1 }} onClick={() => setRole(r)}>
                {r === "admin" ? "🛡️ Admin" : "👤 Sales Rep"}
              </button>
            ))}
          </div>
          <div style={{ marginBottom: 16 }}>
            <label style={{ fontSize: 12, color: "#64748b", fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>Email</label>
            <input className="inp" value={email} onChange={e => setEmail(e.target.value)} />
          </div>
          <div style={{ marginBottom: 24 }}>
            <label style={{ fontSize: 12, color: "#64748b", fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>Password</label>
            <input className="inp" type="password" value={pass} onChange={e => setPass(e.target.value)} />
          </div>
          <button onClick={handleSubmit} style={{ width: "100%", background: "linear-gradient(135deg,#6366f1,#8b5cf6)", color: "#fff", border: "none", padding: "13px", borderRadius: 10, fontSize: 15, fontWeight: 700, cursor: "pointer", letterSpacing: 0.5 }}>
            Sign In →
          </button>
          <div style={{ marginTop: 16, padding: "12px", background: "rgba(99,102,241,0.08)", borderRadius: 8, fontSize: 12, color: "#64748b" }}>
            Demo: admin@company.com / admin123
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── SIDEBAR ──────────────────────────────────────────────────────────────────
function Sidebar({ page, setPage, open, setOpen, user, onLogout }) {
  const adminNav = [
    { id: "admin-dashboard", label: "Dashboard", icon: "⬡" },
    { id: "invoices", label: "Invoices", icon: "🧾" },
    { id: "parties", label: "Parties", icon: "🏢" },
    { id: "reminders", label: "Reminders", icon: "🔔" },
    { id: "salespeople", label: "Sales Team", icon: "👥" },
    { id: "analytics", label: "Analytics", icon: "📈" },
    { id: "reports", label: "Reports", icon: "📋" },
    { id: "settings", label: "Settings", icon: "⚙️" },
  ];
  const spNav = [
    { id: "sp-dashboard", label: "My Dashboard", icon: "⬡" },
    { id: "invoices", label: "My Invoices", icon: "🧾" },
    { id: "parties", label: "My Parties", icon: "🏢" },
    { id: "reminders", label: "Reminders", icon: "🔔" },
  ];
  const nav = user?.role === "admin" ? adminNav : spNav;

  return (
    <div style={{ width: open ? 240 : 64, minWidth: open ? 240 : 64, background: "#13151f", borderRight: "1px solid #1e2235", transition: "all 0.25s", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: open ? "20px 16px" : "20px 12px", borderBottom: "1px solid #1e2235", display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{ width: 36, height: 36, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, flexShrink: 0 }}>📊</div>
        {open && <div><div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 16, color: "#fff" }}>TallyCRM</div><div style={{ fontSize: 10, color: "#475569", marginTop: 1 }}>v2.0 Pro</div></div>}
      </div>

      <nav style={{ flex: 1, padding: "12px 8px", overflowY: "auto" }}>
        {nav.map(item => (
          <div key={item.id} className={`nav-item${page === item.id ? " active" : ""}`}
            onClick={() => setPage(item.id)}
            style={{ display: "flex", alignItems: "center", gap: 10, padding: open ? "10px 12px" : "10px", borderRadius: 8, marginBottom: 2, borderLeft: "3px solid transparent", cursor: "pointer" }}>
            <span style={{ fontSize: 16, flexShrink: 0 }}>{item.icon}</span>
            {open && <span style={{ fontSize: 13, fontWeight: 500, color: page === item.id ? "#a5b4fc" : "#94a3b8", whiteSpace: "nowrap" }}>{item.label}</span>}
          </div>
        ))}
      </nav>

      {open && (
        <div style={{ padding: "12px 16px", borderTop: "1px solid #1e2235" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <div style={{ width: 32, height: 32, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, color: "#fff", flexShrink: 0 }}>
              {user?.name?.split(" ").map(n => n[0]).join("").slice(0, 2)}
            </div>
            <div style={{ overflow: "hidden" }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: "#e2e8f0", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{user?.name}</div>
              <div style={{ fontSize: 11, color: "#475569", textTransform: "capitalize" }}>{user?.role}</div>
            </div>
          </div>
          <button onClick={onLogout} className="btn-secondary" style={{ width: "100%", fontSize: 12 }}>Sign Out</button>
        </div>
      )}
    </div>
  );
}

// ─── TOP BAR ──────────────────────────────────────────────────────────────────
function TopBar({ page, user, sidebarOpen, setSidebarOpen, onImport, invoices, showToast }) {
  const titles = { "admin-dashboard": "Admin Dashboard", "sp-dashboard": "My Dashboard", invoices: "Invoices", parties: "Parties & Clients", reminders: "Reminder Automation", salespeople: "Sales Team", analytics: "Analytics", reports: "Reports", settings: "Settings" };
  const overdue = invoices.filter(i => i.status === "overdue").length;

  return (
    <div style={{ height: 60, background: "#13151f", borderBottom: "1px solid #1e2235", display: "flex", alignItems: "center", padding: "0 20px", gap: 16, flexShrink: 0 }}>
      <button onClick={() => setSidebarOpen(!sidebarOpen)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 18, padding: 4 }}>☰</button>
      <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 600, fontSize: 16, color: "#e2e8f0" }}>{titles[page] || page}</div>
      <div style={{ flex: 1 }} />
      {overdue > 0 && (
        <div style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 8, padding: "5px 12px" }}>
          <span className="pulse" style={{ width: 8, height: 8, background: "#ef4444", borderRadius: "50%", display: "inline-block" }}></span>
          <span style={{ fontSize: 12, color: "#f87171", fontWeight: 600 }}>{overdue} Overdue</span>
        </div>
      )}
      {user?.role === "admin" && (
        <button onClick={onImport} className="btn-primary" style={{ fontSize: 13, padding: "8px 16px" }}>
          ⬆ Import Excel
        </button>
      )}
    </div>
  );
}

// ─── STAT CARD ────────────────────────────────────────────────────────────────
function StatCard({ icon, label, value, sub, color = "#6366f1", trend }) {
  return (
    <div className="card" style={{ padding: "20px 24px" }}>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
        <div>
          <div style={{ fontSize: 12, color: "#64748b", fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>{label}</div>
          <div style={{ fontSize: 24, fontWeight: 700, color: "#fff", fontFamily: "'Space Grotesk',sans-serif", marginBottom: 4 }}>{value}</div>
          {sub && <div style={{ fontSize: 12, color: "#64748b" }}>{sub}</div>}
          {trend !== undefined && (
            <div style={{ fontSize: 12, color: trend >= 0 ? "#10b981" : "#ef4444", marginTop: 4, fontWeight: 600 }}>
              {trend >= 0 ? "↑" : "↓"} {Math.abs(trend)}% vs last month
            </div>
          )}
        </div>
        <div style={{ width: 44, height: 44, background: `${color}20`, border: `1px solid ${color}40`, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>{icon}</div>
      </div>
    </div>
  );
}

// ─── ADMIN DASHBOARD ──────────────────────────────────────────────────────────
function AdminDashboard({ invoices, salespeople, setPage }) {
  const total = invoices.reduce((s, i) => s + i.amount, 0);
  const outstanding = invoices.reduce((s, i) => s + i.outstanding, 0);
  const collected = invoices.reduce((s, i) => s + i.paid, 0);
  const overdue = invoices.filter(i => i.status === "overdue");
  const overdueAmt = overdue.reduce((s, i) => s + i.outstanding, 0);
  const today = new Date().toISOString().split("T")[0];
  const dueToday = invoices.filter(i => i.dueDate === today && i.status !== "paid");
  const recovery = total > 0 ? ((collected / total) * 100).toFixed(1) : 0;

  const topParties = Object.entries(
    invoices.filter(i => i.status !== "paid").reduce((acc, inv) => {
      acc[inv.party] = (acc[inv.party] || 0) + inv.outstanding; return acc;
    }, {})
  ).sort((a, b) => b[1] - a[1]).slice(0, 5);

  const spPerf = salespeople.map(sp => {
    const spInv = invoices.filter(i => i.salespersonId === sp.id);
    const spTotal = spInv.reduce((s, i) => s + i.amount, 0);
    const spCollected = spInv.reduce((s, i) => s + i.paid, 0);
    const spOutstanding = spInv.reduce((s, i) => s + i.outstanding, 0);
    return { ...sp, total: spTotal, collected: spCollected, outstanding: spOutstanding, recovery: spTotal > 0 ? ((spCollected / spTotal) * 100).toFixed(0) : 0 };
  }).sort((a, b) => b.collected - a.collected);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 24 }}>
        <StatCard icon="💰" label="Total Outstanding" value={fmtShort(outstanding)} sub={`${invoices.filter(i => i.status !== "paid").length} invoices`} color="#6366f1" trend={-5.2} />
        <StatCard icon="✅" label="Collected" value={fmtShort(collected)} sub={`Recovery: ${recovery}%`} color="#10b981" trend={8.3} />
        <StatCard icon="⚠️" label="Overdue" value={fmtShort(overdueAmt)} sub={`${overdue.length} invoices`} color="#ef4444" />
        <StatCard icon="📅" label="Due Today" value={dueToday.length} sub={dueToday.length > 0 ? fmtShort(dueToday.reduce((s, i) => s + i.outstanding, 0)) : "All clear"} color="#f59e0b" />
        <StatCard icon="📦" label="Total Business" value={fmtShort(total)} sub={`${invoices.length} invoices`} color="#8b5cf6" />
        <StatCard icon="🔔" label="Reminders Sent" value={invoices.reduce((s, i) => s + i.remindersSent, 0)} sub="This month" color="#06b6d4" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
        {/* Top Pending Parties */}
        <div className="card" style={{ padding: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#e2e8f0" }}>🏆 Top Pending Parties</div>
            <button className="btn-secondary" style={{ fontSize: 11 }} onClick={() => setPage("parties")}>View All</button>
          </div>
          {topParties.map(([name, amt], i) => {
            const pct = (amt / outstanding * 100).toFixed(0);
            return (
              <div key={name} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 22, height: 22, background: `hsl(${i * 60},60%,30%)`, borderRadius: 4, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700, color: "#fff" }}>{i + 1}</div>
                    <span style={{ fontSize: 13, color: "#cbd5e1" }}>{name.length > 22 ? name.slice(0, 22) + "…" : name}</span>
                  </div>
                  <span style={{ fontSize: 13, fontWeight: 700, color: "#f87171" }}>{fmtShort(amt)}</span>
                </div>
                <div className="progress-bar">
                  <div className="progress-fill" style={{ width: `${pct}%`, background: `hsl(${i * 60 + 200},70%,50%)` }} />
                </div>
              </div>
            );
          })}
        </div>

        {/* Sales Team Performance */}
        <div className="card" style={{ padding: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#e2e8f0" }}>👥 Sales Team Performance</div>
            <button className="btn-secondary" style={{ fontSize: 11 }} onClick={() => setPage("salespeople")}>Details</button>
          </div>
          {spPerf.map((sp, i) => (
            <div key={sp.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0", borderBottom: i < spPerf.length - 1 ? "1px solid #1e2235" : "none" }}>
              <div style={{ width: 36, height: 36, background: `linear-gradient(135deg, hsl(${i * 80 + 200},70%,35%), hsl(${i * 80 + 240},70%,25%))`, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, color: "#fff", flexShrink: 0 }}>{sp.avatar}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: "#e2e8f0", marginBottom: 2 }}>{sp.name}</div>
                <div className="progress-bar" style={{ height: 4 }}>
                  <div className="progress-fill" style={{ width: `${sp.recovery}%`, background: sp.recovery >= 70 ? "#10b981" : sp.recovery >= 50 ? "#f59e0b" : "#ef4444" }} />
                </div>
              </div>
              <div style={{ textAlign: "right", flexShrink: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: sp.recovery >= 70 ? "#10b981" : sp.recovery >= 50 ? "#f59e0b" : "#ef4444" }}>{sp.recovery}%</div>
                <div style={{ fontSize: 11, color: "#64748b" }}>{fmtShort(sp.outstanding)} due</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Recent Overdue */}
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#e2e8f0" }}>🚨 Overdue Invoices</div>
          <button className="btn-secondary" style={{ fontSize: 11 }} onClick={() => setPage("invoices")}>View All</button>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ borderBottom: "1px solid #2d3148" }}>
                {["Invoice", "Party", "Amount", "Overdue Days", "Sales Person", "Reminders", ""].map(h => (
                  <th key={h} style={{ textAlign: "left", padding: "8px 12px", color: "#64748b", fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {overdue.slice(0, 6).map(inv => {
                const sp = salespeople.find(s => s.id === inv.salespersonId);
                return (
                  <tr key={inv.id} className="table-row" style={{ borderBottom: "1px solid #1e2235" }}>
                    <td style={{ padding: "10px 12px", color: "#a5b4fc", fontWeight: 600 }}>{inv.id}</td>
                    <td style={{ padding: "10px 12px", color: "#e2e8f0" }}>{inv.party.length > 20 ? inv.party.slice(0, 20) + "…" : inv.party}</td>
                    <td style={{ padding: "10px 12px", color: "#f87171", fontWeight: 700 }}>{fmt(inv.outstanding)}</td>
                    <td style={{ padding: "10px 12px" }}>
                      <span style={{ background: "rgba(239,68,68,0.15)", color: "#f87171", padding: "2px 8px", borderRadius: 4, fontWeight: 700, fontSize: 12 }}>{inv.overdueDays}d</span>
                    </td>
                    <td style={{ padding: "10px 12px", color: "#94a3b8" }}>{sp?.name.split(" ")[0]}</td>
                    <td style={{ padding: "10px 12px", color: "#64748b" }}>{inv.remindersSent} sent</td>
                    <td style={{ padding: "10px 12px" }}>
                      <div style={{ display: "flex", gap: 6 }}>
                        {["💬", "✉️", "📱"].map(ch => (
                          <button key={ch} style={{ background: "rgba(99,102,241,0.15)", border: "1px solid #3d4270", borderRadius: 6, padding: "3px 7px", cursor: "pointer", fontSize: 12 }}>{ch}</button>
                        ))}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── SP DASHBOARD ─────────────────────────────────────────────────────────────
function SPDashboard({ invoices, user, salespeople }) {
  const myInv = invoices.filter(i => i.salespersonId === user?.spId);
  const sp = salespeople.find(s => s.id === user?.spId);
  const total = myInv.reduce((s, i) => s + i.amount, 0);
  const outstanding = myInv.reduce((s, i) => s + i.outstanding, 0);
  const collected = myInv.reduce((s, i) => s + i.paid, 0);
  const recovery = total > 0 ? ((collected / total) * 100).toFixed(1) : 0;
  const overdue = myInv.filter(i => i.status === "overdue");

  const partyGroups = myInv.reduce((acc, inv) => {
    if (!acc[inv.party]) acc[inv.party] = { name: inv.party, phone: inv.phone, total: 0, paid: 0, outstanding: 0, invoices: [] };
    acc[inv.party].total += inv.amount;
    acc[inv.party].paid += inv.paid;
    acc[inv.party].outstanding += inv.outstanding;
    acc[inv.party].invoices.push(inv);
    return acc;
  }, {});

  return (
    <div>
      <div style={{ marginBottom: 20, padding: "16px 20px", background: "linear-gradient(135deg,rgba(99,102,241,0.15),rgba(139,92,246,0.1))", border: "1px solid rgba(99,102,241,0.3)", borderRadius: 12, display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ width: 52, height: 52, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, fontWeight: 700, color: "#fff" }}>{sp?.avatar}</div>
        <div>
          <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontSize: 20, fontWeight: 700, color: "#fff" }}>{sp?.name}</div>
          <div style={{ fontSize: 13, color: "#94a3b8" }}>Sales Representative • Target: {fmtShort(sp?.target || 0)}</div>
        </div>
        <div style={{ flex: 1 }} />
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 12, color: "#64748b" }}>Target Achievement</div>
          <div style={{ fontSize: 22, fontWeight: 700, color: collected >= sp?.target ? "#10b981" : "#f59e0b" }}>{((collected / (sp?.target || 1)) * 100).toFixed(0)}%</div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 24 }}>
        <StatCard icon="💰" label="Total Sales" value={fmtShort(total)} sub={`${myInv.length} invoices`} color="#6366f1" />
        <StatCard icon="⏳" label="Outstanding" value={fmtShort(outstanding)} sub={`${myInv.filter(i => i.status !== "paid").length} pending`} color="#f59e0b" />
        <StatCard icon="✅" label="Collected" value={fmtShort(collected)} color="#10b981" />
        <StatCard icon="📊" label="Recovery %" value={`${recovery}%`} sub={recovery >= 70 ? "On track 🎯" : "Needs attention"} color={recovery >= 70 ? "#10b981" : "#ef4444"} />
        <StatCard icon="🚨" label="Overdue" value={overdue.length} sub={fmtShort(overdue.reduce((s, i) => s + i.outstanding, 0))} color="#ef4444" />
      </div>

      <div style={{ fontWeight: 700, fontSize: 15, color: "#e2e8f0", marginBottom: 12 }}>Party-wise Pending Payments</div>
      <div style={{ display: "grid", gap: 12 }}>
        {Object.values(partyGroups).filter(p => p.outstanding > 0).sort((a, b) => b.outstanding - a.outstanding).map(party => (
          <div key={party.name} className="card" style={{ padding: "16px 20px" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
              <div>
                <div style={{ fontWeight: 600, color: "#e2e8f0", marginBottom: 4 }}>{party.name}</div>
                <div style={{ fontSize: 12, color: "#64748b" }}>{party.phone} • {party.invoices.length} invoices</div>
              </div>
              <div style={{ display: "flex", gap: 20, alignItems: "center" }}>
                <div style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 11, color: "#64748b" }}>Total</div>
                  <div style={{ fontWeight: 700, color: "#e2e8f0" }}>{fmtShort(party.total)}</div>
                </div>
                <div style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 11, color: "#64748b" }}>Collected</div>
                  <div style={{ fontWeight: 700, color: "#10b981" }}>{fmtShort(party.paid)}</div>
                </div>
                <div style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 11, color: "#64748b" }}>Outstanding</div>
                  <div style={{ fontWeight: 700, color: "#f87171" }}>{fmtShort(party.outstanding)}</div>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  {["💬", "✉️", "📱"].map(ch => (
                    <button key={ch} style={{ background: "rgba(99,102,241,0.1)", border: "1px solid #3d4270", borderRadius: 6, padding: "5px 8px", cursor: "pointer", fontSize: 14 }}>{ch}</button>
                  ))}
                </div>
              </div>
            </div>
            <div style={{ marginTop: 10 }}>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${party.total > 0 ? (party.paid / party.total * 100) : 0}%`, background: "linear-gradient(90deg,#10b981,#34d399)" }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4, fontSize: 11, color: "#64748b" }}>
                <span>{party.total > 0 ? (party.paid / party.total * 100).toFixed(0) : 0}% collected</span>
                <span>{party.total > 0 ? (party.outstanding / party.total * 100).toFixed(0) : 0}% pending</span>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── INVOICES PAGE ────────────────────────────────────────────────────────────
function InvoicesPage({ invoices, salespeople, user, onView, onRemind, showToast, setInvoices }) {
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [spFilter, setSpFilter] = useState("all");
  const [sort, setSort] = useState({ field: "dueDate", dir: "asc" });

  const myInvoices = user?.role === "salesperson" ? invoices.filter(i => i.salespersonId === user.spId) : invoices;

  const filtered = myInvoices.filter(inv => {
    const matchSearch = !search || inv.party.toLowerCase().includes(search.toLowerCase()) || inv.id.toLowerCase().includes(search.toLowerCase());
    const matchStatus = statusFilter === "all" || inv.status === statusFilter;
    const matchSp = spFilter === "all" || inv.salespersonId === spFilter;
    return matchSearch && matchStatus && matchSp;
  }).sort((a, b) => {
    const dir = sort.dir === "asc" ? 1 : -1;
    if (sort.field === "dueDate") return (a.dueDate > b.dueDate ? 1 : -1) * dir;
    if (sort.field === "amount") return (a.outstanding - b.outstanding) * dir;
    return 0;
  });

  const markPaid = (id) => {
    setInvoices(prev => prev.map(i => i.id === id ? { ...i, paid: i.amount, outstanding: 0, status: "paid" } : i));
    showToast("Invoice marked as paid");
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
        <input className="input" placeholder="🔍 Search by party or invoice…" value={search} onChange={e => setSearch(e.target.value)} style={{ flex: 1, minWidth: 200 }} />
        <select className="input" value={statusFilter} onChange={e => setStatusFilter(e.target.value)} style={{ width: 140 }}>
          <option value="all">All Status</option>
          <option value="overdue">Overdue</option>
          <option value="pending">Pending</option>
          <option value="partial">Partial</option>
          <option value="paid">Paid</option>
        </select>
        {user?.role === "admin" && (
          <select className="input" value={spFilter} onChange={e => setSpFilter(e.target.value)} style={{ width: 160 }}>
            <option value="all">All Salespeople</option>
            {salespeople.map(sp => <option key={sp.id} value={sp.id}>{sp.name}</option>)}
          </select>
        )}
        <select className="input" value={`${sort.field}-${sort.dir}`} onChange={e => { const [f, d] = e.target.value.split("-"); setSort({ field: f, dir: d }); }} style={{ width: 160 }}>
          <option value="dueDate-asc">Due Date ↑</option>
          <option value="dueDate-desc">Due Date ↓</option>
          <option value="amount-desc">Amount ↓</option>
          <option value="amount-asc">Amount ↑</option>
        </select>
      </div>

      <div style={{ marginBottom: 12, fontSize: 13, color: "#64748b" }}>
        Showing <strong style={{ color: "#a5b4fc" }}>{filtered.length}</strong> of {myInvoices.length} invoices
        {" • "}Outstanding: <strong style={{ color: "#f87171" }}>{fmtShort(filtered.reduce((s, i) => s + i.outstanding, 0))}</strong>
      </div>

      <div className="card" style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#13151f", borderBottom: "1px solid #2d3148" }}>
                {["Invoice ID", "Party", "Invoice Date", "Due Date", "Amount", "Paid", "Outstanding", "Status", "Sales Person", "Actions"].map(h => (
                  <th key={h} style={{ padding: "12px 14px", textAlign: "left", color: "#64748b", fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(inv => {
                const sp = salespeople.find(s => s.id === inv.salespersonId);
                const sc = statusConfig[inv.status];
                return (
                  <tr key={inv.id} className="table-row" style={{ borderBottom: "1px solid #1e2235" }}>
                    <td style={{ padding: "11px 14px" }}>
                      <button onClick={() => onView(inv)} style={{ background: "none", border: "none", color: "#a5b4fc", cursor: "pointer", fontWeight: 600, fontSize: 13 }}>{inv.id}</button>
                    </td>
                    <td style={{ padding: "11px 14px" }}>
                      <div style={{ color: "#e2e8f0", fontWeight: 500 }}>{inv.party.length > 20 ? inv.party.slice(0, 20) + "…" : inv.party}</div>
                      <div style={{ fontSize: 11, color: "#64748b" }}>{inv.gst}</div>
                    </td>
                    <td style={{ padding: "11px 14px", color: "#94a3b8" }}>{inv.invoiceDate}</td>
                    <td style={{ padding: "11px 14px" }}>
                      <span style={{ color: inv.status === "overdue" ? "#f87171" : "#94a3b8", fontWeight: inv.status === "overdue" ? 700 : 400 }}>
                        {inv.dueDate}
                        {inv.overdueDays > 0 && <span style={{ display: "block", fontSize: 11, color: "#ef4444" }}>+{inv.overdueDays}d overdue</span>}
                      </span>
                    </td>
                    <td style={{ padding: "11px 14px", color: "#e2e8f0", fontWeight: 600 }}>{fmt(inv.amount)}</td>
                    <td style={{ padding: "11px 14px", color: "#10b981" }}>{fmt(inv.paid)}</td>
                    <td style={{ padding: "11px 14px", color: "#f87171", fontWeight: 700 }}>{fmt(inv.outstanding)}</td>
                    <td style={{ padding: "11px 14px" }}>
                      <span className="badge" style={{ background: sc.bg, color: sc.color, fontSize: 11 }}>
                        {sc.icon} {sc.label}
                      </span>
                    </td>
                    <td style={{ padding: "11px 14px", color: "#94a3b8" }}>{sp?.name.split(" ")[0]}</td>
                    <td style={{ padding: "11px 14px" }}>
                      <div style={{ display: "flex", gap: 5 }}>
                        {inv.status !== "paid" && (
                          <>
                            <button onClick={() => onRemind(inv)} style={{ background: "rgba(99,102,241,0.15)", border: "1px solid #3d4270", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 12, color: "#a5b4fc" }}>🔔</button>
                            <button onClick={() => markPaid(inv.id)} style={{ background: "rgba(16,185,129,0.1)", border: "1px solid rgba(16,185,129,0.3)", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 12, color: "#10b981" }}>✓</button>
                          </>
                        )}
                        <button onClick={() => onView(inv)} style={{ background: "rgba(100,116,139,0.1)", border: "1px solid #2d3148", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 12, color: "#94a3b8" }}>👁</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {filtered.length === 0 && (
          <div style={{ textAlign: "center", padding: "60px 20px", color: "#64748b" }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>📭</div>
            <div style={{ fontSize: 16 }}>No invoices found</div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── PARTIES PAGE ─────────────────────────────────────────────────────────────
function PartiesPage({ invoices, salespeople, user, onRemind }) {
  const [search, setSearch] = useState("");
  const myInvoices = user?.role === "salesperson" ? invoices.filter(i => i.salespersonId === user.spId) : invoices;

  const parties = Object.entries(
    myInvoices.reduce((acc, inv) => {
      if (!acc[inv.party]) acc[inv.party] = { name: inv.party, phone: inv.phone, email: inv.email, gst: inv.gst, spId: inv.salespersonId, invoices: [], total: 0, paid: 0, outstanding: 0, overdue: 0 };
      acc[inv.party].invoices.push(inv);
      acc[inv.party].total += inv.amount;
      acc[inv.party].paid += inv.paid;
      acc[inv.party].outstanding += inv.outstanding;
      if (inv.status === "overdue") acc[inv.party].overdue += inv.outstanding;
      return acc;
    }, {})
  ).map(([, v]) => v).filter(p => !search || p.name.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => b.outstanding - a.outstanding);

  return (
    <div>
      <input className="input" placeholder="🔍 Search parties…" value={search} onChange={e => setSearch(e.target.value)} style={{ maxWidth: 400, marginBottom: 20 }} />
      <div style={{ display: "grid", gap: 14 }}>
        {parties.map(party => {
          const sp = salespeople.find(s => s.id === party.spId);
          const recovery = party.total > 0 ? (party.paid / party.total * 100).toFixed(0) : 0;
          const pendingInvoice = party.invoices.find(i => i.status !== "paid");
          return (
            <div key={party.name} className="card" style={{ padding: 20 }}>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 16, alignItems: "flex-start", justifyContent: "space-between" }}>
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                    <div style={{ width: 40, height: 40, background: "linear-gradient(135deg,#1e2235,#2d3148)", border: "1px solid #3d4270", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>🏢</div>
                    <div>
                      <div style={{ fontWeight: 700, color: "#e2e8f0", fontSize: 15 }}>{party.name}</div>
                      <div style={{ fontSize: 12, color: "#64748b" }}>GST: {party.gst}</div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap", fontSize: 12, color: "#94a3b8" }}>
                    <span>📱 {party.phone}</span>
                    <span>✉️ {party.email}</span>
                    {sp && <span>👤 {sp.name}</span>}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
                  {[["Total", party.total, "#e2e8f0"], ["Collected", party.paid, "#10b981"], ["Outstanding", party.outstanding, "#f87171"]].map(([l, v, c]) => (
                    <div key={l} style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 11, color: "#64748b", marginBottom: 2 }}>{l}</div>
                      <div style={{ fontSize: 16, fontWeight: 700, color: c }}>{fmtShort(v)}</div>
                    </div>
                  ))}
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 11, color: "#64748b", marginBottom: 2 }}>Recovery</div>
                    <div style={{ fontSize: 16, fontWeight: 700, color: recovery >= 70 ? "#10b981" : recovery >= 40 ? "#f59e0b" : "#ef4444" }}>{recovery}%</div>
                  </div>
                </div>
              </div>
              <div style={{ marginTop: 14 }}>
                <div className="progress-bar">
                  <div className="progress-fill" style={{ width: `${recovery}%`, background: "linear-gradient(90deg,#6366f1,#8b5cf6)" }} />
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12 }}>
                <div style={{ display: "flex", gap: 6 }}>
                  {party.invoices.map(inv => (
                    <span key={inv.id} className="badge" style={{ background: statusConfig[inv.status].bg, color: statusConfig[inv.status].color, fontSize: 10 }}>
                      {inv.id}
                    </span>
                  ))}
                </div>
                {pendingInvoice && (
                  <div style={{ display: "flex", gap: 6 }}>
                    {["💬 WhatsApp", "✉️ Email", "📱 SMS"].map((ch, i) => (
                      <button key={i} onClick={() => onRemind(pendingInvoice)} className="btn-secondary" style={{ fontSize: 11, padding: "5px 10px" }}>{ch}</button>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── REMINDERS PAGE ───────────────────────────────────────────────────────────
function RemindersPage({ templates, setTemplates, history, showToast }) {
  const [tab, setTab] = useState("automation");

  const toggleTemplate = (id) => {
    setTemplates(prev => prev.map(t => t.id === id ? { ...t, active: !t.active } : t));
    showToast("Reminder rule updated");
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 24 }}>
        {["automation", "history", "templates"].map(t => (
          <button key={t} onClick={() => setTab(t)} className={tab === t ? "btn-primary" : "btn-secondary"} style={{ textTransform: "capitalize", fontSize: 13 }}>{t}</button>
        ))}
      </div>

      {tab === "automation" && (
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div>
              <div style={{ fontWeight: 700, color: "#e2e8f0", fontSize: 15 }}>Reminder Automation Rules</div>
              <div style={{ fontSize: 13, color: "#64748b", marginTop: 4 }}>Configure when and how reminders are sent automatically</div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span className="pulse" style={{ width: 8, height: 8, background: "#10b981", borderRadius: "50%", display: "inline-block" }}></span>
              <span style={{ fontSize: 12, color: "#10b981", fontWeight: 600 }}>{templates.filter(t => t.active).length} rules active</span>
            </div>
          </div>
          <div style={{ display: "grid", gap: 12 }}>
            {templates.map(t => (
              <div key={t.id} className="card" style={{ padding: 20, borderLeft: `3px solid ${t.active ? "#6366f1" : "#2d3148"}` }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                    <div style={{ width: 42, height: 42, background: t.active ? "rgba(99,102,241,0.15)" : "#1e2235", border: `1px solid ${t.active ? "#6366f1" : "#2d3148"}`, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>
                      {t.trigger === "before_due" ? "⏰" : "🚨"}
                    </div>
                    <div>
                      <div style={{ fontWeight: 700, color: "#e2e8f0", marginBottom: 4 }}>{t.name}</div>
                      <div style={{ fontSize: 12, color: "#64748b" }}>
                        {t.trigger === "before_due" ? `📅 ${t.days} day(s) before due date` : `⚠️ ${t.days} day(s) after due date`}
                      </div>
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <div style={{ display: "flex", gap: 6 }}>
                      {t.channel.map(ch => (
                        <span key={ch} style={{ background: "rgba(99,102,241,0.1)", border: "1px solid #3d4270", borderRadius: 6, padding: "4px 10px", fontSize: 12, color: "#a5b4fc" }}>
                          {channelIcons[ch]} {ch}
                        </span>
                      ))}
                    </div>
                    <div onClick={() => toggleTemplate(t.id)} style={{ width: 44, height: 24, background: t.active ? "#6366f1" : "#2d3148", borderRadius: 12, cursor: "pointer", position: "relative", transition: "background 0.2s" }}>
                      <div style={{ position: "absolute", top: 3, left: t.active ? 22 : 3, width: 18, height: 18, background: "#fff", borderRadius: "50%", transition: "left 0.2s" }} />
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div style={{ marginTop: 20, padding: 20, background: "rgba(99,102,241,0.05)", border: "1px dashed #3d4270", borderRadius: 12, textAlign: "center" }}>
            <div style={{ fontSize: 24, marginBottom: 8 }}>➕</div>
            <div style={{ color: "#94a3b8", marginBottom: 12 }}>Add a new automation rule</div>
            <button className="btn-primary" onClick={() => showToast("Rule builder coming soon!")} style={{ fontSize: 13 }}>Create Rule</button>
          </div>
        </div>
      )}

      {tab === "history" && (
        <div className="card" style={{ overflow: "hidden" }}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#13151f", borderBottom: "1px solid #2d3148" }}>
                  {["Invoice", "Party", "Channel", "Template", "Sent At", "Status"].map(h => (
                    <th key={h} style={{ padding: "12px 14px", textAlign: "left", color: "#64748b", fontWeight: 600, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {history.map(rh => (
                  <tr key={rh.id} className="table-row" style={{ borderBottom: "1px solid #1e2235" }}>
                    <td style={{ padding: "11px 14px", color: "#a5b4fc", fontWeight: 600 }}>{rh.invoiceId}</td>
                    <td style={{ padding: "11px 14px", color: "#e2e8f0" }}>{rh.party}</td>
                    <td style={{ padding: "11px 14px" }}>
                      <span style={{ fontSize: 16 }}>{channelIcons[rh.channel]}</span>
                      <span style={{ marginLeft: 6, color: "#94a3b8", textTransform: "capitalize" }}>{rh.channel}</span>
                    </td>
                    <td style={{ padding: "11px 14px", color: "#94a3b8" }}>{rh.template}</td>
                    <td style={{ padding: "11px 14px", color: "#64748b" }}>{rh.sentAt}</td>
                    <td style={{ padding: "11px 14px" }}>
                      <span className="badge" style={{ background: `${reminderStatusColors[rh.status]}20`, color: reminderStatusColors[rh.status], fontSize: 11 }}>
                        {rh.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "templates" && (
        <div>
          <div style={{ display: "grid", gap: 16 }}>
            {[
              { name: "Payment Due Reminder", channel: "WhatsApp", preview: "Dear {party_name}, your invoice #{invoice_id} of ₹{amount} is due on {due_date}. Kindly arrange payment. Contact: {salesperson_name} - {salesperson_phone}" },
              { name: "Overdue Notice", channel: "Email", preview: "Subject: Overdue Payment Notice\n\nDear {party_name},\n\nThis is to remind you that invoice #{invoice_id} amounting to ₹{outstanding} is overdue by {overdue_days} days..." },
              { name: "Friendly Reminder", channel: "SMS", preview: "Hi! Your payment of ₹{amount} for invoice #{invoice_id} is pending. Due: {due_date}. Please pay soon. -{company_name}" },
            ].map(tpl => (
              <div key={tpl.name} className="card" style={{ padding: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <div style={{ fontWeight: 700, color: "#e2e8f0" }}>{tpl.name}</div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <span style={{ background: "rgba(99,102,241,0.1)", border: "1px solid #3d4270", borderRadius: 6, padding: "3px 10px", fontSize: 12, color: "#a5b4fc" }}>{tpl.channel}</span>
                    <button className="btn-secondary" style={{ fontSize: 11 }}>Edit</button>
                  </div>
                </div>
                <div style={{ background: "#13151f", borderRadius: 8, padding: 14, fontSize: 12, color: "#94a3b8", fontFamily: "monospace", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{tpl.preview}</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── SALESPEOPLE PAGE ─────────────────────────────────────────────────────────
function SalespeopePage({ invoices, salespeople }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 20 }}>
      {salespeople.map((sp, idx) => {
        const spInv = invoices.filter(i => i.salespersonId === sp.id);
        const total = spInv.reduce((s, i) => s + i.amount, 0);
        const collected = spInv.reduce((s, i) => s + i.paid, 0);
        const outstanding = spInv.reduce((s, i) => s + i.outstanding, 0);
        const overdue = spInv.filter(i => i.status === "overdue").length;
        const recovery = total > 0 ? ((collected / total) * 100).toFixed(1) : 0;
        const targetPct = sp.target > 0 ? ((collected / sp.target) * 100).toFixed(0) : 0;

        return (
          <div key={sp.id} className="card" style={{ padding: 24 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 20 }}>
              <div style={{ width: 52, height: 52, background: `linear-gradient(135deg, hsl(${idx * 80 + 200},65%,35%), hsl(${idx * 80 + 240},65%,25%))`, borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 700, color: "#fff" }}>{sp.avatar}</div>
              <div>
                <div style={{ fontWeight: 700, fontSize: 16, color: "#fff" }}>{sp.name}</div>
                <div style={{ fontSize: 12, color: "#64748b" }}>{sp.email}</div>
                <div style={{ fontSize: 12, color: "#64748b" }}>{sp.phone}</div>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
              {[["Total Sales", fmtShort(total), "#e2e8f0"], ["Collected", fmtShort(collected), "#10b981"], ["Outstanding", fmtShort(outstanding), "#f87171"], ["Overdue Inv.", overdue, "#f59e0b"]].map(([l, v, c]) => (
                <div key={l} style={{ background: "#13151f", borderRadius: 8, padding: "12px 14px" }}>
                  <div style={{ fontSize: 11, color: "#64748b", marginBottom: 4 }}>{l}</div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: c }}>{v}</div>
                </div>
              ))}
            </div>

            <div style={{ marginBottom: 8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 6 }}>
                <span style={{ color: "#64748b" }}>Recovery Rate</span>
                <span style={{ color: recovery >= 70 ? "#10b981" : "#f59e0b", fontWeight: 700 }}>{recovery}%</span>
              </div>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${recovery}%`, background: `linear-gradient(90deg, hsl(${recovery * 1.2},70%,45%), hsl(${recovery * 1.4},70%,55%))` }} />
              </div>
            </div>

            <div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 6 }}>
                <span style={{ color: "#64748b" }}>Target ({fmtShort(sp.target)})</span>
                <span style={{ color: targetPct >= 100 ? "#10b981" : "#a5b4fc", fontWeight: 700 }}>{Math.min(targetPct, 100)}%</span>
              </div>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${Math.min(targetPct, 100)}%`, background: "linear-gradient(90deg,#6366f1,#8b5cf6)" }} />
              </div>
            </div>

            <div style={{ marginTop: 16, fontSize: 12, color: "#64748b" }}>
              {spInv.length} total invoices • {spInv.filter(i => i.status === "paid").length} paid
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ─── ANALYTICS PAGE ───────────────────────────────────────────────────────────
function AnalyticsPage({ invoices, salespeople }) {
  const months = ["Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"];
  const monthData = months.map((m, i) => ({
    month: m,
    invoiced: (800000 + Math.random() * 600000),
    collected: (500000 + Math.random() * 500000),
  }));
  const maxVal = Math.max(...monthData.map(d => Math.max(d.invoiced, d.collected)));

  const statusBreakdown = ["paid", "pending", "partial", "overdue"].map(s => ({
    status: s, count: invoices.filter(i => i.status === s).length,
    amount: invoices.filter(i => i.status === s).reduce((a, b) => a + b.outstanding, 0),
  }));

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 20, marginBottom: 20 }}>
        {/* Monthly Chart */}
        <div className="card" style={{ padding: 24 }}>
          <div style={{ fontWeight: 700, color: "#e2e8f0", fontSize: 15, marginBottom: 20 }}>Monthly Collections vs Invoiced</div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 16, height: 160, paddingBottom: 8 }}>
            {monthData.map(d => (
              <div key={d.month} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                <div style={{ width: "100%", display: "flex", gap: 3, alignItems: "flex-end", height: 130 }}>
                  <div style={{ flex: 1, background: "rgba(99,102,241,0.6)", borderRadius: "4px 4px 0 0", height: `${(d.invoiced / maxVal) * 100}%` }} title={fmtShort(d.invoiced)} />
                  <div style={{ flex: 1, background: "rgba(16,185,129,0.6)", borderRadius: "4px 4px 0 0", height: `${(d.collected / maxVal) * 100}%` }} title={fmtShort(d.collected)} />
                </div>
                <div style={{ fontSize: 11, color: "#64748b" }}>{d.month}</div>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 16, marginTop: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#94a3b8" }}><div style={{ width: 12, height: 12, background: "rgba(99,102,241,0.6)", borderRadius: 2 }} /> Invoiced</div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#94a3b8" }}><div style={{ width: 12, height: 12, background: "rgba(16,185,129,0.6)", borderRadius: 2 }} /> Collected</div>
          </div>
        </div>

        {/* Status Breakdown */}
        <div className="card" style={{ padding: 24 }}>
          <div style={{ fontWeight: 700, color: "#e2e8f0", fontSize: 15, marginBottom: 20 }}>Invoice Status Breakdown</div>
          {statusBreakdown.map(s => (
            <div key={s.status} style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                <span className="badge" style={{ background: statusConfig[s.status].bg, color: statusConfig[s.status].color, fontSize: 11 }}>{statusConfig[s.status].label}</span>
                <span style={{ fontSize: 13, color: "#e2e8f0", fontWeight: 600 }}>{s.count} invoices</span>
              </div>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${(s.count / invoices.length * 100)}%`, background: statusConfig[s.status].color }} />
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* SP Comparison */}
      <div className="card" style={{ padding: 24 }}>
        <div style={{ fontWeight: 700, color: "#e2e8f0", fontSize: 15, marginBottom: 20 }}>Sales Team Performance Comparison</div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ borderBottom: "1px solid #2d3148" }}>
                {["Sales Person", "Total Invoiced", "Collected", "Outstanding", "Overdue", "Recovery%", "Target%"].map(h => (
                  <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: "#64748b", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {salespeople.map((sp, i) => {
                const spInv = invoices.filter(inv => inv.salespersonId === sp.id);
                const total = spInv.reduce((s, inv) => s + inv.amount, 0);
                const collected = spInv.reduce((s, inv) => s + inv.paid, 0);
                const outstanding = spInv.reduce((s, inv) => s + inv.outstanding, 0);
                const overdue = spInv.filter(inv => inv.status === "overdue").length;
                const recovery = total > 0 ? ((collected / total) * 100).toFixed(1) : 0;
                const targetPct = sp.target > 0 ? ((collected / sp.target) * 100).toFixed(0) : 0;
                return (
                  <tr key={sp.id} className="table-row" style={{ borderBottom: "1px solid #1e2235" }}>
                    <td style={{ padding: "12px 14px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ width: 32, height: 32, background: `linear-gradient(135deg, hsl(${i * 80 + 200},65%,35%), hsl(${i * 80 + 240},65%,25%))`, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, color: "#fff" }}>{sp.avatar}</div>
                        <span style={{ fontWeight: 600, color: "#e2e8f0" }}>{sp.name}</span>
                      </div>
                    </td>
                    <td style={{ padding: "12px 14px", color: "#e2e8f0", fontWeight: 600 }}>{fmtShort(total)}</td>
                    <td style={{ padding: "12px 14px", color: "#10b981", fontWeight: 600 }}>{fmtShort(collected)}</td>
                    <td style={{ padding: "12px 14px", color: "#f87171", fontWeight: 600 }}>{fmtShort(outstanding)}</td>
                    <td style={{ padding: "12px 14px" }}><span style={{ background: "rgba(239,68,68,0.1)", color: "#f87171", padding: "2px 8px", borderRadius: 4, fontSize: 12, fontWeight: 700 }}>{overdue}</span></td>
                    <td style={{ padding: "12px 14px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div className="progress-bar" style={{ width: 60 }}>
                          <div className="progress-fill" style={{ width: `${recovery}%`, background: recovery >= 70 ? "#10b981" : recovery >= 50 ? "#f59e0b" : "#ef4444" }} />
                        </div>
                        <span style={{ color: recovery >= 70 ? "#10b981" : recovery >= 50 ? "#f59e0b" : "#ef4444", fontWeight: 700 }}>{recovery}%</span>
                      </div>
                    </td>
                    <td style={{ padding: "12px 14px", color: targetPct >= 100 ? "#10b981" : "#a5b4fc", fontWeight: 700 }}>{targetPct}%</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── REPORTS PAGE ─────────────────────────────────────────────────────────────
function ReportsPage({ invoices, salespeople, showToast }) {
  const downloadExcel = () => {
    const data = invoices.map(inv => {
      const sp = salespeople.find(s => s.id === inv.salespersonId);
      return { "Invoice ID": inv.id, "Party Name": inv.party, "GST": inv.gst, "Invoice Date": inv.invoiceDate, "Due Date": inv.dueDate, "Amount": inv.amount, "Paid": inv.paid, "Outstanding": inv.outstanding, "Status": inv.status, "Overdue Days": inv.overdueDays, "Sales Person": sp?.name || "", "Reminders Sent": inv.remindersSent };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Outstanding Report");
    XLSX.writeFile(wb, `TallyCRM_Report_${new Date().toISOString().split("T")[0]}.xlsx`);
    showToast("Excel report downloaded!");
  };

  const downloadSPReport = () => {
    const data = salespeople.map(sp => {
      const spInv = invoices.filter(i => i.salespersonId === sp.id);
      const total = spInv.reduce((s, i) => s + i.amount, 0);
      const collected = spInv.reduce((s, i) => s + i.paid, 0);
      return { "Sales Person": sp.name, "Email": sp.email, "Total Invoiced": total, "Collected": collected, "Outstanding": total - collected, "Overdue Invoices": spInv.filter(i => i.status === "overdue").length, "Recovery %": total > 0 ? ((collected / total) * 100).toFixed(1) + "%" : "0%", "Target": sp.target, "Target Achievement": sp.target > 0 ? ((collected / sp.target) * 100).toFixed(0) + "%" : "0%" };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SP Performance");
    XLSX.writeFile(wb, `TallyCRM_SP_Report_${new Date().toISOString().split("T")[0]}.xlsx`);
    showToast("Sales team report downloaded!");
  };

  const reports = [
    { icon: "📊", title: "Outstanding Invoice Report", desc: "All pending/overdue invoices with party details, due dates, amounts, and salesperson info.", action: downloadExcel, color: "#6366f1" },
    { icon: "👥", title: "Sales Team Performance", desc: "Per salesperson: total sales, collections, recovery %, target achievement.", action: downloadSPReport, color: "#10b981" },
    { icon: "🏢", title: "Party-wise Summary", desc: "Consolidated outstanding per party with payment history and contact details.", action: () => showToast("Generating party report…"), color: "#f59e0b" },
    { icon: "🔔", title: "Reminder Activity Log", desc: "All reminder messages sent — channel, status, template, timestamp.", action: () => showToast("Generating reminder log…"), color: "#8b5cf6" },
    { icon: "📈", title: "Monthly Collection Trend", desc: "Month-by-month invoiced vs collected comparison for the last 12 months.", action: () => showToast("Generating trend report…"), color: "#06b6d4" },
    { icon: "⚠️", title: "Aging Analysis Report", desc: "Invoices bucketed by overdue period: 0-30, 31-60, 61-90, 90+ days.", action: () => showToast("Generating aging report…"), color: "#ef4444" },
  ];

  return (
    <div>
      <div style={{ marginBottom: 20, padding: "16px 20px", background: "rgba(16,185,129,0.05)", border: "1px solid rgba(16,185,129,0.2)", borderRadius: 12 }}>
        <div style={{ fontSize: 14, color: "#34d399", fontWeight: 600, marginBottom: 4 }}>📋 Reports Center</div>
        <div style={{ fontSize: 13, color: "#64748b" }}>Generate and download Excel reports for offline analysis, sharing, or audit purposes.</div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16 }}>
        {reports.map(r => (
          <div key={r.title} className="card" style={{ padding: 24 }}>
            <div style={{ width: 48, height: 48, background: `${r.color}15`, border: `1px solid ${r.color}30`, borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, marginBottom: 14 }}>{r.icon}</div>
            <div style={{ fontWeight: 700, color: "#e2e8f0", marginBottom: 8 }}>{r.title}</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 16, lineHeight: 1.5 }}>{r.desc}</div>
            <button onClick={r.action} className="btn-primary" style={{ width: "100%", background: `linear-gradient(135deg, ${r.color}cc, ${r.color})` }}>
              ⬇ Download Excel
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── SETTINGS PAGE ────────────────────────────────────────────────────────────
function SettingsPage({ showToast }) {
  const [wa, setWa] = useState({ apiKey: "waba_xxxxx", phoneId: "1234567890", businessId: "" });
  const [email, setEmail] = useState({ host: "smtp.gmail.com", port: "587", user: "noreply@company.com", pass: "" });
  const [sms, setSms] = useState({ provider: "Twilio", sid: "", token: "", from: "" });

  return (
    <div style={{ maxWidth: 800 }}>
      <div style={{ display: "grid", gap: 20 }}>
        {/* WhatsApp */}
        <div className="card" style={{ padding: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
            <div style={{ width: 40, height: 40, background: "rgba(37,211,102,0.15)", border: "1px solid rgba(37,211,102,0.3)", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>💬</div>
            <div><div style={{ fontWeight: 700, color: "#e2e8f0" }}>WhatsApp Business API</div><div style={{ fontSize: 12, color: "#64748b" }}>Meta WhatsApp Business Platform</div></div>
            <div style={{ flex: 1 }} />
            <span style={{ background: "rgba(37,211,102,0.1)", color: "#25d366", border: "1px solid rgba(37,211,102,0.3)", borderRadius: 6, padding: "4px 10px", fontSize: 12, fontWeight: 600 }}>Connected</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            {[["API Key / Token", wa.apiKey, v => setWa({ ...wa, apiKey: v }), "text"], ["Phone Number ID", wa.phoneId, v => setWa({ ...wa, phoneId: v }), "text"], ["Business Account ID", wa.businessId, v => setWa({ ...wa, businessId: v }), "text"]].map(([l, v, fn, t]) => (
              <div key={l}>
                <label style={{ fontSize: 12, color: "#64748b", fontWeight: 600 }}>{l}</label>
                <input className="input" type={t} value={v} onChange={e => fn(e.target.value)} style={{ marginTop: 6 }} />
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
            <button className="btn-primary" onClick={() => showToast("WhatsApp settings saved!")}>Save</button>
            <button className="btn-secondary" onClick={() => showToast("Test message sent!")}>Send Test</button>
          </div>
        </div>

        {/* Email */}
        <div className="card" style={{ padding: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
            <div style={{ width: 40, height: 40, background: "rgba(99,102,241,0.15)", border: "1px solid rgba(99,102,241,0.3)", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>✉️</div>
            <div><div style={{ fontWeight: 700, color: "#e2e8f0" }}>Email (SMTP)</div><div style={{ fontSize: 12, color: "#64748b" }}>Configure outbound email server</div></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            {[["SMTP Host", email.host, v => setEmail({ ...email, host: v })], ["Port", email.port, v => setEmail({ ...email, port: v })], ["Username", email.user, v => setEmail({ ...email, user: v })], ["Password", email.pass, v => setEmail({ ...email, pass: v })]].map(([l, v, fn]) => (
              <div key={l}><label style={{ fontSize: 12, color: "#64748b", fontWeight: 600 }}>{l}</label><input className="input" value={v} onChange={e => fn(e.target.value)} style={{ marginTop: 6 }} type={l === "Password" ? "password" : "text"} /></div>
            ))}
          </div>
          <button className="btn-primary" style={{ marginTop: 16 }} onClick={() => showToast("Email settings saved!")}>Save</button>
        </div>

        {/* SMS */}
        <div className="card" style={{ padding: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
            <div style={{ width: 40, height: 40, background: "rgba(245,158,11,0.15)", border: "1px solid rgba(245,158,11,0.3)", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>📱</div>
            <div><div style={{ fontWeight: 700, color: "#e2e8f0" }}>SMS Gateway</div><div style={{ fontSize: 12, color: "#64748b" }}>Twilio / MSG91 / Fast2SMS</div></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            {[["Provider", sms.provider, v => setSms({ ...sms, provider: v })], ["Account SID", sms.sid, v => setSms({ ...sms, sid: v })], ["Auth Token", sms.token, v => setSms({ ...sms, token: v })], ["From Number", sms.from, v => setSms({ ...sms, from: v })]].map(([l, v, fn]) => (
              <div key={l}><label style={{ fontSize: 12, color: "#64748b", fontWeight: 600 }}>{l}</label><input className="input" value={v} onChange={e => fn(e.target.value)} style={{ marginTop: 6 }} /></div>
            ))}
          </div>
          <button className="btn-primary" style={{ marginTop: 16 }} onClick={() => showToast("SMS settings saved!")}>Save</button>
        </div>
      </div>
    </div>
  );
}

// ─── IMPORT MODAL ─────────────────────────────────────────────────────────────
function ImportModal({ onClose, onImport }) {
  const [dragOver, setDragOver] = useState(false);
  const [preview, setPreview] = useState(null);
  const [fileName, setFileName] = useState(null);
  const fileRef = useRef();

  const processFile = (file) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);
        setPreview(data.slice(0, 5));
      } catch {
        setPreview([{ "Error": "Could not parse file. Use a valid XLSX or CSV." }]);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleDrop = (e) => {
    e.preventDefault(); setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  };

  const handleConfirm = () => {
    if (preview && preview.length > 0 && !preview[0]["Error"]) onImport(preview);
    else {
      const sample = [
        { "Party Name": "Sample Traders", "GST": "24AABCS0001A1Z1", "Phone": "9999988888", "Email": "sample@test.com", "Invoice Date": "2024-01-01", "Due Date": "2024-02-01", "Amount": "150000", "Paid": "0" },
        { "Party Name": "Demo Industries", "GST": "24AABCD0002B2Z2", "Phone": "9999977777", "Email": "demo@test.com", "Invoice Date": "2024-01-05", "Due Date": "2024-02-05", "Amount": "250000", "Paid": "50000" },
      ];
      onImport(sample);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal fade-in" style={{ maxWidth: 640 }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 18, color: "#fff" }}>Import Tally Export</div>
            <div style={{ fontSize: 13, color: "#64748b", marginTop: 4 }}>Upload XLSX or CSV from Tally ERP / Prime</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 22 }}>×</button>
        </div>

        <div
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
          onClick={() => fileRef.current.click()}
          style={{ border: `2px dashed ${dragOver ? "#6366f1" : "#2d3148"}`, borderRadius: 12, padding: 40, textAlign: "center", cursor: "pointer", background: dragOver ? "rgba(99,102,241,0.05)" : "transparent", transition: "all 0.2s", marginBottom: 16 }}>
          <input ref={fileRef} type="file" accept=".xlsx,.csv,.xls" style={{ display: "none" }} onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
          <div style={{ fontSize: 40, marginBottom: 12 }}>📁</div>
          <div style={{ color: "#e2e8f0", fontWeight: 600, marginBottom: 6 }}>
            {fileName || "Drop your Tally export here"}
          </div>
          <div style={{ color: "#64748b", fontSize: 13 }}>Supports XLSX, XLS, CSV • Max 10MB</div>
          {preview && <div style={{ marginTop: 10, color: "#10b981", fontSize: 13 }}>✓ {preview.length} rows detected</div>}
        </div>

        {preview && (
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 12, color: "#64748b", fontWeight: 600, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>Preview (first 5 rows)</div>
            <div style={{ overflowX: "auto", background: "#13151f", borderRadius: 8, border: "1px solid #2d3148" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead>
                  <tr style={{ borderBottom: "1px solid #2d3148" }}>
                    {Object.keys(preview[0]).slice(0, 6).map(k => (
                      <th key={k} style={{ padding: "8px 12px", textAlign: "left", color: "#64748b", fontWeight: 600 }}>{k}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.map((row, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #1e2235" }}>
                      {Object.values(row).slice(0, 6).map((v, j) => (
                        <td key={j} style={{ padding: "7px 12px", color: "#94a3b8" }}>{String(v).slice(0, 20)}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div style={{ padding: 14, background: "rgba(99,102,241,0.08)", borderRadius: 8, fontSize: 12, color: "#94a3b8", marginBottom: 16 }}>
          <strong style={{ color: "#a5b4fc" }}>Expected columns:</strong> Party Name, GST, Phone, Email, Invoice Date, Due Date, Amount, Paid, Sales Person
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={handleConfirm} className="btn-primary" style={{ flex: 1 }}>
            {preview ? `Import ${preview.length} Records` : "Import Demo Data"}
          </button>
          <button onClick={onClose} className="btn-secondary">Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ─── INVOICE DETAIL MODAL ─────────────────────────────────────────────────────
function InvoiceDetailModal({ invoice, salespeople, onClose, onRemind }) {
  const sp = salespeople.find(s => s.id === invoice.salespersonId);
  const sc = statusConfig[invoice.status];

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal fade-in" style={{ maxWidth: 520 }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 20, color: "#a5b4fc" }}>{invoice.id}</div>
            <span className="badge" style={{ background: sc.bg, color: sc.color, marginTop: 6 }}>{sc.icon} {sc.label}</span>
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 22 }}>×</button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 20 }}>
          {[
            ["Party", invoice.party], ["GST Number", invoice.gst],
            ["Phone", invoice.phone], ["Email", invoice.email],
            ["Invoice Date", invoice.invoiceDate], ["Due Date", invoice.dueDate],
            ["Sales Person", sp?.name || "—"], ["Reminders Sent", invoice.remindersSent],
          ].map(([l, v]) => (
            <div key={l}>
              <div style={{ fontSize: 11, color: "#64748b", fontWeight: 600, marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>{l}</div>
              <div style={{ fontSize: 14, color: "#e2e8f0" }}>{v}</div>
            </div>
          ))}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 20 }}>
          {[["Invoice Amount", invoice.amount, "#e2e8f0"], ["Amount Paid", invoice.paid, "#10b981"], ["Outstanding", invoice.outstanding, "#f87171"]].map(([l, v, c]) => (
            <div key={l} style={{ background: "#13151f", borderRadius: 10, padding: "14px 16px", textAlign: "center" }}>
              <div style={{ fontSize: 11, color: "#64748b", marginBottom: 6 }}>{l}</div>
              <div style={{ fontSize: 18, fontWeight: 700, color: c }}>{fmt(v)}</div>
            </div>
          ))}
        </div>

        {invoice.overdueDays > 0 && (
          <div style={{ padding: "12px 16px", background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 8, marginBottom: 16, fontSize: 13, color: "#f87171", fontWeight: 600 }}>
            ⚠️ Overdue by {invoice.overdueDays} days
          </div>
        )}

        {invoice.status !== "paid" && (
          <div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 10 }}>Send reminder via:</div>
            <div style={{ display: "flex", gap: 8 }}>
              {[["💬 WhatsApp", "whatsapp"], ["✉️ Email", "email"], ["📱 SMS", "sms"]].map(([label, ch]) => (
                <button key={ch} onClick={() => onRemind(invoice)} className="btn-primary" style={{ flex: 1, fontSize: 13 }}>{label}</button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── SEND REMINDER MODAL ──────────────────────────────────────────────────────
function SendReminderModal({ invoice, onClose, onSend }) {
  const [channels, setChannels] = useState(["whatsapp"]);
  const [msgType, setMsgType] = useState("overdue");

  const toggleChannel = (ch) => setChannels(prev => prev.includes(ch) ? prev.filter(c => c !== ch) : [...prev, ch]);

  const preview = `Dear ${invoice.party.split(" ")[0]},\n\nThis is a payment reminder for Invoice #${invoice.id}.\n\nInvoice Amount: ${fmt(invoice.amount)}\nAmount Paid: ${fmt(invoice.paid)}\nOutstanding: ${fmt(invoice.outstanding)}\nDue Date: ${invoice.dueDate}${invoice.overdueDays > 0 ? `\nOverdue by: ${invoice.overdueDays} days` : ""}\n\nKindly arrange the payment at the earliest.\n\nRegards,\nAccounts Team`;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal fade-in" onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 18, color: "#fff" }}>Send Payment Reminder</div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 22 }}>×</button>
        </div>

        <div style={{ padding: "12px 16px", background: "rgba(99,102,241,0.08)", borderRadius: 8, marginBottom: 16, fontSize: 13 }}>
          <span style={{ color: "#a5b4fc", fontWeight: 600 }}>{invoice.id}</span>
          <span style={{ color: "#64748b" }}> • </span>
          <span style={{ color: "#e2e8f0" }}>{invoice.party}</span>
          <span style={{ color: "#64748b" }}> • Outstanding: </span>
          <span style={{ color: "#f87171", fontWeight: 700 }}>{fmt(invoice.outstanding)}</span>
        </div>

        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 12, color: "#64748b", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>Send Via</div>
          <div style={{ display: "flex", gap: 10 }}>
            {[["💬 WhatsApp", "whatsapp", "#25d366"], ["✉️ Email", "email", "#6366f1"], ["📱 SMS", "sms", "#f59e0b"]].map(([label, ch, color]) => (
              <button key={ch} onClick={() => toggleChannel(ch)}
                style={{ flex: 1, padding: "12px", borderRadius: 10, border: `2px solid ${channels.includes(ch) ? color : "#2d3148"}`, background: channels.includes(ch) ? `${color}15` : "#13151f", color: channels.includes(ch) ? color : "#64748b", cursor: "pointer", fontWeight: 600, fontSize: 13, transition: "all 0.2s" }}>
                {label}
              </button>
            ))}
          </div>
        </div>

        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 12, color: "#64748b", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Message Preview</div>
          <div style={{ background: "#13151f", border: "1px solid #2d3148", borderRadius: 8, padding: 14, fontSize: 12, color: "#94a3b8", fontFamily: "monospace", lineHeight: 1.7, whiteSpace: "pre-wrap", maxHeight: 180, overflowY: "auto" }}>{preview}</div>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={() => channels.length > 0 && onSend(invoice, channels)} className="btn-primary" style={{ flex: 1 }} disabled={channels.length === 0}>
            🚀 Send Reminder {channels.length > 0 && `(${channels.join(", ")})`}
          </button>
          <button onClick={onClose} className="btn-secondary">Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ─── TOAST ────────────────────────────────────────────────────────────────────
function Toast({ msg, type }) {
  const colors = { success: "#10b981", error: "#ef4444", info: "#6366f1" };
  return (
    <div style={{ position: "fixed", bottom: 24, right: 24, background: "#1a1d27", border: `1px solid ${colors[type]}40`, borderRadius: 12, padding: "14px 20px", display: "flex", alignItems: "center", gap: 10, boxShadow: `0 8px 32px ${colors[type]}20`, zIndex: 9999, animation: "fadeIn 0.3s ease", maxWidth: 380 }}>
      <div style={{ width: 8, height: 8, background: colors[type], borderRadius: "50%", flexShrink: 0 }} />
      <span style={{ fontSize: 13, color: "#e2e8f0" }}>{msg}</span>
    </div>
  );
}
