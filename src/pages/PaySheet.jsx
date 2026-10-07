import { useState, useEffect, memo } from "react";
import { apiFetch, API_BASE } from "../api";
import * as XLSX from "xlsx";
import { Download } from "lucide-react";
import SearchableDropdown from "../components/SearchableDropdown";
import { getFallbackFinancialYears, pickDefaultFinancialYear } from "../utils/financialYear";
import { textColor as txColor } from "../styles/adminTheme";

const formatNumberWithCommas = (number) => {
    if (number == null) return "N/A";
    return Math.round(Number(number)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
};

const getDaysInMonth = (monthIndex, financialYear) => {
    const [startYear] = financialYear.split("-").map(Number);
    const actualYear = monthIndex < 9 ? startYear : startYear + 1;
    const monthInYear = (monthIndex + 3) % 12;
    return new Date(actualYear, monthInYear + 1, 0).getDate();
};

const monthsList = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];

const PaySheet = () => {
    const [financialYears, setFinancialYears] = useState([]);
    const [selectedYear, setSelectedYear] = useState("");
    const [employees, setEmployees] = useState([]);
    const [loading, setLoading] = useState(true);
    const [viewMode, setViewMode] = useState("yearly");
    const [selectedMonth, setSelectedMonth] = useState(0);
    const [employeeTotals, setEmployeeTotals] = useState([]);
    const [monthlyData, setMonthlyData] = useState({});
    const [deptFilter, setDeptFilter] = useState("All");
    const filteredEmployees = deptFilter === "All" ? employees : employees.filter(e => e.department === deptFilter);

    const calculateSalary = (employee, monthIndex) => {
        const fgs = Number(employee.gross_salary) || Number(employee.salary) || 0;
        const basic = Number(employee.basic_da) || Math.round(fgs * 0.5);
        const dim = getDaysInMonth(monthIndex, selectedYear);
        const pd = dim;

        const eb = Math.round((basic / dim) * pd);
        const hra = (employee.hra || employee.hra_enabled) ? Math.round(basic * 0.5) : 0;
        const eh = Math.round((hra / dim) * pd);
        const conv = (employee.conveyance_allowance || employee.conveyance_enabled) ? 1200 : 0;
        const ec = Math.round((conv / dim) * pd);
        const med = (employee.medical_allowance || employee.medical_enabled) ? 1000 : 0;
        const em = Math.round((med / dim) * pd);
        const other = (employee.other_expense_enabled || employee.other_expenses) ? Math.round(fgs - basic - hra - conv - med) : 0;
        const eo = Math.round((other / dim) * pd);
        const gross = Math.round(basic + hra + conv + med + other);
        const eg = Math.round(eb + eh + ec + em + eo);
        const pfW = Math.round(eg - eh);
        const pfEnabled = employee.employee_pf_enabled ?? employee.epfEmployee ?? true;
        const empPfEnabled = employee.employer_pf_enabled ?? employee.epfEmployer ?? true;
        const pf = (pfEnabled && empPfEnabled && pfW < 15000) ? Math.round(pfW * 0.12) : (pfEnabled && empPfEnabled ? 1800 : 0);
        const isMale = (employee.gender || '').toLowerCase() === 'male';
        // PT threshold is based on the employee's fixed monthly gross salary, not earned components
        const pt = isMale ? 200 : (fgs > 25000 ? 200 : 0);
        const esicEnabled = employee.esic_enabled ?? false;
        const esic = (esicEnabled && eg < 21000) ? Math.round(eg * 0.0075) : 0;
        const lwfEnabled = employee.lwf_enabled ?? false;
        const lwf = (lwfEnabled && [2, 8].includes(monthIndex)) ? 25 : 0;
        const tds = (employee.tds_enabled ?? employee.tdsSalary) ? Math.round(eg * 0.10) : 0;
        const totalDed = Math.round(pf + esic + pt + lwf + tds);
        const net = Math.round(eg - totalDed);
        const gratuity = (employee.gratuity_enabled ?? employee.gratuityProvision ?? true) ? Math.round(eb * 0.0481) : 0;
        const empPf = (pfEnabled && empPfEnabled) ? pf : 0;
        const empEsic = (esicEnabled && eg < 21000) ? Math.round(eg * 0.0325) : 0;
        const empLwf = (lwfEnabled && [2, 8].includes(monthIndex)) ? 75 : 0;
        const ctc = Math.round(eg + empPf + empEsic + empLwf + gratuity);

        return { fgs: Math.round(fgs), basic, eb, hra, eh, conv, ec, med, em, other, eo, gross, eg, pfW, pf, esic, pt, lwf, tds, totalDed, net, gratuity, empPf, empEsic, empLwf, ctc };
    };

    useEffect(() => {
        setLoading(true);
        Promise.all([
            apiFetch(`${API_BASE}/api/payroll/financial-years`),
            apiFetch(`${API_BASE}/api/employees`),
        ])
            .then(async ([yearsRes, empRes]) => {
                const years = await yearsRes.json();
                const emps = await empRes.json();
                const finalYears = years.length ? years : getFallbackFinancialYears();
                setFinancialYears(finalYears);
                setEmployees(emps);
                if (!selectedYear) setSelectedYear(pickDefaultFinancialYear(finalYears));
            })
            .catch(err => console.error(err))
            .finally(() => setLoading(false));
    }, []);

    useEffect(() => {
        if (!selectedYear || !employees.length) return;
        const totals = employees.map(emp => {
            const t = { id: emp.id, name: emp.full_name || emp.name || "Unknown", fgs: 0, basic: 0, eb: 0, hra: 0, eh: 0, conv: 0, ec: 0, med: 0, em: 0, other: 0, eo: 0, gross: 0, eg: 0, pfW: 0, pf: 0, esic: 0, pt: 0, lwf: 0, tds: 0, totalDed: 0, net: 0, gratuity: 0, empPf: 0, empEsic: 0, empLwf: 0, ctc: 0 };
            for (let m = 0; m < 12; m++) {
                const s = calculateSalary(emp, m);
                if (s) Object.keys(t).forEach(k => { if (k !== "id" && k !== "name") t[k] += s[k] || 0; });
            }
            return t;
        });
        setEmployeeTotals(totals);

        const monthly = {};
        employees.forEach(emp => {
            monthly[emp.id] = [];
            for (let m = 0; m < 12; m++) monthly[emp.id][m] = calculateSalary(emp, m) || {};
        });
        setMonthlyData(monthly);
    }, [selectedYear, employees]);

    const handleDownloadExcel = () => {
        let data;
        if (viewMode === "yearly") {
            data = employeeTotals.map(t => ({
                "Employee ID": t.id, "Employee Name": t.name, "Fixed Gross": formatNumberWithCommas(t.fgs), "Basic+DA": formatNumberWithCommas(t.basic),
                "Earn Basic": formatNumberWithCommas(t.eb), "HRA": formatNumberWithCommas(t.eh), "Conveyance": formatNumberWithCommas(t.ec), "Medical": formatNumberWithCommas(t.em),
                "Other": formatNumberWithCommas(t.eo), "Earn Gross": formatNumberWithCommas(t.eg), "PF": formatNumberWithCommas(t.pf), "ESIC": formatNumberWithCommas(t.esic),
                "PT": formatNumberWithCommas(t.pt), "LWF": formatNumberWithCommas(t.lwf), "TDS": formatNumberWithCommas(t.tds),
                "Total Deduction": formatNumberWithCommas(t.totalDed), "Net Payable": formatNumberWithCommas(t.net),
                "Gratuity": formatNumberWithCommas(t.gratuity), "Employer PF": formatNumberWithCommas(t.empPf),
                "Employer ESIC": formatNumberWithCommas(t.empEsic), "CTC": formatNumberWithCommas(t.ctc),
            }));
        } else {
            data = employees.map(emp => {
                const s = monthlyData[emp.id]?.[selectedMonth] || {};
                return { "Employee ID": emp.id, "Employee Name": emp.full_name || emp.name || "Unknown", "Fixed Gross": formatNumberWithCommas(s.fgs), "Basic+DA": formatNumberWithCommas(s.basic), "Earn Basic": formatNumberWithCommas(s.eb), "HRA": formatNumberWithCommas(s.eh), "Conveyance": formatNumberWithCommas(s.ec), "Medical": formatNumberWithCommas(s.em), "Other": formatNumberWithCommas(s.eo), "Earn Gross": formatNumberWithCommas(s.eg), "PF": formatNumberWithCommas(s.pf), "ESIC": formatNumberWithCommas(s.esic), "PT": formatNumberWithCommas(s.pt), "LWF": formatNumberWithCommas(s.lwf), "TDS": formatNumberWithCommas(s.tds), "Total Deduction": formatNumberWithCommas(s.totalDed), "Net Payable": formatNumberWithCommas(s.net), "Gratuity": formatNumberWithCommas(s.gratuity), "Employer PF": formatNumberWithCommas(s.empPf), "Employer ESIC": formatNumberWithCommas(s.empEsic), "CTC": formatNumberWithCommas(s.ctc) };
            });
        }
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, viewMode === "yearly" ? "Yearly" : monthsList[selectedMonth]);
        XLSX.writeFile(wb, `PaySheet_${selectedYear}_${viewMode === "yearly" ? "Yearly" : monthsList[selectedMonth]}.xlsx`);
    };

    if (loading) {
        return (<div className="flex justify-center items-center h-64"><div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-purple-600"></div></div>);
    }

    const selectStyle = { padding: "12px 16px", borderRadius: "12px", border: "1px solid #d1d5db", background: "#ffffff", color: txColor('primary'), fontSize: "13px", fontWeight: "600", outline: "none", cursor: "pointer" };
    const thStyle = { padding: "16px 14px", textAlign: "left", fontSize: "12px", fontWeight: "700", color: txColor('secondary'), borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap", background: "#f8fafc", letterSpacing: ".3px", textTransform: "uppercase", position: "sticky", top: 0, zIndex: 5 };
    const tdStyle = { padding: "14px", fontSize: "13px", color: txColor('primary'), borderBottom: "1px solid #f1f5f9", whiteSpace: "nowrap" };
    const startYear = Number(String(selectedYear).split("-")[0]);
    const medicalLabel = startYear >= 2026 ? "Mobile Internet" : "Medical";
    const columns = [
        { key: "name", label: "Employee", w: "180px" },
        { key: "fgs", label: "Fixed Gross" },
        { key: "basic", label: "Basic+DA" },
        { key: "eb", label: "Earn Basic" },
        { key: "eh", label: "HRA" },
        { key: "ec", label: "Convey." },
        { key: "em", label: medicalLabel },
        { key: "eo", label: "Other" },
        { key: "eg", label: "Earn Gross", highlight: true },
        { key: "pf", label: "PF" },
        { key: "esic", label: "ESIC" },
        { key: "pt", label: "PT" },
        { key: "lwf", label: "LWF" },
        { key: "tds", label: "TDS" },
        { key: "totalDed", label: "Total Ded.", color: "#ef4444" },
        { key: "net", label: "Net Payable", highlight: true },
        { key: "gratuity", label: "Gratuity" },
        { key: "empPf", label: "Empl. PF" },
        { key: "empEsic", label: "Empl. ESIC" },
        { key: "ctc", label: "CTC", highlight: true },
    ];

    const filteredTotals = deptFilter === "All" ? employeeTotals : employeeTotals.filter(t => employees.find(e => e.id === t.id)?.department === deptFilter);
    const rows = viewMode === "yearly" ? filteredTotals : filteredEmployees.map(emp => ({ id: emp.id, name: emp.full_name || emp.name, ...monthlyData[emp.id]?.[selectedMonth] || {} }));

    return (
        <div style={{ width: "100%" }}>
            {/* Controls */}
            <div style={{ display: "flex", flexWrap: "wrap", gap: "12px", marginBottom: "24px", alignItems: "center", justifyContent: "space-between" }}>
                <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                    {["yearly", "monthly"].map(mode => (
                        <button key={mode} onClick={() => setViewMode(mode)} style={{
                            padding: "12px 24px", borderRadius: "12px", border: viewMode === mode ? ( "1px solid rgba(8,145,178,.35)") : ( "1px solid #d1d5db"),
                            background: viewMode === mode ? ( "linear-gradient(135deg,rgba(8,145,178,.1),rgba(2,132,199,.1))") : ( "rgba(255,255,255,.8)"),
                            color: viewMode === mode ? ( "#0891b2") : txColor('secondary'), cursor: "pointer", fontWeight: "700", fontSize: "13px", letterSpacing: ".3px",
                            transition: "all .3s ease",
                        }}>
                            {mode === "yearly" ? "Yearly" : "Monthly"}
                        </button>
                    ))}
                    <select value={viewMode === "yearly" ? selectedYear : selectedMonth}
                        onChange={(e) => viewMode === "yearly" ? setSelectedYear(e.target.value) : setSelectedMonth(Number(e.target.value))}
                        style={{ ...selectStyle, marginLeft: "8px" }}>
                        {viewMode === "yearly"
                            ? financialYears.map(y => <option key={y} value={y}>{y}</option>)
                            : monthsList.map((m, i) => <option key={i} value={i}>{m}</option>)}
                    </select>
                    <SearchableDropdown value={deptFilter} onChange={setDeptFilter} width={200} />
                    <span style={{ color: "#0891b2", fontSize: "13px", fontWeight: "700", marginLeft: "8px" }}>
                        {rows.length} Employees
                    </span>
                </div>
                <button onClick={handleDownloadExcel} style={{
                    padding: "12px 24px", borderRadius: "12px", border: "1px solid rgba(34,197,94,.3)",
                    background: "linear-gradient(135deg,rgba(34,197,94,.2),rgba(16,185,129,.2))",
                    color: "#22c55e", cursor: "pointer", fontWeight: "700", fontSize: "13px",
                    display: "flex", alignItems: "center", gap: "8px", transition: ".3s",
                }}
                    onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "0 0 22px rgba(34,197,94,.15)"; }}
                    onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "none"; }}
                >
                    <Download size={16} /> Download as Excel
                </button>
            </div>

            {/* Table */}
            <div style={{ borderRadius: "20px", border: "1px solid #e2e8f0", background: "#ffffff", overflow: "hidden" }}>
                <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: "70vh" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead>
                            <tr>
                                {columns.map(col => (
                                    <th key={col.key} style={{ ...thStyle, minWidth: col.w || "auto" }}>{col.label}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {rows.map((row, idx) => (
                                <tr key={idx} style={{ transition: "background .2s" }}
                                    onMouseEnter={e => e.currentTarget.style.background = "rgba(8,145,178,.04)"}
                                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                                    {columns.map(col => (
                                        <td key={col.key} style={{
                                            ...tdStyle,
                                            color: col.key === "name" ? txColor('primary') : col.highlight ? ( "#0891b2") : ( txColor('secondary')),
                                            fontWeight: col.key === "name" ? "700" : col.highlight ? "800" : "500",
                                            textShadow: col.highlight ? ( "none") : "none",
                                        }}>
                                            {col.key === "name" ? (
                                                <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                                                    <div style={{ width: "36px", height: "36px", borderRadius: "50%", background: "linear-gradient(135deg,#0891b2,#0284c7)", display: "flex", alignItems: "center", justifyContent: "center", color: "#ffffff", fontWeight: "700", fontSize: "14px", boxShadow: "none", flexShrink: 0 }}>
                                                        {(row.name || "?").charAt(0).toUpperCase()}
                                                    </div>
                                                    <span>{row.name}</span>
                                                </div>
                                            ) : (
                                                <>₹{formatNumberWithCommas(row[col.key])}</>
                                            )}
                                        </td>
                                    ))}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

export default memo(PaySheet);
