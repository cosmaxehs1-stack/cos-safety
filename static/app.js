let TOKEN = sessionStorage.getItem("token") || "";
let chartInstances = {};
let locationViewMode = "grade"; // "grade" or "disaster"
let lastSummaryData = null;
let editingRecordId = null; // null = add mode, string = edit mode

// --- Auth ---
async function login() {
    const pw = document.getElementById("password-input").value;
    const errEl = document.getElementById("login-error");
    try {
        const res = await fetch("/api/login", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ password: pw }),
        });
        if (!res.ok) {
            errEl.textContent = "비밀번호가 올바르지 않습니다.";
            return;
        }
        const data = await res.json();
        TOKEN = data.token;
        sessionStorage.setItem("token", TOKEN);
        showDashboard();
    } catch (e) {
        errEl.textContent = "서버 연결 실패";
    }
}

function logout() {
    TOKEN = "";
    sessionStorage.removeItem("token");
    document.getElementById("login-screen").style.display = "flex";
    document.getElementById("dashboard").style.display = "none";
}

let initialLoad = true;
function showDashboard() {
    document.getElementById("login-screen").style.display = "none";
    document.getElementById("dashboard").style.display = "block";
    initialLoad = true;
    fetchSummary();
}

function authHeaders() {
    return { Authorization: "Bearer " + TOKEN };
}

// --- Upload ---
async function uploadFile(input) {
    const file = input.files[0];
    if (!file) return;
    const channel = document.getElementById("upload-channel").value;
    const formData = new FormData();
    formData.append("file", file);
    formData.append("channel", channel);

    try {
        const res = await fetch("/api/upload", {
            method: "POST",
            headers: authHeaders(),
            body: formData,
        });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();
        alert(data.message || "업로드 완료");
        fetchSummary();
    } catch (e) {
        alert("업로드 실패: " + e.message);
    }
    input.value = "";
}

// --- Filters ---
function getFilterValue(id, fallback) {
    const el = document.getElementById(id);
    return el ? el.value : (fallback || "전체");
}

function getFilters() {
    return {
        team: getFilterValue("f-team"),
        channel: getFilterValue("f-channel"),
        year: getFilterValue("f-year"),
        month: getFilterValue("f-month"),
        location: getFilterValue("f-location"),
        grade: getFilterValue("f-grade"),
        disaster_type: getFilterValue("f-disaster"),
        process: getFilterValue("f-process"),
        person: getFilterValue("f-person"),
        week: getFilterValue("f-week", "0"),
        completion: getFilterValue("f-completion"),
        repeat: getFilterValue("f-repeat"),
        keyword: (document.getElementById("f-keyword") || {}).value || "",
    };
}

function setFilterValue(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val;
}

function resetFilters() {
    setFilterValue("f-team", "전체");
    setFilterValue("f-channel", "전체");
    setFilterValue("f-year", "전체");
    setFilterValue("f-month", "전체");
    setFilterValue("f-location", "전체");
    setFilterValue("f-grade", "전체");
    setFilterValue("f-disaster", "전체");
    setFilterValue("f-process", "전체");
    setFilterValue("f-person", "전체");
    setFilterValue("f-week", "0");
    setFilterValue("f-completion", "전체");
    setFilterValue("f-repeat", "전체");
    setFilterValue("f-keyword", "");
    fetchSummary();
}

function populateFilter(selectId, options, keepValue) {
    const sel = document.getElementById(selectId);
    if (!sel || !options) return;
    const prev = sel.value;
    const defaultOption = sel.options[0].outerHTML;
    sel.innerHTML = defaultOption;
    options.forEach(opt => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        sel.appendChild(o);
    });
    if (keepValue && options.includes(prev)) sel.value = prev;
}

// --- Fetch Data ---
async function fetchSummary() {
    const filters = getFilters();
    const params = new URLSearchParams();
    Object.entries(filters).forEach(([k, v]) => {
        if (v && v !== "전체" && v !== "0") params.append(k, v);
    });

    try {
        const res = await fetch("/api/summary?" + params.toString(), {
            headers: authHeaders(),
        });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();

        if (data.total === 0 && !filters.keyword && filters.month === "전체") {
            document.getElementById("no-data").style.display = "block";
        } else {
            document.getElementById("no-data").style.display = "none";
        }

        // Client-side repeat filter
        let displayRecords = data.records;
        const repeatFilter = document.getElementById("f-repeat").value;
        if (repeatFilter === "반복") {
            displayRecords = displayRecords.filter(r => r.is_repeat);
        } else if (repeatFilter === "단건") {
            displayRecords = displayRecords.filter(r => !r.is_repeat);
        }

        lastSummaryData = data;
        updateStats(data);
        updateCharts(data);
        updateTable(displayRecords);
        updateFilters(data.filters);
        document.getElementById("total-badge").textContent = data.total + " cases";
    } catch (e) {
        console.error("Fetch error:", e);
    }
}

// --- Stats ---
function updateStats(data) {
    document.getElementById("s-total").textContent = data.total;
    document.getElementById("s-a").textContent = data.grade_a;
    document.getElementById("s-b").textContent = data.grade_b;
    document.getElementById("s-c").textContent = data.grade_c;
    document.getElementById("s-d").textContent = data.grade_d;
    document.getElementById("s-a-after").textContent = data.grade_a_after || 0;
    document.getElementById("s-b-after").textContent = data.grade_b_after || 0;
    document.getElementById("s-c-after").textContent = data.grade_c_after || 0;
    document.getElementById("s-d-after").textContent = data.grade_d_after || 0;
    document.getElementById("s-complete").textContent = data.complete;
    document.getElementById("s-pending").textContent = data.incomplete;
    document.getElementById("s-repeat").textContent = data.repeat_total || 0;
    const rate = data.improvement_rate != null ? data.improvement_rate : 0;
    const rateEl = document.getElementById("s-improvement");
    if (rateEl) {
        rateEl.textContent = rate + "%";
    }
}

// --- Charts ---
const GRADE_COLORS = {
    A: "#27ae60",
    B: "#3498db",
    C: "#f39c12",
    D: "#e74c3c",
    "-": "#bdc3c7",
};

function destroyChart(id) {
    if (chartInstances[id]) {
        chartInstances[id].destroy();
        delete chartInstances[id];
    }
}

const DISASTER_COLORS = [
    "#e74c3c", "#3498db", "#f39c12", "#27ae60", "#9b59b6",
    "#1abc9c", "#e67e22", "#34495e", "#e91e63", "#00bcd4",
    "#8bc34a", "#ff5722", "#607d8b", "#795548", "#cddc39",
];

function toggleLocationView() {
    locationViewMode = locationViewMode === "grade" ? "disaster" : "grade";
    const btn = document.getElementById("btn-loc-toggle");
    if (locationViewMode === "grade") {
        btn.textContent = "재해유형별";
        btn.classList.remove("active");
    } else {
        btn.textContent = "등급별";
        btn.classList.add("active");
    }
    if (lastSummaryData) renderLocationChart(lastSummaryData);
}

function renderLocationChart(data) {
    destroyChart("chart-location");

    const MAJOR_ORDER = ["화성", "평택", "고렴", "판교", "기타"];
    const hierarchy = data.location_hierarchy || {};
    const subStats = data.location_stats || {};
    const subDisaster = data.location_disaster_stats || {};

    // Build flat list of all sub-locations, ordered by major category
    const locLabels = [];
    const majorGroupMap = []; // tracks which major group each label belongs to
    MAJOR_ORDER.forEach(major => {
        const subs = hierarchy[major] || [];
        subs.forEach(sub => {
            locLabels.push(sub);
            majorGroupMap.push(major);
        });
    });

    if (locLabels.length === 0) return;

    let datasets;
    let stacked = false;

    if (locationViewMode === "disaster") {
        stacked = true;
        const allTypes = new Set();
        locLabels.forEach(l => {
            const obj = subDisaster[l] || {};
            Object.keys(obj).forEach(k => allTypes.add(k));
        });
        const typeList = [...allTypes].sort((a, b) => {
            const totalA = locLabels.reduce((s, l) => s + ((subDisaster[l] || {})[a] || 0), 0);
            const totalB = locLabels.reduce((s, l) => s + ((subDisaster[l] || {})[b] || 0), 0);
            return totalB - totalA;
        });
        datasets = typeList.map((dt, i) => ({
            label: dt,
            data: locLabels.map(l => (subDisaster[l] || {})[dt] || 0),
            backgroundColor: DISASTER_COLORS[i % DISASTER_COLORS.length],
            borderRadius: 4,
        }));
    } else {
        datasets = ["A", "B", "C", "D"].map(g => ({
            label: g + "등급",
            data: locLabels.map(l => (subStats[l] || {})[g] || 0),
            backgroundColor: GRADE_COLORS[g],
            borderRadius: 4,
        }));
    }

    // Build short sub-location labels (strip major prefix for readability)
    const displayLabels = locLabels.map((sub, i) => {
        const major = majorGroupMap[i];
        return sub.startsWith(major) ? sub.slice(major.length) || sub : sub;
    });

    // Compute group ranges for background bands and major labels
    const majorGroups = [];
    let groupStart = 0;
    for (let i = 1; i <= majorGroupMap.length; i++) {
        if (i === majorGroupMap.length || majorGroupMap[i] !== majorGroupMap[i - 1]) {
            majorGroups.push({ major: majorGroupMap[groupStart], start: groupStart, end: i - 1 });
            groupStart = i;
        }
    }

    const BAND_COLORS = ["rgba(59,130,246,0.06)", "rgba(16,185,129,0.06)", "rgba(245,158,11,0.06)", "rgba(139,92,246,0.06)", "rgba(107,114,128,0.06)"];

    const groupPlugin = {
        id: "locationGroupBands",
        beforeDraw(chart) {
            const ctx = chart.ctx;
            const xScale = chart.scales.x;
            const yScale = chart.scales.y;
            const chartArea = chart.chartArea;
            ctx.save();

            // Draw alternating background bands
            majorGroups.forEach((g, gi) => {
                const x1 = xScale.getPixelForValue(g.start) - (xScale.getPixelForValue(1) - xScale.getPixelForValue(0)) / 2;
                const x2 = xScale.getPixelForValue(g.end) + (xScale.getPixelForValue(1) - xScale.getPixelForValue(0)) / 2;
                ctx.fillStyle = BAND_COLORS[gi % BAND_COLORS.length];
                ctx.fillRect(x1, chartArea.top, x2 - x1, chartArea.bottom - chartArea.top);
            });

            ctx.restore();
        },
        afterDraw(chart) {
            const ctx = chart.ctx;
            const xScale = chart.scales.x;
            const chartArea = chart.chartArea;
            ctx.save();

            // Draw vertical separator lines between major groups
            for (let i = 1; i < majorGroups.length; i++) {
                const prevEnd = majorGroups[i - 1].end;
                const currStart = majorGroups[i].start;
                const x = (xScale.getPixelForValue(prevEnd) + xScale.getPixelForValue(currStart)) / 2;
                ctx.strokeStyle = "#d1d5db";
                ctx.lineWidth = 1;
                ctx.setLineDash([]);
                ctx.beginPath();
                ctx.moveTo(x, chartArea.top);
                ctx.lineTo(x, chartArea.bottom + 20);
                ctx.stroke();
            }

            // Draw major category label below the x-axis ticks
            const labelY = chartArea.bottom + 38;
            majorGroups.forEach((g, gi) => {
                const x1 = xScale.getPixelForValue(g.start);
                const x2 = xScale.getPixelForValue(g.end);
                const cx = (x1 + x2) / 2;

                // Draw bracket line
                const bracketY = chartArea.bottom + 24;
                const bx1 = x1 - 4;
                const bx2 = x2 + 4;
                ctx.strokeStyle = "#999";
                ctx.lineWidth = 1;
                ctx.setLineDash([]);
                ctx.beginPath();
                ctx.moveTo(bx1, bracketY);
                ctx.lineTo(bx1, bracketY + 4);
                ctx.lineTo(bx2, bracketY + 4);
                ctx.lineTo(bx2, bracketY);
                ctx.stroke();

                // Draw major label
                ctx.fillStyle = "#374151";
                ctx.font = "bold 12px sans-serif";
                ctx.textAlign = "center";
                ctx.textBaseline = "top";
                ctx.fillText(g.major, cx, labelY);
            });

            ctx.restore();
        }
    };

    chartInstances["chart-location"] = new Chart(
        document.getElementById("chart-location"),
        {
            type: "bar",
            data: { labels: displayLabels, datasets },
            plugins: [groupPlugin],
            options: {
                responsive: true,
                layout: {
                    padding: { bottom: 30 }
                },
                plugins: {
                    legend: { display: true, position: "top" },
                    tooltip: {
                        callbacks: {
                            title: function(items) {
                                const idx = items[0].dataIndex;
                                return locLabels[idx];
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: stacked,
                        grid: { display: false },
                        ticks: {
                            autoSkip: false,
                            maxRotation: 0,
                            font: { size: 11 },
                        }
                    },
                    y: { stacked: stacked, beginAtZero: true, ticks: { stepSize: 5 } },
                },
            },
        }
    );
}

function updateCharts(data) {
    // 1. Location bar chart
    renderLocationChart(data);


    // 2. Monthly effort chart (발굴 vs 개선 + 개선률 line)
    destroyChart("chart-monthly-effort");
    const effort = data.monthly_effort || {};
    const currentMonth = new Date().getMonth() + 1;
    const effortStart = currentMonth <= 6 ? 1 : currentMonth - 5;
    const effortMonths = [];
    for (let i = effortStart; i <= effortStart + 5; i++) effortMonths.push(i + "월");

    const riskTrend = data.risk_trend || {};
    const riskBefore = effortMonths.map(m => (riskTrend[m] || {}).avg_grade_before || null);
    const riskAfter = effortMonths.map(m => (riskTrend[m] || {}).avg_grade_after || null);

    chartInstances["chart-monthly-effort"] = new Chart(
        document.getElementById("chart-monthly-effort"),
        {
            type: "bar",
            data: {
                labels: effortMonths,
                datasets: [
                    {
                        label: "개선 전",
                        data: riskBefore,
                        backgroundColor: "#e74c3c",
                        borderRadius: 4,
                    },
                    {
                        label: "개선 후",
                        data: riskAfter,
                        backgroundColor: "#27ae60",
                        borderRadius: 4,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: "top" },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                const v = ctx.parsed.y;
                                if (v === null) return null;
                                const grade = v <= 1.5 ? "A" : v <= 2.5 ? "B" : v <= 3.5 ? "C" : "D";
                                return ctx.dataset.label + ": " + v.toFixed(1) + " (" + grade + "급)";
                            }
                        }
                    }
                },
                scales: {
                    x: { grid: { display: false } },
                    y: {
                        min: 0, max: 4,
                        title: { display: true, text: "평균 등급" },
                        ticks: {
                            stepSize: 1,
                            callback: function(v) {
                                return {1: "A", 2: "B", 3: "C", 4: "D"}[v] || "";
                            }
                        },
                    },
                },
            },
        }
    );

    // 2a. Monthly grade stacked bar chart (6-month rolling window)
    destroyChart("chart-grade-monthly");
    const gradeTrend = data.grade_trend || {};
    const gradeTrendAfter = data.grade_trend_after || {};
    const gmMonths = effortMonths;

    // Build labels: ["1월 전", "1월 후", "2월 전", "2월 후", ...]
    const gmLabels = [];
    const gmData = { dB: [], cB: [], bB: [], aB: [], dA: [], cA: [], bA: [], aA: [] };
    for (const m of gmMonths) {
        gmLabels.push(m + " 전");
        gmLabels.push(m + " 후");
        const b = gradeTrend[m] || {};
        const a = gradeTrendAfter[m] || {};
        gmData.dB.push(b.D || 0); gmData.dB.push(null);
        gmData.cB.push(b.C || 0); gmData.cB.push(null);
        gmData.bB.push(b.B || 0); gmData.bB.push(null);
        gmData.aB.push(b.A || 0); gmData.aB.push(null);
        gmData.dA.push(null); gmData.dA.push(a.D || 0);
        gmData.cA.push(null); gmData.cA.push(a.C || 0);
        gmData.bA.push(null); gmData.bA.push(a.B || 0);
        gmData.aA.push(null); gmData.aA.push(a.A || 0);
    }

    chartInstances["chart-grade-monthly"] = new Chart(
        document.getElementById("chart-grade-monthly"),
        {
            type: "bar",
            data: {
                labels: gmLabels,
                datasets: [
                    { label: "D등급", data: gmData.dB, backgroundColor: GRADE_COLORS.D, borderRadius: 3, stack: "before", skipNull: true },
                    { label: "C등급", data: gmData.cB, backgroundColor: GRADE_COLORS.C + "99", borderRadius: 3, stack: "before", skipNull: true },
                    { label: "B등급", data: gmData.bB, backgroundColor: GRADE_COLORS.B + "99", borderRadius: 3, stack: "before", skipNull: true },
                    { label: "A등급", data: gmData.aB, backgroundColor: GRADE_COLORS.A + "99", borderRadius: 3, stack: "before", skipNull: true },
                    { label: "D등급(후)", data: gmData.dA, backgroundColor: GRADE_COLORS.D, borderRadius: 3, stack: "after", skipNull: true },
                    { label: "C등급(후)", data: gmData.cA, backgroundColor: GRADE_COLORS.C + "99", borderRadius: 3, stack: "after", skipNull: true },
                    { label: "B등급(후)", data: gmData.bA, backgroundColor: GRADE_COLORS.B + "99", borderRadius: 3, stack: "after", skipNull: true },
                    { label: "A등급(후)", data: gmData.aA, backgroundColor: GRADE_COLORS.A + "99", borderRadius: 3, stack: "after", skipNull: true },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                skipNull: true,
                plugins: {
                    legend: {
                        position: "top",
                        labels: {
                            generateLabels: function() {
                                return [
                                    { text: "D등급", fillStyle: GRADE_COLORS.D, strokeStyle: "transparent", lineWidth: 0 },
                                    { text: "C등급", fillStyle: GRADE_COLORS.C, strokeStyle: "transparent", lineWidth: 0 },
                                    { text: "B등급", fillStyle: GRADE_COLORS.B, strokeStyle: "transparent", lineWidth: 0 },
                                    { text: "A등급", fillStyle: GRADE_COLORS.A, strokeStyle: "transparent", lineWidth: 0 },
                                ];
                            }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                if (ctx.parsed.y === null || ctx.parsed.y === 0) return null;
                                const name = ctx.dataset.label.replace("(후)", "");
                                return name + ": " + ctx.parsed.y + "건";
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        grid: {
                            display: true,
                            color: function(ctx) {
                                // 매 2번째 라인(월 경계)에만 세로줄
                                return ctx.index % 2 === 0 ? "#e5e7eb" : "transparent";
                            },
                        },
                    },
                    y: { stacked: true, beginAtZero: true, title: { display: true, text: "건수" } },
                },
                barPercentage: 0.5,
                categoryPercentage: 0.7,
            },
        }
    );

    // 2b. Grade before/after mini donuts
    destroyChart("chart-grade-before");
    destroyChart("chart-grade-after");
    const gradeColors = [GRADE_COLORS.A, GRADE_COLORS.B, GRADE_COLORS.C, GRADE_COLORS.D];
    const gradeMiniOpts = {
        responsive: true, cutout: "60%",
        plugins: { legend: { display: false }, tooltip: {
            callbacks: { label: function(ctx) { return ctx.label + ": " + ctx.parsed + "건"; } }
        }}
    };
    chartInstances["chart-grade-before"] = new Chart(
        document.getElementById("chart-grade-before"),
        { type: "doughnut", data: { labels: ["A","B","C","D"],
            datasets: [{ data: [data.grade_a, data.grade_b, data.grade_c, data.grade_d], backgroundColor: gradeColors, borderWidth: 0 }]
        }, options: gradeMiniOpts }
    );
    const afterA = data.grade_a_after || 0;
    const afterB = data.grade_b_after || 0;
    const afterC = data.grade_c_after || 0;
    const afterD = data.grade_d_after || 0;
    const afterIncomplete = data.incomplete || 0;
    chartInstances["chart-grade-after"] = new Chart(
        document.getElementById("chart-grade-after"),
        { type: "doughnut", data: { labels: ["A","B","C","D","미완료"],
            datasets: [{ data: [afterA, afterB, afterC, afterD, afterIncomplete > 0 ? afterIncomplete : 0], backgroundColor: [...gradeColors, "#e5e7eb"], borderWidth: 0 }]
        }, options: gradeMiniOpts }
    );

    // 3. Completion donut
    destroyChart("chart-completion");
    const compTotal = data.complete + data.incomplete;
    const compPct = compTotal > 0 ? Math.round(data.complete / compTotal * 100) : 0;
    const incompPct = compTotal > 0 ? 100 - compPct : 0;

    chartInstances["chart-completion"] = new Chart(
        document.getElementById("chart-completion"),
        {
            type: "doughnut",
            data: {
                labels: ["완료", "미완료"],
                datasets: [{
                    data: [data.complete, data.incomplete],
                    backgroundColor: ["#27ae60", "#e5e7eb"],
                    borderWidth: 0,
                }],
            },
            options: {
                responsive: true,
                cutout: "65%",
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                return ctx.label + ": " + ctx.parsed + "건";
                            }
                        }
                    }
                },
            },
        }
    );

    // 4. Weekly bar chart with cumulative target line (recent 3 months)
    destroyChart("chart-week");
    const allWeekLabels = Object.keys(data.week_stats);
    const allWeekData = allWeekLabels.map(k => data.week_stats[k]);
    const weeklyTarget = Math.ceil(2000 / 52);

    // Cumulative over all weeks
    let cumulAll = 0;
    const allCumulData = allWeekData.map(v => { cumulAll += v; return cumulAll; });
    const allCumulTarget = allWeekLabels.map((_, i) => weeklyTarget * (i + 1));

    // Slice to recent 3 months (~13 weeks)
    const recent3m = 13;
    const startIdx = Math.max(0, allWeekLabels.length - recent3m);
    const weekLabels = allWeekLabels.slice(startIdx);
    const weekData = allWeekData.slice(startIdx);
    const cumulData = allCumulData.slice(startIdx);
    const cumulTarget = allCumulTarget.slice(startIdx);

    chartInstances["chart-week"] = new Chart(
        document.getElementById("chart-week"),
        {
            type: "bar",
            data: {
                labels: weekLabels,
                datasets: [
                    {
                        label: "주간 발굴",
                        data: weekData,
                        backgroundColor: "#3498db",
                        borderRadius: 4,
                        order: 3,
                        yAxisID: "y",
                    },
                    {
                        label: "누적 발굴",
                        data: cumulData,
                        type: "line",
                        borderColor: "#27ae60",
                        borderWidth: 2.5,
                        pointRadius: 4,
                        pointBackgroundColor: "#27ae60",
                        tension: 0.3,
                        fill: false,
                        order: 1,
                        yAxisID: "y1",
                    },
                    {
                        label: "누적 목표 (연 2,000건)",
                        data: cumulTarget,
                        type: "line",
                        borderColor: "#95a5a6",
                        borderWidth: 2,
                        borderDash: [6, 4],
                        pointRadius: 0,
                        fill: false,
                        order: 2,
                        yAxisID: "y1",
                    },
                ],
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: "top", labels: { font: { size: 11 } } },
                },
                scales: {
                    x: { grid: { display: false } },
                    y: { beginAtZero: true, position: "left", title: { display: true, text: "주간 건수" } },
                    y1: { beginAtZero: true, position: "right", title: { display: true, text: "누적 건수" }, grid: { display: false } },
                },
            },
        }
    );

    // 5. Disaster type chart (category tab)
    destroyChart("chart-disaster");
    const disLabels = Object.keys(data.disaster_stats);
    const disData = disLabels.map(k => data.disaster_stats[k]);
    const disColors = disLabels.map((_, i) =>
        ["#e74c3c", "#3498db", "#f39c12", "#27ae60", "#9b59b6", "#1abc9c", "#e67e22", "#34495e", "#e91e63", "#00bcd4"][i % 10]
    );
    chartInstances["chart-disaster"] = new Chart(
        document.getElementById("chart-disaster"),
        {
            type: "doughnut",
            data: {
                labels: disLabels,
                datasets: [{
                    data: disData,
                    backgroundColor: disColors,
                    borderWidth: 0,
                }],
            },
            options: {
                responsive: true,
                cutout: "50%",
                plugins: { legend: { position: "right" } },
            },
        }
    );

    // 6. Process chart (category tab)
    destroyChart("chart-process");
    const procLabels = Object.keys(data.process_stats);
    const procData = procLabels.map(k => data.process_stats[k]);
    const procColors = procLabels.map((_, i) =>
        ["#3498db", "#e74c3c", "#27ae60", "#f39c12", "#9b59b6", "#1abc9c", "#e67e22", "#34495e", "#e91e63", "#00bcd4"][i % 10]
    );
    chartInstances["chart-process"] = new Chart(
        document.getElementById("chart-process"),
        {
            type: "bar",
            data: {
                labels: procLabels,
                datasets: [{
                    label: "건수",
                    data: procData,
                    backgroundColor: procColors,
                    borderRadius: 4,
                }],
            },
            options: {
                responsive: true,
                indexAxis: "y",
                plugins: { legend: { display: false } },
                scales: {
                    x: { beginAtZero: true },
                    y: { grid: { display: false } },
                },
            },
        }
    );

    // 7. Channel summary table (visible only with multi-channel data)
    const chStats = data.channel_stats || {};
    const chGradeStats = data.channel_grade_stats || {};
    const chKeys = Object.keys(chStats);

    if (chKeys.length > 1) {
        const chSorted = chKeys.sort((a, b) => chStats[b] - chStats[a]);
        const tableWrap = document.getElementById("channel-table-wrap");
        tableWrap.style.display = "";
        const tbody = document.getElementById("channel-summary-tbody");
        tbody.innerHTML = "";
        let totalRow = { count: 0, A: 0, B: 0, C: 0, D: 0, comp: 0, incomp: 0 };
        chSorted.forEach(ch => {
            const g = chGradeStats[ch] || {};
            const count = chStats[ch] || 0;
            const A = g.A || 0, B = g.B || 0, C = g.C || 0, D = g.D || 0;
            const comp = g.complete || 0, incomp = g.incomplete || 0;
            const chRate = count > 0 ? (comp / count * 100).toFixed(1) : 0;
            totalRow.count += count; totalRow.A += A; totalRow.B += B;
            totalRow.C += C; totalRow.D += D; totalRow.comp += comp; totalRow.incomp += incomp;
            const tr = document.createElement("tr");
            tr.innerHTML =
                "<td>" + escapeHtml(ch) + "</td>" +
                "<td><strong>" + count + "</strong></td>" +
                '<td class="green">' + A + "</td>" +
                '<td class="blue">' + B + "</td>" +
                '<td class="orange">' + C + "</td>" +
                '<td class="red">' + D + "</td>" +
                '<td class="status-complete">' + comp + "</td>" +
                '<td class="status-incomplete">' + incomp + "</td>" +
                '<td style="font-weight:600;color:' + (chRate >= 80 ? '#27ae60' : chRate >= 50 ? '#f39c12' : '#e74c3c') + '">' + chRate + '%</td>';
            tbody.appendChild(tr);
        });
        // Total row
        const totalRate = totalRow.count > 0 ? (totalRow.comp / totalRow.count * 100).toFixed(1) : 0;
        const totalTr = document.createElement("tr");
        totalTr.style.background = "#f0f4ff";
        totalTr.style.fontWeight = "700";
        totalTr.innerHTML =
            "<td>합계</td>" +
            "<td>" + totalRow.count + "</td>" +
            '<td class="green">' + totalRow.A + "</td>" +
            '<td class="blue">' + totalRow.B + "</td>" +
            '<td class="orange">' + totalRow.C + "</td>" +
            '<td class="red">' + totalRow.D + "</td>" +
            '<td class="status-complete">' + totalRow.comp + "</td>" +
            '<td class="status-incomplete">' + totalRow.incomp + "</td>" +
            '<td style="color:' + (totalRate >= 80 ? '#27ae60' : totalRate >= 50 ? '#f39c12' : '#e74c3c') + '">' + totalRate + '%</td>';
        tbody.appendChild(totalTr);
    } else {
        document.getElementById("channel-table-wrap").style.display = "none";
    }
}

// --- Table ---
function updateTable(records) {
    const tbody = document.getElementById("data-tbody");
    tbody.innerHTML = "";
    records.forEach(r => {
        const tr = document.createElement("tr");
        const imgBefore = r.image
            ? '<img src="' + escapeHtml(r.image) + '" class="table-thumb" onclick="showImageModal(\'' + escapeHtml(r.image) + '\')">'
            : '-';
        const imgAfter = r.image_after
            ? '<img src="' + escapeHtml(r.image_after) + '" class="table-thumb" onclick="showImageModal(\'' + escapeHtml(r.image_after) + '\')">'
            : '-';
        const rid = escapeHtml(r._id || "");
        tr.innerHTML =
            '<td>' + r.no + '</td>' +
            '<td>' + escapeHtml(r.month) + '</td>' +
            '<td>' + escapeHtml(r.person) + '</td>' +
            '<td>' + (r.date || "-") + '</td>' +
            '<td>' + escapeHtml(r.location || "-") + '</td>' +
            '<td title="' + escapeHtml(r.content_full) + '">' + escapeHtml(r.content) + '</td>' +
            '<td>' + escapeHtml(r.disaster_type || "-") + '</td>' +
            '<td><span class="grade-badge grade-' + r.grade_before + '">' + r.grade_before + '</span></td>' +
            '<td><span class="grade-badge grade-' + (r.grade_after || "-") + '">' + (r.grade_after || "-") + '</span></td>' +
            '<td class="' + (r.completion === "완료" ? "status-complete" : "status-incomplete") + '">' + (r.completion || "-") + '</td>' +
            '<td>' + (r.is_repeat ? '<span class="repeat-badge">' + r.repeat_count + '회</span>' : '<span class="repeat-badge single">1회</span>') + '</td>' +
            '<td>' + (r.week || "-") + '</td>' +
            '<td>' + imgBefore + '</td>' +
            '<td>' + imgAfter + '</td>' +
            '<td class="action-cell">' +
                '<button class="btn-edit" onclick="editRecord(\'' + rid + '\')">수정</button>' +
                '<button class="btn-row-del" onclick="deleteRecord(\'' + rid + '\')">삭제</button>' +
            '</td>';
        tbody.appendChild(tr);
    });
}

function escapeHtml(str) {
    if (!str) return "";
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// --- Filters Update ---
function updateFilters(filters) {
    if (filters.channels) populateFilter("f-channel", filters.channels, true);
    if (filters.years) populateFilter("f-year", filters.years, true);
    populateFilter("f-month", filters.months, true);
    populateFilter("f-location", filters.locations, true);
    populateFilter("f-disaster", filters.disaster_types, true);
    populateFilter("f-process", filters.processes, true);
    populateFilter("f-person", filters.persons, true);

    const weekSel = document.getElementById("f-week");
    const prevWeek = weekSel.value;
    weekSel.innerHTML = '<option value="0">전체</option>';
    filters.weeks.forEach(w => {
        const o = document.createElement("option");
        o.value = w;
        o.textContent = w + "주차";
        weekSel.appendChild(o);
    });
    if (prevWeek !== "0") weekSel.value = prevWeek;

    // 첫 로드 시 2026년 기본 선택 후 재조회
    if (initialLoad && filters.years && filters.years.includes("2026")) {
        initialLoad = false;
        setFilterValue("f-year", "2026");
        fetchSummary();
    }
}

// --- Tabs ---
function switchTab(tabName, btn) {
    document.querySelectorAll(".tab-content").forEach(el => el.style.display = "none");
    document.querySelectorAll(".tab").forEach(el => el.classList.remove("active"));
    document.getElementById("tab-" + tabName).style.display = "block";
    btn.classList.add("active");
}

// --- Data Management ---
async function showManageModal() {
    document.getElementById("manage-modal").style.display = "flex";
    const list = document.getElementById("channel-status-list");
    list.innerHTML = '<div style="text-align:center;padding:20px;color:#888;">불러오는 중...</div>';

    try {
        const res = await fetch("/api/channels/status", { headers: authHeaders() });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();

        list.innerHTML = "";
        data.channels.forEach(ch => {
            const count = data.counts[ch] || 0;
            const item = document.createElement("div");
            item.className = "channel-item";
            item.innerHTML =
                '<span class="channel-name">' + escapeHtml(ch) + '</span>' +
                '<span class="channel-count ' + (count === 0 ? 'empty' : '') + '">' +
                    (count > 0 ? count + '건' : '미업로드') +
                '</span>' +
                '<button class="btn-del" ' + (count === 0 ? 'disabled' : '') +
                    ' onclick="deleteChannelData(\'' + escapeHtml(ch).replace(/'/g, "\\'") + '\')">' +
                    '삭제</button>';
            list.appendChild(item);
        });

        document.getElementById("manage-total").textContent = "전체 " + data.total + "건";
    } catch (e) {
        list.innerHTML = '<div style="text-align:center;padding:20px;color:#e74c3c;">불러오기 실패</div>';
    }
}

function closeManageModal() {
    document.getElementById("manage-modal").style.display = "none";
}

async function deleteChannelData(channel) {
    if (!confirm("[" + channel + "] 데이터를 삭제하시겠습니까?")) return;
    try {
        const res = await fetch("/api/channels/delete", {
            method: "POST",
            headers: { ...authHeaders(), "Content-Type": "application/json" },
            body: JSON.stringify({ channel: channel }),
        });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();
        alert(data.message);
        showManageModal();
        fetchSummary();
    } catch (e) {
        alert("삭제 실패: " + e.message);
    }
}

// --- Report Generation ---
async function printReport() {
    const btn = document.querySelector('.btn-pdf');
    btn.textContent = '생성 중...';
    btn.disabled = true;
    try {
        const filters = getFilters();
        const params = new URLSearchParams();
        Object.entries(filters).forEach(([k, v]) => {
            if (v && v !== "전체" && v !== "0") params.append(k, v);
        });
        const res = await fetch("/api/summary?" + params.toString(), { headers: authHeaders() });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();
        if (!data.records || data.records.length === 0) {
            alert('리포트를 생성할 데이터가 없습니다.');
            return;
        }
        window.__reportData = data;
        window.open("/static/report.html#" + encodeURIComponent(TOKEN), "_blank");
    } catch (e) {
        alert('리포트 생성 실패: ' + e.message);
    } finally {
        btn.textContent = '리포트 출력';
        btn.disabled = false;
    }
}

// --- Direct Input ---
function showAddRecordModal() {
    editingRecordId = null;
    document.getElementById("ar-modal-title").textContent = "위험요소 직접입력";
    document.getElementById("ar-modal-desc").textContent = "개별 위험요소를 직접 등록합니다.";
    document.getElementById("ar-submit-btn").textContent = "등록";
    document.getElementById("add-record-form").reset();
    document.getElementById("ar-grade-before").textContent = "-";
    document.getElementById("ar-grade-after").textContent = "-";
    document.getElementById("ar-grade-before").className = "calc-result";
    document.getElementById("ar-grade-after").className = "calc-result";
    document.getElementById("ar-image-url").value = "";
    document.getElementById("ar-image-name").textContent = "선택된 파일 없음";
    document.getElementById("ar-image-preview").style.display = "none";
    document.getElementById("ar-image-after-url").value = "";
    document.getElementById("ar-image-after-name").textContent = "선택된 파일 없음";
    document.getElementById("ar-image-after-preview").style.display = "none";
    document.getElementById("add-record-modal").style.display = "flex";
}

function closeAddRecordModal() {
    document.getElementById("add-record-modal").style.display = "none";
    editingRecordId = null;
}

function calcGrade(phase) {
    const lh = parseInt(document.getElementById("ar-lh-" + phase).value) || 0;
    const sv = parseInt(document.getElementById("ar-sv-" + phase).value) || 0;
    const el = document.getElementById("ar-grade-" + phase);
    if (lh > 0 && sv > 0) {
        const risk = lh * sv;
        const grade = risk <= 4 ? "A" : risk <= 8 ? "B" : risk <= 12 ? "C" : "D";
        el.textContent = risk + " (" + grade + "등급)";
        el.className = "calc-result grade-text-" + grade;
    } else {
        el.textContent = "-";
        el.className = "calc-result";
    }
}

async function previewImage(phase) {
    const suffix = phase === "after" ? "-after" : "";
    const input = document.getElementById("ar-image" + suffix);
    const file = input.files[0];
    if (!file) return;
    document.getElementById("ar-image" + suffix + "-name").textContent = file.name;
    const reader = new FileReader();
    reader.onload = function(e) {
        document.getElementById("ar-image" + suffix + "-thumb").src = e.target.result;
        document.getElementById("ar-image" + suffix + "-preview").style.display = "flex";
    };
    reader.readAsDataURL(file);
    const formData = new FormData();
    formData.append("file", file);
    try {
        const res = await fetch("/api/image/upload", {
            method: "POST",
            headers: authHeaders(),
            body: formData,
        });
        if (res.status === 401) { logout(); return; }
        if (!res.ok) { alert("이미지 업로드 실패"); return; }
        const data = await res.json();
        document.getElementById("ar-image" + suffix + "-url").value = data.url;
    } catch (e) {
        alert("이미지 업로드 실패: " + e.message);
    }
}

function removeImage(phase) {
    const suffix = phase === "after" ? "-after" : "";
    document.getElementById("ar-image" + suffix).value = "";
    document.getElementById("ar-image" + suffix + "-url").value = "";
    document.getElementById("ar-image" + suffix + "-name").textContent = "선택된 파일 없음";
    document.getElementById("ar-image" + suffix + "-preview").style.display = "none";
}

async function submitAddRecord(e) {
    e.preventDefault();
    const payload = {
        channel: document.getElementById("ar-channel").value,
        month: document.getElementById("ar-month").value,
        person: document.getElementById("ar-person").value,
        date: document.getElementById("ar-date").value,
        location: document.getElementById("ar-location").value,
        content: document.getElementById("ar-content").value,
        process: document.getElementById("ar-process").value,
        disaster_type: document.getElementById("ar-disaster").value,
        likelihood_before: parseInt(document.getElementById("ar-lh-before").value) || 0,
        severity_before: parseInt(document.getElementById("ar-sv-before").value) || 0,
        improvement_plan: document.getElementById("ar-improvement").value,
        likelihood_after: parseInt(document.getElementById("ar-lh-after").value) || 0,
        severity_after: parseInt(document.getElementById("ar-sv-after").value) || 0,
        completion: document.getElementById("ar-completion").value,
        week: parseInt(document.getElementById("ar-week").value) || 0,
        image: document.getElementById("ar-image-url").value,
        image_after: document.getElementById("ar-image-after-url").value,
    };

    const isEdit = !!editingRecordId;
    const url = isEdit ? "/api/record/update" : "/api/record/add";
    if (isEdit) payload._id = editingRecordId;

    try {
        const res = await fetch(url, {
            method: "POST",
            headers: { ...authHeaders(), "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();
        if (!res.ok) { alert(data.detail || (isEdit ? "수정 실패" : "등록 실패")); return; }
        alert(data.message);
        closeAddRecordModal();
        fetchSummary();
    } catch (e) {
        alert((isEdit ? "수정" : "등록") + " 실패: " + e.message);
    }
}

// --- Edit / Delete Record ---
function editRecord(id) {
    if (!lastSummaryData) return;
    const r = lastSummaryData.records.find(rec => rec._id === id);
    if (!r) { alert("레코드를 찾을 수 없습니다."); return; }

    editingRecordId = id;
    document.getElementById("ar-modal-title").textContent = "위험요소 수정";
    document.getElementById("ar-modal-desc").textContent = "No." + r.no + " 레코드를 수정합니다.";
    document.getElementById("ar-submit-btn").textContent = "수정";

    document.getElementById("ar-channel").value = r.channel || "안전점검";
    document.getElementById("ar-month").value = r.month || "";
    document.getElementById("ar-person").value = r.person || "";
    document.getElementById("ar-date").value = r.date || "";
    document.getElementById("ar-location").value = r.location || "";
    document.getElementById("ar-content").value = r.content_full || "";
    document.getElementById("ar-process").value = r.process || "";
    document.getElementById("ar-disaster").value = r.disaster_type || "";
    document.getElementById("ar-week").value = r.week || "";
    document.getElementById("ar-lh-before").value = r.likelihood_before || "";
    document.getElementById("ar-sv-before").value = r.severity_before || "";
    document.getElementById("ar-improvement").value = r.improvement_plan || "";
    document.getElementById("ar-lh-after").value = r.likelihood_after || "";
    document.getElementById("ar-sv-after").value = r.severity_after || "";
    document.getElementById("ar-completion").value = r.completion || "미완료";

    calcGrade("before");
    calcGrade("after");

    // Image (before)
    if (r.image) {
        document.getElementById("ar-image-url").value = r.image;
        document.getElementById("ar-image-name").textContent = "기존 사진";
        document.getElementById("ar-image-thumb").src = r.image;
        document.getElementById("ar-image-preview").style.display = "flex";
    } else {
        document.getElementById("ar-image-url").value = "";
        document.getElementById("ar-image-name").textContent = "선택된 파일 없음";
        document.getElementById("ar-image-preview").style.display = "none";
    }
    // Image (after)
    if (r.image_after) {
        document.getElementById("ar-image-after-url").value = r.image_after;
        document.getElementById("ar-image-after-name").textContent = "기존 사진";
        document.getElementById("ar-image-after-thumb").src = r.image_after;
        document.getElementById("ar-image-after-preview").style.display = "flex";
    } else {
        document.getElementById("ar-image-after-url").value = "";
        document.getElementById("ar-image-after-name").textContent = "선택된 파일 없음";
        document.getElementById("ar-image-after-preview").style.display = "none";
    }

    document.getElementById("add-record-modal").style.display = "flex";
}

async function deleteRecord(id) {
    if (!confirm("이 위험요소를 삭제하시겠습니까?")) return;
    try {
        const res = await fetch("/api/record/delete", {
            method: "POST",
            headers: { ...authHeaders(), "Content-Type": "application/json" },
            body: JSON.stringify({ _id: id }),
        });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();
        if (!res.ok) { alert(data.detail || "삭제 실패"); return; }
        alert(data.message);
        fetchSummary();
    } catch (e) {
        alert("삭제 실패: " + e.message);
    }
}

// --- Image Viewer ---
function showImageModal(src) {
    let modal = document.getElementById("image-viewer-modal");
    if (!modal) {
        modal = document.createElement("div");
        modal.id = "image-viewer-modal";
        modal.className = "modal-overlay";
        modal.style.cursor = "pointer";
        modal.onclick = function() { modal.style.display = "none"; };
        modal.innerHTML = '<img id="image-viewer-img" class="image-viewer-img" src="">';
        document.body.appendChild(modal);
    }
    document.getElementById("image-viewer-img").src = src;
    modal.style.display = "flex";
}

// --- Init ---
if (TOKEN) {
    fetch("/api/summary", { headers: authHeaders() })
        .then(res => {
            if (res.ok) showDashboard();
            else logout();
        })
        .catch(() => logout());
}
