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

function showDashboard() {
    document.getElementById("login-screen").style.display = "none";
    document.getElementById("dashboard").style.display = "block";
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
    document.getElementById("s-complete").textContent = data.complete;
    document.getElementById("s-pending").textContent = data.incomplete;
    document.getElementById("s-repeat").textContent = data.repeat_total || 0;
    const rate = data.improvement_rate != null ? data.improvement_rate : 0;
    const rateEl = document.getElementById("s-improvement");
    if (rateEl) {
        rateEl.textContent = rate + "%";
        rateEl.className = "stat-value improvement-rate";
        if (rate >= 80) rateEl.classList.add("high");
        else if (rate >= 50) rateEl.classList.add("mid");
        else rateEl.classList.add("low");
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
    const locLabels = Object.keys(data.location_stats);

    if (locationViewMode === "disaster") {
        // Disaster type view
        const locDisaster = data.location_disaster_stats || {};
        const allTypes = new Set();
        Object.values(locDisaster).forEach(obj => Object.keys(obj).forEach(k => allTypes.add(k)));
        const typeList = [...allTypes].sort((a, b) => {
            const totalA = locLabels.reduce((s, l) => s + ((locDisaster[l] || {})[a] || 0), 0);
            const totalB = locLabels.reduce((s, l) => s + ((locDisaster[l] || {})[b] || 0), 0);
            return totalB - totalA;
        });

        const datasets = typeList.map((dt, i) => ({
            label: dt,
            data: locLabels.map(l => (locDisaster[l] || {})[dt] || 0),
            backgroundColor: DISASTER_COLORS[i % DISASTER_COLORS.length],
            borderRadius: 4,
        }));

        chartInstances["chart-location"] = new Chart(
            document.getElementById("chart-location"),
            {
                type: "bar",
                data: { labels: locLabels, datasets },
                options: {
                    responsive: true,
                    plugins: { legend: { display: true, position: "top" } },
                    scales: {
                        x: { stacked: true, grid: { display: false } },
                        y: { stacked: true, beginAtZero: true, ticks: { stepSize: 5 } },
                    },
                },
            }
        );
    } else {
        // Grade view (default)
        const locDatasets = ["A", "B", "C", "D"].map(g => ({
            label: g + "등급",
            data: locLabels.map(l => data.location_stats[l][g] || 0),
            backgroundColor: GRADE_COLORS[g],
            borderRadius: 4,
        }));
        chartInstances["chart-location"] = new Chart(
            document.getElementById("chart-location"),
            {
                type: "bar",
                data: { labels: locLabels, datasets: locDatasets },
                options: {
                    responsive: true,
                    plugins: { legend: { display: true, position: "top" } },
                    scales: {
                        x: { grid: { display: false } },
                        y: { beginAtZero: true, ticks: { stepSize: 5 } },
                    },
                },
            }
        );
    }
}

function updateCharts(data) {
    // 1. Location bar chart
    renderLocationChart(data);

    // 2. Cumulative remaining incomplete by grade (bar + trend lines)
    destroyChart("chart-grade");
    const gradeCumul = data.grade_cumulative || {};
    const cumulMonths = Object.keys(gradeCumul);
    const cumulLabels = cumulMonths.map(m => m + (m.endsWith("월") ? "" : "월"));
    const dData = cumulMonths.map(m => gradeCumul[m].D || 0);
    const cData = cumulMonths.map(m => gradeCumul[m].C || 0);
    const bData = cumulMonths.map(m => gradeCumul[m].B || 0);
    const aData = cumulMonths.map(m => gradeCumul[m].A || 0);

    chartInstances["chart-grade"] = new Chart(
        document.getElementById("chart-grade"),
        {
            type: "bar",
            data: {
                labels: cumulLabels,
                datasets: [
                    // Stacked bars
                    { label: "D등급", data: dData, backgroundColor: GRADE_COLORS.D + "66", borderRadius: 2, stack: "bar", order: 2 },
                    { label: "C등급", data: cData, backgroundColor: GRADE_COLORS.C + "66", borderRadius: 2, stack: "bar", order: 2 },
                    { label: "B등급", data: bData, backgroundColor: GRADE_COLORS.B + "66", borderRadius: 2, stack: "bar", order: 2 },
                    { label: "A등급", data: aData, backgroundColor: GRADE_COLORS.A + "66", borderRadius: 2, stack: "bar", order: 2 },
                    // Trend lines
                    { label: "D추세", data: dData, type: "line", borderColor: GRADE_COLORS.D, borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: GRADE_COLORS.D, tension: 0.3, fill: false, order: 1 },
                    { label: "C추세", data: cData, type: "line", borderColor: GRADE_COLORS.C, borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: GRADE_COLORS.C, tension: 0.3, fill: false, order: 1 },
                    { label: "B추세", data: bData, type: "line", borderColor: GRADE_COLORS.B, borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: GRADE_COLORS.B, tension: 0.3, fill: false, order: 1 },
                    { label: "A추세", data: aData, type: "line", borderColor: GRADE_COLORS.A, borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: GRADE_COLORS.A, tension: 0.3, fill: false, order: 1 },
                ],
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: "top",
                        labels: {
                            filter: function(item) {
                                return !item.text.includes("추세");
                            },
                        },
                    },
                    tooltip: {
                        callbacks: {
                            footer: function(items) {
                                const idx = items[0].dataIndex;
                                const m = cumulMonths[idx];
                                return "잔여 합계: " + gradeCumul[m].total_remaining + "건";
                            },
                        },
                    },
                },
                scales: {
                    x: { stacked: true, grid: { display: false } },
                    y: { stacked: true, beginAtZero: true, ticks: { stepSize: 5 } },
                },
            },
        }
    );

    // 3. Completion donut
    destroyChart("chart-completion");
    chartInstances["chart-completion"] = new Chart(
        document.getElementById("chart-completion"),
        {
            type: "doughnut",
            data: {
                labels: ["완료", "미완료"],
                datasets: [{
                    data: [data.complete, data.incomplete],
                    backgroundColor: ["#27ae60", "#f39c12"],
                    borderWidth: 0,
                }],
            },
            options: {
                responsive: true,
                cutout: "60%",
                plugins: {
                    legend: { position: "top" },
                },
            },
        }
    );

    // 4. Weekly bar chart
    destroyChart("chart-week");
    const weekLabels = Object.keys(data.week_stats);
    const weekData = weekLabels.map(k => data.week_stats[k]);
    chartInstances["chart-week"] = new Chart(
        document.getElementById("chart-week"),
        {
            type: "bar",
            data: {
                labels: weekLabels,
                datasets: [{
                    label: "발굴 건수",
                    data: weekData,
                    backgroundColor: "#e74c3c",
                    borderRadius: 4,
                }],
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { display: false } },
                    y: { beginAtZero: true },
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

    // 7. Channel charts (visible only with multi-channel data)
    destroyChart("chart-channel");
    destroyChart("chart-channel-grade");
    const channelRow = document.getElementById("channel-chart-row");
    const chStats = data.channel_stats || {};
    const chGradeStats = data.channel_grade_stats || {};
    const chKeys = Object.keys(chStats);

    if (chKeys.length > 1) {
        channelRow.style.display = "";
        const chSorted = chKeys.sort((a, b) => chStats[b] - chStats[a]);
        const chColors = chSorted.map((_, i) =>
            ["#e74c3c", "#3498db", "#f39c12", "#27ae60", "#9b59b6", "#1abc9c", "#e67e22", "#34495e"][i % 8]
        );

        chartInstances["chart-channel"] = new Chart(
            document.getElementById("chart-channel"),
            {
                type: "bar",
                data: {
                    labels: chSorted,
                    datasets: [{
                        label: "건수",
                        data: chSorted.map(k => chStats[k]),
                        backgroundColor: chColors,
                        borderRadius: 4,
                    }],
                },
                options: {
                    responsive: true,
                    indexAxis: "y",
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { beginAtZero: true, grid: { display: false } },
                        y: { grid: { display: false } },
                    },
                },
            }
        );

        chartInstances["chart-channel-grade"] = new Chart(
            document.getElementById("chart-channel-grade"),
            {
                type: "bar",
                data: {
                    labels: chSorted,
                    datasets: [
                        { label: "A등급", data: chSorted.map(k => (chGradeStats[k] || {}).A || 0), backgroundColor: GRADE_COLORS.A, borderRadius: 2 },
                        { label: "B등급", data: chSorted.map(k => (chGradeStats[k] || {}).B || 0), backgroundColor: GRADE_COLORS.B, borderRadius: 2 },
                        { label: "C등급", data: chSorted.map(k => (chGradeStats[k] || {}).C || 0), backgroundColor: GRADE_COLORS.C, borderRadius: 2 },
                        { label: "D등급", data: chSorted.map(k => (chGradeStats[k] || {}).D || 0), backgroundColor: GRADE_COLORS.D, borderRadius: 2 },
                    ],
                },
                options: {
                    responsive: true,
                    plugins: { legend: { position: "top" } },
                    scales: {
                        x: { stacked: true, grid: { display: false } },
                        y: { stacked: true, beginAtZero: true },
                    },
                },
            }
        );
        // Channel summary table
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
        channelRow.style.display = "none";
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
