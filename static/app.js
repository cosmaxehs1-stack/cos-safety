let TOKEN = sessionStorage.getItem("token") || "";

// 바깥 클릭 시 info tooltip 닫기
document.addEventListener("click", function(e) {
    if (!e.target.classList.contains("info-btn")) {
        document.querySelectorAll(".info-tooltip.show").forEach(function(t) { t.classList.remove("show"); });
    }
});
let chartInstances = {};
let locationViewMode = "grade";
let lastSummaryData = null;
let editingRecordId = null;
let currentPage = "summary";
let ADMIN_TOKEN = sessionStorage.getItem("admin_token") || "";
let weeklyLiveData = null;
let weeklySavedSnapshots = [];

const ALL_CHANNELS = [
    "정기위험성평가(코스맥스)", "정기위험성평가(협력사)", "수시위험성평가",
    "안전점검", "부서별 위험요소발굴", "근로자 제안", "5S/EHS평가"
];

// ===== Auth =====
async function login() {
    const pw = document.getElementById("password-input").value;
    const errEl = document.getElementById("login-error");
    try {
        const res = await fetch("/api/login", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ password: pw }),
        });
        if (!res.ok) { errEl.textContent = "비밀번호가 올바르지 않습니다."; return; }
        const data = await res.json();
        TOKEN = data.token;
        sessionStorage.setItem("token", TOKEN);
        showDashboard();
    } catch (e) { errEl.textContent = "서버 연결 실패"; }
}

function logout() {
    TOKEN = "";
    ADMIN_TOKEN = "";
    sessionStorage.removeItem("token");
    sessionStorage.removeItem("admin_token");
    document.getElementById("login-screen").style.display = "flex";
    document.getElementById("dashboard").style.display = "none";
}

function showDashboard() {
    document.getElementById("login-screen").style.display = "none";
    document.getElementById("dashboard").style.display = "flex";
    initDateDefaults();
    updateAdminUI();
    fetchSummary();
}

function authHeaders() { return { Authorization: "Bearer " + TOKEN }; }

// ===== Page Navigation =====
function switchPage(pageName, el) {
    currentPage = pageName;
    document.querySelectorAll(".page").forEach(p => { p.style.display = "none"; p.classList.remove("active"); });
    const page = document.getElementById("page-" + pageName);
    if (page) { page.style.display = "block"; page.classList.add("active"); }

    document.querySelectorAll(".nav-item").forEach(n => n.classList.remove("active"));
    if (el) el.classList.add("active");

    // Close mobile sidebar
    document.getElementById("sidebar").classList.remove("open");

    if (pageName === "analysis") {
        // Load charts if needed
        if (lastSummaryData) {
            updateAnalysisCharts(lastSummaryData);
        }
        // Check if weekly subtab is active
        const weeklyTab = document.getElementById("subtab-weekly");
        if (weeklyTab && weeklyTab.style.display !== "none") {
            loadWeeklyTab();
        }
    }
    if (pageName === "register") {
        if (lastSummaryData) updateTable(getDisplayRecords(lastSummaryData));
        showRegisterSelect();
    }
    if (pageName !== "records") {
        // Collapse records sub when navigating away
        document.getElementById("records-sub").classList.remove("open");
        document.querySelectorAll(".nav-sub-item").forEach(function(s) { s.classList.remove("active"); });
    }
}

function toggleSidebar() {
    document.getElementById("sidebar").classList.toggle("open");
}

// ===== Register Methods =====
function showRegisterMethod(method) {
    document.getElementById("register-select").style.display = "none";
    document.getElementById("register-back-btn").style.display = "inline-block";
    document.getElementById("method-direct").style.display = method === "direct" ? "block" : "none";
    document.getElementById("method-excel").style.display = method === "excel" ? "block" : "none";
}

function showRegisterSelect() {
    document.getElementById("register-select").style.display = "flex";
    document.getElementById("register-back-btn").style.display = "none";
    document.getElementById("method-direct").style.display = "none";
    document.getElementById("method-excel").style.display = "none";
}

function switchRegisterMethod(method, btn) {
    showRegisterMethod(method);
}

// ===== Analysis Sub-Tabs =====
function switchAnalysisTab(tabName, btn) {
    document.querySelectorAll(".subtab-content").forEach(el => el.style.display = "none");
    document.querySelectorAll(".analysis-tab").forEach(el => el.classList.remove("active"));
    const el = document.getElementById("subtab-" + tabName);
    if (el) el.style.display = "block";
    if (btn) btn.classList.add("active");

    if (tabName === "weekly") loadWeeklyTab();
    if (lastSummaryData) updateAnalysisCharts(lastSummaryData);
}

// ===== Filters =====
function getFilterValue(id, fallback) {
    const el = document.getElementById(id);
    return el ? el.value : (fallback || "전체");
}

function getFilters() {
    return {
        team: getFilterValue("f-team"),
        channel: getFilterValue("f-channel"),
        year: currentPage === "records" ? getFilterValue("f-rec-year") : getFilterValue("f-year"),
        month: getSelectedMonth(),
        location: getFilterValue("f-location"),
        grade: getFilterValue("f-grade"),
        disaster_type: getFilterValue("f-disaster"),
        process: getFilterValue("f-process"),
        completion: getFilterValue("f-completion"),
        repeat: getFilterValue("f-repeat"),
        keyword: (document.getElementById("f-keyword") || {}).value || "",
    };
}

function setFilterValue(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val;
}

// ===== Month Sheet Tabs =====
function selectMonthTab(tabEl) {
    document.querySelectorAll(".month-tab").forEach(function(t) { t.classList.remove("active"); });
    tabEl.classList.add("active");
    tabEl.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
    // 월 변경 시 주차 초기화
    document.querySelectorAll(".week-card").forEach(function(c) { c.classList.remove("active"); });
    var weekAll = document.querySelector('.week-card[data-week="0"]');
    if (weekAll) weekAll.classList.add("active");
    fetchSummary();
}

function getSelectedMonth() {
    var active = document.querySelector(".month-tab.active");
    return active ? active.getAttribute("data-month") : "전체";
}

function updateMonthTabs(months) {
    var container = document.getElementById("month-sheet-tabs");
    if (!container) return;
    var currentMonth = getSelectedMonth();
    var selectedYear = getFilterValue("f-rec-year");
    var now = new Date();
    var thisYear = String(now.getFullYear());
    var thisMonth = now.getMonth() + 1; // 1~12

    // 선택 연도가 현재 연도면 당월까지만 표시
    var maxMonth = (selectedYear === thisYear) ? thisMonth : 12;

    container.innerHTML = '<div class="month-tab' + (currentMonth === "전체" ? " active" : "") + '" data-month="전체" onclick="selectMonthTab(this)">전체</div>';
    for (var i = 1; i <= maxMonth; i++) {
        var m = i + "월";
        var isActive = (m === currentMonth) ? " active" : "";
        container.innerHTML += '<div class="month-tab' + isActive + '" data-month="' + m + '" onclick="selectMonthTab(this)">' + m + '</div>';
    }

    // 선택된 월이 범위를 넘으면 전체로 리셋
    if (currentMonth !== "전체" && parseInt(currentMonth) > maxMonth) {
        container.querySelector('.month-tab[data-month="전체"]').classList.add("active");
    }
}

function onRecordYearChange() {
    updateMonthTabs();
    fetchSummary();
}

// ===== Week Cards =====
function getSelectedWeek() {
    var active = document.querySelector(".week-card.active");
    return active ? parseInt(active.getAttribute("data-week")) : 0;
}

function selectWeekCard(el) {
    document.querySelectorAll(".week-card").forEach(function(c) { c.classList.remove("active"); });
    el.classList.add("active");
    fetchSummary();
}

function updateWeekCards(records) {
    var container = document.getElementById("week-cards");
    if (!container) return;
    var selectedMonth = getSelectedMonth();
    // 월이 전체이면 주차 카드 숨김
    if (selectedMonth === "전체") { container.innerHTML = ""; return; }

    // 해당 월의 주차 수집
    var weeks = {};
    (records || []).forEach(function(r) {
        if (r.month === selectedMonth) {
            var w = r.week > 0 ? r.week : getWeekFromDate(r.date);
            if (w > 0) weeks[w] = true;
        }
    });
    var weekList = Object.keys(weeks).map(Number).sort(function(a,b){ return a-b; });
    // 최소 1~현재주차까지는 표시
    var now = new Date();
    var curMonthStr = (now.getMonth() + 1) + "월";
    var curYear = String(now.getFullYear());
    var selYear = getFilterValue("f-rec-year");
    var curWeek = 0;
    if ((selYear === curYear || selYear === "전체") && selectedMonth === curMonthStr) {
        curWeek = getWeekFromDate(now.toISOString().split("T")[0]);
        for (var i = 1; i <= curWeek; i++) {
            if (!weeks[i]) weekList.push(i);
        }
        weekList = weekList.filter(function(v, idx, arr) { return arr.indexOf(v) === idx; }).sort(function(a,b){ return a-b; });
    }

    if (weekList.length === 0) { container.innerHTML = ""; return; }

    var prevWeek = getSelectedWeek();
    var html = '<div class="week-card' + (prevWeek === 0 ? ' active' : '') + '" data-week="0" onclick="selectWeekCard(this)">전체</div>';
    weekList.forEach(function(w) {
        var isCurrent = (w === curWeek && selectedMonth === curMonthStr && (selYear === curYear || selYear === "전체"));
        var isActive = (w === prevWeek) ? " active" : "";
        var currentCls = isCurrent ? " current" : "";
        html += '<div class="week-card' + isActive + currentCls + '" data-week="' + w + '" onclick="selectWeekCard(this)">' + w + '주차' + (isCurrent ? ' ★' : '') + '</div>';
    });
    container.innerHTML = html;
}

function resetFilters() {
    setFilterValue("f-channel", "전체");
    setFilterValue("f-rec-year", "전체");
    // Reset month tab to 전체
    document.querySelectorAll(".month-tab").forEach(function(t) { t.classList.remove("active"); });
    var allTab = document.querySelector('.month-tab[data-month="전체"]');
    if (allTab) allTab.classList.add("active");
    setFilterValue("f-location", "전체");
    setFilterValue("f-grade", "전체");
    setFilterValue("f-disaster", "전체");
    setFilterValue("f-process", "전체");
    setFilterValue("f-completion", "전체");
    setFilterValue("f-repeat", "전체");
    setFilterValue("f-keyword", "");
    // 주차 카드 초기화
    document.querySelectorAll(".week-card").forEach(function(c) { c.classList.remove("active"); });
    var weekAll = document.querySelector('.week-card[data-week="0"]');
    if (weekAll) weekAll.classList.add("active");
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
        o.value = opt; o.textContent = opt; sel.appendChild(o);
    });
    if (keepValue && options.includes(prev)) sel.value = prev;
}

// ===== Fetch Data =====
// Records sidebar sub-navigation
function toggleRecordsSub(el) {
    var sub = document.getElementById("records-sub");
    sub.classList.toggle("open");
    // Activate nav item without loading data
    document.querySelectorAll(".nav-item").forEach(function(n) { n.classList.remove("active"); });
    if (el) el.classList.add("active");
}

function openRecordsChannel(channel, el) {
    // Activate sub-item
    document.querySelectorAll(".nav-sub-item").forEach(function(s) { s.classList.remove("active"); });
    if (el) el.classList.add("active");
    // Keep parent nav-item active
    var parent = document.querySelector('.nav-item[data-page="records"]');
    document.querySelectorAll(".nav-item").forEach(function(n) { n.classList.remove("active"); });
    if (parent) parent.classList.add("active");
    // Switch page and set filter
    var page = document.getElementById("page-records");
    document.querySelectorAll(".page").forEach(function(p) { p.style.display = "none"; p.classList.remove("active"); });
    if (page) { page.style.display = "block"; page.classList.add("active"); }
    currentPage = "records";
    document.getElementById("f-channel").value = channel;
    // 기본 필터: 현재 연도
    var currentYear = String(new Date().getFullYear());
    setFilterValue("f-rec-year", currentYear);
    // 기본 필터: 현재 월 탭 선택
    var currentMonth = (new Date().getMonth() + 1) + "월";
    var monthTab = document.querySelector('.month-tab[data-month="' + currentMonth + '"]');
    if (monthTab) {
        document.querySelectorAll(".month-tab").forEach(function(t) { t.classList.remove("active"); });
        monthTab.classList.add("active");
    }
    document.getElementById("records-page-title").textContent =
        channel === "전체" ? "개별 위험요소 확인 - 전체" : "개별 위험요소 확인 - " + channel;
    fetchSummary();
    // Close mobile sidebar
    document.getElementById("sidebar").classList.remove("open");
}

async function fetchSummary() {
    const filters = getFilters();
    const params = new URLSearchParams();
    Object.entries(filters).forEach(([k, v]) => {
        if (v && v !== "전체" && v !== "0") params.append(k, v);
    });
    params.append("page", currentPage || "summary");
    var selW = getSelectedWeek();
    if (selW > 0) params.append("sel_week", selW);

    try {
        const res = await fetch("/api/summary?" + params.toString(), { headers: authHeaders() });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();

        const noDataEl = document.getElementById("no-data");
        if (data.total === 0 && !filters.keyword && filters.month === "전체") {
            noDataEl.style.display = "block";
        } else {
            noDataEl.style.display = "none";
        }

        lastSummaryData = data;
        updateStats(data);
        updatePeriodStats(data);
        updateSummaryCharts(data);
        var displayRecords = getDisplayRecords(data);
        updateTable(displayRecords);
        updateFilters(data.filters);
        updateWeekCards(data.records);
        updateViewSummaryFromRecords(displayRecords, data.view_summary);

        if (currentPage === "analysis") {
            updateAnalysisCharts(data);
        }
    } catch (e) {
        console.error("Fetch error:", e);
    }
}

function getDisplayRecords(data) {
    let records = data.records;
    const repeatFilter = getFilterValue("f-repeat");
    if (repeatFilter === "반복") records = records.filter(r => r.is_repeat);
    else if (repeatFilter === "단건") records = records.filter(r => !r.is_repeat);

    var selWeek = getSelectedWeek();
    if (selWeek > 0) {
        records = records.filter(function(r) {
            var w = r.week > 0 ? r.week : getWeekFromDate(r.date);
            return w === selWeek;
        });
    }
    return records;
}

// ===== Mini Stats + Filter Toggle =====
function toggleFilters() {
    var panel = document.getElementById("filter-panel");
    if (panel.classList.contains("filter-collapsed")) {
        panel.classList.remove("filter-collapsed");
        panel.classList.add("filter-expanded");
    } else {
        panel.classList.remove("filter-expanded");
        panel.classList.add("filter-collapsed");
    }
}

function getTeamFromLocation(loc) {
    if (loc === "평택1공장" || loc === "평택2공장") return "2팀";
    if (loc === "고렴창고") return "2팀"; // 간소화 (1월 예외는 서버에서 처리)
    return "1팀";
}

function updateViewSummaryFromRecords(records, vs) {
    var container = document.getElementById("channel-team-summary");
    if (!container) return;

    var total = records.length;
    if (total === 0) { container.innerHTML = ""; return; }

    var complete = 0, team1 = 0, team1_complete = 0, team2 = 0, team2_complete = 0;

    records.forEach(function(r) {
        var isComplete = r.completion === "완료";
        if (isComplete) complete++;
        var tm = getTeamFromLocation(r.location_group || "");
        if (tm === "1팀") {
            team1++;
            if (isComplete) team1_complete++;
        } else {
            team2++;
            if (isComplete) team2_complete++;
        }
    });
    var incomplete = total - complete;

    var selWeek = getSelectedWeek();
    var label = selWeek > 0 ? selWeek + '주차' : '조회결과';

    container.innerHTML =
        '<div class="ms-card">' +
        '<span class="ms-label">' + label + ' <b class="ms-num">' + total + '</b></span>' +
        '<span class="ms-pair green">개선 <b>' + complete + '</b></span>' +
        '<span class="ms-pair red">미완료 <b>' + incomplete + '</b></span>' +
        '</div>' +
        '<div class="ms-card">' +
        '<span class="ms-label">1팀</span>' +
        '<span class="ms-pair">발굴 <b>' + team1 + '</b></span>' +
        '<span class="ms-pair green">개선 <b>' + team1_complete + '</b></span>' +
        '</div>' +
        '<div class="ms-card">' +
        '<span class="ms-label">2팀</span>' +
        '<span class="ms-pair">발굴 <b>' + team2 + '</b></span>' +
        '<span class="ms-pair green">개선 <b>' + team2_complete + '</b></span>' +
        '</div>';
}

// ===== Period Stats =====
function getWeekFromDate(dateStr) {
    if (!dateStr) return 0;
    const parts = dateStr.split("-");
    if (parts.length < 3) return 0;
    const month = parseInt(parts[1]);
    const day = parseInt(parts[2]);
    if (month === 1) {
        if (day <= 11) return 1;
        return Math.min(Math.floor((day - 12) / 7) + 2, 5);
    }
    // 목요일 기준: 해당 날짜가 속한 주의 목요일로 주차 판단
    const d = new Date(parseInt(parts[0]), month - 1, day);
    const wd = d.getDay(); // 0=일,1=월...4=목,5=금,6=토
    let diffToThu = (4 - wd + 7) % 7;
    if (wd === 5 || wd === 6 || wd === 0) diffToThu -= 7; // 금,토,일 → 이번주 목요일은 과거
    const thu = new Date(d);
    thu.setDate(d.getDate() + diffToThu);
    if (thu.getMonth() !== d.getMonth()) {
        return thu.getMonth() > d.getMonth() ? 5 : 1;
    }
    return Math.min(Math.floor((thu.getDate() - 1) / 7) + 1, 5);
}

function updatePeriodStats(data) {
    const now = new Date();
    const curYear = String(now.getFullYear());
    const curMonth = now.getMonth() + 1;
    const curMonthStr = curMonth + "월";
    const curWeek = getWeekFromDate(now.toISOString().split("T")[0]);

    // 타이틀 설정
    setText("pt-week", curMonthStr + " " + curWeek + "주차");
    setText("pt-month", curMonthStr);
    setText("pt-year", curYear + "년");

    const records = data.records || [];

    // This year
    const yearRecs = records.filter(r => (r.date || "").startsWith(curYear));
    const yearDisc = yearRecs.length;
    const yearImp = yearRecs.filter(r => r.completion === "완료").length;
    const yearRate = yearDisc > 0 ? Math.round(yearImp / yearDisc * 100) : 0;

    // This month
    const monthRecs = yearRecs.filter(r => r.month === curMonthStr);
    const monthDisc = monthRecs.length;
    const monthImp = monthRecs.filter(r => r.completion === "완료").length;
    const monthRate = monthDisc > 0 ? Math.round(monthImp / monthDisc * 100) : 0;

    // This week
    const weekRecs = monthRecs.filter(r => {
        const w = r.week > 0 ? r.week : getWeekFromDate(r.date);
        return w === curWeek;
    });
    const weekDisc = weekRecs.length;
    const weekImp = weekRecs.filter(r => r.completion === "완료").length;
    const weekRate = weekDisc > 0 ? Math.round(weekImp / weekDisc * 100) : 0;

    // 이전 발굴 개선: 이번주 이전에 발굴된 건 중, actual_date가 이번주인 건
    const prevWeekImproved = records.filter(r => {
        if (r.completion !== "완료" || !r.actual_date) return false;
        // 발굴일이 이번주 이전인지
        var rd = r.date || "";
        if (!rd.startsWith(curYear)) return false;
        var rMonth = r.month;
        var rWeek = r.week > 0 ? r.week : getWeekFromDate(rd);
        var isBeforeThisWeek = false;
        var rMonthNum = parseInt(rMonth);
        if (rMonthNum < curMonth) isBeforeThisWeek = true;
        else if (rMonthNum === curMonth && rWeek < curWeek) isBeforeThisWeek = true;
        if (!isBeforeThisWeek) return false;
        // actual_date가 이번주인지
        if (!r.actual_date.startsWith(curYear)) return false;
        var aMonth = parseInt(r.actual_date.split("-")[1]);
        if (aMonth !== curMonth) return false;
        return getWeekFromDate(r.actual_date) === curWeek;
    }).length;

    setText("pw-discovered", weekDisc);
    setText("pw-improved", weekImp);
    setText("pw-rate", weekRate + "%");
    setText("pw-prev-improved", prevWeekImproved);
    var prevWrap = document.getElementById("pw-prev-wrap");
    if (prevWrap) prevWrap.style.display = prevWeekImproved > 0 ? "flex" : "none";
    setText("pm-discovered", monthDisc);
    setText("pm-improved", monthImp);
    setText("pm-rate", monthRate + "%");
    setText("py-discovered", yearDisc);
    setText("py-improved", yearImp);
    setText("py-rate", yearRate + "%");
}

function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
}

// ===== Stats =====
function updateStats(data) {
    setText("s-total", data.total);
    setText("s-a", data.grade_a);
    setText("s-b", data.grade_b);
    setText("s-c", data.grade_c);
    setText("s-d", data.grade_d);
    setText("s-a-after", data.grade_a_current || 0);
    setText("s-b-after", data.grade_b_current || 0);
    setText("s-c-after", data.grade_c_current || 0);
    setText("s-d-after", data.grade_d_current || 0);
    setText("s-complete", data.complete);
    setText("s-pending", data.incomplete);
    setText("s-repeat", data.repeat_total || 0);
    const rate = data.improvement_rate != null ? data.improvement_rate : 0;
    setText("s-improvement", rate + "%");
}

// ===== Charts =====
const GRADE_COLORS = { A: "rgba(39,174,96,0.5)", B: "rgba(52,152,219,0.5)", C: "rgba(243,156,18,0.5)", D: "#e74c3c", "-": "#bdc3c7" };
const DISASTER_COLORS = ["#e74c3c","#3498db","#f39c12","#27ae60","#9b59b6","#1abc9c","#e67e22","#34495e","#e91e63","#00bcd4","#8bc34a","#ff5722","#607d8b","#795548","#cddc39"];

function destroyChart(id) {
    if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
}

function updateSummaryCharts(data) {
    // Grade donuts
    destroyChart("chart-grade-before");
    destroyChart("chart-grade-after");
    const gradeColors = [GRADE_COLORS.A, GRADE_COLORS.B, GRADE_COLORS.C, GRADE_COLORS.D];
    const gradeMiniOpts = {
        responsive: true, cutout: "60%",
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.label + ": " + ctx.parsed + "건" } } }
    };
    chartInstances["chart-grade-before"] = new Chart(document.getElementById("chart-grade-before"), {
        type: "doughnut",
        data: { labels: ["A","B","C","D"], datasets: [{ data: [data.grade_a, data.grade_b, data.grade_c, data.grade_d], backgroundColor: gradeColors, borderWidth: 0 }] },
        options: gradeMiniOpts
    });
    chartInstances["chart-grade-after"] = new Chart(document.getElementById("chart-grade-after"), {
        type: "doughnut",
        data: { labels: ["A","B","C","D"], datasets: [{ data: [data.grade_a_current||0, data.grade_b_current||0, data.grade_c_current||0, data.grade_d_current||0], backgroundColor: gradeColors, borderWidth: 0 }] },
        options: gradeMiniOpts
    });

    // Completion donut
    destroyChart("chart-completion");
    chartInstances["chart-completion"] = new Chart(document.getElementById("chart-completion"), {
        type: "doughnut",
        data: { labels: ["완료","미완료"], datasets: [{ data: [data.complete, data.incomplete], backgroundColor: ["#27ae60","#e5e7eb"], borderWidth: 0 }] },
        options: { responsive: true, cutout: "65%", plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.label + ": " + ctx.parsed + "건" } } } }
    });
}

function updateAnalysisCharts(data) {
    renderLocationChart(data);
    renderWeekChart(data);
    renderChannelTable(data);
    renderDisasterChart(data);
    renderProcessChart(data);
    renderRepeatTable(data);
}

// Location chart
function toggleLocationView() {
    locationViewMode = locationViewMode === "grade" ? "disaster" : "grade";
    const btn = document.getElementById("btn-loc-toggle");
    if (locationViewMode === "grade") { btn.textContent = "재해유형별"; btn.classList.remove("active"); }
    else { btn.textContent = "등급별"; btn.classList.add("active"); }
    if (lastSummaryData) renderLocationChart(lastSummaryData);
}

function renderLocationChart(data) {
    const canvas = document.getElementById("chart-location");
    if (!canvas || canvas.offsetParent === null) return;
    destroyChart("chart-location");

    const MAJOR_ORDER = ["화성","평택","고렴","판교"];
    const hierarchy = data.location_hierarchy || {};
    const subStats = data.location_stats || {};
    const subDisaster = data.location_disaster_stats || {};

    const locLabels = [];
    const majorGroupMap = [];
    MAJOR_ORDER.forEach(major => {
        (hierarchy[major] || []).forEach(sub => { locLabels.push(sub); majorGroupMap.push(major); });
    });
    if (locLabels.length === 0) return;

    let datasets, stacked = false;
    if (locationViewMode === "disaster") {
        stacked = true;
        const allTypes = new Set();
        locLabels.forEach(l => Object.keys(subDisaster[l] || {}).forEach(k => allTypes.add(k)));
        const typeList = [...allTypes].sort((a, b) => {
            const totalA = locLabels.reduce((s, l) => s + ((subDisaster[l]||{})[a]||0), 0);
            const totalB = locLabels.reduce((s, l) => s + ((subDisaster[l]||{})[b]||0), 0);
            return totalB - totalA;
        });
        datasets = typeList.map((dt, i) => ({
            label: dt, data: locLabels.map(l => (subDisaster[l]||{})[dt]||0),
            backgroundColor: DISASTER_COLORS[i % DISASTER_COLORS.length], borderRadius: 4,
        }));
    } else {
        datasets = ["A","B","C","D"].map(g => ({
            label: g + "등급", data: locLabels.map(l => (subStats[l]||{})[g]||0),
            backgroundColor: GRADE_COLORS[g], borderRadius: 4,
        }));
    }

    const displayLabels = locLabels.map((sub, i) => {
        const major = majorGroupMap[i];
        return sub.startsWith(major) ? sub.slice(major.length) || sub : sub;
    });

    const majorGroups = [];
    let groupStart = 0;
    for (let i = 1; i <= majorGroupMap.length; i++) {
        if (i === majorGroupMap.length || majorGroupMap[i] !== majorGroupMap[i-1]) {
            majorGroups.push({ major: majorGroupMap[groupStart], start: groupStart, end: i-1 });
            groupStart = i;
        }
    }

    const BAND_COLORS = ["rgba(59,130,246,0.06)","rgba(16,185,129,0.06)","rgba(245,158,11,0.06)","rgba(139,92,246,0.06)","rgba(107,114,128,0.06)"];
    const groupPlugin = {
        id: "locationGroupBands",
        beforeDraw(chart) {
            const ctx = chart.ctx, xScale = chart.scales.x, chartArea = chart.chartArea;
            ctx.save();
            majorGroups.forEach((g, gi) => {
                const x1 = xScale.getPixelForValue(g.start) - (xScale.getPixelForValue(1)-xScale.getPixelForValue(0))/2;
                const x2 = xScale.getPixelForValue(g.end) + (xScale.getPixelForValue(1)-xScale.getPixelForValue(0))/2;
                ctx.fillStyle = BAND_COLORS[gi % BAND_COLORS.length];
                ctx.fillRect(x1, chartArea.top, x2-x1, chartArea.bottom-chartArea.top);
            });
            ctx.restore();
        },
        afterDraw(chart) {
            const ctx = chart.ctx, xScale = chart.scales.x, chartArea = chart.chartArea;
            ctx.save();
            for (let i = 1; i < majorGroups.length; i++) {
                const x = (xScale.getPixelForValue(majorGroups[i-1].end) + xScale.getPixelForValue(majorGroups[i].start))/2;
                ctx.strokeStyle = "#d1d5db"; ctx.lineWidth = 1; ctx.setLineDash([]);
                ctx.beginPath(); ctx.moveTo(x, chartArea.top); ctx.lineTo(x, chartArea.bottom+20); ctx.stroke();
            }
            const labelY = chartArea.bottom + 38;
            majorGroups.forEach(g => {
                const cx = (xScale.getPixelForValue(g.start) + xScale.getPixelForValue(g.end))/2;
                const bracketY = chartArea.bottom + 24;
                const bx1 = xScale.getPixelForValue(g.start)-4, bx2 = xScale.getPixelForValue(g.end)+4;
                ctx.strokeStyle = "#999"; ctx.lineWidth = 1; ctx.setLineDash([]);
                ctx.beginPath(); ctx.moveTo(bx1,bracketY); ctx.lineTo(bx1,bracketY+4); ctx.lineTo(bx2,bracketY+4); ctx.lineTo(bx2,bracketY); ctx.stroke();
                ctx.fillStyle = "#374151"; ctx.font = "bold 12px sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "top";
                ctx.fillText(g.major, cx, labelY);
            });
            ctx.restore();
        }
    };

    chartInstances["chart-location"] = new Chart(canvas, {
        type: "bar", data: { labels: displayLabels, datasets }, plugins: [groupPlugin],
        options: {
            responsive: true, layout: { padding: { bottom: 30 } },
            plugins: { legend: { display: true, position: "top" }, tooltip: { callbacks: { title: items => locLabels[items[0].dataIndex] } } },
            scales: {
                x: { stacked, grid: { display: false }, ticks: { autoSkip: false, maxRotation: 0, font: { size: 11 } } },
                y: { stacked, beginAtZero: true, ticks: { stepSize: 5 } },
            },
        },
    });
}

function renderWeekChart(data) {
    const canvas = document.getElementById("chart-week");
    if (!canvas || canvas.offsetParent === null) return;
    destroyChart("chart-week");

    const WEEKLY_TARGET = 39;
    const allWeekLabels = Object.keys(data.week_stats);
    const allWeekData = allWeekLabels.map(k => data.week_stats[k]);
    const recent = 13;
    const startIdx = Math.max(0, allWeekLabels.length - recent);
    const weekLabels = allWeekLabels.slice(startIdx);
    const weekData = allWeekData.slice(startIdx);

    chartInstances["chart-week"] = new Chart(canvas, {
        type: "bar",
        data: {
            labels: weekLabels,
            datasets: [
                { label: "주간 발굴", data: weekData, backgroundColor: weekData.map(v => v >= WEEKLY_TARGET ? "#27ae60" : "#3498db"), borderRadius: 4, order: 2 },
                { label: "목표 (39건)", data: weekLabels.map(() => WEEKLY_TARGET), type: "line", borderColor: "#e74c3c", borderWidth: 2, borderDash: [6,4], pointRadius: 0, fill: false, order: 1 },
            ],
        },
        options: {
            responsive: true,
            plugins: { legend: { position: "top", labels: { font: { size: 11 } } } },
            scales: { x: { grid: { display: false } }, y: { beginAtZero: true, title: { display: true, text: "건수" }, suggestedMax: WEEKLY_TARGET + 10 } },
        },
    });
}

function renderChannelTable(data) {
    const tbody = document.getElementById("channel-summary-tbody");
    if (!tbody) return;
    tbody.innerHTML = "";

    const chStats = data.channel_stats || {};
    const chGradeStats = data.channel_grade_stats || {};
    const chKeys = Object.keys(chStats).sort((a, b) => chStats[b] - chStats[a]);

    if (chKeys.length === 0) return;

    let totalRow = { count: 0, A: 0, B: 0, C: 0, D: 0, comp: 0, incomp: 0 };
    chKeys.forEach(ch => {
        const g = chGradeStats[ch] || {};
        const count = chStats[ch] || 0;
        const A = g.A||0, B = g.B||0, C = g.C||0, D = g.D||0;
        const comp = g.complete||0, incomp = g.incomplete||0;
        const chRate = count > 0 ? (comp/count*100).toFixed(1) : 0;
        totalRow.count += count; totalRow.A += A; totalRow.B += B;
        totalRow.C += C; totalRow.D += D; totalRow.comp += comp; totalRow.incomp += incomp;
        const tr = document.createElement("tr");
        tr.innerHTML = "<td>" + escapeHtml(ch) + "</td><td><strong>" + count + "</strong></td>" +
            '<td class="green">' + A + "</td>" + '<td class="blue">' + B + "</td>" +
            '<td class="orange">' + C + "</td>" + '<td class="red">' + D + "</td>" +
            '<td class="status-complete">' + comp + "</td>" + '<td class="status-incomplete">' + incomp + "</td>" +
            '<td style="font-weight:600;color:' + (chRate>=80?'#27ae60':chRate>=50?'#f39c12':'#e74c3c') + '">' + chRate + '%</td>';
        tbody.appendChild(tr);
    });
    const totalRate = totalRow.count > 0 ? (totalRow.comp/totalRow.count*100).toFixed(1) : 0;
    const totalTr = document.createElement("tr");
    totalTr.style.background = "#f0f4ff"; totalTr.style.fontWeight = "700";
    totalTr.innerHTML = "<td>합계</td><td>" + totalRow.count + "</td>" +
        '<td class="green">' + totalRow.A + "</td>" + '<td class="blue">' + totalRow.B + "</td>" +
        '<td class="orange">' + totalRow.C + "</td>" + '<td class="red">' + totalRow.D + "</td>" +
        '<td class="status-complete">' + totalRow.comp + "</td>" + '<td class="status-incomplete">' + totalRow.incomp + "</td>" +
        '<td style="color:' + (totalRate>=80?'#27ae60':totalRate>=50?'#f39c12':'#e74c3c') + '">' + totalRate + '%</td>';
    tbody.appendChild(totalTr);
}

function renderDisasterChart(data) {
    const canvas = document.getElementById("chart-disaster");
    if (!canvas || canvas.offsetParent === null) return;
    destroyChart("chart-disaster");

    const disLabels = Object.keys(data.disaster_stats);
    const disData = disLabels.map(k => data.disaster_stats[k]);
    const disColors = disLabels.map((_, i) => DISASTER_COLORS[i % DISASTER_COLORS.length]);
    chartInstances["chart-disaster"] = new Chart(canvas, {
        type: "doughnut",
        data: { labels: disLabels, datasets: [{ data: disData, backgroundColor: disColors, borderWidth: 0 }] },
        options: { responsive: true, cutout: "50%", plugins: { legend: { position: "right" } } },
    });
}

function renderProcessChart(data) {
    const canvas = document.getElementById("chart-process");
    if (!canvas || canvas.offsetParent === null) return;
    destroyChart("chart-process");

    const procLabels = Object.keys(data.process_stats);
    const procData = procLabels.map(k => data.process_stats[k]);
    const procColors = procLabels.map((_, i) =>
        ["#3498db","#e74c3c","#27ae60","#f39c12","#9b59b6","#1abc9c","#e67e22","#34495e","#e91e63","#00bcd4"][i % 10]
    );
    chartInstances["chart-process"] = new Chart(canvas, {
        type: "bar",
        data: { labels: procLabels, datasets: [{ label: "건수", data: procData, backgroundColor: procColors, borderRadius: 4 }] },
        options: { responsive: true, indexAxis: "y", plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true }, y: { grid: { display: false } } } },
    });
}

function renderRepeatTable(data) {
    const tbody = document.getElementById("repeat-tbody");
    if (!tbody) return;
    tbody.innerHTML = "";
    const records = (data.records || []).filter(r => r.is_repeat);
    records.sort((a, b) => (b.repeat_count || 0) - (a.repeat_count || 0));
    records.forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = "<td>" + r.no + "</td>" +
            '<td title="' + escapeHtml(r.content_full) + '">' + escapeHtml(r.content) + "</td>" +
            '<td><span class="repeat-badge">' + r.repeat_count + "회</span></td>" +
            "<td>" + escapeHtml(r.location || "-") + "</td>" +
            "<td>" + escapeHtml(r.disaster_type || "-") + "</td>" +
            '<td><span class="grade-badge grade-' + r.grade_before + '">' + r.grade_before + "</span></td>" +
            '<td class="' + (r.completion === "완료" ? "status-complete" : "status-incomplete") + '">' + (r.completion || "-") + "</td>";
        tbody.appendChild(tr);
    });
}

// ===== Data Table =====
function updateTable(records) {
    const tbody = document.getElementById("data-tbody");
    if (!tbody) return;
    tbody.innerHTML = "";
    records.forEach(r => {
        const tr = document.createElement("tr");
        const imgBefore = r.has_image ? '<button class="btn-img-load" onclick="loadRecordImage(\'' + escapeHtml(r._id) + '\',\'image\',this)">📷</button>' : '-';
        const imgAfter = r.has_image_after ? '<button class="btn-img-load" onclick="loadRecordImage(\'' + escapeHtml(r._id) + '\',\'image_after\',this)">📷</button>' : '-';
        const rid = escapeHtml(r._id || "");
        tr.innerHTML =
            '<td>' + r.no + '</td><td>' + escapeHtml(r.month) + '</td><td>' + escapeHtml(r.person) + '</td>' +
            '<td>' + (r.date || "-") + '</td><td>' + escapeHtml(r.location || "-") + '</td>' +
            '<td title="' + escapeHtml(r.content_full) + '">' + escapeHtml(r.content) + '</td>' +
            '<td>' + escapeHtml(r.disaster_type || "-") + '</td>' +
            '<td><span class="grade-badge grade-' + r.grade_before + '">' + r.grade_before + '</span></td>' +
            '<td><span class="grade-badge grade-' + (r.grade_after || "-") + '">' + (r.grade_after || "-") + '</span></td>' +
            '<td class="' + (r.completion === "완료" ? "status-complete" : "status-incomplete") + '">' + (r.completion || "-") + '</td>' +
            '<td>' + (r.actual_date || "-") + '</td>' +
            '<td>' + (r.is_repeat ? '<span class="repeat-badge">' + r.repeat_count + '회</span>' : '<span class="repeat-badge single">1회</span>') + '</td>' +
            '<td>' + (r.week || "-") + '</td>' +
            '<td>' + imgBefore + '</td><td>' + imgAfter + '</td>' +
            '<td class="action-cell">' +
                '<button class="btn-edit" onclick="editRecord(\'' + rid + '\')">수정</button>' +
                '<button class="btn-row-del" onclick="deleteRecord(\'' + rid + '\')">삭제</button>' +
            '</td>';
        tbody.appendChild(tr);
    });
}

function escapeHtml(str) {
    if (!str) return "";
    return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ===== Filter Update =====
function updateFilters(filters) {
    if (filters.channels) populateFilter("f-channel", filters.channels, true);
    if (filters.years) {
        populateFilter("f-year", filters.years, true);
        populateFilter("f-rec-year", filters.years, true);
    }
    updateMonthTabs(filters.months);
    populateFilter("f-location", filters.locations, true);
    populateFilter("f-disaster", filters.disaster_types, true);
    populateFilter("f-process", filters.processes, true);
}

// ===== Upload =====
async function uploadFile(input) {
    const file = input.files[0];
    if (!file) return;
    const channel = document.getElementById("upload-channel").value;
    const formData = new FormData();
    formData.append("file", file);
    formData.append("channel", channel);

    try {
        const res = await fetch("/api/upload", { method: "POST", headers: authHeaders(), body: formData });
        if (res.status === 401) { logout(); return; }
        const text = await res.text();
        let data;
        try { data = JSON.parse(text); } catch { data = null; }
        if (!res.ok) { alert("업로드 실패: " + (data?.detail || text || "서버 오류")); return; }
        alert(data?.message || "업로드 완료");
        fetchSummary();
    } catch (e) { alert("업로드 실패: " + e.message); }
    input.value = "";
}

async function downloadExcel() {
    const channel = document.getElementById("upload-channel").value;
    try {
        const res = await fetch("/api/download-excel?channel=" + encodeURIComponent(channel), { headers: authHeaders() });
        if (res.status === 404) { alert("[" + channel + "] 데이터가 없습니다."); return; }
        if (!res.ok) { alert("다운로드 실패"); return; }
        const blob = await res.blob();
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "위험요소_" + channel + ".xlsx";
        a.click();
        URL.revokeObjectURL(a.href);
    } catch (e) { alert("다운로드 실패: " + e.message); }
}

// ===== Direct Input =====
// 작업장 옵션
var WORKPLACE_MAP = {
    "화성1공장": ["1F","2F","3F","4F","옥상","기계실","외부","기타"],
    "화성2공장": ["1F","2F","3F","4F","옥상","기계실","외부","기타"],
    "화성3공장": ["1F","2F","3F","4F","옥상","기계실","외부","기타"],
    "화성5공장": ["1F","2F","3F","4F","옥상","기계실","외부","기타"],
    "평택1공장": ["1F","2F","3F","4F","옥상","기계실","외부","기타"],
    "평택2공장": ["1F","2F","3F","4F","옥상","기계실","외부","기타"],
    "고렴창고": ["1F","2F","외부","기타"],
    "판교사업장": ["사무실","기타"],
    "기타": ["기타"]
};
function updateWorkplaceOptions() {
    var loc = document.getElementById("ar-location").value;
    var wp = document.getElementById("ar-workplace");
    wp.innerHTML = '<option value="">선택하세요</option>';
    var list = WORKPLACE_MAP[loc] || [];
    list.forEach(function(v) {
        var opt = document.createElement("option");
        opt.value = v; opt.textContent = v;
        wp.appendChild(opt);
    });
}

function initDateDefaults() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const dd = String(today.getDate()).padStart(2, "0");
    const dateStr = yyyy + "-" + mm + "-" + dd;
    const monthStr = (today.getMonth() + 1) + "월";
    const weekNum = getWeekFromDate(dateStr);

    document.getElementById("ar-date").value = dateStr;
    document.getElementById("ar-week").value = monthStr + " " + weekNum + "주차";
}

function updateWeekFromDate() {
    const dateVal = document.getElementById("ar-date").value;
    if (dateVal) {
        const weekNum = getWeekFromDate(dateVal);
        const parts = dateVal.split("-");
        const monthLabel = parseInt(parts[1]) + "월";
        document.getElementById("ar-week").value = monthLabel + " " + weekNum + "주차";
    }
}

// Wizard navigation
var currentWizardStep = 1;
function wizardGo(step) {
    // Hide all pages
    for (var i = 1; i <= 3; i++) {
        document.getElementById("wizard-page-" + i).style.display = (i === step) ? "block" : "none";
    }
    // Update step indicators
    document.querySelectorAll(".wizard-step").forEach(function(el) {
        var s = parseInt(el.getAttribute("data-step"));
        el.classList.remove("active", "done");
        if (s === step) el.classList.add("active");
        else if (s < step) el.classList.add("done");
    });
    currentWizardStep = step;
    window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetForm() {
    editingRecordId = null;
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
    document.getElementById("ar-submit-btn").textContent = "등록";
    document.getElementById("ar-cancel-btn").style.display = "none";
    document.getElementById("register-page-title").textContent = "위험요소 등록";
    clearRatingCards();
    wizardGo(1);
    initDateDefaults();
    updateChannelOptions();
}

function cancelEdit() {
    resetForm();
    window.scrollTo({ top: 0, behavior: "smooth" });
}

// Rating card selection
function selectRating(type, phase, value) {
    var containerId = "cards-" + type + "-" + phase;
    var container = document.getElementById(containerId);
    container.querySelectorAll(".rating-card").forEach(function(c) { c.classList.remove("selected"); });
    // Find clicked card by value
    var cards = container.querySelectorAll(".rating-card");
    cards.forEach(function(c) {
        if (c.querySelector(".rating-card-num").textContent.trim() == value) {
            c.classList.add("selected");
        }
    });
    document.getElementById("ar-" + type + "-" + phase).value = value;
    calcGrade(phase);
}

function clearRatingCards() {
    document.querySelectorAll(".rating-card").forEach(function(c) { c.classList.remove("selected"); });
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
    } else { el.textContent = "-"; el.className = "calc-result"; }
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
        const res = await fetch("/api/image/upload", { method: "POST", headers: authHeaders(), body: formData });
        if (res.status === 401) { logout(); return; }
        if (!res.ok) { alert("이미지 업로드 실패"); return; }
        const data = await res.json();
        document.getElementById("ar-image" + suffix + "-url").value = data.url;
    } catch (e) { alert("이미지 업로드 실패: " + e.message); }
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
        month: (function(){ var d = document.getElementById("ar-date").value; return d ? parseInt(d.split("-")[1]) + "월" : ""; })(),
        person: document.getElementById("ar-person").value,
        date: document.getElementById("ar-date").value,
        location: document.getElementById("ar-location").value,
        workplace: document.getElementById("ar-workplace").value,
        content: document.getElementById("ar-content").value,
        cause_object: document.getElementById("ar-cause-object").value,
        process: "",
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
        resetForm();
        fetchSummary();
    } catch (e) { alert((isEdit ? "수정" : "등록") + " 실패: " + e.message); }
}

// ===== Edit / Delete =====
function editRecord(id) {
    if (!lastSummaryData) return;
    const r = lastSummaryData.records.find(rec => rec._id === id);
    if (!r) { alert("레코드를 찾을 수 없습니다."); return; }

    // Navigate to register page
    switchPage("register", document.querySelector('[data-page="register"]'));
    switchRegisterMethod("direct", document.querySelector('.register-tab'));

    editingRecordId = id;
    document.getElementById("register-page-title").textContent = "위험요소 수정 (No." + r.no + ")";
    document.getElementById("ar-submit-btn").textContent = "수정";
    document.getElementById("ar-cancel-btn").style.display = "inline-block";

    // Update channel select to allow all options if editing
    updateChannelOptions(true);

    document.getElementById("ar-channel").value = r.channel || "부서별 위험요소발굴";
    document.getElementById("ar-person").value = r.person || "";
    document.getElementById("ar-date").value = r.date || "";
    document.getElementById("ar-location").value = r.location || "";
    updateWorkplaceOptions();
    document.getElementById("ar-workplace").value = r.workplace || "";
    document.getElementById("ar-content").value = r.content_full || "";
    document.getElementById("ar-cause-object").value = r.cause_object || "";
    document.getElementById("ar-disaster").value = r.disaster_type || "";
    document.getElementById("ar-week").value = r.date ? parseInt(r.date.split("-")[1]) + "월 " + getWeekFromDate(r.date) + "주차" : "";
    document.getElementById("ar-improvement").value = r.improvement_plan || "";
    document.getElementById("ar-completion").value = r.completion || "미완료";

    clearRatingCards();
    if (r.likelihood_before) selectRating("lh", "before", r.likelihood_before);
    if (r.severity_before) selectRating("sv", "before", r.severity_before);
    if (r.likelihood_after) selectRating("lh", "after", r.likelihood_after);
    if (r.severity_after) selectRating("sv", "after", r.severity_after);

    // Lazy load images for edit
    if (r.has_image) {
        fetch("/api/record-image/" + id + "?field=image", { headers: authHeaders() })
            .then(res => res.json()).then(data => {
                if (data.url) {
                    document.getElementById("ar-image-url").value = data.url;
                    document.getElementById("ar-image-name").textContent = "기존 사진";
                    document.getElementById("ar-image-thumb").src = data.url;
                    document.getElementById("ar-image-preview").style.display = "flex";
                }
            });
    }
    if (r.has_image_after) {
        fetch("/api/record-image/" + id + "?field=image_after", { headers: authHeaders() })
            .then(res => res.json()).then(data => {
                if (data.url) {
                    document.getElementById("ar-image-after-url").value = data.url;
                    document.getElementById("ar-image-after-name").textContent = "기존 사진";
                    document.getElementById("ar-image-after-thumb").src = data.url;
                    document.getElementById("ar-image-after-preview").style.display = "flex";
                }
            });
    }

    window.scrollTo({ top: 0, behavior: "smooth" });
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
    } catch (e) { alert("삭제 실패: " + e.message); }
}

// ===== Lazy Image Loading =====
async function loadRecordImage(recordId, field, btn) {
    btn.textContent = "...";
    try {
        const res = await fetch("/api/record-image/" + recordId + "?field=" + field, { headers: authHeaders() });
        const data = await res.json();
        if (data.url) {
            const img = document.createElement("img");
            img.src = data.url;
            img.className = "table-thumb";
            img.onclick = function() { showImageModal(data.url); };
            btn.replaceWith(img);
        } else {
            btn.textContent = "-";
        }
    } catch (e) {
        btn.textContent = "오류";
    }
}

// ===== Image Viewer =====
function showImageModal(src) {
    let modal = document.getElementById("image-viewer-modal");
    if (!modal) {
        modal = document.createElement("div");
        modal.id = "image-viewer-modal"; modal.className = "modal-overlay";
        modal.style.cssText = "cursor:pointer;background:rgba(0,0,0,0.85);";
        modal.onclick = function(e) { if (e.target === modal) modal.style.display = "none"; };
        modal.innerHTML = '<div style="position:relative;display:flex;align-items:center;justify-content:center;">' +
            '<img id="image-viewer-img" class="image-viewer-img" src="">' +
            '<button onclick="document.getElementById(\'image-viewer-modal\').style.display=\'none\'" ' +
            'style="position:absolute;top:-16px;right:-16px;width:36px;height:36px;border-radius:50%;border:none;' +
            'background:rgba(255,255,255,0.9);color:#333;font-size:20px;cursor:pointer;display:flex;align-items:center;' +
            'justify-content:center;box-shadow:0 2px 8px rgba(0,0,0,0.3);">&times;</button></div>';
        document.body.appendChild(modal);
    }
    document.getElementById("image-viewer-img").src = src;
    modal.style.display = "flex";
}

// ===== Admin =====
function adminLogin() {
    return new Promise(resolve => {
        let modal = document.getElementById("admin-login-modal");
        if (!modal) {
            modal = document.createElement("div");
            modal.id = "admin-login-modal"; modal.className = "modal-overlay";
            modal.innerHTML = '<div class="modal-box" style="max-width:360px;">' +
                '<h3 style="margin:0 0 12px;">관리자 인증</h3>' +
                '<input type="password" id="admin-pw-input" placeholder="관리자 비밀번호" ' +
                'style="width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px;box-sizing:border-box;">' +
                '<p id="admin-login-error" style="color:#e74c3c;font-size:13px;margin:8px 0 0;display:none;"></p>' +
                '<div style="display:flex;gap:8px;margin-top:14px;justify-content:flex-end;">' +
                '<button id="admin-cancel-btn" style="padding:8px 16px;border:1px solid #ddd;border-radius:6px;background:#f5f5f5;cursor:pointer;">취소</button>' +
                '<button id="admin-submit-btn" style="padding:8px 16px;border:none;border-radius:6px;background:#1e40af;color:#fff;font-weight:600;cursor:pointer;">확인</button>' +
                '</div></div>';
            document.body.appendChild(modal);
        }
        const input = document.getElementById("admin-pw-input");
        const errEl = document.getElementById("admin-login-error");
        input.value = ""; errEl.style.display = "none";
        modal.style.display = "flex";
        setTimeout(() => input.focus(), 100);

        async function submit() {
            const pw = input.value;
            if (!pw) { errEl.textContent = "비밀번호를 입력하세요."; errEl.style.display = ""; return; }
            try {
                const res = await fetch("/api/admin/login", {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ password: pw }),
                });
                if (!res.ok) { errEl.textContent = "비밀번호가 올바르지 않습니다."; errEl.style.display = ""; return; }
                const data = await res.json();
                ADMIN_TOKEN = data.admin_token;
                sessionStorage.setItem("admin_token", ADMIN_TOKEN);
                cleanup(); resolve(true);
            } catch (e) { errEl.textContent = "서버 연결 실패"; errEl.style.display = ""; }
        }
        function cancel() { cleanup(); resolve(false); }
        function onKey(e) { if (e.key === "Enter") submit(); if (e.key === "Escape") cancel(); }
        function cleanup() {
            modal.style.display = "none";
            document.getElementById("admin-submit-btn").removeEventListener("click", submit);
            document.getElementById("admin-cancel-btn").removeEventListener("click", cancel);
            input.removeEventListener("keydown", onKey);
        }
        document.getElementById("admin-submit-btn").addEventListener("click", submit);
        document.getElementById("admin-cancel-btn").addEventListener("click", cancel);
        input.addEventListener("keydown", onKey);
    });
}

function adminHeaders() { return { "Content-Type": "application/json", "X-Admin-Token": ADMIN_TOKEN }; }

async function toggleAdmin() {
    const btn = document.getElementById("btn-admin");
    if (ADMIN_TOKEN) {
        ADMIN_TOKEN = "";
        sessionStorage.removeItem("admin_token");
        updateAdminUI();
    } else {
        const ok = await adminLogin();
        if (ok) updateAdminUI();
    }
}

function updateAdminUI() {
    const btn = document.getElementById("btn-admin");
    const saveBtn = document.querySelector(".btn-weekly-save");
    const adminOnlyEls = document.querySelectorAll(".admin-only");

    if (ADMIN_TOKEN) {
        btn.textContent = "관리자 ✓";
        btn.classList.add("btn-admin-active");
        if (saveBtn) saveBtn.style.display = "";
        adminOnlyEls.forEach(el => el.style.display = "");
    } else {
        btn.textContent = "관리자";
        btn.classList.remove("btn-admin-active");
        if (saveBtn) saveBtn.style.display = "none";
        adminOnlyEls.forEach(el => el.style.display = "none");
    }
    updateChannelOptions();
}

function updateChannelOptions(forceAll) {
    const isAdmin = !!ADMIN_TOKEN || forceAll;

    // Direct input channel
    const arChannel = document.getElementById("ar-channel");
    const prevArVal = arChannel.value;
    arChannel.innerHTML = "";
    if (isAdmin) {
        ALL_CHANNELS.forEach(ch => {
            const o = document.createElement("option");
            o.value = ch; o.textContent = ch; arChannel.appendChild(o);
        });
        arChannel.value = ALL_CHANNELS.includes(prevArVal) ? prevArVal : "안전점검";
    } else {
        const o = document.createElement("option");
        o.value = "부서별 위험요소발굴"; o.textContent = "부서별 위험요소발굴";
        arChannel.appendChild(o);
    }

    // Upload channel
    const uploadChannel = document.getElementById("upload-channel");
    const prevUpVal = uploadChannel.value;
    uploadChannel.innerHTML = "";
    if (isAdmin) {
        ALL_CHANNELS.forEach(ch => {
            const o = document.createElement("option");
            o.value = ch; o.textContent = ch; uploadChannel.appendChild(o);
        });
        uploadChannel.value = ALL_CHANNELS.includes(prevUpVal) ? prevUpVal : "안전점검";
    } else {
        const o = document.createElement("option");
        o.value = "부서별 위험요소발굴"; o.textContent = "부서별 위험요소발굴";
        uploadChannel.appendChild(o);
    }
}

// ===== Manage Modal =====
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
                '<span class="channel-count ' + (count === 0 ? 'empty' : '') + '">' + (count > 0 ? count + '건' : '미업로드') + '</span>' +
                '<button class="btn-del" ' + (count === 0 ? 'disabled' : '') +
                ' onclick="deleteChannelData(\'' + escapeHtml(ch).replace(/'/g, "\\'") + '\')">삭제</button>';
            list.appendChild(item);
        });
        document.getElementById("manage-total").textContent = "전체 " + data.total + "건";
    } catch (e) {
        list.innerHTML = '<div style="text-align:center;padding:20px;color:#e74c3c;">불러오기 실패</div>';
    }
}

function closeManageModal() { document.getElementById("manage-modal").style.display = "none"; }

async function deleteChannelData(channel) {
    if (!confirm("[" + channel + "] 데이터를 삭제하시겠습니까?")) return;
    try {
        const res = await fetch("/api/channels/delete", {
            method: "POST",
            headers: { ...authHeaders(), "Content-Type": "application/json" },
            body: JSON.stringify({ channel }),
        });
        if (res.status === 401) { logout(); return; }
        const data = await res.json();
        alert(data.message);
        showManageModal();
        fetchSummary();
    } catch (e) { alert("삭제 실패: " + e.message); }
}

// ===== Report =====
async function printReport() {
    const filters = getFilters();
    const params = new URLSearchParams();
    Object.entries(filters).forEach(([k, v]) => {
        if (v && v !== "전체" && v !== "0") params.append(k, v);
    });
    // Open window immediately (user gesture context) to avoid popup blocker
    const reportWin = window.open("about:blank", "_blank");
    try {
        const res = await fetch("/api/summary?" + params.toString(), { headers: authHeaders() });
        if (res.status === 401) { logout(); if (reportWin) reportWin.close(); return; }
        const data = await res.json();
        if (!data.records || data.records.length === 0) {
            alert("리포트를 생성할 데이터가 없습니다.");
            if (reportWin) reportWin.close();
            return;
        }
        // Strip base64 images to avoid sessionStorage quota
        const lightData = JSON.parse(JSON.stringify(data));
        lightData.records = lightData.records.map(r => {
            const copy = { ...r };
            if (copy.image && copy.image.length > 500) copy.image = "";
            if (copy.image_after && copy.image_after.length > 500) copy.image_after = "";
            return copy;
        });
        sessionStorage.setItem("reportData", JSON.stringify(lightData));
        if (reportWin) reportWin.location.href = "/static/report.html";
    } catch (e) {
        alert("리포트 생성 실패: " + e.message);
        if (reportWin) reportWin.close();
    }
}

// ===== Weekly Tab =====
async function loadWeeklyTab() {
    const year = document.getElementById("w-year").value;
    const quarter = document.getElementById("w-quarter").value;

    try {
        const [liveRes, listRes] = await Promise.all([
            fetch("/api/weekly/quarter?year=" + year + "&quarter=" + quarter, { headers: authHeaders() }),
            fetch("/api/weekly/list?year=" + year + "&quarter=" + quarter, { headers: authHeaders() }),
        ]);
        weeklyLiveData = await liveRes.json();
        const listData = await listRes.json();
        weeklySavedSnapshots = listData.snapshots || [];

        const badgeContainer = document.getElementById("w-saved-badges");
        if (weeklySavedSnapshots.length === 0) {
            badgeContainer.innerHTML = '<span style="color:#999;font-size:12px;">저장 이력 없음</span>';
        } else {
            badgeContainer.innerHTML = weeklySavedSnapshots.map(s =>
                '<span class="saved-badge">' + s.year + '년 ' + s.month + '월 ' + s.week + '주차 업데이트 완료 <span style="color:#64748b;font-size:10px;">(' + s.saved_at.slice(0,16).replace("T"," ") + ')</span></span>'
            ).join(" ");
        }

        const compareSelect = document.getElementById("w-compare");
        const compareWrap = document.getElementById("weekly-compare-select");
        if (weeklySavedSnapshots.length > 0) {
            compareWrap.style.display = "";
            let opts = '<option value="">비교 안 함</option>';
            weeklySavedSnapshots.forEach(s => {
                opts += '<option value="' + s.id + '">' + s.month + '월 ' + s.week + '주 (' + s.saved_at.slice(0,10) + ')</option>';
            });
            compareSelect.innerHTML = opts;
            compareSelect.value = weeklySavedSnapshots[weeklySavedSnapshots.length-1].id;
        } else {
            compareWrap.style.display = "none";
            compareSelect.innerHTML = "";
        }
        renderCurrentWeekly();
    } catch (e) { console.error("Weekly load error:", e); }
}

async function renderCurrentWeekly() {
    const compareId = document.getElementById("w-compare").value;
    let prevData = null;
    if (compareId) {
        try {
            const res = await fetch("/api/weekly/get?id=" + compareId, { headers: authHeaders() });
            const result = await res.json();
            if (result.snapshot) prevData = result.snapshot.data;
        } catch (e) {}
    }
    renderQuarterTable(weeklyLiveData, prevData);
}

function renderQuarterTable(data, prevData) {
    const container = document.getElementById("weekly-tables");
    if (!data || !data.sites) { container.innerHTML = ""; return; }

    const months = data.months || [];
    const channels = data.channel_order || [];
    const siteNames = ["전체","환경안전1팀","환경안전2팀"];
    const curMonth = parseInt(document.getElementById("w-cur-month").value);
    const curWeek = parseInt(document.getElementById("w-cur-week").value);

    const totalSite = data.sites["전체"] || {};
    const has5th = {};
    months.forEach(m => {
        has5th[m] = false;
        [...channels, "합계"].forEach(ch => {
            const wk = (totalSite[ch] || {}).weeks || {};
            const w5 = wk[m + "-5"];
            if (w5 && (w5.discovered > 0 || w5.improved > 0)) has5th[m] = true;
        });
    });

    let html = "";
    siteNames.forEach(siteName => {
        const siteData = data.sites[siteName] || {};
        const prevSiteData = prevData ? (prevData.sites || {})[siteName] || {} : null;

        html += '<div class="weekly-table-wrap"><h4 class="weekly-site-title">' + siteName + '</h4>';
        html += '<table class="weekly-table"><thead>';

        html += '<tr><th rowspan="2" class="wt-fixed">구분</th><th rowspan="2" class="wt-fixed"></th>';
        months.forEach(m => {
            const weekCount = has5th[m] ? 5 : 4;
            html += '<th colspan="' + (weekCount+1) + '" class="wt-month-group">' + m + '월</th>';
        });
        html += '<th colspan="1" class="wt-month-group">' + data.year + '년 ' + data.quarter + '분기</th>';
        html += '<th rowspan="2">개선률</th></tr>';

        html += '<tr>';
        months.forEach(m => {
            const maxW = has5th[m] ? 5 : 4;
            for (let w = 1; w <= maxW; w++) {
                const isCur = (m === curMonth && w === curWeek);
                html += '<th class="' + (isCur?"wt-current":"") + ' ' + (w===1?"wt-month-start":"") + '">' + w + '주' + (isCur?" ★":"") + '</th>';
            }
            html += '<th class="wt-sub">소계</th>';
        });
        html += '<th class="wt-month-start">합계</th></tr></thead><tbody>';

        const allCh = [...channels, "합계"];
        allCh.forEach((ch, idx) => {
            const d = siteData[ch] || {};
            const p = prevSiteData ? (prevSiteData[ch] || {}) : null;
            const isTotal = ch === "합계";
            const isLastBeforeTotal = (idx === allCh.length - 2);
            const rowCls = isTotal ? "weekly-total-row" : "";

            html += '<tr class="' + rowCls + ' wt-ch-first">';
            html += '<td class="ch-name" rowspan="2">' + ch + '</td>';
            html += '<td class="row-type">발굴</td>';
            months.forEach(m => {
                const maxW = has5th[m] ? 5 : 4;
                for (let w = 1; w <= maxW; w++) {
                    const wk = d.weeks ? (d.weeks[m+"-"+w] || {}) : {};
                    const pw = p && p.weeks ? (p.weeks[m+"-"+w] || {}) : null;
                    const val = wk.discovered || 0;
                    const pval = pw ? (pw.discovered || 0) : null;
                    const isCur = (m === curMonth && w === curWeek);
                    const diff = pval !== null ? val - pval : null;
                    html += '<td class="num ' + (isCur?"wt-current":"") + ' ' + (w===1?"wt-month-start":"") + '">' + (val||"-") +
                        (diff !== null && diff !== 0 ? '<span class="wt-diff ' + (diff>0?"diff-up":"diff-down") + '">' + (diff>0?"+"+diff:diff) + '</span>' : '') + '</td>';
                }
                const sub = d.month_subs ? (d.month_subs[String(m)] || {}) : {};
                html += '<td class="num wt-sub">' + (sub.discovered||0) + '</td>';
            });
            html += '<td class="num wt-qtr wt-month-start">' + (d.quarter_discovered||0) + '</td>';
            html += '<td class="num rate" rowspan="2">' + (d.quarter_rate ? Math.round(d.quarter_rate*100)+"%" : "-") + '</td>';
            html += '</tr>';

            html += '<tr class="' + rowCls + ' ' + (isLastBeforeTotal?"wt-before-total":"") + '">';
            html += '<td class="row-type">개선</td>';
            months.forEach(m => {
                const maxW = has5th[m] ? 5 : 4;
                for (let w = 1; w <= maxW; w++) {
                    const wk = d.weeks ? (d.weeks[m+"-"+w] || {}) : {};
                    const pw = p && p.weeks ? (p.weeks[m+"-"+w] || {}) : null;
                    const val = wk.improved || 0;
                    const pval = pw ? (pw.improved || 0) : null;
                    const isCur = (m === curMonth && w === curWeek);
                    const diff = pval !== null ? val - pval : null;
                    html += '<td class="num ' + (isCur?"wt-current":"") + ' ' + (w===1?"wt-month-start":"") + '">' + (val||"-") +
                        (diff !== null && diff !== 0 ? '<span class="wt-diff ' + (diff>0?"diff-up":"diff-down") + '">' + (diff>0?"+"+diff:diff) + '</span>' : '') + '</td>';
                }
                const sub = d.month_subs ? (d.month_subs[String(m)] || {}) : {};
                html += '<td class="num wt-sub">' + (sub.improved||0) + '</td>';
            });
            html += '<td class="num wt-qtr wt-month-start">' + (d.quarter_improved||0) + '</td>';
            html += '</tr>';
        });

        html += '</tbody></table></div>';
    });

    container.innerHTML = html;
}

async function saveWeeklySnapshot() {
    const year = document.getElementById("w-year").value;
    const quarter = document.getElementById("w-quarter").value;
    const curMonth = document.getElementById("w-cur-month").value;
    const curWeek = document.getElementById("w-cur-week").value;

    if (!confirm(year + "년 " + quarter + "분기 (이번주: " + curMonth + "월 " + curWeek + "주)를 확정 저장하시겠습니까?")) return;

    try {
        const res = await fetch("/api/weekly/save", {
            method: "POST", headers: adminHeaders(),
            body: JSON.stringify({ year: parseInt(year), quarter: parseInt(quarter), current_month: parseInt(curMonth), current_week: parseInt(curWeek) }),
        });
        if (!res.ok) { const err = await res.json(); alert(err.detail || "저장 실패"); return; }
        const result = await res.json();
        alert(result.message);
        loadWeeklyTab();
    } catch (e) { alert("서버 연결 실패"); }
}

// ===== Init =====
TOKEN = "public";
sessionStorage.setItem("token", TOKEN);
showDashboard();
