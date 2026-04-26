let TOKEN = sessionStorage.getItem("token") || "";

// 로컬 시간대 기준 YYYY-MM-DD 문자열 반환 (toISOString은 UTC라 KST 새벽~오전에 하루 밀림)
function getLocalDateStr(d) {
    d = d || new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
}

// 바깥 클릭 시 info tooltip 닫기
document.addEventListener("click", function(e) {
    if (!e.target.classList.contains("info-btn")) {
        document.querySelectorAll(".info-tooltip.show").forEach(function(t) { t.classList.remove("show"); });
    }
});

// 뷰포트 변경 시 등급 차트 박스 높이 재적용
window.addEventListener("resize", function() {
    var isNarrow = window.innerWidth <= 1024;
    document.querySelectorAll(".grade-chart-box").forEach(function(box) {
        box.style.height = isNarrow ? "110px" : "180px";
    });
});
let chartInstances = {};
let locationViewMode = "grade";
let lastSummaryData = null;
let editingRecordId = null;
let currentPage = "summary";
let ADMIN_TOKEN = sessionStorage.getItem("admin_token") || "";
let weeklyLiveData = null;
let summaryWeeklyData = null;
let weeklyCurrentMonth = null;
let weeklyCurrentWeek = null;

const ALL_CHANNELS = [
    "정기위험성평가(코스맥스)", "정기위험성평가(협력사)", "수시위험성평가",
    "안전점검", "부서별 위험요소발굴", "근로자 제안", "5S/EHS평가"
];

// ===== Auth =====
function logout() {
    ADMIN_TOKEN = "";
    sessionStorage.removeItem("admin_token");
    updateAdminUI();
}

function showDashboard() {
    initDateDefaults();
    updateAdminUI();
    fetchSummary();
    loadSummaryWeekly();
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
    // 현황요약 진입 시 모든 필터 무조건 초기화
    if (pageName === "summary") {
        _activeTeamFilter = "";
        _activeStatusFilter = "";
        _activePrevWeekImproved = false;
        setFilterValue("f-rec-year", "전체");
        setFilterValue("f-rec-month", "전체");
        setFilterValue("f-rec-week", "0");
        setFilterValue("f-channel", "전체");
        setFilterValue("f-location", "전체");
        setFilterValue("f-grade", "전체");
        setFilterValue("f-disaster", "전체");
        setFilterValue("f-process", "전체");
        setFilterValue("f-completion", "전체");
        setFilterValue("f-repeat", "전체");
        const kw = document.getElementById("f-keyword");
        if (kw) kw.value = "";
        fetchSummary();
        loadSummaryWeekly();
    }
}

function toggleSidebar() {
    document.getElementById("sidebar").classList.toggle("open");
}

function goHome() {
    const summaryNav = document.querySelector('.nav-item[data-page="summary"]');
    switchPage('summary', summaryNav);
}

function collapseSidebar() {
    document.getElementById("sidebar").classList.add("collapsed");
    document.getElementById("main-content").classList.add("full-width");
    document.getElementById("sidebar-reopen").style.display = "block";
}

function openSidebar() {
    document.getElementById("sidebar").classList.remove("collapsed");
    document.getElementById("main-content").classList.remove("full-width");
    document.getElementById("sidebar-reopen").style.display = "none";
}

document.addEventListener("click", function(e) {
    const sidebar = document.getElementById("sidebar");
    if (!sidebar || sidebar.classList.contains("collapsed")) return;
    if (sidebar.contains(e.target)) return;
    if (e.target.closest("#sidebar-reopen")) return;
    // 모달이나 팝업 클릭은 무시
    if (e.target.closest(".modal-overlay, .custom-confirm-overlay")) return;
    // 현황요약(홈)에서는 데스크톱에서만 자동으로 접지 않음 (모바일은 화면 공간 확보를 위해 접힘)
    if (currentPage === "summary" && window.innerWidth > 768) return;
    collapseSidebar();
});

// ===== Register Methods =====
function showRegisterMethod(method) {
    document.getElementById("register-select").style.display = "none";
    document.getElementById("register-back-btn").style.display = "inline-block";
    document.getElementById("method-direct").style.display = method === "direct" ? "block" : "none";
    document.getElementById("method-excel").style.display = method === "excel" ? "block" : "none";
    // 직접 입력 진입 시 이전 작업 내용 초기화 (수정 중이 아닐 때만)
    if (method === "direct" && !editingRecordId) {
        resetForm();
    }
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

// ===== Month & Week Dropdowns =====
function getSelectedMonth() {
    return getFilterValue("f-rec-month");
}

function getSelectedWeek() {
    var val = getFilterValue("f-rec-week");
    return parseInt(val) || 0;
}

function updateMonthDropdown() {
    var sel = document.getElementById("f-rec-month");
    if (!sel) return;
    var prev = sel.value;
    var selectedYear = getFilterValue("f-rec-year");
    var now = new Date();
    var thisYear = String(now.getFullYear());
    var thisMonth = now.getMonth() + 1;
    var maxMonth = (selectedYear === thisYear) ? thisMonth : 12;

    sel.innerHTML = '<option value="전체">전체</option>';
    for (var i = 1; i <= maxMonth; i++) {
        var m = i + "월";
        sel.innerHTML += '<option value="' + m + '"' + (m === prev ? ' selected' : '') + '>' + m + '</option>';
    }
    if (prev !== "전체" && parseInt(prev) > maxMonth) sel.value = "전체";
}

function updateWeekDropdown(records) {
    var sel = document.getElementById("f-rec-week");
    if (!sel) return;
    var prev = parseInt(sel.value) || 0;
    var selectedMonth = getSelectedMonth();

    if (selectedMonth === "전체") {
        sel.innerHTML = '<option value="0">전체</option>';
        return;
    }

    var weeks = {};
    (records || []).forEach(function(r) {
        if (r.month === selectedMonth) {
            var w = r.week > 0 ? r.week : getWeekFromDate(r.date);
            if (w > 0) weeks[w] = true;
        }
    });
    var weekList = Object.keys(weeks).map(Number).sort(function(a,b){ return a-b; });

    var now = new Date();
    var curMonthStr = (now.getMonth() + 1) + "월";
    var curYear = String(now.getFullYear());
    var selYear = getFilterValue("f-rec-year");
    if ((selYear === curYear || selYear === "전체") && selectedMonth === curMonthStr) {
        var curWeek = getWeekFromDate(getLocalDateStr(now));
        for (var i = 1; i <= curWeek; i++) {
            if (!weeks[i]) weekList.push(i);
        }
        weekList = weekList.filter(function(v, idx, arr) { return arr.indexOf(v) === idx; }).sort(function(a,b){ return a-b; });
    }

    sel.innerHTML = '<option value="0">전체</option>';
    weekList.forEach(function(w) {
        sel.innerHTML += '<option value="' + w + '"' + (w === prev ? ' selected' : '') + '>' + w + '주차</option>';
    });
}

function onRecordYearChange() {
    if (_activePrevWeekImproved) _activePrevWeekImproved = false;
    updateMonthDropdown();
    setFilterValue("f-rec-week", "0");
    fetchSummary();
}

function onRecordMonthChange() {
    if (_activePrevWeekImproved) _activePrevWeekImproved = false;
    setFilterValue("f-rec-week", "0");
    fetchSummary();
}

function onRecordWeekChange() {
    if (_activePrevWeekImproved) _activePrevWeekImproved = false;
    fetchSummary();
}

function resetFilters() {
    _activeTeamFilter = "";
    _activeStatusFilter = "";
    _activePrevWeekImproved = false;
    setFilterValue("f-channel", "전체");
    setFilterValue("f-rec-year", "전체");
    setFilterValue("f-rec-month", "전체");
    setFilterValue("f-rec-week", "0");
    setFilterValue("f-location", "전체");
    setFilterValue("f-grade", "전체");
    setFilterValue("f-disaster", "전체");
    setFilterValue("f-process", "전체");
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
    // pending filter가 없을 때만 추가개선 필터 클리어 (있으면 applyPendingRecordsFilter가 다시 set)
    if (!window._pendingRecordsFilter) {
        _activePrevWeekImproved = false;
    }
    document.getElementById("f-channel").value = channel;
    // 기본 필터: 현재 연도
    var currentYear = String(new Date().getFullYear());
    setFilterValue("f-rec-year", currentYear);
    // 기본 필터: 현재 월 선택
    var currentMonth = (new Date().getMonth() + 1) + "월";
    updateMonthDropdown();
    setFilterValue("f-rec-month", currentMonth);
    setFilterValue("f-rec-week", "0");
    document.getElementById("records-page-title").textContent =
        channel === "전체" ? "개별 위험요소 관리 - 전체" : "개별 위험요소 관리 - " + channel;
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
        updateWeekDropdown(data.records);
        // 카드는 팀/상태 필터 없는 레코드 기준
        var savedTeam = _activeTeamFilter;
        var savedStatus = _activeStatusFilter;
        _activeTeamFilter = "";
        _activeStatusFilter = "";
        var allRecords = getDisplayRecords(data);
        _activeTeamFilter = savedTeam;
        _activeStatusFilter = savedStatus;
        updateViewSummaryFromRecords(allRecords, data.view_summary);
        renderChipBar();

        if (currentPage === "analysis") {
            updateAnalysisCharts(data);
        }

        // Apply pending filter (from "발굴/개선" click on summary)
        if (window._pendingRecordsFilter) {
            applyPendingRecordsFilter();
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

    if (_activePrevWeekImproved) {
        // 이전주차 개선 모드: 발굴 주차 필터 무시하고 현재 실제 주차 기준으로 필터
        // (올해 이전 주차 발굴 + 이번 주차 완료)
        var now = new Date();
        var refYearStr = String(now.getFullYear());
        var refMonthNum = now.getMonth() + 1;
        var refWeek = getWeekFromDate(getLocalDateStr(now));
        records = records.filter(function(r) {
            if (r.completion !== "완료" || !r.actual_date) return false;
            // 발굴일이 올해 + 이번 주차 이전인지
            var rd = r.date || "";
            if (!rd.startsWith(refYearStr)) return false;
            var rWeek = r.week > 0 ? r.week : getWeekFromDate(rd);
            var rMonthNum = parseInt(r.month);
            var isBefore = false;
            if (rMonthNum < refMonthNum) isBefore = true;
            else if (rMonthNum === refMonthNum && rWeek < refWeek) isBefore = true;
            if (!isBefore) return false;
            // actual_date(개선완료일)가 이번 주차에 속하는지
            if (!r.actual_date.startsWith(refYearStr)) return false;
            var aMonth = parseInt(r.actual_date.split("-")[1]);
            if (aMonth !== refMonthNum) return false;
            return getWeekFromDate(r.actual_date) === refWeek;
        });
    } else {
        // 일반 모드: 선택 주차 필터 적용
        var selWeek = getSelectedWeek();
        if (selWeek > 0) {
            records = records.filter(function(r) {
                var w = r.week > 0 ? r.week : getWeekFromDate(r.date);
                return w === selWeek;
            });
        }
    }

    var teamFilter = getActiveTeamFilter();
    if (teamFilter) {
        records = records.filter(function(r) {
            return getTeamFromLocation(r.location_group || "") === teamFilter;
        });
    }
    var statusFilter = getActiveStatusFilter();
    if (statusFilter === "개선") {
        records = records.filter(function(r) { return r.completion === "완료"; });
    } else if (statusFilter === "미개선") {
        records = records.filter(function(r) { return r.completion !== "완료"; });
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

var _activeTeamFilter = "";  // "", "1팀", "2팀"
var _activeStatusFilter = ""; // "", "발굴"(전체), "개선"(완료만)
var _activePrevWeekImproved = false; // 이전 주차 발굴 → 선택 주차 완료 필터 (선택 주차 기준 동적)

function getActiveTeamFilter() { return _activeTeamFilter; }
function getActiveStatusFilter() { return _activeStatusFilter; }

function toggleTeamFilter(team, status) {
    if (_activeTeamFilter === team && _activeStatusFilter === (status || "")) {
        _activeTeamFilter = "";
        _activeStatusFilter = "";
    } else {
        _activeTeamFilter = team;
        _activeStatusFilter = status || "";
    }
    if (lastSummaryData) {
        var displayRecords = getDisplayRecords(lastSummaryData);
        updateTable(displayRecords);
        var savedTeam = _activeTeamFilter;
        var savedStatus = _activeStatusFilter;
        _activeTeamFilter = "";
        _activeStatusFilter = "";
        var allRecords = getDisplayRecords(lastSummaryData);
        _activeTeamFilter = savedTeam;
        _activeStatusFilter = savedStatus;
        updateViewSummaryFromRecords(allRecords, lastSummaryData.view_summary);
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

    // 이전주차 개선 필터 활성 시: 강조 인디케이터 + 해제 링크
    if (_activePrevWeekImproved && currentPage === "records") {
        var pwNow = new Date();
        var pwMonth = pwNow.getMonth() + 1;
        var pwWeek = getWeekFromDate(getLocalDateStr(pwNow));
        container.innerHTML =
            '<div class="prev-week-active">' +
            '<span>' + pwMonth + '월 ' + pwWeek + '주차 추가 개선' +
            '<span class="info-btn" onclick="event.stopPropagation(); this.nextElementSibling.classList.toggle(\'show\')">?</span>' +
            '<span class="info-tooltip">' + pwMonth + '월 ' + pwWeek + '주차 이전에 발굴된 위험요소 중, 이번주에 개선완료 처리된 건수입니다.</span>' +
            '</span>' +
            '<span class="prev-week-active-count">' + total + '건</span>' +
            '<span class="prev-week-deactivate" onclick="togglePrevWeekImproved()">뒤로</span>' +
            '</div>';
        return;
    }

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
    var team1_incomplete = team1 - team1_complete;
    var team2_incomplete = team2 - team2_complete;

    var selWeek = getSelectedWeek();

    var at = getActiveTeamFilter();
    var as = getActiveStatusFilter();

    function cell(team, status, label, count, color) {
        var isActive = (at === team && as === status);
        var cls = 'ms-cell' + (color ? ' ' + color : '') + (isActive ? ' ms-cell-active' : '');
        return '<div class="' + cls + '" onclick="toggleTeamFilter(\'' + team + '\', \'' + status + '\')"><div class="ms-cell-label">' + label + '</div><div class="ms-cell-num">' + count + '</div></div>';
    }

    var label = selWeek > 0 ? '전체 (' + selWeek + '주차)' : '전체';

    container.innerHTML =
        '<div class="ms-grid">' +
        '<div class="ms-grid-label">' + label + '</div>' +
        cell('', '', '발굴', total, '') +
        cell('', '미개선', '미개선', incomplete, 'orange') +
        cell('', '개선', '개선', complete, 'green') +
        '</div>' +
        '<div class="ms-grid">' +
        '<div class="ms-grid-label">1팀</div>' +
        cell('1팀', '', '발굴', team1, '') +
        cell('1팀', '미개선', '미개선', team1_incomplete, 'orange') +
        cell('1팀', '개선', '개선', team1_complete, 'green') +
        '</div>' +
        '<div class="ms-grid">' +
        '<div class="ms-grid-label">2팀</div>' +
        cell('2팀', '', '발굴', team2, '') +
        cell('2팀', '미개선', '미개선', team2_incomplete, 'orange') +
        cell('2팀', '개선', '개선', team2_complete, 'green') +
        '</div>' +
        // 트리거 텍스트 링크: 이전주차 발굴 → 이번주 개선건 보기
        '<span class="prev-week-link" onclick="togglePrevWeekImproved()" title="이전 주차에 발굴된 위험요소 중 이번주에 개선완료된 건만 보기">+ 이번주 추가 개선건 보기 (이전 주차 발굴분)</span>';
}

// ===== Chip Bar (통합 필터) =====
function renderChipBar() {
    function renderChips(rowId, selectId, triggerFn) {
        const sel = document.getElementById(selectId);
        const row = document.getElementById(rowId);
        if (!sel || !row) return;
        const cur = sel.value;
        row.innerHTML = "";
        Array.from(sel.options).forEach(opt => {
            const chip = document.createElement("button");
            chip.type = "button";
            chip.className = "chip" + (String(opt.value) === String(cur) ? " active" : "");
            chip.textContent = opt.textContent;
            chip.onclick = function() {
                sel.value = opt.value;
                if (triggerFn) triggerFn();
                renderChipBar();
            };
            row.appendChild(chip);
        });
    }
    renderChips("chip-row-year", "f-rec-year", onRecordYearChange);
    renderChips("chip-row-month", "f-rec-month", onRecordMonthChange);
    renderChips("chip-row-week", "f-rec-week", onRecordWeekChange);

    renderTeamStatusChips();
}

function renderTeamStatusChips() {
    const teamRow = document.getElementById("chip-row-team");
    const statusRow = document.getElementById("chip-row-status");
    if (!teamRow || !statusRow) return;

    const at = getActiveTeamFilter();
    const as = getActiveStatusFilter();

    // 카운트 계산: 시간 필터는 적용, 팀/상태 필터는 각 칩별로 별도 적용
    let baseRecords = [];
    if (lastSummaryData) {
        const savedT = _activeTeamFilter, savedS = _activeStatusFilter;
        _activeTeamFilter = ""; _activeStatusFilter = "";
        baseRecords = getDisplayRecords(lastSummaryData);
        _activeTeamFilter = savedT; _activeStatusFilter = savedS;
    }

    function countFor(team, status) {
        return baseRecords.filter(r => {
            if (team) {
                const tm = getTeamFromLocation(r.location_group || "");
                if (tm !== team) return false;
            }
            if (status === "개선" && r.completion !== "완료") return false;
            if (status === "미개선" && r.completion === "완료") return false;
            return true;
        }).length;
    }

    const teams = [["", "전체"], ["1팀", "1팀"], ["2팀", "2팀"]];
    teamRow.innerHTML = "";
    teams.forEach(([val, label]) => {
        const n = countFor(val, as);
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "chip chip-count" + (at === val ? " active" : "");
        chip.innerHTML = label + '<span class="chip-badge">' + n + '</span>';
        chip.onclick = function() {
            if (at === val) return;
            _activeTeamFilter = val;
            applyCardFilter();
        };
        teamRow.appendChild(chip);
    });

    const statuses = [["", "전체"], ["미개선", "미개선"], ["개선", "개선"]];
    statusRow.innerHTML = "";
    statuses.forEach(([val, label]) => {
        const n = countFor(at, val);
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "chip chip-count chip-status-" + (val || "all") +
            (as === val ? " active" : "");
        chip.innerHTML = label + '<span class="chip-badge">' + n + '</span>';
        chip.onclick = function() {
            if (as === val) return;
            _activeStatusFilter = val;
            applyCardFilter();
        };
        statusRow.appendChild(chip);
    });
}

function applyCardFilter() {
    if (!lastSummaryData) { renderChipBar(); return; }
    const displayRecords = getDisplayRecords(lastSummaryData);
    updateTable(displayRecords);
    const savedTeam = _activeTeamFilter;
    const savedStatus = _activeStatusFilter;
    _activeTeamFilter = "";
    _activeStatusFilter = "";
    const allRecords = getDisplayRecords(lastSummaryData);
    _activeTeamFilter = savedTeam;
    _activeStatusFilter = savedStatus;
    updateViewSummaryFromRecords(allRecords, lastSummaryData.view_summary);
    renderChipBar();
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
    const curWeek = getWeekFromDate(getLocalDateStr(now));

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

    // 이전 발굴 개선: 올해(현재 년도) 이전 주차에 발굴된 건 중, actual_date가 이번주인 건
    const prevWeekImproved = records.filter(r => {
        if (r.completion !== "완료" || !r.actual_date) return false;
        // 발굴일이 올해 + 이번주 이전인지
        var rd = r.date || "";
        if (!rd.startsWith(curYear)) return false;
        var rWeek = r.week > 0 ? r.week : getWeekFromDate(rd);
        var rMonthNum = parseInt(r.month);
        var isBeforeThisWeek = false;
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
    setText("pw-prev-improved", prevWeekImproved);

    var addWrap = document.getElementById("pw-add-wrap");
    var eqWrap = document.getElementById("pw-eq-wrap");
    var opPlus = document.getElementById("pw-op-plus");
    var opEq = document.getElementById("pw-op-eq");
    var noteEl = document.getElementById("pw-note");
    var improvedItem = document.querySelector('.period-item.clickable[onclick*="improved"]');

    if (prevWeekImproved > 0) {
        var actualTotal = weekImp + prevWeekImproved;
        var actualRate = weekDisc > 0 ? Math.round(actualTotal / weekDisc * 100) : 0;
        setText("pw-actual-total", actualTotal);
        setText("pw-actual-rate", actualRate + "%");
        if (addWrap) addWrap.style.display = "";
        if (eqWrap) eqWrap.style.display = "";
        if (opPlus) opPlus.style.display = "";
        if (opEq) opEq.style.display = "";
        if (noteEl) noteEl.style.display = "";
        if (improvedItem) improvedItem.classList.add("period-item-compact");
    } else {
        setText("pw-actual-rate", weekRate + "%");
        if (addWrap) addWrap.style.display = "none";
        if (eqWrap) eqWrap.style.display = "none";
        if (opPlus) opPlus.style.display = "none";
        if (improvedItem) improvedItem.classList.remove("period-item-compact");
        if (opEq) opEq.style.display = "none";
        if (noteEl) noteEl.style.display = "none";
    }
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
    // Grade horizontal bar charts: 개선 전 (left) / 현황 (right)
    destroyChart("chart-grade-before");
    destroyChart("chart-grade-after");
    var gradeBarColors = ["rgba(212,206,193,0.7)", "rgba(193,198,199,0.7)", "rgba(200,100,80,0.7)", "rgba(234,29,34,0.7)"];
    var gradeBarBorders = ["#D4CEC1", "#C1C6C7", "#c0432b", "#EA1D22"];
    var gradeLabels = ["A등급", "B등급", "C등급", "D등급"];
    var maxVal = Math.max(data.grade_a, data.grade_b, data.grade_c, data.grade_d, data.grade_a_current||0, data.grade_b_current||0, data.grade_c_current||0, data.grade_d_current||0, 1);
    var gradeBarOpts = function(showY) {
        return {
            indexAxis: "y", responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) { return ctx.parsed.x + "건"; } } } },
            scales: {
                x: { beginAtZero: true, max: Math.max(50, Math.ceil((maxVal * 1.1) / 50) * 50), ticks: { stepSize: 50, font: { size: 10 } }, grid: { color: "#f0f0f0" } },
                y: { ticks: { display: showY, font: { size: 12, weight: "600" } }, grid: { display: false } }
            }
        };
    };
    // 뷰포트 폭에 따라 차트 박스 높이 강제 설정 (CSS 캐시/우선순위 이슈 우회)
    var isNarrow = window.innerWidth <= 1024;
    document.querySelectorAll(".grade-chart-box").forEach(function(box) {
        box.style.height = isNarrow ? "110px" : "180px";
    });

    chartInstances["chart-grade-before"] = new Chart(document.getElementById("chart-grade-before"), {
        type: "bar",
        data: { labels: gradeLabels, datasets: [{ data: [data.grade_a, data.grade_b, data.grade_c, data.grade_d], backgroundColor: gradeBarColors, borderColor: gradeBarBorders, borderWidth: 1, borderRadius: 2 }] },
        options: gradeBarOpts(true)
    });
    chartInstances["chart-grade-after"] = new Chart(document.getElementById("chart-grade-after"), {
        type: "bar",
        data: { labels: gradeLabels, datasets: [{ data: [data.grade_a_current||0, data.grade_b_current||0, data.grade_c_current||0, data.grade_d_current||0], backgroundColor: gradeBarColors, borderColor: gradeBarBorders, borderWidth: 1, borderRadius: 2 }] },
        options: gradeBarOpts(false)
    });
    // 렌더 후 한번 더 리사이즈 (일부 브라우저에서 초기 크기 캐싱 대응)
    setTimeout(function() {
        if (chartInstances["chart-grade-before"]) chartInstances["chart-grade-before"].resize();
        if (chartInstances["chart-grade-after"]) chartInstances["chart-grade-after"].resize();
    }, 50);

    // Completion donut
    destroyChart("chart-completion");
    chartInstances["chart-completion"] = new Chart(document.getElementById("chart-completion"), {
        type: "doughnut",
        data: { labels: ["완료","미완료"], datasets: [{ data: [data.complete, data.incomplete], backgroundColor: ["#555","#e5e7eb"], borderWidth: 0 }] },
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
let _allTableRecords = [];

function formatShortDate(dateStr) {
    if (!dateStr) return "-";
    const m = String(dateStr).match(/^(\d{2})(\d{2})-(\d{2})-(\d{2})/);
    if (!m) return dateStr;
    return m[2] + "/" + m[3] + "/" + m[4];
}

function updateTable(records) {
    _allTableRecords = records;
    const tbody = document.getElementById("data-tbody");
    if (!tbody) return;
    tbody.innerHTML = "";
    records.forEach((r, i) => {
        const tr = document.createElement("tr");
        tr.style.cursor = "pointer";
        tr.onclick = function(e) {
            if (e.target.closest("button") || e.target.closest("img")) return;
            showRecordDetail(r);
        };
        const rowNo = i + 1;
        const imgBefore = r.image
            ? '<img src="' + escapeHtml(r.image) + '" class="table-thumb" onclick="event.stopPropagation();showImageModal(\'' + escapeHtml(r.image) + '\')">'
            : (r.has_image
                ? '<button class="btn-img-load" onclick="event.stopPropagation();loadRecordImage(\'' + escapeHtml(r._id) + '\',\'image\',this)">📷</button>'
                : '<span class="table-thumb-empty">-</span>');
        const imgAfter = r.image_after
            ? '<img src="' + escapeHtml(r.image_after) + '" class="table-thumb" onclick="event.stopPropagation();showImageModal(\'' + escapeHtml(r.image_after) + '\')">'
            : (r.has_image_after
                ? '<button class="btn-img-load" onclick="event.stopPropagation();loadRecordImage(\'' + escapeHtml(r._id) + '\',\'image_after\',this)">📷</button>'
                : '<span class="table-thumb-empty">-</span>');
        const rid = escapeHtml(r._id || "");
        const contentFull = r.content_full || r.content || "";
        const planFull = r.improvement_plan || "";
        tr.innerHTML =
            '<td class="td-no">' + rowNo + '</td>' +
            '<td>' + formatShortDate(r.date) + '</td><td>' + escapeHtml(r.location || "-") + '</td>' +
            '<td>' + escapeHtml(r.disaster_type || "-") + '</td>' +
            '<td class="td-content td-content-risk" title="' + escapeHtml(contentFull) + '">' + escapeHtml(contentFull) + '</td>' +
            '<td class="td-content td-content-plan" title="' + escapeHtml(planFull) + '">' + escapeHtml(planFull || "-") + '</td>' +
            '<td class="td-status ' + (r.completion === "완료" ? "status-complete" : "status-incomplete") + '">' + (r.completion || "-") + '</td>' +
            '<td class="td-img-pair"><div class="img-pair">' +
                '<div class="img-col">' + imgBefore + '<span class="grade-badge grade-' + r.grade_before + '">' + r.grade_before + '</span></div>' +
                '<span class="grade-arrow">→</span>' +
                '<div class="img-col">' + imgAfter + '<span class="grade-badge grade-' + (r.grade_after || "-") + '">' + (r.grade_after || "-") + '</span></div>' +
            '</div></td>' +
            (ADMIN_TOKEN
                ? '<td class="action-cell">' +
                    '<button class="btn-icon btn-edit" title="수정" onclick="event.stopPropagation();editRecord(\'' + rid + '\')">✏️</button>' +
                    '<button class="btn-icon btn-row-del" title="삭제" onclick="event.stopPropagation();deleteRecord(\'' + rid + '\')">🗑️</button>' +
                  '</td>'
                : '');
        tbody.appendChild(tr);
    });
}

function goToRecordsFiltered(period, type) {
    const now = new Date();
    const curYear = String(now.getFullYear());
    const curMonth = (now.getMonth() + 1) + "월";
    const curWeek = getWeekFromDate(getLocalDateStr(now));

    // Stash desired filters to apply after page switch + data load
    window._pendingRecordsFilter = {
        year: curYear,
        month: (period === "month" || period === "week") ? curMonth : "전체",
        week: period === "week" ? curWeek : 0,
        statusFilter: type === "improved" ? "개선" : "",
        prevWeekImproved: false,
    };

    // Switch to records page (전체 channel)
    const allLink = document.querySelector('#records-sub .nav-sub-item');
    if (allLink) {
        openRecordsChannel('전체', allLink);
    } else {
        switchPage('records', document.querySelector('[data-page="records"]'));
    }
}

// "+ N건 추가 개선" 클릭 핸들러: 이전 주차 발굴 + 이번주 완료된 건들로 이동
function onPrevImprovedClick(e) {
    if (e && (e.target.closest('.info-btn') || e.target.closest('.info-tooltip'))) return;
    var prevImproved = parseInt((document.getElementById("pw-prev-improved") || {}).textContent || "0");
    if (prevImproved <= 0) return;
    goToRecordsImprovedInWeek();
}

function goToRecordsImprovedInWeek() {
    window._pendingRecordsFilter = {
        year: "전체",
        month: "전체",
        week: 0,
        statusFilter: "",
        prevWeekImproved: true,
    };

    const allLink = document.querySelector('#records-sub .nav-sub-item');
    if (allLink) {
        openRecordsChannel('전체', allLink);
    } else {
        switchPage('records', document.querySelector('[data-page="records"]'));
    }
}

// 이전주차 개선 모드 토글 — 활성 시 모든 발굴일 필터 해제, 비활성 시 현재 월로 복원
function togglePrevWeekImproved() {
    if (_activePrevWeekImproved) {
        // 비활성화: 현재 월로 드롭다운 복원
        _activePrevWeekImproved = false;
        var now = new Date();
        setFilterValue("f-rec-year", String(now.getFullYear()));
        if (typeof updateMonthDropdown === "function") updateMonthDropdown();
        setFilterValue("f-rec-month", (now.getMonth() + 1) + "월");
        setFilterValue("f-rec-week", "0");
    } else {
        // 활성화: 모든 발굴일 필터 해제 + 팀/상태 필터 클리어
        _activePrevWeekImproved = true;
        _activeTeamFilter = "";
        _activeStatusFilter = "";
        setFilterValue("f-rec-year", "전체");
        if (typeof updateMonthDropdown === "function") updateMonthDropdown();
        setFilterValue("f-rec-month", "전체");
        setFilterValue("f-rec-week", "0");
    }
    fetchSummary();
}

function applyPendingRecordsFilter() {
    const f = window._pendingRecordsFilter;
    if (!f) return;
    window._pendingRecordsFilter = null;

    setFilterValue("f-rec-year", f.year);
    if (typeof updateMonthDropdown === "function") updateMonthDropdown();
    setFilterValue("f-rec-month", f.month);
    if (lastSummaryData && lastSummaryData.records) {
        updateWeekDropdown(lastSummaryData.records);
    }
    setFilterValue("f-rec-week", String(f.week));

    // 발굴/개선 카드 활성화
    _activeTeamFilter = "";
    _activeStatusFilter = f.statusFilter;
    _activePrevWeekImproved = !!f.prevWeekImproved;

    fetchSummary();
}

function showRecordDetail(r) {
    let modal = document.getElementById("record-detail-modal");
    if (!modal) {
        modal = document.createElement("div");
        modal.id = "record-detail-modal";
        modal.className = "custom-confirm-overlay";
        modal.onclick = function(e) {
            if (e.target === modal) modal.style.display = "none";
        };
        modal.innerHTML = '<div class="record-detail-box"><button class="record-detail-close" onclick="document.getElementById(\'record-detail-modal\').style.display=\'none\'">×</button><div id="record-detail-content"></div></div>';
        document.body.appendChild(modal);
    }
    const content = document.getElementById("record-detail-content");
    function row(label, value) {
        if (!value) return '';
        return '<div class="rd-row"><div class="rd-label">' + label + '</div><div class="rd-value">' + escapeHtml(String(value)) + '</div></div>';
    }
    let html = '<h3 style="margin:0 0 16px;">No.' + r.no + ' 위험요소 상세</h3>';
    html += row('채널', r.channel);
    html += row('월', r.month);
    html += row('담당자', r.person);
    html += row('일시', r.date);
    html += row('장소', r.location);
    html += row('위험요소 내용', r.content_full || r.content);
    html += row('재해유형', r.disaster_type);
    html += row('공정', r.process);
    html += row('가능성(전)', r.likelihood_before);
    html += row('중대성(전)', r.severity_before);
    html += '<div class="rd-row"><div class="rd-label">위험등급(전)</div><div class="rd-value"><span class="grade-badge grade-' + r.grade_before + '">' + r.grade_before + '</span></div></div>';
    html += row('개선대책', r.improvement_plan);
    html += row('가능성(후)', r.likelihood_after);
    html += row('중대성(후)', r.severity_after);
    html += '<div class="rd-row"><div class="rd-label">위험등급(후)</div><div class="rd-value"><span class="grade-badge grade-' + (r.grade_after || "-") + '">' + (r.grade_after || "-") + '</span></div></div>';
    html += row('완료여부', r.completion);
    html += row('완료일', r.actual_date);
    html += row('주차', r.week);
    if (r.is_repeat) html += row('반복', r.repeat_count + '회');
    if (r.image) html += '<div class="rd-row"><div class="rd-label">개선 전 사진</div><div class="rd-value"><img src="' + escapeHtml(r.image) + '" style="max-width:300px;max-height:300px;cursor:pointer;" onclick="showImageModal(\'' + escapeHtml(r.image) + '\')"></div></div>';
    if (r.image_after) html += '<div class="rd-row"><div class="rd-label">개선 후 사진</div><div class="rd-value"><img src="' + escapeHtml(r.image_after) + '" style="max-width:300px;max-height:300px;cursor:pointer;" onclick="showImageModal(\'' + escapeHtml(r.image_after) + '\')"></div></div>';
    content.innerHTML = html;
    modal.style.display = "flex";
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
    updateMonthDropdown();
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

    // 로딩 표시
    const btn = input.parentElement.querySelector('.btn-upload');
    const origText = btn ? btn.textContent : '';
    if (btn) { btn.textContent = '업로드 중...'; btn.disabled = true; btn.style.opacity = '0.6'; }

    try {
        const res = await fetch("/api/upload", { method: "POST", headers: authHeaders(), body: formData });
        if (res.status === 401) { logout(); return; }
        const text = await res.text();
        let data;
        try { data = JSON.parse(text); } catch { data = null; }
        if (!res.ok) { alert("업로드 실패: " + (data?.detail || text || "서버 오류")); return; }
        alert((data?.message || "업로드 완료") + "\n(이미지는 백그라운드에서 처리 중입니다)");
        fetchSummary();
    } catch (e) { alert("업로드 실패: " + e.message); }
    finally {
        if (btn) { btn.textContent = origText; btn.disabled = false; btn.style.opacity = '1'; }
    }
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
function checkStep2Required() {
    const causeObject = document.getElementById("ar-cause-object").value.trim();
    const disaster = document.getElementById("ar-disaster").value.trim();
    const content = document.getElementById("ar-content").value.trim();
    const lh = document.getElementById("ar-lh-before").value;
    const sv = document.getElementById("ar-sv-before").value;
    const improvement = document.getElementById("ar-improvement").value.trim();
    const btn = document.getElementById("btn-step2-next");
    btn.disabled = !(causeObject && disaster && content && lh && sv && improvement);
}

function checkStep1Required() {
    const person = document.getElementById("ar-person").value.trim();
    const location = document.getElementById("ar-location").value;
    const workplace = document.getElementById("ar-workplace").value;
    const btn = document.getElementById("btn-step1-next");
    btn.disabled = !(person && location && workplace);
}

function toggleCompletion() {
    const input = document.getElementById("ar-completion");
    const track = document.getElementById("completion-track");
    const label = document.getElementById("completion-label");
    const fields = document.getElementById("completion-fields");
    if (input.value === "미완료") {
        input.value = "완료";
        track.classList.add("active");
        label.textContent = "완료";
        fields.style.display = "block";
    } else {
        input.value = "미완료";
        track.classList.remove("active");
        label.textContent = "미완료";
        fields.style.display = "none";
    }
    updateCompletionDateDisplay();
    checkStep3Submit();
}

function updateCompletionDateDisplay() {
    const completion = document.getElementById("ar-completion").value;
    const wrap = document.getElementById("completion-date-display");
    const valEl = document.getElementById("completion-date-value");
    if (!wrap || !valEl) return;
    if (completion === "완료") {
        const existing = document.getElementById("ar-actual-date").value;
        const today = getLocalDateStr();
        valEl.textContent = existing || today;
        wrap.style.display = "";
    } else {
        wrap.style.display = "none";
    }
}

function setCompletionToggle(value) {
    const input = document.getElementById("ar-completion");
    const track = document.getElementById("completion-track");
    const label = document.getElementById("completion-label");
    const fields = document.getElementById("completion-fields");
    if (value === "완료") {
        input.value = "완료";
        track.classList.add("active");
        label.textContent = "완료";
        fields.style.display = "block";
    } else {
        input.value = "미완료";
        track.classList.remove("active");
        label.textContent = "미완료";
        fields.style.display = "none";
    }
    updateCompletionDateDisplay();
    checkStep3Submit();
}

function checkStep3Submit() {
    const btn = document.getElementById("ar-submit-btn");
    const warning = document.getElementById("after-risk-warning");
    const completion = document.getElementById("ar-completion").value;

    if (completion === "완료") {
        const lhAfter = parseInt(document.getElementById("ar-lh-after").value) || 0;
        const svAfter = parseInt(document.getElementById("ar-sv-after").value) || 0;
        const lhBefore = parseInt(document.getElementById("ar-lh-before").value) || 0;
        const svBefore = parseInt(document.getElementById("ar-sv-before").value) || 0;
        const riskBefore = lhBefore * svBefore;
        const riskAfter = lhAfter * svAfter;

        if (!lhAfter || !svAfter) {
            btn.disabled = true;
            warning.style.display = "none";
        } else if (riskAfter >= riskBefore) {
            btn.disabled = true;
            warning.style.display = "block";
        } else {
            btn.disabled = false;
            warning.style.display = "none";
        }
    } else {
        btn.disabled = false;
        warning.style.display = "none";
    }
}

function updateBeforeRiskDisplay() {
    const lh = parseInt(document.getElementById("ar-lh-before").value) || 0;
    const sv = parseInt(document.getElementById("ar-sv-before").value) || 0;
    const el = document.getElementById("before-risk-display");
    if (lh > 0 && sv > 0) {
        const risk = lh * sv;
        const grade = risk <= 4 ? "A" : risk <= 8 ? "B" : risk <= 12 ? "C" : "D";
        el.textContent = "가능성 " + lh + " × 중대성 " + sv + " = " + risk + " (" + grade + "등급)";
    } else {
        el.textContent = "-";
    }
}

function customConfirm(msg) {
    return new Promise(function(resolve) {
        var overlay = document.getElementById("custom-confirm");
        document.getElementById("custom-confirm-msg").textContent = msg;
        overlay.style.display = "flex";
        document.getElementById("custom-confirm-ok").onclick = function() {
            overlay.style.display = "none"; resolve(true);
        };
        document.getElementById("custom-confirm-cancel").onclick = function() {
            overlay.style.display = "none"; resolve(false);
        };
    });
}

async function goStep3() {
    const img = document.getElementById("ar-image-url").value;
    if (!img) {
        const ok = await customConfirm("개선 전 사진이 등록되지 않았습니다.\n사진 없이 진행하시겠습니까?");
        if (!ok) return;
    }
    wizardGo(3);
}

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
    if (step === 3) { updateBeforeRiskDisplay(); checkStep3Submit(); }
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
    document.getElementById("ar-actual-date").value = "";
    const toggleEl = document.getElementById("completion-toggle");
    toggleEl.classList.remove("locked");
    toggleEl.onclick = toggleCompletion;
    setFormMode(null);
    setCompletionToggle("미완료");
    wizardGo(1);
    document.getElementById("btn-step1-next").disabled = true;
    document.getElementById("btn-step2-next").disabled = true;
    initDateDefaults();
    updateChannelOptions();
}

function getEditFormSnapshot() {
    const ids = [
        "ar-channel", "ar-person", "ar-date", "ar-location", "ar-workplace",
        "ar-content", "ar-cause-object", "ar-disaster", "ar-week",
        "ar-improvement", "ar-actual-date", "ar-completion",
        "ar-lh-before", "ar-sv-before", "ar-lh-after", "ar-sv-after",
        "ar-image-url", "ar-image-after-url",
    ];
    return ids.map(id => {
        const el = document.getElementById(id);
        return el ? (el.value || "") : "";
    }).join("\u0001");
}

let editFormSnapshot = null;

function cancelEdit() {
    if (editingRecordId) {
        const changed = editFormSnapshot === null || getEditFormSnapshot() !== editFormSnapshot;
        if (changed) {
            if (!confirm("수정 중인 내용이 저장되지 않습니다. 닫으시겠습니까?")) return;
        }
    }
    closeEditModal();
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
    if (phase === "before") checkStep2Required();
    if (phase === "after") checkStep3Submit();
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
    if (document.getElementById("ar-submit-btn").disabled) return;
    const _completion = document.getElementById("ar-completion").value;
    if (_completion === "완료") {
        const imgAfter = document.getElementById("ar-image-after-url").value;
        if (!imgAfter) {
            const ok = await customConfirm("개선 후 사진이 등록되지 않았습니다.\n사진 없이 등록하시겠습니까?");
            if (!ok) return;
        }
    }
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
        if (isEdit) closeEditModal();
        resetForm();
        fetchSummary();
    } catch (e) { alert((isEdit ? "수정" : "등록") + " 실패: " + e.message); }
}

// ===== Edit / Delete =====
function editRecord(id) {
    if (!lastSummaryData) return;
    const r = lastSummaryData.records.find(rec => rec._id === id);
    if (!r) { alert("레코드를 찾을 수 없습니다."); return; }
    doEditRecord(id, "edit");
}

function confirmCompletionTransition() {
    const current = document.getElementById("ar-completion").value;
    if (current === "미완료") {
        const lh = parseInt(document.getElementById("ar-lh-after").value) || 0;
        const sv = parseInt(document.getElementById("ar-sv-after").value) || 0;
        if (!lh || !sv) {
            alert("개선 후 가능성/중대성 평가를 먼저 입력해주세요.");
            return;
        }
        if (!confirm("이 위험요소를 완료로 전환하시겠습니까?\n완료 등록일이 오늘 날짜로 기록됩니다.\n(전환 후 하단의 '수정' 버튼을 눌러 저장해주세요.)")) return;
        setCompletionToggle("완료");
    } else {
        if (!confirm("완료 상태를 취소하시겠습니까?\n완료 등록일이 초기화됩니다.\n(전환 후 하단의 '수정' 버튼을 눌러 저장해주세요.)")) return;
        setCompletionToggle("미완료");
    }
}

let editFormOriginalParent = null;

function setStep12ReadOnly(readOnly) {
    ["wizard-page-1", "wizard-page-2"].forEach(pid => {
        const page = document.getElementById(pid);
        if (!page) return;
        page.classList.toggle("view-only", readOnly);
        page.querySelectorAll("input, select, textarea").forEach(el => {
            if (el.type === "hidden") return;
            el.disabled = readOnly;
        });
    });
}

// mode: "complete" | "edit" | null (기본 위자드)
function setFormMode(mode) {
    const form = document.getElementById("add-record-form");
    const stepsEl = form.querySelector(".wizard-steps");
    const page1 = document.getElementById("wizard-page-1");
    const page2 = document.getElementById("wizard-page-2");
    const page3 = document.getElementById("wizard-page-3");
    const toggleWrap = document.getElementById("view-collapse-toggle");
    const prevBtn = document.getElementById("btn-step3-prev");
    const step1NavEl = page1 ? page1.querySelector(".wizard-nav") : null;
    const step2NavEl = page2 ? page2.querySelector(".wizard-nav") : null;

    if (mode === "complete") {
        // 완료 등록: Step 3만 보이고 1,2는 접힘 토글로 열람
        stepsEl.style.display = "none";
        toggleWrap.style.display = "";
        page1.style.display = "none";
        page2.style.display = "none";
        page3.style.display = "";
        if (step1NavEl) step1NavEl.style.display = "none";
        if (step2NavEl) step2NavEl.style.display = "none";
        if (prevBtn) prevBtn.style.display = "none";
        document.getElementById("view-collapse-icon").textContent = "▶";
        setStep12ReadOnly(true);
    } else if (mode === "edit") {
        // 내용 수정: 1,2,3 모두 펼쳐서 스크롤로 확인·편집
        stepsEl.style.display = "none";
        toggleWrap.style.display = "none";
        page1.style.display = "";
        page2.style.display = "";
        page3.style.display = "";
        if (step1NavEl) step1NavEl.style.display = "none";
        if (step2NavEl) step2NavEl.style.display = "none";
        if (prevBtn) prevBtn.style.display = "none";
        setStep12ReadOnly(false);
    } else {
        // 기본 위자드 (신규 등록)
        stepsEl.style.display = "";
        toggleWrap.style.display = "none";
        if (step1NavEl) step1NavEl.style.display = "";
        if (step2NavEl) step2NavEl.style.display = "";
        if (prevBtn) prevBtn.style.display = "";
        setStep12ReadOnly(false);
    }
}

// 하위 호환
function setCompleteMode(on) { setFormMode(on ? "complete" : null); }

function toggleViewCollapse() {
    const page1 = document.getElementById("wizard-page-1");
    const page2 = document.getElementById("wizard-page-2");
    const icon = document.getElementById("view-collapse-icon");
    const expanded = page1.style.display !== "none";
    if (expanded) {
        page1.style.display = "none";
        page2.style.display = "none";
        icon.textContent = "▶";
    } else {
        page1.style.display = "";
        page2.style.display = "";
        icon.textContent = "▼";
    }
}

function openEditModal(mode) {
    const formEl = document.getElementById("method-direct");
    const modal = document.getElementById("edit-record-modal");
    const body = document.getElementById("edit-record-modal-body");
    if (!editFormOriginalParent) editFormOriginalParent = formEl.parentNode;
    body.appendChild(formEl);
    formEl.style.display = "";
    document.getElementById("edit-record-modal-title").textContent =
        mode === "complete" ? "완료 등록" : "위험요소 수정";
    modal.style.display = "flex";
}

function closeEditModal() {
    const formEl = document.getElementById("method-direct");
    const modal = document.getElementById("edit-record-modal");
    if (editFormOriginalParent && formEl.parentNode !== editFormOriginalParent) {
        editFormOriginalParent.appendChild(formEl);
        formEl.style.display = "none";
    }
    modal.style.display = "none";
}

function doEditRecord(id, mode) {
    if (!lastSummaryData) return;
    const r = lastSummaryData.records.find(rec => rec._id === id);
    if (!r) { alert("레코드를 찾을 수 없습니다."); return; }

    openEditModal(mode);

    editingRecordId = id;
    document.getElementById("register-page-title").textContent = "위험요소 수정 (No." + r.no + ")";
    document.getElementById("ar-submit-btn").textContent = "수정";
    document.getElementById("ar-cancel-btn").style.display = "inline-block";

    // Update channel select to allow all options if editing
    updateChannelOptions(true);

    document.getElementById("ar-channel").value = r.channel || "부서별 위험요소발굴";
    syncChannelCard();
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
    document.getElementById("ar-actual-date").value = r.actual_date || "";
    setCompletionToggle(r.completion || "미완료");

    clearRatingCards();
    if (r.likelihood_before) selectRating("lh", "before", r.likelihood_before);
    if (r.severity_before) selectRating("sv", "before", r.severity_before);
    if (r.likelihood_after) selectRating("lh", "after", r.likelihood_after);
    if (r.severity_after) selectRating("sv", "after", r.severity_after);

    // Lazy load images for edit
    const imgPromises = [];
    if (r.has_image) {
        imgPromises.push(fetch("/api/record-image/" + id + "?field=image", { headers: authHeaders() })
            .then(res => res.json()).then(data => {
                if (data.url) {
                    document.getElementById("ar-image-url").value = data.url;
                    document.getElementById("ar-image-name").textContent = "기존 사진";
                    document.getElementById("ar-image-thumb").src = data.url;
                    document.getElementById("ar-image-preview").style.display = "flex";
                }
            }).catch(() => {}));
    }
    if (r.has_image_after) {
        imgPromises.push(fetch("/api/record-image/" + id + "?field=image_after", { headers: authHeaders() })
            .then(res => res.json()).then(data => {
                if (data.url) {
                    document.getElementById("ar-image-after-url").value = data.url;
                    document.getElementById("ar-image-after-name").textContent = "기존 사진";
                    document.getElementById("ar-image-after-thumb").src = data.url;
                    document.getElementById("ar-image-after-preview").style.display = "flex";
                }
            }).catch(() => {}));
    }
    editFormSnapshot = getEditFormSnapshot();
    if (imgPromises.length) {
        Promise.all(imgPromises).then(() => {
            if (editingRecordId === id) editFormSnapshot = getEditFormSnapshot();
        });
    }

    const toggleEl = document.getElementById("completion-toggle");
    toggleEl.classList.remove("locked");
    toggleEl.onclick = confirmCompletionTransition;
    setFormMode("edit");

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
    const adminOnlyEls = document.querySelectorAll(".admin-only");

    if (ADMIN_TOKEN) {
        btn.textContent = "관리자 ✓";
        btn.classList.add("btn-admin-active");
        adminOnlyEls.forEach(el => el.style.display = "");
    } else {
        btn.textContent = "관리자";
        btn.classList.remove("btn-admin-active");
        adminOnlyEls.forEach(el => el.style.display = "none");
    }
    updateChannelOptions();
    // Show/hide 작업 column header
    document.querySelectorAll(".th-action").forEach(el => {
        el.style.display = ADMIN_TOKEN ? "" : "none";
    });
    // Redraw table so edit/delete buttons reflect admin state
    if (lastSummaryData) updateTable(getDisplayRecords(lastSummaryData));
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
        ["부서별 위험요소발굴", "5S/EHS평가"].forEach(ch => {
            const o = document.createElement("option");
            o.value = ch; o.textContent = ch; arChannel.appendChild(o);
        });
        arChannel.value = ["부서별 위험요소발굴", "5S/EHS평가"].includes(prevArVal) ? prevArVal : "부서별 위험요소발굴";
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
        ["부서별 위험요소발굴", "5S/EHS평가"].forEach(ch => {
            const o = document.createElement("option");
            o.value = ch; o.textContent = ch; uploadChannel.appendChild(o);
        });
        uploadChannel.value = ["부서별 위험요소발굴", "5S/EHS평가"].includes(prevUpVal) ? prevUpVal : "부서별 위험요소발굴";
    }

    syncChannelCard();
}

const CHANNEL_DESCRIPTIONS = {
    "부서별 위험요소발굴": "평소 업무 중 직접 발견하신 위험요소를 등록해주세요.",
    "5S/EHS평가": "5S 점검 활동 중 발견한 위험요소를 등록해주세요."
};

function syncChannelCard() {
    const arChannel = document.getElementById("ar-channel");
    const optionsWrap = document.getElementById("ar-channel-options");
    if (!arChannel || !optionsWrap) return;

    const currentVal = arChannel.value;
    const channels = Array.from(arChannel.options).map(o => o.value);
    const isAdminMode = channels.length > 2;

    optionsWrap.className = "channel-card-options" + (isAdminMode ? " admin" : "");
    optionsWrap.style.gridTemplateColumns = "repeat(" + channels.length + ", minmax(0, 1fr))";

    optionsWrap.innerHTML = "";
    channels.forEach(ch => {
        const label = document.createElement("label");
        label.className = "channel-option" + (ch === currentVal ? " selected" : "");
        const input = document.createElement("input");
        input.type = "radio";
        input.name = "ar-channel-radio";
        input.value = ch;
        if (ch === currentVal) input.checked = true;
        input.addEventListener("change", () => {
            arChannel.value = ch;
            optionsWrap.querySelectorAll(".channel-option").forEach(el => el.classList.remove("selected"));
            label.classList.add("selected");
        });
        const title = document.createElement("span");
        title.className = "channel-option-title";
        title.textContent = ch;
        label.appendChild(input);
        label.appendChild(title);

        if (!isAdminMode && CHANNEL_DESCRIPTIONS[ch]) {
            const desc = document.createElement("span");
            desc.className = "channel-option-desc";
            desc.textContent = CHANNEL_DESCRIPTIONS[ch];
            label.appendChild(desc);
        }

        optionsWrap.appendChild(label);
    });
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
        // 주간현황표 URL (현재 연도·분기)
        const now = new Date();
        const curYear = String(now.getFullYear());
        const curQuarter = Math.ceil((now.getMonth() + 1) / 3);

        // 두 요청을 병렬 실행
        const [res, wkRes] = await Promise.all([
            fetch("/api/summary?" + params.toString(), { headers: authHeaders() }),
            fetch("/api/weekly/quarter?year=" + curYear + "&quarter=" + curQuarter, { headers: authHeaders() })
        ]);

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
        lightData._team = filters.team || "전체";

        if (wkRes && wkRes.ok) {
            try { lightData._weeklyQuarter = await wkRes.json(); }
            catch (e) { console.error("Weekly quarter parse failed:", e); }
        }

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

    // 현재 날짜 기준 이번주 월/주차 자동 설정
    var now = new Date();
    weeklyCurrentMonth = now.getMonth() + 1;
    weeklyCurrentWeek = getWeekFromDate(getLocalDateStr(now));

    try {
        const liveRes = await fetch("/api/weekly/quarter?year=" + year + "&quarter=" + quarter, { headers: authHeaders() });
        weeklyLiveData = await liveRes.json();
        renderQuarterTable(weeklyLiveData);
    } catch (e) { console.error("Weekly load error:", e); }
}

async function loadSummaryWeekly() {
    const yearEl = document.getElementById("sw-year");
    const quarterEl = document.getElementById("sw-quarter");
    if (!yearEl || !quarterEl) return;

    // 분기 셀렉트가 아직 초기값이면 현재 분기로 자동 세팅
    var now = new Date();
    if (!quarterEl.dataset.initialized) {
        quarterEl.value = String(Math.ceil((now.getMonth() + 1) / 3));
        quarterEl.dataset.initialized = "1";
    }

    const year = yearEl.value;
    const quarter = quarterEl.value;

    weeklyCurrentMonth = now.getMonth() + 1;
    weeklyCurrentWeek = getWeekFromDate(getLocalDateStr(now));

    try {
        const liveRes = await fetch("/api/weekly/quarter?year=" + year + "&quarter=" + quarter, { headers: authHeaders() });
        summaryWeeklyData = await liveRes.json();
        renderQuarterTable(summaryWeeklyData, { containerId: "summary-weekly-tables", siteSelectId: "f-team" });
    } catch (e) { console.error("Summary weekly load error:", e); }
}

function onSummaryTeamChange() {
    fetchSummary();
    if (summaryWeeklyData) {
        renderQuarterTable(summaryWeeklyData, { containerId: "summary-weekly-tables", siteSelectId: "f-team" });
    } else {
        loadSummaryWeekly();
    }
}

function renderQuarterTable(data, opts) {
    opts = opts || {};
    const containerId = opts.containerId || "weekly-tables";
    const siteSelectId = opts.siteSelectId || "w-site";
    const container = document.getElementById(containerId);
    if (!container) return;
    if (!data || !data.sites) { container.innerHTML = ""; return; }

    const months = data.months || [];
    const channels = data.channel_order || [];
    const priorQuarters = data.prior_quarters || [];
    const selectedSite = (document.getElementById(siteSelectId) || {}).value || "전체";
    const siteNames = [selectedSite];
    const curMonth = weeklyCurrentMonth;
    const curWeek = weeklyCurrentWeek;

    function _deltaBadge(n) {
        if (!n || n <= 0) return "";
        return '<div class="wt-delta">+' + n + '</div>';
    }

    const totalSite = data.sites["전체"] || {};
    const has5th = {};
    months.forEach(m => {
        has5th[m] = false;
        // 달력상 5주차(마지막 목요일이 29/30/31일) 존재 여부
        const lastDay = new Date(parseInt(data.year), m, 0).getDate();
        for (let day = 29; day <= lastDay; day++) {
            const dt = new Date(parseInt(data.year), m - 1, day);
            if (dt.getDay() === 4) { has5th[m] = true; break; }
        }
        // 데이터 기반도 확인
        [...channels, "합계"].forEach(ch => {
            const wk = (totalSite[ch] || {}).weeks || {};
            const w5 = wk[m + "-5"];
            if (w5 && (w5.discovered > 0 || w5.improved > 0)) has5th[m] = true;
        });
    });

    let html = "";
    siteNames.forEach(siteName => {
        const siteData = data.sites[siteName] || {};

        html += '<div class="weekly-table-wrap"><h4 class="weekly-site-title">' + siteName + '</h4>';
        html += '<table class="weekly-table"><thead>';

        html += '<tr><th rowspan="2" class="wt-fixed">구분</th><th rowspan="2" class="wt-fixed"></th>';
        if (priorQuarters.length > 0) {
            html += '<th colspan="2" class="wt-month-group wt-month-start wt-prior-group">' + priorQuarters.join('+') + '분기</th>';
        }
        months.forEach(m => {
            const weekCount = has5th[m] ? 5 : 4;
            html += '<th colspan="' + (weekCount+1) + '" class="wt-month-group">' + m + '월</th>';
        });
        html += '<th colspan="3" class="wt-month-group wt-month-start wt-year-group" style="border-right:1px solid #999;">' + data.year + '년</th>';
        html += '</tr>';

        html += '<tr>';
        if (priorQuarters.length > 0) {
            html += '<th class="wt-prior-q wt-month-start">누적</th>';
            html += '<th class="wt-prior-avg">월평균</th>';
        }
        months.forEach(m => {
            const maxW = has5th[m] ? 5 : 4;
            for (let w = 1; w <= maxW; w++) {
                const isCur = (m === curMonth && w === curWeek);
                const isFuture = (m > curMonth || (m === curMonth && w > curWeek));
                html += '<th class="' + (isCur?"wt-current":"") + (isFuture?" wt-future":"") + ' ' + (w===1?"wt-month-start":"") + '">' + w + '주</th>';
            }
            html += '<th class="wt-sub' + (m > curMonth?" wt-future":"") + '">소계</th>';
        });
        html += '<th class="wt-month-start wt-year-cell">' + data.quarter + '분기</th>';
        html += '<th class="wt-year-cell">합계</th>';
        html += '<th class="wt-year-cell" style="border-left:1px solid #999;">개선률</th>';
        html += '</tr></thead><tbody>';

        const allCh = [...channels, "합계"];
        allCh.forEach((ch, idx) => {
            const d = siteData[ch] || {};
            const isTotal = ch === "합계";
            const isLastBeforeTotal = (idx === allCh.length - 2);
            const rowCls = isTotal ? "weekly-total-row" : "";

            const priorData = d.prior_quarters || {};
            const rateVal = (d.cum_rate !== undefined) ? d.cum_rate : d.quarter_rate;
            const cumDisc = (d.cum_discovered !== undefined) ? d.cum_discovered : (d.quarter_discovered || 0);
            const cumImp = (d.cum_improved !== undefined) ? d.cum_improved : (d.quarter_improved || 0);

            // 전 분기 누적 합산 (여러 분기 묶어서 하나의 열로)
            let priorDiscSum = 0, priorImpSum = 0;
            priorQuarters.forEach(q => {
                const pq = priorData[q] || {};
                priorDiscSum += (pq.discovered || 0);
                priorImpSum += (pq.improved || 0);
            });
            const priorMonths = priorQuarters.length * 3;
            const priorDiscAvg = priorMonths > 0 ? Math.round(priorDiscSum / priorMonths) : 0;
            const priorImpAvg = priorMonths > 0 ? Math.round(priorImpSum / priorMonths) : 0;

            html += '<tr class="' + rowCls + ' wt-ch-first">';
            html += '<td class="ch-name' + (isLastBeforeTotal?" wt-border-bottom":"") + '" rowspan="2">' + ch + '</td>';
            html += '<td class="row-type">발굴</td>';
            if (priorQuarters.length > 0) {
                let priorDiscDeltaSum = 0;
                priorQuarters.forEach(q => {
                    const pq = priorData[q] || {};
                    priorDiscDeltaSum += (pq.week_discovered_delta || 0);
                });
                html += '<td class="num wt-qtr wt-month-start wt-prior-q">' + priorDiscSum + _deltaBadge(priorDiscDeltaSum) + '</td>';
                html += '<td class="num wt-qtr wt-prior-avg">' + priorDiscAvg + '</td>';
            }
            months.forEach(m => {
                const maxW = has5th[m] ? 5 : 4;
                for (let w = 1; w <= maxW; w++) {
                    const wk = d.weeks ? (d.weeks[m+"-"+w] || {}) : {};
                    const val = wk.discovered || 0;
                    const delta = wk.week_discovered_delta || 0;
                    const isCur = (m === curMonth && w === curWeek);
                    const isFuture = (m > curMonth || (m === curMonth && w > curWeek));
                    const display = isFuture ? "" : val;
                    // 현재 주차 셀은 delta가 값과 동일하니 생략 (중복 방지)
                    const showDelta = !isFuture && !isCur;
                    const deltaHtml = showDelta ? _deltaBadge(delta) : "";
                    html += '<td class="num ' + (isCur?"wt-current":"") + (isFuture?" wt-future":"") + ' ' + (w===1?"wt-month-start":"") + '">' + display + deltaHtml + '</td>';
                }
                const sub = d.month_subs ? (d.month_subs[String(m)] || {}) : {};
                var monthFuture = m > curMonth;
                const subDelta = sub.week_discovered_delta || 0;
                html += '<td class="num wt-sub' + (monthFuture?" wt-future":"") + '">' + (monthFuture ? "" : (sub.discovered||0)) + (monthFuture ? "" : _deltaBadge(subDelta)) + '</td>';
            });
            const qDiscDelta = d.quarter_week_discovered_delta || 0;
            const cumDiscDelta = d.cum_week_discovered_delta || 0;
            html += '<td class="num wt-qtr wt-month-start wt-year-cell">' + (d.quarter_discovered||0) + _deltaBadge(qDiscDelta) + '</td>';
            html += '<td class="num wt-qtr wt-year-cell">' + cumDisc + _deltaBadge(cumDiscDelta) + '</td>';
            html += '<td class="num rate wt-year-cell' + (isLastBeforeTotal?" wt-border-bottom":"") + '" rowspan="2">' + (rateVal ? Math.round(rateVal*100)+"%" : "-") + '</td>';
            html += '</tr>';

            html += '<tr class="' + rowCls + ' ' + (isLastBeforeTotal?"wt-before-total":"") + '">';
            html += '<td class="row-type">개선</td>';
            if (priorQuarters.length > 0) {
                // 전 분기 누적의 증감: 각 분기 week_improved_delta 합
                let priorDeltaSum = 0;
                priorQuarters.forEach(q => {
                    const pq = priorData[q] || {};
                    priorDeltaSum += (pq.week_improved_delta || 0);
                });
                html += '<td class="num wt-qtr wt-month-start wt-prior-q">' + priorImpSum + _deltaBadge(priorDeltaSum) + '</td>';
                html += '<td class="num wt-qtr wt-prior-avg">' + priorImpAvg + '</td>';
            }
            months.forEach(m => {
                const maxW = has5th[m] ? 5 : 4;
                for (let w = 1; w <= maxW; w++) {
                    const wk = d.weeks ? (d.weeks[m+"-"+w] || {}) : {};
                    const val = wk.improved || 0;
                    const delta = wk.week_improved_delta || 0;
                    const isCur = (m === curMonth && w === curWeek);
                    const isFuture = (m > curMonth || (m === curMonth && w > curWeek));
                    const display = isFuture ? "" : val;
                    const deltaHtml = isFuture ? "" : _deltaBadge(delta);
                    html += '<td class="num ' + (isCur?"wt-current":"") + (isFuture?" wt-future":"") + ' ' + (w===1?"wt-month-start":"") + '">' + display + deltaHtml + '</td>';
                }
                const sub = d.month_subs ? (d.month_subs[String(m)] || {}) : {};
                var monthFuture = m > curMonth;
                const subDelta = sub.week_improved_delta || 0;
                html += '<td class="num wt-sub' + (monthFuture?" wt-future":"") + '">' + (monthFuture ? "" : (sub.improved||0)) + (monthFuture ? "" : _deltaBadge(subDelta)) + '</td>';
            });
            const qDelta = d.quarter_week_improved_delta || 0;
            const cumDelta = d.cum_week_improved_delta || 0;
            html += '<td class="num wt-qtr wt-month-start wt-year-cell">' + (d.quarter_improved||0) + _deltaBadge(qDelta) + '</td>';
            html += '<td class="num wt-qtr wt-year-cell">' + cumImp + _deltaBadge(cumDelta) + '</td>';
            html += '</tr>';
        });

        html += '</tbody></table></div>';
    });

    container.innerHTML = html;
}

// ===== Executive Comments =====
let commentPanelOpen = false;
let currentWeekKey = "";
let allWeeksData = [];
let currentDefaults = {};

function toggleCommentPanel() {
    commentPanelOpen = !commentPanelOpen;
    document.getElementById("comment-panel").classList.toggle("open", commentPanelOpen);
    document.getElementById("comment-overlay").classList.toggle("open", commentPanelOpen);
    if (commentPanelOpen) {
        loadWeeks();
        loadNotifications();
    }
}

async function loadWeeks() {
    try {
        const res = await fetch("/api/comment-weeks", { headers: authHeaders() });
        const data = await res.json();
        allWeeksData = data.weeks || [];
        currentDefaults = data.current || {};

        // 연도 드롭다운
        const years = data.years || [];
        years.sort((a, b) => a - b);
        const yearSelect = document.getElementById("comment-year-select");
        yearSelect.innerHTML = years.map(y =>
            `<option value="${y}">${y}년</option>`
        ).join("");
        yearSelect.value = currentDefaults.year;

        updateMonthOptions(currentDefaults.year, currentDefaults.month);
    } catch(e) { console.error("loadWeeks error", e); }
}

function onYearChange() {
    const year = parseInt(document.getElementById("comment-year-select").value);
    updateMonthOptions(year, null);
}

function updateMonthOptions(year, defaultMonth) {
    const filtered = allWeeksData.filter(w => w.year === year);
    const months = [];
    const monthSet = new Set();
    filtered.forEach(w => {
        if (!monthSet.has(w.month)) {
            monthSet.add(w.month);
            months.push(w.month);
        }
    });
    months.sort((a, b) => a - b);

    const monthSelect = document.getElementById("comment-month-select");
    monthSelect.innerHTML = months.map(m =>
        `<option value="${m}">${m}월</option>`
    ).join("");

    const selectMonth = defaultMonth && months.includes(defaultMonth) ? defaultMonth : months[0];
    if (selectMonth) {
        monthSelect.value = selectMonth;
        updateWeekOptions(year, selectMonth);
    }
}

function onMonthChange() {
    const year = parseInt(document.getElementById("comment-year-select").value);
    const month = parseInt(document.getElementById("comment-month-select").value);
    updateWeekOptions(year, month);
}

function updateWeekOptions(year, month) {
    const weekSelect = document.getElementById("comment-week-select");
    const filtered = allWeeksData.filter(w => w.year === year && w.month === month);
    filtered.sort((a, b) => a.week_of_month - b.week_of_month);

    weekSelect.innerHTML = filtered.map(w =>
        `<option value="${w.key}">${w.week_of_month}주차</option>`
    ).join("");

    // 현재 주차가 이 월에 있으면 선택, 아니면 최신 주차
    const currentInList = filtered.find(w => w.key === currentDefaults.week_key);
    if (currentInList) {
        currentWeekKey = currentInList.key;
    } else if (filtered.length > 0) {
        currentWeekKey = filtered[0].key;
    }
    weekSelect.value = currentWeekKey;
    loadComments();
}

function onWeekChange() {
    currentWeekKey = document.getElementById("comment-week-select").value;
    loadComments();
}

const ROLE_LABEL_MAP = {
    "1팀장": "환경안전1팀 팀장",
    "2팀장": "환경안전2팀 팀장",
    "본부장": "생산기술본부장",
    "부문장": "SCM 부문장",
    "대표이사": "대표이사",
};

function switchCommentTab(tab, btn) {
    document.querySelectorAll(".comment-tab").forEach(t => t.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("comment-view-register").style.display = tab === "register" ? "flex" : "none";
    document.getElementById("comment-view-history").style.display = tab === "history" ? "flex" : "none";
    if (tab === "history") initHistoryFilters();
}

function initHistoryFilters() {
    const years = [...new Set(allWeeksData.map(w => w.year))].sort((a, b) => a - b);
    const yearSelect = document.getElementById("history-year-filter");
    yearSelect.innerHTML = years.map(y => `<option value="${y}">${y}년</option>`).join("");
    yearSelect.value = currentDefaults.year;
    updateHistoryMonths(currentDefaults.year, currentDefaults.month);
}

function onHistoryYearChange() {
    const year = parseInt(document.getElementById("history-year-filter").value);
    updateHistoryMonths(year, null);
}

function updateHistoryMonths(year, defaultMonth) {
    const months = [...new Set(allWeeksData.filter(w => w.year === year).map(w => w.month))].sort((a, b) => a - b);
    const monthSelect = document.getElementById("history-month-filter");
    monthSelect.innerHTML = `<option value="전체">전체</option>` + months.map(m => `<option value="${m}">${m}월</option>`).join("");
    const selectMonth = defaultMonth && months.includes(defaultMonth) ? defaultMonth : "전체";
    monthSelect.value = selectMonth;
    if (selectMonth === "전체") {
        updateHistoryWeeksAll(year);
    } else {
        updateHistoryWeeks(year, selectMonth);
    }
}

function onHistoryMonthChange() {
    const year = parseInt(document.getElementById("history-year-filter").value);
    const monthVal = document.getElementById("history-month-filter").value;
    if (monthVal === "전체") {
        updateHistoryWeeksAll(year);
    } else {
        updateHistoryWeeks(year, parseInt(monthVal));
    }
}

function updateHistoryWeeksAll(year) {
    const weekSelect = document.getElementById("history-week-filter");
    weekSelect.innerHTML = `<option value="전체">전체</option>`;
    weekSelect.value = "전체";
    loadCommentHistory();
}

function updateHistoryWeeks(year, month) {
    const filtered = allWeeksData.filter(w => w.year === year && w.month === month).sort((a, b) => a.week_of_month - b.week_of_month);
    const weekSelect = document.getElementById("history-week-filter");
    weekSelect.innerHTML = `<option value="전체">전체</option>` + filtered.map(w =>
        `<option value="${w.key}">${w.week_of_month}주차</option>`
    ).join("");
    weekSelect.value = "전체";
    loadCommentHistory();
}

async function loadCommentHistory() {
    const year = document.getElementById("history-year-filter").value;
    const monthVal = document.getElementById("history-month-filter").value;
    const weekVal = document.getElementById("history-week-filter").value;

    try {
        let url;
        if (weekVal !== "전체") {
            url = `/api/comments?week=${encodeURIComponent(weekVal)}`;
        } else {
            let filtered = allWeeksData.filter(w => w.year === parseInt(year));
            if (monthVal !== "전체") {
                filtered = filtered.filter(w => w.month === parseInt(monthVal));
            }
            const weekKeys = filtered.map(w => w.key);
            url = `/api/comments/all?weeks=${encodeURIComponent(weekKeys.join(","))}`;
        }
        const res = await fetch(url, { headers: authHeaders() });
        const data = await res.json();
        const comments = data.comments || [];
        const list = document.getElementById("comment-history-list");
        if (comments.length === 0) {
            list.innerHTML = '<div style="font-size:13px;color:#999;text-align:center;padding:20px;">코멘트가 없습니다.</div>';
            return;
        }
        list.innerHTML = comments.map(c => {
            const roleLabel = ROLE_LABEL_MAP[c.role] || c.role;
            const weekKey = c.week_key || "";
            const weekInfo = allWeeksData.find(w => w.key === weekKey);
            const weekLabel = weekInfo ? `${weekInfo.month}월 ${weekInfo.week_of_month}주차` : weekKey;
            const deleteBtn = ADMIN_TOKEN ? `<button class="comment-item-delete" onclick="deleteComment('${c.id}')" title="삭제">삭제</button>` : "";
            return `
                <div class="history-item">
                    <div class="history-item-header">
                        <span class="history-item-role">${escapeHtml(roleLabel)}</span>
                        <span class="history-item-week">${escapeHtml(weekLabel)}</span>
                    </div>
                    <div class="history-item-content">${escapeHtml(c.content)}</div>
                    <div class="history-item-footer">
                        <span class="history-item-time">${c.created_at}</span>
                        ${deleteBtn}
                    </div>
                </div>
            `;
        }).join("");
    } catch(e) { console.error("loadCommentHistory error", e); }
}

async function loadComments() {
    try {
        const url = currentWeekKey ? `/api/comments?week=${currentWeekKey}` : "/api/comments";
        const res = await fetch(url, { headers: authHeaders() });
        const data = await res.json();
        const comments = data.comments || [];
        ["1팀장","2팀장","본부장","부문장","대표이사"].forEach(role => {
            const list = document.getElementById("comments-" + role);
            const roleComments = comments.filter(c => c.role === role);
            if (roleComments.length === 0) {
                list.innerHTML = '<div style="font-size:12px;color:#999;">코멘트가 없습니다.</div>';
                return;
            }
            list.innerHTML = roleComments.map(c => `
                <div class="comment-item">
                    <div class="comment-item-content">${escapeHtml(c.content)}</div>
                    <div class="comment-item-footer">
                        <span class="comment-item-time">${c.created_at}</span>
                        <button class="comment-item-delete" onclick="deleteComment('${c.id}')" title="삭제">삭제</button>
                    </div>
                </div>
            `).join("");
        });
    } catch(e) { console.error("loadComments error", e); }
}

function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

async function submitCommentNew() {
    const role = document.getElementById("comment-role-select").value;
    const input = document.getElementById("comment-input");
    const content = input.value.trim();
    if (!content) return;
    try {
        const res = await fetch("/api/comments", {
            method: "POST",
            headers: { ...authHeaders(), "Content-Type": "application/json" },
            body: JSON.stringify({ role, content })
        });
        if (res.ok) {
            input.value = "";
            loadNotifications();
            loadSummaryNotifications();
            const data = await res.json();
            if (data.notification) showToast(data.notification.message);
            // 코멘트 확인 탭으로 전환
            const historyTab = document.querySelectorAll(".comment-tab")[1];
            switchCommentTab("history", historyTab);
        }
    } catch(e) { console.error("submitCommentNew error", e); }
}

async function deleteComment(commentId) {
    try {
        await fetch("/api/comments/" + commentId, { method: "DELETE", headers: authHeaders() });
        loadCommentHistory();
        loadNotifications();
        loadSummaryNotifications();
    } catch(e) { console.error("deleteComment error", e); }
}

async function loadNotifications() {
    try {
        const res = await fetch("/api/notifications", { headers: authHeaders() });
        const data = await res.json();
        const notifs = data.notifications || [];
        const badge = document.getElementById("comment-badge");
        if (notifs.length > 0) {
            badge.textContent = notifs.length;
            badge.style.display = "flex";
        } else {
            badge.style.display = "none";
        }
        const container = document.getElementById("comment-notifications");
        if (notifs.length === 0) {
            container.innerHTML = "";
            return;
        }
        container.innerHTML = notifs.map(n => `
            <div class="notif-item">
                <span class="notif-msg">${escapeHtml(n.message)}</span>
                <span class="notif-time">${n.created_at}</span>
                <button class="notif-dismiss" onclick="dismissNotification('${n.id}')" title="알림 삭제">&times;</button>
            </div>
        `).join("");
    } catch(e) { console.error("loadNotifications error", e); }
}

async function dismissNotification(notifId) {
    try {
        await fetch("/api/notifications/" + notifId, { method: "DELETE", headers: authHeaders() });
        loadNotifications();
    } catch(e) { console.error("dismissNotification error", e); }
}

function showToast(message) {
    const container = document.getElementById("toast-container");
    const toast = document.createElement("div");
    toast.className = "toast-item";
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => {
        toast.classList.add("fade-out");
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

async function loadSummaryNotifications() {
    try {
        const res = await fetch("/api/notifications?current_week=true", { headers: authHeaders() });
        const data = await res.json();
        const notifs = data.notifications || [];
        const container = document.getElementById("summary-notifications");
        if (notifs.length === 0) {
            container.style.display = "none";
            return;
        }
        container.style.display = "flex";
        container.innerHTML = notifs.map(n => `
            <div class="summary-notif-item" id="summary-notif-${n.id}">
                <span class="notif-dot"></span>
                <span class="notif-msg">${escapeHtml(n.message)} <span class="summary-notif-hint" onclick="event.stopPropagation();toggleCommentPanel()">💬 자세히 보기</span></span>
                <span class="notif-time">${n.created_at}</span>
                <button class="notif-close" onclick="closeSummaryNotif('${n.id}')">&times;</button>
            </div>
        `).join("");
    } catch(e) { console.error("loadSummaryNotifications error", e); }
}

function closeSummaryNotif(notifId) {
    const el = document.getElementById("summary-notif-" + notifId);
    if (el) el.remove();
    const container = document.getElementById("summary-notifications");
    if (container && container.children.length === 0) {
        container.style.display = "none";
    }
}

// Poll notifications every 30 seconds
setInterval(() => { loadNotifications(); loadSummaryNotifications(); }, 30000);

// ===== Init =====
TOKEN = "public";
sessionStorage.setItem("token", TOKEN);
showDashboard();
loadNotifications();
loadSummaryNotifications();
