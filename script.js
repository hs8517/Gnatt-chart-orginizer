/**
 * Gantt Chart Application Logic
 */

// --- State Management ---
const state = {
    tasks: [],
    viewScale: 'day', // 'day', 'week', 'month'
    viewStartDate: new Date(),
    columnWidth: 50,
    lang: localStorage.getItem('gantt_lang') || 'en',
    theme: localStorage.getItem('gantt_theme') || 'black',
    team: loadTeam(), // Load using helper to handle migration
    daysToRender: 60,
    endingBuffer: localStorage.getItem('gantt_ending_buffer') ? parseInt(localStorage.getItem('gantt_ending_buffer')) : 60,
    editingTaskId: null,
    dataSource: localStorage.getItem('gantt_data_source') || 'local'
};

// Helper to load/migrate team data
function loadTeam() {
    const stored = JSON.parse(localStorage.getItem('gantt_team'));
    if (!stored) return [{ name: 'Alice', email: '' }, { name: 'Bob', email: '' }, { name: 'Charlie', email: '' }, { name: 'David', email: '' }];

    // Migration: If array of strings, convert to objects
    if (stored.length > 0 && typeof stored[0] === 'string') {
        return stored.map(name => ({ name, email: '' }));
    }
    return stored;
}

// --- Localization ---
const translations = {
    en: {
        scale_day: "Day",
        scale_week: "Week",
        scale_month: "Month",
        btn_new_task: "New Task",
        modal_title_add: "Add New Task",
        modal_title_edit: "Edit Task",
        modal_title_settings: "Settings",
        tab_appearance: "Appearance",
        tab_team: "Team",
        tab_summary: "Summary",
        label_manage_team: "Manage Team Members",
        label_task_name: "Task Name",
        label_start_date: "Start Date",
        label_duration: "Duration (Days)",
        label_color: "Color",
        label_owner: "Responsible Person",
        label_theme: "Theme",
        btn_create: "Create Task",
        btn_save: "Save Changes",
        btn_delete: "Delete",
        btn_add: "Add",
        btn_calendar: "Add to Calendar",
        btn_confirm: "Confirm",
        theme_black: "Black",
        theme_white: "White",
        theme_ivory: "Ivory",
        label_ending_buffer: "Ending Buffer (Days)",
        // Dynamic
        export_filename_excel: "Gantt_Export",
        export_filename_pdf: "Gantt_Chart",
        col_task_name: "Task Name",
        col_owner: "Owner",
        col_start_date: "Start Date",
        col_end_date: "End Date",
        col_duration: "Duration (Days)",
        tab_data: "Data",
        label_data_source: "Data Source",
        desc_data_source: "Switching sources will reload the application with data from the selected source."
    },
    zh: {
        scale_day: "日",
        scale_week: "週",
        scale_month: "月",
        btn_new_task: "新建任務",
        modal_title_add: "添加新任務",
        modal_title_edit: "編輯任務",
        modal_title_settings: "設置",
        tab_appearance: "外觀",
        tab_team: "團隊",
        tab_summary: "概覽",
        label_manage_team: "管理團隊成員",
        label_task_name: "任務名稱",
        label_start_date: "開始日期",
        label_duration: "時長 (天)",
        label_color: "顏色",
        label_owner: "負責人",
        label_theme: "主題",
        btn_create: "創建任務",
        btn_save: "保存更改",
        btn_delete: "刪除",
        btn_add: "添加",
        btn_calendar: "添加到日曆",
        btn_confirm: "確認",
        theme_black: "暗黑",
        theme_white: "純白",
        theme_ivory: "象牙",
        label_ending_buffer: "結束緩衝 (天)",
        // Dynamic
        export_filename_excel: "甘特圖導出",
        export_filename_pdf: "甘特圖",
        col_task_name: "任務名稱",
        col_owner: "負責人",
        col_start_date: "開始日期",
        col_end_date: "結束日期",
        col_duration: "時長 (天)",
        tab_data: "數據",
        label_data_source: "數據源",
        desc_data_source: "切換數據源將使用所選源的數據重新加載應用程序。",
        // Summary
        label_total_days: "總天數",
        label_total_tasks: "總任務數",
        label_team_size: "團隊人數",
        label_longest_task: "最長任務"
    },
    ja: {
        scale_day: "日",
        scale_week: "週",
        scale_month: "月",
        btn_new_task: "新規タスク",
        modal_title_add: "新しいタスクを追加",
        modal_title_edit: "タスクを編集",
        modal_title_settings: "設定",
        tab_appearance: "外観",
        tab_team: "チーム",
        tab_summary: "概要",
        label_manage_team: "チームメンバー管理",
        label_task_name: "タスク名",
        label_start_date: "開始日",
        label_duration: "期間 (日)",
        label_color: "色",
        label_owner: "担当者",
        label_theme: "テーマ",
        btn_create: "タスクを作成",
        btn_save: "変更を保存",
        btn_delete: "削除",
        btn_add: "追加",
        btn_calendar: "カレンダーに追加",
        btn_confirm: "確認",
        theme_black: "ブラック",
        theme_white: "ホワイト",
        theme_ivory: "アイボリー",
        label_ending_buffer: "終了バッファ (日)",
        // Dynamic
        export_filename_excel: "ガントチャート出力",
        export_filename_pdf: "ガントチャート",
        col_task_name: "タスク名",
        col_owner: "担当者",
        col_start_date: "開始日",
        col_end_date: "終了日",
        col_duration: "期間 (日)",
        tab_data: "データ",
        label_data_source: "データソース",
        desc_data_source: "ソースを切り替えると、選択したソースのデータでアプリケーションが再読み込みされます。",
        // Summary
        label_total_days: "合計日数",
        label_total_tasks: "タスク総数",
        label_team_size: "チーム人数",
        label_longest_task: "最長タスク"
    }
};

function t(key) {
    return translations[state.lang][key] || key;
}

function updateLanguage(lang) {
    state.lang = lang;
    localStorage.setItem('gantt_lang', lang);
    saveToServer();
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        el.textContent = t(key);
    });
    render();
    renderTeamSettings();
}

function updateTheme(theme) {
    state.theme = theme;
    localStorage.setItem('gantt_theme', theme);
    saveToServer();
    document.body.setAttribute('data-theme', theme);

    document.querySelectorAll('.theme-card').forEach(card => {
        if (card.dataset.theme === theme) card.classList.add('active');
        else card.classList.remove('active');
    });
}

function updateEndingBuffer(days) {
    state.endingBuffer = parseInt(days) || 60;
    localStorage.setItem('gantt_ending_buffer', state.endingBuffer);
    saveToServer();
    render();
}

function saveTeam() {
    localStorage.setItem('gantt_team', JSON.stringify(state.team));
    saveToServer();
}

// Initialize
function initData() {
    const stored = localStorage.getItem('gantt_tasks');
    if (stored) {
        state.tasks = JSON.parse(stored);
        state.tasks.forEach(t => {
            t.startDate = new Date(t.startDate);
            t.endDate = new Date(t.endDate);
        });
    } else {
        const today = new Date();
        state.tasks = [
            {
                id: crypto.randomUUID(),
                name: 'Project Kickoff',
                owner: 'Alice',
                startDate: new Date(today),
                duration: 2,
                endDate: addDays(today, 2),
                color: '#3b82f6'
            },
            {
                id: crypto.randomUUID(),
                name: 'Design Phase',
                owner: 'Bob',
                startDate: addDays(today, 2),
                duration: 5,
                endDate: addDays(today, 7),
                color: '#8b5cf6'
            }
        ];
        saveTasks();
    }

    if (state.tasks.length > 0) {
        state.viewStartDate = state.tasks.reduce((min, t) => t.startDate < min ? t.startDate : min, state.tasks[0].startDate);
        // Optional: Subtract a small fixed buffer (e.g., 2 days) just for visual comfort, or 0 if strictly "no start buffer"
        state.viewStartDate = addDays(state.viewStartDate, -2);
    } else {
        state.viewStartDate = addDays(new Date(), -2);
    }
}

function saveTasks() {
    localStorage.setItem('gantt_tasks', JSON.stringify(state.tasks));
    saveToServer();
}

// --- Helper Functions ---
function getWeekNumber(d) {
    // Copy date so don't modify original
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    // Set to nearest Thursday: current date + 4 - current day number
    // Make Sunday's day number 7
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    // Get first day of year
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    // Calculate full weeks to nearest Thursday
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
}

function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}


function getDiffDays(start, end) {
    const oneDay = 24 * 60 * 60 * 1000;
    return Math.round((end - start) / oneDay);
}

function formatDate(date) {
    return date.toISOString().split('T')[0];
}
const elements = {
    headerTop: document.getElementById('gantt-header-top'),
    headerMiddle: document.getElementById('gantt-header-middle'),
    headerBottom: document.getElementById('gantt-header-bottom'),
    ganttBody: document.getElementById('gantt-body'),
    sidebarBody: document.getElementById('sidebar-body'),
    scaleBtns: document.querySelectorAll('.scale-btn'),
    modalOverlay: document.getElementById('modal-overlay'),
    btnAdd: document.getElementById('btn-add-task'),
    btnCloseModal: document.getElementById('btn-close-modal'),
    form: document.getElementById('add-task-form'),
    inpName: document.getElementById('task-name'),
    inpOwner: document.getElementById('task-owner'), // Now a SELECT
    inpStart: document.getElementById('task-start'),
    inpDuration: document.getElementById('task-duration'),
    modalTitle: document.getElementById('modal-title'),

    btnSaveTask: document.getElementById('btn-save-task'),
    btnDeleteTask: document.getElementById('btn-delete-task'),
    btnExportExcel: document.getElementById('btn-export-excel'),
    btnExportPDF: document.getElementById('btn-export-pdf'),

    // Team Settings
    tabBtns: document.querySelectorAll('.tab-btn'),
    tabContents: document.querySelectorAll('.tab-content'),
    teamList: document.getElementById('team-list'),
    newMemberName: document.getElementById('new-member-name'),
    newMemberEmail: document.getElementById('new-member-email'),
    btnAddMember: document.getElementById('btn-add-member'),
    btnAddCalendar: document.getElementById('btn-add-calendar')
};

if (!document.getElementById('modal-title')) {
    const h2 = document.querySelector('.modal-header h2');
    if (h2) h2.id = 'modal-title';
}
elements.modalTitle = document.getElementById('modal-title');


// --- Rendering ---

function updateScaleConfig() {
    // Calculate required timeline length based on tasks
    let maxTaskDate = new Date();
    if (state.tasks && state.tasks.length > 0) {
        maxTaskDate = new Date(state.tasks[0].endDate);
        state.tasks.forEach(t => {
            if (t.endDate > maxTaskDate) maxTaskDate = new Date(t.endDate);
        });
    } else {
        maxTaskDate = addDays(new Date(), 30);
    }

    // Determine days needed from view start to max task date
    // Note: viewStartDate is initialized in initData/render
    const viewStart = new Date(state.viewStartDate);
    const diffDays = getDiffDays(viewStart, maxTaskDate);
    // Add extra space for buffers that might extend past endDate
    const bufferDays = state.endingBuffer;
    const requiredDays = Math.max(0, diffDays) + bufferDays;

    if (state.viewScale === 'day') {
        state.columnWidth = 60;
        state.daysToRender = Math.max(60, requiredDays);
    } else if (state.viewScale === 'week') {
        state.columnWidth = 45; // Widen to fit "M/D"
        state.daysToRender = Math.max(90, requiredDays); // Show more days
    } else if (state.viewScale === 'month') {
        state.columnWidth = 100; // Per month column
        state.daysToRender = Math.max(730, requiredDays); // 2 years default or fit tasks
    }

    elements.ganttBody.style.backgroundSize = `${state.columnWidth}px 100%`;
}

function renderHeader() {
    elements.headerTop.innerHTML = '';
    elements.headerMiddle.innerHTML = '';
    elements.headerBottom.innerHTML = '';

    let totalWidth = 0;

    if (state.viewScale === 'week') {
        // ... Week View logic ...
        let currentYear = -1;
        let lastYearCell = null;
        let lastYearWidth = 0;

        let currentWeekKey = -1;
        let lastWeekCell = null;
        let lastWeekWidth = 0;

        let currentDate = new Date(state.viewStartDate);

        for (let i = 0; i < state.daysToRender; i++) {
            const dObj = new Date(currentDate);
            const day = dObj.getDay();
            const diff = dObj.getDate() - day + (day == 0 ? -6 : 1);
            const monday = new Date(dObj.setDate(diff));
            const weekKey = monday.toISOString().split('T')[0];

            if (weekKey !== currentWeekKey) {
                if (lastWeekCell) {
                    lastWeekCell.style.width = `${lastWeekWidth}px`;
                    elements.headerTop.appendChild(lastWeekCell);
                }
                lastWeekCell = document.createElement('div');
                lastWeekCell.className = 'timeline-cell group';

                const weekNum = getWeekNumber(monday);
                lastWeekCell.textContent = `${t('scale_week')} ${weekNum} - ${monday.toLocaleDateString(state.lang)}`;

                lastWeekWidth = 0;
                currentWeekKey = weekKey;
            }
            lastWeekWidth += state.columnWidth;

            const cell = document.createElement('div');
            cell.className = 'timeline-cell';
            cell.style.width = `${state.columnWidth}px`;
            const m = currentDate.getMonth() + 1;
            const dateNum = currentDate.getDate();
            cell.textContent = `${m}/${dateNum}`;
            if (currentDate.getDay() === 0 || currentDate.getDay() === 6) {
                cell.style.backgroundColor = "rgba(255,255,255,0.02)";
            }
            elements.headerMiddle.appendChild(cell);

            const y = currentDate.getFullYear();
            if (y !== currentYear) {
                if (lastYearCell) {
                    lastYearCell.style.width = `${lastYearWidth}px`;
                    elements.headerBottom.appendChild(lastYearCell);
                }
                lastYearCell = document.createElement('div');
                lastYearCell.className = 'timeline-cell group';
                lastYearCell.textContent = y;
                lastYearWidth = 0;
                currentYear = y;
            }
            lastYearWidth += state.columnWidth;

            totalWidth += state.columnWidth;
            currentDate.setDate(currentDate.getDate() + 1);
        }
        if (lastWeekCell) {
            lastWeekCell.style.width = `${lastWeekWidth}px`;
            elements.headerTop.appendChild(lastWeekCell);
        }
        if (lastYearCell) {
            lastYearCell.style.width = `${lastYearWidth}px`;
            elements.headerBottom.appendChild(lastYearCell);
        }

    } else if (state.viewScale === 'day') {
        let currentYear = -1;
        let lastYearCell = null;
        let lastYearWidth = 0;
        let currentMid = -1;
        let lastMidCell = null;
        let lastMidWidth = 0;
        let currentDate = new Date(state.viewStartDate);

        for (let i = 0; i < state.daysToRender; i++) {
            const cell = document.createElement('div');
            cell.className = 'timeline-cell';
            cell.style.width = `${state.columnWidth}px`;
            const d = currentDate.getDate();
            const w = currentDate.toLocaleString(state.lang, { weekday: 'narrow' });
            cell.innerHTML = `<div>${d}</div><div style='font-size:0.7em'>${w}</div>`;
            elements.headerTop.appendChild(cell);

            let midKey = `${currentDate.getFullYear()}-${currentDate.getMonth()}`;
            let midLabel = currentDate.toLocaleString(state.lang === 'en' ? 'default' : state.lang, { month: 'long' });
            if (midKey !== currentMid) {
                if (lastMidCell) {
                    lastMidCell.style.width = `${lastMidWidth}px`;
                    elements.headerMiddle.appendChild(lastMidCell);
                }
                lastMidCell = document.createElement('div');
                lastMidCell.className = 'timeline-cell group';
                lastMidCell.textContent = midLabel;
                lastMidWidth = 0;
                currentMid = midKey;
            }
            lastMidWidth += state.columnWidth;

            const y = currentDate.getFullYear();
            if (y !== currentYear) {
                if (lastYearCell) {
                    lastYearCell.style.width = `${lastYearWidth}px`;
                    elements.headerBottom.appendChild(lastYearCell);
                }
                lastYearCell = document.createElement('div');
                lastYearCell.className = 'timeline-cell group';
                lastYearCell.textContent = y;
                lastYearWidth = 0;
                currentYear = y;
            }
            lastYearWidth += state.columnWidth;
            totalWidth += state.columnWidth;
            currentDate.setDate(currentDate.getDate() + 1);
        }
        if (lastMidCell) {
            lastMidCell.style.width = `${lastMidWidth}px`;
            elements.headerMiddle.appendChild(lastMidCell);
        }
        if (lastYearCell) {
            lastYearCell.style.width = `${lastYearWidth}px`;
            elements.headerBottom.appendChild(lastYearCell);
        }
    } else {
        // Month View
        let currentDate = new Date(state.viewStartDate);
        currentDate.setDate(1);
        let currentYear = -1;
        let lastYearCell = null;
        let lastYearWidth = 0;
        let currentQ = -1;
        let lastQCell = null;
        let lastQWidth = 0;

        for (let i = 0; i < 24; i++) {
            const cell = document.createElement('div');
            cell.className = 'timeline-cell';
            cell.style.width = `${state.columnWidth}px`;
            cell.textContent = currentDate.toLocaleString(state.lang === 'en' ? 'default' : state.lang, { month: 'short' });
            elements.headerTop.appendChild(cell);

            const q = Math.floor(currentDate.getMonth() / 3) + 1;
            const qKey = `${currentDate.getFullYear()}-Q${q}`;
            if (qKey !== currentQ) {
                if (lastQCell) {
                    lastQCell.style.width = `${lastQWidth}px`;
                    elements.headerMiddle.appendChild(lastQCell);
                }
                lastQCell = document.createElement('div');
                lastQCell.className = 'timeline-cell group';
                lastQCell.textContent = `Q${q}`;
                lastQWidth = 0;
                currentQ = qKey;
            }
            lastQWidth += state.columnWidth;

            const y = currentDate.getFullYear();
            if (y !== currentYear) {
                if (lastYearCell) {
                    lastYearCell.style.width = `${lastYearWidth}px`;
                    elements.headerBottom.appendChild(lastYearCell);
                }
                lastYearCell = document.createElement('div');
                lastYearCell.className = 'timeline-cell group';
                lastYearCell.textContent = y;
                lastYearWidth = 0;
                currentYear = y;
            }
            lastYearWidth += state.columnWidth;

            totalWidth += state.columnWidth;
            currentDate.setMonth(currentDate.getMonth() + 1);
        }
        if (lastQCell) {
            lastQCell.style.width = `${lastQWidth}px`;
            elements.headerMiddle.appendChild(lastQCell);
        }
        if (lastYearCell) {
            lastYearCell.style.width = `${lastYearWidth}px`;
            elements.headerBottom.appendChild(lastYearCell);
        }
    }

    const widthPx = `${totalWidth}px`;
    elements.headerTop.style.width = widthPx;
    elements.headerMiddle.style.width = widthPx;
    elements.headerBottom.style.width = widthPx;
    elements.ganttBody.style.width = widthPx;
}

function renderSidebar() {
    elements.sidebarBody.innerHTML = '';
    const sortedTasks = [...state.tasks].sort((a, b) => a.startDate - b.startDate);

    sortedTasks.forEach((task, index) => {
        const row = document.createElement('div');
        row.className = 'sidebar-row';
        row.dataset.id = task.id;

        const cellIndex = document.createElement('div');
        cellIndex.className = 'sidebar-cell index-cell';
        cellIndex.textContent = index + 1;

        const cellTask = document.createElement('div');
        cellTask.className = 'sidebar-cell task-name';
        cellTask.textContent = task.name;

        const cellOwner = document.createElement('div');
        cellOwner.className = 'sidebar-cell task-owner';
        cellOwner.textContent = task.owner || '-';

        row.appendChild(cellIndex);
        row.appendChild(cellTask);
        row.appendChild(cellOwner);

        row.addEventListener('dblclick', () => openModal(task));

        elements.sidebarBody.appendChild(row);
    });
}

function renderTasks() {
    elements.ganttBody.innerHTML = '';
    const sortedTasks = [...state.tasks].sort((a, b) => a.startDate - b.startDate);

    sortedTasks.forEach(task => {
        const row = document.createElement('div');
        row.className = 'gantt-row';
        row.style.width = elements.ganttBody.style.width;

        let left = 0;
        let width = 0;

        if (state.viewScale === 'month') {
            const vStart = new Date(state.viewStartDate);
            vStart.setDate(1);
            const getMonthDiff = (d1, d2) => {
                let months = (d2.getFullYear() - d1.getFullYear()) * 12;
                months -= d1.getMonth();
                months += d2.getMonth();
                months += (d2.getDate() - 1) / 30;
                return months;
            }
            const mStart = getMonthDiff(vStart, task.startDate);
            const mDur = task.duration / 30;
            left = mStart * state.columnWidth;
            width = mDur * state.columnWidth;
        } else {
            const daysFromStart = getDiffDays(state.viewStartDate, task.startDate);
            left = daysFromStart * state.columnWidth;
            width = task.duration * state.columnWidth;
        }

        const bar = document.createElement('div');
        bar.className = 'task-bar';
        bar.style.left = `${left}px`;
        bar.style.width = `${Math.max(width, 2)}px`;
        bar.style.backgroundColor = task.color;
        bar.title = `${task.name} (${task.duration} days)`;

        bar.addEventListener('dblclick', (e) => {
            e.stopPropagation();
            openModal(task);
        });

        row.appendChild(bar);
        elements.ganttBody.appendChild(row);
    });
}

function renderTeamSettings() {
    if (!elements.teamList) return;
    elements.teamList.innerHTML = '';
    state.team.forEach((member, index) => {
        const name = typeof member === 'object' ? member.name : member;
        const email = (typeof member === 'object' && member.email) ? member.email : '';

        const li = document.createElement('li');
        li.className = 'team-item';
        li.innerHTML = `
            <div style="display: flex; flex-direction: column; gap: 2px;">
                <span style="font-weight: 500;">${name}</span>
                ${email ? `<span style="font-size: 0.8rem; color: var(--text-secondary);">${email}</span>` : ''}
            </div>
            <button class="btn-remove-member" data-index="${index}">&times;</button>
        `;
        elements.teamList.appendChild(li);
    });

    elements.teamList.querySelectorAll('.btn-remove-member').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.target.dataset.index);
            state.team.splice(idx, 1);
            saveTeam();
            renderTeamSettings();
            updateOwnerDropdown();
        });
    });
}

function updateOwnerDropdown() {
    if (!elements.inpOwner) return;
    elements.inpOwner.innerHTML = '';
    state.team.forEach(member => {
        const name = typeof member === 'object' ? member.name : member;
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        elements.inpOwner.appendChild(option);
    });
}

function render() {
    // Re-calculate view start date based on current tasks
    if (state.tasks.length > 0) {
        const minDate = state.tasks.reduce((min, t) => t.startDate < min ? t.startDate : min, state.tasks[0].startDate);
        state.viewStartDate = addDays(minDate, -2);
    } else {
        // If no tasks, anchor to today minus small fixed buffer
        state.viewStartDate = addDays(new Date(), -2);
    }

    updateScaleConfig();
    renderHeader();
    renderSidebar();
    renderTasks();
}

// --- Export Functions ---
function getExportFilename() {
    const rawName = document.getElementById('project-name').value || "Untitled_Project";
    return rawName.replace(/[^a-z0-9_\-\s\u4e00-\u9fa5]/gi, '').replace(/\s+/g, '_');
}

function exportToExcel() {
    // 1. Sort tasks by Start Date (same as UI)
    const sortedTasks = [...state.tasks].sort((a, b) => a.startDate - b.startDate);

    if (sortedTasks.length === 0) {
        alert("No tasks to export.");
        return;
    }

    // 2. Determine Timeline Range
    // Find min start and max end to define the timeline range
    let minDate = sortedTasks[0].startDate;
    let maxDate = sortedTasks[0].endDate;
    sortedTasks.forEach(t => {
        if (t.startDate < minDate) minDate = t.startDate;
        if (t.endDate > maxDate) maxDate = t.endDate;
    });
    // Add buffer
    minDate = addDays(minDate, -2);
    maxDate = addDays(maxDate, 5);

    // Generate dates array for columns
    const dates = [];
    let curr = new Date(minDate);
    while (curr <= maxDate) {
        dates.push(new Date(curr));
        curr.setDate(curr.getDate() + 1);
    }

    // 3. Build HTML Table String
    let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <!--[if gte mso 9]>
            <xml>
                <x:ExcelWorkbook>
                    <x:ExcelWorksheets>
                        <x:ExcelWorksheet>
                            <x:Name>Gantt Project</x:Name>
                            <x:WorksheetOptions>
                                <x:DisplayGridlines/>
                            </x:WorksheetOptions>
                        </x:ExcelWorksheet>
                    </x:ExcelWorksheets>
                </x:ExcelWorkbook>
            </xml>
            <![endif]-->
            <meta charset="UTF-8">
            <style>
                table { border-collapse: collapse; }
                th, td { border: 1px solid #ccc; padding: 5px; text-align: center; font-family: Arial, sans-serif; font-size: 10pt; }
                th { background-color: #f0f0f0; font-weight: bold; }
                .task-header { background-color: #e0e0e0; text-align: left; }
                .timeline-header { background-color: #f8f8f8; }
                .timeline-cell { width: 30px; }
            </style>
        </head>
        <body>
        <table>
    `;

    // Header Row 1: Fixed Columns + Month/Year Labels
    html += `<tr>
        <th rowspan="2" style="width: 40px;">#</th>
        <th rowspan="2" style="width: 200px;">${t("col_task_name")}</th>
        <th rowspan="2" style="width: 100px;">${t("col_owner")}</th>
        <th rowspan="2" style="width: 90px;">${t("col_start_date")}</th>
        <th rowspan="2" style="width: 90px;">${t("col_end_date")}</th>
        <th rowspan="2" style="width: 60px;">${t("col_duration")}</th>
    `;

    // Timeline Month Headers
    let currentMonthLabel = "";
    let colspan = 0;

    dates.forEach((d, i) => {
        const mLabel = d.toLocaleString(state.lang === 'en' ? 'default' : state.lang, { month: 'short', year: 'numeric' });
        if (mLabel !== currentMonthLabel) {
            if (currentMonthLabel !== "") {
                html += `<th colspan="${colspan}">${currentMonthLabel}</th>`;
            }
            currentMonthLabel = mLabel;
            colspan = 1;
        } else {
            colspan++;
        }
        // Last one
        if (i === dates.length - 1) {
            html += `<th colspan="${colspan}">${currentMonthLabel}</th>`;
        }
    });

    html += `</tr>`;

    // Header Row 2: Days
    html += `<tr>`;
    dates.forEach(d => {
        const day = d.getDate();
        html += `<th class="timeline-cell">${day}</th>`;
    });
    html += `</tr>`;

    // 4. Data Rows
    sortedTasks.forEach((task, index) => {
        html += `<tr>
            <td>${index + 1}</td>
            <td style="text-align: left;">${task.name}</td>
            <td>${task.owner || '-'}</td>
            <td>${formatDate(task.startDate)}</td>
            <td>${formatDate(task.endDate)}</td>
            <td>${task.duration}</td>
        `;

        // Timeline Cells
        dates.forEach(d => {
            // Check if date is within task range
            // Reset times to 00:00:00 for accurate comparison
            const checkTime = d.getTime();
            const sTime = new Date(task.startDate).setHours(0, 0, 0, 0);
            const eTime = new Date(task.endDate).setHours(0, 0, 0, 0);

            // Note: task.endDate is typically exclusive or inclusive depending on logic.
            // In our data, endDate = startDate + duration. 
            // So visually [startDate, endDate) usually. 
            // But let's check our render logic: width = duration * colWidth.
            // So if duration is 1 day, it covers 1 day (startDate).
            // Logic: checkTime >= sTime && checkTime < eTime

            if (checkTime >= sTime && checkTime < eTime) {
                html += `<td style="background-color: ${task.color};"></td>`;
            } else {
                html += `<td></td>`;
            }
        });

        html += `</tr>`;
    });

    html += `
        </table>
        </body>
        </html>
    `;

    // 5. Trigger Download
    const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${getExportFilename()}_${formatDate(new Date())}.xls`; // Note .xls extension
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function exportToPDF() {
    const element = document.querySelector('.split-view'); // Capture visible area
    const fname = getExportFilename();
    const opt = {
        margin: 10,
        filename: `${fname}_${formatDate(new Date())}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
    };
    html2pdf().set(opt).from(element).save();
}

function generateCalendarUrl(task) {
    // Format YYYYMMDD
    const formatDateICal = (d) => d.toISOString().replace(/-|:|\.\d\d\d/g, "").substring(0, 8);

    // Google Calendar expects UTC or local formatted string. 
    // Since task dates are objects, we can convert.
    // However, All-day events in GCal typically work best with YYYYMMDD.
    // End date is exclusive in GCal for all-day events, which matches our system.

    const start = formatDateICal(task.startDate);
    const end = formatDateICal(task.endDate);

    const title = encodeURIComponent(task.name);
    const details = encodeURIComponent(`Task in GanttFlow.\nOwner: ${task.owner}`);

    return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${title}&dates=${start}/${end}&details=${details}`;
}

// --- Event Listeners ---
if (elements.ganttBody && elements.sidebarBody) {
    elements.ganttBody.addEventListener('scroll', () => {
        elements.sidebarBody.scrollTop = elements.ganttBody.scrollTop;
        const scrollLeft = elements.ganttBody.scrollLeft;
        elements.headerTop.style.transform = `translateX(-${scrollLeft}px)`;
        elements.headerMiddle.style.transform = `translateX(-${scrollLeft}px)`;
        elements.headerBottom.style.transform = `translateX(-${scrollLeft}px)`;
    });

    elements.sidebarBody.addEventListener('scroll', () => {
        elements.ganttBody.scrollTop = elements.sidebarBody.scrollTop;
    });
}

const langSelect = document.getElementById('lang-select');
if (langSelect) {
    langSelect.value = state.lang;
    langSelect.addEventListener('change', (e) => {
        updateLanguage(e.target.value);
    });
}

elements.scaleBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        elements.scaleBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.viewScale = btn.dataset.scale;
        render();
    });
});

// Modal Logic
function openModal(taskToEdit = null) {
    elements.modalOverlay.classList.add('open');
    updateOwnerDropdown();

    if (taskToEdit) {
        state.editingTaskId = taskToEdit.id;
        elements.modalTitle.setAttribute('data-i18n', 'modal_title_edit');
        elements.modalTitle.textContent = t('modal_title_edit');

        elements.inpName.value = taskToEdit.name;

        const ownerName = taskToEdit.owner;
        const teamNames = state.team.map(m => typeof m === 'object' ? m.name : m);

        if (teamNames.includes(ownerName)) {
            elements.inpOwner.value = ownerName;
        } else {
            const opt = document.createElement('option');
            opt.value = ownerName;
            opt.textContent = ownerName + " (Unknown)";
            opt.selected = true;
            elements.inpOwner.appendChild(opt);
        }

        elements.inpStart.valueAsDate = taskToEdit.startDate;
        elements.inpDuration.value = taskToEdit.duration;

        const radio = document.querySelector(`input[name="task-color"][value="${taskToEdit.color}"]`);
        if (radio) radio.checked = true;

        elements.btnSaveTask.textContent = t('btn_save');
        elements.btnSaveTask.setAttribute('data-i18n', 'btn_save');
        elements.btnSaveTask.style.background = taskToEdit.color;
        elements.btnSaveTask.style.boxShadow = `0 4px 12px ${taskToEdit.color}80`;
        elements.btnDeleteTask.style.display = 'block';

        // Calendar Button
        elements.btnAddCalendar.style.display = 'flex';
        elements.btnAddCalendar.onclick = () => {
            const url = generateCalendarUrl(taskToEdit);
            window.open(url, '_blank');
        };

    } else {
        state.editingTaskId = null;
        elements.modalTitle.setAttribute('data-i18n', 'modal_title_add');
        elements.modalTitle.textContent = t('modal_title_add');

        elements.form.reset();
        elements.inpStart.valueAsDate = new Date();

        const defaultColor = '#3b82f6';
        elements.btnSaveTask.style.background = defaultColor;
        elements.btnSaveTask.style.boxShadow = `0 4px 12px ${defaultColor}80`;

        if (state.team.length > 0) {
            const m = state.team[0];
            elements.inpOwner.value = typeof m === 'object' ? m.name : m;
        }

        elements.btnSaveTask.textContent = t('btn_create');
        elements.btnSaveTask.setAttribute('data-i18n', 'btn_create');

        elements.btnDeleteTask.style.display = 'none';
        elements.btnAddCalendar.style.display = 'none';
    }
}

function closeModal() {
    elements.modalOverlay.classList.remove('open');
    state.editingTaskId = null;
    elements.form.reset();
}

if (elements.btnAdd) elements.btnAdd.addEventListener('click', () => openModal());
if (elements.btnCloseModal) elements.btnCloseModal.addEventListener('click', closeModal);
if (elements.modalOverlay) elements.modalOverlay.addEventListener('click', (e) => {
    if (e.target === elements.modalOverlay) closeModal();
});

if (elements.btnDeleteTask) {
    elements.btnDeleteTask.addEventListener('click', () => {
        if (state.editingTaskId && confirm("Delete this task?")) {
            state.tasks = state.tasks.filter(t => t.id !== state.editingTaskId);
            saveTasks();
            render();
            closeModal();
        }
    });
}

elements.form.addEventListener('submit', (e) => {
    e.preventDefault();
    const name = elements.inpName.value;
    const owner = elements.inpOwner.value;
    const startDate = elements.inpStart.valueAsDate;
    const duration = parseInt(elements.inpDuration.value, 10);
    const color = document.querySelector('input[name="task-color"]:checked').value;

    if (!name || !startDate || !duration) return;

    if (state.editingTaskId) {
        const taskIndex = state.tasks.findIndex(t => t.id === state.editingTaskId);
        if (taskIndex !== -1) {
            state.tasks[taskIndex] = {
                ...state.tasks[taskIndex],
                name,
                owner,
                startDate,
                duration,
                endDate: addDays(startDate, duration),
                color
            };
        }
    } else {
        const newTask = {
            id: crypto.randomUUID(),
            name,
            owner,
            startDate,
            duration,
            endDate: addDays(startDate, duration),
            color
        };
        state.tasks.push(newTask);
    }

    saveTasks();
    render();
    closeModal();
});

// Color Selection Logic
const colorRadios = document.querySelectorAll('input[name="task-color"]');
colorRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
        const color = e.target.value;
        if (elements.btnSaveTask) {
            elements.btnSaveTask.style.background = color;
            elements.btnSaveTask.style.boxShadow = `0 4px 12px ${color}80`;
        }
    });
});


if (elements.btnExportExcel) elements.btnExportExcel.addEventListener('click', exportToExcel);
if (elements.btnExportPDF) elements.btnExportPDF.addEventListener('click', exportToPDF);

const projectNameInput = document.getElementById('project-name');
if (projectNameInput) {
    const savedName = localStorage.getItem('gantt_project_name');
    if (savedName) projectNameInput.value = savedName;
    else projectNameInput.value = "My Project";
    projectNameInput.addEventListener('input', (e) => {
        localStorage.setItem('gantt_project_name', e.target.value);
    });
}

const btnSettings = document.getElementById('btn-settings');
const settingsModal = document.getElementById('settings-modal-overlay');
const btnCloseSettings = document.getElementById('btn-close-settings');
const themeCards = document.querySelectorAll('.theme-card');

const endingBufferInput = document.getElementById('ending-buffer-input');

if (endingBufferInput) {
    endingBufferInput.value = state.endingBuffer;
    endingBufferInput.addEventListener('change', (e) => {
        updateEndingBuffer(e.target.value);
    });
}

if (btnSettings) {
    btnSettings.addEventListener('click', () => {
        settingsModal.classList.add('open');
        renderTeamSettings();
        if (endingBufferInput) endingBufferInput.value = state.endingBuffer;
    });
}

const btnSettingsConfirm = document.getElementById('btn-settings-confirm');
if (btnSettingsConfirm) {
    btnSettingsConfirm.addEventListener('click', () => {
        settingsModal.classList.remove('open');
    });
}

if (btnCloseSettings) {
    btnCloseSettings.addEventListener('click', () => {
        settingsModal.classList.remove('open');
    });
}

if (settingsModal) {
    settingsModal.addEventListener('click', (e) => {
        if (e.target === settingsModal) settingsModal.classList.remove('open');
    });
}

themeCards.forEach(card => {
    card.addEventListener('click', () => {
        const theme = card.dataset.theme;
        updateTheme(theme);
    });
});

elements.tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        elements.tabBtns.forEach(b => b.classList.remove('active'));
        elements.tabContents.forEach(c => c.classList.remove('active'));

        btn.classList.add('active');
        const tabId = btn.dataset.tab;
        document.getElementById(`tab-${tabId}`).classList.add('active');
    });
});

if (elements.btnAddMember) {
    elements.btnAddMember.addEventListener('click', () => {
        const name = elements.newMemberName.value.trim();
        const email = elements.newMemberEmail ? elements.newMemberEmail.value.trim() : '';

        if (name) {
            const exists = state.team.some(m => (typeof m === 'object' ? m.name : m) === name);
            if (!exists) {
                state.team.push({ name, email });
                saveTeam();
                renderTeamSettings();
                updateOwnerDropdown();
                elements.newMemberName.value = '';
                if (elements.newMemberEmail) elements.newMemberEmail.value = '';
            } else {
                alert("Team member with this name already exists.");
            }
        }
    });
}

// --- Server Persistence ---
async function loadFromServer() {
    try {
        const response = await fetch(`/api/data?source=${state.dataSource}`);
        if (!response.ok) throw new Error('Server not available');
        const data = await response.json();
        // Always reset tasks/team/etc from the loaded data, or empty if new source
        state.tasks = data.tasks || [];
        state.team = data.team || loadTeam(); // Fallback to local default if missing

        // Settings could be synced or kept local. Let's sync them if they exist in valid data
        if (data.lang) state.lang = data.lang;
        if (data.theme) state.theme = data.theme;
        if (data.endingBuffer) state.endingBuffer = data.endingBuffer;

        // Update LocalStorage to match what we just loaded, effectively caching it
        localStorage.setItem('gantt_tasks', JSON.stringify(state.tasks));
        localStorage.setItem('gantt_team', JSON.stringify(state.team));
        localStorage.setItem('gantt_lang', state.lang);
        localStorage.setItem('gantt_theme', state.theme);
        localStorage.setItem('gantt_ending_buffer', state.endingBuffer);

        return true;
    } catch (e) {
        // console.log('Local server not found or error, using localStorage');
    }
    return false;
}

async function saveToServer() {
    const data = {
        tasks: state.tasks,
        team: state.team,
        lang: state.lang,
        theme: state.theme,
        endingBuffer: state.endingBuffer
    };
    try {
        await fetch(`/api/data?source=${state.dataSource}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
    } catch (e) {
        // console.log('Failed to save to server');
    }
}

// Data Source UI Binding
const dataSourceSelect = document.getElementById('data-source-select');
if (dataSourceSelect) {
    dataSourceSelect.value = state.dataSource;
    dataSourceSelect.addEventListener('change', async (e) => {
        const newSource = e.target.value;
        state.dataSource = newSource;
        localStorage.setItem('gantt_data_source', newSource);

        // Reload data from the new source
        const success = await loadFromServer();
        if (success) {
            // Re-render everything
            initData(); // Re-calc viewStartDate checks
            updateLanguage(state.lang);
            updateTheme(state.theme);
            renderTeamSettings();
            updateOwnerDropdown();
            render();
            alert(`Switched to ${newSource} storage.`);
        } else {
            alert("Failed to load data from server. Is it running?");
        }
    });
}

loadFromServer().then(() => {
    initData();
    updateLanguage(state.lang);
    updateTheme(state.theme);
});

function updateSummaryStats() {
    const totalTasks = state.tasks.length;
    const totalDays = state.tasks.reduce((sum, t) => sum + parseInt(t.duration || 0), 0);
    const teamSize = state.team.length;

    let longestTaskName = "-";
    if (state.tasks.length > 0) {
        const longest = state.tasks.reduce((prev, current) => (prev.duration > current.duration) ? prev : current);
        longestTaskName = `${longest.name} (${longest.duration}d)`;
    }

    document.getElementById("summary-total-tasks").textContent = totalTasks;
    document.getElementById("summary-total-days").textContent = totalDays;
    document.getElementById("summary-team-size").textContent = teamSize;
    document.getElementById("summary-longest-task").textContent = longestTaskName;
}

// Bind to Summary Tab
document.querySelector(".tab-btn[data-tab=\"summary\"]").addEventListener("click", () => {
    updateSummaryStats();
});

