/* global XLSX, Chart */
(() => {
  const STORAGE_KEY = "agenda_semanal_v1";

  const DAYS = [
    { name: "Lunes", idx: 0 },
    { name: "Martes", idx: 1 },
    { name: "Miércoles", idx: 2 },
    { name: "Jueves", idx: 3 },
    { name: "Viernes", idx: 4 },
  ];

  const STATUS = ["Pendiente", "En progreso", "Hecho"];
  const PRIORITIES = ["Alta", "Media", "Baja"];

  // Colores para departamentos (clásicos)
  const DEPT_COLORS = {
    "RH": "#3b82f6",      // Azul
    "IT": "#10b981",      // Verde
    "Ventas": "#f59e0b",  // Amarillo/Naranja
    "Marketing": "#ef4444", // Rojo
    "Finanzas": "#8b5cf6", // Morado
    "Operaciones": "#06b6d4", // Cian
    "Legal": "#f97316",   // Naranja
    "Compras": "#14b8a6", // Turquesa
  };

  // Colores para prioridades
  const PRIORITY_COLORS = {
    "Alta": "#ef4444",    // Rojo
    "Media": "#f59e0b",   // Amarillo/Naranja
    "Baja": "#10b981",    // Verde
  };

  // Colores para estados (progreso)
  const STATUS_COLORS = {
    "Pendiente": "#ef4444",      // Rojo
    "En progreso": "#f59e0b",    // Amarillo/Naranja
    "Hecho": "#10b981",          // Verde
  };

  /** @type {{tasks: any[], weekStartISO: string}} */
  let state = { tasks: [], weekStartISO: "" };
  let chartWeek = null;

  const el = {
    weekStart: document.getElementById("weekStart"),
    weekLabel: document.getElementById("weekLabel"),
    weekGrid: document.getElementById("weekGrid"),
    excelFile: document.getElementById("excelFile"),
    btnNewTask: document.getElementById("btnNewTask"),
    btnDownload: document.getElementById("btnDownload"),
    btnTodayWeek: document.getElementById("btnTodayWeek"),
    btnClearWeek: document.getElementById("btnClearWeek"),
    btnLoadSample: document.getElementById("btnLoadSample"),
    btnTemplate: document.getElementById("btnTemplate"),
    filterDept: document.getElementById("filterDept"),
    filterOwner: document.getElementById("filterOwner"),
    filterStatus: document.getElementById("filterStatus"),
    weekProgressText: document.getElementById("weekProgressText"),
    weekTasksText: document.getElementById("weekTasksText"),
    chartCanvas: document.getElementById("chartWeek"),
    modal: document.getElementById("taskModal"),
    form: document.getElementById("taskForm"),
    modalTitle: document.getElementById("modalTitle"),
    taskId: document.getElementById("taskId"),
    taskDay: document.getElementById("taskDay"),
    taskDate: document.getElementById("taskDate"),
    taskTitle: document.getElementById("taskTitle"),
    taskDept: document.getElementById("taskDept"),
    taskOwner: document.getElementById("taskOwner"),
    taskStatus: document.getElementById("taskStatus"),
    taskProgress: document.getElementById("taskProgress"),
    taskPriority: document.getElementById("taskPriority"),
    btnDeleteTask: document.getElementById("btnDeleteTask"),
    tasksList: document.getElementById("tasksList"),
  };

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function toISODate(d) {
    const y = d.getFullYear();
    const m = pad2(d.getMonth() + 1);
    const day = pad2(d.getDate());
    return `${y}-${m}-${day}`;
  }

  function fromISODate(iso) {
    // ISO yyyy-mm-dd -> local Date at midnight
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(y, m - 1, d);
  }

  function startOfWeekMonday(d) {
    const date = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const day = date.getDay(); // 0 Sun .. 6 Sat
    const diff = (day === 0 ? -6 : 1) - day; // shift to Monday
    date.setDate(date.getDate() + diff);
    return date;
  }

  function addDays(isoStart, days) {
    const d = fromISODate(isoStart);
    d.setDate(d.getDate() + days);
    return toISODate(d);
  }

  function clampProgress(p) {
    const n = Number(p);
    if (Number.isNaN(n)) return 0;
    return Math.min(100, Math.max(0, Math.round(n)));
  }

  function normalizeStatus(s, progress) {
    const val = (s ?? "").toString().trim();
    if (STATUS.includes(val)) return val;
    if (progress >= 100) return "Hecho";
    if (progress > 0) return "En progreso";
    return "Pendiente";
  }

  function uid() {
    return `t_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
  }

  function loadState() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        state.tasks = Array.isArray(parsed.tasks) ? parsed.tasks : [];
        state.weekStartISO = typeof parsed.weekStartISO === "string" ? parsed.weekStartISO : "";
      }
    } catch {
      // ignore
    }
  }

  function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function getWeekStartISO() {
    if (el.weekStart.value) return el.weekStart.value;
    const monday = startOfWeekMonday(new Date());
    return toISODate(monday);
  }

  function setWeekStartISO(iso) {
    el.weekStart.value = iso;
    state.weekStartISO = iso;
    saveState();
  }

  function getTaskWeekISO(task) {
    const d = fromISODate(task.dateISO);
    return toISODate(startOfWeekMonday(d));
  }

  function getWeekTasks() {
    const weekStart = getWeekStartISO();
    const dept = el.filterDept.value;
    const owner = el.filterOwner.value;
    const status = el.filterStatus.value;

    return state.tasks
      .filter((t) => getTaskWeekISO(t) === weekStart)
      .filter((t) => (dept ? t.dept === dept : true))
      .filter((t) => (owner ? t.owner === owner : true))
      .filter((t) => (status ? t.status === status : true))
      .sort((a, b) => (a.dateISO === b.dateISO ? a.title.localeCompare(b.title) : a.dateISO.localeCompare(b.dateISO)));
  }

  function computeDayStats(weekStartISO) {
    const weekTasks = state.tasks.filter((t) => getTaskWeekISO(t) === weekStartISO);
    const byDay = new Map(DAYS.map((d) => [d.idx, []]));
    for (const t of weekTasks) {
      const idx = dayIndexFromISODate(weekStartISO, t.dateISO);
      if (idx >= 0 && idx <= 4) byDay.get(idx).push(t);
    }

    const stats = DAYS.map((d) => {
      const list = byDay.get(d.idx);
      const total = list.length;
      const avg = total ? Math.round(list.reduce((s, t) => s + clampProgress(t.progress), 0) / total) : 0;
      const done = list.filter((t) => clampProgress(t.progress) >= 100 || t.status === "Hecho").length;
      return { dayIdx: d.idx, total, avg, done };
    });

    const weekTotal = weekTasks.length;
    const weekAvg = weekTotal ? Math.round(weekTasks.reduce((s, t) => s + clampProgress(t.progress), 0) / weekTotal) : 0;
    return { stats, weekTotal, weekAvg };
  }

  function dayIndexFromISODate(weekStartISO, isoDate) {
    const ws = fromISODate(weekStartISO);
    const d = fromISODate(isoDate);
    const diffMs = d.getTime() - ws.getTime();
    const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
    return diffDays;
  }

  function formatWeekLabel(weekStartISO) {
    const start = fromISODate(weekStartISO);
    const end = fromISODate(addDays(weekStartISO, 4));
    const fmt = (x) => `${pad2(x.getDate())}/${pad2(x.getMonth() + 1)}/${x.getFullYear()}`;
    return `Semana: ${fmt(start)} – ${fmt(end)}`;
  }

  function rebuildFilters() {
    const weekStart = getWeekStartISO();
    const weekTasks = state.tasks.filter((t) => getTaskWeekISO(t) === weekStart);
    const depts = [...new Set(weekTasks.map((t) => t.dept).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    const owners = [...new Set(weekTasks.map((t) => t.owner).filter(Boolean))].sort((a, b) => a.localeCompare(b));

    const prevDept = el.filterDept.value;
    const prevOwner = el.filterOwner.value;

    el.filterDept.innerHTML = `<option value="">Todos</option>${depts.map((d) => `<option>${escapeHtml(d)}</option>`).join("")}`;
    el.filterOwner.innerHTML = `<option value="">Todos</option>${owners.map((o) => `<option>${escapeHtml(o)}</option>`).join("")}`;

    if (depts.includes(prevDept)) el.filterDept.value = prevDept;
    if (owners.includes(prevOwner)) el.filterOwner.value = prevOwner;
  }

  function escapeHtml(s) {
    return String(s)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function getDeptColor(dept) {
    if (!dept) return "#6b7280";
    const normalized = dept.trim();
    return DEPT_COLORS[normalized] || DEPT_COLORS[Object.keys(DEPT_COLORS)[Math.abs(normalized.split("").reduce((a, b) => a + b.charCodeAt(0), 0)) % Object.keys(DEPT_COLORS).length]];
  }

  function getPriorityColor(priority) {
    return PRIORITY_COLORS[priority] || PRIORITY_COLORS["Media"];
  }

  function getStatusColor(status, progress) {
    const p = clampProgress(progress);
    if (status === "Hecho" || p >= 100) return STATUS_COLORS["Hecho"];
    if (status === "En progreso" || p > 0) return STATUS_COLORS["En progreso"];
    return STATUS_COLORS["Pendiente"];
  }

  function statusDot(task) {
    const p = clampProgress(task.progress);
    if (task.status === "Hecho" || p >= 100) return "dot dot--ok";
    if (task.status === "Pendiente" && p === 0) return "dot dot--danger";
    return "dot dot--warn";
  }

  function render() {
    const weekStart = getWeekStartISO();
    el.weekLabel.textContent = formatWeekLabel(weekStart);

    rebuildFilters();

    const tasks = getWeekTasks();
    const tasksByDay = new Map(DAYS.map((d) => [d.idx, []]));
    for (const t of tasks) {
      const idx = dayIndexFromISODate(weekStart, t.dateISO);
      if (idx >= 0 && idx <= 4) tasksByDay.get(idx).push(t);
    }

    const cols = DAYS.map((d) => {
      const dayISO = addDays(weekStart, d.idx);
      const list = tasksByDay.get(d.idx);
      const avg = list.length ? Math.round(list.reduce((s, t) => s + clampProgress(t.progress), 0) / list.length) : 0;
      const dotClass = avg >= 90 ? "dot dot--ok" : avg === 0 ? "dot dot--danger" : "dot";
      const body = list.length
        ? list
            .map((t) => {
              const p = clampProgress(t.progress);
              const deptColor = getDeptColor(t.dept);
              const priorityColor = getPriorityColor(t.priority || "Media");
              const statusColor = getStatusColor(t.status, p);
              return `
                <article class="task" data-id="${t.id}">
                  <div class="task__title">${escapeHtml(t.title)}</div>
                  <div class="task__meta">
                    <span class="tag" style="border-left: 3px solid ${statusColor}"><span class="${statusDot(t)}"></span>${escapeHtml(t.status)}</span>
                    <span class="tag" style="border-left: 3px solid ${deptColor}">Depto: <b>${escapeHtml(t.dept)}</b></span>
                    <span class="tag" style="border-left: 3px solid ${priorityColor}">Prioridad: <b>${escapeHtml(t.priority || "Media")}</b></span>
                    <span class="tag">Resp: <b>${escapeHtml(t.owner)}</b></span>
                  </div>
                  <div class="task__foot">
                    <div class="bar" aria-label="Progreso"><div style="width:${p}%; background: ${statusColor}"></div></div>
                    <div class="pct">${p}%</div>
                    <div class="task__actions">
                      <button class="linkBtn" type="button" data-action="edit">Editar</button>
                    </div>
                  </div>
                </article>
              `;
            })
            .join("")
        : `<div class="empty">Sin tareas para este día.</div>`;

      return `
        <div class="dayCol">
          <div class="dayCol__head">
            <div class="dayCol__name">${d.name}</div>
            <div class="dayCol__date">${dayISO}</div>
            <div class="dayCol__meta">
              <span class="pill"><span class="${dotClass}"></span>${avg}%</span>
              <span class="pill">${list.length} tarea(s)</span>
            </div>
          </div>
          <div class="dayCol__body">
            ${body}
          </div>
        </div>
      `;
    }).join("");

    el.weekGrid.innerHTML = cols;

    // Update summary + chart (use unfiltered tasks for a stable weekly dashboard)
    const { stats, weekTotal, weekAvg } = computeDayStats(weekStart);
    el.weekProgressText.textContent = `${weekAvg}%`;
    el.weekTasksText.textContent = String(weekTotal);
    renderChart(stats);
    renderTasksList();
  }

  function renderChart(stats) {
    if (!window.Chart) return;
    const labels = DAYS.map((d) => d.name);
    const data = stats.map((s) => s.avg);
    const totals = stats.map((s) => s.total);

    // Colores para la gráfica
    const progressColors = data.map((val) => {
      if (val >= 90) return "rgba(16, 185, 129, 0.7)"; // Verde
      if (val >= 50) return "rgba(245, 158, 11, 0.7)"; // Amarillo/Naranja
      return "rgba(239, 68, 68, 0.7)"; // Rojo
    });

    const cfg = {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Progreso promedio (%)",
            data,
            backgroundColor: progressColors,
            borderColor: progressColors.map(c => c.replace("0.7", "1")),
            borderWidth: 2,
            borderRadius: 8,
          },
          {
            label: "Tareas (conteo)",
            data: totals,
            backgroundColor: "rgba(59, 130, 246, 0.5)", // Azul
            borderColor: "rgba(59, 130, 246, 1)",
            borderWidth: 2,
            borderRadius: 8,
            yAxisID: "y2",
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: getComputedStyle(document.documentElement).getPropertyValue("--muted") } },
          tooltip: { mode: "index", intersect: false },
        },
        interaction: { mode: "index", intersect: false },
        scales: {
          y: {
            min: 0,
            max: 100,
            ticks: { color: getComputedStyle(document.documentElement).getPropertyValue("--muted") },
            grid: { color: "rgba(17,24,39,.10)" },
          },
          y2: {
            position: "right",
            beginAtZero: true,
            ticks: { color: getComputedStyle(document.documentElement).getPropertyValue("--muted") },
            grid: { drawOnChartArea: false },
          },
          x: {
            ticks: { color: getComputedStyle(document.documentElement).getPropertyValue("--muted") },
            grid: { display: false },
          },
        },
      },
    };

    if (chartWeek) {
      chartWeek.data.labels = cfg.data.labels;
      chartWeek.data.datasets = cfg.data.datasets;
      chartWeek.update();
      return;
    }

    chartWeek = new Chart(el.chartCanvas, cfg);
  }

  function renderTasksList() {
    if (!el.tasksList) return;
    const weekStart = getWeekStartISO();
    // Usar las mismas tareas filtradas que el tablero
    const weekTasks = getWeekTasks()
      .slice()
      .sort((a, b) => {
        // Ordenar por estado (Hecho al final), luego por prioridad, luego por fecha
        const statusOrder = { "Pendiente": 0, "En progreso": 1, "Hecho": 2 };
        const priorityOrder = { "Alta": 0, "Media": 1, "Baja": 2 };
        const statusDiff = (statusOrder[a.status] || 2) - (statusOrder[b.status] || 2);
        if (statusDiff !== 0) return statusDiff;
        const priorityDiff = (priorityOrder[a.priority] || 1) - (priorityOrder[b.priority] || 1);
        if (priorityDiff !== 0) return priorityDiff;
        return a.dateISO.localeCompare(b.dateISO);
      });

    const byStatus = {
      "Pendiente": weekTasks.filter((t) => t.status === "Pendiente" || (clampProgress(t.progress) === 0 && t.status !== "Hecho")),
      "En progreso": weekTasks.filter((t) => t.status === "En progreso" || (clampProgress(t.progress) > 0 && clampProgress(t.progress) < 100)),
      "Hecho": weekTasks.filter((t) => t.status === "Hecho" || clampProgress(t.progress) >= 100),
    };

    const sections = STATUS.map((status) => {
      const tasks = byStatus[status];
      if (!tasks.length) return "";

      const items = tasks.map((t) => {
        const p = clampProgress(t.progress);
        const deptColor = getDeptColor(t.dept);
        const priorityColor = getPriorityColor(t.priority || "Media");
        const statusColor = getStatusColor(t.status, p);
        const dayName = dayNameFromISO(t.dateISO);

        return `
          <div class="taskListItem" data-id="${t.id}">
            <div class="taskListItem__content">
              <div class="taskListItem__title">${escapeHtml(t.title)}</div>
              <div class="taskListItem__meta">
                <span class="taskListItem__tag" style="border-left: 3px solid ${deptColor}">
                  <b>${escapeHtml(t.dept)}</b>
                </span>
                <span class="taskListItem__tag" style="border-left: 3px solid ${priorityColor}">
                  <b>${escapeHtml(t.priority || "Media")}</b>
                </span>
                <span class="taskListItem__tag">
                  ${escapeHtml(t.owner)}
                </span>
                <span class="taskListItem__tag">
                  ${dayName} ${t.dateISO}
                </span>
              </div>
              <div class="taskListItem__progress">
                <div class="bar" aria-label="Progreso"><div style="width:${p}%; background: ${statusColor}"></div></div>
                <span class="pct">${p}%</span>
              </div>
            </div>
            <button class="linkBtn" type="button" data-action="edit">Editar</button>
          </div>
        `;
      }).join("");

      return `
        <div class="tasksList__section">
          <div class="tasksList__header" style="border-left: 4px solid ${STATUS_COLORS[status]}">
            <h3 class="tasksList__title">${status}</h3>
            <span class="tasksList__count">${tasks.length} tarea(s)</span>
          </div>
          <div class="tasksList__items">
            ${items}
          </div>
        </div>
      `;
    }).join("");

    el.tasksList.innerHTML = sections || '<div class="empty">No hay tareas para mostrar.</div>';

    // Wire edit buttons
    el.tasksList.addEventListener("click", (ev) => {
      const btn = ev.target.closest("button[data-action='edit']");
      if (!btn) return;
      const item = ev.target.closest(".taskListItem");
      if (!item) return;
      const id = item.getAttribute("data-id");
      const t = state.tasks.find((x) => x.id === id);
      if (t) openModal(t);
    });
  }

  function openModal(task) {
    const weekStart = getWeekStartISO();
    if (!task) {
      el.modalTitle.textContent = "Nueva tarea";
      el.taskId.value = "";
      el.taskDay.value = "0";
      el.taskDate.value = addDays(weekStart, 0);
      el.taskTitle.value = "";
      el.taskDept.value = "";
      el.taskOwner.value = "";
      el.taskStatus.value = "Pendiente";
      el.taskProgress.value = "0";
      el.taskPriority.value = "Media";
      el.btnDeleteTask.style.display = "none";
    } else {
      el.modalTitle.textContent = "Editar tarea";
      el.taskId.value = task.id;
      el.taskDate.value = task.dateISO;
      const idx = dayIndexFromISODate(weekStart, task.dateISO);
      el.taskDay.value = String(Math.min(4, Math.max(0, idx)));
      el.taskTitle.value = task.title ?? "";
      el.taskDept.value = task.dept ?? "";
      el.taskOwner.value = task.owner ?? "";
      el.taskStatus.value = task.status ?? "Pendiente";
      el.taskProgress.value = String(clampProgress(task.progress));
      el.taskPriority.value = task.priority ?? "Media";
      el.btnDeleteTask.style.display = "inline-flex";
    }

    // Keep date/day in sync
    el.taskDay.onchange = () => {
      el.taskDate.value = addDays(weekStart, Number(el.taskDay.value));
    };
    el.taskDate.onchange = () => {
      const idx = dayIndexFromISODate(weekStart, el.taskDate.value);
      if (idx >= 0 && idx <= 4) el.taskDay.value = String(idx);
    };

    el.modal.showModal();
  }

  function upsertTaskFromModal() {
    const weekStart = getWeekStartISO();
    const id = el.taskId.value || uid();
    const dateISO = el.taskDate.value;
    const idx = dayIndexFromISODate(weekStart, dateISO);
    const safeIdx = idx >= 0 && idx <= 4 ? idx : Number(el.taskDay.value);
    const fixedDate = addDays(weekStart, safeIdx);

    const title = el.taskTitle.value.trim();
    const dept = el.taskDept.value.trim();
    const owner = el.taskOwner.value.trim();
    const progress = clampProgress(el.taskProgress.value);
    const status = normalizeStatus(el.taskStatus.value, progress);
    const priority = el.taskPriority.value || "Media";

    if (!title || !dept || !owner) return null;

    const task = { id, dateISO: fixedDate, title, dept, owner, progress, status, priority };
    const idxExisting = state.tasks.findIndex((t) => t.id === id);
    if (idxExisting >= 0) state.tasks[idxExisting] = task;
    else state.tasks.push(task);
    saveState();
    return task;
  }

  function deleteTask(id) {
    state.tasks = state.tasks.filter((t) => t.id !== id);
    saveState();
  }

  function toExcelRows(tasks) {
    return tasks
      .slice()
      .sort((a, b) => (a.dateISO === b.dateISO ? a.title.localeCompare(b.title) : a.dateISO.localeCompare(b.dateISO)))
      .map((t) => ({
        Fecha: t.dateISO,
        Día: dayNameFromISO(t.dateISO),
        Tarea: t.title,
        Departamento: t.dept,
        Responsable: t.owner,
        Progreso: clampProgress(t.progress),
        Estado: normalizeStatus(t.status, clampProgress(t.progress)),
        Prioridad: t.priority || "Media",
      }));
  }

  function dayNameFromISO(isoDate) {
    const d = fromISODate(isoDate);
    // JS: 0 Sunday ... 6 Saturday. We only use Mon-Fri but keep mapping.
    const map = {
      1: "Lunes",
      2: "Martes",
      3: "Miércoles",
      4: "Jueves",
      5: "Viernes",
      6: "Sábado",
      0: "Domingo",
    };
    return map[d.getDay()];
  }

  function downloadExcelCurrentWeek() {
    if (!window.XLSX) {
      alert("No se pudo cargar la librería XLSX. Si tu red bloquea CDN, puedo dejarla local.");
      return;
    }
    const weekStart = getWeekStartISO();
    const weekTasks = state.tasks.filter((t) => getTaskWeekISO(t) === weekStart);
    const rows = toExcelRows(weekTasks);
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Agenda");
    const fileName = `agenda_${weekStart}.xlsx`;
    XLSX.writeFile(wb, fileName);
  }

  function downloadTemplate() {
    if (!window.XLSX) {
      alert("No se pudo cargar la librería XLSX. Si tu red bloquea CDN, puedo dejarla local.");
      return;
    }
    const weekStart = getWeekStartISO();
    const rows = [
      {
        Fecha: addDays(weekStart, 0),
        Día: "Lunes",
        Tarea: "Ej. Actualizar expedientes",
        Departamento: "RH",
        Responsable: "Nombre Apellido",
        Progreso: 0,
        Estado: "Pendiente",
      },
    ];
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Agenda");
    XLSX.writeFile(wb, "plantilla_agenda_semanal.xlsx");
  }

  function parseDayToIdx(day) {
    const d = (day ?? "").toString().trim().toLowerCase();
    const normalized = d
      .replaceAll("á", "a")
      .replaceAll("é", "e")
      .replaceAll("í", "i")
      .replaceAll("ó", "o")
      .replaceAll("ú", "u");
    const map = {
      lunes: 0,
      martes: 1,
      miercoles: 2,
      "miércoles": 2,
      jueves: 3,
      viernes: 4,
    };
    return map[normalized] ?? null;
  }

  function excelDateToISO(v) {
    if (!v && v !== 0) return "";
    if (typeof v === "string") {
      // try yyyy-mm-dd or dd/mm/yyyy
      const s = v.trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (m) {
        const dd = pad2(Number(m[1]));
        const mm = pad2(Number(m[2]));
        const yy = Number(m[3]);
        return `${yy}-${mm}-${dd}`;
      }
      const dt = new Date(s);
      if (!Number.isNaN(dt.getTime())) return toISODate(dt);
      return "";
    }
    if (typeof v === "number") {
      // Excel serial date
      const parsed = XLSX.SSF.parse_date_code(v);
      if (!parsed) return "";
      const dt = new Date(parsed.y, parsed.m - 1, parsed.d);
      return toISODate(dt);
    }
    if (v instanceof Date) return toISODate(v);
    return "";
  }

  function pickAny(obj, keys) {
    for (const k of keys) {
      if (k in obj) return obj[k];
      // case-insensitive match
      const found = Object.keys(obj).find((x) => x.toLowerCase() === k.toLowerCase());
      if (found) return obj[found];
    }
    return undefined;
  }

  function importRows(rows, weekStartISO) {
    /** @type {any[]} */
    const imported = [];
    for (const r of rows) {
      const title = (pickAny(r, ["Tarea", "tarea", "Actividad", "actividad", "Task"]) ?? "").toString().trim();
      if (!title) continue;

      const dept = (pickAny(r, ["Departamento", "depto", "Depto", "Area", "Área"]) ?? "").toString().trim();
      const owner = (pickAny(r, ["Responsable", "responsable", "Owner", "Encargado"]) ?? "").toString().trim();
      const progRaw = pickAny(r, ["Progreso", "progreso", "%", "Avance", "avance"]);
      const progress = clampProgress(progRaw ?? 0);

      const statusRaw = pickAny(r, ["Estado", "estado", "Status"]);
      const status = normalizeStatus(statusRaw, progress);

      const priorityRaw = pickAny(r, ["Prioridad", "prioridad", "Priority"]);
      const priority = PRIORITIES.includes(priorityRaw?.toString().trim()) ? priorityRaw.toString().trim() : "Media";

      const dateRaw = pickAny(r, ["Fecha", "fecha", "Date"]);
      const dayRaw = pickAny(r, ["Día", "Dia", "día", "dia", "Day"]);
      let dateISO = excelDateToISO(dateRaw);

      if (!dateISO) {
        const idx = parseDayToIdx(dayRaw);
        if (idx !== null) dateISO = addDays(weekStartISO, idx);
      }

      // As a last resort, place on Monday of the chosen week
      if (!dateISO) dateISO = addDays(weekStartISO, 0);

      // Snap tasks into the selected week (Mon-Fri) if they fall outside
      const idxInWeek = dayIndexFromISODate(weekStartISO, dateISO);
      const safeIdx = idxInWeek >= 0 && idxInWeek <= 4 ? idxInWeek : 0;
      const fixedDate = addDays(weekStartISO, safeIdx);

      imported.push({
        id: uid(),
        dateISO: fixedDate,
        title,
        dept: dept || "—",
        owner: owner || "—",
        progress,
        status,
        priority,
      });
    }
    return imported;
  }

  async function readExcelFile(file) {
    if (!window.XLSX) {
      alert("No se pudo cargar la librería XLSX. Si tu red bloquea CDN, puedo dejarla local.");
      return;
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    return rows;
  }

  function loadSample() {
    const weekStart = getWeekStartISO();
    const sample = [
      { Fecha: addDays(weekStart, 0), Tarea: "Actualizar expedientes", Departamento: "RH", Responsable: "María", Progreso: 0, Estado: "Pendiente", Prioridad: "Alta" },
      { Día: "Martes", Tarea: "Revisión de incidencias", Departamento: "RH", Responsable: "Carlos", Progreso: 35, Estado: "En progreso", Prioridad: "Media" },
      { Día: "Miércoles", Tarea: "Capacitación (seguimiento)", Departamento: "RH", Responsable: "María", Progreso: 60, Estado: "En progreso", Prioridad: "Baja" },
      { Día: "Jueves", Tarea: "Reporte semanal a gerencia", Departamento: "RH", Responsable: "Ana", Progreso: 0, Estado: "Pendiente", Prioridad: "Alta" },
      { Día: "Viernes", Tarea: "Cierre de pendientes", Departamento: "RH", Responsable: "Carlos", Progreso: 100, Estado: "Hecho", Prioridad: "Media" },
    ];
    state.tasks = importRows(sample, weekStart);
    saveState();
    render();
  }

  function wireEvents() {
    el.btnTodayWeek.addEventListener("click", () => {
      setWeekStartISO(toISODate(startOfWeekMonday(new Date())));
      render();
    });

    el.weekStart.addEventListener("change", () => {
      if (!el.weekStart.value) return;
      // Snap chosen date to Monday
      const d = fromISODate(el.weekStart.value);
      setWeekStartISO(toISODate(startOfWeekMonday(d)));
      render();
    });

    el.btnClearWeek.addEventListener("click", () => {
      if (!confirm("¿Seguro? Se borrarán las tareas guardadas localmente en este navegador.")) return;
      localStorage.removeItem(STORAGE_KEY);
      state = { tasks: [], weekStartISO: getWeekStartISO() };
      saveState();
      render();
    });

    el.btnNewTask.addEventListener("click", () => openModal(null));

    el.weekGrid.addEventListener("click", (ev) => {
      const btn = ev.target.closest("button[data-action='edit']");
      if (!btn) return;
      const card = ev.target.closest(".task");
      if (!card) return;
      const id = card.getAttribute("data-id");
      const t = state.tasks.find((x) => x.id === id);
      if (t) openModal(t);
    });

    el.form.addEventListener("submit", (ev) => {
      // Dialog will close by default; we keep it, validate, then close.
      ev.preventDefault();
      const task = upsertTaskFromModal();
      if (!task) {
        alert("Completa Tarea, Departamento y Responsable.");
        return;
      }
      el.modal.close();
      render();
    });

    el.btnDeleteTask.addEventListener("click", () => {
      const id = el.taskId.value;
      if (!id) return;
      if (!confirm("¿Eliminar esta tarea?")) return;
      deleteTask(id);
      el.modal.close();
      render();
    });

    el.filterDept.addEventListener("change", () => {
      render();
      renderTasksList();
    });
    el.filterOwner.addEventListener("change", () => {
      render();
      renderTasksList();
    });
    el.filterStatus.addEventListener("change", () => {
      render();
      renderTasksList();
    });

    el.btnDownload.addEventListener("click", downloadExcelCurrentWeek);
    el.btnTemplate.addEventListener("click", downloadTemplate);
    el.btnLoadSample.addEventListener("click", loadSample);

    el.excelFile.addEventListener("change", async () => {
      const f = el.excelFile.files?.[0];
      if (!f) return;
      try {
        const weekStart = getWeekStartISO();
        const rows = await readExcelFile(f);
        const imported = importRows(rows, weekStart);
        if (!imported.length) {
          alert("No encontré filas válidas. Revisa columnas: Fecha/Día, Tarea, Departamento, Responsable, Progreso, Estado.");
          return;
        }
        // Merge strategy: replace the week
        state.tasks = state.tasks.filter((t) => getTaskWeekISO(t) !== weekStart).concat(imported);
        saveState();
        render();
        alert(`Listo: se importaron ${imported.length} tarea(s) para la semana seleccionada.`);
      } catch (e) {
        console.error(e);
        alert("No se pudo leer el archivo. Si es Excel, intenta guardarlo como .xlsx y reintenta.");
      } finally {
        el.excelFile.value = "";
      }
    });
  }

  function init() {
    loadState();

    const initialWeek = state.weekStartISO || toISODate(startOfWeekMonday(new Date()));
    setWeekStartISO(initialWeek);

    wireEvents();
    render();
  }

  init();
})();


