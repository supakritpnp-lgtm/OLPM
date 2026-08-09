import React, { useState, useMemo, useEffect, useCallback } from "react";
import {
  Search,
  RefreshCw,
  ExternalLink,
  Bug,
  MessageSquarePlus,
  GitPullRequest,
  Wrench,
  ChevronRight,
  ChevronLeft,
  CircleDot,
  AlertTriangle,
  Plus,
  Pencil,
  Trash2,
  X,
  Check,
} from "lucide-react";
import { API_URL } from "./config.js";

// ---------------------------------------------------------------------------
// Helpers to normalize rows coming back from Apps Script (raw sheet headers)
// ---------------------------------------------------------------------------

function normalizeProject(row) {
  return {
    id: row["Project ID"],
    name: row["Project name"] || "(ไม่มีชื่อ)",
    type: row["Type"] || "",
    pm: row["PM"] || "—",
    sales: row["Sales Owner"] || "—",
    start: row["Start Date"],
    due: row["Due Date"],
    value: Number(row["Project Value"]) || 0,
    progress: Number(row["Task progress"]) || 0,
    platform: row["Platform"] || "",
  };
}

function normalizeTask(row) {
  return {
    rowIndex: row.rowIndex,
    type: row["Type"] || "Request",
    status: row["Status"] || "Not start",
    progress: Number(row["Progress"]) || 0,
    name: row["Task Name"] || "(ไม่มีชื่องาน)",
    link: row["Link"] || "",
    detail: row["Task Detail"] || "",
  };
}

// ---------------------------------------------------------------------------
// Design tokens — minimal, bright, "office desk" theme
// ---------------------------------------------------------------------------

const C = {
  bg: "#F6F6F3",          // paper-white app background
  surface: "#FFFFFF",     // cards, sidebar, rows
  surfaceSoft: "#FBFBF9", // hover
  border: "#E4E3DD",
  borderStrong: "#D6D5CD",
  ink: "#24262B",         // primary text
  inkMuted: "#8B8D8F",    // secondary text
  inkFaint: "#B7B7B0",
  accent: "#33538F",      // stamp-ink blue
  accentSoft: "rgba(51,83,143,0.10)",
  amber: "#B9822E",
  amberSoft: "rgba(185,130,46,0.12)",
  green: "#3E8264",
  greenSoft: "rgba(62,130,100,0.12)",
  red: "#B85042",
  redSoft: "rgba(184,80,66,0.12)",
  violet: "#6E5CA6",
  violetSoft: "rgba(110,92,166,0.12)",
};

const TYPE_META = {
  CR: { label: "CR", icon: GitPullRequest, color: C.violet, bg: C.violetSoft },
  Request: { label: "Request", icon: MessageSquarePlus, color: C.accent, bg: C.accentSoft },
  Bug: { label: "Bug", icon: Bug, color: C.red, bg: C.redSoft },
  Sol: { label: "Sol", icon: Wrench, color: C.amber, bg: C.amberSoft },
};

const STATUS_META = {
  "Not start": { color: C.inkMuted, bg: "#EFEFEC" },
  Que: { color: C.amber, bg: C.amberSoft },
  Done: { color: C.green, bg: C.greenSoft },
};

const thb = (n) =>
  n ? n.toLocaleString("th-TH", { style: "currency", currency: "THB", maximumFractionDigits: 0 }) : "—";

const dmy = (s) => {
  if (!s) return "—";
  const d = new Date(s);
  if (isNaN(d)) return "—";
  return d.toLocaleDateString("th-TH", { day: "2-digit", month: "short", year: "2-digit" });
};

const PM_USERS = ["Art", "Jay", "Joon", "Pla"];
const USER_STORAGE_KEY = "project_desk_user";

function fmtLogTime(v) {
  const d = new Date(v);
  if (isNaN(d)) return String(v);
  return d.toLocaleDateString("th-TH", { day: "2-digit", month: "short", year: "2-digit" }) +
    " " + d.toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit" });
}

// Tally-mark progress — reads like a ledger checkbox row, not a gauge
function SignalBar({ value = 0, segments = 8, height = 14 }) {
  const filled = Math.round(value * segments);
  return (
    <div style={{ display: "flex", gap: 3, alignItems: "flex-end" }}>
      {Array.from({ length: segments }).map((_, i) => {
        const isFilled = i < filled;
        const h = height * (0.5 + (i / segments) * 0.5);
        return (
          <div
            key={i}
            style={{
              width: 3.5,
              height: h,
              borderRadius: 1,
              background: isFilled ? (value >= 1 ? C.green : C.accent) : C.border,
              transition: "background 0.2s",
            }}
          />
        );
      })}
    </div>
  );
}

function TypeBadge({ type }) {
  const meta = TYPE_META[type] || TYPE_META.Request;
  const Icon = meta.icon;
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        padding: "3px 8px",
        borderRadius: 6,
        fontSize: 11,
        fontFamily: "'IBM Plex Mono', monospace",
        letterSpacing: 0.3,
        color: meta.color,
        background: meta.bg,
        whiteSpace: "nowrap",
      }}
    >
      <Icon size={11} strokeWidth={2.2} />
      {meta.label}
    </span>
  );
}

function StatusPill({ status }) {
  const meta = STATUS_META[status] || STATUS_META["Not start"];
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        padding: "3px 9px",
        borderRadius: 999,
        fontSize: 11,
        fontWeight: 600,
        color: meta.color,
        background: meta.bg,
      }}
    >
      <CircleDot size={9} />
      {status}
    </span>
  );
}

function UserGate({ onSelect }) {
  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans Thai', 'IBM Plex Sans', sans-serif",
        background: C.bg,
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
    >
      <style>{`@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap');`}</style>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: 32, width: 340, boxShadow: "0 4px 16px rgba(0,0,0,0.05)" }}>
        <div style={{ fontSize: 11, letterSpacing: 1.5, color: C.inkMuted, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 6 }}>
          PROJECT DESK
        </div>
        <h2 style={{ fontSize: 18, fontWeight: 700, margin: "0 0 6px", color: C.ink }}>คุณคือใคร?</h2>
        <p style={{ fontSize: 12.5, color: C.inkMuted, margin: "0 0 18px" }}>
          เลือกชื่อของคุณ — ใช้เพื่อบันทึกใน log ว่าใครแก้ไขอะไร (ไม่มีรหัสผ่าน)
        </p>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {PM_USERS.map((name) => (
            <button
              key={name}
              onClick={() => onSelect(name)}
              style={{
                padding: "10px 14px",
                borderRadius: 9,
                border: `1px solid ${C.border}`,
                background: C.bg,
                color: C.ink,
                fontSize: 14,
                fontWeight: 600,
                textAlign: "left",
              }}
            >
              {name}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

function LogView({ entries, loading, onBack }) {
  return (
    <>
      <div style={{ padding: "20px 32px 16px", borderBottom: `1px solid ${C.border}` }}>
        <button
          onClick={onBack}
          style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: C.inkMuted, fontSize: 12.5, marginBottom: 10, padding: 0 }}
        >
          <ChevronLeft size={14} /> โปรเจกต์
        </button>
        <h2 style={{ fontSize: 22, fontWeight: 700, margin: 0, letterSpacing: -0.2 }}>Log การแก้ไข</h2>
        <div style={{ fontSize: 12.5, color: C.inkMuted, marginTop: 4 }}>
          บันทึกล่าสุดก่อน — เพราะระบบไม่มีการ login จึงระบุได้แค่ชื่อที่ผู้ใช้เลือกไว้เอง
        </div>
      </div>
      <div style={{ flex: 1, overflow: "auto", padding: "16px 32px 32px" }}>
        {loading ? (
          <div style={{ color: C.inkMuted, padding: "60px 0", textAlign: "center" }}>กำลังโหลด log...</div>
        ) : entries.length === 0 ? (
          <div style={{ color: C.inkMuted, padding: "60px 0", textAlign: "center" }}>ยังไม่มีการแก้ไขใดๆ</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {entries.map((e, i) => (
              <div
                key={i}
                style={{
                  display: "flex",
                  alignItems: "baseline",
                  gap: 12,
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  padding: "10px 14px",
                  fontSize: 12.5,
                }}
              >
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.inkFaint, fontSize: 11, minWidth: 130 }}>
                  {fmtLogTime(e.Timestamp)}
                </span>
                <span style={{ fontWeight: 700, color: C.accent, minWidth: 44 }}>{e.User}</span>
                <span style={{ color: C.ink }}>{e.Action}</span>
                {e.Project && (
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.inkMuted, fontSize: 11 }}>{e.Project}</span>
                )}
                {e.Detail && <span style={{ color: C.inkMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>— {e.Detail}</span>}
              </div>
            ))}
          </div>
        )}
      </div>
    </>
  );
}


export default function App() {
  const [currentUser, setCurrentUser] = useState(() => {
    try {
      return localStorage.getItem(USER_STORAGE_KEY) || null;
    } catch {
      return null;
    }
  });
  const [mode, setMode] = useState("projects"); // "projects" | "log"
  const [logEntries, setLogEntries] = useState([]);
  const [loadingLog, setLoadingLog] = useState(false);

  const [projects, setProjects] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [tasks, setTasks] = useState([]);
  const [query, setQuery] = useState("");
  const [typeFilter, setTypeFilter] = useState("All");
  const [loadingProjects, setLoadingProjects] = useState(true);
  const [loadingTasks, setLoadingTasks] = useState(false);
  const [error, setError] = useState(null);
  const [lastSync, setLastSync] = useState(null);
  const [savingRow, setSavingRow] = useState(null);
  const [showAddForm, setShowAddForm] = useState(false);
  const [newTask, setNewTask] = useState({ type: "Request", name: "", detail: "", link: "" });
  const [addingTask, setAddingTask] = useState(false);
  const [editingRowIndex, setEditingRowIndex] = useState(null);
  const [editFields, setEditFields] = useState({ type: "", name: "", detail: "", link: "" });
  const [savingEdit, setSavingEdit] = useState(false);
  const [deletingRow, setDeletingRow] = useState(null);
  const [showAddProjectForm, setShowAddProjectForm] = useState(false);
  const [newProject, setNewProject] = useState({ name: "", pm: PM_USERS[0], sales: "", start: "", due: "", platform: "" });
  const [addingProject, setAddingProject] = useState(false);

  const loadProjects = useCallback(async () => {
    setLoadingProjects(true);
    setError(null);
    try {
      const res = await fetch(`${API_URL}?action=overview`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setProjects(data.map(normalizeProject).filter((p) => p.id));
      setLastSync(new Date());
    } catch (err) {
      setError("โหลดข้อมูลจาก Google Sheet ไม่สำเร็จ: " + err.message);
    } finally {
      setLoadingProjects(false);
    }
  }, []);

  const loadTasks = useCallback(async (projectId) => {
    setLoadingTasks(true);
    setError(null);
    try {
      const res = await fetch(`${API_URL}?action=tasks&projectId=${encodeURIComponent(projectId)}`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setTasks(data.map(normalizeTask));
    } catch (err) {
      setError("โหลดรายการงานไม่สำเร็จ: " + err.message);
    } finally {
      setLoadingTasks(false);
    }
  }, []);

  const loadLog = useCallback(async () => {
    setLoadingLog(true);
    setError(null);
    try {
      const res = await fetch(`${API_URL}?action=log&limit=200`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setLogEntries(data);
    } catch (err) {
      setError("โหลด log ไม่สำเร็จ: " + err.message);
    } finally {
      setLoadingLog(false);
    }
  }, []);

  useEffect(() => {
    if (mode === "log") loadLog();
  }, [mode, loadLog]);

  useEffect(() => {
    loadProjects();
  }, [loadProjects]);

  useEffect(() => {
    if (selectedId) loadTasks(selectedId);
  }, [selectedId, loadTasks]);

  const selected = projects.find((p) => p.id === selectedId) || null;

  const filteredProjects = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return projects;
    return projects.filter(
      (p) =>
        p.name.toLowerCase().includes(q) ||
        p.id.toLowerCase().includes(q) ||
        p.pm.toLowerCase().includes(q)
    );
  }, [projects, query]);

  const filteredTasks = useMemo(() => {
    if (typeFilter === "All") return tasks;
    return tasks.filter((t) => t.type === typeFilter);
  }, [tasks, typeFilter]);

  const totals = useMemo(
    () => ({
      projects: projects.length,
      done: projects.filter((p) => p.progress >= 1).length,
    }),
    [projects]
  );

  async function cycleStatus(task) {
    const order = ["Not start", "Que", "Done"];
    const nextStatus = order[(order.indexOf(task.status) + 1) % order.length];
    const nextProgress = nextStatus === "Done" ? 1 : nextStatus === "Que" ? Math.max(task.progress, 0.3) : 0;

    setTasks((prev) =>
      prev.map((t) => (t.rowIndex === task.rowIndex ? { ...t, status: nextStatus, progress: nextProgress } : t))
    );
    setSavingRow(task.rowIndex);
    try {
      const res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "updateTaskStatus",
          projectId: selectedId,
          rowIndex: task.rowIndex,
          status: nextStatus,
          progress: nextProgress,
          user: currentUser,
        }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
    } catch (err) {
      setError("บันทึกสถานะไม่สำเร็จ: " + err.message + " (กำลังโหลดข้อมูลล่าสุดใหม่)");
      loadTasks(selectedId);
    } finally {
      setSavingRow(null);
    }
  }

  function handleSyncClick() {
    if (mode === "log") {
      loadLog();
      return;
    }
    if (selectedId) loadTasks(selectedId);
    loadProjects();
  }

  async function submitNewTask() {
    if (!newTask.name.trim()) {
      setError("กรอกชื่องานก่อนบันทึก");
      return;
    }
    setAddingTask(true);
    setError(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "addTask",
          projectId: selectedId,
          user: currentUser,
          task: {
            Type: newTask.type,
            Status: "Not start",
            Progress: 0,
            "Task Name": newTask.name,
            Link: newTask.link,
            "Task Detail": newTask.detail,
          },
        }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setNewTask({ type: "Request", name: "", detail: "", link: "" });
      setShowAddForm(false);
      await loadTasks(selectedId);
    } catch (err) {
      setError("เพิ่มงานไม่สำเร็จ: " + err.message);
    } finally {
      setAddingTask(false);
    }
  }

  function startEdit(task) {
    setEditingRowIndex(task.rowIndex);
    setEditFields({ type: task.type, name: task.name, detail: task.detail, link: task.link });
  }

  function cancelEdit() {
    setEditingRowIndex(null);
  }

  async function saveEdit(rowIndex) {
    if (!editFields.name.trim()) {
      setError("ชื่องานห้ามว่าง");
      return;
    }
    setSavingEdit(true);
    setError(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "editTask",
          projectId: selectedId,
          rowIndex,
          user: currentUser,
          fields: {
            Type: editFields.type,
            "Task Name": editFields.name,
            Link: editFields.link,
            "Task Detail": editFields.detail,
          },
        }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setTasks((prev) =>
        prev.map((t) =>
          t.rowIndex === rowIndex
            ? { ...t, type: editFields.type, name: editFields.name, link: editFields.link, detail: editFields.detail }
            : t
        )
      );
      setEditingRowIndex(null);
    } catch (err) {
      setError("บันทึกการแก้ไขไม่สำเร็จ: " + err.message);
    } finally {
      setSavingEdit(false);
    }
  }

  async function confirmDelete(task) {
    if (!window.confirm(`ลบงาน "${task.name}" ใช่ไหม? ลบแล้วกู้คืนไม่ได้`)) return;
    setDeletingRow(task.rowIndex);
    setError(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "deleteTask", projectId: selectedId, rowIndex: task.rowIndex, user: currentUser }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      await loadTasks(selectedId);
    } catch (err) {
      setError("ลบงานไม่สำเร็จ: " + err.message);
    } finally {
      setDeletingRow(null);
    }
  }

  function selectUser(name) {
    try {
      if (name) localStorage.setItem(USER_STORAGE_KEY, name);
      else localStorage.removeItem(USER_STORAGE_KEY);
    } catch {}
    setCurrentUser(name);
  }

  async function submitNewProject() {
    if (!newProject.name.trim()) {
      setError("กรอกชื่อโปรเจกต์ก่อนบันทึก");
      return;
    }
    setAddingProject(true);
    setError(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "addProject", user: currentUser, project: newProject }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setNewProject({ name: "", pm: PM_USERS[0], sales: "", start: "", due: "", platform: "" });
      setShowAddProjectForm(false);
      await loadProjects();
      if (data.projectId) setSelectedId(data.projectId);
    } catch (err) {
      setError("เพิ่มโปรเจกต์ไม่สำเร็จ: " + err.message);
    } finally {
      setAddingProject(false);
    }
  }

  if (!currentUser) {
    return <UserGate onSelect={selectUser} />;
  }

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans Thai', 'IBM Plex Sans', sans-serif",
        background: C.bg,
        color: C.ink,
        minHeight: "100vh",
        display: "flex",
        fontSize: 14,
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${C.borderStrong}; border-radius: 4px; }
        .row-hover:hover { background: ${C.surfaceSoft} !important; }
        button { font-family: inherit; cursor: pointer; }
        input { font-family: inherit; }
        @keyframes spin { from { transform: rotate(0); } to { transform: rotate(360deg); } }
      `}</style>

      {/* Sidebar */}
      <aside
        style={{
          width: 300,
          minWidth: 300,
          borderRight: `1px solid ${C.border}`,
          display: "flex",
          flexDirection: "column",
          background: C.surface,
        }}
      >
        <div style={{ padding: "20px 18px 14px" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div
                style={{
                  width: 7,
                  height: 7,
                  borderRadius: "50%",
                  background: loadingProjects ? C.inkFaint : C.green,
                }}
              />
              <span style={{ fontSize: 11, letterSpacing: 1.5, color: C.inkMuted, fontFamily: "'IBM Plex Mono', monospace" }}>
                PROJECT DESK
              </span>
            </div>
            <button
              onClick={() => selectUser(null)}
              title="เปลี่ยนผู้ใช้"
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: C.accent,
                background: C.accentSoft,
                border: "none",
                borderRadius: 999,
                padding: "3px 9px",
              }}
            >
              {currentUser}
            </button>
          </div>
          <h1 style={{ fontSize: 19, fontWeight: 700, margin: "6px 0 0", letterSpacing: -0.2 }}>งานภายในทีม</h1>
        </div>

        <div style={{ display: "flex", gap: 6, padding: "0 18px 14px" }}>
          <button
            onClick={() => setMode("projects")}
            style={{
              flex: 1,
              padding: "7px 0",
              borderRadius: 8,
              border: `1px solid ${mode === "projects" ? C.accent : C.border}`,
              background: mode === "projects" ? C.accentSoft : "transparent",
              color: mode === "projects" ? C.accent : C.inkMuted,
              fontSize: 12.5,
              fontWeight: 600,
            }}
          >
            โปรเจกต์
          </button>
          <button
            onClick={() => setMode("log")}
            style={{
              flex: 1,
              padding: "7px 0",
              borderRadius: 8,
              border: `1px solid ${mode === "log" ? C.accent : C.border}`,
              background: mode === "log" ? C.accentSoft : "transparent",
              color: mode === "log" ? C.accent : C.inkMuted,
              fontSize: 12.5,
              fontWeight: 600,
            }}
          >
            Log
          </button>
        </div>

        <div style={{ padding: "0 18px 14px" }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              background: C.bg,
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              padding: "8px 10px",
            }}
          >
            <Search size={15} color={C.inkFaint} />
            <input
              placeholder="ค้นหาโปรเจกต์ / PM"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              style={{ background: "transparent", border: "none", outline: "none", color: C.ink, fontSize: 13, width: "100%" }}
            />
          </div>
        </div>

        <div style={{ flex: 1, overflowY: "auto", padding: "0 10px" }}>
          {loadingProjects && (
            <div style={{ color: C.inkMuted, fontSize: 12.5, padding: "16px 12px" }}>กำลังโหลดโปรเจกต์...</div>
          )}
          {!loadingProjects &&
            filteredProjects.map((p) => {
              const active = p.id === selectedId;
              return (
                <button
                  key={p.id}
                  onClick={() => { setSelectedId(p.id); setMode("projects"); }}
                  className="row-hover"
                  style={{
                    width: "100%",
                    textAlign: "left",
                    background: active ? C.accentSoft : "transparent",
                    border: active ? `1px solid ${C.accent}33` : "1px solid transparent",
                    borderRadius: 10,
                    padding: "10px 12px",
                    marginBottom: 4,
                    display: "block",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: C.inkMuted }}>{p.id}</span>
                    <SignalBar value={p.progress} segments={5} height={10} />
                  </div>
                  <div style={{ fontSize: 13.5, fontWeight: 600, marginTop: 3, color: C.ink }}>{p.name}</div>
                  <div style={{ fontSize: 11.5, color: C.inkMuted, marginTop: 2 }}>
                    PM: {p.pm} {p.platform ? `· ${p.platform}` : ""}
                  </div>
                </button>
              );
            })}
        </div>

        <div style={{ padding: 16, borderTop: `1px solid ${C.border}` }}>
          <button
            onClick={handleSyncClick}
            style={{
              width: "100%",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 8,
              background: C.surface,
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              padding: "9px 10px",
              color: C.accent,
              fontSize: 12.5,
              fontWeight: 600,
              boxShadow: "0 1px 2px rgba(0,0,0,0.03)",
            }}
          >
            <RefreshCw size={13} style={{ animation: loadingProjects || loadingTasks ? "spin 0.8s linear infinite" : "none" }} />
            {loadingProjects || loadingTasks ? "กำลังซิงก์..." : "ซิงก์กับ Google Sheet"}
          </button>
          <div style={{ fontSize: 10.5, color: C.inkFaint, marginTop: 6, textAlign: "center", fontFamily: "'IBM Plex Mono', monospace" }}>
            {lastSync
              ? `อัปเดตล่าสุด ${lastSync.toLocaleDateString("th-TH", { day: "2-digit", month: "short" })} ${lastSync.toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit" })}`
              : "ยังไม่ได้ซิงก์"}
          </div>
        </div>
      </aside>

      {/* Main */}
      <main style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        {error && (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              background: C.redSoft,
              borderBottom: `1px solid ${C.red}33`,
              color: C.red,
              padding: "10px 32px",
              fontSize: 12.5,
            }}
          >
            <AlertTriangle size={14} />
            {error}
          </div>
        )}

        {mode === "log" ? (
          <LogView entries={logEntries} loading={loadingLog} onBack={() => setMode("projects")} />
        ) : !selected ? (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", padding: "22px 32px", borderBottom: `1px solid ${C.border}`, flexWrap: "wrap", gap: 12 }}>
              <div style={{ display: "flex", gap: 24 }}>
                {[
                  { label: "โปรเจกต์ทั้งหมด", value: totals.projects },
                  { label: "เสร็จสมบูรณ์", value: totals.done, color: C.green },
                ].map((s) => (
                  <div key={s.label}>
                    <div style={{ fontSize: 26, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: s.color || C.ink }}>
                      {s.value}
                    </div>
                    <div style={{ fontSize: 11.5, color: C.inkMuted, marginTop: 2 }}>{s.label}</div>
                  </div>
                ))}
              </div>
              <button
                onClick={() => setShowAddProjectForm((v) => !v)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  padding: "8px 16px",
                  borderRadius: 8,
                  border: `1px solid ${showAddProjectForm ? C.border : C.accent}`,
                  background: showAddProjectForm ? C.surface : C.accent,
                  color: showAddProjectForm ? C.inkMuted : "#fff",
                  fontSize: 12.5,
                  fontWeight: 600,
                }}
              >
                {showAddProjectForm ? <X size={14} /> : <Plus size={14} />}
                {showAddProjectForm ? "ยกเลิก" : "เพิ่มโปรเจกต์"}
              </button>
            </div>

            {showAddProjectForm && (
              <div style={{ margin: "16px 32px 0", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, padding: 16, boxShadow: "0 1px 2px rgba(0,0,0,0.03)" }}>
                <div style={{ fontSize: 11, color: C.inkFaint, marginBottom: 10, fontFamily: "'IBM Plex Mono', monospace" }}>
                  รหัสโปรเจกต์จะถูกสร้างให้อัตโนมัติ (เช่น PM-{new Date().getFullYear().toString().slice(-2)}{String(new Date().getMonth() + 1).padStart(2, "0")}01)
                </div>
                <div style={{ display: "flex", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                  <input
                    placeholder="ชื่อโปรเจกต์"
                    value={newProject.name}
                    onChange={(e) => setNewProject({ ...newProject, name: e.target.value })}
                    style={{ flex: 1, minWidth: 220, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                  />
                  <select
                    value={newProject.pm}
                    onChange={(e) => setNewProject({ ...newProject, pm: e.target.value })}
                    style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, background: C.bg }}
                  >
                    {PM_USERS.map((u) => (
                      <option key={u} value={u}>{u}</option>
                    ))}
                  </select>
                </div>
                <div style={{ display: "flex", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                  <input
                    placeholder="Sales Owner"
                    value={newProject.sales}
                    onChange={(e) => setNewProject({ ...newProject, sales: e.target.value })}
                    style={{ flex: 1, minWidth: 160, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                  />
                  <input
                    placeholder="Platform (เช่น FTS, iOV, Ceiba2)"
                    value={newProject.platform}
                    onChange={(e) => setNewProject({ ...newProject, platform: e.target.value })}
                    style={{ flex: 1, minWidth: 160, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                  />
                </div>
                <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
                  <label style={{ flex: 1, minWidth: 160, fontSize: 11, color: C.inkMuted }}>
                    วันเริ่ม
                    <input
                      type="date"
                      value={newProject.start}
                      onChange={(e) => setNewProject({ ...newProject, start: e.target.value })}
                      style={{ display: "block", width: "100%", marginTop: 4, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                    />
                  </label>
                  <label style={{ flex: 1, minWidth: 160, fontSize: 11, color: C.inkMuted }}>
                    วันกำหนดส่ง
                    <input
                      type="date"
                      value={newProject.due}
                      onChange={(e) => setNewProject({ ...newProject, due: e.target.value })}
                      style={{ display: "block", width: "100%", marginTop: 4, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                    />
                  </label>
                </div>
                <button
                  onClick={submitNewProject}
                  disabled={addingProject}
                  style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: C.accent, color: "#fff", fontSize: 13, fontWeight: 600, opacity: addingProject ? 0.6 : 1 }}
                >
                  {addingProject ? "กำลังบันทึก..." : "สร้างโปรเจกต์"}
                </button>
              </div>
            )}

            <div style={{ flex: 1, overflow: "auto", padding: "8px 32px 32px" }}>
              {loadingProjects ? (
                <div style={{ color: C.inkMuted, padding: "60px 0", textAlign: "center" }}>กำลังโหลดข้อมูลจาก Google Sheet...</div>
              ) : filteredProjects.length === 0 ? (
                <div style={{ color: C.inkMuted, padding: "60px 0", textAlign: "center" }}>ไม่พบโปรเจกต์</div>
              ) : (
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                      {["Project", "PM", "Sales", "กำหนดส่ง", "มูลค่า", "ความคืบหน้า", ""].map((h) => (
                        <th
                          key={h}
                          style={{ textAlign: "left", padding: "12px 10px", fontSize: 11, letterSpacing: 0.5, color: C.inkMuted, fontWeight: 600, textTransform: "uppercase" }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredProjects.map((p) => (
                      <tr
                        key={p.id}
                        className="row-hover"
                        onClick={() => { setSelectedId(p.id); setMode("projects"); }}
                        style={{ borderBottom: `1px solid ${C.border}`, cursor: "pointer" }}
                      >
                        <td style={{ padding: "13px 10px" }}>
                          <div style={{ fontWeight: 600 }}>{p.name}</div>
                          <div style={{ fontSize: 11, color: C.inkMuted, fontFamily: "'IBM Plex Mono', monospace" }}>{p.id}</div>
                        </td>
                        <td style={{ padding: "13px 10px", color: C.ink }}>{p.pm}</td>
                        <td style={{ padding: "13px 10px", color: C.ink }}>{p.sales}</td>
                        <td style={{ padding: "13px 10px", color: C.ink, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{dmy(p.due)}</td>
                        <td style={{ padding: "13px 10px", color: C.ink, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{thb(p.value)}</td>
                        <td style={{ padding: "13px 10px" }}>
                          <SignalBar value={p.progress} segments={8} />
                        </td>
                        <td style={{ padding: "13px 10px", textAlign: "right" }}>
                          <ChevronRight size={16} color={C.inkFaint} />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </>
        ) : (
          <>
            <div style={{ padding: "20px 32px 16px", borderBottom: `1px solid ${C.border}` }}>
              <button
                onClick={() => setSelectedId(null)}
                style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: C.inkMuted, fontSize: 12.5, marginBottom: 10, padding: 0 }}
              >
                <ChevronLeft size={14} /> ภาพรวมทั้งหมด
              </button>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: C.inkMuted }}>{selected.id}</div>
                  <h2 style={{ fontSize: 22, fontWeight: 700, margin: "4px 0 0", letterSpacing: -0.2 }}>{selected.name}</h2>
                  <div style={{ fontSize: 12.5, color: C.inkMuted, marginTop: 4 }}>
                    PM {selected.pm} · Sales {selected.sales} · กำหนดส่ง {dmy(selected.due)}
                  </div>
                </div>
                <SignalBar value={selected.progress} segments={10} height={18} />
              </div>
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 32px 0", flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", gap: 6 }}>
                {["All", "CR", "Request", "Bug", "Sol"].map((t) => (
                  <button
                    key={t}
                    onClick={() => setTypeFilter(t)}
                    style={{
                      padding: "6px 14px",
                      borderRadius: 999,
                      border: "1px solid " + (typeFilter === t ? C.accent : C.border),
                      background: typeFilter === t ? C.accentSoft : "transparent",
                      color: typeFilter === t ? C.accent : C.inkMuted,
                      fontSize: 12.5,
                      fontWeight: 600,
                    }}
                  >
                    {t}
                  </button>
                ))}
              </div>
              <button
                onClick={() => setShowAddForm((v) => !v)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: `1px solid ${showAddForm ? C.border : C.accent}`,
                  background: showAddForm ? C.surface : C.accent,
                  color: showAddForm ? C.inkMuted : "#fff",
                  fontSize: 12.5,
                  fontWeight: 600,
                }}
              >
                {showAddForm ? <X size={14} /> : <Plus size={14} />}
                {showAddForm ? "ยกเลิก" : "เพิ่มงาน"}
              </button>
            </div>

            {showAddForm && (
              <div style={{ margin: "14px 32px 0", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, padding: 16, boxShadow: "0 1px 2px rgba(0,0,0,0.03)" }}>
                <div style={{ display: "flex", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                  <select
                    value={newTask.type}
                    onChange={(e) => setNewTask({ ...newTask, type: e.target.value })}
                    style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, background: C.bg }}
                  >
                    {["CR", "Request", "Bug", "Sol"].map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                  <input
                    placeholder="ชื่องาน"
                    value={newTask.name}
                    onChange={(e) => setNewTask({ ...newTask, name: e.target.value })}
                    style={{ flex: 1, minWidth: 200, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                  />
                </div>
                <input
                  placeholder="ลิงก์ (ClickUp ฯลฯ) — ไม่บังคับ"
                  value={newTask.link}
                  onChange={(e) => setNewTask({ ...newTask, link: e.target.value })}
                  style={{ width: "100%", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, marginBottom: 10 }}
                />
                <textarea
                  placeholder="รายละเอียดเพิ่มเติม — ไม่บังคับ"
                  value={newTask.detail}
                  onChange={(e) => setNewTask({ ...newTask, detail: e.target.value })}
                  rows={2}
                  style={{ width: "100%", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, fontFamily: "inherit", resize: "vertical", marginBottom: 12 }}
                />
                <button
                  onClick={submitNewTask}
                  disabled={addingTask}
                  style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: C.accent, color: "#fff", fontSize: 13, fontWeight: 600, opacity: addingTask ? 0.6 : 1 }}
                >
                  {addingTask ? "กำลังบันทึก..." : "บันทึกงานใหม่"}
                </button>
              </div>
            )}

            <div style={{ flex: 1, overflow: "auto", padding: "16px 32px 32px" }}>
              {loadingTasks ? (
                <div style={{ color: C.inkMuted, padding: "40px 0", textAlign: "center" }}>กำลังโหลดรายการงาน...</div>
              ) : filteredTasks.length === 0 ? (
                <div style={{ color: C.inkMuted, fontSize: 13, padding: "40px 0", textAlign: "center" }}>ยังไม่มีงานประเภทนี้</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {filteredTasks.map((t) => {
                    const isEditing = editingRowIndex === t.rowIndex;
                    const isDeleting = deletingRow === t.rowIndex;
                    if (isEditing) {
                      return (
                        <div
                          key={t.rowIndex}
                          style={{ background: C.surface, border: `1px solid ${C.accent}`, borderRadius: 10, padding: 14, boxShadow: "0 1px 2px rgba(0,0,0,0.03)" }}
                        >
                          <div style={{ display: "flex", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                            <select
                              value={editFields.type}
                              onChange={(e) => setEditFields({ ...editFields, type: e.target.value })}
                              style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, background: C.bg }}
                            >
                              {["CR", "Request", "Bug", "Sol"].map((tp) => (
                                <option key={tp} value={tp}>{tp}</option>
                              ))}
                            </select>
                            <input
                              value={editFields.name}
                              onChange={(e) => setEditFields({ ...editFields, name: e.target.value })}
                              style={{ flex: 1, minWidth: 200, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink }}
                            />
                          </div>
                          <input
                            placeholder="ลิงก์"
                            value={editFields.link}
                            onChange={(e) => setEditFields({ ...editFields, link: e.target.value })}
                            style={{ width: "100%", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, marginBottom: 10 }}
                          />
                          <textarea
                            placeholder="รายละเอียด"
                            value={editFields.detail}
                            onChange={(e) => setEditFields({ ...editFields, detail: e.target.value })}
                            rows={2}
                            style={{ width: "100%", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 13, color: C.ink, fontFamily: "inherit", resize: "vertical", marginBottom: 12 }}
                          />
                          <div style={{ display: "flex", gap: 8 }}>
                            <button
                              onClick={() => saveEdit(t.rowIndex)}
                              disabled={savingEdit}
                              style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 8, border: "none", background: C.accent, color: "#fff", fontSize: 12.5, fontWeight: 600, opacity: savingEdit ? 0.6 : 1 }}
                            >
                              <Check size={13} /> บันทึก
                            </button>
                            <button
                              onClick={cancelEdit}
                              style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 8, border: `1px solid ${C.border}`, background: C.surface, color: C.inkMuted, fontSize: 12.5, fontWeight: 600 }}
                            >
                              <X size={13} /> ยกเลิก
                            </button>
                          </div>
                        </div>
                      );
                    }
                    return (
                      <div
                        key={t.rowIndex}
                        className="row-hover"
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 12,
                          background: C.surface,
                          border: `1px solid ${C.border}`,
                          borderRadius: 10,
                          padding: "12px 14px",
                          boxShadow: "0 1px 2px rgba(0,0,0,0.03)",
                          opacity: savingRow === t.rowIndex || isDeleting ? 0.5 : 1,
                        }}
                      >
                        <TypeBadge type={t.type} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 13.5, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.name}</div>
                          {t.detail && (
                            <div style={{ fontSize: 11.5, color: C.inkMuted, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                              {t.detail}
                            </div>
                          )}
                        </div>
                        <SignalBar value={t.progress} segments={5} height={11} />
                        <button onClick={() => cycleStatus(t)} disabled={savingRow === t.rowIndex}>
                          <StatusPill status={t.status} />
                        </button>
                        {t.link ? (
                          <a href={t.link} target="_blank" rel="noreferrer" style={{ color: C.inkFaint, display: "flex" }}>
                            <ExternalLink size={15} />
                          </a>
                        ) : (
                          <span style={{ width: 15 }} />
                        )}
                        <button onClick={() => startEdit(t)} title="แก้ไข" style={{ color: C.inkFaint, display: "flex", background: "none", border: "none" }}>
                          <Pencil size={15} />
                        </button>
                        <button onClick={() => confirmDelete(t)} disabled={isDeleting} title="ลบ" style={{ color: C.inkFaint, display: "flex", background: "none", border: "none" }}>
                          <Trash2 size={15} />
                        </button>
                      </div>
                    );
                  })}
                </div>
              )}
              <div style={{ fontSize: 11, color: C.inkFaint, marginTop: 14, textAlign: "center" }}>
                คลิกที่สถานะเพื่อเปลี่ยน Not start → Que → Done — บันทึกกลับไปที่ Google Sheet ทันที
              </div>
            </div>
          </>
        )}
      </main>
    </div>
  );
}
