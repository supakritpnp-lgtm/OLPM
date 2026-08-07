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

const TYPE_META = {
  CR: { label: "CR", icon: GitPullRequest, color: "#8B7CF6" },
  Request: { label: "Request", icon: MessageSquarePlus, color: "#3ED9C4" },
  Bug: { label: "Bug", icon: Bug, color: "#F2545B" },
  Sol: { label: "Sol", icon: Wrench, color: "#F2A93B" },
};

const STATUS_META = {
  "Not start": { color: "#5B6B87", bg: "rgba(91,107,135,0.15)" },
  Que: { color: "#F2A93B", bg: "rgba(242,169,59,0.15)" },
  Done: { color: "#3ED9C4", bg: "rgba(62,217,196,0.15)" },
};

const thb = (n) =>
  n ? n.toLocaleString("th-TH", { style: "currency", currency: "THB", maximumFractionDigits: 0 }) : "—";

const dmy = (s) => {
  if (!s) return "—";
  const d = new Date(s);
  if (isNaN(d)) return "—";
  return d.toLocaleDateString("th-TH", { day: "2-digit", month: "short", year: "2-digit" });
};

function SignalBar({ value = 0, segments = 8, height = 14 }) {
  const filled = Math.round(value * segments);
  return (
    <div style={{ display: "flex", gap: 2, alignItems: "flex-end" }}>
      {Array.from({ length: segments }).map((_, i) => {
        const isFilled = i < filled;
        const h = height * (0.45 + (i / segments) * 0.55);
        return (
          <div
            key={i}
            style={{
              width: 4,
              height: h,
              borderRadius: 1,
              background: isFilled ? (value >= 1 ? "#3ED9C4" : "#F2A93B") : "#223049",
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
        background: `${meta.color}1A`,
        border: `1px solid ${meta.color}33`,
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

export default function App() {
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

    // optimistic update
    setTasks((prev) =>
      prev.map((t) => (t.rowIndex === task.rowIndex ? { ...t, status: nextStatus, progress: nextProgress } : t))
    );
    setSavingRow(task.rowIndex);
    try {
      const res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" }, // avoids CORS preflight to Apps Script
        body: JSON.stringify({
          action: "updateTaskStatus",
          projectId: selectedId,
          rowIndex: task.rowIndex,
          status: nextStatus,
          progress: nextProgress,
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
    if (selectedId) loadTasks(selectedId);
    loadProjects();
  }

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans Thai', 'IBM Plex Sans', sans-serif",
        background: "#0B1220",
        color: "#E7ECF3",
        minHeight: "100vh",
        display: "flex",
        fontSize: 14,
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #223049; border-radius: 4px; }
        .row-hover:hover { background: #131C2E !important; }
        button { font-family: inherit; cursor: pointer; }
        input { font-family: inherit; }
        @keyframes spin { from { transform: rotate(0); } to { transform: rotate(360deg); } }
      `}</style>

      {/* Sidebar */}
      <aside
        style={{
          width: 300,
          minWidth: 300,
          borderRight: "1px solid #182238",
          display: "flex",
          flexDirection: "column",
          background: "#0E1626",
        }}
      >
        <div style={{ padding: "20px 18px 14px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
            <div
              style={{
                width: 8,
                height: 8,
                borderRadius: "50%",
                background: loadingProjects ? "#5B6B87" : "#3ED9C4",
                boxShadow: loadingProjects ? "none" : "0 0 8px #3ED9C4",
              }}
            />
            <span style={{ fontSize: 11, letterSpacing: 1.5, color: "#7C8AA5", fontFamily: "'IBM Plex Mono', monospace" }}>
              PROJECT CONSOLE
            </span>
          </div>
          <h1 style={{ fontSize: 19, fontWeight: 700, margin: "6px 0 0" }}>งานภายในทีม</h1>
        </div>

        <div style={{ padding: "0 18px 14px" }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              background: "#131C2E",
              border: "1px solid #223049",
              borderRadius: 8,
              padding: "8px 10px",
            }}
          >
            <Search size={15} color="#5B6B87" />
            <input
              placeholder="ค้นหาโปรเจกต์ / PM"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              style={{ background: "transparent", border: "none", outline: "none", color: "#E7ECF3", fontSize: 13, width: "100%" }}
            />
          </div>
        </div>

        <div style={{ flex: 1, overflowY: "auto", padding: "0 10px" }}>
          {loadingProjects && (
            <div style={{ color: "#5B6B87", fontSize: 12.5, padding: "16px 12px" }}>กำลังโหลดโปรเจกต์...</div>
          )}
          {!loadingProjects &&
            filteredProjects.map((p) => {
              const active = p.id === selectedId;
              return (
                <button
                  key={p.id}
                  onClick={() => setSelectedId(p.id)}
                  className="row-hover"
                  style={{
                    width: "100%",
                    textAlign: "left",
                    background: active ? "#152036" : "transparent",
                    border: active ? "1px solid #2A3B5C" : "1px solid transparent",
                    borderRadius: 10,
                    padding: "10px 12px",
                    marginBottom: 4,
                    display: "block",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#5B6B87" }}>{p.id}</span>
                    <SignalBar value={p.progress} segments={5} height={10} />
                  </div>
                  <div style={{ fontSize: 13.5, fontWeight: 600, marginTop: 3, color: "#E7ECF3" }}>{p.name}</div>
                  <div style={{ fontSize: 11.5, color: "#7C8AA5", marginTop: 2 }}>
                    PM: {p.pm} {p.platform ? `· ${p.platform}` : ""}
                  </div>
                </button>
              );
            })}
        </div>

        <div style={{ padding: 16, borderTop: "1px solid #182238" }}>
          <button
            onClick={handleSyncClick}
            style={{
              width: "100%",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 8,
              background: "#131C2E",
              border: "1px solid #223049",
              borderRadius: 8,
              padding: "9px 10px",
              color: "#3ED9C4",
              fontSize: 12.5,
              fontWeight: 600,
            }}
          >
            <RefreshCw size={13} style={{ animation: loadingProjects || loadingTasks ? "spin 0.8s linear infinite" : "none" }} />
            {loadingProjects || loadingTasks ? "กำลังซิงก์..." : "ซิงก์กับ Google Sheet"}
          </button>
          <div style={{ fontSize: 10.5, color: "#5B6B87", marginTop: 6, textAlign: "center", fontFamily: "'IBM Plex Mono', monospace" }}>
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
              background: "rgba(242,84,91,0.12)",
              borderBottom: "1px solid rgba(242,84,91,0.3)",
              color: "#F2545B",
              padding: "10px 32px",
              fontSize: 12.5,
            }}
          >
            <AlertTriangle size={14} />
            {error}
          </div>
        )}

        {!selected ? (
          <>
            <div style={{ display: "flex", gap: 24, padding: "22px 32px", borderBottom: "1px solid #182238" }}>
              {[
                { label: "โปรเจกต์ทั้งหมด", value: totals.projects },
                { label: "เสร็จสมบูรณ์", value: totals.done, color: "#3ED9C4" },
              ].map((s) => (
                <div key={s.label}>
                  <div style={{ fontSize: 26, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: s.color || "#E7ECF3" }}>
                    {s.value}
                  </div>
                  <div style={{ fontSize: 11.5, color: "#7C8AA5", marginTop: 2 }}>{s.label}</div>
                </div>
              ))}
            </div>

            <div style={{ flex: 1, overflow: "auto", padding: "8px 32px 32px" }}>
              {loadingProjects ? (
                <div style={{ color: "#5B6B87", padding: "60px 0", textAlign: "center" }}>กำลังโหลดข้อมูลจาก Google Sheet...</div>
              ) : filteredProjects.length === 0 ? (
                <div style={{ color: "#5B6B87", padding: "60px 0", textAlign: "center" }}>ไม่พบโปรเจกต์</div>
              ) : (
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ borderBottom: "1px solid #182238" }}>
                      {["Project", "PM", "Sales", "กำหนดส่ง", "มูลค่า", "ความคืบหน้า", ""].map((h) => (
                        <th
                          key={h}
                          style={{ textAlign: "left", padding: "12px 10px", fontSize: 11, letterSpacing: 0.5, color: "#5B6B87", fontWeight: 600, textTransform: "uppercase" }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredProjects.map((p) => (
                      <tr key={p.id} className="row-hover" onClick={() => setSelectedId(p.id)} style={{ borderBottom: "1px solid #131C2E", cursor: "pointer" }}>
                        <td style={{ padding: "13px 10px" }}>
                          <div style={{ fontWeight: 600 }}>{p.name}</div>
                          <div style={{ fontSize: 11, color: "#5B6B87", fontFamily: "'IBM Plex Mono', monospace" }}>{p.id}</div>
                        </td>
                        <td style={{ padding: "13px 10px", color: "#B7C1D4" }}>{p.pm}</td>
                        <td style={{ padding: "13px 10px", color: "#B7C1D4" }}>{p.sales}</td>
                        <td style={{ padding: "13px 10px", color: "#B7C1D4", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{dmy(p.due)}</td>
                        <td style={{ padding: "13px 10px", color: "#B7C1D4", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{thb(p.value)}</td>
                        <td style={{ padding: "13px 10px" }}>
                          <SignalBar value={p.progress} segments={8} />
                        </td>
                        <td style={{ padding: "13px 10px", textAlign: "right" }}>
                          <ChevronRight size={16} color="#5B6B87" />
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
            <div style={{ padding: "20px 32px 16px", borderBottom: "1px solid #182238" }}>
              <button
                onClick={() => setSelectedId(null)}
                style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: "#5B6B87", fontSize: 12.5, marginBottom: 10, padding: 0 }}
              >
                <ChevronLeft size={14} /> ภาพรวมทั้งหมด
              </button>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: "#5B6B87" }}>{selected.id}</div>
                  <h2 style={{ fontSize: 22, fontWeight: 700, margin: "4px 0 0" }}>{selected.name}</h2>
                  <div style={{ fontSize: 12.5, color: "#7C8AA5", marginTop: 4 }}>
                    PM {selected.pm} · Sales {selected.sales} · กำหนดส่ง {dmy(selected.due)}
                  </div>
                </div>
                <SignalBar value={selected.progress} segments={10} height={18} />
              </div>
            </div>

            <div style={{ display: "flex", gap: 6, padding: "14px 32px 0" }}>
              {["All", "CR", "Request", "Bug", "Sol"].map((t) => (
                <button
                  key={t}
                  onClick={() => setTypeFilter(t)}
                  style={{
                    padding: "6px 14px",
                    borderRadius: 999,
                    border: "1px solid " + (typeFilter === t ? "#3ED9C4" : "#223049"),
                    background: typeFilter === t ? "rgba(62,217,196,0.12)" : "transparent",
                    color: typeFilter === t ? "#3ED9C4" : "#7C8AA5",
                    fontSize: 12.5,
                    fontWeight: 600,
                  }}
                >
                  {t}
                </button>
              ))}
            </div>

            <div style={{ flex: 1, overflow: "auto", padding: "16px 32px 32px" }}>
              {loadingTasks ? (
                <div style={{ color: "#5B6B87", padding: "40px 0", textAlign: "center" }}>กำลังโหลดรายการงาน...</div>
              ) : filteredTasks.length === 0 ? (
                <div style={{ color: "#5B6B87", fontSize: 13, padding: "40px 0", textAlign: "center" }}>ยังไม่มีงานประเภทนี้</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {filteredTasks.map((t) => (
                    <div
                      key={t.rowIndex}
                      className="row-hover"
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 14,
                        background: "#0E1626",
                        border: "1px solid #182238",
                        borderRadius: 10,
                        padding: "12px 14px",
                        opacity: savingRow === t.rowIndex ? 0.6 : 1,
                      }}
                    >
                      <TypeBadge type={t.type} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13.5, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.name}</div>
                        {t.detail && (
                          <div style={{ fontSize: 11.5, color: "#5B6B87", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                            {t.detail}
                          </div>
                        )}
                      </div>
                      <SignalBar value={t.progress} segments={5} height={11} />
                      <button onClick={() => cycleStatus(t)} disabled={savingRow === t.rowIndex}>
                        <StatusPill status={t.status} />
                      </button>
                      {t.link ? (
                        <a href={t.link} target="_blank" rel="noreferrer" style={{ color: "#5B6B87", display: "flex" }}>
                          <ExternalLink size={15} />
                        </a>
                      ) : (
                        <span style={{ width: 15 }} />
                      )}
                    </div>
                  ))}
                </div>
              )}
              <div style={{ fontSize: 11, color: "#3A465F", marginTop: 14, textAlign: "center" }}>
                คลิกที่สถานะเพื่อเปลี่ยน Not start → Que → Done — บันทึกกลับไปที่ Google Sheet ทันที
              </div>
            </div>
          </>
        )}
      </main>
    </div>
  );
}
