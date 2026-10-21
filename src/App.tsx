import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
} from "react";
import OBSWebSocket from "obs-websocket-js";
import {
  buildRoomRecords,
  getRoomLevel,
  type RoomRecord,
  type WorkbookSheet,
} from "./lib/workbook";

type ObsSource = {
  label: string;
  inputName: string;
  value: (room: RoomRecord, roomRank: string) => string;
};

const OBS_SOURCES: ObsSource[] = [
  {
    label: "อันดับที่",
    inputName: "text_room_rank",
    value: (_room, roomRank) => roomRank,
  },
  {
    label: "ห้อง",
    inputName: "text_room",
    value: (room) => room.room,
  },
  {
    label: "นักเรียนคนที่ 1",
    inputName: "text_student_1",
    value: (room) => room.students[0] ?? "",
  },
  {
    label: "นักเรียนคนที่ 2",
    inputName: "text_student_2",
    value: (room) => room.students[1] ?? "",
  },
  {
    label: "ครูที่ปรึกษาคนที่ 1",
    inputName: "text_teacher_1",
    value: (room) => room.teachers[0] ?? "",
  },
  {
    label: "ครูที่ปรึกษาคนที่ 2",
    inputName: "text_teacher_2",
    value: (room) => room.teachers[1] ?? "",
  },
  {
    label: "ครูที่ปรึกษาคนที่ 3",
    inputName: "text_teacher_3",
    value: (room) => room.teachers[2] ?? "",
  },
] as const;

type LogEntry = {
  time: string;
  message: string;
};

function formatError(error: unknown): string {
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  return "เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ";
}

function App() {
  const obs = useRef<OBSWebSocket | null>(null);
  const [records, setRecords] = useState<RoomRecord[]>([]);
  const [selectedRoom, setSelectedRoom] = useState<string | null>(null);
  const [ranksByRoom, setRanksByRoom] = useState<Record<string, string>>({});
  const [level, setLevel] = useState("");
  const [search, setSearch] = useState("");
  const [fileName, setFileName] = useState("");
  const [isReading, setIsReading] = useState(false);
  const [host, setHost] = useState("localhost");
  const [port, setPort] = useState("4455");
  const [password, setPassword] = useState("");
  const [connected, setConnected] = useState(false);
  const [isConnecting, setIsConnecting] = useState(false);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [notice, setNotice] = useState("");
  const [isError, setIsError] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  const levels = useMemo(
    () =>
      [...new Set(records.map((record) => getRoomLevel(record.room)))].sort(
        (a, b) => Number(a) - Number(b),
      ),
    [records],
  );

  const visibleRecords = useMemo(() => {
    const query = search.trim().toLocaleLowerCase();
    return records.filter((record) => {
      const matchesLevel = !level || getRoomLevel(record.room) === level;
      const searchable = [
        record.room,
        ...record.students,
        ...record.teachers,
      ]
        .join(" ")
        .toLocaleLowerCase();
      return matchesLevel && (!query || searchable.includes(query));
    });
  }, [level, records, search]);

  const selectedRecord =
    records.find((record) => record.room === selectedRoom) ?? null;
  const selectedRoomRank = selectedRoom ? ranksByRoom[selectedRoom] ?? "" : "";
  const selectedVisibleIndex = visibleRecords.findIndex(
    (record) => record.room === selectedRoom,
  );

  function addLog(message: string) {
    setLogs((current) => [
      { time: new Date().toLocaleTimeString(), message },
      ...current,
    ].slice(0, 40));
  }

  function showNotice(message: string, error = false) {
    setNotice(message);
    setIsError(error);
  }

  function updateRoomRank(value: string) {
    if (!selectedRoom) return;
    setRanksByRoom((current) => ({
      ...current,
      [selectedRoom]: value,
    }));
  }

  useEffect(
    () => () => {
      void obs.current?.disconnect();
    },
    [],
  );

  async function loadWorkbook(file?: File) {
    if (!file) return;
    setIsReading(true);
    showNotice("");
    try {
      const XLSX = await import("@e965/xlsx");
      const workbook = XLSX.read(await file.arrayBuffer(), {
        type: "array",
        cellDates: false,
      });
      const sheets: WorkbookSheet[] = workbook.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], {
          header: 1,
          defval: "",
          raw: false,
        }),
      }));
      const nextRecords = buildRoomRecords(sheets);
      if (nextRecords.length === 0) {
        throw new Error(
          "ไม่พบข้อมูลห้อง ตรวจสอบว่ามีห้องรูปแบบ ม.1/2 ในไฟล์ Excel",
        );
      }

      setRecords(nextRecords);
      setSelectedRoom(nextRecords[0].room);
      setRanksByRoom({});
      setLevel("");
      setSearch("");
      setFileName(file.name);
      addLog(`โหลด ${file.name} สำเร็จ พบ ${nextRecords.length} ห้อง`);
      showNotice(`โหลดไฟล์สำเร็จ · พบ ${nextRecords.length} ห้อง`);
    } catch (error) {
      const message = `อ่านไฟล์ไม่สำเร็จ: ${formatError(error)}`;
      addLog(message);
      showNotice(message, true);
    } finally {
      setIsReading(false);
    }
  }

  async function toggleConnection() {
    if (connected) {
      try {
        await obs.current?.disconnect();
        obs.current = null;
        setConnected(false);
        addLog("ตัดการเชื่อมต่อ OBS แล้ว");
        showNotice("ตัดการเชื่อมต่อ OBS แล้ว");
      } catch (error) {
        const message = `ตัดการเชื่อมต่อไม่สำเร็จ: ${formatError(error)}`;
        addLog(message);
        showNotice(message, true);
      }
      return;
    }

    const parsedPort = Number(port);
    if (!host.trim() || !Number.isInteger(parsedPort) || parsedPort < 1 || parsedPort > 65535) {
      showNotice("กรุณาตรวจสอบ Host และ Port (1–65535)", true);
      return;
    }

    setIsConnecting(true);
    showNotice("");
    const client = new OBSWebSocket();
    obs.current = client;
    client.on("ConnectionClosed", () => {
      if (obs.current === client) {
        obs.current = null;
        setConnected(false);
        addLog("การเชื่อมต่อ OBS ถูกปิด");
      }
    });

    try {
      await client.connect(`ws://${host.trim()}:${parsedPort}`, password || undefined);
      setConnected(true);
      addLog(`เชื่อมต่อ OBS ที่ ${host.trim()}:${parsedPort} สำเร็จ`);
      showNotice("เชื่อมต่อ OBS แล้ว");
    } catch (error) {
      if (obs.current === client) obs.current = null;
      const message = `เชื่อมต่อ OBS ไม่สำเร็จ: ${formatError(error)}`;
      addLog(message);
      showNotice(message, true);
    } finally {
      setIsConnecting(false);
    }
  }

  async function pushSelected() {
    const client = obs.current;
    if (!client || !connected || !selectedRecord) return;

    try {
      const { inputs } = await client.call("GetInputList");
      const hasRoomRankSource = inputs.some(
        (input) => input.inputName === "text_room_rank",
      );
      if (!hasRoomRankSource) {
        const textSourceTemplate = inputs.find(
          (input) =>
            input.inputName === "text_room" &&
            typeof input.inputKind === "string",
        );
        if (!textSourceTemplate || typeof textSourceTemplate.inputKind !== "string") {
          throw new Error(
            "สร้าง Text Source อันดับไม่ได้: ไม่พบ Source text_room ให้สร้าง Text Source ชื่อ text_room ก่อน",
          );
        }

        const { currentProgramSceneName } =
          await client.call("GetCurrentProgramScene");
        await client.call("CreateInput", {
          sceneName: currentProgramSceneName,
          inputName: "text_room_rank",
          inputKind: textSourceTemplate.inputKind,
          inputSettings: {
            text: selectedRoomRank,
          },
          sceneItemEnabled: true,
        });
        addLog(`สร้าง Text Source text_room_rank ในฉาก ${currentProgramSceneName}`);
      }

      for (const source of OBS_SOURCES) {
        await client.call("SetInputSettings", {
          inputName: source.inputName,
          inputSettings: { text: source.value(selectedRecord, selectedRoomRank) },
          overlay: true,
        });
      }
      const message = `ส่งข้อมูลห้อง ${selectedRecord.room} และอันดับห้องไป OBS สำเร็จ`;
      addLog(message);
      showNotice(message);
    } catch (error) {
      const message = `ส่งข้อมูลไป OBS ไม่สำเร็จ: ${formatError(error)} ตรวจสอบชื่อ Text Source`;
      addLog(message);
      showNotice(message, true);
    }
  }

  function selectRelative(offset: -1 | 1) {
    if (visibleRecords.length === 0) return;
    const startIndex =
      selectedVisibleIndex >= 0
        ? selectedVisibleIndex
        : offset > 0
          ? -1
          : 0;
    const nextIndex =
      (startIndex + offset + visibleRecords.length) % visibleRecords.length;
    setSelectedRoom(visibleRecords[nextIndex].room);
  }

  function handleTableKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key === "ArrowDown" || event.key === "ArrowRight") {
      event.preventDefault();
      selectRelative(1);
    }
    if (event.key === "ArrowUp" || event.key === "ArrowLeft") {
      event.preventDefault();
      selectRelative(-1);
    }
  }

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="topbar">
          <div className="brand">
            <div className="brand-mark" aria-hidden="true">O</div>
            <div>
              <h1>OBS Text Display</h1>
              <p>จัดการรายชื่อห้องเรียน · Excel → OBS</p>
            </div>
          </div>
          <div className="header-actions">
            <div className="selection-indicator">
              <span>ห้องที่เลือก</span>
              <strong>{selectedRoom ?? "—"}</strong>
            </div>
            <button
              className="icon-button"
              onClick={() => selectRelative(-1)}
              disabled={!visibleRecords.length}
              aria-label="ห้องก่อนหน้า"
              title="ห้องก่อนหน้า"
            >
              ←
            </button>
            <button
              className="icon-button"
              onClick={() => selectRelative(1)}
              disabled={!visibleRecords.length}
              aria-label="ห้องถัดไป"
              title="ห้องถัดไป"
            >
              →
            </button>
            <button
              className="button button-primary"
              onClick={() => void pushSelected()}
              disabled={!connected || !selectedRecord}
            >
              <span aria-hidden="true">↗</span> ส่งไป OBS
            </button>
          </div>
        </header>

        <div className="content-grid">
          <section className="roster-panel" aria-label="รายชื่อห้องเรียน">
            <div className="panel-heading">
              <div>
                <span className="eyebrow">CLASS ROSTER</span>
                <h2>รายชื่อห้องเรียน</h2>
              </div>
              {records.length > 0 && (
                <span className="count-pill">{visibleRecords.length} ห้อง</span>
              )}
            </div>

            <div className="file-picker">
              <div className="file-icon" aria-hidden="true">↥</div>
              <div className="file-copy">
                <strong>{fileName || "นำเข้ารายชื่อจาก Excel"}</strong>
                <span>
                  {fileName
                    ? `${records.length} ห้องพร้อมใช้งาน`
                    : "รองรับไฟล์ .xlsx และ .xls"}
                </span>
              </div>
              <input
                ref={fileInput}
                className="visually-hidden"
                type="file"
                accept=".xlsx,.xls"
                onChange={(event) => { const file = event.currentTarget.files?.[0]; event.currentTarget.value = ""; void loadWorkbook(file); }}
                aria-label="เลือกไฟล์ Excel"
              />
              <button
                className="button button-secondary"
                onClick={() => fileInput.current?.click()}
                disabled={isReading}
              >
                {isReading ? "กำลังอ่าน..." : "เลือกไฟล์"}
              </button>
            </div>

            <div className="filters">
              <label className="filter-field level-field">
                <span>ระดับชั้น</span>
                <select
                  value={level}
                  onChange={(event) => setLevel(event.target.value)}
                  disabled={!levels.length}
                >
                  <option value="">ทุกระดับชั้น</option>
                  {levels.map((item) => (
                    <option value={item} key={item}>มัธยมศึกษาปีที่ {item}</option>
                  ))}
                </select>
              </label>
              <label className="filter-field search-field">
                <span>ค้นหา</span>
                <div className="search-box">
                  <span aria-hidden="true">⌕</span>
                  <input
                    type="search"
                    value={search}
                    onChange={(event) => setSearch(event.target.value)}
                    placeholder="ค้นหาห้อง นักเรียน หรือครู"
                  />
                </div>
              </label>
            </div>

            <div
              className="table-wrap"
              tabIndex={0}
              onKeyDown={handleTableKeyDown}
              aria-label="ตารางห้องเรียน ใช้ปุ่มลูกศรเพื่อเปลี่ยนห้องที่เลือก"
            >
              {records.length === 0 ? (
                <div className="empty-state">
                  <div className="empty-icon" aria-hidden="true">▤</div>
                  <strong>เริ่มต้นด้วยการเลือกไฟล์ Excel</strong>
                  <p>ระบบจะจับคู่รายชื่อนักเรียนและครูที่ปรึกษาตามห้องให้อัตโนมัติ</p>
                  <button
                    className="button button-secondary"
                    onClick={() => fileInput.current?.click()}
                  >
                    เลือกไฟล์รายชื่อ
                  </button>
                </div>
              ) : visibleRecords.length === 0 ? (
                <div className="empty-state compact-empty">
                  <strong>ไม่พบห้องที่ตรงกับการค้นหา</strong>
                  <p>ลองเปลี่ยนคำค้นหาหรือเลือกระดับชั้นอื่น</p>
                </div>
              ) : (
                <table>
                  <thead>
                    <tr>
                      <th className="room-column">ห้องเรียน</th>
                      <th>นักเรียน</th>
                      <th>ครูที่ปรึกษา</th>
                      <th className="arrow-column" aria-label="เลือกห้อง" />
                    </tr>
                  </thead>
                  <tbody>
                    {visibleRecords.map((record) => {
                      const selected = record.room === selectedRoom;
                      return (
                        <tr
                          key={record.room}
                          className={selected ? "selected-row" : ""}
                          onClick={() => setSelectedRoom(record.room)}
                          aria-selected={selected}
                        >
                          <td>
                            <span className="room-badge">{record.room}</span>
                          </td>
                          <td>
                            <span className="cell-primary">
                              {record.students[0] ?? "—"}
                            </span>
                            <span className="cell-secondary">
                              {record.students[1] ?? "ยังไม่มีรายชื่อคนที่ 2"}
                            </span>
                          </td>
                          <td>
                            <span className="cell-primary">
                              {record.teachers[0] ?? "—"}
                            </span>
                            <span className="cell-secondary">
                              {record.teachers.length > 1
                                ? record.teachers.slice(1).join(", ")
                                : "ยังไม่มีข้อมูลครู"}
                            </span>
                          </td>
                          <td className="row-arrow" aria-hidden="true">→</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            {selectedRecord && (
              <section className="preview-panel" aria-label="ตัวอย่างข้อมูลที่จะส่งไป OBS">
                <div className="preview-heading">
                  <div>
                    <span className="eyebrow">LIVE PREVIEW</span>
                    <h3>ตัวอย่างข้อมูลที่จะส่ง</h3>
                  </div>
                  <span className="preview-room">{selectedRecord.room}</span>
                </div>
                <div className="preview-grid">
                  {OBS_SOURCES.map((source) => (
                    <div className="preview-item" key={source.inputName}>
                      <span>{source.label}</span>
                      {source.inputName === "text_room_rank" ? (
                        <input
                          type="text"
                          value={selectedRoomRank}
                          onChange={(event) => updateRoomRank(event.target.value)}
                          placeholder="กรอกอันดับของห้องนี้"
                          aria-label={`อันดับของห้อง ${selectedRecord.room}`}
                        />
                      ) : (
                        <strong>
                          {source.value(selectedRecord, selectedRoomRank) || "—"}
                        </strong>
                      )}
                    </div>
                  ))}
                </div>
              </section>
            )}
          </section>

          <aside className="sidebar">
            <section className="side-card connection-card">
              <div className="side-title">
                <div className="side-icon" aria-hidden="true">⌁</div>
                <div>
                  <h2>การเชื่อมต่อ OBS</h2>
                  <p>OBS WebSocket 5</p>
                </div>
              </div>

              <label className="input-label" htmlFor="obs-host">Host</label>
              <input
                id="obs-host"
                value={host}
                onChange={(event) => setHost(event.target.value)}
                placeholder="localhost"
                autoComplete="url"
                disabled={connected}
              />
              <div className="split-fields">
                <label>
                  <span className="input-label">Port</span>
                  <input
                    type="number"
                    min="1"
                    max="65535"
                    value={port}
                    onChange={(event) => setPort(event.target.value)}
                    disabled={connected}
                  />
                </label>
                <label>
                  <span className="input-label">Password</span>
                  <input
                    type="password"
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                    placeholder="ถ้ามี"
                    autoComplete="current-password"
                    disabled={connected}
                  />
                </label>
              </div>

              <button
                className={`button connection-button ${connected ? "button-disconnect" : "button-primary"}`}
                onClick={() => void toggleConnection()}
                disabled={isConnecting}
              >
                <span className={`status-dot ${connected ? "is-online" : ""}`} />
                {isConnecting
                  ? "กำลังเชื่อมต่อ..."
                  : connected
                    ? "ตัดการเชื่อมต่อ"
                    : "เชื่อมต่อกับ OBS"}
              </button>
              <p className={`connection-status ${connected ? "online-text" : ""}`}>
                {connected ? "เชื่อมต่อแล้ว · พร้อมส่งข้อมูล" : "ยังไม่ได้เชื่อมต่อ"}
              </p>
            </section>

            <section className="side-card source-card">
              <div className="side-title">
                <div className="side-icon source-icon" aria-hidden="true">▤</div>
                <div>
                  <h2>OBS Text Sources</h2>
                  <p>ชื่อแหล่งข้อความใน OBS</p>
                </div>
              </div>
              <div className="source-list">
                {OBS_SOURCES.map((source, index) => (
                  <div className="source-row" key={source.inputName}>
                    <span className="source-index">{String(index + 1).padStart(2, "0")}</span>
                    <span className="source-label">{source.label}</span>
                    <code>{source.inputName}</code>
                  </div>
                ))}
              </div>
              <p className="source-hint">
                สร้าง Text Source ใน OBS ให้ตรงกับชื่อด้านบน ข้อมูลที่ไม่มีจะถูกล้างเป็นข้อความว่าง
              </p>
            </section>

            <section className="side-card log-card">
              <div className="log-heading">
                <div className="side-title">
                  <div className="side-icon log-icon" aria-hidden="true">≋</div>
                  <div>
                    <h2>บันทึกการทำงาน</h2>
                    <p>สถานะและกิจกรรมล่าสุด</p>
                  </div>
                </div>
                <span className="log-count">{logs.length}</span>
              </div>
              <div className="log-list" aria-live="polite">
                {logs.length === 0 ? (
                  <p className="no-logs">กิจกรรมจะแสดงที่นี่</p>
                ) : (
                  logs.map((entry, index) => (
                    <div className="log-entry" key={`${entry.time}-${index}`}>
                      <span className="log-time">{entry.time}</span>
                      <span>{entry.message}</span>
                    </div>
                  ))
                )}
              </div>
            </section>
          </aside>
        </div>

        {notice && (
          <div
            className={`notice ${isError ? "notice-error" : ""}`}
            role={isError ? "alert" : "status"}
          >
            <span aria-hidden="true">{isError ? "!" : "✓"}</span>
            {notice}
            <button
              onClick={() => showNotice("")}
              aria-label="ปิดข้อความแจ้งเตือน"
            >
              ×
            </button>
          </div>
        )}
        <footer className="footer">
          <span>ประมวลผลไฟล์ในเบราว์เซอร์ของคุณ · ไม่ส่งข้อมูลออกจากเครื่อง</span>
          <span>OBS WebSocket 5 <i /> พร้อมใช้งาน</span>
        </footer>
      </section>
    </main>
  );
}

export default App;
