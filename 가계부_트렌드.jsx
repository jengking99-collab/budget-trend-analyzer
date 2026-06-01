import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ReferenceLine
} from "recharts";

/* ── 엑셀 시리얼 날짜 → JS Date ── */
function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  return new Date(utcDays * 86400 * 1000);
}

function toYearMonth(date) {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function formatLabel(ym) {
  const [y, m] = ym.split("-");
  return `${y.slice(2)}.${m}`;
}

const PALETTE = [
  "#38bdf8","#34d399","#fb923c","#f472b6",
  "#a78bfa","#fbbf24","#f87171","#4ade80",
  "#e879f9","#22d3ee","#818cf8","#facc15",
];

/* ── 금액 포맷터 ── */
function fmtAmt(v) {
  if (v >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`;
  if (v >= 1_000)     return `${(v / 1_000).toFixed(0)}K`;
  return String(v);
}

/* ── 커스텀 툴팁 ── */
function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: "#0d1b2a", border: "1px solid #1e3a5f",
      borderRadius: 10, padding: "12px 16px", minWidth: 180,
      boxShadow: "0 8px 32px rgba(0,0,0,0.5)"
    }}>
      <p style={{ margin: "0 0 8px", color: "#94a3b8", fontSize: 12 }}>
        📅 {label.replace("-", "년 ")}월
      </p>
      {payload.map(p => (
        <div key={p.dataKey} style={{ display: "flex", justifyContent: "space-between", gap: 16, marginBottom: 4 }}>
          <span style={{ color: p.color, fontSize: 13, fontWeight: 600 }}>{p.dataKey}</span>
          <span style={{ color: "#e2e8f0", fontSize: 13 }}>
            {Number(p.value).toLocaleString()}원
          </span>
        </div>
      ))}
    </div>
  );
}

export default function App() {
  const [transactions, setTransactions] = useState(null);
  const [allCategories, setAllCategories] = useState([]);
  const [catsByType, setCatsByType] = useState({ 지출: [], 수입: [] });
  const [selectedCats, setSelectedCats] = useState([]);
  const [startMonth, setStartMonth] = useState("");
  const [endMonth,   setEndMonth]   = useState("");
  const [yMaxInput, setYMaxInput]   = useState("");
  const [yMax,      setYMax]        = useState(null);
  const [dragOver,  setDragOver]    = useState(false);
  const [typeFilter, setTypeFilter] = useState("지출"); // 지출 | 수입 | 전체
  const [viewMode, setViewMode] = useState("월별"); // 월별 | 누적
  const [selectedPoint, setSelectedPoint] = useState(null); // { ym, cat }
  const detailRef = useRef(null);
  const fileRef = useRef();

  /* ── 파일 파싱 ── */
  const parseFile = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const header = rows[0];
      const idx = (name) => header.indexOf(name);
      const iDate    = idx("날짜");
      const iCat     = idx("분류");
      const iAmt     = idx("금액(원)");
      const iType    = idx("수입/지출");
      const iAsset   = idx("자산");
      const iSubCat  = idx("소분류");
      const iMemo    = idx("내용");

      const parsed = [];
      const expendSet = new Set();
      const incomeSet = new Set();

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row[iDate] || typeof row[iDate] !== "number") continue;
        const date  = excelSerialToDate(row[iDate]);
        const ym    = toYearMonth(date);
        const dateStr = `${date.getUTCFullYear()}-${String(date.getUTCMonth()+1).padStart(2,"0")}-${String(date.getUTCDate()).padStart(2,"0")}`;
        const cat   = row[iCat] || "기타";
        const amt   = Number(row[iAmt]) || 0;
        const type  = row[iType] || "";
        const asset  = iAsset  >= 0 ? (row[iAsset]  || "") : "";
        const subCat = iSubCat >= 0 ? (row[iSubCat] || "") : "";
        const memo   = iMemo   >= 0 ? (row[iMemo]   || "") : "";
        if (type === "수입") incomeSet.add(cat);
        else expendSet.add(cat);
        parsed.push({ ym, dateStr, cat, amt, type, asset, subCat, memo });
      }

      const expendCats = Array.from(expendSet).sort();
      const incomeCats = Array.from(incomeSet).sort();
      const cats = [...expendCats, ...incomeCats.filter(c => !expendSet.has(c))];
      const yms  = parsed.map(t => t.ym).sort();

      setTransactions(parsed);
      setAllCategories(cats);
      setCatsByType({ 지출: expendCats, 수입: incomeCats });
      // 기본 선택: 지출 카테고리 중 상위 5개 (금액 합 기준)
      const catTotals = {};
      parsed.forEach(t => { catTotals[t.cat] = (catTotals[t.cat] || 0) + t.amt; });
      const top5 = [...expendCats].sort((a,b) => (catTotals[b]||0) - (catTotals[a]||0)).slice(0, 5);
      setSelectedCats(top5);
      setStartMonth(yms[0] || "");
      setEndMonth(yms[yms.length - 1] || "");
    };
    reader.readAsBinaryString(file);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragOver(false);
    const f = e.dataTransfer.files[0];
    if (f) parseFile(f);
  }, [parseFile]);

  /* ── 차트 데이터 계산 ── */
  const chartData = useMemo(() => {
    if (!transactions) return [];
    const filtered = transactions.filter(t => {
      if (startMonth && t.ym < startMonth) return false;
      if (endMonth   && t.ym > endMonth)   return false;
      if (typeFilter === "지출" && t.type !== "지출" && t.type !== "이체지출") return false;
      if (typeFilter === "수입" && t.type !== "수입") return false;
      return true;
    });
    const months = Array.from(new Set(filtered.map(t => t.ym))).sort();
    const monthly = months.map(ym => {
      const obj = { month: ym, label: formatLabel(ym) };
      for (const cat of selectedCats) {
        obj[cat] = filtered
          .filter(t => t.ym === ym && t.cat === cat)
          .reduce((s, t) => s + t.amt, 0);
      }
      return obj;
    });
    if (viewMode === "월별") return monthly;
    // 누적 모드: 각 카테고리를 앞에서부터 합산
    const acc = {};
    return monthly.map(row => {
      const obj = { month: row.month, label: row.label };
      for (const cat of selectedCats) {
        acc[cat] = (acc[cat] || 0) + (row[cat] || 0);
        obj[cat] = acc[cat];
      }
      return obj;
    });
  }, [transactions, startMonth, endMonth, selectedCats, typeFilter, viewMode]);

  const toggleCat = (cat) =>
    setSelectedCats(p => p.includes(cat) ? p.filter(c => c !== cat) : [...p, cat]);

  const applyYMax = () => {
    const raw = String(yMaxInput).replace(/,/g, "").trim();
    const v = raw === "" ? 0 : Number(raw);
    setYMax(!isNaN(v) && v > 0 ? v : null);
  };

  /* ──────────── UI ──────────── */
  if (!transactions) {
    return (
      <div style={{
        minHeight: "100vh", background: "#060d18",
        display: "flex", flexDirection: "column",
        alignItems: "center", justifyContent: "center",
        fontFamily: "'Noto Sans KR', sans-serif"
      }}>
        {/* 배경 그리드 */}
        <div style={{
          position: "fixed", inset: 0, zIndex: 0,
          backgroundImage: "linear-gradient(#1e3a5f1a 1px, transparent 1px), linear-gradient(90deg, #1e3a5f1a 1px, transparent 1px)",
          backgroundSize: "40px 40px"
        }} />

        <div style={{ position: "relative", zIndex: 1, textAlign: "center", padding: 32 }}>
          <div style={{
            fontSize: 52, marginBottom: 4,
            filter: "drop-shadow(0 0 20px #38bdf855)"
          }}>📊</div>
          <h1 style={{
            margin: "0 0 8px", fontSize: 32, fontWeight: 800,
            background: "linear-gradient(135deg, #38bdf8, #818cf8)",
            WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent"
          }}>가계부 트렌드 분석기</h1>
          <p style={{ color: "#475569", marginBottom: 40, fontSize: 15 }}>
            엑셀 파일을 업로드하면 월별 지출 트렌드를 시각화해드립니다
          </p>

          <div
            onDrop={handleDrop}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onClick={() => fileRef.current.click()}
            style={{
              width: 380, border: `2px dashed ${dragOver ? "#38bdf8" : "#1e3a5f"}`,
              borderRadius: 20, padding: "48px 32px", cursor: "pointer",
              background: dragOver
                ? "linear-gradient(135deg, #0c1e3388, #0c1e33cc)"
                : "linear-gradient(135deg, #0a1628cc, #0d1f38cc)",
              backdropFilter: "blur(12px)",
              transition: "all 0.2s",
              boxShadow: dragOver ? "0 0 40px #38bdf822" : "0 0 0 #0000"
            }}
          >
            <div style={{ fontSize: 40, marginBottom: 16 }}>
              {dragOver ? "🎯" : "📂"}
            </div>
            <p style={{ color: "#cbd5e1", fontWeight: 600, fontSize: 16, margin: "0 0 8px" }}>
              파일을 드래그하거나 클릭하여 업로드
            </p>
            <p style={{ color: "#334155", fontSize: 13, margin: 0 }}>
              .xlsx / .xls 파일 지원
            </p>
          </div>
          <input
            ref={fileRef} type="file" accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={e => { const f = e.target.files[0]; if (f) parseFile(f); }}
          />
        </div>
      </div>
    );
  }

  /* ── 메인 대시보드 ── */
  return (
    <div style={{
      minHeight: "100vh", background: "#060d18",
      fontFamily: "'Noto Sans KR', sans-serif", color: "#e2e8f0"
    }}>
      {/* 배경 그리드 */}
      <div style={{
        position: "fixed", inset: 0, zIndex: 0, pointerEvents: "none",
        backgroundImage: "linear-gradient(#1e3a5f0e 1px, transparent 1px), linear-gradient(90deg, #1e3a5f0e 1px, transparent 1px)",
        backgroundSize: "40px 40px"
      }} />

      {/* 헤더 */}
      <div style={{
        position: "relative", zIndex: 1,
        padding: "18px 28px", display: "flex",
        alignItems: "center", justifyContent: "space-between",
        borderBottom: "1px solid #0f2035",
        background: "linear-gradient(180deg, #0a1628ee, transparent)"
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 22 }}>📊</span>
          <h1 style={{
            margin: 0, fontSize: 20, fontWeight: 800,
            background: "linear-gradient(135deg, #38bdf8, #818cf8)",
            WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent"
          }}>가계부 트렌드 분석기</h1>
          <span style={{
            fontSize: 11, fontWeight: 700, color: "#334155",
            border: "1px solid #1e3a5f", borderRadius: 6,
            padding: "2px 7px", marginLeft: 4
          }}>v1.3.2</span>
        </div>
        <button
          onClick={() => { setTransactions(null); setAllCategories([]); }}
          style={{
            padding: "6px 14px", borderRadius: 8,
            background: "#0f2035", color: "#64748b",
            border: "1px solid #1e3a5f", cursor: "pointer", fontSize: 13
          }}
        >
          ← 파일 변경
        </button>
      </div>

      <div style={{ position: "relative", zIndex: 1, padding: "20px 28px", display: "flex", flexDirection: "column", gap: 18 }}>

        {/* ── 컨트롤 패널 ── */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 14 }}>

          {/* 기간 설정 */}
          <div style={{ background: "#0a1628", border: "1px solid #0f2035", borderRadius: 14, padding: 18 }}>
            <p style={{ margin: "0 0 12px", fontSize: 12, color: "#475569", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              📅 분석 기간
            </p>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input
                type="month" value={startMonth}
                onChange={e => setStartMonth(e.target.value)}
                style={{
                  flex: 1, padding: "7px 10px", borderRadius: 8,
                  border: "1px solid #1e3a5f", background: "#060d18",
                  color: "#cbd5e1", fontSize: 13, outline: "none"
                }}
              />
              <span style={{ color: "#1e3a5f", fontSize: 16 }}>–</span>
              <input
                type="month" value={endMonth}
                onChange={e => setEndMonth(e.target.value)}
                style={{
                  flex: 1, padding: "7px 10px", borderRadius: 8,
                  border: "1px solid #1e3a5f", background: "#060d18",
                  color: "#cbd5e1", fontSize: 13, outline: "none"
                }}
              />
            </div>
          </div>

          {/* 수입/지출 필터 */}
          <div style={{ background: "#0a1628", border: "1px solid #0f2035", borderRadius: 14, padding: 18 }}>
            <p style={{ margin: "0 0 12px", fontSize: 12, color: "#475569", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              🔀 유형 필터
            </p>
            <div style={{ display: "flex", gap: 6 }}>
              {["지출", "수입", "전체"].map(t => (
                <button
                  key={t}
                  onClick={() => setTypeFilter(t)}
                  style={{
                    flex: 1, padding: "7px 0", borderRadius: 8,
                    border: `1px solid ${typeFilter === t ? "#38bdf8" : "#1e3a5f"}`,
                    background: typeFilter === t ? "#0c2d47" : "transparent",
                    color: typeFilter === t ? "#38bdf8" : "#475569",
                    cursor: "pointer", fontSize: 13, fontWeight: typeFilter === t ? 700 : 400,
                    transition: "all 0.15s"
                  }}
                >{t}</button>
              ))}
            </div>
          </div>

          {/* 누적 / 월별 */}
          <div style={{ background: "#0a1628", border: "1px solid #0f2035", borderRadius: 14, padding: 18 }}>
            <p style={{ margin: "0 0 12px", fontSize: 12, color: "#475569", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              📈 합계 방식
            </p>
            <div style={{ display: "flex", gap: 6 }}>
              {["월별", "누적"].map(m => (
                <button
                  key={m}
                  onClick={() => setViewMode(m)}
                  style={{
                    flex: 1, padding: "7px 0", borderRadius: 8,
                    border: `1px solid ${viewMode === m ? "#38bdf8" : "#1e3a5f"}`,
                    background: viewMode === m ? "#0c2d47" : "transparent",
                    color: viewMode === m ? "#38bdf8" : "#475569",
                    cursor: "pointer", fontSize: 13, fontWeight: viewMode === m ? 700 : 400,
                    transition: "all 0.15s"
                  }}
                >{m} 합계</button>
              ))}
            </div>
          </div>

          {/* Y축 최대값 */}
          <div style={{ background: "#0a1628", border: "1px solid #0f2035", borderRadius: 14, padding: 18 }}>
            <p style={{ margin: "0 0 12px", fontSize: 12, color: "#475569", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              📏 Y축 최대값
            </p>
            <div style={{ display: "flex", gap: 6 }}>
              <input
                type="text" value={yMaxInput}
                onChange={e => setYMaxInput(e.target.value)}
                placeholder="자동"
                style={{
                  flex: 1, padding: "7px 10px", borderRadius: 8,
                  border: "1px solid #1e3a5f", background: "#060d18",
                  color: "#cbd5e1", fontSize: 13, outline: "none"
                }}
                onKeyDown={e => { if (e.key === "Enter") applyYMax(); }}
              />
              <button
                onClick={applyYMax}
                style={{
                  padding: "7px 12px", borderRadius: 8, cursor: "pointer",
                  background: "linear-gradient(135deg, #1d4ed8, #4338ca)",
                  color: "#fff", border: "none", fontWeight: 700, fontSize: 13
                }}
              >적용</button>
              <button
                onClick={() => { setYMax(null); setYMaxInput(""); }}
                style={{
                  padding: "7px 10px", borderRadius: 8, cursor: "pointer",
                  background: "#0f2035", color: "#475569",
                  border: "1px solid #1e3a5f", fontSize: 13
                }}
              >↺</button>
            </div>
            {yMax && (
              <p style={{ margin: "8px 0 0", fontSize: 11, color: "#38bdf8" }}>
                최대: {yMax.toLocaleString()}원
              </p>
            )}
          </div>
        </div>

        {/* ── 카테고리 선택 ── */}
        <div style={{ background: "#0a1628", border: "1px solid #0f2035", borderRadius: 14, padding: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <p style={{ margin: 0, fontSize: 12, color: "#475569", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>
              🏷️ 카테고리 선택 ({selectedCats.length}/{allCategories.length})
            </p>
            <div style={{ display: "flex", gap: 6 }}>
              <button
                onClick={() => setSelectedCats([...allCategories])}
                style={{ padding: "4px 10px", borderRadius: 6, cursor: "pointer", background: "#0c2d47", color: "#38bdf8", border: "1px solid #1e3a5f", fontSize: 11 }}
              >전체 선택</button>
              <button
                onClick={() => setSelectedCats([])}
                style={{ padding: "4px 10px", borderRadius: 6, cursor: "pointer", background: "#1a0a0a", color: "#f87171", border: "1px solid #3f1414", fontSize: 11 }}
              >전체 해제</button>
            </div>
          </div>

          {/* 지출 카테고리 그룹 */}
          {catsByType.지출.length > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                <span style={{ fontSize: 11, fontWeight: 700, color: "#f87171", letterSpacing: "0.06em" }}>💸 지출</span>
                <div style={{ flex: 1, height: 1, background: "#1e3a5f" }} />
                <button
                  onClick={() => setSelectedCats(p => [...new Set([...p, ...catsByType.지출])])}
                  style={{ padding: "2px 8px", borderRadius: 5, cursor: "pointer", background: "transparent", color: "#475569", border: "1px solid #1e3a5f", fontSize: 10 }}
                >전체 선택</button>
                <button
                  onClick={() => setSelectedCats(p => p.filter(c => !catsByType.지출.includes(c)))}
                  style={{ padding: "2px 8px", borderRadius: 5, cursor: "pointer", background: "transparent", color: "#475569", border: "1px solid #1e3a5f", fontSize: 10 }}
                >전체 해제</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
                {catsByType.지출.map((cat) => {
                  const colorIdx = allCategories.indexOf(cat) % PALETTE.length;
                  const color = PALETTE[colorIdx];
                  const sel = selectedCats.includes(cat);
                  return (
                    <button key={cat} onClick={() => toggleCat(cat)} style={{
                      padding: "5px 13px", borderRadius: 20, cursor: "pointer",
                      border: `1.5px solid ${sel ? color : "#1e3a5f"}`,
                      background: sel ? `${color}18` : "transparent",
                      color: sel ? color : "#334155",
                      fontSize: 12, fontWeight: sel ? 700 : 400,
                      transition: "all 0.15s",
                      boxShadow: sel ? `0 0 8px ${color}33` : "none"
                    }}>
                      {sel && <span style={{ marginRight: 4, fontSize: 9 }}>●</span>}
                      {cat}
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* 수입 카테고리 그룹 */}
          {catsByType.수입.length > 0 && (
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                <span style={{ fontSize: 11, fontWeight: 700, color: "#34d399", letterSpacing: "0.06em" }}>💰 수입</span>
                <div style={{ flex: 1, height: 1, background: "#1e3a5f" }} />
                <button
                  onClick={() => setSelectedCats(p => [...new Set([...p, ...catsByType.수입])])}
                  style={{ padding: "2px 8px", borderRadius: 5, cursor: "pointer", background: "transparent", color: "#475569", border: "1px solid #1e3a5f", fontSize: 10 }}
                >전체 선택</button>
                <button
                  onClick={() => setSelectedCats(p => p.filter(c => !catsByType.수입.includes(c)))}
                  style={{ padding: "2px 8px", borderRadius: 5, cursor: "pointer", background: "transparent", color: "#475569", border: "1px solid #1e3a5f", fontSize: 10 }}
                >전체 해제</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
                {catsByType.수입.map((cat) => {
                  const colorIdx = allCategories.indexOf(cat) % PALETTE.length;
                  const color = PALETTE[colorIdx];
                  const sel = selectedCats.includes(cat);
                  return (
                    <button key={cat} onClick={() => toggleCat(cat)} style={{
                      padding: "5px 13px", borderRadius: 20, cursor: "pointer",
                      border: `1.5px solid ${sel ? color : "#1e3a5f"}`,
                      background: sel ? `${color}18` : "transparent",
                      color: sel ? color : "#334155",
                      fontSize: 12, fontWeight: sel ? 700 : 400,
                      transition: "all 0.15s",
                      boxShadow: sel ? `0 0 8px ${color}33` : "none"
                    }}>
                      {sel && <span style={{ marginRight: 4, fontSize: 9 }}>●</span>}
                      {cat}
                    </button>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* ── 차트 ── */}
        <div style={{ background: "#0a1628", border: "1px solid #0f2035", borderRadius: 14, padding: "20px 24px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
            <div>
              <h2 style={{ margin: 0, fontSize: 17, fontWeight: 800, color: "#f1f5f9" }}>{viewMode} 금액 트렌드</h2>
              <p style={{ margin: "3px 0 0", fontSize: 12, color: "#334155" }}>
                {startMonth} ~ {endMonth} · {chartData.length}개월
              </p>
            </div>
            <div style={{
              padding: "4px 12px", borderRadius: 20,
              background: "#0c2d47", color: "#38bdf8",
              fontSize: 12, fontWeight: 600, border: "1px solid #1e3a5f"
            }}>
              {typeFilter} 기준
            </div>
          </div>

          {chartData.length === 0 || selectedCats.length === 0 ? (
            <div style={{ textAlign: "center", padding: "60px 0", color: "#1e3a5f" }}>
              <div style={{ fontSize: 36, marginBottom: 12 }}>📉</div>
              <p style={{ margin: 0, fontSize: 14 }}>
                {selectedCats.length === 0 ? "카테고리를 선택해주세요" : "해당 기간에 데이터가 없습니다"}
              </p>
            </div>
          ) : (
            <ResponsiveContainer width="100%" height={380}>
              <LineChart data={chartData} margin={{ top: 10, right: 20, left: 10, bottom: 60 }}>
                <defs>
                  {selectedCats.map((cat) => {
                    const color = PALETTE[allCategories.indexOf(cat) % PALETTE.length];
                    return (
                      <linearGradient key={cat} id={`grad-${cat}`} x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={color} stopOpacity={0.3} />
                        <stop offset="100%" stopColor={color} stopOpacity={0} />
                      </linearGradient>
                    );
                  })}
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#0f2035" vertical={false} />
                <XAxis
                  dataKey="label"
                  stroke="#0f2035"
                  tick={{ fill: "#334155", fontSize: 11 }}
                  tickLine={false}
                  angle={-50}
                  textAnchor="end"
                  height={70}
                  interval={Math.max(0, Math.floor(chartData.length / 24))}
                />
                <YAxis
                  stroke="#0f2035"
                  tick={{ fill: "#334155", fontSize: 11 }}
                  tickLine={false}
                  axisLine={false}
                  tickFormatter={fmtAmt}
                  domain={yMax ? [0, yMax] : [0, "auto"]}
                  allowDataOverflow={!!yMax}
                  width={52}
                />
                <Tooltip content={<CustomTooltip />} />
                <Legend
                  wrapperStyle={{ paddingTop: 12, fontSize: 12, color: "#64748b" }}
                />
                {selectedCats.map((cat) => {
                  const color = PALETTE[allCategories.indexOf(cat) % PALETTE.length];
                  return (
                    <Line
                      key={cat}
                      type="monotone"
                      dataKey={cat}
                      stroke={color}
                      strokeWidth={2}
                      dot={{ r: 2.5, fill: color, strokeWidth: 0 }}
                      activeDot={{
                        r: 7, fill: color, stroke: "#060d18", strokeWidth: 2,
                        cursor: "pointer",
                        onClick: (e, payload) => {
                          const ym = payload?.payload?.month;
                          if (!ym) return;
                          setSelectedPoint({ ym, cat });
                          setTimeout(() => detailRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
                        }
                      }}
                      connectNulls
                    />
                  );
                })}
              </LineChart>
            </ResponsiveContainer>
          )}
        </div>

        {/* ── 요약 통계 ── */}
        {chartData.length > 0 && selectedCats.length > 0 && (
          <div style={{ display: "grid", gridTemplateColumns: `repeat(${Math.min(selectedCats.length, 4)}, 1fr)`, gap: 10 }}>
            {selectedCats.slice(0, 8).map((cat) => {
              const color = PALETTE[allCategories.indexOf(cat) % PALETTE.length];
              const vals  = chartData.map(d => d[cat] || 0).filter(v => v > 0);
              const total = vals.reduce((s, v) => s + v, 0);
              const avg   = vals.length ? Math.round(total / vals.length) : 0;
              const max   = vals.length ? Math.max(...vals) : 0;
              return (
                <div key={cat} style={{
                  background: "#0a1628", border: `1px solid ${color}33`,
                  borderRadius: 12, padding: "14px 16px",
                  boxShadow: `0 0 20px ${color}11`
                }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: color, display: "inline-block" }} />
                    <span style={{ color, fontSize: 12, fontWeight: 700 }}>{cat}</span>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ color: "#334155", fontSize: 11 }}>합계</span>
                      <span style={{ color: "#94a3b8", fontSize: 12, fontWeight: 600 }}>{total.toLocaleString()}원</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ color: "#334155", fontSize: 11 }}>월평균</span>
                      <span style={{ color: "#94a3b8", fontSize: 12 }}>{avg.toLocaleString()}원</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ color: "#334155", fontSize: 11 }}>최대</span>
                      <span style={{ color: "#94a3b8", fontSize: 12 }}>{max.toLocaleString()}원</span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
        {/* ── 일별 상세 테이블 ── */}
        {selectedPoint && (() => {
          const MAX_ROWS = 30;
          const { ym, cat } = selectedPoint;
          const allRows = transactions
            .filter(t => t.ym === ym && t.cat === cat)
            .sort((a, b) => a.dateStr.localeCompare(b.dateStr));
          const rows = allRows.slice(0, MAX_ROWS);
          const hasMore = allRows.length > MAX_ROWS;
          const total = allRows.reduce((s, t) => s + t.amt, 0);
          const color = PALETTE[allCategories.indexOf(cat) % PALETTE.length];
          const [y, m] = ym.split("-");
          return (
            <div ref={detailRef} style={{
              background: "#0a1628", border: `1px solid ${color}44`,
              borderRadius: 14, padding: "20px 24px",
              boxShadow: `0 0 32px ${color}11`
            }}>
              {/* 테이블 헤더 */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ width: 10, height: 10, borderRadius: "50%", background: color, display: "inline-block" }} />
                  <h2 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: "#f1f5f9" }}>
                    {y}년 {m}월 · <span style={{ color }}>{cat}</span> 일별 내역
                  </h2>
                  <span style={{
                    fontSize: 11, padding: "2px 8px", borderRadius: 10,
                    background: `${color}22`, color, border: `1px solid ${color}44`
                  }}>{allRows.length}건</span>
                  {hasMore && (
                    <span style={{ fontSize: 11, color: "#f59e0b", background: "#f59e0b18", border: "1px solid #f59e0b44", borderRadius: 10, padding: "2px 8px" }}>
                      상위 {MAX_ROWS}건 표시
                    </span>
                  )}
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 13, color: "#94a3b8", fontWeight: 600 }}>
                    합계: <span style={{ color }}>{total.toLocaleString()}원</span>
                  </span>
                  <button
                    onClick={() => setSelectedPoint(null)}
                    style={{
                      padding: "5px 12px", borderRadius: 8, cursor: "pointer",
                      background: "#0f2035", color: "#64748b",
                      border: "1px solid #1e3a5f", fontSize: 12
                    }}
                  >✕ 닫기</button>
                </div>
              </div>

              {rows.length === 0 ? (
                <p style={{ color: "#334155", textAlign: "center", padding: "24px 0" }}>데이터가 없습니다</p>
              ) : (
                <div style={{ overflowX: "auto", overflowY: hasMore ? "auto" : "visible", maxHeight: hasMore ? 520 : "none" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ borderBottom: "1px solid #1e3a5f", position: "sticky", top: 0, background: "#0a1628", zIndex: 1 }}>
                        {[
                          { label: "날짜",   align: "left"  },
                          { label: "자산",   align: "left"  },
                          { label: "카테고리", align: "left" },
                          { label: "소분류", align: "left"  },
                          { label: "내용",   align: "left"  },
                          { label: "금액",   align: "right" },
                          { label: "구분",   align: "left"  },
                        ].map(({ label, align }) => (
                          <th key={label} style={{
                            padding: "8px 12px", textAlign: align,
                            color: "#475569", fontWeight: 700, fontSize: 11,
                            letterSpacing: "0.06em", textTransform: "uppercase",
                            whiteSpace: "nowrap"
                          }}>{label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((t, i) => (
                        <tr key={i} style={{
                          borderBottom: "1px solid #0f2035",
                          background: i % 2 === 0 ? "transparent" : "#060d1844",
                          transition: "background 0.1s"
                        }}
                          onMouseEnter={e => e.currentTarget.style.background = `${color}0d`}
                          onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "transparent" : "#060d1844"}
                        >
                          <td style={{ padding: "9px 12px", color: "#94a3b8", whiteSpace: "nowrap" }}>{t.dateStr}</td>
                          <td style={{ padding: "9px 12px", color: "#64748b", whiteSpace: "nowrap" }}>{t.asset || "—"}</td>
                          <td style={{ padding: "9px 12px" }}>
                            <span style={{
                              display: "inline-block", padding: "2px 9px",
                              borderRadius: 10, fontSize: 11, fontWeight: 600,
                              background: `${color}18`, color, border: `1px solid ${color}33`
                            }}>{t.cat}</span>
                          </td>
                          <td style={{ padding: "9px 12px", color: "#64748b" }}>{t.subCat || "—"}</td>
                          <td style={{ padding: "9px 12px", color: "#94a3b8", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
                            title={t.memo}
                          >{t.memo || "—"}</td>
                          <td style={{ padding: "9px 12px", textAlign: "right", fontWeight: 700, color: "#e2e8f0", whiteSpace: "nowrap" }}>
                            {t.amt.toLocaleString()}원
                          </td>
                          <td style={{ padding: "9px 12px" }}>
                            <span style={{
                              display: "inline-block", padding: "2px 8px",
                              borderRadius: 6, fontSize: 11, fontWeight: 600,
                              background: t.type === "수입" ? "#06402022" : "#3f141422",
                              color: t.type === "수입" ? "#34d399" : "#f87171",
                              border: `1px solid ${t.type === "수입" ? "#34d39944" : "#f8717144"}`
                            }}>{t.type || "—"}</span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr style={{ borderTop: "2px solid #1e3a5f" }}>
                        <td colSpan={5} style={{ padding: "10px 12px", color: "#475569", fontSize: 12, fontWeight: 700 }}>
                          총 {allRows.length}건{hasMore ? ` (상위 ${MAX_ROWS}건 표시)` : ""}
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 800, color, fontSize: 14 }}>
                          {total.toLocaleString()}원
                        </td>
                        <td />
                      </tr>
                    </tfoot>
                  </table>
                </div>
              )}
            </div>
          );
        })()}

      </div>
    </div>
  );
}
