import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Constants ──
const STORAGE_KEY = "mart-helper-v2";
const SOJU_ML = 360;
const SOJU_ABV = 17;
const SOJU_ALCOHOL_ML = (SOJU_ML * SOJU_ABV) / 100;

const DRINK_PRESETS = [
  { name: "와인 (레드/화이트)", volumeMl: 750, abv: 13 },
  { name: "맥주 500ml 캔", volumeMl: 500, abv: 5 },
  { name: "맥주 355ml 캔", volumeMl: 355, abv: 5 },
  { name: "막걸리", volumeMl: 750, abv: 6 },
  { name: "하이볼 캔", volumeMl: 350, abv: 7 },
  { name: "위스키", volumeMl: 700, abv: 40 },
  { name: "사케", volumeMl: 720, abv: 15 },
  { name: "직접 입력", volumeMl: 0, abv: 0 },
];

const CART_TABS = [
  { key: "mart", label: "🛒 마트", color: "#16a34a" },
  { key: "online", label: "📦 온라인", color: "#2563eb" },
  { key: "donated", label: "🎁 기증", color: "#d97706" },
];

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const fmt = (n) => Number(n).toLocaleString("ko-KR");

const initial = () => ({
  cart: { mart: [], online: [], donated: [] },
  alc: { people: 5, sojuPer: 2, drinks: [] },
});

// ── localStorage helpers ──
function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {
    console.warn("Failed to load saved data:", e);
  }
  return null;
}

function saveToStorage(data) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (e) {
    console.warn("Failed to save data:", e);
  }
}

// ── Shared Styles ──
const s = {
  page: {
    fontFamily: "'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif",
    background: "#f1f5f0", minHeight: "100vh", display: "flex", flexDirection: "column",
    maxWidth: 520, margin: "0 auto", color: "#1a1a1a",
  },
  header: { background: "linear-gradient(135deg, #1b5e37 0%, #2d8a56 100%)", padding: "20px 20px 14px", color: "#fff" },
  headerTitle: { fontSize: 22, fontWeight: 800, letterSpacing: -0.5 },
  headerSub: { fontSize: 12, opacity: 0.7, marginTop: 2 },
  mainTabs: { display: "flex", gap: 0, background: "#fff", borderBottom: "2px solid #e5e7eb", position: "sticky", top: 0, zIndex: 20 },
  mainTab: (active) => ({
    flex: 1, padding: "13px 0", textAlign: "center", fontSize: 15,
    fontWeight: active ? 700 : 500, color: active ? "#1b5e37" : "#888",
    background: active ? "#f0faf4" : "#fff",
    borderBottom: active ? "3px solid #1b5e37" : "3px solid transparent",
    cursor: "pointer", transition: "all 0.15s", userSelect: "none",
  }),
  body: { flex: 1, padding: "12px 14px 100px", overflowY: "auto" },
  card: { background: "#fff", borderRadius: 14, padding: "14px 16px", marginBottom: 10, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" },
  input: (w) => ({
    width: w, padding: "8px 10px", borderRadius: 8, border: "1.5px solid #ddd",
    fontSize: 14, outline: "none", background: "#fafafa", transition: "border 0.15s",
  }),
  btn: (bg, color = "#fff") => ({
    background: bg, color, border: "none", borderRadius: 10, padding: "10px 16px",
    fontSize: 14, fontWeight: 600, cursor: "pointer", display: "flex",
    alignItems: "center", justifyContent: "center", gap: 6, transition: "opacity 0.15s", userSelect: "none",
  }),
  smallBtn: (bg, color = "#fff") => ({
    background: bg, color, border: "none", borderRadius: 8,
    padding: "6px 10px", fontSize: 12, fontWeight: 600, cursor: "pointer", lineHeight: 1,
  }),
  delBtn: { background: "none", border: "none", color: "#ccc", fontSize: 20, cursor: "pointer", padding: "2px 6px", lineHeight: 1 },
  label: { fontSize: 13, color: "#666", fontWeight: 500 },
  tag: (bg, color) => ({ display: "inline-block", padding: "3px 8px", borderRadius: 6, fontSize: 11, fontWeight: 600, background: bg, color }),
  bottomBar: {
    position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)",
    width: "100%", maxWidth: 520, background: "#fff", borderTop: "1px solid #e0e0e0",
    padding: "12px 20px", boxShadow: "0 -2px 12px rgba(0,0,0,0.06)", zIndex: 10,
  },
};

// ── Cart Item Row ──
function CartItem({ item, onUpdate, onDelete, showPrice }) {
  return (
    <div style={{ ...s.card, padding: "10px 12px" }}>
      <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
        <input
          style={{ ...s.input("auto"), flex: 1, minWidth: 80 }}
          placeholder="품목명"
          value={item.name}
          onChange={(e) => onUpdate(item.id, "name", e.target.value)}
        />
        <button style={s.delBtn} onClick={() => onDelete(item.id)} title="삭제">×</button>
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "center" }}>
        {showPrice && (
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>단가 (₩)</div>
            <input
              style={{ ...s.input("100%"), textAlign: "right" }}
              type="number" inputMode="numeric" placeholder="0"
              value={item.price || ""}
              onChange={(e) => onUpdate(item.id, "price", Number(e.target.value) || 0)}
            />
          </div>
        )}
        <div style={{ width: 70 }}>
          <div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>수량</div>
          <div style={{ display: "flex", alignItems: "center" }}>
            <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "8px 0 0 8px", padding: "8px 10px" }}
              onClick={() => onUpdate(item.id, "quantity", Math.max(1, item.quantity - 1))}>−</button>
            <input
              style={{ width: 36, textAlign: "center", padding: "7px 0", border: "1.5px solid #ddd", borderLeft: "none", borderRight: "none", fontSize: 14, outline: "none", background: "#fafafa" }}
              type="number" inputMode="numeric" value={item.quantity}
              onChange={(e) => onUpdate(item.id, "quantity", Math.max(1, Number(e.target.value) || 1))}
            />
            <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "0 8px 8px 0", padding: "8px 10px" }}
              onClick={() => onUpdate(item.id, "quantity", item.quantity + 1)}>+</button>
          </div>
        </div>
        {showPrice && (
          <div style={{ minWidth: 80, textAlign: "right", paddingTop: 16 }}>
            <span style={{ fontSize: 15, fontWeight: 700, color: "#1b5e37" }}>₩{fmt(item.price * item.quantity)}</span>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Cart View ──
function CartView({ data, save }) {
  const [tab, setTab] = useState("mart");
  const items = data.cart[tab] || [];
  const showPrice = tab !== "donated";

  const addItem = () => {
    const newItem = { id: uid(), name: "", price: 0, quantity: 1 };
    save({ ...data, cart: { ...data.cart, [tab]: [...items, newItem] } });
  };
  const updateItem = (id, field, value) => {
    save({ ...data, cart: { ...data.cart, [tab]: items.map((i) => (i.id === id ? { ...i, [field]: value } : i)) } });
  };
  const deleteItem = (id) => {
    save({ ...data, cart: { ...data.cart, [tab]: items.filter((i) => i.id !== id) } });
  };

  const catTotal = items.reduce((a, i) => a + (showPrice ? i.price * i.quantity : 0), 0);
  const catQty = items.reduce((a, i) => a + i.quantity, 0);
  const allTotal = Object.entries(data.cart).reduce(
    (a, [k, v]) => a + (k !== "donated" ? v.reduce((sum, i) => sum + i.price * i.quantity, 0) : 0), 0
  );
  const allQty = Object.values(data.cart).reduce((a, v) => a + v.reduce((sum, i) => sum + i.quantity, 0), 0);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    let hasData = false;
    CART_TABS.forEach(({ key, label }) => {
      const list = data.cart[key];
      if (list.length === 0) return;
      hasData = true;
      const isD = key === "donated";
      const rows = list.map((i) => ({
        품목: i.name || "(미입력)",
        ...(isD ? {} : { "단가(₩)": i.price }),
        수량: i.quantity,
        ...(isD ? {} : { "소계(₩)": i.price * i.quantity }),
      }));
      if (!isD) {
        rows.push({ 품목: "【합계】", "단가(₩)": "", 수량: list.reduce((a, i) => a + i.quantity, 0), "소계(₩)": list.reduce((a, i) => a + i.price * i.quantity, 0) });
      } else {
        rows.push({ 품목: "【합계】", 수량: list.reduce((a, i) => a + i.quantity, 0) });
      }
      const ws = XLSX.utils.json_to_sheet(rows);
      ws["!cols"] = Object.keys(rows[0]).map((k) => ({ wch: Math.max(k.length * 2, 12) }));
      XLSX.utils.book_append_sheet(wb, ws, label.replace(/[^\w가-힣]/g, "").trim());
    });
    if (!hasData) { alert("내보낼 데이터가 없습니다."); return; }
    XLSX.writeFile(wb, `장보기_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <>
      {/* Sub-tabs */}
      <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
        {CART_TABS.map((t) => {
          const active = tab === t.key;
          const count = (data.cart[t.key] || []).length;
          return (
            <button key={t.key} onClick={() => setTab(t.key)} style={{
              flex: 1, padding: "9px 4px", borderRadius: 10,
              border: active ? `2px solid ${t.color}` : "2px solid #e5e7eb",
              background: active ? `${t.color}11` : "#fff",
              color: active ? t.color : "#888",
              fontSize: 13, fontWeight: active ? 700 : 500, cursor: "pointer", transition: "all 0.15s",
            }}>
              {t.label} {count > 0 && <span style={s.tag(active ? t.color : "#e5e7eb", active ? "#fff" : "#888")}>{count}</span>}
            </button>
          );
        })}
      </div>

      {/* Items */}
      {items.length === 0 ? (
        <div style={{ textAlign: "center", padding: "40px 20px", color: "#bbb" }}>
          <div style={{ fontSize: 40, marginBottom: 8 }}>{tab === "mart" ? "🛒" : tab === "online" ? "📦" : "🎁"}</div>
          <div style={{ fontSize: 14 }}>아직 추가된 물품이 없어요</div>
          <div style={{ fontSize: 12, marginTop: 4 }}>아래 버튼으로 추가해보세요</div>
        </div>
      ) : (
        items.map((item) => (
          <CartItem key={item.id} item={item} onUpdate={updateItem} onDelete={deleteItem} showPrice={showPrice} />
        ))
      )}

      {/* Add & Export */}
      <button onClick={addItem} style={{ ...s.btn("#1b5e37"), width: "100%", marginTop: 4, padding: "12px 0", borderRadius: 12, fontSize: 15 }}>
        + 물품 추가
      </button>
      <button onClick={exportExcel} style={{ ...s.btn("#f5f5f5", "#555"), width: "100%", marginTop: 8, padding: "10px 0", borderRadius: 12, fontSize: 13 }}>
        📊 Excel로 내보내기
      </button>

      {/* Bottom bar */}
      <div style={s.bottomBar}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 11, color: "#999" }}>현재 탭 ({CART_TABS.find((t) => t.key === tab)?.label})</div>
            <div style={{ fontSize: 14, fontWeight: 600 }}>
              {catQty}건{showPrice && <> · <span style={{ color: "#1b5e37" }}>₩{fmt(catTotal)}</span></>}
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 11, color: "#999" }}>전체 합계</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#1b5e37" }}>
              ₩{fmt(allTotal)}
              <span style={{ fontSize: 12, fontWeight: 500, color: "#999", marginLeft: 6 }}>{allQty}건</span>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

// ── Result Box ──
function ResultBox({ label, value, sub, highlight }) {
  return (
    <div style={{ padding: "10px 12px", background: highlight ? "#1b5e37" : "#fff", borderRadius: 10, border: highlight ? "none" : "1px solid #eee" }}>
      <div style={{ fontSize: 11, color: highlight ? "rgba(255,255,255,0.7)" : "#999", marginBottom: 2 }}>{label}</div>
      <div style={{ fontSize: 18, fontWeight: 800, color: highlight ? "#fff" : "#1b5e37" }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: highlight ? "rgba(255,255,255,0.6)" : "#bbb", marginTop: 1 }}>{sub}</div>}
    </div>
  );
}

// ── Alcohol Calculator ──
function AlcoholCalc({ data, save }) {
  const { people, sojuPer, drinks } = data.alc;
  const totalSojuNeeded = people * sojuPer;

  const drinkSojuEquiv = (d) => {
    if (!d.volumeMl || !d.abv) return 0;
    return ((d.volumeMl * d.abv) / 100 / SOJU_ALCOHOL_ML) * (d.quantity || 1);
  };
  const totalEquiv = drinks.reduce((a, d) => a + drinkSojuEquiv(d), 0);
  const remaining = Math.max(0, totalSojuNeeded - totalEquiv);

  const setAlc = (patch) => save({ ...data, alc: { ...data.alc, ...patch } });

  const addDrink = (preset) => {
    const d = { id: uid(), name: preset.name === "직접 입력" ? "" : preset.name, volumeMl: preset.volumeMl, abv: preset.abv, quantity: 1 };
    setAlc({ drinks: [...drinks, d] });
  };
  const updateDrink = (id, field, value) => {
    setAlc({ drinks: drinks.map((d) => (d.id === id ? { ...d, [field]: value } : d)) });
  };
  const deleteDrink = (id) => {
    setAlc({ drinks: drinks.filter((d) => d.id !== id) });
  };
  const [showPresets, setShowPresets] = useState(false);

  return (
    <>
      {/* Info */}
      <div style={{ ...s.card, background: "linear-gradient(135deg, #fef9ee 0%, #fff7e6 100%)", border: "1.5px solid #f0d98d" }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: "#92702a", marginBottom: 6 }}>🍶 기준: 소주 1병</div>
        <div style={{ fontSize: 12, color: "#a08339" }}>
          {SOJU_ML}ml · {SOJU_ABV}% · 순수 알코올 {SOJU_ALCOHOL_ML.toFixed(1)}ml
        </div>
      </div>

      {/* People & consumption */}
      <div style={s.card}>
        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>👥 모임 설정</div>
        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <div style={s.label}>참여 인원</div>
            <div style={{ display: "flex", alignItems: "center", marginTop: 4 }}>
              <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "8px 0 0 8px", padding: "10px 14px", fontSize: 16 }} onClick={() => setAlc({ people: Math.max(1, people - 1) })}>−</button>
              <input style={{ width: 48, textAlign: "center", padding: "9px 0", border: "1.5px solid #ddd", borderLeft: "none", borderRight: "none", fontSize: 16, fontWeight: 700, outline: "none", background: "#fafafa" }}
                type="number" inputMode="numeric" value={people}
                onChange={(e) => setAlc({ people: Math.max(1, Number(e.target.value) || 1) })} />
              <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "0 8px 8px 0", padding: "10px 14px", fontSize: 16 }} onClick={() => setAlc({ people: people + 1 })}>+</button>
            </div>
          </div>
          <div style={{ flex: 1 }}>
            <div style={s.label}>1인당 소주 (병)</div>
            <div style={{ display: "flex", alignItems: "center", marginTop: 4 }}>
              <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "8px 0 0 8px", padding: "10px 14px", fontSize: 16 }} onClick={() => setAlc({ sojuPer: Math.max(0.5, sojuPer - 0.5) })}>−</button>
              <input style={{ width: 48, textAlign: "center", padding: "9px 0", border: "1.5px solid #ddd", borderLeft: "none", borderRight: "none", fontSize: 16, fontWeight: 700, outline: "none", background: "#fafafa" }}
                type="number" inputMode="decimal" step="0.5" value={sojuPer}
                onChange={(e) => setAlc({ sojuPer: Math.max(0, Number(e.target.value) || 0) })} />
              <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "0 8px 8px 0", padding: "10px 14px", fontSize: 16 }} onClick={() => setAlc({ sojuPer: sojuPer + 0.5 })}>+</button>
            </div>
          </div>
        </div>
        <div style={{ marginTop: 12, padding: "10px 14px", background: "#f0faf4", borderRadius: 10, textAlign: "center" }}>
          <span style={{ fontSize: 13, color: "#666" }}>총 필요 소주</span>
          <span style={{ fontSize: 22, fontWeight: 800, color: "#1b5e37", marginLeft: 10 }}>{totalSojuNeeded}병</span>
        </div>
      </div>

      {/* Other drinks */}
      <div style={s.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ fontSize: 14, fontWeight: 700 }}>🍷 다른 주류 추가</div>
          <button onClick={() => setShowPresets(!showPresets)} style={{ ...s.smallBtn("#1b5e37"), borderRadius: 8, padding: "6px 12px" }}>
            + 추가
          </button>
        </div>

        {showPresets && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12, padding: 10, background: "#f9f9f9", borderRadius: 10 }}>
            {DRINK_PRESETS.map((p, i) => (
              <button key={i} onClick={() => { addDrink(p); setShowPresets(false); }}
                style={{ ...s.smallBtn("#fff", "#333"), border: "1.5px solid #ddd", borderRadius: 8, padding: "8px 12px", fontSize: 13 }}>
                {p.name}
              </button>
            ))}
          </div>
        )}

        {drinks.length === 0 ? (
          <div style={{ textAlign: "center", padding: "20px 0", color: "#ccc", fontSize: 13 }}>
            추가된 주류가 없습니다
          </div>
        ) : (
          drinks.map((d) => {
            const equiv = drinkSojuEquiv(d);
            return (
              <div key={d.id} style={{ padding: "10px 0", borderBottom: "1px solid #f0f0f0" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                  <input style={{ ...s.input("auto"), flex: 1, fontSize: 14, fontWeight: 600 }} placeholder="주류명"
                    value={d.name} onChange={(e) => updateDrink(d.id, "name", e.target.value)} />
                  <button style={s.delBtn} onClick={() => deleteDrink(d.id)}>×</button>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "flex-end" }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>용량(ml)</div>
                    <input style={{ ...s.input("100%"), textAlign: "right" }} type="number" inputMode="numeric"
                      value={d.volumeMl || ""} onChange={(e) => updateDrink(d.id, "volumeMl", Number(e.target.value) || 0)} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>도수(%)</div>
                    <input style={{ ...s.input("100%"), textAlign: "right" }} type="number" inputMode="decimal" step="0.1"
                      value={d.abv || ""} onChange={(e) => updateDrink(d.id, "abv", Number(e.target.value) || 0)} />
                  </div>
                  <div style={{ width: 70 }}>
                    <div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>수량</div>
                    <div style={{ display: "flex" }}>
                      <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "8px 0 0 8px", padding: "8px 8px" }}
                        onClick={() => updateDrink(d.id, "quantity", Math.max(1, d.quantity - 1))}>−</button>
                      <input style={{ width: 30, textAlign: "center", padding: "7px 0", border: "1.5px solid #ddd", borderLeft: "none", borderRight: "none", fontSize: 13, outline: "none", background: "#fafafa" }}
                        type="number" inputMode="numeric" value={d.quantity}
                        onChange={(e) => updateDrink(d.id, "quantity", Math.max(1, Number(e.target.value) || 1))} />
                      <button style={{ ...s.smallBtn("#eee", "#333"), borderRadius: "0 8px 8px 0", padding: "8px 8px" }}
                        onClick={() => updateDrink(d.id, "quantity", d.quantity + 1)}>+</button>
                    </div>
                  </div>
                </div>
                <div style={{ marginTop: 6, fontSize: 12, color: "#1b5e37", fontWeight: 600 }}>
                  ≈ 소주 {equiv.toFixed(1)}병 분량{" "}
                  <span style={{ color: "#999", fontWeight: 400 }}>
                    ({((d.volumeMl * d.abv / 100) * d.quantity).toFixed(1)}ml 순수알코올)
                  </span>
                </div>
              </div>
            );
          })
        )}
      </div>

      {/* Result */}
      <div style={{
        ...s.card,
        background: totalEquiv > 0 ? "linear-gradient(135deg, #f0faf4 0%, #e6f7ed 100%)" : "#fff",
        border: totalEquiv > 0 ? "1.5px solid #86d4a0" : "1.5px solid #eee",
      }}>
        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 10 }}>📊 계산 결과</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <ResultBox label="총 필요 소주" value={`${totalSojuNeeded}병`} />
          <ResultBox label="다른 주류 합산" value={`≈ ${totalEquiv.toFixed(1)}병`} sub="소주 환산" />
          <ResultBox label="차감 후 소주" value={`${remaining.toFixed(1)}병`} highlight />
          <ResultBox label="1인당 소주" value={`${(remaining / people).toFixed(1)}병`} sub={`÷ ${people}명`} />
        </div>
        {totalEquiv > 0 && totalEquiv >= totalSojuNeeded && (
          <div style={{ marginTop: 10, padding: "8px 12px", background: "#fff3cd", borderRadius: 8, fontSize: 12, color: "#856404", textAlign: "center" }}>
            ⚠️ 다른 주류만으로 목표 알코올량을 초과합니다. 소주를 추가로 살 필요가 없어요!
          </div>
        )}
      </div>
      <div style={{ height: 20 }} />
    </>
  );
}

// ── Main App ──
export default function App() {
  const [data, setData] = useState(() => loadData() || initial());
  const [mainTab, setMainTab] = useState("cart");
  const [saved, setSaved] = useState(false);
  const saveTimer = useRef(null);

  const save = useCallback((newData) => {
    setData(newData);
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      saveToStorage(newData);
      setSaved(true);
      setTimeout(() => setSaved(false), 1200);
    }, 300);
  }, []);

  const resetAll = () => {
    if (window.confirm("모든 데이터를 초기화하시겠습니까?")) {
      const fresh = initial();
      setData(fresh);
      saveToStorage(fresh);
    }
  };

  return (
    <div style={s.page}>
      {/* Header */}
      <div style={s.header}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={s.headerTitle}>장보기 도우미</div>
            <div style={s.headerSub}>장바구니 · 주류 계산기</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {saved && <span style={{ fontSize: 11, opacity: 0.8, color: "#a5f3c0" }}>✓ 저장됨</span>}
            <button onClick={resetAll} style={{
              background: "rgba(255,255,255,0.15)", border: "none", color: "#fff",
              borderRadius: 8, padding: "6px 10px", fontSize: 12, cursor: "pointer",
            }}>초기화</button>
          </div>
        </div>
      </div>

      {/* Main tabs */}
      <div style={s.mainTabs}>
        <div style={s.mainTab(mainTab === "cart")} onClick={() => setMainTab("cart")}>🛒 장바구니</div>
        <div style={s.mainTab(mainTab === "alcohol")} onClick={() => setMainTab("alcohol")}>🍶 주류 계산기</div>
      </div>

      {/* Body */}
      <div style={s.body}>
        {mainTab === "cart" ? <CartView data={data} save={save} /> : <AlcoholCalc data={data} save={save} />}
      </div>
    </div>
  );
}
