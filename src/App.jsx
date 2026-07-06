import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════
// CONSTANTS & HELPERS
// ═══════════════════════════════════════
const AUTH_KEY = "mt-auth-v4";
const SOJU_ML = 360, SOJU_ABV = 17, SOJU_ALC = SOJU_ML * SOJU_ABV / 100;
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const fmt = (n) => Number(n).toLocaleString("ko-KR");
const shuffle = (arr) => { const a = [...arr]; for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const timeStr = (t) => new Date(t).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
const dateStr = (t) => new Date(t).toLocaleString("ko-KR", { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
const NOTICE_SEEN_KEY = "mt-notices-seen";

// 이미지를 캔버스로 리사이즈 후 JPEG dataURL 반환 (최대 800px, ~수십-수백KB)
const resizeImage = (file, maxDim = 800) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
      const draw = (dim, q) => {
        const sc = Math.min(1, dim / Math.max(img.width, img.height));
        const c = document.createElement("canvas");
        c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        return c.toDataURL("image/jpeg", q);
      };
      let out = draw(maxDim, 0.72);
      if (out.length > 500000) out = draw(maxDim, 0.5);
      if (out.length > 500000) out = draw(600, 0.5);
      if (out.length > 700000) reject(new Error("이미지 용량을 줄일 수 없습니다. 더 작은 이미지를 사용해주세요."));
      else resolve(out);
    };
    img.onerror = () => reject(new Error("이미지를 읽을 수 없습니다."));
    img.src = e.target.result;
  };
  reader.onerror = () => reject(new Error("파일 읽기 실패"));
  reader.readAsDataURL(file);
});

const DRINK_PRESETS = [
  { name: "와인", v: 750, a: 13 }, { name: "맥주 500ml", v: 500, a: 5 },
  { name: "맥주 355ml", v: 355, a: 5 }, { name: "막걸리", v: 750, a: 6 },
  { name: "하이볼 캔", v: 350, a: 7 }, { name: "위스키", v: 700, a: 40 },
  { name: "사케", v: 720, a: 15 }, { name: "직접 입력", v: 0, a: 0 },
];
const CART_TABS = [
  { key: "mart", label: "🛒 마트", c: "#16a34a" },
  { key: "online", label: "📦 온라인", c: "#2563eb" },
  { key: "donated", label: "🎁 기증", c: "#d97706" },
  { key: "pre", label: "✅ 사전구매", c: "#7c3aed" },
];
const SHEET_TO_TAB = { "마트": "mart", "온라인": "online", "기증": "donated", "사전구매": "pre" };
const TAB_TO_SHEET = { mart: "마트", online: "온라인", donated: "기증", pre: "사전구매" };

const CATEGORIES = [
  { key: "games", icon: "🎮", label: "게임 & 진행", desc: "룰렛 · 팀배정 · 점수판 · 미니게임", color: "#7c3aed", bg: "linear-gradient(135deg,#ede9fe,#ddd6fe)" },
  { key: "manage", icon: "📋", label: "정산 & 관리", desc: "장바구니 · 주류계산 · N빵정산 · 일정표", color: "#0d9488", bg: "linear-gradient(135deg,#ccfbf1,#99f6e4)" },
  { key: "fun", icon: "🎉", label: "분위기 & 기타", desc: "공지 · 아이스브레이킹 · 투표", color: "#e11d48", bg: "linear-gradient(135deg,#ffe4e6,#fecdd3)" },
  { key: "admin", icon: "⚙️", label: "관리", desc: "참가자 계정 · 시스템", color: "#475569", bg: "linear-gradient(135deg,#f1f5f9,#e2e8f0)", staffOnly: true },
];
const SUBS = {
  games: [
    { key: "roulette", label: "🎯 룰렛" }, { key: "teams", label: "👥 팀배정" },
    { key: "score", label: "🏆 점수판" }, { key: "mini", label: "🎲 미니게임" },
  ],
  manage: [
    { key: "cart", label: "🛒 장바구니" }, { key: "alc", label: "🍶 주류" },
    { key: "settle", label: "💸 N빵" }, { key: "sched", label: "📅 일정" },
  ],
  fun: [{ key: "notice", label: "📢 공지" }, { key: "ice", label: "❓ 아이스브레이킹" }, { key: "vote", label: "🗳️ 투표" }],
  admin: [{ key: "system", label: "🖥️ 시스템", adminOnly: true }, { key: "members", label: "👤 참가자 계정" }],
};

const CHOSUNG_SET = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ".split("");
const ICE_QUESTIONS = [
  "무인도에 딱 하나만 가져갈 수 있다면?","지금까지 가장 부끄러웠던 순간은?",
  "로또 1등 당첨되면 제일 먼저 할 일은?","다시 태어나면 되고 싶은 직업은?",
  "최근 가장 많이 웃었던 순간은?","인생 최고의 여행지는?",
  "핸드폰에서 가장 많이 쓰는 앱은?","어렸을 때 장래희망은?",
  "가장 좋아하는 음식과 그 이유는?","초능력 하나를 갖는다면? (텔레포트/독심술/투명인간/시간정지)",
  "인생에서 가장 잘한 결정은?","100만원이 생기면 당장 사고 싶은 것은?",
  "지금 바로 여행 갈 수 있다면 어디로?","나만의 스트레스 해소법은?",
  "죽기 전에 꼭 해보고 싶은 것은?","나를 동물로 표현한다면?",
  "이상형을 3가지 키워드로 표현하면?","요즘 빠져있는 취미는?",
  "가장 감명 깊게 본 영화/드라마는?","10년 후 나는 어디서 뭘 하고 있을까?",
  "마지막 식사로 먹고 싶은 메뉴는?","내 인생 TMI 하나를 공개한다면?",
  "지금 카톡 프사의 의미는?","최근 가장 많이 들은 노래는?",
  "가장 오래된 친구는 몇 년째?","나만 아는 맛집을 하나 공개한다면?",
  "올해 가장 잘 산 물건은?","가장 좋아하는 계절과 이유는?",
  "처음 술 마셨을 때 에피소드는?","지금 가장 갖고 싶은 것은?",
];

// ═══════════════════════════════════════
// DARK MODE — 반전 필터 방식 (기기별 localStorage 저장)
// 이미지(<img>)는 카운터 필터로 원색 유지
// ═══════════════════════════════════════
const DARK_KEY = "mt-dark";
if (typeof document !== "undefined" && !document.getElementById("mt-dark-style")) {
  const st = document.createElement("style");
  st.id = "mt-dark-style";
  st.textContent = `
    html.mt-dark { filter: invert(0.93) hue-rotate(180deg); background: #0b0f14; }
    html.mt-dark img { filter: invert(1.075) hue-rotate(180deg); }
  `;
  document.head.appendChild(st);
}
function useDark() {
  const [dark, setDark] = useState(() => localStorage.getItem(DARK_KEY) === "1");
  useEffect(() => {
    document.documentElement.classList.toggle("mt-dark", dark);
    localStorage.setItem(DARK_KEY, dark ? "1" : "0");
  }, [dark]);
  return [dark, () => setDark(d => !d)];
}

// ═══════════════════════════════════════
// API CLIENT
// ═══════════════════════════════════════
const api = {
  token: null,
  async req(path, method = "GET", body) {
    const r = await fetch(`/api/${path}`, {
      method,
      headers: { "Content-Type": "application/json", ...(this.token ? { Authorization: `Bearer ${this.token}` } : {}) },
      body: body ? JSON.stringify(body) : undefined,
    });
    let d; try { d = await r.json(); } catch { throw new Error("서버 응답 오류"); }
    if (!r.ok) { const e = new Error(d.message || d.error || "오류가 발생했습니다"); e.code = d.error; throw e; }
    return d;
  },
};

// ═══════════════════════════════════════
// STYLES
// ═══════════════════════════════════════
const s = {
  page: { fontFamily: "'Pretendard',-apple-system,BlinkMacSystemFont,system-ui,sans-serif", background: "#f3f4f6", minHeight: "100vh", display: "flex", flexDirection: "column", maxWidth: 520, margin: "0 auto", color: "#1a1a1a" },
  card: { background: "#fff", borderRadius: 14, padding: "14px 16px", marginBottom: 10, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" },
  input: (w) => ({ width: w, padding: "8px 10px", borderRadius: 8, border: "1.5px solid #ddd", fontSize: 14, outline: "none", background: "#fafafa", boxSizing: "border-box" }),
  btn: (bg, c = "#fff") => ({ background: bg, color: c, border: "none", borderRadius: 10, padding: "10px 16px", fontSize: 14, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, userSelect: "none" }),
  smBtn: (bg, c = "#fff") => ({ background: bg, color: c, border: "none", borderRadius: 8, padding: "6px 10px", fontSize: 12, fontWeight: 600, cursor: "pointer", lineHeight: 1 }),
  del: { background: "none", border: "none", color: "#ccc", fontSize: 20, cursor: "pointer", padding: "2px 6px", lineHeight: 1 },
  label: { fontSize: 13, color: "#666", fontWeight: 500, marginBottom: 4 },
  tag: (bg, c) => ({ display: "inline-block", padding: "3px 8px", borderRadius: 6, fontSize: 11, fontWeight: 600, background: bg, color: c }),
  section: { fontSize: 14, fontWeight: 700, marginBottom: 10 },
  roNote: { fontSize: 12, color: "#94a3b8", textAlign: "center", padding: "6px 0 10px" },
};

const Stepper = ({ value, onChange, min = 1, step = 1, decimal }) => (
  <div style={{ display: "flex", alignItems: "center" }}>
    <button style={{ background: "#eee", color: "#333", border: "none", borderRadius: "8px 0 0 8px", padding: "8px 12px", fontSize: 16, cursor: "pointer" }} onClick={() => onChange(Math.max(min, +(value - step).toFixed(1)))}>−</button>
    <input style={{ width: 44, textAlign: "center", padding: "7px 0", border: "1.5px solid #ddd", borderLeft: "none", borderRight: "none", fontSize: 15, fontWeight: 700, outline: "none", background: "#fafafa" }} type="number" inputMode={decimal ? "decimal" : "numeric"} value={value} onChange={e => onChange(Math.max(min, Number(e.target.value) || min))} />
    <button style={{ background: "#eee", color: "#333", border: "none", borderRadius: "0 8px 8px 0", padding: "8px 12px", fontSize: 16, cursor: "pointer" }} onClick={() => onChange(+(value + step).toFixed(1))}>+</button>
  </div>
);

const RoNote = () => <div style={s.roNote}>👀 보기 전용 — 편집은 호스트만 가능합니다</div>;

// ═══════════════════════════════════════
// HELP CONTENT (역할별 사용 안내)
// ═══════════════════════════════════════
const HELP = {
  guest: {
    icon: "🙋", label: "게스트", color: "#0d9488",
    intro: "게스트는 호스트가 등록해준 이름으로 입장해, 게임에 참여하고 정보를 확인합니다.",
    sections: [
      { t: "입장하기", items: [
        "로그인 화면에서 게스트를 선택하고, 호스트가 등록해준 이름을 입력하면 입장됩니다.",
        "비밀번호는 본인 이름과 동일합니다 (비워두면 자동으로 채워져요).",
        "\"등록되지 않은 참가자\"라고 나오면 호스트에게 이름 등록을 요청하세요. 띄어쓰기·오타까지 정확히 같아야 합니다.",
      ]},
      { t: "무엇을 할 수 있나요", items: [
        "📢 공지 확인 — 홈 상단 배너나 분위기&기타 → 공지에서 호스트 공지를 봅니다. 안 읽은 공지엔 New! 표시가 붙어요.",
        "🎲 미니게임 참여 — 호스트가 문제를 내면(업다운·초성·자유출제) 답을 제출합니다.",
        "🗳️ 투표 — 진행 중인 투표에 참여합니다. 1인 1표이고, 마감 전에는 선택을 바꿀 수 있어요.",
        "👀 결과 보기 — 룰렛 당첨, 팀 배정(내 팀 ⭐), 점수판 순위, N빵 정산(내 금액 ⭐), 일정표를 확인합니다.",
      ]},
      { t: "알아두면 좋아요", items: [
        "화면은 10초마다 자동 갱신됩니다. 즉시 새로고침하려면 상단 🔄 버튼을 누르세요.",
        "편집·진행은 호스트만 할 수 있어요. 게스트 화면의 '👀 보기 전용' 항목은 눈으로 확인하는 용도입니다.",
        "다른 기기·브라우저로 다시 접속해도 이름만 넣으면 내 투표·제출 기록이 그대로 유지됩니다.",
        "🌙 버튼으로 다크 모드를 켤 수 있습니다 (내 기기에만 적용).",
      ]},
    ],
  },
  host: {
    icon: "👑", label: "호스트", color: "#d97706",
    intro: "호스트는 MT 전반을 준비하고 진행합니다. 참가자 등록부터 게임 진행, 정산까지 대부분의 편집 권한을 가집니다.",
    sections: [
      { t: "처음 시작하기", items: [
        "로그인 화면에서 호스트를 선택하고 비밀번호를 정해 입력하세요. 최초 입력한 비밀번호가 그대로 호스트 비번으로 설정됩니다 (4자 이상).",
        "이후에는 그 비밀번호로 로그인합니다. 잊어버리면 Admin이 초기화해줄 수 있어요.",
        "가장 먼저 ⚙️ 관리 → 참가자 계정에서 참가자 이름을 등록하세요. 쉼표로 여러 명을 한 번에 넣을 수 있습니다.",
      ]},
      { t: "게임 & 진행", items: [
        "🎯 룰렛 — 등록된 참가자 중 랜덤 １명을 뽑습니다. 결과가 게스트에게도 공유됩니다.",
        "👥 팀배정 — 팀 수를 정하고 자동으로 섞어 배정합니다.",
        "🏆 점수판 — 팀을 만들고 라운드별 점수를 입력하면 순위가 자동 집계됩니다.",
        "🎲 미니게임 — 업다운·초성은 문제를 내면 게스트가 답을 제출하고, 자유출제는 직접 문제를 쓰고 정답자를 선택합니다. (정답 숫자는 호스트에게만 보여요.)",
      ]},
      { t: "정산 & 관리", items: [
        "🛒 장바구니 — 마트·온라인·기증·사전구매로 나눠 물품을 관리합니다. 엑셀로 내보내기/가져오기가 가능해요.",
        "🍶 주류계산 — 인원과 1인당 소주량으로 필요한 술을 계산합니다.",
        "💸 N빵 — 지출 항목별로 금액과 제외 인원을 정하면 1인당 정산액이 나옵니다.",
        "📅 일정 — 시간과 내용을 입력하면 게스트가 볼 수 있는 타임테이블이 됩니다.",
      ]},
      { t: "분위기 & 공지", items: [
        "📢 공지 — 텍스트와 이미지로 공지를 올립니다. 이미지는 자동으로 축소 저장돼요.",
        "❓ 아이스브레이킹 — 질문을 뽑습니다. 한 번 나온 질문은 다시 안 나오며, 초기화하면 전부 다시 나옵니다.",
        "🗳️ 투표 — 주제와 선택지로 투표를 만들고 마감/재개할 수 있습니다.",
      ]},
      { t: "진행 팁", items: [
        "호스트가 게임에 직접 참여하려면, 참가자 계정에 본인 이름을 등록한 뒤 시크릿 창에서 게스트로도 로그인하세요.",
        "게임 데이터를 정리하려면 ⚙️ 관리 → 참가자 계정 하단의 '게임 데이터 초기화'를 쓰세요. 참가자 계정과 비밀번호는 유지됩니다.",
      ]},
    ],
  },
  admin: {
    icon: "⚙️", label: "Admin", color: "#475569",
    intro: "Admin은 시스템 관리자입니다. 호스트의 모든 권한에 더해 백업·복원·시스템 초기화 등 상위 관리 기능을 담당합니다.",
    sections: [
      { t: "입장하기", items: [
        "로그인 화면에서 Admin을 선택하고 관리자 비밀번호를 입력합니다.",
        "이 비밀번호는 Cloudflare의 ADMIN_PASSWORD 환경변수로 설정됩니다 (앱 데이터가 아니라 서버 설정).",
        "Admin으로 관리 카테고리에 들어가면 🖥️ 시스템 탭이 먼저 열립니다.",
      ]},
      { t: "시스템 (Admin 전용)", items: [
        "🖥️ System Info — 저장 용량, 참가자·공지·투표 수, 호스트 비번/테마 설정 여부 등을 확인합니다.",
        "💾 백업(Export) — 전체 데이터를 JSON 파일로 내려받습니다. MT가 끝난 뒤 아카이브 용도로 보관할 수 있어요.",
        "📥 복원(Import) — 백업 JSON을 올려 그 시점 상태로 되돌립니다.",
        "🔑 호스트 비번 초기화 — 호스트가 비번을 잊었을 때 초기화합니다. 다음 호스트 로그인 시 새 비번이 설정돼요.",
        "💥 시스템 전체 초기화 — 참가자 계정까지 모두 삭제합니다. 새 MT를 시작할 때 사용하세요.",
      ]},
      { t: "테마 설정", items: [
        "🎨 앱 테마 — 로그인 화면의 배경 이미지, 팀명, MT 기간, 장소를 설정합니다. 배경 이미지는 자동 축소됩니다.",
        "설정 후 로그아웃하면 로그인 화면에서 \"○○ MT에 오신 것을 환영합니다!\" 문구와 함께 확인할 수 있어요.",
      ]},
      { t: "참가자 계정 관리", items: [
        "Admin이 참가자 계정 탭을 열려면 호스트 비밀번호가 필요합니다 (개인정보 보호). 브라우저를 닫으면 다시 잠겨요.",
        "이 잠금은 관리 화면에만 적용됩니다. 팀배정·정산 등 다른 화면과 백업 파일에는 이름이 포함됩니다.",
      ]},
      { t: "게임 참여 팁", items: [
        "Admin도 게임에 끼려면 참가자 계정에 이름을 등록하고, 시크릿 창에서 게스트로 로그인하면 됩니다.",
        "Admin/호스트/게스트는 로그인 단위로 구분되며, 시크릿 창이나 다른 브라우저로 여러 역할을 동시에 쓸 수 있습니다.",
      ]},
    ],
  },
};

function HelpView({ roleKey, onClose, embedded }) {
  const h = HELP[roleKey];
  return (
    <div style={embedded ? {} : { padding: "4px 2px" }}>
      <div style={{ ...s.card, background: `${h.color}0d`, border: `1.5px solid ${h.color}44` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ fontSize: 30 }}>{h.icon}</div>
          <div>
            <div style={{ fontSize: 17, fontWeight: 800, color: h.color }}>{h.label} 사용법</div>
            <div style={{ fontSize: 12.5, color: "#666", marginTop: 3, lineHeight: 1.5 }}>{h.intro}</div>
          </div>
        </div>
      </div>
      {h.sections.map((sec, i) => (
        <div key={i} style={s.card}>
          <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8, color: h.color }}>{sec.t}</div>
          {sec.items.map((it, j) => (
            <div key={j} style={{ display: "flex", gap: 8, padding: "5px 0", fontSize: 13, lineHeight: 1.55, borderBottom: j < sec.items.length - 1 ? "1px solid #f4f4f5" : "none" }}>
              <span style={{ color: h.color, flexShrink: 0 }}>•</span>
              <span>{it}</span>
            </div>
          ))}
        </div>
      ))}
      {onClose && <button onClick={onClose} style={{ ...s.btn("#f1f5f9", "#475569"), width: "100%", borderRadius: 12, marginTop: 2 }}>닫기</button>}
    </div>
  );
}

// ═══════════════════════════════════════
// LOGIN
// ═══════════════════════════════════════
function LoginView({ onLogin, dark, toggleDark }) {
  const [role, setRole] = useState("guest");
  const [name, setName] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [theme, setTheme] = useState(null);
  const [helpRole, setHelpRole] = useState(null);

  useEffect(() => { api.req("theme").then(setTheme).catch(() => setTheme({})); }, []);

  const submit = async () => {
    setErr(""); setBusy(true);
    try {
      const d = await api.req("login", "POST", { role, name: role === "guest" ? name : undefined, password: role === "guest" && !pw ? name : pw });
      onLogin(d);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };

  const ROLES = [
    { k: "guest", l: "🙋 게스트", d: "등록된 이름으로 입장" },
    { k: "host", l: "👑 호스트", d: "전체 관리 및 진행" },
    { k: "admin", l: "⚙️ Admin", d: "시스템 관리자" },
  ];

  const t = theme || {};
  const hasBg = !!t.bgImg;
  const fD = (d) => (d || "").replace(/-/g, ".");
  const period = t.dateStart || t.dateEnd ? `${fD(t.dateStart) || "?"} ~ ${fD(t.dateEnd) || "?"}` : null;
  const locked = (role === "guest" && t.lockGuest) || (role === "host" && t.lockHost);

  return (
    <div style={{ ...s.page, position: "relative", justifyContent: "center", padding: "40px 20px", overflow: "hidden", minHeight: "100vh", boxSizing: "border-box" }}>
      {hasBg && <>
        <img src={t.bgImg} alt="" style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", objectFit: "cover", zIndex: 0 }} />
        <div style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", background: "linear-gradient(rgba(15,23,42,0.5), rgba(15,23,42,0.78))", zIndex: 1 }} />
      </>}
      <button onClick={toggleDark} title="다크 모드" style={{ position: "absolute", top: 14, right: 14, zIndex: 3, background: hasBg ? "rgba(255,255,255,0.25)" : "#fff", border: "none", borderRadius: 10, padding: "8px 10px", fontSize: 16, cursor: "pointer", boxShadow: "0 1px 4px rgba(0,0,0,0.12)" }}>{dark ? "☀️" : "🌙"}</button>

      <div style={{ position: "relative", zIndex: 2 }}>
        <div style={{ textAlign: "center", marginBottom: 22, color: hasBg ? "#fff" : "inherit" }}>
          {!hasBg && <div style={{ fontSize: 48 }}>🏕️</div>}
          {t.teamName ? <>
            <div style={{ fontSize: 30, fontWeight: 800, marginTop: 4, textShadow: hasBg ? "0 2px 10px rgba(0,0,0,0.45)" : "none", lineHeight: 1.3 }}>{t.teamName}</div>
            <div style={{ fontSize: 16, fontWeight: 600, marginTop: 8, textShadow: hasBg ? "0 1px 6px rgba(0,0,0,0.4)" : "none" }}>{t.welcome || "MT에 오신 것을 환영합니다! 🎉"}</div>
          </> : <>
            <div style={{ fontSize: 26, fontWeight: 800, marginTop: 4, textShadow: hasBg ? "0 2px 10px rgba(0,0,0,0.45)" : "none" }}>MT 도우미</div>
            <div style={{ fontSize: 13, opacity: hasBg ? 0.9 : 0.55, marginTop: 4 }}>준비부터 게임까지 한 곳에서</div>
          </>}
          {(period || t.place) && <div style={{ marginTop: 12, fontSize: 13.5, fontWeight: 500, textShadow: hasBg ? "0 1px 6px rgba(0,0,0,0.4)" : "none", opacity: hasBg ? 0.95 : 0.75, display: "flex", flexDirection: "column", gap: 4 }}>
            {period && <div>📅 {period}</div>}
            {t.place && <div>📍 {t.placeUrl ? <a href={t.placeUrl} target="_blank" rel="noopener noreferrer" style={{ color: hasBg ? "#fff" : "#7c3aed", textDecoration: "underline" }}>{t.place}</a> : t.place}</div>}
          </div>}
        </div>

        <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
          {ROLES.map(r => (
            <button key={r.k} onClick={() => { setRole(r.k); setErr(""); }} style={{ flex: 1, padding: "12px 4px", borderRadius: 12, border: role === r.k ? "2px solid #7c3aed" : "2px solid #e5e7eb", background: role === r.k ? "#ede9fe" : "#fff", cursor: "pointer" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: role === r.k ? "#7c3aed" : "#666" }}>{r.l}</div>
              <div style={{ fontSize: 10, color: "#999", marginTop: 2 }}>{r.d}</div>
            </button>
          ))}
        </div>
        <div style={s.card}>
          {locked && <div style={{ padding: "12px 14px", background: "#f1f5f9", borderRadius: 10, marginBottom: 12, textAlign: "center", border: "1.5px solid #cbd5e1" }}>
            <div style={{ fontSize: 22 }}>🔒</div>
            <div style={{ fontSize: 13.5, fontWeight: 700, color: "#475569", marginTop: 4 }}>현재 Admin이 접속을 잠궈두었습니다!</div>
            <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 3 }}>잠시 후 다시 시도해주세요</div>
          </div>}
          {role === "guest" && <>
            <div style={s.label}>이름</div>
            <input disabled={locked} style={{ ...s.input("100%"), marginBottom: 10, background: locked ? "#e5e7eb" : "#fafafa", color: locked ? "#9ca3af" : "inherit", cursor: locked ? "not-allowed" : "text" }} placeholder={locked ? "접속이 잠겨 있습니다" : "호스트가 등록한 이름"} value={locked ? "" : name} onChange={e => setName(e.target.value)} onKeyDown={e => e.key === "Enter" && submit()} />
            {!locked && <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 4 }}>💡 비밀번호는 본인 이름과 동일합니다 (자동 입력됨)</div>}
          </>}
          {role !== "guest" && <>
            <div style={s.label}>{role === "host" ? "호스트 비밀번호" : "관리자 비밀번호"}</div>
            <input disabled={locked} type="password" style={{ ...s.input("100%"), marginBottom: 6, background: locked ? "#e5e7eb" : "#fafafa", color: locked ? "#9ca3af" : "inherit", cursor: locked ? "not-allowed" : "text" }} placeholder={locked ? "접속이 잠겨 있습니다" : "비밀번호"} value={locked ? "" : pw} onChange={e => setPw(e.target.value)} onKeyDown={e => e.key === "Enter" && submit()} />
            {role === "host" && !locked && <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 4 }}>💡 최초 로그인 시 입력한 비밀번호가 그대로 설정됩니다 (4자 이상)</div>}
          </>}
          {err && <div style={{ padding: "8px 10px", background: "#fee2e2", borderRadius: 8, fontSize: 13, color: "#dc2626", marginTop: 6 }}>{err}</div>}
          <button onClick={submit} disabled={busy || locked} style={{ ...s.btn(locked ? "#cbd5e1" : "#7c3aed"), width: "100%", marginTop: 12, padding: "13px", borderRadius: 12, fontSize: 15, opacity: (busy || locked) ? 0.6 : 1, cursor: locked ? "not-allowed" : "pointer" }}>
            {locked ? "🔒 접속 잠김" : busy ? "확인 중..." : "입장하기"}
          </button>
          <button onClick={() => setHelpRole(role)} style={{ background: "none", border: "none", color: "#7c3aed", fontSize: 12.5, cursor: "pointer", textDecoration: "underline", width: "100%", marginTop: 12, padding: 0 }}>
            ❓ {HELP[role].label} 사용법 보기
          </button>
        </div>
      </div>

      {helpRole && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", zIndex: 50, display: "flex", justifyContent: "center", alignItems: "flex-start", overflowY: "auto", padding: "24px 0" }} onClick={() => setHelpRole(null)}>
          <div style={{ background: "#f3f4f6", borderRadius: 16, width: "100%", maxWidth: 480, margin: "0 14px", padding: 14, boxShadow: "0 8px 40px rgba(0,0,0,0.3)" }} onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
              {["guest", "host", "admin"].map(rk => (
                <button key={rk} onClick={() => setHelpRole(rk)} style={{ flex: 1, padding: "9px 4px", borderRadius: 10, border: helpRole === rk ? `2px solid ${HELP[rk].color}` : "2px solid #e5e7eb", background: helpRole === rk ? `${HELP[rk].color}12` : "#fff", color: helpRole === rk ? HELP[rk].color : "#888", fontSize: 12.5, fontWeight: helpRole === rk ? 700 : 500, cursor: "pointer" }}>{HELP[rk].icon} {HELP[rk].label}</button>
              ))}
            </div>
            <HelpView roleKey={helpRole} onClose={() => setHelpRole(null)} embedded />
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════
// FEATURE: MEMBERS (참가자 계정 관리)
// ═══════════════════════════════════════
function MembersView({ S, save, canEdit, role }) {
  const [inp, setInp] = useState("");
  const members = S.members || [];
  const add = () => {
    const nn = inp.split(/[,\n]/).map(n => n.trim()).filter(n => n && !members.find(m => m.name === n));
    if (nn.length) save({ members: [...members, ...nn.map(n => ({ id: uid(), name: n }))] });
    setInp("");
  };
  const [confirmReset, setConfirmReset] = useState(false);
  // Admin은 호스트 비밀번호 확인 후에만 계정 관리 화면 접근 (호스트는 잠금 없음)
  const [unlocked, setUnlocked] = useState(() => role !== "admin" || sessionStorage.getItem("mt-members-unlock") === "1");
  const [lockPw, setLockPw] = useState("");
  const [lockErr, setLockErr] = useState("");
  const [lockBusy, setLockBusy] = useState(false);
  const tryUnlock = async () => {
    setLockErr(""); setLockBusy(true);
    try {
      await api.req("admin/unlock-members", "POST", { password: lockPw });
      sessionStorage.setItem("mt-members-unlock", "1");
      setUnlocked(true);
    } catch (e) { setLockErr(e.message); }
    setLockBusy(false);
  };

  if (!unlocked) return (
    <div style={{ ...s.card, textAlign: "center", padding: "30px 20px" }}>
      <div style={{ fontSize: 40 }}>🔒</div>
      <div style={{ fontSize: 15, fontWeight: 700, margin: "10px 0 4px" }}>참가자 계정 관리 잠금</div>
      <div style={{ fontSize: 12.5, color: "#888", lineHeight: 1.6, marginBottom: 14 }}>
        개인정보 보호를 위해 Admin은 <b>호스트 비밀번호</b>를 입력해야<br />참가자 계정을 열람·관리할 수 있습니다.
      </div>
      <input style={{ ...s.input("100%"), textAlign: "center", marginBottom: 8 }} type="password" placeholder="호스트 비밀번호" value={lockPw} onChange={e => setLockPw(e.target.value)} onKeyDown={e => e.key === "Enter" && tryUnlock()} />
      {lockErr && <div style={{ padding: "8px 10px", background: "#fee2e2", borderRadius: 8, fontSize: 13, color: "#dc2626", marginBottom: 8 }}>{lockErr}</div>}
      <button onClick={tryUnlock} disabled={lockBusy} style={{ ...s.btn("#475569"), width: "100%", borderRadius: 10 }}>{lockBusy ? "확인 중..." : "잠금 해제"}</button>
      <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 10 }}>이 잠금은 브라우저를 닫으면 다시 걸립니다</div>
    </div>
  );

  return (
    <>
      <div style={s.card}>
        <div style={s.section}>👤 참가자 계정 ({members.length}명)</div>
        {canEdit && <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
          <input style={{ ...s.input("auto"), flex: 1 }} placeholder="이름 (쉼표로 여러 명 가능)" value={inp} onChange={e => setInp(e.target.value)} onKeyDown={e => e.key === "Enter" && add()} />
          <button onClick={add} style={s.btn("#475569")}>추가</button>
        </div>}
        {members.length === 0 ? <div style={{ textAlign: "center", padding: 20, color: "#bbb", fontSize: 13 }}>참가자를 등록하면 그 이름으로 게스트 로그인이 가능해집니다</div> :
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {members.map(m => <span key={m.id} style={{ ...s.tag("#f1f5f9", "#475569"), padding: "6px 12px", fontSize: 13, cursor: canEdit ? "pointer" : "default" }} onClick={() => canEdit && save({ members: members.filter(x => x.id !== m.id) })}>{m.name}{canEdit && " ×"}</span>)}
          </div>}
        {canEdit && <div style={{ marginTop: 10, padding: "8px 10px", background: "#f8fafc", borderRadius: 8, fontSize: 12, color: "#64748b" }}>
          💡 게스트는 <b>이름 = 비밀번호</b>로 로그인합니다. 이 명단이 룰렛·팀배정·N빵·투표에 공통 사용됩니다.
        </div>}
      </div>
      {canEdit && <div style={s.card}>
        <div style={s.section}>🔄 게임 데이터 초기화</div>
        <div style={{ fontSize: 12, color: "#888", marginBottom: 10 }}>장바구니·점수·투표 등 게임 데이터를 초기화합니다. <b>참가자 계정과 비밀번호는 유지</b>됩니다.</div>
        {!confirmReset ? <button onClick={() => setConfirmReset(true)} style={{ ...s.btn("#fef2f2", "#dc2626"), width: "100%", borderRadius: 10 }}>게임 데이터 초기화</button> :
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={async () => { const d = await api.req("host/reset", "POST", {}); save.replace(d); setConfirmReset(false); }} style={{ ...s.btn("#dc2626"), flex: 1 }}>확인 (초기화)</button>
            <button onClick={() => setConfirmReset(false)} style={{ ...s.btn("#f5f5f5", "#666"), flex: 1 }}>취소</button>
          </div>}
      </div>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: SYSTEM (Admin 전용)
// ═══════════════════════════════════════
function SystemView({ S, save }) {
  const [info, setInfo] = useState(null);
  const [err, setErr] = useState("");
  const [msg, setMsg] = useState("");
  const fileRef = useRef(null);
  const bgRef = useRef(null);
  const [confirmWipe, setConfirmWipe] = useState(false);
  // 테마 편집 로컬 상태 (저장 버튼 눌러야 서버 반영)
  const [th, setTh] = useState(() => ({ bgImg: null, teamName: "", welcome: "", dateStart: "", dateEnd: "", place: "", placeUrl: "", ...(S?.theme || {}) }));
  const [thBusy, setThBusy] = useState(false);

  const attachBg = async (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    setThBusy(true);
    try { setTh(p => ({ ...p, bgImg: null })); const img = await resizeImage(f, 1200); setTh(p => ({ ...p, bgImg: img })); }
    catch (er) { alert(er.message); }
    setThBusy(false); e.target.value = "";
  };
  const saveTheme = async () => {
    setErr(""); setThBusy(true);
    try {
      const d = await api.req("update", "POST", { patch: { theme: th } });
      save.replace(d);
      setMsg("테마가 저장되었습니다. 로그아웃하면 로그인 화면에서 확인할 수 있어요.");
    } catch (e) { setErr(e.message); }
    setThBusy(false);
  };

  const loadInfo = async () => { setErr(""); try { const d = await api.req("admin/sysinfo"); setInfo(d); setLock({ guest: d.lockGuest, host: d.lockHost }); } catch (e) { setErr(e.message); } };
  useEffect(() => { loadInfo(); }, []);
  const [lock, setLock] = useState({ guest: false, host: false });
  const toggleLock = async (which) => {
    setErr("");
    const next = { ...lock, [which]: !lock[which] };
    setLock(next); // 낙관적
    try { await api.req("admin/lock", "POST", { lockGuest: next.guest, lockHost: next.host }); }
    catch (e) { setErr(e.message); setLock(lock); }
  };

  const doExport = async () => {
    try {
      const d = await api.req("admin/export");
      const blob = new Blob([JSON.stringify(d, null, 2)], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `mt-helper-backup_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.json`;
      a.click(); URL.revokeObjectURL(a.href);
      setMsg("백업 파일이 다운로드되었습니다.");
    } catch (e) { setErr(e.message); }
  };

  const doImport = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader();
    r.onload = async (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        const st = parsed.state || parsed;
        if (!window.confirm("현재 데이터를 백업 파일 내용으로 완전히 덮어씁니다. 계속할까요?")) return;
        const d = await api.req("admin/import", "POST", { state: st });
        save.replace(d.state);
        setMsg("복원 완료!"); loadInfo();
      } catch (e2) { setErr("복원 실패: " + e2.message); }
      e.target.value = "";
    };
    r.readAsText(f);
  };

  const Info = ({ l, v }) => <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f1f5f9", fontSize: 13 }}><span style={{ color: "#64748b" }}>{l}</span><span style={{ fontWeight: 600 }}>{v}</span></div>;

  return (
    <>
      {msg && <div style={{ padding: "10px 12px", background: "#dcfce7", borderRadius: 10, fontSize: 13, color: "#166534", marginBottom: 10 }}>{msg}</div>}
      {err && <div style={{ padding: "10px 12px", background: "#fee2e2", borderRadius: 10, fontSize: 13, color: "#dc2626", marginBottom: 10 }}>{err}</div>}

      <div style={s.card}>
        <div style={s.section}>🎨 앱 테마 (로그인 화면)</div>
        <div style={s.label}>프로젝트 팀명</div>
        <input style={{ ...s.input("100%"), marginBottom: 10 }} placeholder="예: 컴퓨터공학과 26학번" maxLength={40} value={th.teamName} onChange={e => setTh(p => ({ ...p, teamName: e.target.value }))} />
        <div style={s.label}>환영 문구</div>
        <input style={{ ...s.input("100%"), marginBottom: 10 }} placeholder="MT에 오신 것을 환영합니다! 🎉" maxLength={60} value={th.welcome} onChange={e => setTh(p => ({ ...p, welcome: e.target.value }))} />
        <div style={s.label}>MT 기간</div>
        <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 10 }}>
          <input style={{ ...s.input("auto"), flex: 1 }} type="date" value={th.dateStart} onChange={e => setTh(p => ({ ...p, dateStart: e.target.value }))} />
          <span style={{ color: "#999" }}>~</span>
          <input style={{ ...s.input("auto"), flex: 1 }} type="date" value={th.dateEnd} onChange={e => setTh(p => ({ ...p, dateEnd: e.target.value }))} />
        </div>
        <div style={s.label}>장소명</div>
        <input style={{ ...s.input("100%"), marginBottom: 8 }} placeholder="예: 가평 OO펜션" maxLength={40} value={th.place} onChange={e => setTh(p => ({ ...p, place: e.target.value }))} />
        <div style={s.label}>장소 링크 (선택 — 지도 URL)</div>
        <input style={{ ...s.input("100%"), marginBottom: 10 }} placeholder="https://naver.me/... (비워두면 링크 없음)" value={th.placeUrl} onChange={e => setTh(p => ({ ...p, placeUrl: e.target.value }))} />
        <div style={s.label}>배경 이미지</div>
        {th.bgImg ? <div style={{ position: "relative", marginBottom: 8 }}>
          <img src={th.bgImg} alt="배경 미리보기" style={{ width: "100%", height: 120, objectFit: "cover", borderRadius: 10, display: "block" }} />
          <button onClick={() => setTh(p => ({ ...p, bgImg: null }))} style={{ position: "absolute", top: 6, right: 6, background: "rgba(0,0,0,0.55)", color: "#fff", border: "none", borderRadius: 8, padding: "4px 10px", fontSize: 12, cursor: "pointer" }}>제거</button>
        </div> : <button onClick={() => bgRef.current?.click()} disabled={thBusy} style={{ ...s.btn("#f5f5f5", "#555"), width: "100%", borderRadius: 10, marginBottom: 8, fontSize: 13 }}>{thBusy ? "처리 중..." : "🖼️ 이미지 선택 (자동 축소, 최대 1200px)"}</button>}
        <input ref={bgRef} type="file" accept="image/*" style={{ display: "none" }} onChange={attachBg} />
        <button onClick={saveTheme} disabled={thBusy} style={{ ...s.btn("#475569"), width: "100%", borderRadius: 10 }}>{thBusy ? "저장 중..." : "테마 저장"}</button>
        <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 6 }}>💡 로그인 화면에 팀명·환영 문구·기간·장소가 표시됩니다. 장소 링크를 넣으면 장소명이 눌러지는 링크가 됩니다.</div>
      </div>

      <div style={s.card}>
        <div style={s.section}>🔒 로그인 잠금</div>
        <div style={{ fontSize: 12.5, color: "#666", marginBottom: 12, lineHeight: 1.6 }}>잠그면 해당 역할은 로그인 화면에서 입력칸이 막히고 "Admin이 접속을 잠궈두었습니다" 안내가 뜹니다. <b>배포 직후 미리 들어오지 못하게</b> 막고, 준비가 되면 순서대로 풀어주세요. (Admin은 잠금과 무관하게 항상 로그인 가능)</div>
        {[{ k: "host", label: "👑 호스트", desc: "세팅할 호스트만 먼저 열어줄 때 해제" }, { k: "guest", label: "🙋 게스트", desc: "호스트 세팅이 끝나면 마지막에 해제" }].map(row => (
          <div key={row.k} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderTop: "1px solid #f1f5f9" }}>
            <div>
              <div style={{ fontSize: 14, fontWeight: 600 }}>{row.label}</div>
              <div style={{ fontSize: 11.5, color: "#94a3b8", marginTop: 2 }}>{row.desc}</div>
            </div>
            <button onClick={() => toggleLock(row.k)} style={{ border: "none", cursor: "pointer", borderRadius: 999, padding: "7px 16px", fontSize: 13, fontWeight: 700, background: lock[row.k] ? "#fee2e2" : "#dcfce7", color: lock[row.k] ? "#dc2626" : "#16a34a", minWidth: 92 }}>
              {lock[row.k] ? "🔒 잠김" : "🔓 열림"}
            </button>
          </div>
        ))}
        {(lock.guest || lock.host) && <div style={{ marginTop: 10, padding: "8px 10px", background: "#fef2f2", borderRadius: 8, fontSize: 12, color: "#b91c1c" }}>현재 잠긴 역할: {[lock.host && "호스트", lock.guest && "게스트"].filter(Boolean).join(", ")}</div>}
      </div>

      <div style={s.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={s.section}>🖥️ System Info</div>
          <button onClick={loadInfo} style={s.smBtn("#f1f5f9", "#475569")}>갱신</button>
        </div>
        {info ? <>
          <Info l="저장 데이터 크기" v={`${fmt(info.storedBytes)} bytes (${(info.storedBytes / info.kvLimitBytes * 100).toFixed(2)}% / 25MB)`} />
          <Info l="참가자 계정" v={`${info.members}명`} />
          <Info l="장바구니 물품" v={`${info.cartItems}건`} />
          <Info l="투표" v={`${info.votes}건`} />
          <Info l="일정 / 정산항목" v={`${info.scheduleItems}건 / ${info.settleItems}건`} />
          <Info l="점수판 (팀/라운드)" v={`${info.sbTeams} / ${info.sbRounds}`} />
          <Info l="호스트 비밀번호" v={info.hostPwSet ? "설정됨" : "미설정"} />
          <Info l="테마" v={info.themeSet ? "설정됨 🎨" : "미설정 (기본 화면)"} />
          <Info l="로그인 잠금" v={info.lockHost || info.lockGuest ? `🔒 ${[info.lockHost && "호스트", info.lockGuest && "게스트"].filter(Boolean).join(", ")}` : "없음"} />
          <Info l="ADMIN_PASSWORD 환경변수" v={info.adminPwCustom ? "설정됨 ✅" : "⚠️ 미설정 (기본값 'admin' 사용 중)"} />
          <Info l="방 생성일" v={new Date(info.createdAt).toLocaleString("ko-KR")} />
          <Info l="서버 시간" v={new Date(info.serverTime).toLocaleString("ko-KR")} />
        </> : <div style={{ color: "#bbb", fontSize: 13, textAlign: "center", padding: 10 }}>불러오는 중...</div>}
      </div>

      <div style={s.card}>
        <div style={s.section}>💾 백업 / 복원</div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={doExport} style={{ ...s.btn("#475569"), flex: 1, borderRadius: 10 }}>📤 DB Export</button>
          <button onClick={() => fileRef.current?.click()} style={{ ...s.btn("#f1f5f9", "#475569"), flex: 1, borderRadius: 10 }}>📥 복원 (Import)</button>
          <input ref={fileRef} type="file" accept=".json" style={{ display: "none" }} onChange={doImport} />
        </div>
        <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 8 }}>Export에는 호스트 비밀번호를 포함한 전체 DB가 담깁니다.</div>
      </div>

      <div style={s.card}>
        <div style={s.section}>🛠️ 관리 작업</div>
        <button onClick={async () => { if (window.confirm("호스트 비밀번호를 초기화할까요? 다음 호스트 로그인 시 새 비밀번호가 설정됩니다.")) { const d = await api.req("admin/hostpw", "POST", {}); setMsg(d.message); loadInfo(); } }}
          style={{ ...s.btn("#fff7ed", "#c2410c"), width: "100%", borderRadius: 10, marginBottom: 8 }}>🔑 호스트 비밀번호 초기화</button>
        {!confirmWipe ? <button onClick={() => setConfirmWipe(true)} style={{ ...s.btn("#fef2f2", "#dc2626"), width: "100%", borderRadius: 10 }}>💥 시스템 전체 초기화 (계정 포함)</button> :
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={async () => { const d = await api.req("admin/reset", "POST", {}); setMsg(d.message); setConfirmWipe(false); const st = await api.req("state"); save.replace(st); loadInfo(); }} style={{ ...s.btn("#dc2626"), flex: 1 }}>진짜 전체 초기화</button>
            <button onClick={() => setConfirmWipe(false)} style={{ ...s.btn("#f5f5f5", "#666"), flex: 1 }}>취소</button>
          </div>}
      </div>
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: SHOPPING CART
// ═══════════════════════════════════════
function CartView({ S, save, canEdit }) {
  const [tab, setTab] = useState("mart");
  const fileRef = useRef(null);
  const cart = S.cart || { mart: [], online: [], donated: [], pre: [] };
  const items = cart[tab] || [];
  const showP = tab !== "donated";
  const upd = (id, f, v) => save({ cart: { ...cart, [tab]: items.map(i => i.id === id ? { ...i, [f]: v } : i) } });

  const catT = items.reduce((a, i) => a + (showP ? (i.price || 0) * i.quantity : 0), 0);
  const catQ = items.reduce((a, i) => a + i.quantity, 0);
  const buyT = ["mart", "online"].reduce((a, k) => a + (cart[k] || []).reduce((x, i) => x + (i.price || 0) * i.quantity, 0), 0);
  const preT = (cart.pre || []).reduce((a, i) => a + (i.price || 0) * i.quantity, 0);

  const exportXl = () => {
    const wb = XLSX.utils.book_new(); let has = false;
    CART_TABS.forEach(({ key }) => {
      const l = cart[key] || []; if (!l.length) return; has = true;
      const isD = key === "donated";
      const rows = l.map(i => ({ 품목: i.name || "(미입력)", ...(isD ? {} : { "단가(₩)": i.price }), 수량: i.quantity, ...(isD ? {} : { "소계(₩)": (i.price || 0) * i.quantity }) }));
      rows.push(isD ? { 품목: "【합계】", 수량: l.reduce((a, i) => a + i.quantity, 0) } : { 품목: "【합계】", "단가(₩)": "", 수량: l.reduce((a, i) => a + i.quantity, 0), "소계(₩)": l.reduce((a, i) => a + (i.price || 0) * i.quantity, 0) });
      const ws = XLSX.utils.json_to_sheet(rows); ws["!cols"] = Object.keys(rows[0]).map(k => ({ wch: Math.max(k.length * 2, 12) }));
      XLSX.utils.book_append_sheet(wb, ws, TAB_TO_SHEET[key]);
    });
    if (!has) { alert("내보낼 데이터가 없습니다."); return; }
    XLSX.writeFile(wb, `장보기_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const importXl = (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" }); const nc = { ...cart }; let cnt = 0;
        wb.SheetNames.forEach(sn => {
          const tk = SHEET_TO_TAB[sn.trim()]; if (!tk) return;
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn]); const isD = tk === "donated";
          const parsed = rows.filter(r => r["품목"] && r["품목"] !== "【합계】").map(r => ({ id: uid(), name: String(r["품목"] || ""), price: isD ? 0 : (Number(r["단가(₩)"]) || 0), quantity: Number(r["수량"]) || 1 }));
          if (parsed.length) { nc[tk] = [...parsed, ...(nc[tk] || [])]; cnt += parsed.length; }
        });
        if (cnt) { save({ cart: nc }); alert(`${cnt}개 물품을 가져왔습니다.`); } else alert("가져올 데이터가 없습니다. 시트 이름이 '마트/온라인/기증/사전구매'인지 확인하세요.");
      } catch { alert("파일 읽기 오류"); } e.target.value = "";
    }; reader.readAsArrayBuffer(file);
  };

  return (
    <>
      {!canEdit && <RoNote />}
      <div style={{ display: "flex", gap: 5, marginBottom: 12, overflowX: "auto" }}>
        {CART_TABS.map(t => { const a = tab === t.key; const cnt = (cart[t.key] || []).length;
          return <button key={t.key} onClick={() => setTab(t.key)} style={{ flex: "1 0 auto", padding: "9px 6px", borderRadius: 10, border: a ? `2px solid ${t.c}` : "2px solid #e5e7eb", background: a ? `${t.c}11` : "#fff", color: a ? t.c : "#888", fontSize: 12, fontWeight: a ? 700 : 500, cursor: "pointer", whiteSpace: "nowrap" }}>
            {t.label}{cnt > 0 && <span style={{ ...s.tag(a ? t.c : "#e5e7eb", a ? "#fff" : "#888"), marginLeft: 4 }}>{cnt}</span>}
          </button>;
        })}
      </div>
      {tab === "pre" && <div style={{ padding: "8px 12px", background: "#f5f3ff", borderRadius: 10, fontSize: 12, color: "#6d28d9", marginBottom: 10 }}>✅ 이미 구매 완료된 물품 목록입니다. 전체 예산 합계와 별도로 집계됩니다.</div>}
      {canEdit && <button onClick={() => save({ cart: { ...cart, [tab]: [{ id: uid(), name: "", price: 0, quantity: 1 }, ...items] } })} style={{ ...s.btn("#0d9488"), width: "100%", marginBottom: 10, padding: "12px 0", borderRadius: 12, fontSize: 15 }}>+ 물품 추가</button>}

      {items.length === 0 ? <div style={{ textAlign: "center", padding: "30px", color: "#bbb" }}><div style={{ fontSize: 36 }}>{tab === "mart" ? "🛒" : tab === "online" ? "📦" : tab === "donated" ? "🎁" : "✅"}</div><div style={{ fontSize: 13, marginTop: 6 }}>{canEdit ? "물품을 추가해보세요" : "등록된 물품이 없습니다"}</div></div> :
        canEdit ? items.map(item => (
          <div key={item.id} style={{ ...s.card, padding: "10px 12px" }}>
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              <input style={{ ...s.input("auto"), flex: 1, minWidth: 80 }} placeholder="품목명" value={item.name} onChange={e => upd(item.id, "name", e.target.value)} />
              <button style={s.del} onClick={() => save({ cart: { ...cart, [tab]: items.filter(i => i.id !== item.id) } })}>×</button>
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "center" }}>
              {showP && <div style={{ flex: 1 }}><div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>단가 (₩)</div><input style={{ ...s.input("100%"), textAlign: "right" }} type="number" inputMode="numeric" placeholder="0" value={item.price || ""} onChange={e => upd(item.id, "price", +e.target.value || 0)} /></div>}
              <div style={{ width: 70 }}><div style={{ fontSize: 11, color: "#999", marginBottom: 2 }}>수량</div>
                <div style={{ display: "flex" }}>
                  <button style={{ ...s.smBtn("#eee", "#333"), borderRadius: "8px 0 0 8px", padding: "8px 10px" }} onClick={() => upd(item.id, "quantity", Math.max(1, item.quantity - 1))}>−</button>
                  <input style={{ width: 36, textAlign: "center", padding: "7px 0", border: "1.5px solid #ddd", borderLeft: "none", borderRight: "none", fontSize: 14, outline: "none", background: "#fafafa" }} type="number" inputMode="numeric" value={item.quantity} onChange={e => upd(item.id, "quantity", Math.max(1, +e.target.value || 1))} />
                  <button style={{ ...s.smBtn("#eee", "#333"), borderRadius: "0 8px 8px 0", padding: "8px 10px" }} onClick={() => upd(item.id, "quantity", item.quantity + 1)}>+</button>
                </div>
              </div>
              {showP && <div style={{ minWidth: 75, textAlign: "right", paddingTop: 16 }}><span style={{ fontSize: 15, fontWeight: 700, color: "#0d9488" }}>₩{fmt((item.price || 0) * item.quantity)}</span></div>}
            </div>
          </div>
        )) : (
          <div style={s.card}>
            {items.map(item => (
              <div key={item.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #f5f5f5" }}>
                <span style={{ fontSize: 14 }}>{item.name || "(미입력)"} <span style={{ color: "#999", fontSize: 12 }}>×{item.quantity}</span></span>
                {showP && <span style={{ fontSize: 14, fontWeight: 700, color: "#0d9488" }}>₩{fmt((item.price || 0) * item.quantity)}</span>}
              </div>
            ))}
          </div>
        )}

      {canEdit && <div style={{ display: "flex", gap: 8, marginTop: 4, marginBottom: 10 }}>
        <button onClick={exportXl} style={{ ...s.btn("#f5f5f5", "#555"), flex: 1, padding: "10px 0", borderRadius: 12, fontSize: 13 }}>📤 내보내기</button>
        <button onClick={() => fileRef.current?.click()} style={{ ...s.btn("#f5f5f5", "#555"), flex: 1, padding: "10px 0", borderRadius: 12, fontSize: 13 }}>📥 가져오기</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={importXl} />
      </div>}

      <div style={{ ...s.card, background: "linear-gradient(135deg,#f0fdf9,#ccfbf1)", border: "1.5px solid #5eead4" }}>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}><span style={{ color: "#666" }}>현재 탭 ({CART_TABS.find(t => t.key === tab)?.label})</span><b>{catQ}건{showP && ` · ₩${fmt(catT)}`}</b></div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}><span style={{ color: "#666" }}>구매 예정 (마트+온라인)</span><b style={{ color: "#0d9488" }}>₩{fmt(buyT)}</b></div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}><span style={{ color: "#666" }}>사전구매 (지출 완료)</span><b style={{ color: "#7c3aed" }}>₩{fmt(preT)}</b></div>
      </div>
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: ALCOHOL
// ═══════════════════════════════════════
function AlcView({ S, save, canEdit }) {
  const alc = S.alc || { people: 5, sojuPer: 2, drinks: [] };
  const { people, sojuPer, drinks } = alc;
  const totalNeed = people * sojuPer;
  const equiv = (d) => d.v && d.a ? (d.v * d.a / 100 / SOJU_ALC) * (d.qty || 1) : 0;
  const totalEq = drinks.reduce((a, d) => a + equiv(d), 0);
  const remain = Math.max(0, totalNeed - totalEq);
  const setA = (p) => save({ alc: { ...alc, ...p } });
  const [showP, setShowP] = useState(false);
  const RB = ({ l, v, h }) => <div style={{ padding: "10px 12px", background: h ? "#0d9488" : "#fff", borderRadius: 10, border: h ? "none" : "1px solid #eee" }}><div style={{ fontSize: 11, color: h ? "rgba(255,255,255,.7)" : "#999" }}>{l}</div><div style={{ fontSize: 18, fontWeight: 800, color: h ? "#fff" : "#0d9488" }}>{v}</div></div>;

  return (
    <>
      {!canEdit && <RoNote />}
      <div style={{ ...s.card, background: "linear-gradient(135deg,#fef9ee,#fff7e6)", border: "1.5px solid #f0d98d" }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: "#92702a" }}>🍶 기준: 소주 1병 = {SOJU_ML}ml · {SOJU_ABV}% · 알코올 {SOJU_ALC.toFixed(1)}ml</div>
      </div>
      <div style={s.card}>
        <div style={s.section}>👥 모임 설정</div>
        {canEdit ? <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}><div style={s.label}>인원</div><Stepper value={people} onChange={v => setA({ people: v })} /></div>
          <div style={{ flex: 1 }}><div style={s.label}>1인당 소주(병)</div><Stepper value={sojuPer} onChange={v => setA({ sojuPer: v })} min={0.5} step={0.5} decimal /></div>
        </div> : <div style={{ fontSize: 14 }}>인원 <b>{people}명</b> · 1인당 <b>{sojuPer}병</b></div>}
        <div style={{ marginTop: 12, padding: "10px", background: "#f0fdf9", borderRadius: 10, textAlign: "center" }}>
          <span style={{ fontSize: 13, color: "#666" }}>필요 소주</span><span style={{ fontSize: 22, fontWeight: 800, color: "#0d9488", marginLeft: 8 }}>{totalNeed}병</span>
        </div>
      </div>
      <div style={s.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={s.section}>🍷 다른 주류</div>
          {canEdit && <button onClick={() => setShowP(!showP)} style={{ ...s.smBtn("#0d9488"), borderRadius: 8, padding: "6px 12px" }}>+ 추가</button>}
        </div>
        {showP && canEdit && <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12, padding: 10, background: "#f9f9f9", borderRadius: 10 }}>
          {DRINK_PRESETS.map((p, i) => <button key={i} onClick={() => { setA({ drinks: [{ id: uid(), name: p.name === "직접 입력" ? "" : p.name, v: p.v, a: p.a, qty: 1 }, ...drinks] }); setShowP(false); }} style={{ ...s.smBtn("#fff", "#333"), border: "1.5px solid #ddd", borderRadius: 8, padding: "8px 12px", fontSize: 13 }}>{p.name}</button>)}
        </div>}
        {drinks.length === 0 && <div style={{ fontSize: 13, color: "#bbb", textAlign: "center", padding: 8 }}>등록된 주류가 없습니다</div>}
        {drinks.map(d => canEdit ? <div key={d.id} style={{ padding: "10px 0", borderBottom: "1px solid #f0f0f0" }}>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 6 }}>
            <input style={{ ...s.input("auto"), flex: 1, fontWeight: 600 }} placeholder="주류명" value={d.name} onChange={e => setA({ drinks: drinks.map(x => x.id === d.id ? { ...x, name: e.target.value } : x) })} />
            <button style={s.del} onClick={() => setA({ drinks: drinks.filter(x => x.id !== d.id) })}>×</button>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <div style={{ flex: 1 }}><div style={{ fontSize: 11, color: "#999" }}>ml</div><input style={{ ...s.input("100%"), textAlign: "right" }} type="number" inputMode="numeric" value={d.v || ""} onChange={e => setA({ drinks: drinks.map(x => x.id === d.id ? { ...x, v: +e.target.value || 0 } : x) })} /></div>
            <div style={{ flex: 1 }}><div style={{ fontSize: 11, color: "#999" }}>도수%</div><input style={{ ...s.input("100%"), textAlign: "right" }} type="number" inputMode="decimal" value={d.a || ""} onChange={e => setA({ drinks: drinks.map(x => x.id === d.id ? { ...x, a: +e.target.value || 0 } : x) })} /></div>
            <div style={{ width: 65 }}><div style={{ fontSize: 11, color: "#999" }}>수량</div><Stepper value={d.qty || 1} onChange={v => setA({ drinks: drinks.map(x => x.id === d.id ? { ...x, qty: v } : x) })} /></div>
          </div>
          <div style={{ marginTop: 4, fontSize: 12, color: "#0d9488", fontWeight: 600 }}>≈ 소주 {equiv(d).toFixed(1)}병</div>
        </div> : <div key={d.id} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", fontSize: 14, borderBottom: "1px solid #f5f5f5" }}>
          <span>{d.name || "(무명)"} {d.v}ml·{d.a}% ×{d.qty || 1}</span><b style={{ color: "#0d9488" }}>≈{equiv(d).toFixed(1)}병</b>
        </div>)}
      </div>
      <div style={{ ...s.card, background: totalEq > 0 ? "linear-gradient(135deg,#f0fdf9,#ccfbf1)" : "#fff", border: totalEq > 0 ? "1.5px solid #5eead4" : "1.5px solid #eee" }}>
        <div style={s.section}>📊 결과</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <RB l="필요 소주" v={`${totalNeed}병`} /><RB l="다른 주류" v={`≈${totalEq.toFixed(1)}병`} />
          <RB l="남은 소주" v={`${remain.toFixed(1)}병`} h /><RB l="1인당" v={`${(remain / people).toFixed(1)}병`} />
        </div>
        {totalEq >= totalNeed && totalEq > 0 && <div style={{ marginTop: 8, padding: "8px", background: "#fff3cd", borderRadius: 8, fontSize: 12, color: "#856404", textAlign: "center" }}>⚠️ 소주 추가 구매 불필요!</div>}
      </div>
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: ROULETTE
// ═══════════════════════════════════════
function RouletteView({ S, save, canEdit }) {
  const names = (S.members || []).map(m => m.name);
  const rl = S.roulette || { current: null, history: [] };
  const [spinning, setSpinning] = useState(false);
  const [display, setDisplay] = useState(null);

  const spin = () => {
    if (names.length < 2 || spinning) return;
    setSpinning(true);
    let count = 0; const total = 20 + Math.floor(Math.random() * 10);
    const finalIdx = Math.floor(Math.random() * names.length);
    const run = () => {
      if (count < total) { setDisplay(names[count % names.length]); count++; setTimeout(run, 60 + count * 15); }
      else {
        const picked = names[finalIdx]; setDisplay(picked); setSpinning(false);
        save({ roulette: { current: picked, history: [{ name: picked, t: Date.now() }, ...rl.history].slice(0, 20) } });
      }
    }; run();
  };

  const shown = spinning ? display : (rl.current || "?");

  return (
    <>
      <div style={s.card}>
        <div style={s.section}>👤 참가자 ({names.length}명)</div>
        {names.length === 0 ? <div style={{ fontSize: 13, color: "#bbb", textAlign: "center" }}>⚙️ 관리 → 참가자 계정에서 등록하세요</div> :
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{names.map(n => <span key={n} style={{ ...s.tag("#ede9fe", "#7c3aed"), padding: "5px 10px", fontSize: 13 }}>{n}</span>)}</div>}
      </div>
      <div style={{ ...s.card, textAlign: "center", padding: "24px 16px" }}>
        <div style={{ fontSize: 44, fontWeight: 800, color: spinning ? "#7c3aed" : rl.current ? "#059669" : "#ddd", minHeight: 56, display: "flex", alignItems: "center", justifyContent: "center", transition: "color 0.2s" }}>{shown}</div>
        {!spinning && rl.current && <div style={{ fontSize: 14, color: "#666", marginTop: 4 }}>🎉 최근 당첨</div>}
        {canEdit ? <button onClick={spin} disabled={spinning || names.length < 2} style={{ ...s.btn(names.length < 2 ? "#ccc" : "#7c3aed"), width: "100%", marginTop: 16, padding: "14px", borderRadius: 12, fontSize: 16 }}>{spinning ? "돌리는 중..." : "🎯 룰렛 돌리기"}</button>
          : <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 12 }}>호스트가 룰렛을 돌리면 결과가 여기 표시됩니다 (새로고침)</div>}
      </div>
      {rl.history.length > 0 && <div style={s.card}>
        <div style={s.section}>📜 히스토리</div>
        {rl.history.map((h, i) => <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 13, color: "#666", borderBottom: "1px solid #f5f5f5" }}><span>{h.name}</span><span style={{ color: "#bbb" }}>{timeStr(h.t)}</span></div>)}
        {canEdit && <button onClick={() => save({ roulette: { current: null, history: [] } })} style={{ ...s.smBtn("#f5f5f5", "#999"), marginTop: 8 }}>초기화</button>}
      </div>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: TEAMS
// ═══════════════════════════════════════
function TeamsView({ S, save, canEdit, myName }) {
  const names = (S.members || []).map(m => m.name);
  const tm = S.teams || { numTeams: 2, result: null };
  const TC = ["#7c3aed", "#0d9488", "#e11d48", "#d97706", "#2563eb", "#16a34a", "#dc2626", "#7c2d12"];
  const assign = () => {
    if (names.length < tm.numTeams) return;
    const sh = shuffle(names); const t = Array.from({ length: tm.numTeams }, () => []);
    sh.forEach((n, i) => t[i % tm.numTeams].push(n));
    save({ teams: { ...tm, result: t } });
  };

  return (
    <>
      <div style={s.card}>
        <div style={s.section}>👤 참가자 ({names.length}명)</div>
        {names.length === 0 ? <div style={{ fontSize: 13, color: "#bbb", textAlign: "center" }}>⚙️ 관리 → 참가자 계정에서 등록하세요</div> :
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{names.map(n => <span key={n} style={{ ...s.tag("#ede9fe", "#7c3aed"), padding: "5px 10px", fontSize: 13 }}>{n}</span>)}</div>}
      </div>
      {canEdit && <div style={s.card}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}><div style={s.label}>팀 수</div><Stepper value={tm.numTeams} onChange={v => save({ teams: { numTeams: v, result: null } })} min={2} /></div>
        <button onClick={assign} disabled={names.length < tm.numTeams} style={{ ...s.btn(names.length < tm.numTeams ? "#ccc" : "#7c3aed"), width: "100%", padding: "12px", borderRadius: 12, fontSize: 15 }}>🔀 팀 배정하기</button>
      </div>}
      {tm.result ? tm.result.map((team, i) => {
        const mine = team.includes(myName);
        return (
          <div key={i} style={{ ...s.card, borderLeft: `4px solid ${TC[i % TC.length]}`, ...(mine ? { boxShadow: `0 0 0 2px ${TC[i % TC.length]}55` } : {}) }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: TC[i % TC.length], marginBottom: 6 }}>팀 {i + 1} ({team.length}명){mine && " · ⭐ 내 팀"}</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{team.map(n => <span key={n} style={{ ...s.tag(n === myName ? TC[i % TC.length] : "#f3f4f6", n === myName ? "#fff" : "#333") }}>{n}</span>)}</div>
          </div>
        );
      }) : !canEdit && <div style={{ textAlign: "center", padding: 20, color: "#bbb", fontSize: 13 }}>아직 팀이 배정되지 않았습니다</div>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: SCOREBOARD
// ═══════════════════════════════════════
function ScoreView({ S, save, canEdit }) {
  const sb = S.sb || { teams: [], rounds: [] };
  const { teams, rounds } = sb;
  const [nt, setNt] = useState("");
  const addT = () => { if (nt.trim() && !teams.includes(nt.trim())) { save({ sb: { ...sb, teams: [...teams, nt.trim()] } }); setNt(""); } };
  const totals = teams.map(t => ({ team: t, total: rounds.reduce((a, r) => a + (r.scores[t] || 0), 0) })).sort((a, b) => b.total - a.total);

  return (
    <>
      {canEdit ? <div style={s.card}>
        <div style={s.section}>팀 등록</div>
        <div style={{ display: "flex", gap: 6 }}><input style={{ ...s.input("auto"), flex: 1 }} placeholder="팀 이름" value={nt} onChange={e => setNt(e.target.value)} onKeyDown={e => e.key === "Enter" && addT()} /><button onClick={addT} style={s.btn("#7c3aed")}>추가</button></div>
        {teams.length > 0 && <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8 }}>{teams.map(t => <span key={t} style={{ ...s.tag("#ede9fe", "#7c3aed"), padding: "5px 10px", cursor: "pointer" }} onClick={() => save({ sb: { teams: teams.filter(x => x !== t), rounds: rounds.map(r => { const ns = { ...r.scores }; delete ns[t]; return { ...r, scores: ns }; }) } })}>{t} ×</span>)}</div>}
      </div> : !teams.length && <div style={{ textAlign: "center", padding: 20, color: "#bbb", fontSize: 13 }}>아직 점수판이 준비되지 않았습니다</div>}

      {teams.length >= 2 && totals.some(t => t.total > 0) || rounds.length > 0 ? <div style={{ ...s.card, background: "linear-gradient(135deg,#faf5ff,#ede9fe)" }}>
        <div style={s.section}>🏆 현재 순위</div>
        {totals.map((t, i) => <div key={t.team} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid rgba(0,0,0,0.05)" }}><span style={{ fontSize: 15 }}>{["🥇", "🥈", "🥉"][i] || `${i + 1}위`} {t.team}</span><span style={{ fontSize: 18, fontWeight: 800, color: "#7c3aed" }}>{t.total}점</span></div>)}
      </div> : null}

      {canEdit && teams.length >= 2 && <>
        <button onClick={() => { const sc = {}; teams.forEach(t => sc[t] = 0); save({ sb: { ...sb, rounds: [...rounds, { id: uid(), name: `라운드 ${rounds.length + 1}`, scores: sc }] } }); }} style={{ ...s.btn("#7c3aed"), width: "100%", borderRadius: 12, marginBottom: 10, padding: "12px" }}>+ 라운드 추가</button>
        {rounds.map(r => (
          <div key={r.id} style={s.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <input style={{ ...s.input("auto"), fontWeight: 700, flex: 1 }} value={r.name} onChange={e => save({ sb: { ...sb, rounds: rounds.map(x => x.id === r.id ? { ...x, name: e.target.value } : x) } })} />
              <button style={s.del} onClick={() => save({ sb: { ...sb, rounds: rounds.filter(x => x.id !== r.id) } })}>×</button>
            </div>
            {teams.map(t => <div key={t} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 0" }}><span style={{ fontSize: 13 }}>{t}</span><input style={{ ...s.input("80px"), textAlign: "center" }} type="number" inputMode="numeric" value={r.scores[t] || ""} onChange={e => save({ sb: { ...sb, rounds: rounds.map(x => x.id === r.id ? { ...x, scores: { ...x.scores, [t]: +e.target.value || 0 } } : x) } })} /></div>)}
          </div>
        ))}
      </>}

      {!canEdit && rounds.length > 0 && <div style={s.card}>
        <div style={s.section}>라운드별 점수</div>
        {rounds.map(r => <div key={r.id} style={{ marginBottom: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 2 }}>{r.name}</div>
          <div style={{ fontSize: 13, color: "#666" }}>{teams.map(t => `${t}: ${r.scores[t] || 0}점`).join(" · ")}</div>
        </div>)}
      </div>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: MINI GAMES
// ═══════════════════════════════════════
function MiniView({ S, save, canEdit, myName, act }) {
  const [game, setGame] = useState("updown");
  const ud = S.updown || { target: null, range: 100, over: false, winner: null, submissions: [] };
  const ch = S.chosung || { current: null, submissions: [] };
  const qz = S.quiz || { question: null, submissions: [], winner: null };
  const [guess, setGuess] = useState("");
  const [ans, setAns] = useState("");
  const [qAns, setQAns] = useState("");
  const [chLen, setChLen] = useState(2);
  const [qInput, setQInput] = useState("");
  const [err, setErr] = useState("");

  const doSubmit = async (g, v, clear) => {
    setErr("");
    try { await act.submit(g, v); clear(); } catch (e) { setErr(e.message); }
  };

  const GameTab = ({ k, l }) => <button onClick={() => { setGame(k); setErr(""); }} style={{ flex: 1, padding: "10px 4px", borderRadius: 10, border: game === k ? "2px solid #7c3aed" : "2px solid #e5e7eb", background: game === k ? "#ede9fe" : "#fff", color: game === k ? "#7c3aed" : "#888", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{l}</button>;

  return (
    <>
      <div style={{ display: "flex", gap: 5, marginBottom: 12 }}>
        <GameTab k="updown" l="🔢 업다운" /><GameTab k="chosung" l="🔤 초성" /><GameTab k="quiz" l="✍️ 자유출제" />
      </div>
      {err && <div style={{ padding: "8px 10px", background: "#fee2e2", borderRadius: 8, fontSize: 13, color: "#dc2626", marginBottom: 8 }}>{err}</div>}

      {/* ─── 업다운 ─── */}
      {game === "updown" && <div style={s.card}>
        <div style={s.section}>🔢 업다운 — {canEdit ? "호스트가 출제, 게스트가 도전" : "숫자를 맞춰보세요!"}</div>
        {canEdit && <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
          {[50, 100, 200].map(r => <button key={r} onClick={() => save({ updown: { target: Math.floor(Math.random() * r) + 1, range: r, over: false, winner: null, submissions: [] } })} style={{ ...s.btn("#f5f5f5", "#555"), flex: 1, borderRadius: 10, fontSize: 13 }}>1~{r} 출제</button>)}
        </div>}
        {(() => { const running = canEdit ? (!!ud.target && !ud.over) : !!ud.active; return (running || ud.over || ud.submissions.length > 0) ? <>
          {(running || ud.over) ? <>
            <div style={{ textAlign: "center", fontSize: 14, color: "#666", marginBottom: 8 }}>범위: 1 ~ {ud.range}</div>
            {canEdit && ud.target && !ud.over && <div style={{ textAlign: "center", fontSize: 12, color: "#94a3b8", marginBottom: 8 }}>🤫 정답: <b>{ud.target}</b> (호스트에게만 보임)</div>}
            {ud.over ? <div style={{ textAlign: "center", padding: 12 }}>
              <div style={{ fontSize: 32, fontWeight: 800, color: "#059669" }}>🎉 정답!</div>
              <div style={{ fontSize: 15, marginTop: 4 }}><b>{ud.winner}</b>님이 맞혔습니다{canEdit && ud.submissions[0] ? ` (정답: ${ud.submissions[0].val})` : ""}</div>
            </div> : !canEdit && <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
              <input style={{ ...s.input("auto"), flex: 1, textAlign: "center", fontSize: 18 }} type="number" inputMode="numeric" placeholder="?" value={guess} onChange={e => setGuess(e.target.value)} onKeyDown={e => e.key === "Enter" && doSubmit("updown", guess, () => setGuess(""))} />
              <button onClick={() => doSubmit("updown", guess, () => setGuess(""))} style={s.btn("#7c3aed")}>제출</button>
            </div>}
          </> : <div style={{ textAlign: "center", padding: 12, color: "#bbb", fontSize: 13 }}>{canEdit ? "범위를 골라 출제하세요" : "호스트가 문제를 내면 시작됩니다"}</div>}
          {ud.submissions.length > 0 && <div style={{ marginTop: 8 }}>
            <div style={{ fontSize: 12, color: "#999", marginBottom: 4 }}>제출 기록 ({ud.submissions.length}회)</div>
            {ud.submissions.map((g, i) => <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "4px 8px", fontSize: 13, background: g.dir.startsWith("정답") ? "#f0fdf4" : "#fafafa", borderRadius: 6, marginBottom: 2 }}>
              <span><b>{g.name}</b>: {g.val}</span><span style={{ fontWeight: 600, color: g.dir.startsWith("정답") ? "#059669" : "#7c3aed" }}>{g.dir}</span>
            </div>)}
          </div>}
        </> : <div style={{ textAlign: "center", padding: 12, color: "#bbb", fontSize: 13 }}>{canEdit ? "범위를 골라 출제하세요" : "호스트가 문제를 내면 시작됩니다"}</div>; })()}
      </div>}

      {/* ─── 초성 ─── */}
      {game === "chosung" && <div style={s.card}>
        <div style={s.section}>🔤 초성 게임 — {canEdit ? "초성을 뽑고 정답 판정" : "초성으로 단어를 맞춰 제출!"}</div>
        {canEdit && <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 }}>
          <div style={s.label}>글자 수</div><Stepper value={chLen} onChange={setChLen} min={2} />
          <button onClick={() => save({ chosung: { current: Array.from({ length: chLen }, () => CHOSUNG_SET[Math.floor(Math.random() * CHOSUNG_SET.length)]).join(" "), submissions: [] } })} style={{ ...s.btn("#7c3aed"), flex: 1, borderRadius: 10 }}>{ch.current ? "🔄 다음 초성" : "🎲 초성 뽑기"}</button>
        </div>}
        <div style={{ textAlign: "center", padding: "16px 0" }}>
          <div style={{ fontSize: 44, fontWeight: 800, color: "#7c3aed", letterSpacing: 10, minHeight: 54 }}>{ch.current || "?"}</div>
        </div>
        {ch.current && !canEdit && <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
          <input style={{ ...s.input("auto"), flex: 1 }} placeholder="정답 입력" value={ans} onChange={e => setAns(e.target.value)} onKeyDown={e => e.key === "Enter" && doSubmit("chosung", ans, () => setAns(""))} />
          <button onClick={() => doSubmit("chosung", ans, () => setAns(""))} style={s.btn("#7c3aed")}>제출</button>
        </div>}
        {ch.submissions.length > 0 && <div style={{ marginTop: 4 }}>
          <div style={{ fontSize: 12, color: "#999", marginBottom: 4 }}>제출 순서 (빠른 순)</div>
          {[...ch.submissions].sort((a, b) => a.t - b.t).map((x, i) => <div key={x.name} style={{ display: "flex", justifyContent: "space-between", padding: "5px 8px", fontSize: 13, background: i === 0 ? "#fef9c3" : "#fafafa", borderRadius: 6, marginBottom: 2 }}>
            <span>{i + 1}. <b>{x.name}</b>{canEdit || x.name === myName ? ` — ${x.answer}` : ""}</span><span style={{ color: "#bbb", fontSize: 11 }}>{timeStr(x.t)}</span>
          </div>)}
          {!canEdit && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>💡 다른 사람의 답 내용은 호스트만 볼 수 있어요</div>}
        </div>}
      </div>}

      {/* ─── 자유 출제 ─── */}
      {game === "quiz" && <div style={s.card}>
        <div style={s.section}>✍️ 자유 출제 — {canEdit ? "문제를 내고 정답자 선택" : "답을 제출하세요!"}</div>
        {canEdit && <div style={{ marginBottom: 10 }}>
          <textarea style={{ ...s.input("100%"), minHeight: 56, resize: "vertical", fontFamily: "inherit" }} placeholder="문제를 입력하세요 (예: 우리 학과 조교님 성함은?)" value={qInput} onChange={e => setQInput(e.target.value)} />
          <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
            <button onClick={() => { if (qInput.trim()) { save({ quiz: { question: qInput.trim(), submissions: [], winner: null } }); setQInput(""); } }} style={{ ...s.btn("#7c3aed"), flex: 1, borderRadius: 10 }}>📢 출제</button>
            {qz.question && <button onClick={() => save({ quiz: { question: null, submissions: [], winner: null } })} style={{ ...s.btn("#f5f5f5", "#666"), flex: 1, borderRadius: 10 }}>문제 내리기</button>}
          </div>
        </div>}
        {qz.question ? <>
          <div style={{ padding: "14px", background: "#faf5ff", borderRadius: 12, fontSize: 17, fontWeight: 700, color: "#6d28d9", textAlign: "center", marginBottom: 10, lineHeight: 1.5 }}>Q. {qz.question}</div>
          {qz.winner && <div style={{ padding: "10px", background: "#fef9c3", borderRadius: 10, textAlign: "center", fontSize: 15, marginBottom: 8 }}>🏅 정답자: <b>{qz.winner}</b></div>}
          {!canEdit && !qz.winner && <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
            <input style={{ ...s.input("auto"), flex: 1 }} placeholder="답 입력" value={qAns} onChange={e => setQAns(e.target.value)} onKeyDown={e => e.key === "Enter" && doSubmit("quiz", qAns, () => setQAns(""))} />
            <button onClick={() => doSubmit("quiz", qAns, () => setQAns(""))} style={s.btn("#7c3aed")}>제출</button>
          </div>}
          {qz.submissions.length > 0 && <div>
            <div style={{ fontSize: 12, color: "#999", marginBottom: 4 }}>제출 답안 (빠른 순){canEdit && " — 탭하여 정답자 선정"}</div>
            {[...qz.submissions].sort((a, b) => a.t - b.t).map((x, i) => (
              <div key={x.name} onClick={() => canEdit && save({ quiz: { ...qz, winner: x.name } })} style={{ display: "flex", justifyContent: "space-between", padding: "7px 10px", fontSize: 13, background: qz.winner === x.name ? "#fef9c3" : "#fafafa", borderRadius: 8, marginBottom: 3, cursor: canEdit ? "pointer" : "default", border: qz.winner === x.name ? "1.5px solid #eab308" : "1.5px solid transparent" }}>
                <span>{i + 1}. <b>{x.name}</b>{canEdit || x.name === myName ? ` — ${x.answer}` : ""}</span><span style={{ color: "#bbb", fontSize: 11 }}>{timeStr(x.t)}</span>
              </div>
            ))}
            {!canEdit && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>💡 다른 사람의 답 내용은 호스트만 볼 수 있어요</div>}
          </div>}
        </> : <div style={{ textAlign: "center", padding: 12, color: "#bbb", fontSize: 13 }}>{canEdit ? "위에서 문제를 출제하세요" : "호스트가 문제를 내면 여기 표시됩니다"}</div>}
      </div>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: SETTLEMENT
// ═══════════════════════════════════════
function SettleView({ S, save, canEdit, myName }) {
  const members = (S.members || []).map(m => m.name);
  const settle = S.settle || { items: [] };
  const items = settle.items;
  const updI = (id, f, v) => save({ settle: { items: items.map(i => i.id === id ? { ...i, [f]: v } : i) } });
  const totalA = items.reduce((a, i) => a + (i.amount || 0), 0);
  const perP = members.map(n => ({ name: n, share: Math.round(items.reduce((a, i) => { const exc = i.exclude || []; const p = members.length - exc.length; return p <= 0 || exc.includes(n) ? a : a + (i.amount || 0) / p; }, 0)) }));

  return (
    <>
      {!canEdit && <RoNote />}
      <div style={s.card}>
        <div style={s.section}>👤 정산 대상: 등록된 참가자 {members.length}명</div>
        {members.length === 0 ? <div style={{ fontSize: 13, color: "#bbb", textAlign: "center" }}>⚙️ 관리 → 참가자 계정에서 등록하세요</div> :
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{members.map(n => <span key={n} style={{ ...s.tag(n === myName ? "#0d9488" : "#ccfbf1", n === myName ? "#fff" : "#0d9488"), padding: "5px 10px" }}>{n}</span>)}</div>}
      </div>
      {canEdit && <button onClick={() => save({ settle: { items: [{ id: uid(), name: "", amount: 0, exclude: [] }, ...items] } })} style={{ ...s.btn("#0d9488"), width: "100%", borderRadius: 12, marginBottom: 10, padding: "12px", fontSize: 15 }}>+ 지출 항목 추가</button>}

      {items.map(item => canEdit ? (
        <div key={item.id} style={s.card}>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 8 }}><input style={{ ...s.input("auto"), flex: 1 }} placeholder="항목명 (예: 고기값)" value={item.name} onChange={e => updI(item.id, "name", e.target.value)} /><button style={s.del} onClick={() => save({ settle: { items: items.filter(i => i.id !== item.id) } })}>×</button></div>
          <div style={{ marginBottom: 6 }}><div style={{ fontSize: 11, color: "#999" }}>금액 (₩)</div><input style={{ ...s.input("100%"), textAlign: "right" }} type="number" inputMode="numeric" value={item.amount || ""} onChange={e => updI(item.id, "amount", +e.target.value || 0)} /></div>
          {members.length > 0 && <div><div style={{ fontSize: 11, color: "#999", marginBottom: 4 }}>제외할 사람 (탭하여 토글)</div><div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>{members.map(n => { const ex = (item.exclude || []).includes(n); return <button key={n} onClick={() => updI(item.id, "exclude", ex ? item.exclude.filter(x => x !== n) : [...(item.exclude || []), n])} style={{ ...s.smBtn(ex ? "#fee2e2" : "#f3f4f6", ex ? "#dc2626" : "#333"), fontSize: 12, textDecoration: ex ? "line-through" : "none" }}>{n}</button>; })}</div></div>}
        </div>
      ) : (
        <div key={item.id} style={{ ...s.card, padding: "10px 14px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14 }}><b>{item.name || "(무명)"}</b><span style={{ fontWeight: 700, color: "#0d9488" }}>₩{fmt(item.amount || 0)}</span></div>
          {(item.exclude || []).length > 0 && <div style={{ fontSize: 12, color: "#999", marginTop: 2 }}>제외: {item.exclude.join(", ")}</div>}
        </div>
      ))}

      {members.length > 0 && items.length > 0 && <div style={{ ...s.card, background: "linear-gradient(135deg,#f0fdf9,#ccfbf1)", border: "1.5px solid #5eead4" }}>
        <div style={s.section}>💸 1인당 정산</div>
        <div style={{ fontSize: 12, color: "#666", marginBottom: 8 }}>총 지출: ₩{fmt(totalA)}</div>
        {perP.map(p => <div key={p.name} style={{ display: "flex", justifyContent: "space-between", padding: "6px 8px", borderRadius: 8, background: p.name === myName ? "rgba(13,148,136,0.15)" : "transparent", borderBottom: "1px solid rgba(0,0,0,0.05)" }}>
          <span style={{ fontSize: 14, fontWeight: p.name === myName ? 700 : 400 }}>{p.name}{p.name === myName && " ⭐"}</span><span style={{ fontSize: 16, fontWeight: 700, color: "#0d9488" }}>₩{fmt(p.share)}</span>
        </div>)}
      </div>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: SCHEDULE
// ═══════════════════════════════════════
function SchedView({ S, save, canEdit }) {
  const sched = S.schedule || [];
  const sorted = [...sched].sort((a, b) => (a.time || "").localeCompare(b.time || ""));
  return (
    <>
      {!canEdit && <RoNote />}
      {canEdit && <button onClick={() => save({ schedule: [{ id: uid(), time: "", event: "" }, ...sched] })} style={{ ...s.btn("#0d9488"), width: "100%", borderRadius: 12, marginBottom: 10, padding: "12px", fontSize: 15 }}>+ 일정 추가</button>}
      {sorted.length === 0 ? <div style={{ textAlign: "center", padding: 30, color: "#bbb" }}>📅 등록된 일정이 없습니다</div> :
        canEdit ? sorted.map(item => (
          <div key={item.id} style={{ ...s.card, display: "flex", gap: 10, alignItems: "center", padding: "10px 14px" }}>
            <input style={{ ...s.input("90px"), textAlign: "center", fontWeight: 700 }} type="time" value={item.time} onChange={e => save({ schedule: sched.map(x => x.id === item.id ? { ...x, time: e.target.value } : x) })} />
            <input style={{ ...s.input("auto"), flex: 1 }} placeholder="일정 내용" value={item.event} onChange={e => save({ schedule: sched.map(x => x.id === item.id ? { ...x, event: e.target.value } : x) })} />
            <button style={s.del} onClick={() => save({ schedule: sched.filter(x => x.id !== item.id) })}>×</button>
          </div>
        )) : (
          <div style={s.card}>
            {sorted.map((item, i) => (
              <div key={item.id} style={{ display: "flex", gap: 12, alignItems: "center", padding: "10px 0", borderBottom: i < sorted.length - 1 ? "1px solid #f0f0f0" : "none" }}>
                <div style={{ fontSize: 15, fontWeight: 800, color: "#0d9488", minWidth: 52 }}>{item.time || "--:--"}</div>
                <div style={{ fontSize: 14 }}>{item.event || "(내용 없음)"}</div>
              </div>
            ))}
          </div>
        )}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: NOTICES (공지)
// ═══════════════════════════════════════
function NoticeView({ S, save, canEdit }) {
  const notices = S.notices || [];
  const [text, setText] = useState("");
  const [img, setImg] = useState(null);
  const [busy, setBusy] = useState(false);
  const fileRef = useRef(null);
  // 이 화면에 들어온 시점의 "마지막으로 본 시각"을 캡처 → 그보다 새 공지에 New! 표시
  const [seenAt] = useState(() => Number(localStorage.getItem(NOTICE_SEEN_KEY) || 0));
  useEffect(() => { localStorage.setItem(NOTICE_SEEN_KEY, String(Date.now())); }, [notices.length]);

  const attach = async (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    setBusy(true);
    try { setImg(await resizeImage(f)); } catch (err) { alert(err.message); }
    setBusy(false); e.target.value = "";
  };
  const post = () => {
    if (!text.trim() && !img) { alert("내용 또는 이미지를 입력하세요."); return; }
    save({ notices: [{ id: uid(), text: text.trim(), img, t: Date.now() }, ...notices] });
    setText(""); setImg(null);
  };

  return (
    <>
      {canEdit && <div style={s.card}>
        <div style={s.section}>📢 공지 작성</div>
        <textarea style={{ ...s.input("100%"), minHeight: 64, resize: "vertical", fontFamily: "inherit" }} placeholder="공지 내용 (예: 19시 바베큐장 집합!)" value={text} onChange={e => setText(e.target.value)} />
        {img && <div style={{ position: "relative", marginTop: 8 }}>
          <img src={img} alt="첨부" style={{ width: "100%", borderRadius: 10, display: "block" }} />
          <button onClick={() => setImg(null)} style={{ position: "absolute", top: 6, right: 6, background: "rgba(0,0,0,0.55)", color: "#fff", border: "none", borderRadius: 8, padding: "4px 10px", fontSize: 12, cursor: "pointer" }}>이미지 제거</button>
        </div>}
        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <button onClick={() => fileRef.current?.click()} disabled={busy} style={{ ...s.btn("#f5f5f5", "#555"), flex: 1, borderRadius: 10, fontSize: 13 }}>{busy ? "처리 중..." : img ? "🖼️ 이미지 변경" : "🖼️ 이미지 첨부"}</button>
          <button onClick={post} style={{ ...s.btn("#e11d48"), flex: 2, borderRadius: 10 }}>공지 올리기</button>
          <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={attach} />
        </div>
        <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 6 }}>💡 이미지는 자동으로 축소 저장됩니다 (최대 800px)</div>
      </div>}

      {notices.length === 0 ? <div style={{ textAlign: "center", padding: 30, color: "#bbb" }}><div style={{ fontSize: 36 }}>📢</div><div style={{ fontSize: 13, marginTop: 6 }}>등록된 공지가 없습니다</div></div> :
        notices.map(n => (
          <div key={n.id} style={{ ...s.card, borderLeft: n.t > seenAt ? "4px solid #e11d48" : "4px solid #e5e7eb" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                {n.t > seenAt && <span style={{ ...s.tag("#e11d48", "#fff"), fontSize: 10 }}>New!</span>}
                <span style={{ fontSize: 11, color: "#999" }}>{dateStr(n.t)}</span>
              </div>
              {canEdit && <button style={s.del} onClick={() => window.confirm("이 공지를 삭제할까요?") && save({ notices: notices.filter(x => x.id !== n.id) })}>×</button>}
            </div>
            {n.text && <div style={{ fontSize: 14, lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{n.text}</div>}
            {n.img && <img src={n.img} alt="공지 이미지" style={{ width: "100%", borderRadius: 10, marginTop: n.text ? 8 : 0, display: "block" }} />}
          </div>
        ))}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: ICEBREAKER
// ═══════════════════════════════════════
function IceView({ S, save, canEdit }) {
  const ice = S.ice || { usedQs: [], customQs: [], current: null };
  const allQ = [...ICE_QUESTIONS, ...(ice.customQs || [])];
  const remaining = allQ.filter((_, i) => !ice.usedQs.includes(i));
  const [nq, setNq] = useState("");
  const draw = () => {
    if (!remaining.length) return;
    const pool = allQ.map((q, i) => ({ q, i })).filter(x => !ice.usedQs.includes(x.i));
    const pick = pool[Math.floor(Math.random() * pool.length)];
    save({ ice: { ...ice, usedQs: [...ice.usedQs, pick.i], current: pick.q } });
  };

  return (
    <>
      <div style={{ ...s.card, textAlign: "center", padding: "28px 20px", minHeight: 110, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
        {ice.current ? <div style={{ fontSize: 20, fontWeight: 700, color: "#e11d48", lineHeight: 1.5 }}>{ice.current}</div> : <div style={{ fontSize: 16, color: "#bbb" }}>{canEdit ? "질문을 뽑아보세요!" : "호스트가 질문을 뽑으면 여기 표시됩니다"}</div>}
      </div>
      <div style={{ textAlign: "center", fontSize: 12, color: "#999", marginBottom: 8 }}>남은 질문: {remaining.length} / {allQ.length} (한번 나온 질문은 반복되지 않아요)</div>
      {canEdit && <>
        <button onClick={draw} disabled={!remaining.length} style={{ ...s.btn(!remaining.length ? "#ccc" : "#e11d48"), width: "100%", borderRadius: 12, padding: "14px", fontSize: 16, marginBottom: 8 }}>❓ 질문 뽑기</button>
        {ice.usedQs.length > 0 && <button onClick={() => save({ ice: { ...ice, usedQs: [], current: null } })} style={{ ...s.btn("#f5f5f5", "#999"), width: "100%", borderRadius: 12, fontSize: 13, marginBottom: 8 }}>사용한 질문 초기화 (전부 다시 나옴)</button>}
        <div style={s.card}>
          <div style={s.section}>✏️ 질문 직접 추가</div>
          <div style={{ display: "flex", gap: 6 }}><input style={{ ...s.input("auto"), flex: 1 }} placeholder="질문을 입력하세요" value={nq} onChange={e => setNq(e.target.value)} onKeyDown={e => e.key === "Enter" && nq.trim() && (save({ ice: { ...ice, customQs: [...ice.customQs, nq.trim()] } }), setNq(""))} /><button onClick={() => { if (nq.trim()) { save({ ice: { ...ice, customQs: [...ice.customQs, nq.trim()] } }); setNq(""); } }} style={s.btn("#e11d48")}>추가</button></div>
          {ice.customQs.length > 0 && <div style={{ marginTop: 8, fontSize: 12, color: "#999" }}>추가된 질문 {ice.customQs.length}개</div>}
        </div>
      </>}
    </>
  );
}

// ═══════════════════════════════════════
// FEATURE: VOTING
// ═══════════════════════════════════════
function VoteView({ S, save, canEdit, myName, act }) {
  const votes = S.votes || [];
  const [topic, setTopic] = useState(""); const [opts, setOpts] = useState("");
  const [err, setErr] = useState("");
  const create = () => {
    const o = opts.split(/[,\n]/).map(x => x.trim()).filter(Boolean);
    if (!topic.trim() || o.length < 2) { alert("주제와 2개 이상의 선택지를 입력하세요"); return; }
    save({ votes: [{ id: uid(), topic: topic.trim(), options: o.map(x => ({ id: uid(), text: x })), voters: {}, closed: false }, ...votes] });
    setTopic(""); setOpts("");
  };
  const castVote = async (vId, oId) => { setErr(""); try { await act.vote(vId, oId); } catch (e) { setErr(e.message); } };

  return (
    <>
      {canEdit && <div style={s.card}>
        <div style={s.section}>🗳️ 새 투표</div>
        <input style={{ ...s.input("100%"), marginBottom: 8 }} placeholder='주제 (예: "저녁 뭐 먹을까?")' value={topic} onChange={e => setTopic(e.target.value)} />
        <textarea style={{ ...s.input("100%"), minHeight: 60, resize: "vertical", fontFamily: "inherit" }} placeholder={"선택지 (쉼표 또는 줄바꿈)\n예: 삼겹살, 치킨, 피자"} value={opts} onChange={e => setOpts(e.target.value)} />
        <button onClick={create} style={{ ...s.btn("#e11d48"), width: "100%", marginTop: 8, borderRadius: 12, padding: "12px" }}>투표 생성</button>
      </div>}
      {err && <div style={{ padding: "8px 10px", background: "#fee2e2", borderRadius: 8, fontSize: 13, color: "#dc2626", marginBottom: 8 }}>{err}</div>}
      {votes.length === 0 && !canEdit && <div style={{ textAlign: "center", padding: 30, color: "#bbb", fontSize: 13 }}>진행 중인 투표가 없습니다</div>}

      {votes.map(v => {
        const voters = v.voters || {};
        const counts = {}; Object.values(voters).forEach(oId => counts[oId] = (counts[oId] || 0) + 1);
        const tot = Object.keys(voters).length;
        const maxC = Math.max(...v.options.map(o => counts[o.id] || 0), 1);
        const myVote = voters[myName];
        return (
          <div key={v.id} style={s.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ fontSize: 15, fontWeight: 700 }}>{v.topic}</div>
              {canEdit && <div style={{ display: "flex", gap: 4 }}>
                <button onClick={() => save({ votes: votes.map(x => x.id === v.id ? { ...x, closed: !x.closed } : x) })} style={s.smBtn(v.closed ? "#dcfce7" : "#fee2e2", v.closed ? "#16a34a" : "#dc2626")}>{v.closed ? "재개" : "마감"}</button>
                <button style={s.del} onClick={() => save({ votes: votes.filter(x => x.id !== v.id) })}>×</button>
              </div>}
            </div>
            {v.options.map(o => {
              const c = counts[o.id] || 0; const isMine = myVote === o.id;
              return (
                <div key={o.id} style={{ marginBottom: 8 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 2 }}>
                    <button onClick={() => !v.closed && castVote(v.id, o.id)} disabled={v.closed} style={{ background: "none", border: "none", fontSize: 14, cursor: v.closed ? "default" : "pointer", color: isMine ? "#e11d48" : "#333", fontWeight: isMine ? 700 : 400, textAlign: "left", padding: 0 }}>
                      {!v.closed && "👆 "}{o.text}{isMine && " ✔️ 내 선택"}
                    </button>
                    <span style={{ fontSize: 13, fontWeight: 600, color: "#e11d48" }}>{c}표{tot > 0 ? ` (${Math.round(c / tot * 100)}%)` : ""}</span>
                  </div>
                  <div style={{ background: "#f3f4f6", borderRadius: 4, height: 8, overflow: "hidden" }}><div style={{ width: `${(c / maxC) * 100}%`, height: "100%", background: isMine ? "#be123c" : "#e11d48", borderRadius: 4, transition: "width 0.3s" }} /></div>
                </div>
              );
            })}
            <div style={{ fontSize: 12, color: "#999", marginTop: 6 }}>
              {tot}명 참여{v.closed && " · 마감됨"}{!v.closed && !myVote && " · 아직 투표하지 않았어요 (마감 전 변경 가능)"}
            </div>
          </div>
        );
      })}
    </>
  );
}

// ═══════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════
export default function App() {
  const [auth, setAuth] = useState(() => { try { return JSON.parse(localStorage.getItem(AUTH_KEY)) || null; } catch { return null; } });
  const [S, setS] = useState(null);
  const [cat, setCat] = useState(null);
  const [sub, setSub] = useState(null);
  const [saved, setSaved] = useState(false);
  const [fatal, setFatal] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const [showHelp, setShowHelp] = useState(false);
  const [dark, toggleDark] = useDark();
  const timer = useRef(null);
  const pending = useRef({});
  const lastEdit = useRef(0);

  if (auth) api.token = auth.token;
  const role = auth?.role;
  const myName = auth?.name;
  const canEdit = role === "host" || role === "admin";

  const refresh = useCallback(async (manual) => {
    if (!api.token) return;
    if (manual) setSyncing(true);
    try { const d = await api.req("state"); setS(d); setFatal(null); }
    catch (e) {
      if (e.code === "KV_NOT_BOUND") setFatal(e.message);
      else if (e.message.includes("로그인")) { localStorage.removeItem(AUTH_KEY); setAuth(null); }
    }
    if (manual) setSyncing(false);
  }, []);

  // 최초 로드 + 폴링 (10초, 읽기 전용 — 편집 직후 8초간은 폴링 스킵)
  useEffect(() => {
    if (!auth) return;
    refresh();
    const iv = setInterval(() => { if (Date.now() - lastEdit.current > 8000) refresh(); }, 10000);
    return () => clearInterval(iv);
  }, [auth, refresh]);

  const flush = useCallback(async () => {
    const p = pending.current; pending.current = {};
    if (!Object.keys(p).length) return;
    try {
      const d = await api.req("update", "POST", { patch: p });
      setS(d); setSaved(true); setTimeout(() => setSaved(false), 1200);
    } catch (e) { alert("저장 실패: " + e.message); refresh(); }
  }, [refresh]);

  const save = useCallback((patch) => {
    lastEdit.current = Date.now();
    setS(prev => ({ ...prev, ...patch }));
    pending.current = { ...pending.current, ...patch };
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(flush, 600);
  }, [flush]);
  save.replace = (newState) => { pending.current = {}; setS(newState); };

  const act = {
    vote: async (voteId, optionId) => { lastEdit.current = Date.now(); const d = await api.req("vote", "POST", { voteId, optionId }); setS(d); },
    submit: async (game, value) => { lastEdit.current = Date.now(); const d = await api.req("submit", "POST", { game, value }); setS(d); },
  };

  const logout = () => { localStorage.removeItem(AUTH_KEY); setAuth(null); setS(null); setCat(null); setSub(null); api.token = null; };
  const onLogin = (d) => {
    const a = { token: d.token, role: d.role, name: d.name };
    localStorage.setItem(AUTH_KEY, JSON.stringify(a));
    api.token = d.token; setAuth(a); if (d.state) setS(d.state);
    if (d.firstSetup) setTimeout(() => alert("호스트 비밀번호가 설정되었습니다! 이 비밀번호를 기억해주세요."), 100);
  };

  if (!auth) return <LoginView onLogin={onLogin} dark={dark} toggleDark={toggleDark} />;

  if (fatal) return (
    <div style={{ ...s.page, justifyContent: "center", padding: 20 }}>
      <div style={{ ...s.card, border: "2px solid #fca5a5" }}>
        <div style={{ fontSize: 32, textAlign: "center" }}>⚠️</div>
        <div style={{ fontSize: 15, fontWeight: 700, textAlign: "center", margin: "8px 0" }}>서버 설정 필요</div>
        <div style={{ fontSize: 13, color: "#666", lineHeight: 1.6 }}>{fatal}</div>
        <button onClick={() => refresh(true)} style={{ ...s.btn("#7c3aed"), width: "100%", marginTop: 12 }}>다시 시도</button>
      </div>
    </div>
  );

  if (!S) return <div style={{ ...s.page, justifyContent: "center", alignItems: "center" }}><div style={{ color: "#999" }}>불러오는 중...</div></div>;

  const visibleCats = CATEGORIES.filter(c => !c.staffOnly || canEdit);
  const subs = (SUBS[cat] || []).filter(t => !t.adminOnly || role === "admin");
  const catInfo = CATEGORIES.find(c => c.key === cat);
  const goCat = (c) => { setCat(c); setSub(SUBS[c].filter(t => !t.adminOnly || role === "admin")[0].key); };

  const ROLE_BADGE = { admin: ["⚙️ Admin", "#475569"], host: ["👑 호스트", "#d97706"], guest: [`🙋 ${myName}`, "#0d9488"] };
  const [badgeTxt, badgeCol] = ROLE_BADGE[role];

  const renderFeature = () => {
    const p = { S, save, canEdit, myName, role, act };
    switch (sub) {
      case "cart": return <CartView {...p} />;
      case "alc": return <AlcView {...p} />;
      case "roulette": return <RouletteView {...p} />;
      case "teams": return <TeamsView {...p} />;
      case "score": return <ScoreView {...p} />;
      case "mini": return <MiniView {...p} />;
      case "settle": return <SettleView {...p} />;
      case "sched": return <SchedView {...p} />;
      case "ice": return <IceView {...p} />;
      case "notice": return <NoticeView {...p} />;
      case "vote": return <VoteView {...p} />;
      case "members": return <MembersView {...p} />;
      case "system": return <SystemView S={S} save={save} />;
      default: return null;
    }
  };

  return (
    <div style={s.page}>
      <div style={{ background: "linear-gradient(135deg,#1e293b,#334155)", padding: "14px 16px", color: "#fff" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {cat && <button onClick={() => { setCat(null); setSub(null); }} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", borderRadius: 8, padding: "6px 10px", fontSize: 14, cursor: "pointer" }}>←</button>}
            <div style={{ fontSize: 17, fontWeight: 800 }}>{cat ? `${catInfo.icon} ${catInfo.label}` : "🏕️ MT 도우미"}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
            {saved && <span style={{ fontSize: 11, color: "#86efac" }}>저장됨 ✓</span>}
            <button onClick={() => setShowHelp(true)} style={{ background: "rgba(255,255,255,0.12)", border: "none", borderRadius: 8, padding: "5px 8px", fontSize: 14, cursor: "pointer" }} title="사용법">❓</button>
            <button onClick={toggleDark} style={{ background: "rgba(255,255,255,0.12)", border: "none", borderRadius: 8, padding: "5px 8px", fontSize: 14, cursor: "pointer" }} title="다크 모드">{dark ? "☀️" : "🌙"}</button>
            <button onClick={() => refresh(true)} style={{ background: "rgba(255,255,255,0.12)", border: "none", borderRadius: 8, padding: "5px 8px", fontSize: 14, cursor: "pointer", color: "#fff", opacity: syncing ? 0.5 : 1 }} title="새로고침">{syncing ? "⏳" : "🔄"}</button>
            {cat && visibleCats.filter(c => c.key !== cat).map(c => (
              <button key={c.key} onClick={() => goCat(c.key)} style={{ background: "rgba(255,255,255,0.12)", border: "none", borderRadius: 8, padding: "5px 8px", fontSize: 14, cursor: "pointer" }} title={c.label}>{c.icon}</button>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
          <span style={{ ...s.tag(badgeCol, "#fff"), padding: "4px 10px", fontSize: 12 }}>{badgeTxt}</span>
          <button onClick={logout} style={{ background: "none", border: "none", color: "rgba(255,255,255,0.5)", fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>로그아웃</button>
        </div>
      </div>

      {!cat && <div style={{ padding: "16px 14px", flex: 1 }}>
        {(S.notices || []).length > 0 && (() => {
          const latest = S.notices[0];
          const unread = latest.t > Number(localStorage.getItem(NOTICE_SEEN_KEY) || 0);
          return (
            <button onClick={() => { goCat("fun"); setSub("notice"); }} style={{ display: "block", width: "100%", padding: "12px 14px", marginBottom: 12, borderRadius: 14, border: unread ? "1.5px solid #fda4af" : "1.5px solid #e5e7eb", background: unread ? "#fff1f2" : "#fff", cursor: "pointer", textAlign: "left" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                <span style={{ fontSize: 13, fontWeight: 700, color: "#e11d48" }}>📢 공지</span>
                {unread && <span style={{ ...s.tag("#e11d48", "#fff"), fontSize: 10 }}>New!</span>}
                <span style={{ fontSize: 11, color: "#999", marginLeft: "auto" }}>{dateStr(latest.t)}</span>
              </div>
              <div style={{ fontSize: 13, color: "#444", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{latest.text || "🖼️ 이미지 공지"}</div>
            </button>
          );
        })()}
        {visibleCats.map(c => (
          <button key={c.key} onClick={() => goCat(c.key)} style={{ display: "flex", alignItems: "center", gap: 14, width: "100%", padding: "20px 18px", marginBottom: 12, borderRadius: 16, border: "none", background: c.bg, cursor: "pointer", textAlign: "left", boxShadow: "0 2px 8px rgba(0,0,0,0.06)" }}>
            <div style={{ fontSize: 36 }}>{c.icon}</div>
            <div><div style={{ fontSize: 17, fontWeight: 700, color: c.color }}>{c.label}</div><div style={{ fontSize: 12, color: "#666", marginTop: 2 }}>{c.desc}</div></div>
          </button>
        ))}
        <div style={{ textAlign: "center", fontSize: 11, color: "#bbb", marginTop: 10 }}>10초마다 자동으로 최신 상태를 불러옵니다 · 🔄 버튼으로 즉시 갱신</div>
      </div>}

      {cat && <>
        <div style={{ display: "flex", gap: 0, background: "#fff", borderBottom: "2px solid #e5e7eb", overflowX: "auto", position: "sticky", top: 0, zIndex: 20 }}>
          {subs.map(t => (
            <button key={t.key} onClick={() => setSub(t.key)} style={{ flex: "0 0 auto", padding: "11px 14px", fontSize: 13, whiteSpace: "nowrap", fontWeight: sub === t.key ? 700 : 500, color: sub === t.key ? catInfo.color : "#888", background: sub === t.key ? `${catInfo.color}08` : "#fff", borderBottom: sub === t.key ? `3px solid ${catInfo.color}` : "3px solid transparent", border: "none", cursor: "pointer" }}>{t.label}</button>
          ))}
        </div>
        <div style={{ flex: 1, padding: "12px 14px 24px", overflowY: "auto" }}>{renderFeature()}</div>
      </>}

      {showHelp && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", zIndex: 50, display: "flex", justifyContent: "center", alignItems: "flex-start", overflowY: "auto", padding: "24px 0" }} onClick={() => setShowHelp(false)}>
          <div style={{ background: "#f3f4f6", borderRadius: 16, width: "100%", maxWidth: 480, margin: "0 14px", padding: 14, boxShadow: "0 8px 40px rgba(0,0,0,0.3)" }} onClick={e => e.stopPropagation()}>
            <HelpView roleKey={role} onClose={() => setShowHelp(false)} embedded />
          </div>
        </div>
      )}
    </div>
  );
}
