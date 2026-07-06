// ═══════════════════════════════════════════════════════
// MT 도우미 API — Cloudflare Pages Functions
// 위치: functions/api/[[path]].js
// 필요 설정:
//   1) Pages > Settings > Functions > KV namespace bindings
//      → Variable name: MT_KV (KV 네임스페이스 연결)
//   2) Pages > Settings > Environment variables
//      → ADMIN_PASSWORD = 관리자 비밀번호 (미설정 시 기본값 "admin")
// ═══════════════════════════════════════════════════════

const KEY = "room:main";
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type,Authorization",
};
const j = (d, s = 200) =>
  new Response(JSON.stringify(d), { status: s, headers: { "Content-Type": "application/json", ...CORS } });

// 호스트/Admin이 patch로 수정할 수 있는 최상위 키 (config는 제외 — 별도 경로로만)
const PATCH_KEYS = [
  "members", "notices", "cart", "alc", "schedule", "settle", "sb",
  "roulette", "teams", "updown", "chosung", "quiz", "ice", "votes",
];

const initialGame = () => ({
  members: [],
  notices: [],
  cart: { mart: [], online: [], donated: [], pre: [] },
  alc: { people: 5, sojuPer: 2, drinks: [] },
  schedule: [],
  settle: { items: [] },
  sb: { teams: [], rounds: [] },
  roulette: { current: null, history: [] },
  teams: { numTeams: 2, result: null },
  updown: { target: null, range: 100, over: false, winner: null, submissions: [] },
  chosung: { current: null, submissions: [] },
  quiz: { question: null, submissions: [], winner: null },
  ice: { usedQs: [], customQs: [], current: null },
  votes: [],
});
const initialState = () => ({
  config: { hostPw: null, createdAt: Date.now() },
  theme: { bgImg: null, teamName: "", dateStart: "", dateEnd: "", place: "" },
  ...initialGame(),
});

// 간이 토큰 (파티용 앱 — 의도적으로 보안 최소화)
const enc = (o) => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
const dec = (s) => { try { return JSON.parse(decodeURIComponent(escape(atob(s)))); } catch { return null; } };

export async function onRequest({ request, env, params }) {
  if (request.method === "OPTIONS") return new Response(null, { headers: CORS });

  if (!env.MT_KV)
    return j({ error: "KV_NOT_BOUND", message: "KV 바인딩(MT_KV)이 설정되지 않았습니다. Cloudflare Pages 설정 > Functions > KV namespace bindings에서 MT_KV를 연결한 뒤 재배포하세요." }, 500);

  const path = Array.isArray(params.path) ? params.path.join("/") : (params.path || "");
  const method = request.method;

  const raw = await env.MT_KV.get(KEY);
  let state = raw ? Object.assign(initialState(), JSON.parse(raw)) : initialState();
  const put = () => env.MT_KV.put(KEY, JSON.stringify(state));

  const auth = dec((request.headers.get("Authorization") || "").replace(/^Bearer\s+/i, ""));
  const role = auth?.role || null;
  const name = auth?.name || null;
  const isStaff = role === "host" || role === "admin";
  const isAdmin = role === "admin";
  const body = method === "POST" ? await request.json().catch(() => ({})) : {};

  // 역할별로 민감 데이터 제거 후 반환
  const sanitize = (forRole) => {
    const s = JSON.parse(JSON.stringify(state));
    s.config = { hostPwSet: !!state.config.hostPw, createdAt: state.config.createdAt };
    s.updown.active = !!state.updown.target && !state.updown.over; // 게임 진행 여부 (정답 노출 없이)
    if (forRole !== "host" && forRole !== "admin") {
      s.updown = { ...s.updown, target: null }; // 게스트에게 정답 숨김
    }
    return s;
  };

  // ── 로그인 ──────────────────────────────
  if (path === "login" && method === "POST") {
    const r = body.role, n = (body.name || "").trim(), p = body.password || "";
    if (r === "admin") {
      const ap = env.ADMIN_PASSWORD || "admin";
      if (p !== ap) return j({ error: "관리자 비밀번호가 틀렸습니다." }, 401);
      return j({ token: enc({ role: "admin", name: "관리자" }), role: "admin", name: "관리자", state: sanitize("admin") });
    }
    if (r === "host") {
      if (!state.config.hostPw) {
        if (p.length < 4) return j({ error: "첫 로그인입니다. 4자 이상의 호스트 비밀번호를 정해서 입력하세요." }, 400);
        state.config.hostPw = p;
        await put();
        return j({ token: enc({ role: "host", name: "호스트" }), role: "host", name: "호스트", firstSetup: true, state: sanitize("host") });
      }
      if (p !== state.config.hostPw) return j({ error: "호스트 비밀번호가 틀렸습니다." }, 401);
      return j({ token: enc({ role: "host", name: "호스트" }), role: "host", name: "호스트", state: sanitize("host") });
    }
    if (r === "guest") {
      if (!n) return j({ error: "이름을 입력하세요." }, 400);
      const found = state.members.find((m) => m.name === n);
      if (!found) return j({ error: "등록되지 않은 참가자입니다. 호스트에게 등록을 요청하세요." }, 401);
      if (p !== n) return j({ error: "비밀번호가 틀렸습니다. (비밀번호는 본인 이름과 동일)" }, 401);
      return j({ token: enc({ role: "guest", name: n }), role: "guest", name: n, state: sanitize("guest") });
    }
    return j({ error: "잘못된 역할입니다." }, 400);
  }

  // ── 테마 조회 (공개 — 로그인 화면에서 사용) ──
  if (path === "theme" && method === "GET") return j(state.theme || {});

  // 이하 모든 경로는 로그인 필요
  if (!role) return j({ error: "로그인이 필요합니다." }, 401);

  // ── 상태 조회 (폴링/새로고침) ──────────────
  if (path === "state" && method === "GET") return j(sanitize(role));

  // ── 데이터 수정 (호스트/Admin) ─────────────
  if (path === "update" && method === "POST") {
    if (!isStaff) return j({ error: "호스트만 수정할 수 있습니다." }, 403);
    const patch = body.patch || {};
    if ("theme" in patch) {
      if (!isAdmin) return j({ error: "테마는 관리자만 수정할 수 있습니다." }, 403);
      state.theme = patch.theme;
    }
    for (const k of PATCH_KEYS) if (k in patch) state[k] = patch[k];
    await put();
    return j(sanitize(role));
  }

  // ── 투표 (전원, 1인 1표 — 마감 전 변경 가능) ──
  if (path === "vote" && method === "POST") {
    const v = state.votes.find((x) => x.id === body.voteId);
    if (!v) return j({ error: "투표를 찾을 수 없습니다." }, 404);
    if (v.closed) return j({ error: "마감된 투표입니다." }, 400);
    if (!v.options.find((o) => o.id === body.optionId)) return j({ error: "잘못된 선택지입니다." }, 400);
    v.voters = v.voters || {};
    v.voters[name] = body.optionId;
    await put();
    return j(sanitize(role));
  }

  // ── 미니게임 답 제출 (전원) ─────────────────
  if (path === "submit" && method === "POST") {
    const g = body.game;
    if (g === "updown") {
      if (!state.updown.target || state.updown.over) return j({ error: "진행 중인 게임이 없습니다." }, 400);
      const val = Number(body.value);
      if (!val) return j({ error: "숫자를 입력하세요." }, 400);
      const dir = val === state.updown.target ? "정답 🎉" : val < state.updown.target ? "UP ⬆️" : "DOWN ⬇️";
      state.updown.submissions.unshift({ name, val, dir, t: Date.now() });
      if (val === state.updown.target) { state.updown.over = true; state.updown.winner = name; }
    } else if (g === "chosung" || g === "quiz") {
      const ans = String(body.value || "").trim().slice(0, 100);
      if (!ans) return j({ error: "답을 입력하세요." }, 400);
      const box = g === "chosung" ? state.chosung : state.quiz;
      const target = g === "chosung" ? state.chosung.current : state.quiz.question;
      if (!target) return j({ error: "진행 중인 문제가 없습니다." }, 400);
      const prev = box.submissions.find((x) => x.name === name);
      if (prev) { prev.answer = ans; prev.t = Date.now(); }
      else box.submissions.push({ name, answer: ans, t: Date.now() });
    } else return j({ error: "알 수 없는 게임입니다." }, 400);
    await put();
    return j(sanitize(role));
  }

  // ── 게임 데이터 초기화 (호스트/Admin — 참가자·계정 유지) ──
  if (path === "host/reset" && method === "POST") {
    if (!isStaff) return j({ error: "권한이 없습니다." }, 403);
    const g = initialGame();
    g.members = state.members;
    state = { config: state.config, theme: state.theme, ...g };
    await put();
    return j(sanitize(role));
  }

  // ── Admin 전용 ─────────────────────────────
  if (path.startsWith("admin/")) {
    if (!isAdmin) return j({ error: "관리자 권한이 필요합니다." }, 403);

    if (path === "admin/unlock-members" && method === "POST") {
      if (!state.config.hostPw) return j({ ok: true, note: "호스트 비밀번호가 아직 설정되지 않아 잠금 없이 열립니다." });
      if ((body.password || "") !== state.config.hostPw) return j({ error: "호스트 비밀번호가 틀렸습니다." }, 401);
      return j({ ok: true });
    }
    if (path === "admin/export" && method === "GET") {
      // 전체 원본 DB (호스트 비번 포함)
      return j({ exportedAt: new Date().toISOString(), state });
    }
    if (path === "admin/import" && method === "POST") {
      const st = body.state;
      if (!st || typeof st !== "object" || !("config" in st)) return j({ error: "올바른 백업 파일이 아닙니다." }, 400);
      state = Object.assign(initialState(), st);
      await put();
      return j({ ok: true, state: sanitize(role) });
    }
    if (path === "admin/sysinfo" && method === "GET") {
      const bytes = JSON.stringify(state).length;
      return j({
        storedBytes: bytes,
        kvLimitBytes: 26214400,
        members: state.members.length,
        notices: (state.notices || []).length,
        cartItems: Object.values(state.cart).reduce((a, v) => a + v.length, 0),
        votes: state.votes.length,
        scheduleItems: state.schedule.length,
        settleItems: state.settle.items.length,
        sbTeams: state.sb.teams.length,
        sbRounds: state.sb.rounds.length,
        hostPwSet: !!state.config.hostPw,
        themeSet: !!(state.theme && (state.theme.teamName || state.theme.bgImg)),
        adminPwCustom: !!env.ADMIN_PASSWORD,
        createdAt: state.config.createdAt,
        serverTime: Date.now(),
      });
    }
    if (path === "admin/hostpw" && method === "POST") {
      state.config.hostPw = null;
      await put();
      return j({ ok: true, message: "호스트 비밀번호가 초기화되었습니다. 다음 호스트 로그인 시 새 비밀번호가 설정됩니다." });
    }
    if (path === "admin/reset" && method === "POST") {
      state = initialState();
      await put();
      return j({ ok: true, message: "시스템이 완전히 초기화되었습니다." });
    }
  }

  return j({ error: "Not found: " + path }, 404);
}
