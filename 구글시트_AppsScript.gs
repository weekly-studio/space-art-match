/**
 * 공간 작품 매칭 — 응답 수집기
 *
 * [설치 방법]
 *  1. 구글 드라이브에서 새 스프레드시트를 하나 만듭니다. (이름 예: 공간진단 응답)
 *  2. 상단 메뉴 → 확장 프로그램 → Apps Script
 *  3. 편집기에 있던 내용을 모두 지우고 이 파일 내용을 통째로 붙여넣습니다.
 *  4. 저장(디스크 아이콘) → 오른쪽 위 [배포] → [새 배포]
 *  5. 유형 선택(톱니바퀴) → [웹 앱]
 *       설명:        공간진단 수집기
 *       실행 계정:   나
 *       액세스 권한: 모든 사용자          ← 이걸 꼭 "모든 사용자"로 하셔야 합니다
 *  6. [배포] → 권한 요청이 뜨면 허용 (본인 계정이니 안전합니다)
 *  7. 마지막에 나오는 "웹 앱 URL"을 복사해서 알려주세요.
 *       https://script.google.com/macros/s/AKfy...../exec  형태입니다.
 *
 * 시트는 자동으로 두 장이 생깁니다.
 *   "응답" — 고객이 제출한 진단 결과 (사진은 드라이브 링크)
 *   "초대" — 큐레이터가 발급한 초대 링크 대장
 *
 * ※ 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전 "새 버전" → [배포]
 *    를 해야 바뀐 내용이 반영됩니다. URL은 그대로 유지됩니다.
 */

/** 담당자 접속 코드. 페이지에는 없고 여기서만 확인한다. 바꾸려면 이 줄만 고치면 된다. */
const ACCESS_CODE = "1440";

/* ================= AI 사진 분석 =================
 *
 * [키 넣는 법 — 한 번만 하면 됩니다]
 *  1. 이 편집기 왼쪽 톱니바퀴(프로젝트 설정) 클릭
 *  2. 맨 아래 "스크립트 속성" → [속성 추가]
 *  3. 속성 이름:  ANTHROPIC_API_KEY
 *     값:         sk-ant-...  (발급받은 키 붙여넣기)
 *  4. 저장 → [배포] → [배포 관리] → 연필 → 새 버전 → [배포]
 *
 *  키는 이 스크립트 안에만 있습니다. 웹페이지로는 절대 나가지 않습니다.
 *  키를 넣지 않으면 관리자 화면의 [AI로 사진 분석] 버튼이 회색으로 남고,
 *  담당자가 직접 입력하는 방식은 그대로 동작합니다.
 */
const AI_MODEL   = "claude-sonnet-5";     // 더 정밀하게 하려면 "claude-opus-5"
const AI_MAXTOK  = 1200;

function aiKey() {
  try { return PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY") || ""; }
  catch (e) { return ""; }
}
function aiReady() { return aiKey().length > 10; }

/** 응답 행에 붙어 있는 공간 사진을 드라이브에서 꺼내 온다 */
function rowPhotoBlobs(d) {
  const f = findRespRow(d);
  if (!f.row) return [];
  const col = RESP_HEAD.indexOf("사진1") + 1;
  const out = [];
  for (let i = 0; i < 3; i++) {
    const cell = f.sh.getRange(f.row, col + i);
    const src = String(cell.getNote() || "") + " " + String(cell.getFormula() || "") + " " + String(cell.getValue() || "");
    const m = /\/d\/([A-Za-z0-9_-]+)/.exec(src) || /id=([A-Za-z0-9_-]+)/.exec(src);
    if (!m) continue;
    try { out.push(DriveApp.getFileById(m[1]).getBlob()); } catch (e) { /* 지워진 파일 */ }
  }
  return out;
}

/** 화면이 보낸 data URL 을 블롭으로 */
function dataUrlBlobs(list) {
  return (list || []).map(function (u) {
    const m = /^data:([^;]+);base64,(.+)$/.exec(String(u || ""));
    return m ? Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], "photo.jpg") : null;
  }).filter(Boolean);
}

/** Claude 에 이미지와 함께 묻고, 돌려받은 JSON 을 그대로 돌려준다 */
function askClaude(prompt, blobs) {
  const key = aiKey();
  if (!key) return { ok: false, error: "API 키가 아직 설정되지 않았습니다. 앱스스크립트 → 프로젝트 설정 → 스크립트 속성에 ANTHROPIC_API_KEY 를 넣어 주세요." };

  const content = [];
  (blobs || []).slice(0, 3).forEach(function (b) {
    content.push({
      type: "image",
      source: { type: "base64", media_type: b.getContentType() || "image/jpeg",
                data: Utilities.base64Encode(b.getBytes()) }
    });
  });
  content.push({ type: "text", text: prompt });

  let res;
  try {
    res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      contentType: "application/json",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
      payload: JSON.stringify({ model: AI_MODEL, max_tokens: AI_MAXTOK,
                                messages: [{ role: "user", content: content }] }),
      muteHttpExceptions: true
    });
  } catch (e) {
    return { ok: false, error: "AI 서버에 연결하지 못했습니다: " + e };
  }

  const code = res.getResponseCode();
  const raw  = res.getContentText();
  if (code !== 200) {
    let msg = raw;
    try { msg = JSON.parse(raw).error.message; } catch (e) {}
    return { ok: false, error: "AI 응답 오류(" + code + "): " + String(msg).slice(0, 300) };
  }

  let text = "";
  try {
    const body = JSON.parse(raw);
    text = (body.content || []).map(function (c) { return c.text || ""; }).join("");
  } catch (e) { return { ok: false, error: "AI 응답을 읽지 못했습니다." }; }

  // ```json 울타리가 붙어 와도 벗겨서 읽는다
  const fence = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
  const jsonText = fence ? fence[1] : text;
  const start = jsonText.indexOf("{");
  const end   = jsonText.lastIndexOf("}");
  if (start < 0 || end < start) return { ok: false, error: "AI가 형식에 맞지 않게 답했습니다." };
  try { return { ok: true, data: JSON.parse(jsonText.slice(start, end + 1)) }; }
  catch (e) { return { ok: false, error: "AI 답을 해석하지 못했습니다." }; }
}

/**
 * 공간 사진 분석. 화면(recommend)이 바로 쓰는 네 값을 반드시 포함한다.
 *   wallIdx 0~3 · lightIdx 0~3 · roomIdx 0~3 · toneShift -30~30
 */
function aiAnalyzeSpace(d) {
  let blobs = dataUrlBlobs(d.photos);
  if (!blobs.length) blobs = rowPhotoBlobs(d);
  if (!blobs.length) return { ok: false, error: "분석할 공간 사진이 없습니다. 사진을 먼저 넣어 주세요." };

  const a = d.answers || {};
  const prompt =
    "당신은 한국의 주거·상업 공간 인테리어와 미술품 배치를 전문으로 하는 공간 디자이너입니다. " +
    "첨부된 공간 사진(최대 3장)을 보고 미술 작품을 걸기 위한 조건을 분석하세요.\n" +
    "고객 자가응답(참고용, 틀릴 수 있음): 공간=" + (a.place || "") + "; 벽 폭=" + (a.wall || "") +
    "; 조도=" + (a.light || "") + "; 실내 톤=" + (a.tone_room || "") + "; 메모=" + (d.memo || "없음") + ".\n" +
    "반드시 아래 JSON만 출력하세요(설명 금지). 문장은 고객에게 보여줄 수 있는 존댓말로:\n" +
    '{"wall":"벽 색과 마감 한 문장","light":"빛의 종류·방향·세기 한 문장","style":"가구·바닥·전체 스타일 한 문장",' +
    '"wallWidth":"작품이 걸릴 벽의 가로 폭 추정 한 문장(예: 소파 폭 기준 약 2.4m로 보입니다)",' +
    '"placement":"작품이 걸릴 위치와 권장 높이·크기 한 문장","advice":"추천 시 유의할 점 한두 문장",' +
    '"wallIdx":0~3 정수(0:1m미만,1:1~2m,2:2~3m,3:3m이상),' +
    '"lightIdx":0~3 정수(0:종일 자연광,1:오후만,2:자연광 거의 없음,3:스팟조명),' +
    '"roomIdx":0~3 정수(0:화이트+원목,1:그레이·블랙 모던,2:베이지·아이보리,3:색·패턴 많음),' +
    '"toneShift":-30~30 정수(실내가 이미 색이 많으면 음수, 무채·어두우면 양수, 보통 0),' +
    '"quality":"사진 품질 한 문장"}';

  const r = askClaude(prompt, blobs);
  if (!r.ok) return r;

  const o = r.data;
  ["wallIdx", "lightIdx", "roomIdx"].forEach(function (k) {
    o[k] = Math.max(0, Math.min(3, Math.round(Number(o[k]) || 0)));
  });
  o.toneShift = Math.max(-30, Math.min(30, Math.round((Number(o.toneShift) || 0) / 10) * 10));
  o.by = "ai";
  o.at = new Date().toISOString();

  // 분석 결과를 시트에도 남겨 다른 기기에서 같은 기준으로 추천되게 한다
  const f = findRespRow(d);
  if (f.row) {
    ensureHead(f.sh, 28, "확정작품");
    ensureHead(f.sh, 29, "공간분석");
    f.sh.getRange(f.row, 29).setValue(JSON.stringify(o));
  }
  return { ok: true, wrote: true, analysis: o };
}

/** 추천 5점 각각에 "왜 이 공간에 맞는지" 코멘트 */
function aiWorkComments(d) {
  let blobs = dataUrlBlobs(d.photos);
  if (!blobs.length) blobs = rowPhotoBlobs(d);

  const prompt =
    "당신은 미술품 큐레이터입니다. 첨부된 고객의 공간 사진을 보고, 아래 추천 작품 각각에 대해 " +
    '"왜 이 작품이 이 공간에 맞는지"를 고객에게 보내는 말투(존댓말, 2~3문장, 사진 속 실제 요소—벽 색, 소파, 창, 바닥—를 구체적으로 언급)로 써주세요. ' +
    "과장 금지. 작품 외형은 주어진 성격 설명 범위 안에서만 쓰세요.\n" +
    "고객 취향 유형: " + (d.type || "") + ". 공간 분석: " + (d.space || "없음") + "\n" +
    "추천 작품:\n" + (d.works || "") + "\n" +
    'JSON만 출력: {"작품제목":"코멘트", ...} — 키는 〈 〉 안의 제목을 그대로.';

  const r = askClaude(prompt, blobs);
  if (!r.ok) return r;
  return { ok: true, wrote: true, comments: r.data };
}


/** 공간 사진이 쌓일 드라이브 폴더 이름 */
const PHOTO_FOLDER = "공간진단 사진";

const Q_TAGS = ["공간", "규격", "조도", "실내 톤", "정서", "화면", "인상", "색", "재료", "동기", "예산", "단계"];
const Q_IDS = ["place", "wall", "light", "tone_room", "mood", "density", "form", "color", "matter", "motive", "budget", "stage"];

const RESP_HEAD = ["접수시각", "성함", "호칭", "연락처", "초대코드", "담당", "유형", "채도", "추상", "밀도", "추천동기"]
  .concat(Q_TAGS)
  .concat(["메모", "사진1", "사진2", "사진3", "확정작품", "공간분석"]);

const INV_HEAD = ["발급시각", "성함", "호칭", "연락처", "초대코드", "담당", "상담메모", "초대링크"];


function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.kind === "check")  return json(checkGuest(d));
    if (d.kind === "invite") return json(saveInvite(d));

    // 담당자용 조회는 접속 코드를 확인한다
    if (d.kind === "auth")      return json({ ok: authed(d), auth: authed(d) });
    if (d.kind === "invites")   return json(authed(d) ? listInvites()   : denied());
    if (d.kind === "responses") return json(authed(d) ? listResponses() : denied());

    // 시트를 고치고 지우는 것도 담당자만
    if (d.kind === "invite.update")   return json(authed(d) ? updateInvite(d)   : denied());
    if (d.kind === "invite.delete")   return json(authed(d) ? deleteInvite(d)   : denied());
    if (d.kind === "response.update") return json(authed(d) ? updateResponse(d) : denied());
    if (d.kind === "response.delete") return json(authed(d) ? deleteResponse(d) : denied());

    // 담당자가 관리자 화면에서 넣는 공간 사진
    if (d.kind === "response.photo")        return json(authed(d) ? putResponsePhoto(d)  : denied());
    if (d.kind === "response.photo.delete") return json(authed(d) ? dropResponsePhoto(d) : denied());

    // AI 사진 분석 — API 키는 여기(서버)에만 있고 화면에는 나가지 않는다
    if (d.kind === "ai.status")  return json(authed(d) ? { ok: true, ready: aiReady(), model: AI_MODEL } : denied());
    if (d.kind === "ai.space")   return json(authed(d) ? aiAnalyzeSpace(d)  : denied());
    if (d.kind === "ai.comment") return json(authed(d) ? aiWorkComments(d)  : denied());

    return json(saveResponse(d));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** 배포된 코드가 어느 버전인지 알려준다. 주소를 브라우저로 열면 보인다. */
const VERSION = "v12-시트 사진 직접 얹기";

function doGet() {
  return json({
    ok: true,
    version: VERSION,
    features: ["응답저장", "초대기록", "명단조회", "사진표시", "AI분석"],
    aiReady: aiReady(),
    msg: "공간 작품 매칭 수집기가 정상 동작 중입니다."
  });
}


function saveResponse(d) {
  const a = d.answers || {};
  const ax = d.axis || {};
  const sh = sheet("응답", RESP_HEAD);

  // 사진 업로드보다 먼저 행을 남긴다. 업로드가 실패해도 응답은 보존된다.
  sh.appendRow(
    [
      new Date(),
      d.name || "",
      (d.invite && d.invite.title) || "",
      asText(d.phone),
      d.inviteCode || "",
      (d.invite && d.invite.by) || "",
      d.type || "",
      ax.tone, ax.form, ax.density,
      d.motiveLabel || d.motive || ""
    ]
      .concat(Q_IDS.map(function (k) { return a[k] || ""; }))
      .concat([d.memo || "", "", "", ""])
  );

  const row = sh.getLastRow();          // 사진을 넣는 동안 다른 응답이 들어와도 이 행에 쓴다
  const files = savePhotos(d);
  const shown = writePhotoCells(sh, row, files);
  return {
    ok: true,
    photos: files.filter(function (f) { return f && f.id; }).length,
    shown: shown
  };
}


/**
 * 초대 명단에서 성함을 찾는다. 발급할 때 연락처를 남긴 분은 뒤 4자리까지
 * 맞아야 통과한다. 명단 전체를 돌려주지 않으므로 목록이 새어 나가지 않는다.
 */
function checkGuest(d) {
  const name = String(d.name || "").replace(/\s/g, "");
  const tail = String(d.tail || "").replace(/\D/g, "");
  if (!name) return { ok: false };

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("초대");
  if (!sh || sh.getLastRow() < 2) return { ok: false };

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, INV_HEAD.length).getValues();
  let nameHit = false;

  for (let i = rows.length - 1; i >= 0; i--) {   // 최근 발급부터 본다
    const r = rows[i];
    if (String(r[1] || "").replace(/\s/g, "") !== name) continue;
    nameHit = true;

    const phone = String(r[3] || "").replace(/^'/, "");
    const digits = phone.replace(/\D/g, "");
    if (tail && digits && digits.slice(-4) !== tail) continue;   // 뒤 4자리를 보냈을 때만 대조

    return {
      ok: true,
      name:  String(r[1] || ""),
      title: String(r[2] || ""),
      phone: phone,
      code:  String(r[4] || ""),
      by:    String(r[5] || ""),
      at:    r[0] ? new Date(r[0]).toISOString() : ""
    };
  }
  // 이름은 있는데 뒤 4자리가 틀린 경우를 구분해 안내를 다르게 준다
  return { ok: false, nameHit: nameHit };
}


function authed(d) { return String((d && d.code) || "") === ACCESS_CODE; }

/** 초대 시트에서 코드로 행 번호를 찾는다. 없으면 0 */
function findInviteRow(code) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("초대");
  if (!sh || sh.getLastRow() < 2 || !code) return { sh: sh, row: 0 };
  const col = sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues();   // E열 = 초대코드
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0] || "").trim() === String(code).trim()) return { sh: sh, row: i + 2 };
  }
  return { sh: sh, row: 0 };
}

function updateInvite(d) {
  const f = findInviteRow(d.inviteCode || d.icode);   // 접속 코드(code)와 헷갈리지 않게
  if (!f.row) return { ok: false, error: "초대를 찾지 못했습니다." };
  f.sh.getRange(f.row, 2, 1, 2).setValues([[d.name || "", d.title || ""]]);   // B,C
  f.sh.getRange(f.row, 4).setValue(asText(d.phone));                          // D
  f.sh.getRange(f.row, 6, 1, 2).setValues([[d.by || "", d.memo || ""]]);      // F,G
  return { ok: true, wrote: true };
}

function deleteInvite(d) {
  const f = findInviteRow(d.inviteCode || d.icode);
  if (!f.row) return { ok: false, error: "초대를 찾지 못했습니다." };
  f.sh.deleteRow(f.row);
  return { ok: true, wrote: true };
}

/** 응답 시트에서 행을 찾는다. id 는 "s12" 형태이며 성함으로 한 번 더 대조한다. */
function findRespRow(d) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("응답");
  const row = Number(String(d.id || "").replace(/^s/, ""));
  if (!sh || !row || row < 2 || row > sh.getLastRow()) return { sh: sh, row: 0 };
  const name = String(sh.getRange(row, 2).getValue() || "").trim();
  if (d.at) {                       // 행이 밀렸는지 접수시각으로 확인
    const at = sh.getRange(row, 1).getValue();
    const iso = at ? new Date(at).toISOString() : "";
    if (iso && String(d.at) !== iso) return { sh: sh, row: 0 };
  } else if (d.was && name !== String(d.was).trim()) {
    return { sh: sh, row: 0 };
  }
  return { sh: sh, row: row };
}

function updateResponse(d) {
  const f = findRespRow(d);
  if (!f.row) return { ok: false, error: "응답을 찾지 못했습니다. 새로고침 후 다시 시도해 주세요." };
  f.sh.getRange(f.row, 2, 1, 2).setValues([[d.name || "", d.title || ""]]);   // B,C
  f.sh.getRange(f.row, 4).setValue(asText(d.phone));                          // D
  f.sh.getRange(f.row, 6).setValue(d.by || "");                               // F
  f.sh.getRange(f.row, 24).setValue(d.memo || "");                            // X = 메모
  if (d.picks !== undefined) {                                                // AB = 확정작품
    ensureHead(f.sh, 28, "확정작품");
    f.sh.getRange(f.row, 28).setValue(Array.isArray(d.picks) ? d.picks.join(",") : String(d.picks || ""));
  }
  if (d.space !== undefined) {                                                // AC = 공간분석
    ensureHead(f.sh, 28, "확정작품");                                          // 28번을 건너뛰고 29번만 생기지 않게
    ensureHead(f.sh, 29, "공간분석");
    f.sh.getRange(f.row, 29).setValue(String(d.space || ""));
  }
  return { ok: true, wrote: true };
}

function deleteResponse(d) {
  const f = findRespRow(d);
  if (!f.row) return { ok: false, error: "응답을 찾지 못했습니다. 새로고침 후 다시 시도해 주세요." };
  f.sh.deleteRow(f.row);
  return { ok: true, wrote: true };
}
function denied()   { return { ok: false, error: "접속 코드가 맞지 않습니다." }; }


/**
 * 담당자 화면에 뿌릴 응답 목록. 어느 기기에서 열어도 같은 목록이 보인다.
 * 사진은 드라이브 링크로 넘긴다.
 */
function listResponses() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("응답");
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };

  const width = Math.max(sh.getLastColumn(), 1);
  const range = sh.getRange(2, 1, sh.getLastRow() - 1, width);
  const rows = range.getValues();
  const forms = range.getFormulas();
  const notes = range.getNotes();             // 사진 주소는 칸 메모에 있다
  const items = rows.map(function (r, i) {
    return {
      id:          "s" + (i + 2),                 // 시트 행 번호
      at:          r[0] ? new Date(r[0]).toISOString() : "",
      name:        String(r[1] || ""),
      title:       String(r[2] || ""),
      phone:       String(r[3] || "").replace(/^'/, ""),
      inviteCode:  String(r[4] || ""),
      by:          String(r[5] || ""),
      type:        String(r[6] || ""),
      tone:        Number(r[7]) || 50,
      form:        Number(r[8]) || 50,
      density:     Number(r[9]) || 50,
      motiveLabel: String(r[10] || ""),
      answers:     Q_IDS.map(function (k, j) { return String(r[11 + j] || ""); }),
      memo:        String(r[23] || ""),
      photos:      photoLinks(r, forms[i], notes[i]),
      picks:       String(r[27] || "").split(",").map(function (x) { return x.trim(); }),
      space:       String(r[28] || "")
    };
  }).filter(function (x) { return x.name; });

  items.reverse();                                // 최근 접수가 위로
  return { ok: true, items: items };
}


/**
 * 담당자 화면에 뿌릴 초대 명단. 어느 기기에서 열어도 같은 목록이 보이게 한다.
 * 이 주소는 공개되어 있으므로 연락처(D열)는 내보내지 않는다.
 */
function listInvites() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("초대");
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };

  // 응답 시트의 초대코드를 모아 완료 여부를 판단한다
  const done = {};
  const rs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("응답");
  if (rs && rs.getLastRow() > 1) {
    rs.getRange(2, 5, rs.getLastRow() - 1, 1).getValues().forEach(function (r) {
      const c = String(r[0] || "").trim();
      if (c) done[c] = true;
    });
  }

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, INV_HEAD.length).getValues();
  const items = rows.map(function (r) {
    const code = String(r[4] || "");
    return {
      at:    r[0] ? new Date(r[0]).toISOString() : "",
      name:  String(r[1] || ""),
      title: String(r[2] || ""),
      code:  code,
      by:    String(r[5] || ""),
      memo:  String(r[6] || ""),
      done:  !!done[code]
    };
  }).filter(function (x) { return x.name && x.code; });

  items.reverse();                    // 최근 발급이 위로
  return { ok: true, items: items };
}


function saveInvite(d) {
  sheet("초대", INV_HEAD).appendRow([
    new Date(),
    d.name || "", d.title || "", asText(d.phone),
    d.code || "", d.by || "", d.memo || "", d.url || ""
  ]);
  return { ok: true };
}


/**
 * 사진(data URL)을 드라이브에 올린다. 파일은 비공개로 남는다.
 * 시트에는 그림을 직접 얹기 때문에(=IMAGE() 를 쓰지 않는다) 공개할 필요가 없다.
 */
function savePhotos(d) {
  const photos = d.photos || [];
  if (!photos.length) return [];

  const first = Number(d.startIndex) || 0;      // 담당자가 2번 칸만 바꿀 때 파일 이름을 맞춘다
  const folder = getFolder();
  return photos.map(function (dataUrl, i) {
    try {
      const m = /^data:([^;]+);base64,(.+)$/.exec(dataUrl);
      if (!m) return null;
      const name = [d.name || "고객", d.inviteCode || "", first + i + 1].join("_") + ".jpg";
      const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
      // 사진은 비공개로 둔다. 시트에는 그림을 직접 얹으므로 공개할 필요가 없다.
      return { id: file.getId(), url: file.getUrl(), blob: file.getBlob() };
    } catch (err) {
      return { id: "", url: "저장 실패: " + err, shared: false };
    }
  });
}


/**
 * 사진1~3 칸에 사진을 보이게 넣는다.
 *
 * 구글이 =IMAGE() 로 드라이브 사진 불러오는 것을 막아 두어(#REF! 가 뜬다)
 * 그림을 시트 위에 직접 얹는다. 파일을 공개하지 않아도 보이고, 링크 정책이
 * 바뀌어도 깨지지 않는다. 원본 주소는 칸 메모에 남겨 두므로 마우스를 올리면 보인다.
 */
function writePhotoCells(sh, row, files) {
  if (!row || !files || !files.length) return 0;
  const col = RESP_HEAD.indexOf("사진1") + 1;      // 머릿글에서 직접 찾는다
  if (col < 1) return 0;
  let shown = 0;

  files.slice(0, 3).forEach(function (f, i) {      // i 가 곧 사진 칸 번호다. null 은 건드리지 않는다
    if (!f) return;
    const cell = sh.getRange(row, col + i);
    cell.clearContent();                            // 예전 =IMAGE() 수식·#REF! 를 걷어낸다
    dropCellImage(sh, col + i, row);                // 같은 칸의 옛 그림도 치운다

    if (!f.id) { cell.setValue(f.url || ""); return; }   // 저장 실패 메시지만 남는 경우

    cell.setNote(f.url);                            // 원본 드라이브 주소
    try {
      const blob = f.blob || DriveApp.getFileById(f.id).getBlob();
      const im = sh.insertImage(blob, col + i, row);
      const w = im.getWidth() || 1, h = im.getHeight() || 1;
      im.setWidth(150).setHeight(Math.max(40, Math.round(150 * h / w)));
      shown++;
    } catch (e) {
      cell.setValue(f.url || "");                   // 그림을 못 얹으면 주소라도 남긴다
    }
  });

  if (shown) {
    sh.setRowHeight(row, 125);
    for (let c = 0; c < 3; c++) {
      if (sh.getColumnWidth(col + c) < 170) sh.setColumnWidth(col + c, 170);
    }
  }
  return shown;
}

/** 특정 칸에 얹힌 그림을 치운다 */
function dropCellImage(sh, col, row) {
  try {
    sh.getImages().forEach(function (im) {
      const a = im.getAnchorCell();
      if (a.getColumn() === col && a.getRow() === row) im.remove();
    });
  } catch (e) {}
}


/** 사진 칸에서 원본 링크를 되살린다. =IMAGE() 가 들어 있으면 파일 id 로 주소를 만든다. */
function photoLinks(values, formulas, notes) {
  const out = [];
  const first = RESP_HEAD.indexOf("사진1");
  if (first < 0) return out;
  for (let c = first; c < first + 3; c++) {
    const note = String((notes && notes[c]) || "");
    if (note.indexOf("http") === 0) { out.push(note.split(/\s/)[0]); continue; }   // 지금 방식
    const v = String(values[c] || "");
    if (v.indexOf("http") === 0) { out.push(v); continue; }                        // 옛 방식(주소만)
    const m = /id=([A-Za-z0-9_-]+)/.exec(String((formulas && formulas[c]) || ""));  // 더 옛 방식(=IMAGE)
    if (m) out.push("https://drive.google.com/file/d/" + m[1] + "/view");
  }
  return out;
}


/**
 * 담당자가 관리자 화면에서 넣는 공간 사진 한 장.
 * 고객은 사진을 올리지 않는다. 카카오톡 등으로 받은 사진을 담당자가 여기에 넣으면
 * 드라이브에 저장되고 해당 응답 행의 사진 칸에 그림으로 박힌다.
 */
function putResponsePhoto(d) {
  const f = findRespRow(d);
  if (!f.row) return { ok: false, error: "응답을 찾지 못했습니다. 새로고침 후 다시 시도해 주세요." };
  const slot = Math.max(0, Math.min(2, Number(d.slot) || 0));
  const files = savePhotos({ photos: [d.photo], name: d.name, inviteCode: d.inviteCode, startIndex: slot });
  const file = files[0];
  if (!file || !file.id) return { ok: false, error: (file && file.url) || "사진을 저장하지 못했습니다." };

  const slots = [null, null, null];
  slots[slot] = file;
  writePhotoCells(f.sh, f.row, slots);
  return { ok: true, wrote: true, url: file.url };
}

function dropResponsePhoto(d) {
  const f = findRespRow(d);
  if (!f.row) return { ok: false, error: "응답을 찾지 못했습니다." };
  const slot = Math.max(0, Math.min(2, Number(d.slot) || 0));
  const col = RESP_HEAD.indexOf("사진1") + 1 + slot;
  const cell = f.sh.getRange(f.row, col);
  cell.clearContent();
  cell.clearNote();
  dropCellImage(f.sh, col, f.row);
  return { ok: true, wrote: true };
}


function getFolder() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}


/** 뒤에 붙는 열의 머릿글을 그 칸만 보고 채운다. 마지막 열 번호로 판단하면 한 칸씩 빈다. */
function ensureHead(sh, col, name) {
  if (String(sh.getRange(1, col).getValue() || "").trim() !== name) sh.getRange(1, col).setValue(name);
}

function sheet(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight("bold");
    sh.setColumnWidth(1, 140);
  }
  return sh;
}


/** 010으로 시작하는 번호의 앞자리 0이 사라지지 않게 텍스트로 고정 */
function asText(v) {
  const s = String(v == null ? "" : v);
  return s ? "'" + s : "";
}


function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
