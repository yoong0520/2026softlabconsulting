/**
 * 수업운영 컨설팅 수신기 (Google Apps Script 웹 앱) — 실증 결과보고서 수신기와 별도 프로젝트
 * - 교수자 자가진단 제출 → 접수 코드 발급, PDF는 「1_컨설팅 요청」 폴더, 데이터는 「9_데이터」 폴더에 저장
 * - 컨설턴트가 접수 코드로 자가진단을 불러오고(GET ?code=), 결과 제출 시 PDF·DOCX는 「2_컨설팅 결과」 폴더에 저장
 * - 모든 제출은 「수업운영 컨설팅 현황」 시트에 한 줄씩 기록
 * 배포: "웹 앱" / 실행 주체: 나 / 액세스 권한: 모든 사용자
 */
var REQ_FOLDER_ID  = "129sbRsDgaGedlG4aWVngO6YJ0shFAFrg"; // 1_컨설팅 요청(교수자 자가진단)
var RES_FOLDER_ID  = "10dkmzkQEOF6nWFfOBJHrQpJI-VFtGyeX"; // 2_컨설팅 결과(컨설턴트 검토)
var DATA_FOLDER_ID = "1qWY3KpMTE1IetfUbwMot7AP9YJU9m6_k"; // 9_데이터(자동 저장, 수정 금지)
var LOG_SHEET_ID   = "1mfXnWXz68eK1ues_ZFxWbMLSVY76xjUITVDgYSTzuvU"; // 수업운영 컨설팅 현황
var MAX_BYTES = 25 * 1024 * 1024;

function doGet(e) {
  var code = e && e.parameter && e.parameter.code;
  if (code) return out(loadConsult(String(code).trim().toUpperCase()));
  return out({ ok: true, message: "수업운영 컨설팅 수신기가 동작 중입니다" });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || "{}");
    if (data.kind !== "consult-request" && data.kind !== "consult-result") return out({ ok: false, message: "알 수 없는 요청 종류입니다" });
    return out(saveConsult(data));
  } catch (err) {
    return out({ ok: false, message: "저장 중 오류: " + err });
  }
}

function newCode() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789", c = "";
  for (var i = 0; i < 6; i++) c += chars.charAt(Math.floor(Math.random() * chars.length));
  return c;
}
function findData(code, prefix) {
  var it = DriveApp.getFolderById(DATA_FOLDER_ID).getFilesByName(prefix + "_" + code + ".json");
  return it.hasNext() ? it.next() : null;
}
function saveConsult(data) {
  var isReq = data.kind === "consult-request";
  var given = String(data.code || "").trim().toUpperCase();
  var code = isReq ? (given && findData(given, "req") ? given : newCode()) : given;
  if (!isReq && !findData(code, "req")) return { ok: false, message: "접수 코드 " + code + "에 해당하는 교수자 자가진단이 없습니다" };

  var prefix = isReq ? "req" : "res";
  var payload = JSON.stringify({ code: code, kind: data.kind, savedAt: new Date().toISOString(), meta: data.meta || {}, data: data.data || {} });
  var old = findData(code, prefix);
  if (old) old.setContent(payload); else DriveApp.getFolderById(DATA_FOLDER_ID).createFile(prefix + "_" + code + ".json", payload, "application/json");

  var folder = DriveApp.getFolderById(isReq ? REQ_FOLDER_ID : RES_FOLDER_ID);
  var stamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmm");
  var saved = [];
  (data.files || []).forEach(function (f) {
    var bytes = Utilities.base64Decode(f.base64);
    if (bytes.length > MAX_BYTES) return;
    var name = safeName(f.filename || (prefix + "_" + code));
    if (folder.getFilesByName(name).hasNext()) {
      var dot = name.lastIndexOf(".");
      name = dot > 0 ? name.slice(0, dot) + "_" + stamp + name.slice(dot) : name + "_" + stamp;
    }
    var file = folder.createFile(Utilities.newBlob(bytes, f.mime || "application/octet-stream", name));
    saved.push({ name: name, url: file.getUrl() });
  });
  try {
    var sh = SpreadsheetApp.openById(LOG_SHEET_ID).getSheets()[0];
    var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");
    var m = data.meta || {};
    (saved.length ? saved : [{ name: "", url: "" }]).forEach(function (sv) {
      sh.appendRow([now, isReq ? "자가진단 제출" : "컨설팅 결과 제출", code, m.track || "", m.course || "", m.prof || "", m.dept || "", m.cw || "", m.consultant || "", m.avg || "", sv.name, sv.url]);
    });
  } catch (err) {}
  return { ok: true, code: code, folder: folder.getName(), files: saved };
}
function loadConsult(code) {
  if (!code) return { ok: false, message: "접수 코드가 없습니다" };
  var req = findData(code, "req");
  if (!req) return { ok: false, message: "접수 코드 " + code + "에 해당하는 자가진단이 없습니다. 코드를 확인해 주세요" };
  var res = findData(code, "res");
  var o = { ok: true, code: code, request: JSON.parse(req.getBlob().getDataAsString()) };
  if (res) o.result = JSON.parse(res.getBlob().getDataAsString());
  return o;
}
function safeName(n) { return String(n).replace(/[\\/:*?"<>|]/g, "_").slice(0, 120); }
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
