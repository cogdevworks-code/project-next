// 오프라인 단일 파일 생성: index.plain.html + mp3 base64 임베드 → Project_Next_OFFLINE.html
// plain은 게이트가 없어 동기(cloudInit/afterUnlock)가 호출되지 않음 → 오프라인에서 네트워크 0.
const fs = require("fs");

let html = fs.readFileSync("index.plain.html", "utf8");

const files = { mitech_ko:"audio/mitech_ko.mp3", mitech_en:"audio/mitech_en.mp3", dnotitia_ko:"audio/dnotitia_ko.mp3" };
const audio = {};
for (const k in files) {
  if (fs.existsSync(files[k])) {
    audio[k] = "data:audio/mpeg;base64," + fs.readFileSync(files[k]).toString("base64");
    console.log("embedded", k, "(", Math.round(fs.statSync(files[k]).size/1024), "KB )");
  } else {
    console.log("⚠️ missing", files[k]);
  }
}

const inject = `<script>const OFFLINE_AUDIO=${JSON.stringify(audio)};</script>\n`;
// DATA 스크립트 시작 직전에 OFFLINE_AUDIO 주입
const anchor = html.indexOf("const DATA = {");
const scriptOpen = html.lastIndexOf("<script>", anchor);
html = html.slice(0, scriptOpen) + inject + html.slice(scriptOpen);

fs.writeFileSync("Project_Next_OFFLINE.html", html);
console.log("offline written: Project_Next_OFFLINE.html (", Math.round(fs.statSync("Project_Next_OFFLINE.html").size/1024), "KB )");
