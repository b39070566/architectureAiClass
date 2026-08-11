const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5
pres.author = "vibe coding 工作坊";
pres.title = "三小時，做出自己的網頁";

/* ── 色票：與 demos/ 的紙本色系一致 ── */
const INK      = "1B1C1A";  // 墨
const INK_2    = "55574F";
const PAPER    = "FBFAF7";  // 暖白紙
const CARD     = "FFFFFF";
const MUTED    = "86887F";
const MUTED_LT = "A5A79F";
const RULE     = "DDD9D0";
const RULE_2   = "C3BEB2";
const BRASS    = "8A6524";  // 建案網站段
const BRASS_LT = "D0A75E";
const BLUE     = "35566B";  // 儀表板段
const BLUE_LT  = "8FB3C9";
const GOOD     = "3F6B48";
const WARN     = "8A6524";
const CRIT     = "93402F";

const FS = "Microsoft JhengHei";   // 中文
const FM = "Consolas";             // 只用在真的是程式碼／檔名的地方

const M = 0.65;                    // 邊界
const W = 13.333 - M * 2;          // 內容寬 12.03

/* ── 共用元件 ── */
function darkBg(s) {
  s.background = { color: INK };
}

/* 頁面上方的小標。中文、不加字距、不做全大寫等寬那一套。 */
function eyebrow(s, txt, color, x, y) {
  s.addText(txt, {
    x: x === undefined ? M : x, y: y === undefined ? 0.52 : y, w: 8, h: 0.3,
    fontFace: FS, fontSize: 12, color: color || MUTED,
    margin: 0, valign: "middle",
  });
}

/* 標題，下面壓一條通欄細線——講義的標題就是這樣分段的 */
function title(s, txt, color, y) {
  const ty = y === undefined ? 0.92 : y;
  s.addText(txt, {
    x: M, y: ty, w: W, h: 0.78, fontFace: FS, fontSize: 32,
    bold: true, color: color || INK, margin: 0, valign: "middle",
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: ty + 0.86, w: W, h: 0, line: { color: RULE_2, width: 0.75 },
  });
}

/* 原本是有邊框的白卡片，現在只在區塊頂端畫一條線。
   幾何位置不變，所以每一頁的文字排版都不用動。 */
function card(s, x, y, w, h, opts) {
  const o = opts || {};
  if (o.fill && o.fill !== CARD) {
    s.addShape(pres.ShapeType.rect, {
      x: x, y: y, w: w, h: h, fill: { color: o.fill }, line: { color: o.fill, width: 0 },
    });
  }
  s.addShape(pres.ShapeType.line, {
    x: x, y: y, w: w, h: 0,
    line: { color: o.line || RULE_2, width: o.lw || 0.75 },
  });
}
function note(s, txt) { s.addNotes(txt); }

/* 內容頁底 */
function light() {
  const s = pres.addSlide();
  s.background = { color: PAPER };
  return s;
}

/* ══════════════════ 開場 ══════════════════ */

/* 1 封面 */
{
  const s = pres.addSlide();
  darkBg(s);
  s.addText("建商員工 AI 應用工作坊", {
    x: M, y: 2.2, w: 9, h: 0.34, fontFace: FS, fontSize: 12,
    color: BRASS_LT, charSpacing: 3, margin: 0,
  });
  s.addText("三小時，做出自己的網頁", {
    x: M, y: 2.75, w: 11.5, h: 1.3, fontFace: FS, fontSize: 52,
    bold: true, color: "F2F5F4", margin: 0, valign: "middle",
  });
  s.addText("不寫程式，用說的", {
    x: M, y: 4.15, w: 9, h: 0.5, fontFace: FS, fontSize: 21,
    color: MUTED_LT, margin: 0, charSpacing: 2,
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: 5.2, w: 3.2, h: 0, line: { color: BRASS, width: 1.5 },
  });
  s.addText("全程以「沐森」建案為例，四個階段逐步累積", {
    x: M, y: 5.45, w: 9, h: 0.4, fontFace: FS, fontSize: 13,
    color: MUTED, margin: 0,
  });
  note(s, "開場先問一句：今天在座有多少人自己做過網頁？通常是零。這正是重點——三小時後，每個人手上都會有一個能打開給主管看的東西。");
}

/* 2 今天的四個產出 */
{
  const s = light();
  eyebrow(s, "今天的產出", BRASS);
  title(s, "今天你會做出這四樣東西");
  const items = [
    ["第一段", "資訊網站", "把建案介紹變成一個網頁", BRASS],
    ["第二段", "加上自己的圖片", "換上你部門拍的照片", BRASS],
    ["第三段", "靜態儀表板", "把 Excel 進度表變成圖表", BLUE],
    ["第四段", "可互動儀表板", "換一份資料，畫面跟著算", BLUE],
  ];
  const cw = 2.82, gap = 0.24;
  items.forEach((it, i) => {
    const x = M + i * (cw + gap);
    card(s, x, 2.05, cw, 3.35);
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.3, y: 2.4, w: 0.34, h: 0.34, fill: { color: it[3] },
    });
    s.addText(it[0], {
      x: x + 0.3, y: 2.92, w: cw - 0.6, h: 0.28,
      fontFace: FS, fontSize: 10, bold: true, color: it[3], charSpacing: 1, margin: 0,
    });
    s.addText(it[1], {
      x: x + 0.3, y: 3.28, w: cw - 0.6, h: 0.9,
      fontFace: FS, fontSize: 19, bold: true, color: INK, margin: 0, valign: "top",
    });
    s.addText(it[2], {
      x: x + 0.3, y: 4.24, w: cw - 0.6, h: 0.9,
      fontFace: FS, fontSize: 13, color: MUTED, margin: 0, valign: "top",
    });
  });
  s.addText("四個階段都用同一個建案，只多一個新東西 —— 不用每次重新理解情境。", {
    x: M, y: 5.72, w: W, h: 0.4, fontFace: FS, fontSize: 14, color: INK_2, margin: 0,
  });
  note(s, "強調：四階段是疊加，不是四個獨立主題。每一階段只多學一件事。");
}

/* 3 期待管理 */
{
  const s = light();
  eyebrow(s, "期待管理", BRASS);
  title(s, "先說清楚：今天不教你寫程式");
  const rows = [
    ["今天教的是", "怎麼把「你要什麼」講清楚給 AI 聽", GOOD],
    ["今天不教的是", "HTML、CSS、JavaScript 的語法", MUTED],
    ["你需要準備的是", "一份自己部門的資料，或一個想做的東西", BRASS],
  ];
  rows.forEach((r, i) => {
    const y = 2.15 + i * 1.18;
    card(s, M, y, W, 0.98);
    s.addText(r[0], {
      x: M + 0.4, y: y, w: 2.6, h: 0.98,
      fontFace: FS, fontSize: 11, bold: true, color: r[2], margin: 0, valign: "middle",
    });
    s.addText(r[1], {
      x: M + 3.2, y: y, w: W - 3.6, h: 0.98,
      fontFace: FS, fontSize: 18, color: INK, margin: 0, valign: "middle",
    });
  });
  s.addText("會下指令的人，不需要會寫程式，也做得出東西。", {
    x: M, y: 5.9, w: W, h: 0.45, fontFace: FS, fontSize: 17,
    bold: true, color: BRASS, margin: 0,
  });
  note(s, "這頁是全場最重要的期待管理。非資訊背景學員最怕的就是「看不懂 code」，先把這個焦慮拆掉。");
}

/* 4 時間分配 */
{
  const s = light();
  eyebrow(s, "時間分配", BRASS);
  title(s, "三小時怎麼跑");
  const steps = [
    ["00:00", "45 min", "資訊網站", "學會下指令，做出第一個網頁", BRASS],
    ["00:45", "35 min", "加上圖片", "把自己的照片放上去", BRASS],
    ["01:20", "50 min", "靜態儀表板", "把 Excel 變成圖表", BLUE],
    ["02:10", "50 min", "可互動儀表板", "換資料，畫面跟著算", BLUE],
  ];
  steps.forEach((st, i) => {
    const y = 2.1 + i * 0.98;
    s.addText(st[0], {
      x: M, y: y, w: 1.05, h: 0.8, fontFace: FS, fontSize: 15,
      bold: true, color: st[4], margin: 0, valign: "middle",
    });
    s.addShape(pres.ShapeType.rect, {
      x: M + 1.25, y: y + 0.3, w: 0.2, h: 0.2, fill: { color: st[4] },
    });
    s.addText(st[2], {
      x: M + 1.75, y: y, w: 3.1, h: 0.8, fontFace: FS, fontSize: 19,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(st[3], {
      x: M + 4.95, y: y, w: 5.3, h: 0.8, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "middle",
    });
    s.addText(st[1], {
      x: M + 10.4, y: y, w: 1.6, h: 0.8, fontFace: FS, fontSize: 12,
      color: MUTED, align: "right", margin: 0, valign: "middle",
    });
    if (i < 3) {
      s.addShape(pres.ShapeType.line, {
        x: M, y: y + 0.88, w: W, h: 0, line: { color: RULE, width: 0.75 },
      });
    }
  });
  s.addText("每一段都是：先看範例 → 講觀念 → 自己動手", {
    x: M, y: 6.2, w: W, h: 0.4, fontFace: FS, fontSize: 14, color: INK_2, margin: 0,
  });
  note(s, "時間僅供參考，動手時間不夠就砍講解，不要砍動手。學員的成就感全部來自自己做出來那一刻。");
}

/* ══════════════════ 分隔頁工廠 ══════════════════ */
function divider(no, name, line1, line2, accent) {
  const s = pres.addSlide();
  darkBg(s);
  s.addText(no, {
    x: M, y: 2.35, w: 4, h: 0.4, fontFace: FS, fontSize: 13,
    bold: true, color: accent, charSpacing: 4, margin: 0,
  });
  s.addText(name, {
    x: M, y: 2.9, w: 11, h: 1.05, fontFace: FS, fontSize: 44,
    bold: true, color: "F2F5F4", margin: 0, valign: "middle",
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: 4.22, w: 2.6, h: 0, line: { color: accent, width: 1.5 },
  });
  s.addText(line1, {
    x: M, y: 4.45, w: 10, h: 0.42, fontFace: FS, fontSize: 18,
    color: "D6DDDB", margin: 0,
  });
  s.addText(line2, {
    x: M, y: 4.95, w: 10, h: 0.42, fontFace: FS, fontSize: 14,
    color: MUTED, margin: 0,
  });
  return s;
}

/* ══════════════════ STAGE 01 ══════════════════ */
divider("第一段", "資訊網站", "把一段描述變成一個網頁",
        "這一段的重點是「怎麼說」，不是「怎麼寫」", BRASS_LT);

/* 6 成品 */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addImage({ path: "p1.jpg", x: 5.983, y: 0, w: 7.35, h: 7.5 });
  s.addText("成品", {
    x: M, y: 1.5, w: 4.6, h: 0.32, fontFace: FS, fontSize: 11,
    bold: true, color: BRASS, charSpacing: 2, margin: 0,
  });
  s.addText("這是三十分鐘做出來的", {
    x: M, y: 1.95, w: 4.7, h: 1.5, fontFace: FS, fontSize: 30,
    bold: true, color: INK, margin: 0, valign: "top",
  });
  s.addText([
    { text: "一整份建案介紹網頁：建案名、生活尺度、", options: { breakLine: true } },
    { text: "理念、規劃特色、格局表、聯絡方式。", options: { breakLine: true } },
    { text: "" , options: { breakLine: true } },
    { text: "從頭到尾沒有寫一行程式碼。" , options: { bold: true } },
  ], {
    x: M, y: 3.5, w: 4.7, h: 2.0, fontFace: FS, fontSize: 15,
    color: INK_2, lineSpacing: 26, margin: 0, valign: "top",
  });
  note(s, "先給成品，再講怎麼做。學員看到具體目標才有動力聽觀念。");
}

/* 7 核心觀念 */
{
  const s = light();
  eyebrow(s, "核心觀念", BRASS);
  title(s, "你不是在寫程式，是在描述需求");
  card(s, M, 2.15, W, 1.5, { fill: "FFFFFF" });
  s.addText("「幫我做一個建案介紹網頁，建案叫沐森，在台中七期，主打面湖景觀，要有格局表和生活機能。」", {
    x: M + 0.5, y: 2.15, w: W - 1.0, h: 1.5, fontFace: FS, fontSize: 19,
    color: INK, margin: 0, valign: "middle", lineSpacing: 32,
  });
  s.addText("這句話你本來就會說 —— 你每天跟同事、跟廣告商講的就是這個。", {
    x: M, y: 3.95, w: W, h: 0.45, fontFace: FS, fontSize: 16, color: BRASS, bold: true, margin: 0,
  });
  const pts = [
    ["你已經有的能力", "清楚描述一個建案該有什麼"],
    ["今天要練的", "把這個描述講得讓 AI 接得住"],
  ];
  pts.forEach((p, i) => {
    const x = M + i * (5.9 + 0.25);
    card(s, x, 4.65, 5.9, 1.55);
    s.addText(p[0], {
      x: x + 0.4, y: 4.9, w: 5.1, h: 0.32, fontFace: FS, fontSize: 10,
      bold: true, color: MUTED, charSpacing: 1, margin: 0,
    });
    s.addText(p[1], {
      x: x + 0.4, y: 5.3, w: 5.1, h: 0.7, fontFace: FS, fontSize: 17,
      color: INK, margin: 0, valign: "top",
    });
  });
  note(s, "讓學員唸一次那句話。他們會發現這根本不是技術語言，就是日常工作語言。");
}

/* 8 壞指令 vs 好指令 */
{
  const s = light();
  eyebrow(s, "對照", BRASS);
  title(s, "同樣一件事，兩種說法");
  const cw = 5.9;
  // 壞
  card(s, M, 2.15, cw, 4.0);
  s.addText("太模糊", {
    x: M + 0.4, y: 2.42, w: cw - 0.8, h: 0.36, fontFace: FS, fontSize: 15,
    bold: true, color: CRIT, margin: 0,
  });
  s.addText("「幫我做一個建案網站」", {
    x: M + 0.4, y: 2.9, w: cw - 0.8, h: 0.6, fontFace: FS, fontSize: 18,
    color: INK, margin: 0, valign: "top",
  });
  s.addText([
    { text: "AI 只能自己猜：什麼建案、賣點在哪、", options: { breakLine: true } },
    { text: "要放哪些內容、給誰看。", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "結果就是一個很通用、跟你的案子", options: { breakLine: true } },
    { text: "沒什麼關係的網頁。", options: {} },
  ], {
    x: M + 0.4, y: 3.7, w: cw - 0.8, h: 2.1, fontFace: FS, fontSize: 14,
    color: MUTED, lineSpacing: 24, margin: 0, valign: "top",
  });
  // 好
  const x2 = M + cw + 0.25;
  card(s, x2, 2.15, cw, 4.0, { line: BRASS, lw: 1.5 });
  s.addText("講清楚", {
    x: x2 + 0.4, y: 2.42, w: cw - 0.8, h: 0.36, fontFace: FS, fontSize: 15,
    bold: true, color: GOOD, margin: 0,
  });
  s.addText("「沐森，台中七期，低層電梯宅，主打面湖景觀。要有建案理念、格局表、生活機能、聯絡方式。調性沉靜有質感，不要像廣告傳單。」", {
    x: x2 + 0.4, y: 2.9, w: cw - 0.8, h: 2.1, fontFace: FS, fontSize: 15,
    color: INK, margin: 0, valign: "top", lineSpacing: 25,
  });
  s.addText("多花三十秒描述，省下反覆重做的時間。", {
    x: x2 + 0.4, y: 5.4, w: cw - 0.8, h: 0.5, fontFace: FS, fontSize: 14,
    bold: true, color: BRASS, margin: 0, valign: "top",
  });
  note(s, "現場可以真的跑一次左邊的指令，讓學員看到通用結果有多空洞，再跑右邊。對比最有說服力。");
}

/* 9 好指令的四個要素 */
{
  const s = light();
  eyebrow(s, "檢查清單", BRASS);
  title(s, "一個好指令，講滿這四件事");
  const four = [
    ["01", "做什麼", "一頁建案介紹網站"],
    ["02", "關於誰", "沐森，台中七期，低層電梯宅"],
    ["03", "要有什麼", "理念、特色、格局表、生活機能、聯絡"],
    ["04", "什麼調性", "沉靜有質感，不要像廣告傳單"],
  ];
  four.forEach((f, i) => {
    const y = 2.15 + i * 1.05;
    s.addShape(pres.ShapeType.rect, {
      x: M, y: y, w: 0.52, h: 0.52, fill: { color: BRASS },
    });
    s.addText(f[0], {
      x: M, y: y, w: 0.52, h: 0.52, fontFace: FS, fontSize: 12,
      bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    s.addText(f[1], {
      x: M + 0.85, y: y, w: 2.5, h: 0.52, fontFace: FS, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(f[2], {
      x: M + 3.5, y: y, w: 8.4, h: 0.52, fontFace: FS, fontSize: 15,
      color: MUTED, margin: 0, valign: "middle",
    });
  });
  card(s, M, 6.35, W, 0.72, { fill: "FFFFFF" });
  s.addText("四件事都講到，AI 第一次給的結果通常就有八成可用。", {
    x: M + 0.4, y: 6.35, w: W - 0.8, h: 0.72, fontFace: FS, fontSize: 15,
    bold: true, color: INK, margin: 0, valign: "middle",
  });
  note(s, "請學員拿紙筆，把自己的案子按這四點寫下來。寫完直接貼進 AI，這就是第一次實作。");
}

/* 10 動手練習 */
{
  const s = light();
  eyebrow(s, "動手 · 25 分鐘", BRASS);
  title(s, "換你說說看");
  card(s, M, 2.1, W, 3.05, { fill: "FFFFFF" });
  s.addText("把這段填完，直接貼給 AI", {
    x: M + 0.5, y: 2.35, w: W - 1.0, h: 0.35, fontFace: FS, fontSize: 11,
    bold: true, color: BRASS, charSpacing: 1, margin: 0,
  });
  s.addText([
    { text: "幫我做一個一頁式的", options: {} },
    { text: "「＿＿＿＿」", options: { color: BRASS, bold: true } },
    { text: "介紹網頁。", options: { breakLine: true } },
    { text: "它在", options: {} },
    { text: "＿＿＿＿", options: { color: BRASS, bold: true } },
    { text: "，主打的是", options: {} },
    { text: "＿＿＿＿", options: { color: BRASS, bold: true } },
    { text: "。", options: { breakLine: true } },
    { text: "網頁要有：", options: {} },
    { text: "＿＿＿＿、＿＿＿＿、＿＿＿＿", options: { color: BRASS, bold: true } },
    { text: "。", options: { breakLine: true } },
    { text: "調性希望是", options: {} },
    { text: "＿＿＿＿", options: { color: BRASS, bold: true } },
    { text: "。", options: {} },
  ], {
    x: M + 0.5, y: 2.85, w: W - 1.0, h: 2.0, fontFace: FS, fontSize: 19,
    color: INK, lineSpacing: 40, margin: 0, valign: "top",
  });
  s.addText("不一定要是建案 —— 部門介紹、活動報名、產品說明都可以。挑一個你這週真的用得到的。", {
    x: M, y: 5.45, w: W, h: 0.5, fontFace: FS, fontSize: 15, color: INK_2, margin: 0,
  });
  s.addText("做出來之後，先別急著調細節。整體對了再修。", {
    x: M, y: 6.05, w: W, h: 0.5, fontFace: FS, fontSize: 15,
    bold: true, color: BRASS, margin: 0,
  });
  note(s, "巡場重點：卡住的人多半是「不知道要做什麼」，不是不會下指令。直接幫他挑一個題目。");
}

/* 11 不滿意怎麼辦 */
{
  const s = light();
  eyebrow(s, "卡住的時候", BRASS);
  title(s, "AI 給的不是你要的？");
  const tips = [
    ["不要重寫，用改的", "「其他都好，只要把格局表改成四欄」——AI 會只動那一塊"],
    ["說「不要什麼」", "「顏色太花，改成低彩度」比「弄好看一點」有用得多"],
    ["給它參考", "「像 ○○ 那種質感」——有具體對象，AI 抓得比較準"],
  ];
  tips.forEach((t, i) => {
    const y = 2.15 + i * 1.42;
    card(s, M, y, W, 1.22);
    s.addShape(pres.ShapeType.rect, {
      x: M + 0.42, y: y + 0.45, w: 0.28, h: 0.28, fill: { color: BRASS },
    });
    s.addText(t[0], {
      x: M + 1.0, y: y + 0.16, w: 4.2, h: 0.45, fontFace: FS, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(t[1], {
      x: M + 1.0, y: y + 0.62, w: 10.3, h: 0.48, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "top",
    });
  });
  note(s, "第一點最重要。學員常以為不滿意就要整個重來，其實對話可以持續修正——這也是 Stage 02 的鋪陳。");
}

/* ══════════════════ STAGE 02 ══════════════════ */
divider("第二段", "加上自己的圖片", "把你部門拍的照片放上去",
        "重點：網頁跟圖片檔，是怎麼連在一起的", BRASS_LT);

/* 13 成品 */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addImage({ path: "p2.jpg", x: 0, y: 0, w: 7.35, h: 7.5 });
  s.addText("成品", {
    x: 7.9, y: 1.5, w: 4.8, h: 0.32, fontFace: FS, fontSize: 11,
    bold: true, color: BRASS, charSpacing: 2, margin: 0,
  });
  s.addText("同一個網站，\n換上自己的照片", {
    x: 7.9, y: 1.95, w: 4.9, h: 1.7, fontFace: FS, fontSize: 28, lineSpacing: 42,
    bold: true, color: INK, margin: 0, valign: "top",
  });
  s.addText([
    { text: "文字、版面完全沒動。", options: { breakLine: true } },
    { text: "只是多了資料夾裡的圖片檔。", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "從這一刻起，它不再是範例，", options: { breakLine: true } },
    { text: "而是你的東西。", options: { bold: true } },
  ], {
    x: 7.9, y: 3.65, w: 4.9, h: 2.2, fontFace: FS, fontSize: 15,
    color: INK_2, lineSpacing: 26, margin: 0, valign: "top",
  });
  note(s, "這一頁的情緒點是「擁有感」。放自己拍的工地照片，學員的態度會明顯不一樣。");
}

/* 14 核心觀念 */
{
  const s = light();
  eyebrow(s, "核心觀念", BRASS);
  title(s, "網頁不會把圖片存在裡面");
  card(s, M, 2.2, W, 1.55, { fill: INK, line: INK });
  s.addText("它只記住一件事：這張圖放在哪、叫什麼名字。", {
    x: M + 0.5, y: 2.2, w: W - 1.0, h: 1.55, fontFace: FS, fontSize: 25,
    bold: true, color: "F2F5F4", margin: 0, valign: "middle",
  });
  s.addText("所以換照片，其實不用改網頁 —— 把檔案換掉就好。", {
    x: M, y: 4.05, w: W, h: 0.5, fontFace: FS, fontSize: 18,
    bold: true, color: BRASS, margin: 0,
  });
  const two = [
    ["網頁檔", "index.html", "決定版面長什麼樣、文字寫什麼"],
    ["圖片資料夾", "images/", "放照片，網頁到這裡來拿"],
  ];
  two.forEach((t, i) => {
    const x = M + i * (5.9 + 0.25);
    card(s, x, 4.85, 5.9, 1.75);
    s.addText(t[0], {
      x: x + 0.42, y: 5.08, w: 5.1, h: 0.36, fontFace: FS, fontSize: 16,
      bold: true, color: INK, margin: 0,
    });
    s.addText(t[1], {
      x: x + 0.42, y: 5.5, w: 5.1, h: 0.34, fontFace: FS, fontSize: 14,
      bold: true, color: BRASS, margin: 0,
    });
    s.addText(t[2], {
      x: x + 0.42, y: 5.92, w: 5.1, h: 0.5, fontFace: FS, fontSize: 13,
      color: MUTED, margin: 0, valign: "top",
    });
  });
  note(s, "這句話是整個 Stage 02 的核心。講完先停一下，讓它沉下去再往下走。");
}

/* 15 資料夾結構 */
{
  const s = light();
  eyebrow(s, "資料夾結構", BRASS);
  title(s, "資料夾長什麼樣");
  card(s, M, 2.15, 6.6, 3.5, { fill: "FFFFFF" });
  s.addText([
    { text: "沐森網站/", options: { breakLine: true, bold: true, color: INK } },
    { text: "├─ index.html          ", options: { color: MUTED } },
    { text: "網頁", options: { breakLine: true, color: BRASS } },
    { text: "└─ images/             ", options: { color: MUTED } },
    { text: "圖片資料夾", options: { breakLine: true, color: BRASS } },
    { text: "     ├─ hero.jpg       ", options: { color: MUTED } },
    { text: "首圖", options: { breakLine: true, color: BRASS } },
    { text: "     ├─ interior-1.jpg ", options: { color: MUTED } },
    { text: "客廳", options: { breakLine: true, color: BRASS } },
    { text: "     ├─ interior-2.jpg ", options: { color: MUTED } },
    { text: "主臥", options: { breakLine: true, color: BRASS } },
    { text: "     └─ interior-3.jpg ", options: { color: MUTED } },
    { text: "大廳", options: { color: BRASS } },
  ], {
    x: M + 0.42, y: 2.45, w: 6.0, h: 2.9, fontFace: FM, fontSize: 13,
    lineSpacing: 24, margin: 0, valign: "top",
  });
  card(s, M + 6.85, 2.15, 5.18, 3.5);
  s.addText("兩件事，分開放", {
    x: M + 7.25, y: 2.5, w: 4.4, h: 0.42, fontFace: FS, fontSize: 19,
    bold: true, color: INK, margin: 0,
  });
  s.addText([
    { text: "網頁是一個檔案，照片是另外幾個檔案。", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "它們必須放在正確的相對位置，網頁才找得到照片。", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "整包資料夾一起複製、一起搬，就不會出問題。", options: {} },
  ], {
    x: M + 7.25, y: 3.05, w: 4.4, h: 2.4, fontFace: FS, fontSize: 14,
    color: MUTED, lineSpacing: 23, margin: 0, valign: "top",
  });
  s.addText("寄給同事的時候，要整個資料夾一起寄 —— 只寄 index.html 過去，圖會全部破掉。", {
    x: M, y: 5.95, w: W, h: 0.5, fontFace: FS, fontSize: 15,
    bold: true, color: WARN, margin: 0,
  });
  note(s, "最後這句是實務上最常出包的地方，一定要講。");
}

/* 16 src 拆解 */
{
  const s = light();
  eyebrow(s, "網頁裡是這樣寫的", BRASS);
  title(s, "看懂這一行就夠了");
  card(s, M, 2.15, W, 1.32, { fill: INK, line: INK });
  s.addText([
    { text: "<img src=", options: { color: MUTED_LT } },
    { text: '"images/hero.jpg"', options: { color: BRASS_LT, bold: true } },
    { text: " />", options: { color: MUTED_LT } },
  ], {
    x: M + 0.5, y: 2.15, w: W - 1.0, h: 1.32, fontFace: FM, fontSize: 26,
    margin: 0, valign: "middle",
  });
  const parts = [
    ["src", "「這張圖在哪裡」", MUTED],
    ["images/", "同一層有個叫 images 的資料夾", BRASS],
    ["hero.jpg", "資料夾裡有個叫 hero.jpg 的檔案", BRASS],
  ];
  parts.forEach((p, i) => {
    const y = 3.8 + i * 0.92;
    s.addText(p[0], {
      x: M + 0.5, y: y, w: 2.3, h: 0.6, fontFace: FM, fontSize: 16,
      bold: true, color: p[2], margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: M + 3.0, y: y, w: 8.5, h: 0.6, fontFace: FS, fontSize: 17,
      color: INK, margin: 0, valign: "middle",
    });
    s.addShape(pres.ShapeType.line, {
      x: M + 0.5, y: y + 0.72, w: W - 1.0, h: 0, line: { color: RULE, width: 0.75 },
    });
  });
  s.addText("那張圖是湖景還是你昨天在工地拍的，網頁不在意 —— 名字對就顯示。", {
    x: M, y: 6.5, w: W, h: 0.45, fontFace: FS, fontSize: 15, color: BRASS, bold: true, margin: 0,
  });
  note(s, "不用教 HTML 語法，只要讓他們認得 src 這個字，以及斜線代表「資料夾裡面」。");
}

/* 17 兩種換圖做法 */
{
  const s = light();
  eyebrow(s, "操作步驟", BRASS);
  title(s, "換成自己的照片，兩種做法");
  const cw = 5.9;
  const ways = [
    ["做法一", "改檔名", ["把你的照片改名成 hero.jpg", "複製進 images 資料夾蓋掉原本那張", "回瀏覽器按重新整理"], "最簡單，完全不用碰程式碼", GOOD],
    ["做法二", "改網頁裡的檔名", ["照片放進 images，不用改名", "把網頁裡的 hero.jpg 改成你的檔名", "存檔，重新整理"], "照片多的時候比較好管理", BRASS],
  ];
  ways.forEach((w, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 2.15, cw, 3.75);
    s.addText(w[0], {
      x: x + 0.42, y: 2.42, w: cw - 0.84, h: 0.3, fontFace: FS, fontSize: 11,
      bold: true, color: w[4], charSpacing: 1, margin: 0,
    });
    s.addText(w[1], {
      x: x + 0.42, y: 2.78, w: cw - 0.84, h: 0.5, fontFace: FS, fontSize: 22,
      bold: true, color: INK, margin: 0,
    });
    w[2].forEach((step, j) => {
      const sy = 3.45 + j * 0.6;
      s.addText(String(j + 1), {
        x: x + 0.42, y: sy, w: 0.32, h: 0.32, fontFace: FS, fontSize: 11,
        bold: true, color: w[4], align: "center", valign: "middle", margin: 0,
      });
      s.addText(step, {
        x: x + 0.85, y: sy - 0.04, w: cw - 1.3, h: 0.42, fontFace: FS, fontSize: 14,
        color: INK_2, margin: 0, valign: "middle",
      });
    });
    s.addText(w[3], {
      x: x + 0.42, y: 5.3, w: cw - 0.84, h: 0.42, fontFace: FS, fontSize: 13,
      color: w[4], bold: true, margin: 0,
    });
  });
  card(s, M, 6.15, W, 0.85, { fill: "FFFFFF" });
  s.addText("不會改？把需求講給 AI 聽就好：「照片放進 images 了，檔名是 工地實景.jpg，幫我把首圖換成這張」", {
    x: M + 0.42, y: 6.15, w: W - 0.84, h: 0.85, fontFace: FS, fontSize: 15,
    bold: true, color: INK, margin: 0, valign: "middle",
  });
  note(s, "最後那句要用力講——它把 Stage 02 扣回課程主軸：不會做就描述需求。");
}

/* 18 破圖檢查 */
{
  const s = light();
  eyebrow(s, "排除問題", BRASS);
  title(s, "圖片破掉？先查這三件事");
  const checks = [
    ["檔名有沒有完全一樣", "Hero.jpg 和 hero.jpg 可能被當成兩個不同檔案，大小寫要對"],
    ["副檔名對不對", "手機照片常是 .jpeg 或 .HEIC，不是 .jpg —— 要嘛改網頁，要嘛轉存"],
    ["真的放進 images 了嗎", "放在資料夾外面、或多包了一層，網頁都會找不到"],
  ];
  checks.forEach((c, i) => {
    const y = 2.15 + i * 1.42;
    card(s, M, y, W, 1.22);
    s.addShape(pres.ShapeType.rect, {
      x: M + 0.42, y: y + 0.4, w: 0.4, h: 0.4, fill: { color: WARN },
    });
    s.addText(String(i + 1), {
      x: M + 0.42, y: y + 0.4, w: 0.4, h: 0.4, fontFace: FS, fontSize: 12,
      bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    s.addText(c[0], {
      x: M + 1.1, y: y + 0.16, w: 5.0, h: 0.45, fontFace: FS, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(c[1], {
      x: M + 1.1, y: y + 0.62, w: 10.2, h: 0.48, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "top",
    });
  });
  s.addText("電腦是照字面找檔案的 —— 差一個字母就是找不到。", {
    x: M, y: 6.5, w: W, h: 0.45, fontFace: FS, fontSize: 15, color: INK_2, margin: 0,
  });
  note(s, "建議現場故意打錯一次檔名，讓學員看到破圖，再修好。體驗過一次比講三遍有用。");
}

/* ══════════════════ STAGE 03 ══════════════════ */
divider("第三段", "靜態儀表板", "把 Excel 進度表變成看得懂的圖表",
        "從「做內容」換到「處理資料」", BLUE_LT);

/* 20 成品 */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addImage({ path: "p3.jpg", x: 5.983, y: 0, w: 7.35, h: 7.5 });
  s.addText("成品", {
    x: M, y: 1.5, w: 4.6, h: 0.32, fontFace: FS, fontSize: 11,
    bold: true, color: BLUE, charSpacing: 2, margin: 0,
  });
  s.addText("一張表格，\n變成一眼看懂的畫面", {
    x: M, y: 1.95, w: 4.7, h: 1.8, fontFace: FS, fontSize: 28, lineSpacing: 42,
    bold: true, color: INK, margin: 0, valign: "top",
  });
  s.addText([
    { text: "同樣七個案場的數字。", options: { breakLine: true } },
    { text: "在 Excel 裡要一格一格看，", options: { breakLine: true } },
    { text: "在這裡三秒就知道哪兩個落後。", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "這就是儀表板存在的理由。", options: { bold: true } },
  ], {
    x: M, y: 3.85, w: 4.7, h: 2.3, fontFace: FS, fontSize: 15,
    color: INK_2, lineSpacing: 26, margin: 0, valign: "top",
  });
  note(s, "可以先投影一份真的 Excel，再切到這頁。落差感很強。");
}

/* 21 觀念切換 */
{
  const s = light();
  eyebrow(s, "換一種任務", BLUE);
  title(s, "這一段換了一種任務");
  const cw = 5.9;
  const cols = [
    ["前兩階段", "做內容", ["你決定要放什麼字、什麼圖", "AI 幫你排版", "好不好看是主觀的"], BRASS],
    ["這一階段", "處理資料", ["資料是既有的，不能亂改", "AI 幫你決定怎麼呈現", "對不對是客觀的"], BLUE],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 2.15, cw, 3.4, i === 1 ? { line: BLUE, lw: 1.5 } : {});
    s.addText(c[0], {
      x: x + 0.42, y: 2.42, w: cw - 0.84, h: 0.3, fontFace: FS, fontSize: 11,
      bold: true, color: c[3], charSpacing: 1, margin: 0,
    });
    s.addText(c[1], {
      x: x + 0.42, y: 2.78, w: cw - 0.84, h: 0.55, fontFace: FS, fontSize: 24,
      bold: true, color: INK, margin: 0,
    });
    c[2].forEach((t, j) => {
      s.addText(t, {
        x: x + 0.42, y: 3.55 + j * 0.6, w: cw - 0.84, h: 0.45,
        fontFace: FS, fontSize: 15, color: INK_2, margin: 0, valign: "middle",
        bullet: { code: "25AA" },
      });
    });
  });
  card(s, M, 5.85, W, 0.95, { fill: "FFFFFF" });
  s.addText("好消息：資料本來就在你手上。你每週在填的那張進度表，就是原料。", {
    x: M + 0.42, y: 5.85, w: W - 0.84, h: 0.95, fontFace: FS, fontSize: 16,
    bold: true, color: INK, margin: 0, valign: "middle",
  });
  note(s, "很多學員會覺得「資料」是個很技術的詞。要把它拉回「你每週在填的那張表」。");
}

/* 22 好資料 vs 壞資料 */
{
  const s = light();
  eyebrow(s, "關於資料", BLUE);
  title(s, "AI 看得懂的表格長什麼樣");
  const cw = 5.9;
  // 壞
  card(s, M, 2.15, cw, 3.9);
  s.addText("這種 AI 會做不好", {
    x: M + 0.42, y: 2.42, w: cw - 0.84, h: 0.38, fontFace: FS, fontSize: 16,
    bold: true, color: CRIT, margin: 0,
  });
  [
    "標題列中間有空白行",
    "合併儲存格",
    "一格裡塞好幾個資訊",
    "數字後面帶單位（68%、180戶）",
    "同一欄有中文也有數字",
  ].forEach((t, i) => {
    s.addText(t, {
      x: M + 0.42, y: 2.95 + i * 0.56, w: cw - 0.84, h: 0.45,
      fontFace: FS, fontSize: 14, color: INK_2, margin: 0, valign: "middle",
      bullet: { code: "25AA" },
    });
  });
  // 好
  const x2 = M + cw + 0.25;
  card(s, x2, 2.15, cw, 3.9, { line: BLUE, lw: 1.5 });
  s.addText("這種一次就成", {
    x: x2 + 0.42, y: 2.42, w: cw - 0.84, h: 0.38, fontFace: FS, fontSize: 16,
    bold: true, color: GOOD, margin: 0,
  });
  [
    "第一列就是欄位名稱",
    "一格一個資訊",
    "數字就是純數字，不帶單位",
    "沒有合併儲存格、沒有空行",
    "每一列代表一個案場",
  ].forEach((t, i) => {
    s.addText(t, {
      x: x2 + 0.42, y: 2.95 + i * 0.56, w: cw - 0.84, h: 0.45,
      fontFace: FS, fontSize: 14, color: INK_2, margin: 0, valign: "middle",
      bullet: { code: "25AA" },
    });
  });
  s.addText("整理資料花的十分鐘，比事後修圖表的一小時划算。", {
    x: M, y: 6.35, w: W, h: 0.5, fontFace: FS, fontSize: 16,
    bold: true, color: BLUE, margin: 0,
  });
  note(s, "這頁是 Stage 03 最實用的內容。多數人做不出好圖表，問題出在資料，不在 AI。");
}

/* 23 欄位對應 */
{
  const s = light();
  eyebrow(s, "欄位對應", BLUE);
  title(s, "六個欄位，變成畫面上的什麼");
  const maps = [
    ["案名", "表格第一欄", MUTED],
    ["階段", "表格第二欄", MUTED],
    ["完成率", "進度條的長度 · 顏色深淺", BLUE],
    ["已售戶數 / 總戶數", "上方的銷售率大數字", BLUE],
    ["狀態", "右邊的標籤（正常 / 落後 / 即將完工）", BLUE],
  ];
  maps.forEach((m, i) => {
    const y = 2.15 + i * 0.86;
    s.addText(m[0], {
      x: M, y: y, w: 3.6, h: 0.62, fontFace: FS, fontSize: 17,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText("→", {
      x: M + 3.7, y: y, w: 0.5, h: 0.62, fontFace: FS, fontSize: 15,
      color: m[2], margin: 0, valign: "middle",
    });
    s.addText(m[1], {
      x: M + 4.3, y: y, w: 7.6, h: 0.62, fontFace: FS, fontSize: 16,
      color: m[2], margin: 0, valign: "middle",
    });
    s.addShape(pres.ShapeType.line, {
      x: M, y: y + 0.7, w: W, h: 0, line: { color: RULE, width: 0.75 },
    });
  });
  s.addText("你不用決定「怎麼畫」，只要決定「哪些欄位重要」。", {
    x: M, y: 6.55, w: W, h: 0.45, fontFace: FS, fontSize: 16,
    bold: true, color: BLUE, margin: 0,
  });
  note(s, "帶學員看一次自己的表：哪一欄是要被強調的？那就是 KPI。");
}

/* 24 實作 */
{
  const s = light();
  eyebrow(s, "動手 · 30 分鐘", BLUE);
  title(s, "把你部門的表格丟進去");
  const steps = [
    ["01", "挑一張表", "你每週或每月都要看的那張，不用太大，十列以內就夠"],
    ["02", "整理乾淨", "照前一頁的原則：一格一個資訊、數字不帶單位"],
    ["03", "描述你要看什麼", "「幫我做成儀表板，最上面要看到平均進度和落後幾個」"],
    ["04", "看結果，再修", "顏色、排序、要強調哪一欄，都可以再對話調整"],
  ];
  steps.forEach((st, i) => {
    const y = 2.15 + i * 1.1;
    s.addShape(pres.ShapeType.rect, {
      x: M, y: y, w: 0.52, h: 0.52, fill: { color: BLUE },
    });
    s.addText(st[0], {
      x: M, y: y, w: 0.52, h: 0.52, fontFace: FS, fontSize: 12,
      bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    s.addText(st[1], {
      x: M + 0.85, y: y, w: 3.0, h: 0.52, fontFace: FS, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(st[2], {
      x: M + 4.0, y: y, w: 7.9, h: 0.52, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "middle",
    });
  });
  s.addText("沒有現成資料的人，用主辦提供的建案範例檔就好。", {
    x: M, y: 6.55, w: W, h: 0.45, fontFace: FS, fontSize: 15, color: INK_2, margin: 0,
  });
  note(s, "務必準備好範例 CSV，一定會有人沒帶資料。");
}

/* 25 這版的限制 */
{
  const s = light();
  eyebrow(s, "這一版的限制", BLUE);
  title(s, "但這一版有個麻煩");
  card(s, M, 2.2, W, 1.5, { fill: INK, line: INK });
  s.addText("數字是寫死在網頁裡的。資料一變，就要整個重做一次。", {
    x: M + 0.5, y: 2.2, w: W - 1.0, h: 1.5, fontFace: FS, fontSize: 24,
    bold: true, color: "F2F5F4", margin: 0, valign: "middle",
  });
  const pain = [
    ["這週", "把本週數字給 AI，生出一份網頁"],
    ["下週", "數字更新了 —— 再給一次，再生一份"],
    ["下下週", "再一次⋯⋯"],
  ];
  pain.forEach((p, i) => {
    const y = 4.05 + i * 0.78;
    s.addText(p[0], {
      x: M, y: y, w: 1.7, h: 0.6, fontFace: FS, fontSize: 13,
      bold: true, color: WARN, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: M + 1.9, y: y, w: 10.0, h: 0.6, fontFace: FS, fontSize: 16,
      color: INK_2, margin: 0, valign: "middle",
    });
  });
  s.addText("下一段就是要解決這件事。", {
    x: M, y: 6.5, w: W, h: 0.45, fontFace: FS, fontSize: 17,
    bold: true, color: BLUE, margin: 0,
  });
  note(s, "刻意讓學員感受到痛點，Stage 04 的價值才會被看見。");
}

/* ══════════════════ STAGE 04 ══════════════════ */
divider("第四段", "可互動儀表板", "換一份資料，畫面自己重算",
        "順便看清楚：這種做法的邊界在哪", BLUE_LT);

/* 27 成品 */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addImage({ path: "p4.jpg", x: 0, y: 0, w: 7.35, h: 7.5 });
  s.addText("成品", {
    x: 7.9, y: 1.5, w: 4.8, h: 0.32, fontFace: FS, fontSize: 11,
    bold: true, color: BLUE, charSpacing: 2, margin: 0,
  });
  s.addText("多了一顆「上傳 CSV」", {
    x: 7.9, y: 1.95, w: 4.9, h: 1.1, fontFace: FS, fontSize: 28,
    bold: true, color: INK, margin: 0, valign: "top",
  });
  s.addText([
    { text: "選一個檔案，所有數字、進度條、", options: { breakLine: true } },
    { text: "標籤全部重算。", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "下週要更新報表，換檔案就好，", options: { breakLine: true } },
    { text: "不用再找 AI 重做一次。", options: { bold: true } },
  ], {
    x: 7.9, y: 3.25, w: 4.9, h: 2.4, fontFace: FS, fontSize: 15,
    color: INK_2, lineSpacing: 26, margin: 0, valign: "top",
  });
  note(s, "現場一定要示範上傳一次。看到數字整排跳動的瞬間，是全場反應最大的地方。");
}

/* 28 三 vs 四 */
{
  const s = light();
  eyebrow(s, "對照", BLUE);
  title(s, "同一份儀表板，差在哪");
  const cw = 5.9;
  const cmp = [
    ["第三段", "數字寫死", ["更新資料 = 重做一份網頁", "適合一次性的簡報", "檔案會愈積愈多"], MUTED],
    ["第四段", "可以換資料", ["更新資料 = 換一個檔案", "適合每週固定要出的報表", "網頁只要一份"], BLUE],
  ];
  cmp.forEach((c, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 2.15, cw, 3.3, i === 1 ? { line: BLUE, lw: 1.5 } : {});
    s.addText(c[0], {
      x: x + 0.42, y: 2.42, w: cw - 0.84, h: 0.3, fontFace: FS, fontSize: 11,
      bold: true, color: c[3], charSpacing: 1, margin: 0,
    });
    s.addText(c[1], {
      x: x + 0.42, y: 2.78, w: cw - 0.84, h: 0.55, fontFace: FS, fontSize: 23,
      bold: true, color: INK, margin: 0,
    });
    c[2].forEach((t, j) => {
      s.addText(t, {
        x: x + 0.42, y: 3.5 + j * 0.58, w: cw - 0.84, h: 0.45,
        fontFace: FS, fontSize: 14, color: INK_2, margin: 0, valign: "middle",
        bullet: { code: "25AA" },
      });
    });
  });
  card(s, M, 5.75, W, 0.95, { fill: "FFFFFF" });
  s.addText("多學的那一件事：讓網頁自己去讀檔案，而不是把答案先寫進去。", {
    x: M + 0.42, y: 5.75, w: W - 0.84, h: 0.95, fontFace: FS, fontSize: 16,
    bold: true, color: INK, margin: 0, valign: "middle",
  });
  note(s, "把兩個網頁並排開兩個分頁對照，比投影片有效。");
}

/* 29 邊界 */
{
  const s = light();
  eyebrow(s, "邊界", BLUE);
  title(s, "一個要記住的邊界");
  card(s, M, 2.2, W, 1.5, { fill: INK, line: INK });
  s.addText("這是「快照」，不是「即時系統」。", {
    x: M + 0.5, y: 2.2, w: W - 1.0, h: 1.5, fontFace: FS, fontSize: 26,
    bold: true, color: "F2F5F4", margin: 0, valign: "middle",
  });
  const cw = 5.9;
  const bd = [
    ["做得到", ["上傳檔案，畫面立刻重算", "完全不需要伺服器", "資料不會外流，只在你電腦裡"], GOOD],
    ["做不到", ["關掉頁面，上傳的資料就沒了", "別人打開看不到你的資料", "資料源頭改了，不會自動同步"], CRIT],
  ];
  bd.forEach((b, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 4.0, cw, 2.55);
    s.addText(b[0], {
      x: x + 0.42, y: 4.25, w: cw - 0.84, h: 0.4, fontFace: FS, fontSize: 18,
      bold: true, color: b[3], margin: 0,
    });
    b[1].forEach((t, j) => {
      s.addText(t, {
        x: x + 0.42, y: 4.75 + j * 0.56, w: cw - 0.84, h: 0.45,
        fontFace: FS, fontSize: 14, color: INK_2, margin: 0, valign: "middle",
        bullet: { code: "25AA" },
      });
    });
  });
  note(s, "不要把「做不到」講成缺點。週報、月報本來就是快照，這個做法完全match真實工作情境。");
}

/* 30 何時需要資料庫 */
{
  const s = light();
  eyebrow(s, "再往下一步", BLUE);
  title(s, "什麼時候才真的需要工程師");
  s.addText("以下三種需求，超出今天教的範圍 —— 但你已經知道怎麼把需求講清楚了。", {
    x: M, y: 1.85, w: W, h: 0.4, fontFace: FS, fontSize: 15, color: MUTED, margin: 0,
  });
  const needs = [
    ["多人同時看同一份資料", "業務改了數字，主管那邊要立刻看到"],
    ["資料要長期累積", "要能查半年前的進度，做趨勢比較"],
    ["自動從別的系統抓資料", "不想每次手動匯出 Excel 再上傳"],
  ];
  needs.forEach((n, i) => {
    const y = 2.5 + i * 1.32;
    card(s, M, y, W, 1.12);
    s.addShape(pres.ShapeType.rect, {
      x: M + 0.42, y: y + 0.4, w: 0.32, h: 0.32, fill: { color: BLUE },
    });
    s.addText(n[0], {
      x: M + 1.05, y: y + 0.12, w: 5.2, h: 0.45, fontFace: FS, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(n[1], {
      x: M + 1.05, y: y + 0.58, w: 10.2, h: 0.45, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "top",
    });
  });
  s.addText("知道邊界在哪，也是一種專業 —— 你會知道什麼時候該找人幫忙。", {
    x: M, y: 6.55, w: W, h: 0.45, fontFace: FS, fontSize: 15,
    bold: true, color: BLUE, margin: 0,
  });
  note(s, "這頁的作用是讓學員不要對這套方法過度期待，也不要低估自己已經能做的事。");
}

/* 31 CSV 編碼陷阱 */
{
  const s = light();
  eyebrow(s, "Windows 注意", WARN);
  title(s, "Excel 存 CSV 有個坑");
  card(s, M, 2.15, W, 1.35, { fill: "FFFFFF", line: WARN, lw: 1.5 });
  s.addText("另存新檔時，一定要選「CSV UTF-8（逗號分隔）」", {
    x: M + 0.5, y: 2.15, w: W - 1.0, h: 1.35, fontFace: FS, fontSize: 22,
    bold: true, color: INK, margin: 0, valign: "middle",
  });
  const cw = 5.9;
  const two = [
    ["選錯會怎樣", "Windows 版 Excel 預設存的是 Big5 編碼，網頁讀進去中文會整片變亂碼", CRIT],
    ["怎麼確認", "另存新檔的下拉選單裡，找有寫「UTF-8」那一個，不是最上面那個「CSV」", GOOD],
  ];
  two.forEach((t, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 3.85, cw, 2.1);
    s.addText(t[0], {
      x: x + 0.42, y: 4.1, w: cw - 0.84, h: 0.4, fontFace: FS, fontSize: 17,
      bold: true, color: t[2], margin: 0,
    });
    s.addText(t[1], {
      x: x + 0.42, y: 4.6, w: cw - 0.84, h: 1.2, fontFace: FS, fontSize: 14,
      color: INK_2, lineSpacing: 23, margin: 0, valign: "top",
    });
  });
  s.addText("這是全場最容易卡住的地方，先講在前面。", {
    x: M, y: 6.3, w: W, h: 0.45, fontFace: FS, fontSize: 15, color: WARN, bold: true, margin: 0,
  });
  note(s, "務必在學員動手前講。不然會有一半的人卡在亂碼，時間全部耗在這裡。");
}

/* 32 實作 */
{
  const s = light();
  eyebrow(s, "動手 · 30 分鐘", BLUE);
  title(s, "換一份資料，看畫面自己算");
  const steps = [
    ["01", "先用範例檔跑一次", "確認上傳流程會動，知道畫面會怎麼變"],
    ["02", "換成自己的資料", "欄位名稱對上，存成 CSV UTF-8"],
    ["03", "故意改幾個數字再上傳", "看 KPI 跟著跳 —— 這就是「互動」"],
    ["04", "想想這能用在你哪份報表", "每週要做的那份，是不是可以省下來"],
  ];
  steps.forEach((st, i) => {
    const y = 2.15 + i * 1.1;
    s.addShape(pres.ShapeType.rect, {
      x: M, y: y, w: 0.52, h: 0.52, fill: { color: BLUE },
    });
    s.addText(st[0], {
      x: M, y: y, w: 0.52, h: 0.52, fontFace: FS, fontSize: 12,
      bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    s.addText(st[1], {
      x: M + 0.85, y: y, w: 4.6, h: 0.52, fontFace: FS, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(st[2], {
      x: M + 5.6, y: y, w: 6.3, h: 0.52, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "middle",
    });
  });
  s.addText("第四步才是重點 —— 做出來只是開始，用得上才有價值。", {
    x: M, y: 6.55, w: W, h: 0.45, fontFace: FS, fontSize: 15,
    bold: true, color: BLUE, margin: 0,
  });
  note(s, "留最後十分鐘讓幾位學員分享自己做的東西，收尾效果最好。");
}

/* ══════════════════ 收尾 ══════════════════ */

/* 33 四階段回顧 */
{
  const s = light();
  eyebrow(s, "回顧", BRASS);
  title(s, "你今天學會的四件事");
  const recap = [
    ["01", "把需求講清楚", "四個要素：做什麼、關於誰、要有什麼、什麼調性", BRASS],
    ["02", "檔案怎麼跟網頁連結", "網頁只記住圖片放在哪，換檔案就換圖", BRASS],
    ["03", "資料怎麼變成畫面", "先整理乾淨，再決定哪一欄要被強調", BLUE],
    ["04", "靜態與互動的差別", "以及什麼時候該找工程師幫忙", BLUE],
  ];
  recap.forEach((r, i) => {
    const y = 2.1 + i * 1.15;
    card(s, M, y, W, 0.98);
    s.addText(r[0], {
      x: M + 0.42, y: y, w: 0.6, h: 0.98, fontFace: FS, fontSize: 15,
      bold: true, color: r[3], margin: 0, valign: "middle",
    });
    s.addText(r[1], {
      x: M + 1.2, y: y, w: 4.2, h: 0.98, fontFace: FS, fontSize: 19,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + 5.5, y: y, w: 6.4, h: 0.98, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "middle",
    });
  });
  note(s, "回顧時強調：這四件事沒有一件需要記語法。");
}

/* 34 可以用在哪 */
{
  const s = light();
  eyebrow(s, "帶回去用", BRASS);
  title(s, "回去可以先做哪一件");
  const uses = [
    ["工地週報", "把每週的進度表做成儀表板，開會直接投影", BLUE],
    ["建案介紹頁", "新案子的簡介，不用等廣告公司排版", BRASS],
    ["內部活動報名", "尾牙、教育訓練的說明頁，自己一小時做完", BRASS],
    ["銷售日報", "把每日成交數字做成一頁，主管自己看", BLUE],
  ];
  const cw = 5.9;
  uses.forEach((u, i) => {
    const x = M + (i % 2) * (cw + 0.25);
    const y = 2.15 + Math.floor(i / 2) * 1.9;
    card(s, x, y, cw, 1.6);
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.42, y: y + 0.32, w: 0.3, h: 0.3, fill: { color: u[2] },
    });
    s.addText(u[0], {
      x: x + 1.0, y: y + 0.22, w: cw - 1.4, h: 0.5, fontFace: FS, fontSize: 19,
      bold: true, color: INK, margin: 0, valign: "middle",
    });
    s.addText(u[1], {
      x: x + 1.0, y: y + 0.78, w: cw - 1.4, h: 0.65, fontFace: FS, fontSize: 14,
      color: MUTED, margin: 0, valign: "top",
    });
  });
  s.addText("挑最小、最常做、最煩的那一件開始 —— 一週內做得完的才會真的做。", {
    x: M, y: 6.3, w: W, h: 0.45, fontFace: FS, fontSize: 16,
    bold: true, color: BRASS, margin: 0,
  });
  note(s, "鼓勵當場認領一件，散會前寫在便利貼上。有承諾才有後續。");
}

/* 35 結尾 */
{
  const s = pres.addSlide();
  darkBg(s);
  s.addText("帶走這句", {
    x: M, y: 2.3, w: 8, h: 0.34, fontFace: FS, fontSize: 12,
    color: BRASS_LT, charSpacing: 3, margin: 0,
  });
  s.addText("你缺的從來不是技術，\n是有人告訴你「其實你可以」。", {
    x: M, y: 2.9, w: 11.5, h: 1.9, fontFace: FS, fontSize: 36,
    bold: true, color: "F2F5F4", margin: 0, valign: "middle", lineSpacing: 52,
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: 5.15, w: 3.2, h: 0, line: { color: BRASS, width: 1.5 },
  });
  s.addText("課後材料：四個階段的完整範例，都在共用資料夾裡", {
    x: M, y: 5.45, w: 10, h: 0.4, fontFace: FS, fontSize: 14,
    color: MUTED, margin: 0,
  });
  note(s, "收尾不要再講技術。留一句能帶回去的話就好。");
}

pres.writeFile({ fileName: "vibe-coding-workshop.pptx" })
  .then(f => console.log("written:", f));
