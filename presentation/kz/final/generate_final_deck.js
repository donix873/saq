import pptxgen from 'pptxgenjs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Бигелді Алдияр';
pptx.company = 'SAQ';
pptx.subject = 'SAQ — әлеуметтік инженериядан превентивті қорғаныс';
pptx.title = 'SAQ — Сақ бол. Қауіпсіз бол.';
pptx.lang = 'kk-KZ';
pptx.theme = { headFontFace: 'Noto Sans', bodyFontFace: 'Noto Sans', lang: 'kk-KZ' };

const ROOT = __dirname;
const V = path.join(ROOT, '..', 'visuals');
const A = path.join(ROOT, 'assets');
const QR = path.join(ROOT, '..', 'assets', 'saq-beta-qr.png');
const F = 'Noto Sans';
const ST = pptx.ShapeType;
const C = {
  bg: '071017', panel: '0E1B24', panel2: '122631', line: '2B404C',
  text: 'F6F9FB', muted: 'A6B8C2', cyan: '49EFEA', blue: '58A6FF',
  green: '5CE3A8', orange: 'FF7B45', red: 'FF6B6B', yellow: 'FFD166', ink: '071017'
};

function txt(slide, value, x, y, w, h, o = {}) {
  slide.addText(value, {
    x, y, w, h, fontFace: F, fontSize: o.fontSize || 16, color: o.color || C.text,
    bold: o.bold || false, margin: o.margin ?? 0, valign: o.valign || 'mid',
    align: o.align || 'left', fit: 'shrink', breakLine: false, isTextBox: true,
    paraSpaceAfterPt: o.paraSpaceAfterPt || 0, ...o
  });
}

function rich(slide, runs, x, y, w, h, o = {}) {
  slide.addText(runs, {
    x, y, w, h, fontFace: F, fontSize: o.fontSize || 16, color: o.color || C.text,
    margin: o.margin ?? 0, valign: o.valign || 'mid', align: o.align || 'left',
    fit: 'shrink', breakLine: false, isTextBox: true, ...o
  });
}

function rect(slide, x, y, w, h, fill = C.panel, stroke = C.line, tr = 8, radius = 0.15) {
  slide.addShape(ST.roundRect, {
    x, y, w, h, rectRadius: radius,
    fill: { color: fill, transparency: tr }, line: { color: stroke, width: 1 }
  });
}

function rule(slide, x, y, w, color = C.line, width = 1) {
  slide.addShape(ST.line, { x, y, w, h: 0, line: { color, width } });
}

function dot(slide, x, y, d, color = C.cyan) {
  slide.addShape(ST.ellipse, { x, y, w: d, h: d, fill: { color }, line: { color } });
}

function pill(slide, value, x, y, w, fill = C.cyan, fg = C.ink) {
  slide.addShape(ST.roundRect, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color: fill }, line: { color: fill } });
  txt(slide, value, x, y + 0.01, w, 0.29, { fontSize: 8.5, color: fg, bold: true, align: 'center', charSpacing: 0.9 });
}

function bg(slide, image, transparency = 30) {
  slide.addImage({ path: image, x: 0, y: 0, w: 13.333, h: 7.5 });
  slide.addShape(ST.rect, {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { color: C.bg, transparency }, line: { color: C.bg, transparency: 100 }
  });
}

function brand(slide, n, section, light = false) {
  slide.addImage({ path: path.join(A, 'saq-logo-mark.png'), x: 0.55, y: 0.34, w: 0.47, h: 0.39, transparency: 0 });
  txt(slide, 'SAQ', 1.1, 0.39, 0.68, 0.25, { fontSize: 12, color: C.text, bold: true, charSpacing: 0.8 });
  txt(slide, section.toUpperCase(), 1.96, 0.39, 5.5, 0.25, { fontSize: 8.5, color: light ? C.text : C.orange, bold: true, charSpacing: 1.35 });
  txt(slide, String(n).padStart(2, '0'), 12.18, 0.39, 0.55, 0.24, { fontSize: 9, color: C.muted, bold: true, align: 'right', charSpacing: 1 });
}

function footer(slide, source = '') {
  if (source) txt(slide, source, 0.58, 7.08, 10.8, 0.16, { fontSize: 6.8, color: '6B808C' });
  txt(slide, 'Бигелді Алдияр · SAQ', 11.15, 7.08, 1.58, 0.16, { fontSize: 6.8, color: '6B808C', align: 'right' });
}

function note(slide, value) { slide.addNotes(value); }

// 1 — Жоба
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-01.png'), 55);
  s.addShape(ST.rect, { x: 0, y: 0, w: 6.55, h: 7.5, fill: { color: C.bg, transparency: 10 }, line: { color: C.bg, transparency: 100 } });
  s.addImage({ path: path.join(A, 'saq-logo-full.png'), x: 0.62, y: 0.72, w: 1.72, h: 1.72 });
  txt(s, 'SAQ', 0.72, 2.55, 4.6, 0.96, { fontSize: 58, bold: true, charSpacing: -2.2 });
  txt(s, 'Сақ бол. Қауіпсіз бол.', 0.76, 3.57, 5.0, 0.5, { fontSize: 24, color: C.cyan, bold: true });
  txt(s, 'Адамды әлеуметтік инженериядан қауіпті әрекетке дейін қорғайтын тәуелсіз ЖИ-көмекші.', 0.76, 4.35, 5.25, 1.05, { fontSize: 16, color: C.text, bold: true, valign: 'top' });
  rule(s, 0.76, 5.66, 4.8, C.line, 1);
  rich(s, [
    { text: 'Банк транзакцияны көреді.\n', options: { color: C.muted } },
    { text: 'SAQ соған алып келетін қауіпті әңгімені көреді.', options: { color: C.text, bold: true } }
  ], 0.76, 5.9, 5.25, 0.83, { fontSize: 14, valign: 'top' });
  pill(s, 'ҚАЗАҚСТАН · 2026', 10.45, 6.63, 1.8, C.orange, C.ink);
  note(s, 'SAQ — әлеуметтік инженерияны қауіпті әрекет жасалмай тұрып түсіндіруге арналған ЖИ-көмекші. Атауы қазақтың «сақ» сөзінен шыққан. Негізгі ой: банк транзакцияны көреді, ал SAQ адамды сол транзакцияға итермелейтін қысымды, қорқынышты және жалған сенімді байқайды.');
}

// 2 — Мәселе
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-02.png'), 23);
  brand(s, 2, 'Мәселе');
  txt(s, 'Алаяқтық ақша аударымынан бұрын басталады', 0.6, 0.95, 8.4, 0.72, { fontSize: 30, bold: true });

  rect(s, 0.6, 1.94, 3.38, 3.45, C.panel, C.line, 4);
  txt(s, '$1,03 трлн+', 0.92, 2.25, 2.8, 0.7, { fontSize: 34, color: C.cyan, bold: true });
  txt(s, 'әлемдік бағаланған шығын', 0.92, 3.03, 2.65, 0.32, { fontSize: 12, bold: true });
  txt(s, 'GASA, 2024 · 12 айлық сауалнама бағасы', 0.92, 3.47, 2.65, 0.55, { fontSize: 9.5, color: C.muted, valign: 'top' });
  rule(s, 0.92, 4.35, 1.05, C.cyan, 3);
  txt(s, '$12,5 млрд', 0.92, 4.58, 2.8, 0.54, { fontSize: 25, color: C.orange, bold: true });
  txt(s, 'АҚШ-та 2024 жылы тіркелген шығын', 0.92, 5.06, 2.7, 0.3, { fontSize: 9.5, color: C.muted });

  rect(s, 8.66, 1.94, 4.05, 3.45, '10272A', C.cyan, 5);
  txt(s, 'Адамға қалай әсер етеді?', 8.98, 2.27, 3.3, 0.4, { fontSize: 19, bold: true });
  const signs = [
    ['01', 'Асықтырады', '«Қазір әрекет етпесеңіз…»'],
    ['02', 'Қорқытады', '«Шотыңыз бұғатталады»'],
    ['03', 'Сенімге кіреді', 'Банк, полиция немесе туыс болып көрінеді'],
    ['04', 'Әрекет сұрайды', 'Сілтеме, код, ақша аудару']
  ];
  signs.forEach((v, i) => {
    const y = 2.92 + i * 0.57;
    txt(s, v[0], 8.98, y, 0.36, 0.24, { fontSize: 8, color: C.cyan, bold: true });
    txt(s, v[1], 9.48, y - 0.04, 1.45, 0.3, { fontSize: 11, bold: true });
    txt(s, v[2], 10.8, y - 0.04, 1.53, 0.32, { fontSize: 8.5, color: C.muted, valign: 'top' });
  });
  rect(s, 4.25, 5.75, 5.35, 0.67, '301A1B', C.red, 7);
  txt(s, 'Қорғаныс құрылғыны ғана емес, адамның шешімін де қорғауы керек.', 4.52, 5.89, 4.82, 0.36, { fontSize: 13, bold: true, align: 'center' });
  footer(s, 'Дереккөздер: GASA, Global State of Scams 2024; FTC, Consumer Sentinel Network, 10.03.2025.');
  note(s, 'Мәселе вируспен шектелмейді. Алаяқ адамды асықтырады, қорқытады, сенімге кіреді және әрекет сұрайды. GASA жаһандық шығынды 1,03 триллион доллардан жоғары деп бағалады. FTC АҚШ-та 2024 жылы 12,5 миллиард доллар тіркелген шығын болғанын хабарлады. SAQ дәл осы шешім қабылдау сәтін нысанаға алады.');
}

// 3 — Сұраныс / нарық
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-08.png'), 27);
  brand(s, 3, 'Сұраныс және нарық');
  txt(s, 'Алдымен Қазақстан. Кейін — Орталық Азия.', 0.6, 0.95, 9.6, 0.72, { fontSize: 30, bold: true });
  txt(s, 'Цифрлық төлемдер өскен сайын, қауіпсіз шешімге сұраныс та өседі.', 0.62, 1.68, 7.35, 0.46, { fontSize: 14, color: C.muted });

  rect(s, 0.62, 2.42, 3.25, 2.74, C.panel, C.line, 3);
  txt(s, '85%+', 0.92, 2.72, 2.6, 0.65, { fontSize: 32, color: C.cyan, bold: true });
  txt(s, 'қолма-қолсыз төлемдер', 0.92, 3.42, 2.6, 0.38, { fontSize: 13, bold: true });
  txt(s, 'Қазақстан — цифрлық қаржысы дамыған бастапқы нарық.', 0.92, 4.02, 2.55, 0.72, { fontSize: 10.5, color: C.muted, valign: 'top' });

  const markets = [
    ['TAM', '$35 млрд+', 'Жаһандық санат', C.blue],
    ['SAM', '$420 млн', 'Өңірлік гипотеза', C.cyan],
    ['SOM', '$12 млн', '3 жылдық мақсат-гипотеза', C.orange]
  ];
  markets.forEach((m, i) => {
    const x = 4.22 + i * 2.75;
    rect(s, x, 2.42, 2.4, 2.74, i === 2 ? '3A2118' : C.panel, m[3], 4);
    txt(s, m[0], x + 0.24, 2.71, 0.7, 0.26, { fontSize: 9, color: m[3], bold: true, charSpacing: 1 });
    txt(s, m[1], x + 0.24, 3.23, 1.95, 0.5, { fontSize: 22, bold: true });
    txt(s, m[2], x + 0.24, 3.93, 1.9, 0.5, { fontSize: 10.5, color: C.muted, valign: 'top' });
  });
  rect(s, 0.62, 5.66, 11.85, 0.75, '10242B', C.line, 4);
  txt(s, 'Нарыққа шығу жолы', 0.92, 5.88, 1.7, 0.28, { fontSize: 9, color: C.orange, bold: true });
  txt(s, 'Қазақстан → Өзбекстан → Қырғызстан → халықаралық B2B API', 2.5, 5.81, 9.48, 0.42, { fontSize: 14, bold: true });
  footer(s, 'TAM / SAM / SOM — аудиттен өтпеген жұмыс гипотезалары. 85%+ көрсеткішінің нақты жылы финал алдында Ұлттық банк дерегімен жаңартылады.');
  note(s, 'Қазақстан — бастапқы нарық: цифрлық төлемдер кең таралған, орыс және қазақ тілдеріндегі жергілікті алаяқтық сценарийлері бар. TAM, SAM және SOM сандары әзірге гипотеза, сондықтан олар нақты нарық фактісі емес деп белгіленді. Алғашқы міндет — Қазақстанда сұраныс пен сапаны дәлелдеу.');
}

// 4 — Шешім
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-04.png'), 18);
  brand(s, 4, 'Шешім');
  txt(s, 'Күмәнді хабарламадан қауіпсіз әрекетке дейін', 0.6, 0.92, 9.0, 0.72, { fontSize: 30, bold: true });
  rect(s, 0.62, 1.9, 4.75, 3.95, C.panel, C.line, 5);
  pill(s, 'SAQ ТЕКСЕРУ · MVP', 0.92, 2.2, 1.65, C.cyan, C.ink);
  const flow = [
    ['1', 'Жіберу', 'Мәтін, сілтеме немесе скриншот'],
    ['2', 'Талдау', 'Қысым, асығыстық, жалған бедел'],
    ['3', 'Түсіндіру', 'Қауіптің себебін қарапайым тілмен'],
    ['4', 'Әрекет', 'Ресми арнаға өту, кодты бермеу']
  ];
  flow.forEach((v, i) => {
    const y = 2.86 + i * 0.67;
    dot(s, 0.94, y + 0.02, 0.26, i === 3 ? C.green : C.cyan);
    txt(s, v[0], 1.0, y + 0.055, 0.13, 0.12, { fontSize: 7, color: C.ink, bold: true, align: 'center' });
    txt(s, v[1], 1.38, y - 0.04, 1.15, 0.3, { fontSize: 13, bold: true });
    txt(s, v[2], 2.53, y - 0.04, 2.45, 0.34, { fontSize: 9.5, color: C.muted, valign: 'top' });
  });
  rect(s, 7.72, 2.05, 4.62, 3.65, '10282B', C.cyan, 5);
  txt(s, 'ДЕМО НӘТИЖЕСІ', 8.04, 2.38, 2.0, 0.26, { fontSize: 8.5, color: C.cyan, bold: true, charSpacing: 1 });
  txt(s, '94%', 8.02, 2.85, 1.45, 0.65, { fontSize: 36, color: C.red, bold: true });
  txt(s, 'жоғары қауіп', 9.35, 3.02, 2.2, 0.36, { fontSize: 16, color: C.red, bold: true });
  const why = ['Жасанды асығыстық', 'Банкке еліктеу', 'Күмәнді домен'];
  why.forEach((v, i) => { dot(s, 8.05, 3.78 + i * 0.44, 0.11, C.orange); txt(s, v, 8.32, 3.67 + i * 0.44, 2.95, 0.31, { fontSize: 10.5, bold: true }); });
  pill(s, 'РЕСМИ АРНАҒА ӨТУ', 8.02, 5.15, 2.35, C.green, C.ink);
  rect(s, 0.62, 6.18, 11.72, 0.42, '1B1715', C.orange, 4);
  txt(s, 'Құпиялылық: пайдаланушы нені жіберетінін өзі таңдайды. Жасырын оқу жоқ.', 0.88, 6.25, 11.18, 0.26, { fontSize: 10.5, bold: true, align: 'center' });
  footer(s, '94% — интерфейс симуляциясының демо-мәні; жұмыс істейтін модель сапасы емес.');
  note(s, 'SAQ Scan — бірінші MVP. Пайдаланушы күмәнді мәтінді, сілтемені немесе скриншотты өзі жібереді. Жүйе қауіп белгілерін түсіндіреді және ресми қауіпсіз әрекетті ұсынады. Бұл слайдтағы 94 пайыз — интерфейс симуляциясы, өлшенген модель нәтижесі емес.');
}

// 5 — Бизнес-модель
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-06.png'), 23);
  brand(s, 5, 'Бизнес-модель');
  txt(s, 'Алдымен сенім. Содан кейін — ақылы құндылық.', 0.6, 0.95, 9.5, 0.72, { fontSize: 30, bold: true });
  const tiers = [
    ['B2C', 'Тегін + жазылым', 'Базалық тексеру тегін\nОтбасылық жоспар — ақылы', C.cyan],
    ['B2B', 'Банк және телеком', 'API, қауіп сигнатуралары\nҚызметтік интеграция', C.blue],
    ['B2G', 'Әлеуметтік әріптестік', 'Цифрлық қауіпсіздік\nбағдарламалары', C.green]
  ];
  tiers.forEach((v, i) => {
    const x = 0.62 + i * 4.08;
    rect(s, x, 2.12, 3.65, 3.47, i === 2 ? '10281F' : C.panel, v[3], 5);
    pill(s, v[0], x + 0.3, 2.44, 0.7, v[3], C.ink);
    txt(s, v[1], x + 0.3, 3.13, 3.0, 0.54, { fontSize: 20, bold: true });
    txt(s, v[2], x + 0.3, 4.02, 2.92, 0.82, { fontSize: 11.5, color: C.muted, valign: 'top' });
    rule(s, x + 0.3, 5.14, 1.05, v[3], 3);
  });
  rect(s, 0.62, 5.94, 11.85, 0.66, '10242B', C.line, 4);
  txt(s, 'Тексерілетін экономика', 0.92, 6.12, 1.95, 0.28, { fontSize: 9, color: C.orange, bold: true });
  txt(s, 'ЖИ-сұрау құны · төлемге конверсия · қайта қолдану · отбасылық жоспар сұранысы', 2.82, 6.05, 9.15, 0.4, { fontSize: 12, bold: true });
  footer(s, 'B2B және B2G — жоспарланған бағыттар; жасалған серіктестік ретінде көрсетілмейді.');
  note(s, 'Бизнес-модель үш бағыттан тұрады. B2C-де базалық қорғаныс тегін болып, отбасылық мүмкіндіктер жазылымға өтеді. B2B — банк пен телекомға API немесе қауіп сигнатуралары. B2G — ұзақмерзімді әлеуметтік бағыт. Алғашқы кезеңде экономика гипотезаларын нақты өлшеу қажет.');
}

// 6 — Маркетинг және сату
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-09.png'), 24);
  brand(s, 6, 'Маркетинг және сату');
  txt(s, 'Өнім адамдарға сенімді арналар арқылы жетеді', 0.6, 0.95, 9.2, 0.72, { fontSize: 30, bold: true });
  const channels = [
    ['01', 'Telegram бот', 'Ең жылдам MVP арнасы\nКүмәнді хабарды бөлісу', C.cyan],
    ['02', 'Мектеп және отбасы', 'Киберсауат контенті\nШынайы сценарийлер', C.green],
    ['03', 'Банк және телеком', 'B2B2C қауіпсіздік пакеті\nПилоттық интеграция', C.blue]
  ];
  channels.forEach((v, i) => {
    const x = 0.62 + i * 4.08;
    rect(s, x, 2.05, 3.65, 2.5, C.panel, v[3], 5);
    txt(s, v[0], x + 0.28, 2.36, 0.42, 0.26, { fontSize: 9, color: v[3], bold: true });
    txt(s, v[1], x + 0.28, 2.93, 3.0, 0.45, { fontSize: 19, bold: true });
    txt(s, v[2], x + 0.28, 3.62, 2.9, 0.58, { fontSize: 10.5, color: C.muted, valign: 'top' });
  });
  txt(s, 'Пайдаланушы жолы', 0.62, 5.1, 1.75, 0.3, { fontSize: 9, color: C.orange, bold: true, charSpacing: 1 });
  const funnel = ['Көреді', 'Ботта тексереді', 'Пайдасын сезеді', 'Қайта қолданады', 'Отбасымен бөліседі'];
  funnel.forEach((v, i) => {
    const x = 0.62 + i * 2.43;
    rect(s, x, 5.55, 2.05, 0.72, i === 4 ? '10281F' : C.panel, i === 4 ? C.green : C.line, 4);
    txt(s, v, x + 0.1, 5.72, 1.85, 0.32, { fontSize: 10, bold: true, align: 'center' });
    if (i < 4) rule(s, x + 2.06, 5.91, 0.35, C.cyan, 1.5);
  });
  footer(s, 'Банк және телеком арналары — келісілген серіктестік емес, нарыққа шығу гипотезасы.');
  note(s, 'Алғашқы маркетинг қымбат жарнамадан басталмайды. Telegram бот арқылы адам күмәнді хабарды бір әрекетпен тексереді. Мектеп пен отбасы сенім қалыптастырады. B2B2C арнасы кейін банк пен телеком арқылы масштабтауға мүмкіндік береді, бірақ ол әзірге жоспар.');
}

// 7 — Гипотезалар және метрикалар
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-07.png'), 20);
  brand(s, 7, 'Гипотезалар және метрикалар');
  txt(s, 'Табысты «ұнады» емес, мінез-құлықтың өзгеруі дәлелдейді', 0.6, 0.95, 11.0, 0.72, { fontSize: 29, bold: true });
  const hypotheses = [
    ['H1', 'Түсіндіру әсері', 'Себепті түсінген адам қауіпсіз әрекетті жиі таңдайды.', C.cyan],
    ['H2', 'Қысқа жол', 'Бөлісу → тексеру жолы қолдануды жеңілдетеді.', C.orange],
    ['H3', 'Отбасылық құндылық', 'Жақынын қорғау жазылымға нақты себеп береді.', C.green]
  ];
  hypotheses.forEach((v, i) => {
    const x = 0.62 + i * 4.08;
    rect(s, x, 2.05, 3.65, 2.55, C.panel, v[3], 5);
    pill(s, v[0], x + 0.28, 2.34, 0.66, v[3], C.ink);
    txt(s, v[1], x + 0.28, 3.03, 2.9, 0.42, { fontSize: 17, bold: true });
    txt(s, v[2], x + 0.28, 3.65, 2.95, 0.63, { fontSize: 10.5, color: C.muted, valign: 'top' });
  });
  rect(s, 0.62, 5.05, 11.85, 1.25, '10242B', C.line, 4);
  txt(s, 'АЛҒАШҚЫ ТАБЫС ӨЛШЕМДЕРІ', 0.92, 5.35, 2.25, 0.26, { fontSize: 8.5, color: C.orange, bold: true, charSpacing: 1 });
  const metrics = [
    ['50–100', 'тест пайдаланушы'], ['↓', 'жалған дабыл'], ['↑', 'қауіпсіз әрекет'], ['↑', 'қайта қолдану']
  ];
  metrics.forEach((m, i) => {
    const x = 3.35 + i * 2.15;
    txt(s, m[0], x, 5.25, 1.25, 0.42, { fontSize: 20, color: i === 0 ? C.cyan : C.green, bold: true, align: 'center' });
    txt(s, m[1], x - 0.2, 5.75, 1.65, 0.28, { fontSize: 9, color: C.muted, align: 'center' });
  });
  footer(s, 'Нақты нысаналы мәндер MVP тестінен кейін калибрленеді.');
  note(s, 'Үш негізгі гипотеза тексеріледі: түсіндірме қауіпсіз шешімге әсер ете ме, бөлісу мәзірі арқылы тексеру жеткілікті жеңіл ме, отбасы үшін қорғау ақылы құндылық бере ме. Табыстың өлшемі — жүктеу саны емес, қауіпсіз әрекет, төмен жалған дабыл және қайта қолдану.');
}

// 8 — Команда / негізін қалаушы
{
  const s = pptx.addSlide();
  bg(s, path.join(A, 'founder-story.png'), 22);
  s.addShape(ST.rect, { x: 0, y: 0, w: 7.1, h: 7.5, fill: { color: C.bg, transparency: 7 }, line: { color: C.bg, transparency: 100 } });
  brand(s, 8, 'Команда және негізін қалаушы');
  pill(s, 'КОМАНДА: 1 АДАМ', 0.62, 1.13, 1.75, C.orange, C.ink);
  txt(s, 'Бигелді Алдияр', 0.62, 1.75, 5.7, 0.78, { fontSize: 34, bold: true });
  txt(s, 'Негізін қалаушы · өнім · зерттеу · алғашқы әзірлеу', 0.64, 2.52, 5.65, 0.42, { fontSize: 13, color: C.cyan, bold: true });
  rect(s, 0.62, 3.17, 5.85, 2.45, C.panel, C.line, 3);
  txt(s, 'Неге SAQ?', 0.93, 3.47, 1.55, 0.35, { fontSize: 17, bold: true });
  rich(s, [
    { text: '«Әжемді алаяқтар алдауға тырысқаннан кейін, мен бір нәрсені түсіндім: адамға қауіп туралы ', options: { color: C.text } },
    { text: 'ақша жоғалмай тұрып', options: { color: C.orange, bold: true } },
    { text: ' айту керек. Сол сәттен бастап осындай жағдай қайталанбауы үшін SAQ идеясын ойластыра бастадым.»', options: { color: C.text } }
  ], 0.93, 3.95, 5.18, 1.28, { fontSize: 13, valign: 'top' });
  const roles = ['Жоба идеясы', 'Мәселені зерттеу', 'Өнім тұжырымдамасы', 'Прототип бағыты'];
  roles.forEach((r, i) => { dot(s, 0.67 + i * 1.48, 6.1, 0.1, C.cyan); txt(s, r, 0.85 + i * 1.48, 5.99, 1.22, 0.34, { fontSize: 8.7, bold: true }); });
  pill(s, '11-СЫНЫП ОҚУШЫСЫ · ҚАЗАҚСТАН', 0.62, 6.61, 2.9, C.cyan, C.ink);
  txt(s, 'Көрнекі иллюстрация', 10.78, 6.81, 1.43, 0.2, { fontSize: 7, color: C.muted, align: 'right' });
  footer(s);
  note(s, 'Менің атым — Бигелді Алдияр. SAQ идеясы жеке оқиғадан басталды: әжемді телефон алаяқтары алдауға тырысты. Сол кезде қауіп туралы ақша жоғалғаннан кейін емес, оған дейін ескерту керек екенін түсіндім. Жобаның идеясын, зерттеуін, өнім тұжырымдамасын және алғашқы прототип бағытын әзірге жеке өзім жасап келемін.');
}

// 9 — Статус және жоспар
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-10.png'), 24);
  brand(s, 9, 'Ағымдағы статус және жоспар');
  txt(s, 'Идеядан өлшенетін MVP-ге дейін', 0.6, 0.95, 8.7, 0.72, { fontSize: 30, bold: true });
  rect(s, 0.62, 1.95, 3.18, 3.82, C.panel, C.cyan, 4);
  txt(s, 'ҚАЗІР ДАЙЫН', 0.92, 2.29, 1.6, 0.26, { fontSize: 8.5, color: C.cyan, bold: true, charSpacing: 1 });
  const ready = ['Мәселе және позиционирлеу', 'SAQ бренді мен логотипі', 'Интерфейс концепті', 'Презентация және MVP жоспары'];
  ready.forEach((v, i) => { dot(s, 0.94, 2.92 + i * 0.55, 0.1, C.green); txt(s, v, 1.2, 2.82 + i * 0.55, 2.25, 0.3, { fontSize: 10.5, bold: true }); });

  const plan = [
    ['2026 · IV тоқсан', 'Сұхбат + бот MVP', '20 сұхбат, 30–50 сценарий'],
    ['2027 · I жартыжылдық', 'Қазақ/орыс пилоты', '50–100 тест пайдаланушы'],
    ['2027 · II жартыжылдық', 'Өңірлік тексеру', 'Өзбекстан және B2B пилот']
  ];
  plan.forEach((v, i) => {
    const y = 2.05 + i * 1.25;
    pill(s, String(i + 1).padStart(2, '0'), 4.35, y + 0.07, 0.55, i === 0 ? C.orange : C.cyan, C.ink);
    txt(s, v[0], 5.18, y - 0.02, 2.25, 0.28, { fontSize: 9, color: i === 0 ? C.orange : C.cyan, bold: true });
    txt(s, v[1], 5.18, y + 0.36, 2.6, 0.35, { fontSize: 15, bold: true });
    txt(s, v[2], 8.18, y + 0.32, 3.55, 0.4, { fontSize: 10, color: C.muted, valign: 'top' });
    if (i < 2) rule(s, 5.18, y + 0.92, 6.65, C.line, 1);
  });
  rect(s, 4.35, 5.9, 7.53, 0.6, '10242B', C.line, 4);
  txt(s, 'Келесі қадам: жұмыс істейтін Telegram бот және өлшенетін қазақ/орыс тесті.', 4.62, 6.02, 7.0, 0.34, { fontSize: 11.5, bold: true, align: 'center' });
  footer(s, 'Жол картасы ағымдағы 2026 жылғы қазан айына сәйкестендірілген.');
  note(s, 'Қазір мәселе, позиционирлеу, бренд, интерфейс концепті және MVP жоспары дайын. Келесі қадам — пайдаланушы сұхбаттары, сценарий корпусы және Telegram бот MVP. Содан кейін қазақ және орыс тілдерінде 50–100 адаммен пилот өткізу жоспарланған.');
}

// 10 — Нақты сұраныс
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-11.png'), 24);
  brand(s, 10, 'Қажетті ресурс және ұсыныс');
  txt(s, 'Қазір инвестициядан бұрын — дәлел керек', 0.6, 0.95, 9.4, 0.72, { fontSize: 30, bold: true });
  rect(s, 0.62, 2.0, 5.55, 3.65, C.panel, C.orange, 4);
  txt(s, 'БІЗГЕ ҚАЖЕТ', 0.94, 2.35, 1.65, 0.28, { fontSize: 9, color: C.orange, bold: true, charSpacing: 1 });
  const asks = [
    ['01', '50–100 тест пайдаланушы'],
    ['02', 'ЖИ және киберқауіпсіздік тәлімгері'],
    ['03', 'Алаяқтық сценарийлеріне сараптама'],
    ['04', 'Бір пилоттық серіктес']
  ];
  asks.forEach((v, i) => {
    const y = 2.95 + i * 0.56;
    txt(s, v[0], 0.94, y, 0.36, 0.24, { fontSize: 8, color: C.orange, bold: true });
    txt(s, v[1], 1.45, y - 0.06, 4.05, 0.34, { fontSize: 12, bold: true });
  });
  rect(s, 6.48, 2.0, 5.85, 3.65, '10271F', C.green, 4);
  txt(s, 'SAQ НЕ БЕРЕДІ', 6.8, 2.35, 1.9, 0.28, { fontSize: 9, color: C.green, bold: true, charSpacing: 1 });
  const gives = [
    ['Адамға', 'қауіпті әрекет алдында түсінікті көмек'],
    ['Отбасына', 'ерікті және құпия қорғаныс контуры'],
    ['Қоғамға', 'әр тексеру арқылы цифрлық сауат'],
    ['Серіктеске', 'өлшенетін жергілікті пилот']
  ];
  gives.forEach((v, i) => {
    const y = 2.9 + i * 0.56;
    txt(s, v[0], 6.8, y, 1.05, 0.28, { fontSize: 10.5, color: C.green, bold: true });
    txt(s, v[1], 7.95, y - 0.03, 3.9, 0.34, { fontSize: 10, color: C.text, bold: true });
  });
  rect(s, 0.62, 6.08, 11.72, 0.48, '10242B', C.line, 4);
  txt(s, 'Алғашқы мақсат: «көп қауіп таптық» емес, «адам қауіпсіз әрекетті таңдады».', 0.92, 6.17, 11.1, 0.3, { fontSize: 11.5, bold: true, align: 'center' });
  footer(s);
  note(s, 'Қазір жоба инвестиция сұрамайды. Алдымен дәлел керек: тест пайдаланушылар, техникалық тәлімгер, алаяқтық сценарийлеріне сараптама және бір пилоттық серіктес. SAQ-тың құндылығы — қауіп саны емес, адамның қауіпсіз әрекетті таңдауына көмектесу.');
}

// 11 — Байланыс
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-12.png'), 48);
  s.addShape(ST.rect, { x: 0, y: 0, w: 7.15, h: 7.5, fill: { color: C.bg, transparency: 9 }, line: { color: C.bg, transparency: 100 } });
  s.addImage({ path: path.join(A, 'saq-logo-full.png'), x: 0.62, y: 0.7, w: 1.55, h: 1.55 });
  txt(s, 'Сенбес бұрын ойлан.', 0.66, 2.46, 5.95, 0.78, { fontSize: 34, bold: true });
  txt(s, 'SAQ-ты нақты адамдармен тексеруге көмектесіңіз.', 0.67, 3.28, 5.4, 0.6, { fontSize: 16, color: C.cyan, bold: true });
  rect(s, 0.66, 4.15, 5.5, 1.62, C.panel, C.cyan, 4);
  txt(s, 'Бигелді Алдияр', 0.98, 4.44, 3.1, 0.38, { fontSize: 18, bold: true });
  txt(s, 'saq.protection@gmail.com', 0.98, 4.99, 3.9, 0.3, { fontSize: 12, color: C.text, bold: true });
  txt(s, '@saq_kz_bot', 0.98, 5.39, 2.9, 0.3, { fontSize: 12, color: C.text, bold: true });
  pill(s, 'БЕТА · SAQ-NINE.VERCEL.APP', 0.67, 6.22, 2.72, C.orange, C.ink);
  txt(s, 'SAQ', 9.2, 5.88, 2.55, 0.7, { fontSize: 42, bold: true, align: 'right' });
  txt(s, 'Сақ бол. Қауіпсіз бол.', 8.25, 6.54, 3.5, 0.36, { fontSize: 14, color: C.cyan, bold: true, align: 'right' });
  txt(s, '11', 12.18, 7.08, 0.55, 0.16, { fontSize: 7, color: '6B808C', align: 'right' });
  note(s, 'SAQ — әжемді алдауға тырысқан бір оқиғадан басталған идея. Енді оны нақты адамдармен тексерілетін өнімге айналдырғым келеді. Егер сіз тест пайдаланушы, тәлімгер немесе пилоттық серіктес бола алсаңыз, көрсетілген байланыс арналары арқылы хабарласыңыз. Сенбес бұрын ойлан.');
}

// 12 — Жұмыс істейтін бета-нұсқа
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-04.png'), 38);
  brand(s, 12, 'Жұмыс істейтін бета-нұсқа');
  txt(s, 'SAQ енді жұмыс істейді', 0.6, 0.95, 6.7, 0.72, { fontSize: 30, bold: true });

  pill(s, 'ҚАЗІР ҚОЛЖЕТІМДІ', 0.62, 1.84, 1.72, C.green, C.ink);
  txt(s, 'Күмәнді мәтінді немесе сілтемені әрекет жасамай тұрып тексеріңіз.', 0.62, 2.42, 3.62, 1.0, { fontSize: 20, bold: true, valign: 'top' });
  txt(s, 'Бета-нұсқада:', 0.62, 3.68, 1.5, 0.3, { fontSize: 10, color: C.cyan, bold: true, charSpacing: 1 });
  const betaFeatures = [
    'Қазақша және орысша интерфейс',
    'Мәтін мен сілтемені талдау',
    'Қауіп деңгейі мен түсіндірме',
    'Қауіпсіз келесі қадамдар'
  ];
  betaFeatures.forEach((value, i) => {
    dot(s, 0.65, 4.25 + i * 0.47, 0.11, i === 3 ? C.green : C.cyan);
    txt(s, value, 0.92, 4.14 + i * 0.47, 3.2, 0.32, { fontSize: 10.8, bold: true });
  });
  rect(s, 0.62, 6.22, 3.62, 0.46, '10281F', C.green, 4);
  txt(s, 'Мәтін браузерден сыртқа жіберілмейді.', 0.82, 6.31, 3.2, 0.27, { fontSize: 9.2, color: C.green, bold: true, align: 'center' });

  rect(s, 4.55, 1.78, 7.82, 4.93, '09131A', C.cyan, 2);
  s.addShape(ST.ellipse, { x: 4.83, y: 1.97, w: 0.09, h: 0.09, fill: { color: C.red }, line: { color: C.red } });
  s.addShape(ST.ellipse, { x: 5.03, y: 1.97, w: 0.09, h: 0.09, fill: { color: C.yellow }, line: { color: C.yellow } });
  s.addShape(ST.ellipse, { x: 5.23, y: 1.97, w: 0.09, h: 0.09, fill: { color: C.green }, line: { color: C.green } });
  txt(s, 'saq-nine.vercel.app', 8.15, 1.9, 2.65, 0.22, { fontSize: 7.5, color: C.muted, align: 'center' });
  s.addImage({ path: path.join(A, 'saq-beta-web.png'), x: 4.76, y: 2.24, w: 7.4, h: 4.22 });
  footer(s, 'Бета-нұсқа қателесуі мүмкін; нәтиже абсолютті қауіпсіздік кепілі емес.');
  note(s, 'SAQ енді тек презентациядағы идея емес. Жұмыс істейтін бета-нұсқа saq-nine.vercel.app мекенжайында қолжетімді. Онда пайдаланушы күмәнді мәтінді немесе сілтемені тексеріп, қауіп деңгейін, анықталған белгілерді және қауіпсіз келесі қадамдарды көре алады. Интерфейс қазақ және орыс тілдерінде жұмыс істейді.');
}

// 13 — QR арқылы ашу
{
  const s = pptx.addSlide();
  bg(s, path.join(V, 'slide-12.png'), 48);
  s.addShape(ST.rect, { x: 0, y: 0, w: 6.35, h: 7.5, fill: { color: C.bg, transparency: 7 }, line: { color: C.bg, transparency: 100 } });
  brand(s, 13, 'QR арқылы ашу');
  txt(s, 'SAQ-ты дәл қазір ашыңыз', 0.6, 0.95, 7.2, 0.72, { fontSize: 30, bold: true });

  rect(s, 0.62, 1.82, 4.52, 4.92, 'F7FBFC', C.cyan, 1);
  s.addImage({ path: QR, x: 1.27, y: 2.15, w: 3.22, h: 3.22 });
  txt(s, 'QR-КОДТЫ СКАНЕРЛЕҢІЗ', 1.08, 5.62, 3.62, 0.28, { fontSize: 9, color: '0A5960', bold: true, align: 'center', charSpacing: 1.1 });
  txt(s, 'saq-nine.vercel.app', 1.04, 6.06, 3.7, 0.36, { fontSize: 14, color: C.ink, bold: true, align: 'center' });

  pill(s, 'БЕТА-НҰСҚА', 6.08, 1.88, 1.38, C.orange, C.ink);
  txt(s, 'Сенбес бұрын — тексеріңіз.', 6.08, 2.55, 5.62, 0.72, { fontSize: 27, bold: true });
  txt(s, 'Күмәнді хабарламаны немесе сілтемені SAQ-қа енгізіп, қауіп белгілерін қарапайым тілмен көріңіз.', 6.08, 3.45, 5.38, 0.78, { fontSize: 14, color: C.text, bold: true, valign: 'top' });

  const qrPoints = [
    ['01', 'Сканерлеңіз', 'QR-кодты телефон камерасымен ашыңыз.'],
    ['02', 'Тексеріңіз', 'Күмәнді мәтінді немесе сілтемені енгізіңіз.'],
    ['03', 'Қауіпсіз әрекет етіңіз', 'Нәтиже мен ұсынылған қадамдарды оқыңыз.']
  ];
  qrPoints.forEach((value, i) => {
    const y = 4.62 + i * 0.56;
    pill(s, value[0], 6.1, y, 0.52, i === 2 ? C.green : C.cyan, C.ink);
    txt(s, value[1], 6.88, y - 0.03, 1.82, 0.3, { fontSize: 11.2, bold: true });
    txt(s, value[2], 8.72, y - 0.05, 2.95, 0.38, { fontSize: 8.5, color: C.muted, valign: 'top' });
  });
  footer(s, 'SAQ банк, полиция немесе киберқауіпсіздік маманының орнын алмастырмайды.');
  note(s, 'QR-код SAQ-тың жұмыс істейтін бета-нұсқасына апарады. Телефон камерасымен сканерлеп, күмәнді мәтінді немесе сілтемені тексеруге болады. SAQ қауіптің себебін түсіндіріп, қауіпсіз келесі қадамды ұсынады. Бұл ерте тест нұсқасы, сондықтан нәтиже абсолютті кепілдік емес. Өнімді қолданып көріп, кері байланыс беріңіз.');
}

pptx.writeFile({ fileName: path.join(ROOT, 'SAQ_Pitch_Deck_KZ_Final.pptx'), compression: true });
