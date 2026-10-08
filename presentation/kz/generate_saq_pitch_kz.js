const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'SAQ';
pptx.company = 'SAQ';
pptx.subject = 'SAQ AI digital protection startup — қазақша pitch deck';
pptx.title = 'SAQ — Сақ бол. Қауіпсіз бол.';
pptx.lang = 'kk-KZ';
pptx.theme = { headFontFace: 'Noto Sans', bodyFontFace: 'Noto Sans', lang: 'kk-KZ' };

const ST = pptx.ShapeType;
const F = 'Noto Sans';
const VISUAL_DIR = '/workspace/outputs/saq_kz_visuals';
const C = {
  bg: '081017', panel: '101B24', panel2: '142630', line: '263B47',
  text: 'F5F8FA', muted: '9FB2BE', cyan: '4DEEEA', blue: '5B9DFF',
  green: '5BE3A7', red: 'FF6B6B', yellow: 'FFD166', orange: 'FF7A45', ink: '071017'
};

pptx.defineSlideMaster({
  title: 'SAQ_KZ',
  background: { color: C.bg },
  objects: [
    { rect: { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.orange }, line: { color: C.orange } } },
  ],
  slideNumber: { x: 12.05, y: 7.04, w: 0.7, h: 0.2, color: '5D7280', fontFace: F, fontSize: 8, align: 'right' }
});

function text(slide, value, x, y, w, h, o = {}) {
  slide.addText(value, {
    x, y, w, h, fontFace: F, fontSize: o.fontSize || 16, color: o.color || C.text,
    bold: o.bold || false, margin: o.margin ?? 0, valign: o.valign || 'mid',
    align: o.align || 'left', fit: 'shrink', isTextBox: true,
    breakLine: false, paraSpaceAfterPt: o.paraSpaceAfterPt || 0, ...o
  });
}

function rich(slide, runs, x, y, w, h, o = {}) {
  slide.addText(runs, {
    x, y, w, h, fontFace: F, fontSize: o.fontSize || 16, color: o.color || C.text,
    margin: o.margin ?? 0, valign: o.valign || 'mid', align: o.align || 'left',
    fit: 'shrink', isTextBox: true, breakLine: false, ...o
  });
}

function box(slide, x, y, w, h, fill = C.panel, stroke = C.line, radius = 0.15) {
  slide.addShape(ST.roundRect, { x, y, w, h, rectRadius: radius, fill: { color: fill }, line: { color: stroke, width: 1 } });
}

function line(slide, x, y, w, h = 0, color = C.line, width = 1, dash = 'solid') {
  slide.addShape(ST.line, { x, y, w, h, line: { color, width, dashType: dash } });
}

function dot(slide, x, y, d, color = C.cyan) {
  slide.addShape(ST.ellipse, { x, y, w: d, h: d, fill: { color }, line: { color } });
}

function pill(slide, value, x, y, w, fill = C.cyan, fg = C.ink) {
  slide.addShape(ST.roundRect, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color: fill }, line: { color: fill } });
  text(slide, value, x, y + 0.01, w, 0.3, { fontSize: 8.5, color: fg, bold: true, align: 'center', charSpacing: 1.05 });
}

function visual(slide, no) {
  const overlayTransparency = [0, 54, 28, 30, 24, 28, 28, 28, 30, 28, 30, 26, 48][no];
  slide.addImage({ path: `${VISUAL_DIR}/slide-${String(no).padStart(2, '0')}.png`, x: 0, y: 0, w: 13.333, h: 7.5 });
  slide.addShape(ST.rect, {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { color: C.bg, transparency: overlayTransparency },
    line: { color: C.bg, transparency: 100 }
  });
}

function header(slide, no, kicker, title) {
  text(slide, kicker.toUpperCase(), 0.68, 0.42, 5.8, 0.26, { fontSize: 9, color: C.orange, bold: true, charSpacing: 1.5 });
  text(slide, title, 0.68, 0.82, 11.65, 0.83, { fontSize: 29, bold: true });
  pill(slide, `СЛАЙД ${String(no).padStart(2, '0')}`, 11.3, 0.38, 1.08, C.orange, C.ink);
  text(slide, 'SAQ · Сақ бол. Қауіпсіз бол.', 0.68, 7.04, 5.2, 0.2, { fontSize: 7.5, color: '5D7280', charSpacing: 0.6 });
}

function base(no, kicker, title) {
  const s = pptx.addSlide('SAQ_KZ');
  visual(s, no);
  header(s, no, kicker, title);
  return s;
}

function notes(slide, value) { slide.addNotes(value); }

// 1 — Title
{
  const s = pptx.addSlide('SAQ_KZ');
  visual(s, 1);
  pill(s, 'ЖИ ЦИФРЛЫҚ ҚОРҒАНЫС', 0.72, 0.62, 2.12, C.orange, C.ink);
  text(s, 'SAQ', 0.72, 1.28, 5.2, 1.12, { fontSize: 68, bold: true, charSpacing: -3 });
  text(s, 'Сақ бол. Қауіпсіз бол.', 0.78, 2.48, 6.4, 0.58, { fontSize: 27, bold: true });
  text(s, 'Әлеуметтік инженериядан тәуелсіз превентивті қорғаныс', 0.78, 3.28, 5.75, 0.78, { fontSize: 16, color: C.muted, valign: 'top' });
  line(s, 0.78, 4.42, 5.15, 0, C.line);
  text(s, '«Банк транзакцияны көреді.', 0.78, 4.7, 5.2, 0.32, { fontSize: 14, color: C.muted });
  text(s, 'SAQ соған алып келетін қауіпті контексті көреді.»', 0.78, 5.07, 5.65, 0.58, { fontSize: 17, color: C.text, bold: true, valign: 'top' });

  pill(s, 'ПРЕВЕНТИВТІ ҚОРҒАНЫС', 8.75, 5.92, 2.42, C.cyan, C.ink);
  text(s, 'ҚАЗАҚСТАН · 2026', 0.78, 6.82, 3.8, 0.22, { fontSize: 8, color: '5D7280', charSpacing: 1.7 });
  notes(s, 'SAQ — әлеуметтік инженерия мен цифрлық алаяқтықты қауіпті әрекет жасалмай тұрып түсіндіруге арналған AI-қорғаушы тұжырымдамасы. Атауы қазақтың «сақ» сөзінен шыққан. Негізгі идея: банк транзакцияны көреді, ал SAQ соған алып келетін қауіпті әңгімені, қысымды және контексті байқап, адамға қауіпсіз келесі қадамды ұсынады.');
}

// 2 — Problem and verified data
{
  const s = base(2, 'Өзекті мәселе', 'Алаяқтық ақша аударымынан бұрын басталады');
  box(s, 0.72, 1.92, 3.55, 3.85);
  text(s, '$1,03 трлн+', 1.0, 2.25, 2.95, 0.88, { fontSize: 37, color: C.cyan, bold: true, charSpacing: -1.5 });
  text(s, 'әлемдегі тұтынушылардың 12 айдағы алаяқтық шығыны', 1.0, 3.2, 2.95, 0.92, { fontSize: 14, bold: true, valign: 'top' });
  pill(s, 'GASA · 2024', 1.0, 4.5, 1.32, C.line, C.text);
  text(s, 'Сауалнамаға негізделген жаһандық бағалау', 1.0, 5.02, 2.9, 0.45, { fontSize: 9.5, color: C.muted });

  box(s, 4.53, 1.92, 3.55, 3.85);
  text(s, '$12,5 млрд', 4.82, 2.25, 2.95, 0.88, { fontSize: 36, color: C.orange, bold: true, charSpacing: -1.5 });
  text(s, 'АҚШ-та 2024 жылы тіркелген алаяқтық шығыны', 4.82, 3.2, 2.95, 0.92, { fontSize: 14, bold: true, valign: 'top' });
  pill(s, '+25% ЖЫЛДЫҚ ӨСІМ', 4.82, 4.5, 1.75, C.red, C.ink);
  text(s, 'FTC Consumer Sentinel Network', 4.82, 5.02, 2.9, 0.45, { fontSize: 9.5, color: C.muted });

  box(s, 8.34, 1.92, 4.2, 3.85, '0E2528', C.cyan);
  text(s, 'Көрінбейтін аймақ', 8.65, 2.25, 3.45, 0.5, { fontSize: 21, bold: true });
  const gaps = ['Сыртқы чаттағы манипуляция', 'Жалған жұмыс ұсынысы', 'Код пен құпияны сұрау', 'Қорқыту және асықтыру'];
  gaps.forEach((g, i) => { dot(s, 8.67, 3.12 + i * 0.55, 0.12, C.cyan); text(s, g, 8.95, 3.01 + i * 0.55, 3.0, 0.34, { fontSize: 11.5, bold: true }); });
  text(s, 'Банк өзінің экожүйесінен тыс әңгімені толық көрмеуі мүмкін.', 8.65, 5.38, 3.28, 0.28, { fontSize: 9.5, color: C.muted });
  text(s, 'Дереккөздер: GASA, Global State of Scams 2024; FTC, Consumer Sentinel Network, 10.03.2025.', 0.76, 6.66, 11.6, 0.28, { fontSize: 7.6, color: '667B87' });
  notes(s, 'Мәселенің ауқымы үлкен. GASA-ның 2024 жылғы жаһандық есебі соңғы он екі айдағы шығынды 1,03 триллион доллардан жоғары деп бағалайды. FTC 2024 жылы АҚШ-та 12,5 миллиард доллар тіркелген шығын болғанын хабарлады. Бірақ SAQ үшін маңыздысы — шығыннан бұрынғы кезең: адамды сыртқы чатта асықтырады, қорқытады немесе құпия код сұрайды. 70 пайыздық мессенджер көрсеткіші сенімді алғашқы дереккөз табылмағандықтан қолданылған жоқ.');
}

// 3 — Why now / Kazakhstan
{
  const s = base(3, 'Неге дәл қазір?', 'Қазақстан — цифрлық қорғанысты тексеруге қолайлы бастапқы нарық');
  const items = [
    ['85%+', 'Қолма-қолсыз төлемдер', 'Цифрлық қаржы күнделікті әдетке айналды', C.cyan],
    ['eGov', 'Ресми сервистер экожүйесі', 'Қауіпсіз әрекетке бағыттау мүмкіндігі', C.orange],
    ['3 нарық', 'Қазақстан → Өзбекстан → Қырғызстан', 'Тілдік және сценарийлік жақындық', C.green]
  ];
  items.forEach((it, i) => {
    const x = 0.72 + i * 4.08;
    box(s, x, 2.02, 3.65, 3.58, C.panel);
    text(s, it[0], x + 0.28, 2.32, 3.0, 0.77, { fontSize: 34, color: it[3], bold: true });
    text(s, it[1], x + 0.28, 3.26, 3.0, 0.62, { fontSize: 16, bold: true, valign: 'top' });
    text(s, it[2], x + 0.28, 4.18, 2.95, 0.72, { fontSize: 11.5, color: C.muted, valign: 'top' });
    line(s, x + 0.28, 5.18, 1.0, 0, it[3], 3);
  });
  box(s, 0.75, 6.0, 11.82, 0.57, '0D2028', C.line);
  text(s, 'eGov интеграциясы — дайын функция емес: API, құқықтық негіз және ресми келісім қажет.', 1.03, 6.12, 11.25, 0.3, { fontSize: 11, color: C.text, bold: true, align: 'center' });
  text(s, '85%+ — ҚР Ұлттық банкінің қолма-қолсыз төлемдер динамикасына негізделген бағдар; финал алдында нақты жыл/үлесті жаңарту қажет.', 0.76, 6.76, 11.7, 0.18, { fontSize: 7, color: '667B87' });
  notes(s, 'Неге Қазақстан? Біріншіден, қолма-қолсыз төлемдердің үлесі 85 пайыздан жоғары, сондықтан цифрлық тәуекел күнделікті өмірге жақын. Екіншіден, eGov сияқты ресми сервистер пайдаланушыны тексерілген арнаға бағыттау идеясын күшейтеді. Бірақ интеграция бар деп айтпаймыз: API, құқықтық негіз және келісім қажет. Үшіншіден, Қазақстаннан көрші нарықтарға бейімделуге болады.');
}

// 4 — Interactive MVP simulation
{
  const s = base(4, 'MVP симуляциясы', 'Күмәнді хабарламадан қауіпсіз әрекетке дейін');
  // Left phone
  s.addShape(ST.roundRect, { x: 0.82, y: 1.72, w: 3.65, h: 4.82, rectRadius: 0.32, fill: { color: '05090C' }, line: { color: '3C5360', width: 2 } });
  s.addShape(ST.roundRect, { x: 1.0, y: 1.92, w: 3.29, h: 4.42, rectRadius: 0.2, fill: { color: 'F5F7F8' }, line: { color: 'F5F7F8' } });
  text(s, 'SMS', 1.27, 2.15, 0.7, 0.32, { fontSize: 16, color: C.ink, bold: true });
  text(s, 'Kaspi қауіпсіздік қызметі', 1.27, 2.62, 2.55, 0.3, { fontSize: 10, color: '65737C', bold: true });
  box(s, 1.27, 3.08, 2.74, 1.75, 'E8ECEF', 'D8E0E4');
  rich(s, [
    { text: 'Картаңыз бұғатталады. Деректерді ', options: { color: '25333B' } },
    { text: 'шұғыл', options: { color: 'B22D2D', bold: true } },
    { text: ' растаңыз:\nkaspi-sec-pay.cc/auth', options: { color: '25333B' } }
  ], 1.48, 3.3, 2.32, 1.3, { fontSize: 11, valign: 'top' });
  pill(s, 'КҮМӘНДІ SMS', 1.27, 5.26, 1.25, C.red, C.ink);
  text(s, 'Сілтемені ашпай SAQ-қа бөлісу', 1.27, 5.78, 2.7, 0.35, { fontSize: 10, color: '586A74' });

  line(s, 4.72, 4.1, 0.65, 0, C.cyan, 2);
  s.addShape(ST.chevron, { x: 5.23, y: 3.93, w: 0.28, h: 0.34, fill: { color: C.cyan }, line: { color: C.cyan } });

  // Right result
  box(s, 5.72, 1.72, 6.8, 4.82, C.panel, C.cyan);
  text(s, 'SAQ талдауы', 6.03, 2.02, 2.4, 0.4, { fontSize: 19, bold: true });
  pill(s, 'ДЕМО-НӘТИЖЕ', 10.55, 2.02, 1.52, C.line, C.text);
  text(s, '94%', 6.02, 2.72, 1.72, 0.82, { fontSize: 42, color: C.red, bold: true });
  text(s, 'жоғары қауіп', 7.7, 2.89, 2.25, 0.4, { fontSize: 18, color: C.red, bold: true });
  text(s, 'Түсіндірмелі ЖИ себептері', 6.03, 3.72, 2.9, 0.28, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1 });
  const reasons = ['Жасанды асығыстық', 'Ресми брендке еліктеу', 'Күмәнді домен', 'Қорқыту арқылы қысым'];
  reasons.forEach((r, i) => { dot(s, 6.05, 4.18 + i * 0.4, 0.1, i < 2 ? C.red : C.orange); text(s, r, 6.3, 4.08 + i * 0.4, 2.7, 0.3, { fontSize: 10.5, bold: true }); });
  s.addShape(ST.roundRect, { x: 9.45, y: 3.76, w: 2.62, h: 0.52, rectRadius: 0.16, fill: { color: C.red }, line: { color: C.red } });
  text(s, 'Сілтемені бұғаттау', 9.62, 3.87, 2.28, 0.28, { fontSize: 10.5, color: C.ink, bold: true, align: 'center' });
  s.addShape(ST.roundRect, { x: 9.45, y: 4.5, w: 2.62, h: 0.52, rectRadius: 0.16, fill: { color: C.cyan }, line: { color: C.cyan } });
  text(s, 'Ресми арнаға өту', 9.62, 4.61, 2.28, 0.28, { fontSize: 10.5, color: C.ink, bold: true, align: 'center' });
  s.addShape(ST.roundRect, { x: 9.45, y: 5.24, w: 2.62, h: 0.52, rectRadius: 0.16, fill: { color: C.line }, line: { color: C.line } });
  text(s, 'eGov тыйымы · концепт', 9.62, 5.35, 2.28, 0.28, { fontSize: 9.5, color: C.text, bold: true, align: 'center' });
  text(s, '94% — жұмыс істейтін модель өлшемі емес, интерфейсті көрсетуге арналған демо-мән.', 5.98, 6.06, 5.95, 0.25, { fontSize: 8.2, color: C.muted });
  notes(s, 'Бұл — интерактивті MVP симуляциясы. Сол жақта фишингтік SMS және жалған домен бар. Пайдаланушы оны ашпай, Share арқылы SAQ-қа жібереді. Оң жақта жүйе қауіп деңгейін, себептерін және қауіпсіз қадамды көрсетеді. 94 пайыз — өлшенген өнім метрикасы емес, интерфейстік демо. eGov тыйымы да әзірге концепт, жұмыс істейтін интеграция емес.');
}

// 5 — Access & privacy
{
  const s = base(5, 'Технология және құпиялылық', 'Рұқсат — минимум. Бақылау — пайдаланушыда.');
  const access = [
    ['БІРІНШІ НҰСҚА', 'Бөлісу мәзірі', 'Пайдаланушы мәтінді немесе скриншотты өзі жібереді', C.cyan],
    ['ШЕКТЕУЛІ', 'Арнайы рұқсаттар / SMS', 'Тек нақты қажеттілік, ашық келісім және платформа ережесі сақталса', C.orange],
    ['БОЛАШАҚ', 'Құрылғыдағы шағын модель', 'Сезімтал контентті құрылғыда өңдеу мақсаты', C.green]
  ];
  access.forEach((a, i) => {
    const x = 0.72 + i * 4.08;
    box(s, x, 2.0, 3.65, 3.25, i === 2 ? '0D2722' : C.panel, i === 2 ? C.green : C.line);
    pill(s, a[0], x + 0.28, 2.3, i === 1 ? 1.12 : 0.92, a[3], C.ink);
    text(s, a[1], x + 0.28, 3.02, 3.0, 0.45, { fontSize: 20, bold: true });
    text(s, a[2], x + 0.28, 3.75, 2.98, 0.92, { fontSize: 11.5, color: C.muted, valign: 'top' });
  });
  const principles = ['Жасырын оқу жоқ', 'Әдепкіде сақтау жоқ', 'Шифрлау', 'Өшіру құқығы'];
  principles.forEach((p, i) => {
    const x = 0.82 + i * 3.05;
    dot(s, x, 5.79, 0.14, C.cyan);
    text(s, p, x + 0.3, 5.68, 2.35, 0.36, { fontSize: 11.5, bold: true });
  });
  box(s, 0.76, 6.34, 11.76, 0.38, '1A1716', C.orange);
  text(s, 'Құрылғыдағы ЖИ іске асқан жоқ — бұл ұзақмерзімді техникалық мақсат.', 1.0, 6.4, 11.25, 0.24, { fontSize: 9.5, color: C.text, bold: true, align: 'center' });
  notes(s, 'Құпиялылық архитектурасының бірінші қағидасы — ең аз рұқсат. MVP-де Share Sheet жеткілікті: пайдаланушы өзі таңдап жібереді. Accessibility Services пен SMS фильтрациясы платформа саясатына қатаң тәуелді, сондықтан оларды дайын мүмкіндік деп көрсетпейміз. On-device шағын модель — жеке чаттарды серверге жібермеуге арналған болашақ мақсат.');
}

// 6 — Business model
{
  const s = base(6, 'Бизнес-модель', 'Қолжетімді базалық қорғаныс, кейін — ақылы құндылық');
  const tiers = [
    ['B2C · ТЕГІН + ЖАЗЫЛЫМ', 'Тегін → жазылым', 'Базалық тексеру тегін\nОтбасылық жоспар — ақылы', C.cyan],
    ['B2B · API', 'Банк / телеком', 'Қауіп сигнатуралары\nКоммуникация талдауы', C.blue],
    ['B2G · ҰЗАҚМЕРЗІМ', 'Әлеуметтік әріптестік', 'Цифрлық қауіпсіздік\nбағдарламалары', C.green]
  ];
  tiers.forEach((t, i) => {
    const x = 0.72 + i * 4.08;
    box(s, x, 2.02, 3.65, 3.48, C.panel);
    text(s, t[0], x + 0.28, 2.32, 2.9, 0.3, { fontSize: 9, color: t[3], bold: true, charSpacing: 1 });
    text(s, t[1], x + 0.28, 3.02, 3.0, 0.55, { fontSize: 21, bold: true });
    text(s, t[2], x + 0.28, 3.92, 2.95, 0.85, { fontSize: 12, color: C.muted, valign: 'top' });
    line(s, x + 0.28, 5.03, 1.0, 0, t[3], 3);
  });
  box(s, 0.76, 5.9, 11.76, 0.66, '0D2028', C.line);
  text(s, 'Алғашқы тексерілетін метрикалар', 1.03, 6.06, 2.6, 0.3, { fontSize: 10, color: C.orange, bold: true });
  text(s, 'ЖИ-сұрау құны · төлемге конверсия · отбасылық жоспар сұранысы · қайта оралу', 3.42, 6.0, 8.45, 0.4, { fontSize: 12, bold: true });
  notes(s, 'Монетизация үш бағыттан тұрады. B2C-де базалық тексеру тегін, ал Family Plan мен кеңейтілген мүмкіндіктер жазылымға өтеді. B2B-де банк пен телекомға API немесе қауіп сигнатуралары ұсынылуы мүмкін. B2G — тек ұзақмерзімді мүмкіндік, дайын келісім емес. Алдымен AI-сұрау құны, төлемге конверсия және ұстап қалу өлшенеді.');
}

// 7 — Algorithms
{
  const s = base(7, 'Анықтау алгоритмдері', 'Бір модель емес — бірнеше дәлел қабаты');
  const layers = [
    ['01', 'Тіл және мағына талдауы', 'Манипуляция, эмоциялық қысым, асығыстық және құпия сұрауды табу', C.cyan, 'БІРІНШІ НҰСҚА'],
    ['02', 'Дипфейк және дауыс талдауы', 'Клондалған дауыс пен бейнені анықтау; нақты жылдамдық тестпен өлшенеді', C.orange, 'ЗЕРТТЕУ'],
    ['03', 'Қауіптер графы', 'Домен, нөмір, әмиян және сценарий арасындағы байланысты көру', C.green, 'ЖОСПАР']
  ];
  layers.forEach((l, i) => {
    const y = 1.95 + i * 1.43;
    pill(s, l[0], 0.78, y + 0.08, 0.55, l[3], C.ink);
    text(s, l[1], 1.62, y - 0.02, 3.2, 0.42, { fontSize: 18, bold: true });
    text(s, l[2], 5.0, y - 0.02, 5.62, 0.55, { fontSize: 11.5, color: C.muted, valign: 'top' });
    pill(s, l[4], 11.08, y + 0.05, 1.08, i === 0 ? C.cyan : C.line, i === 0 ? C.ink : C.text);
    if (i < 2) line(s, 1.62, y + 0.85, 10.55, 0, C.line);
  });
  box(s, 0.78, 6.22, 11.38, 0.48, '241A14', C.orange);
  text(s, '«0,3 секунд» уәдесі алынып тасталды: жылдамдық құрылғыда және нақты модельде өлшенуі тиіс.', 1.0, 6.31, 10.96, 0.29, { fontSize: 10, bold: true, align: 'center' });
  notes(s, 'SAQ технологиясы қабаттардан тұрады. Бірінші MVP-де NLP және семантикалық талдау бар. Deepfake пен дауыс талдауы — жеке зерттеу бағыты, ал Graph Threat Intelligence — кейінгі желілік қабат. 0,3 секунд деген уәде дәлелсіз болғандықтан қолданылмайды: latency нақты құрылғыда, нақты модельде және сапа шегінде өлшенуі керек.');
}

// 8 — Market size
{
  const s = base(8, 'Нарық көлемі', 'TAM / SAM / SOM — тексерілетін жұмыс гипотезасы');
  const market = [
    ['TAM', '$35 млрд+', 'Алаяқтықты анықтау және алдын алудың жаһандық санаты', C.blue, 4.1],
    ['SAM', '$420 млн', 'Орталық Азия + Кавказ + Таяу Шығыс және Солтүстік Африка', C.cyan, 3.05],
    ['SOM', '$12 млн', 'Қазақстан + Өзбекстан, алғашқы 3 жыл', C.orange, 1.95]
  ];
  market.forEach((m, i) => {
    s.addShape(ST.ellipse, { x: 0.9 + (4.1 - m[4]) / 2, y: 2.03 + (4.1 - m[4]) / 2, w: m[4], h: m[4], fill: { color: i === 0 ? '172535' : i === 1 ? '0F3D40' : '5A2E1E', transparency: i === 0 ? 8 : 0 }, line: { color: m[3], width: i === 2 ? 2 : 1.2 } });
  });
  text(s, 'SOM', 2.22, 3.24, 1.45, 0.26, { fontSize: 9, color: C.orange, bold: true, align: 'center' });
  text(s, '$12 млн', 1.98, 3.6, 1.95, 0.5, { fontSize: 21, bold: true, align: 'center' });
  text(s, 'мақсат-гипотеза', 2.08, 4.13, 1.75, 0.26, { fontSize: 8.5, color: C.muted, align: 'center' });

  market.forEach((m, i) => {
    const y = 2.0 + i * 1.3;
    text(s, m[0], 5.72, y, 0.72, 0.28, { fontSize: 9, color: m[3], bold: true, charSpacing: 1 });
    text(s, m[1], 6.55, y - 0.09, 1.6, 0.48, { fontSize: 21, bold: true });
    text(s, m[2], 8.2, y - 0.05, 3.8, 0.46, { fontSize: 11, color: C.muted, valign: 'top' });
    if (i < 2) line(s, 5.72, y + 0.72, 6.35, 0, C.line);
  });
  box(s, 5.72, 6.0, 6.35, 0.62, '241A14', C.orange);
  text(s, 'Есептеу әдісі мен алғашқы дереккөздер финал алдында аудиттен өтуі тиіс.', 5.98, 6.12, 5.85, 0.32, { fontSize: 9.8, bold: true, align: 'center' });
  text(s, '$35 млрд / $420 млн / $12 млн — осы презентация үшін берілген алдын ала бағалар; нақты нарықтық факт ретінде қолданылмайды.', 0.78, 6.76, 11.55, 0.18, { fontSize: 7, color: '667B87' });
  notes(s, 'Бұл слайдтағы 35 миллиард, 420 миллион және 12 миллион доллар — аудиттен өткен нарықтық есеп емес, жұмыс гипотезалары. Финалдық өтінімде TAM үшін бір анық санат пен дереккөз, SAM үшін елдер мен мақсатты клиенттер, SOM үшін баға, конверсия және үш жылдық сату формуласы көрсетілуі тиіс. Қазір бұл сандар масштабты талқылауға арналған.');
}

// 9 — GTM
{
  const s = base(9, 'Нарыққа шығу', 'Пайдаланушыға сенімді арна арқылы жету');
  const channels = [
    ['АРНАЛЫ САТУ', 'Банк + телеком', '«Қауіпсіздік пакеті» арқылы тарату', C.blue],
    ['ВИРУСТЫҚ ӨСІМ', 'Telegram бот-сканері', 'Күмәнді хабарламаны бір әрекетпен тексеру', C.cyan],
    ['СЕНІМ', 'Мектеп + отбасы', 'Киберсауат мазмұны және нақты мысалдар', C.green]
  ];
  channels.forEach((ch, i) => {
    const x = 0.72 + i * 4.08;
    box(s, x, 2.0, 3.65, 3.38, C.panel);
    text(s, ch[0], x + 0.28, 2.3, 2.6, 0.28, { fontSize: 9, color: ch[3], bold: true, charSpacing: 1 });
    text(s, ch[1], x + 0.28, 3.0, 3.0, 0.58, { fontSize: 20, bold: true });
    text(s, ch[2], x + 0.28, 3.87, 2.96, 0.75, { fontSize: 11.5, color: C.muted, valign: 'top' });
    line(s, x + 0.28, 4.9, 1.0, 0, ch[3], 3);
  });
  text(s, 'Алғашқы конверсия жолы', 0.78, 5.82, 1.7, 0.28, { fontSize: 9, color: C.orange, bold: true, charSpacing: 1 });
  const funnel = ['ботқа кіру', '1-ші тексеру', 'пайдалы нәтиже', 'қайта қолдану'];
  funnel.forEach((f, i) => {
    const x = 2.42 + i * 2.42;
    box(s, x, 5.68, 1.92, 0.62, i === 3 ? '0D2923' : C.panel, i === 3 ? C.green : C.line);
    text(s, f, x + 0.08, 5.82, 1.76, 0.28, { fontSize: 9.5, bold: true, align: 'center' });
    if (i < 3) line(s, x + 1.93, 5.99, 0.45, 0, C.cyan, 1.5);
  });
  text(s, 'Банк/телеком келісімдері — жоспар; жасалған серіктестік ретінде көрсетілмейді.', 0.78, 6.7, 11.5, 0.2, { fontSize: 7.8, color: '667B87' });
  notes(s, 'Алғашқы өсу стратегиясы үш арнаға сүйенеді. B2B2C — банк пен телекомның қауіпсіздік пакетіне кіру мүмкіндігі. Telegram бот — ең арзан тест арнасы және органикалық бөлісу циклі. Мектептер мен отбасылар — сенім мен киберсауат каналы. Бірақ ешбір банкпен немесе операторлармен келісім жасалды деп айтпаймыз.');
}

// 10 — Roadmap
{
  const s = base(10, 'Даму жоспары', '2026–2027: уәде емес, өлшенетін кезеңдер');
  line(s, 1.12, 3.26, 10.65, 0, C.line, 5);
  const phases = [
    ['2026 · IV тоқсан', 'MVP + Telegram боты', 'eGov API мүмкіндігін зерттеу', C.cyan, 1.2],
    ['2027 · I жартыжылдық', 'Қазақ/орыс пилоты', 'Банк/телекоммен ниет хаттары және дипфейк зерттеуі', C.orange, 5.0],
    ['2027 · II жартыжылдық', 'Өзбекстанда тексеру', '1 млн айлық белсенді аудитория — өршіл мақсат', C.green, 8.8]
  ];
  phases.forEach((p, i) => {
    dot(s, p[4], 3.08, 0.36, p[3]);
    text(s, p[0], p[4] - 0.2, 2.1, 1.8, 0.34, { fontSize: 10, color: p[3], bold: true });
    text(s, p[1], p[4] - 0.2, 3.7, 2.85, 0.48, { fontSize: 17, bold: true });
    text(s, p[2], p[4] - 0.2, 4.3, 3.05, 0.78, { fontSize: 10.5, color: C.muted, valign: 'top' });
  });
  box(s, 0.78, 5.72, 11.35, 0.72, '0D2028', C.line);
  text(s, 'Кезең өлшемдері', 1.05, 5.94, 1.3, 0.26, { fontSize: 9, color: C.orange, bold: true });
  text(s, 'пайдалы түсіндірме → қайта қолдану → қазақ/орыс сапасы → қайталанатын өсу арнасы', 2.25, 5.87, 9.55, 0.4, { fontSize: 11.5, bold: true });
  text(s, '2026 жылғы I–II тоқсан өтіп кеткендіктен, жоспар ағымдағы 2026 жылғы қазан айына сәйкестендірілді.', 0.78, 6.75, 11.4, 0.19, { fontSize: 7.5, color: '667B87' });
  notes(s, 'Бастапқы жол картасындағы 2026 жылдың бірінші және екінші тоқсандары өтіп кетті, сондықтан жоспар ағымдағы күнге сәйкестендірілді. 2026 жылдың соңына дейін MVP пен бот, 2027 жылдың бірінші жартысында RU/KZ пилот пен серіктестік ниеттері, екінші жартысында Өзбекстанда тексеру. Бір миллион MAU — кепіл емес, stretch goal.');
}

// 11 — Social impact
{
  const s = base(11, 'Әлеуметтік маңыз', 'Қорғаныс тек ақша емес — адамның дербестігін сақтау');
  const impact = [
    ['ОСАЛ ТОПТАР', 'Отбасылық қорғаныс', 'Зейнеткерлер, ата-аналар және балаларға ерікті көмек контуры', C.green],
    ['ЦИФРЛЫҚ САУАТ', 'Түсіндірмелі ЖИ', 'Әр тексеру арқылы алаяқтық белгілерін үйрету', C.cyan],
    ['ЭТИКА', 'Жасырын бақылаусыз', 'Жасырын оқу жоқ, адам нені бөлісетінін өзі таңдайды', C.orange]
  ];
  impact.forEach((it, i) => {
    const x = 0.72 + i * 4.08;
    box(s, x, 2.0, 3.65, 3.45, i === 0 ? '0D2923' : C.panel, i === 0 ? C.green : C.line);
    text(s, it[0], x + 0.28, 2.3, 2.7, 0.28, { fontSize: 9, color: it[3], bold: true, charSpacing: 1 });
    text(s, it[1], x + 0.28, 3.0, 3.0, 0.5, { fontSize: 20, bold: true });
    text(s, it[2], x + 0.28, 3.78, 2.96, 0.95, { fontSize: 11.5, color: C.muted, valign: 'top' });
    line(s, x + 0.28, 5.03, 1.0, 0, it[3], 3);
  });
  box(s, 0.76, 5.9, 11.76, 0.63, '0D2028', C.line);
  text(s, 'Әсер көрсеткіші: қанша қауіп табылды емес, қанша адам қауіпсіз әрекетті дұрыс таңдады.', 1.02, 6.03, 11.25, 0.36, { fontSize: 11.5, bold: true, align: 'center' });
  notes(s, 'SAQ-тың әлеуметтік маңызы үш бағытта. Family Shield осал топтарға тек ерікті түрде көмектеседі. Explainable AI әр нәтижені шағын киберсауат сабағына айналдырады. Ал этикалық қағида — жасырын бақылау жоқ. Негізгі impact метрикасы табылған қауіп саны емес, пайдаланушының қауіпсіз әрекетті дұрыс таңдауы.');
}

// 12 — CTA / contacts
{
  const s = pptx.addSlide('SAQ_KZ');
  visual(s, 12);
  pill(s, 'БІРГЕ ІСКЕ АСЫРАЙЫҚ', 0.75, 0.62, 1.85, C.orange, C.ink);
  text(s, 'SAQ-ты нақты адамдармен тексеруге көмектесіңіз.', 0.75, 1.28, 10.9, 1.02, { fontSize: 35, bold: true });
  box(s, 0.75, 2.85, 5.45, 2.72, C.panel);
  text(s, 'ҚАЗІР ҚАЖЕТ', 1.05, 3.2, 2.3, 0.28, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1.1 });
  const asks = ['50–100 тест пайдаланушы', 'ЖИ және киберқауіпсіздік тәлімгері', 'Сценарийлерге сараптамалық баға'];
  asks.forEach((a, i) => { dot(s, 1.08, 3.89 + i * 0.51, 0.13, C.cyan); text(s, a, 1.38, 3.77 + i * 0.51, 4.2, 0.38, { fontSize: 13, bold: true }); });

  box(s, 6.5, 2.85, 5.78, 2.72, '0D2528', C.cyan);
  text(s, 'БАЙЛАНЫС', 6.82, 3.2, 2.3, 0.28, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1.1 });
  text(s, 'saq.protection@gmail.com', 6.82, 3.8, 4.8, 0.38, { fontSize: 15, bold: true });
  text(s, '@saq_kz_bot', 6.82, 4.34, 4.8, 0.38, { fontSize: 15, bold: true });
  text(s, 'Веб-сайт: MVP кезеңі · жуырда', 6.82, 4.9, 4.8, 0.32, { fontSize: 11, color: C.muted });
  text(s, 'SAQ', 0.75, 6.2, 1.6, 0.55, { fontSize: 30, bold: true });
  text(s, 'Сенбес бұрын ойлан.', 2.3, 6.25, 4.7, 0.42, { fontSize: 19, color: C.cyan, bold: true });
  text(s, 'Рақмет', 10.75, 6.28, 1.4, 0.3, { fontSize: 13, color: C.muted, align: 'right' });
  notes(s, 'Қазір бізге ең алдымен нақты тексеру керек: 50–100 тест пайдаланушы, AI және cybersecurity менторы, сондай-ақ алаяқтық сценарийлеріне сараптамалық баға. Байланыс арналары осы слайдта көрсетілген. Веб-сайт дайын деп айтылмайды — ол MVP фазасында. SAQ-тың мақсаты: қауіпті әрекет алдында кідіріп, тексеруді жаңа цифрлық әдетке айналдыру.');
}

pptx.writeFile({ fileName: '/workspace/outputs/SAQ_Pitch_Deck_KZ.pptx', compression: true });
