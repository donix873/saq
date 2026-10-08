const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'SAQ';
pptx.company = 'SAQ';
pptx.subject = 'AI digital protection startup pitch deck';
pptx.title = 'SAQ — Think Before You Trust';
pptx.lang = 'ru-RU';
pptx.theme = {
  headFontFace: 'Noto Sans',
  bodyFontFace: 'Noto Sans',
  lang: 'ru-RU'
};
pptx.defineSlideMaster({
  title: 'SAQ_MASTER',
  background: { color: '0A0F14' },
  objects: [
    { rect: { x: 0, y: 0, w: 0.07, h: 7.5, fill: { color: '4DEEEA' }, line: { color: '4DEEEA' } } },
  ],
  slideNumber: { x: 12.45, y: 7.05, w: 0.35, h: 0.2, color: '5E7180', fontFace: 'Noto Sans', fontSize: 8, align: 'right' }
});

const C = {
  bg: '0A0F14',
  panel: '111922',
  panel2: '14212A',
  line: '263541',
  text: 'F4F7F9',
  muted: '9AAEBC',
  cyan: '4DEEEA',
  blue: '5B9DFF',
  green: '56E3A6',
  red: 'FF6B6B',
  yellow: 'FFD166',
  ink: '0B1319'
};
const F = 'Noto Sans';
const ST = pptx.ShapeType;

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: F,
    fontSize: opts.fontSize || 18,
    color: opts.color || C.text,
    bold: opts.bold || false,
    margin: opts.margin ?? 0,
    breakLine: false,
    valign: opts.valign || 'mid',
    align: opts.align || 'left',
    fit: 'shrink',
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 0,
    lineSpacingMultiple: opts.lineSpacingMultiple,
    isTextBox: true,
    ...opts
  });
}

function addRich(slide, runs, x, y, w, h, opts = {}) {
  slide.addText(runs, {
    x, y, w, h,
    fontFace: F,
    fontSize: opts.fontSize || 18,
    color: opts.color || C.text,
    margin: opts.margin ?? 0,
    valign: opts.valign || 'mid',
    align: opts.align || 'left',
    fit: 'shrink',
    breakLine: false,
    isTextBox: true,
    ...opts
  });
}

function box(slide, x, y, w, h, fill = C.panel, radius = 0.16, line = C.line) {
  slide.addShape(ST.roundRect, {
    x, y, w, h,
    rectRadius: radius,
    fill: { color: fill },
    line: { color: line, width: 1 }
  });
}

function line(slide, x, y, w, h = 0, color = C.line, width = 1, dash = 'solid') {
  slide.addShape(ST.line, { x, y, w, h, line: { color, width, dashType: dash } });
}

function dot(slide, x, y, d, color = C.cyan) {
  slide.addShape(ST.ellipse, { x, y, w: d, h: d, fill: { color }, line: { color } });
}

function pill(slide, text, x, y, w, color = C.cyan, textColor = C.ink) {
  slide.addShape(ST.roundRect, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color } });
  addText(slide, text, x, y + 0.005, w, 0.32, { fontSize: 9, color: textColor, bold: true, align: 'center', charSpacing: 1.3 });
}

function addHeader(slide, kicker, title, num) {
  addText(slide, kicker.toUpperCase(), 0.65, 0.42, 5.5, 0.28, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1.8 });
  addText(slide, title, 0.65, 0.83, 11.8, 0.82, { fontSize: 30, bold: true, breakLine: false });
  addText(slide, String(num).padStart(2, '0'), 11.95, 0.42, 0.65, 0.28, { fontSize: 9, color: C.muted, align: 'right', charSpacing: 1.5 });
}

function addFooter(slide, text = 'SAQ · Think Before You Trust') {
  addText(slide, text, 0.65, 7.05, 8.4, 0.2, { fontSize: 7.5, color: '5E7180', charSpacing: 0.6 });
}

function slideBase(kicker, title, num) {
  const slide = pptx.addSlide('SAQ_MASTER');
  addHeader(slide, kicker, title, num);
  addFooter(slide);
  return slide;
}

function addNotes(slide, text) {
  slide.addNotes(text);
}

// 1 — Cover
{
  const s = pptx.addSlide('SAQ_MASTER');
  pill(s, 'AI DIGITAL PROTECTION', 0.72, 0.62, 2.15);
  addText(s, 'SAQ', 0.72, 1.25, 5.2, 1.2, { fontSize: 70, bold: true, charSpacing: -3.5 });
  addText(s, 'Think Before You Trust.', 0.78, 2.55, 6.2, 0.62, { fontSize: 28, bold: true });
  addText(s, 'Персональный AI-защитник от социальной инженерии', 0.78, 3.35, 5.3, 0.75, { fontSize: 16, color: C.muted, breakLine: false });
  line(s, 0.78, 4.47, 4.5, 0, C.line, 1);
  addText(s, 'Казахстан → Центральная Азия → мир', 0.78, 4.66, 5.4, 0.28, { fontSize: 11, color: C.text });
  addText(s, '«сақ» — осторожный, бдительный', 0.78, 5.15, 5.2, 0.3, { fontSize: 10, color: C.cyan });

  // Radar / shield visual
  for (const [d, tr] of [[4.6, 82], [3.55, 75], [2.5, 66], [1.45, 50]]) {
    s.addShape(ST.ellipse, { x: 7.3 + (4.6 - d) / 2, y: 1.22 + (4.6 - d) / 2, w: d, h: d, fill: { color: C.bg, transparency: 100 }, line: { color: C.cyan, transparency: tr, width: 1.4 } });
  }
  s.addShape(ST.hexagon, { x: 8.72, y: 2.64, w: 1.75, h: 1.9, fill: { color: C.panel2 }, line: { color: C.cyan, width: 2.2 } });
  addText(s, 'S', 9.05, 2.91, 1.08, 1.0, { fontSize: 48, bold: true, color: C.cyan, align: 'center' });
  dot(s, 11.18, 2.0, 0.13, C.red);
  dot(s, 7.87, 4.92, 0.11, C.yellow);
  dot(s, 11.02, 5.14, 0.09, C.green);
  addText(s, 'STARTUP PITCH · 2026', 0.78, 6.82, 4.4, 0.22, { fontSize: 8, color: '5E7180', charSpacing: 1.8 });
  addNotes(s, 'Сегодня я представляю SAQ — идею персонального AI-защитника. Название происходит от казахского слова «сақ»: осторожный и бдительный. Наша цель — помочь человеку распознать манипуляцию до того, как он совершит опасное действие. SAQ не обещает абсолютную защиту. Он даёт человеку понятный второй взгляд именно в момент сомнения.');
}

// 2 — Problem
{
  const s = slideBase('Проблема', 'Атака начинается с доверия, а не с перевода', 2);
  box(s, 0.72, 1.87, 6.1, 4.55, C.panel);
  pill(s, 'ВХОДЯЩЕЕ СООБЩЕНИЕ', 1.04, 2.18, 1.95, C.line, C.text);
  addRich(s, [
    { text: '«Здравствуйте. Служба безопасности обнаружила подозрительную активность. ', options: { color: C.text } },
    { text: 'Срочно', options: { color: C.red, bold: true } },
    { text: ' подтвердите данные по ссылке, иначе ваш счёт будет заблокирован»', options: { color: C.text } }
  ], 1.06, 2.82, 5.34, 1.48, { fontSize: 21, bold: false, breakLine: false, valign: 'top' });
  line(s, 1.06, 4.68, 5.34, 0, C.line);
  addText(s, 'Человек видит знакомый бренд и ограниченное время.', 1.06, 4.98, 5.0, 0.56, { fontSize: 14, color: C.muted });
  addText(s, 'В этот момент антивирус может быть бессилен: ссылка ещё не нажата.', 1.06, 5.52, 5.2, 0.6, { fontSize: 14, color: C.text, bold: true });

  const cues = [
    ['01', 'Авторитет', '«служба безопасности»', C.blue],
    ['02', 'Страх', '«счёт заблокируют»', C.red],
    ['03', 'Срочность', '«действуйте сейчас»', C.yellow],
    ['04', 'Опасное действие', '«перейдите по ссылке»', C.cyan]
  ];
  cues.forEach((c, i) => {
    const y = 1.95 + i * 1.05;
    addText(s, c[0], 7.35, y, 0.45, 0.28, { fontSize: 9, color: c[3], bold: true });
    addText(s, c[1], 8.0, y - 0.08, 2.4, 0.4, { fontSize: 18, bold: true });
    addText(s, c[2], 8.0, y + 0.34, 3.7, 0.3, { fontSize: 11, color: C.muted });
    if (i < 3) line(s, 7.35, y + 0.79, 4.9, 0, C.line);
  });
  pill(s, 'СОЦИАЛЬНАЯ ИНЖЕНЕРИЯ', 9.72, 6.08, 2.55, C.red, C.ink);
  addNotes(s, 'Представьте обычное сообщение: якобы служба безопасности банка просит срочно подтвердить данные. Здесь нет сложного вируса. Есть четыре психологических механизма: авторитет, страх, срочность и призыв к опасному действию. Пользователь принимает решение под давлением. Значит, защищать нужно не только устройство или транзакцию — нужно поддержать человека в момент выбора.');
}

// 3 — Timing gap
{
  const s = slideBase('Момент риска', 'Чем раньше появляется помощь, тем больше шанс остановиться', 3);
  const stages = [
    ['1', 'Контакт', 'Сообщение или звонок', C.blue],
    ['2', 'Манипуляция', 'Страх + срочность', C.yellow],
    ['3', 'Решение', 'Клик, код, перевод', C.red],
    ['4', 'Последствие', 'Деньги или доступ потеряны', C.red]
  ];
  line(s, 1.3, 3.15, 9.9, 0, C.line, 5);
  stages.forEach((st, i) => {
    const x = 1.3 + i * 3.3;
    dot(s, x - 0.18, 2.97, 0.36, st[3]);
    addText(s, st[0], x - 0.13, 3.01, 0.26, 0.18, { fontSize: 8, color: C.ink, bold: true, align: 'center' });
    addText(s, st[1], x - 0.48, 3.52, 2.1, 0.38, { fontSize: 17, bold: true });
    addText(s, st[2], x - 0.48, 3.99, 2.4, 0.62, { fontSize: 11, color: C.muted, valign: 'top' });
  });
  box(s, 0.85, 5.18, 5.45, 1.0, '0E2A2C', 0.15, C.cyan);
  addText(s, 'SAQ', 1.15, 5.38, 0.75, 0.34, { fontSize: 18, color: C.cyan, bold: true });
  addText(s, 'анализирует контекст до действия', 2.0, 5.37, 3.8, 0.38, { fontSize: 15, bold: true });
  box(s, 7.02, 5.18, 5.05, 1.0, '25191C', 0.15, C.red);
  addText(s, 'АНТИФРОД', 7.32, 5.42, 1.35, 0.28, { fontSize: 10, color: C.red, bold: true, charSpacing: 1 });
  addText(s, 'часто видит финансовое действие', 8.82, 5.37, 2.8, 0.4, { fontSize: 14, bold: true });
  addText(s, 'Не противопоставление, а дополнительный слой защиты', 4.15, 6.5, 5.1, 0.28, { fontSize: 10, color: C.muted, align: 'center' });
  addNotes(s, 'Защита сегодня сильна на многих уровнях: браузеры проверяют ссылки, определители — номера, банки — операции. SAQ не утверждает, что они реагируют только после потери денег. Мы добавляем другой момент вмешательства: когда общение уже манипулирует человеком, но опасное действие ещё не совершено. Это дополнительный, а не конкурирующий слой защиты.');
}

// 4 — Gap
{
  const s = slideBase('Пробел', 'Инструменты видят сигнал. Пользователю нужен смысл.', 4);
  const cards = [
    ['БАНКОВСКИЙ АНТИФРОД', 'Транзакция', 'Силен внутри финансовой системы', C.blue],
    ['CALLER ID', 'Номер', 'Показывает репутацию звонящего', C.yellow],
    ['SAFE BROWSING', 'URL', 'Проверяет известную опасную ссылку', C.green]
  ];
  cards.forEach((c, i) => {
    const x = 0.72 + i * 3.05;
    box(s, x, 2.05, 2.72, 3.55, C.panel);
    addText(s, c[0], x + 0.24, 2.3, 2.22, 0.35, { fontSize: 8.5, color: c[3], bold: true, charSpacing: 1.1 });
    addText(s, c[1], x + 0.24, 3.03, 2.18, 0.6, { fontSize: 26, bold: true });
    addText(s, c[2], x + 0.24, 4.05, 2.16, 0.92, { fontSize: 12, color: C.muted, valign: 'top' });
    line(s, x + 0.24, 5.14, 1.0, 0, c[3], 3);
  });
  box(s, 10.05, 1.83, 2.5, 3.98, '0D2B2D', 0.18, C.cyan);
  addText(s, 'SAQ', 10.35, 2.2, 1.9, 0.55, { fontSize: 28, color: C.cyan, bold: true });
  addText(s, 'Контекст', 10.35, 3.0, 1.8, 0.42, { fontSize: 20, bold: true });
  addText(s, 'Почему это выглядит опасно?', 10.35, 3.58, 1.83, 0.64, { fontSize: 12, color: C.text, bold: true, valign: 'top' });
  addText(s, 'Что сделать безопасно прямо сейчас?', 10.35, 4.48, 1.82, 0.74, { fontSize: 12, color: C.muted, valign: 'top' });
  addText(s, 'Фокус продукта', 10.35, 5.38, 1.75, 0.24, { fontSize: 8, color: C.cyan, charSpacing: 1.1 });
  addText(s, 'SAQ соединяет сигналы в понятное решение для человека.', 0.78, 6.14, 8.9, 0.44, { fontSize: 17, bold: true });
  addNotes(s, 'Каждый существующий класс продукта решает важную задачу. Банк видит финансовый сигнал, caller ID — репутацию номера, браузер — URL. Но человек всё равно спрашивает: почему меня торопят, можно ли доверять этому сообщению и что делать сейчас? SAQ объединяет эти сигналы вокруг контекста и следующего безопасного действия.');
}

// 5 — Solution
{
  const s = slideBase('Решение', 'SAQ — система поддержки безопасных решений', 5);
  addText(s, 'The bank sees the transaction.', 0.78, 2.02, 7.0, 0.56, { fontSize: 24, color: C.muted });
  addText(s, 'SAQ sees the situation that leads to it.', 0.78, 2.62, 10.8, 0.74, { fontSize: 32, color: C.text, bold: true });
  line(s, 0.8, 3.67, 11.75, 0, C.cyan, 2);
  const principles = [
    ['01', 'Понять', 'Анализирует текст, ссылку или скриншот в контексте ситуации.'],
    ['02', 'Объяснить', 'Показывает признаки давления и социальной инженерии простым языком.'],
    ['03', 'Защитить', 'Предлагает конкретный безопасный следующий шаг — без абсолютных гарантий.']
  ];
  principles.forEach((p, i) => {
    const x = 0.78 + i * 4.05;
    addText(s, p[0], x, 4.16, 0.45, 0.3, { fontSize: 9, color: C.cyan, bold: true });
    addText(s, p[1], x, 4.58, 2.8, 0.46, { fontSize: 22, bold: true });
    addText(s, p[2], x, 5.18, 3.35, 0.98, { fontSize: 12, color: C.muted, valign: 'top' });
  });
  pill(s, 'НЕ АНТИВИРУС · НЕ БАНК · НЕ СЛЕЖКА', 4.18, 6.46, 4.65, C.line, C.text);
  addNotes(s, 'SAQ можно сформулировать в одной фразе: банк видит транзакцию, а SAQ анализирует ситуацию, которая к ней приводит. Продукт делает три вещи: понимает контекст, объясняет признаки манипуляции и предлагает безопасный шаг. При этом мы честно показываем неопределённость: AI может ошибаться, а финальное решение остаётся за пользователем.');
}

// 6 — Product mockup
{
  const s = slideBase('Продукт', 'Один экран. Один понятный следующий шаг.', 6);
  // Phone frame
  s.addShape(ST.roundRect, { x: 0.98, y: 1.7, w: 4.0, h: 4.95, rectRadius: 0.35, fill: { color: '05080B' }, line: { color: '3C4E5A', width: 2 } });
  s.addShape(ST.roundRect, { x: 1.16, y: 1.9, w: 3.64, h: 4.55, rectRadius: 0.22, fill: { color: 'F4F7F9' }, line: { color: 'F4F7F9' } });
  s.addShape(ST.roundRect, { x: 2.35, y: 1.78, w: 1.24, h: 0.16, rectRadius: 0.08, fill: { color: '202A32' }, line: { color: '202A32' } });
  addText(s, 'SAQ', 1.43, 2.12, 1.2, 0.34, { fontSize: 18, color: C.ink, bold: true });
  addText(s, 'Проверка сообщения', 1.43, 2.5, 2.5, 0.24, { fontSize: 9, color: '667784' });
  box(s, 1.4, 2.92, 3.12, 0.98, 'E8EDF0', 0.14, 'D3DBDF');
  addText(s, '«Срочно подтвердите данные по ссылке, иначе счёт будет заблокирован…»', 1.58, 3.08, 2.72, 0.64, { fontSize: 9, color: '25313A', valign: 'top' });
  s.addShape(ST.roundRect, { x: 1.4, y: 4.16, w: 3.12, h: 0.72, rectRadius: 0.14, fill: { color: 'FFE8E8' }, line: { color: 'FFB4B4' } });
  dot(s, 1.62, 4.37, 0.18, C.red);
  addText(s, 'Высокий риск', 1.93, 4.24, 1.65, 0.28, { fontSize: 12, color: 'A12B2B', bold: true });
  addText(s, 'Не переходите по ссылке', 1.93, 4.52, 2.1, 0.19, { fontSize: 8.5, color: '723D3D' });
  addText(s, 'Найдены признаки', 1.43, 5.08, 1.6, 0.22, { fontSize: 9, color: '667784', bold: true });
  pill(s, 'Срочность', 1.43, 5.43, 0.92, C.yellow, C.ink);
  pill(s, 'Страх', 2.48, 5.43, 0.72, C.red, C.ink);
  pill(s, 'Имитация банка', 3.32, 5.43, 1.2, C.cyan, C.ink);
  s.addShape(ST.roundRect, { x: 1.4, y: 5.92, w: 3.12, h: 0.34, rectRadius: 0.15, fill: { color: C.ink }, line: { color: C.ink } });
  addText(s, 'Открыть официальный канал', 1.53, 5.94, 2.86, 0.27, { fontSize: 8.5, color: C.text, bold: true, align: 'center' });

  pill(s, 'ДЕМО-СЦЕНАРИЙ', 5.72, 1.93, 1.65, C.line, C.text);
  const result = [
    ['01', 'Уровень риска', 'Не процент ради процента, а понятная категория.'],
    ['02', 'Почему', 'Признаки манипуляции выделены простым языком.'],
    ['03', 'Что делать', 'Официальный и безопасный следующий шаг.']
  ];
  result.forEach((r, i) => {
    const y = 2.6 + i * 1.25;
    addText(s, r[0], 5.75, y, 0.48, 0.28, { fontSize: 9, color: C.cyan, bold: true });
    addText(s, r[1], 6.42, y - 0.08, 2.7, 0.4, { fontSize: 18, bold: true });
    addText(s, r[2], 6.42, y + 0.35, 4.8, 0.55, { fontSize: 11.5, color: C.muted, valign: 'top' });
  });
  box(s, 5.72, 6.04, 6.15, 0.58, '0D2B2D', 0.14, C.cyan);
  addText(s, 'Приватность: пользователь сам выбирает, что отправить на анализ.', 5.97, 6.15, 5.68, 0.3, { fontSize: 11, color: C.text, bold: true });
  addNotes(s, 'Так выглядит ключевой экран первого MVP. Пользователь вставляет подозрительное сообщение. SAQ показывает не магический процент, а понятную категорию риска, объясняет найденные признаки и даёт безопасный следующий шаг — например, открыть официальный канал организации. В MVP анализ запускается только по инициативе пользователя: никакого скрытого чтения переписок.');
}

// 7 — Flow
{
  const s = slideBase('Как это работает', 'От сомнения до безопасного действия — за четыре шага', 7);
  const steps = [
    ['01', 'Передать', 'Текст · ссылка · скриншот', C.blue],
    ['02', 'Проверить', 'Контекст + признаки давления', C.cyan],
    ['03', 'Объяснить', 'Почему ситуация рискованна', C.yellow],
    ['04', 'Действовать', 'Официальный безопасный канал', C.green]
  ];
  steps.forEach((st, i) => {
    const x = 0.72 + i * 3.05;
    box(s, x, 2.12, 2.62, 3.72, i === 3 ? '0F2923' : C.panel, 0.17, i === 3 ? C.green : C.line);
    pill(s, st[0], x + 0.24, 2.42, 0.55, st[3], C.ink);
    addText(s, st[1], x + 0.24, 3.18, 2.0, 0.47, { fontSize: 22, bold: true });
    addText(s, st[2], x + 0.24, 3.86, 2.12, 0.83, { fontSize: 12, color: C.muted, valign: 'top' });
    if (i < 3) {
      line(s, x + 2.64, 3.95, 0.35, 0, C.cyan, 2);
      s.addShape(ST.chevron, { x: x + 2.9, y: 3.83, w: 0.18, h: 0.24, fill: { color: C.cyan }, line: { color: C.cyan } });
    }
  });
  addText(s, 'Целевой результат MVP', 0.78, 6.26, 2.4, 0.25, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1.1 });
  addText(s, 'Пользователь понимает риск и знает, что делать дальше.', 3.07, 6.15, 8.6, 0.46, { fontSize: 18, bold: true });
  addNotes(s, 'Путь максимально короткий. Пользователь передаёт только то, что хочет проверить. Система извлекает текст и признаки, анализирует манипуляцию, объясняет вывод и ведёт к безопасному действию. Для конкурса этот сценарий можно показать вживую на двух примерах: явном мошенничестве и нейтральном сообщении, чтобы продемонстрировать не только тревогу, но и аккуратность.');
}

// 8 — Tech
{
  const s = slideBase('Технология', 'Гибридный AI: модель + правила + проверяемые сигналы', 8);
  const cols = [
    ['INPUT', 'Текст\nOCR скриншота\nURL', C.blue],
    ['SIGNALS', 'Срочность\nАвторитет\nЗапрос секрета', C.yellow],
    ['REASONING', 'Классификатор\nLLM-объяснение\nУровень уверенности', C.cyan],
    ['OUTPUT', 'Риск\nПричины\nБезопасный шаг', C.green]
  ];
  cols.forEach((c, i) => {
    const x = 0.72 + i * 3.08;
    box(s, x, 2.0, 2.68, 2.55, C.panel);
    addText(s, c[0], x + 0.24, 2.27, 2.2, 0.3, { fontSize: 9, color: c[3], bold: true, charSpacing: 1.2 });
    addText(s, c[1], x + 0.24, 2.88, 2.14, 1.15, { fontSize: 15, color: C.text, bold: true, breakLine: false, valign: 'top', paraSpaceAfterPt: 7 });
    if (i < 3) line(s, x + 2.7, 3.27, 0.32, 0, C.line, 2);
  });
  addText(s, 'Guardrails', 0.78, 4.98, 1.5, 0.35, { fontSize: 10, color: C.cyan, bold: true, charSpacing: 1.3 });
  const guards = [
    ['НЕОПРЕДЕЛЁННОСТЬ', 'Не выдавать оценку за гарантию'],
    ['PRIVACY BY DESIGN', 'Минимум данных, удаление, согласие'],
    ['HUMAN-IN-THE-LOOP', 'Эскалация к официальной поддержке'],
    ['EVALUATION', 'Тесты RU/KZ, precision/recall, false alarms']
  ];
  guards.forEach((g, i) => {
    const x = 0.78 + i * 3.08;
    line(s, x, 5.54, 2.55, 0, C.line, 1);
    addText(s, g[0], x, 5.72, 2.55, 0.27, { fontSize: 8, color: C.muted, bold: true, charSpacing: 0.8 });
    addText(s, g[1], x, 6.08, 2.52, 0.49, { fontSize: 10.5, color: C.text, bold: true, valign: 'top' });
  });
  addText(s, 'Архитектура — план MVP; качество должно быть измерено на локальном датасете.', 0.78, 6.72, 9.9, 0.21, { fontSize: 8, color: '647986' });
  addNotes(s, 'Технически SAQ не должен зависеть от одной большой модели. Реалистичный MVP — гибрид: OCR для скриншотов, правила для явных сигналов, классификатор риска и LLM для понятного объяснения. Критично измерять не только точность, но и ложные тревоги, отдельно на русском и казахском. Эта архитектура — план, а не утверждение о готовой системе.');
}

// 9 — Positioning
{
  const s = slideBase('Позиционирование', 'Наша ставка — контекст + объяснение + действие', 9);
  const x0 = 0.75, y0 = 2.02;
  const widths = [3.15, 1.65, 1.65, 1.65, 1.65, 1.65];
  const heads = ['КАТЕГОРИЯ', 'Контекст\nобщения', 'Репутация\nномера', 'Проверка\nURL', 'Объяснение\nриска', 'Следующий\nшаг'];
  let x = x0;
  heads.forEach((h, i) => {
    addText(s, h, x, y0, widths[i], 0.62, { fontSize: i === 0 ? 9 : 8, color: i === 0 ? C.muted : C.text, bold: true, align: i === 0 ? 'left' : 'center', breakLine: false });
    x += widths[i];
  });
  const rows = [
    ['Банковский антифрод', '△', '—', '△', '△', '△'],
    ['Caller ID', '—', '●', '—', '△', '—'],
    ['Safe Browsing', '—', '—', '●', '△', '△'],
    ['SAQ · целевая модель', '●', '△', '●', '●', '●']
  ];
  rows.forEach((r, ri) => {
    const y = y0 + 0.76 + ri * 0.82;
    if (ri === 3) box(s, x0 - 0.14, y - 0.08, 11.55, 0.72, '0D2B2D', 0.12, C.cyan);
    let xx = x0;
    r.forEach((v, ci) => {
      const color = v === '●' ? (ri === 3 ? C.cyan : C.green) : v === '△' ? C.yellow : '536572';
      addText(s, v, xx, y, widths[ci], 0.42, { fontSize: ci === 0 ? 13 : 15, color: ci === 0 ? C.text : color, bold: ci === 0 || v === '●', align: ci === 0 ? 'left' : 'center' });
      xx += widths[ci];
    });
    if (ri < 3) line(s, x0, y + 0.59, 11.35, 0, C.line);
  });
  addText(s, '● основной фокус', 0.78, 6.27, 1.55, 0.24, { fontSize: 9, color: C.green });
  addText(s, '△ частично / зависит от продукта', 2.55, 6.27, 2.9, 0.24, { fontSize: 9, color: C.yellow });
  addText(s, '— не основной фокус', 5.72, 6.27, 2.25, 0.24, { fontSize: 9, color: C.muted });
  addText(s, 'Категорийное сравнение; конкретные продукты требуют отдельной верификации.', 0.78, 6.72, 9.0, 0.2, { fontSize: 8, color: '647986' });
  addNotes(s, 'Мы не заявляем, что конкуренты ничего не умеют. У банков, Kaspersky, Getcontact, Truecaller и браузеров сильные технологии и большие данные. Но их основной фокус другой. Целевая модель SAQ — соединить контекст общения, URL, объяснение и безопасное действие в независимом пользовательском продукте с локальной адаптацией. Отдельные функции можно скопировать; экосистему доверия — сложнее.');
}

// 10 — Market
{
  const s = slideBase('Возможность', 'Масштаб проблемы огромен. Начинать нужно узко.', 10);
  box(s, 0.75, 1.95, 5.35, 3.75, C.panel);
  addText(s, '$12.5B', 1.03, 2.23, 4.6, 0.92, { fontSize: 48, color: C.cyan, bold: true, charSpacing: -2 });
  addText(s, 'заявленных потерь от мошенничества в США за 2024 год', 1.05, 3.28, 4.3, 0.74, { fontSize: 16, bold: true, valign: 'top' });
  pill(s, '+25% ГОД К ГОДУ', 1.05, 4.34, 1.9, C.red, C.ink);
  addText(s, 'Это только заявленные потери — не размер рынка SAQ.', 1.05, 4.93, 4.5, 0.42, { fontSize: 10.5, color: C.muted });

  addText(s, 'Стратегия выхода', 6.75, 2.0, 2.8, 0.42, { fontSize: 19, bold: true });
  const rings = [
    ['TAM', 'Мировая цифровая безопасность потребителей', 4.2, C.line],
    ['SAM', 'Русско- и казахоязычные пользователи Центральной Азии', 3.05, '16424A'],
    ['SOM', 'Первые семьи и молодые пользователи в Казахстане', 1.85, '0D4A4A']
  ];
  rings.forEach((r, i) => {
    s.addShape(ST.ellipse, { x: 8.4 + (4.2 - r[2]) / 2, y: 1.9 + (4.2 - r[2]) / 2, w: r[2], h: r[2], fill: { color: r[3], transparency: i === 0 ? 30 : 6 }, line: { color: i === 2 ? C.cyan : C.line, width: i === 2 ? 2 : 1 } });
  });
  addText(s, 'SOM', 9.9, 3.02, 1.2, 0.28, { fontSize: 10, color: C.cyan, bold: true, align: 'center' });
  addText(s, 'Казахстан', 9.55, 3.42, 1.9, 0.36, { fontSize: 15, bold: true, align: 'center' });
  addText(s, 'Проверяем спрос до расчёта выручки', 7.2, 6.0, 4.8, 0.35, { fontSize: 12, color: C.muted, align: 'center' });
  addText(s, 'Источники: FTC, Consumer Sentinel Network, 10.03.2025; НБК — Национальный антифрод-центр работает с июля 2024.', 0.78, 6.72, 11.55, 0.22, { fontSize: 7.5, color: '647986' });
  addNotes(s, 'Для масштаба: FTC сообщила о 12,5 миллиарда долларов заявленных потерь от мошенничества в США в 2024 году — на 25 процентов больше, чем годом ранее. Это не размер рынка SAQ, а индикатор проблемы. В Казахстане уже работает Национальный антифрод-центр. Наша стратегия — не начинать «со всего мира», а проверить продукт на понятном локальном сегменте и только затем считать SAM и выручку. Источник FTC: New FTC Data Show a Big Jump in Reported Losses to Fraud to $12.5 Billion in 2024, 10 марта 2025 года.');
}

// 11 — MVP
{
  const s = slideBase('MVP', '8 недель до честной проверки гипотезы', 11);
  const weeks = [
    ['1–2', 'Исследование', '20 интервью\n30–50 сценариев', C.blue],
    ['3–4', 'Прототип', 'Текстовый ввод\nПравила + AI', C.cyan],
    ['5–6', 'Оценка', 'RU/KZ датасет\nЛожные тревоги', C.yellow],
    ['7–8', 'Пилот', '50–100 тестеров\nОбратная связь', C.green]
  ];
  weeks.forEach((w, i) => {
    const x = 0.72 + i * 3.05;
    box(s, x, 2.0, 2.66, 2.6, C.panel);
    pill(s, `НЕДЕЛИ ${w[0]}`, x + 0.22, 2.28, 1.12, w[3], C.ink);
    addText(s, w[1], x + 0.22, 2.95, 2.16, 0.4, { fontSize: 19, bold: true });
    addText(s, w[2], x + 0.22, 3.56, 2.1, 0.72, { fontSize: 12, color: C.muted, valign: 'top', breakLine: false });
  });
  const scope = [
    ['СЕЙЧАС', 'Текст · ссылка · скриншот · риск · объяснение · действие', C.cyan],
    ['ПОТОМ', 'История с согласием · Family Shield · threat signals', C.blue],
    ['НЕ В MVP', 'Авточтение переписок · скрытый мониторинг · «100% защита»', C.red]
  ];
  scope.forEach((r, i) => {
    const y = 5.0 + i * 0.55;
    addText(s, r[0], 0.82, y, 1.18, 0.28, { fontSize: 8, color: r[2], bold: true, charSpacing: 1.1 });
    addText(s, r[1], 2.02, y - 0.05, 9.85, 0.38, { fontSize: 11.5, color: i === 2 ? C.muted : C.text, bold: i === 0 });
  });
  addNotes(s, 'Первый MVP можно построить за восемь недель даже начинающему разработчику, если жёстко удерживать фокус. Сначала интервью и набор сценариев, затем простой веб-прототип, тестирование на русском и казахском и небольшой пилот. В MVP не входят скрытый мониторинг, автоматическое чтение мессенджеров и обещание стопроцентной защиты. Главный вопрос: помогает ли объяснение человеку принять более безопасное решение?');
}

// 12 — Business model
{
  const s = slideBase('Бизнес-модель', 'Бесплатный момент доверия → платная постоянная защита', 12);
  const tiers = [
    ['B2C FREEMIUM', '0 → подписка', 'Базовые проверки бесплатно\nРасширенная история и лимиты', C.cyan],
    ['FAMILY PLAN', 'общая ценность', 'Добровольная помощь близким\nБез скрытого доступа', C.green],
    ['B2B API', 'масштаб', 'Анализ коммуникаций\nИнтеграции с платформами', C.blue]
  ];
  tiers.forEach((t, i) => {
    const x = 0.75 + i * 4.08;
    box(s, x, 2.02, 3.65, 3.38, C.panel);
    addText(s, t[0], x + 0.28, 2.32, 2.7, 0.3, { fontSize: 9, color: t[3], bold: true, charSpacing: 1.1 });
    addText(s, t[1], x + 0.28, 3.0, 2.9, 0.44, { fontSize: 22, bold: true });
    addText(s, t[2], x + 0.28, 3.74, 2.93, 0.92, { fontSize: 12, color: C.muted, valign: 'top', breakLine: false });
    line(s, x + 0.28, 4.9, 1.0, 0, t[3], 3);
  });
  box(s, 0.75, 5.72, 11.82, 0.75, '0D161D', 0.14, C.line);
  addText(s, 'Экономика — гипотеза', 1.05, 5.93, 2.2, 0.28, { fontSize: 10, color: C.yellow, bold: true });
  addText(s, 'Измерять: стоимость AI-проверки · конверсию в оплату · удержание после первого риска', 3.05, 5.87, 8.9, 0.4, { fontSize: 12.5, color: C.text, bold: true });
  addText(s, 'B2G — возможное долгосрочное направление, не заявленное партнёрство.', 0.78, 6.72, 9.2, 0.2, { fontSize: 8, color: '647986' });
  addNotes(s, 'Монетизация строится постепенно. Базовая проверка должна быть доступна бесплатно, потому что доверие нужно заслужить. Подписка может давать дополнительные лимиты, историю и семейные функции. B2B API — более поздний источник масштаба. Пока это гипотезы. На пилоте нужно измерить стоимость одной проверки, готовность платить и возврат пользователя после первого полезного результата.');
}

// 13 — GTM / roadmap
{
  const s = slideBase('Выход на рынок', 'Сначала доказать доверие. Затем масштабировать защиту.', 13);
  const phases = [
    ['0–3 мес.', 'КАЗАХСТАН', 'Интервью · MVP · 50–100 тестеров', 'Критерий: полезность объяснения', C.cyan],
    ['3–9 мес.', 'ЛОКАЛЬНЫЙ ПРОДУКТ', 'RU/KZ качество · школы · семьи · контент', 'Критерий: повторное использование', C.green],
    ['9–18 мес.', 'ЦЕНТРАЛЬНАЯ АЗИЯ', 'Новые языки · сценарии · пилоты B2B', 'Критерий: воспроизводимый канал', C.blue]
  ];
  phases.forEach((p, i) => {
    const y = 1.95 + i * 1.42;
    addText(s, p[0], 0.78, y, 1.25, 0.3, { fontSize: 10, color: p[4], bold: true });
    line(s, 2.05, y + 0.16, 0.95, 0, p[4], 3);
    addText(s, p[1], 3.18, y - 0.08, 2.4, 0.4, { fontSize: 17, bold: true });
    addText(s, p[2], 5.45, y - 0.08, 3.25, 0.46, { fontSize: 12, color: C.text, bold: true });
    addText(s, p[3], 9.15, y - 0.08, 3.0, 0.48, { fontSize: 10.5, color: C.muted, valign: 'top' });
    if (i < 2) line(s, 3.18, y + 0.72, 8.98, 0, C.line);
  });
  box(s, 0.78, 6.15, 11.35, 0.53, '0D2B2D', 0.13, C.cyan);
  addText(s, 'Долгосрочно: Family Shield → Threat Network → privacy-preserving AI on-device', 1.05, 6.26, 10.82, 0.3, { fontSize: 12, color: C.text, bold: true, align: 'center' });
  addNotes(s, 'Go-to-market начинается не с дорогой рекламы, а с доверия и реальных сценариев. В Казахстане первые каналы — интервью, школы, семьи, образовательный контент и сообщества цифровой грамотности. Переход к следующему этапу зависит от метрики, а не от календаря: сначала полезность объяснения, затем повторное использование, потом воспроизводимый канал роста.');
}

// 14 — Founder / ask
{
  const s = pptx.addSlide('SAQ_MASTER');
  pill(s, 'VISION + ASK', 0.75, 0.62, 1.35);
  addText(s, 'Сделать паузу перед опасным действием — новой цифровой привычкой.', 0.75, 1.28, 10.9, 1.22, { fontSize: 36, bold: true, breakLine: false });
  box(s, 0.75, 3.12, 5.25, 2.35, C.panel);
  addText(s, 'ОСНОВАТЕЛЬ', 1.05, 3.42, 2.0, 0.28, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1.2 });
  addText(s, 'Ученик 11 класса из Казахстана', 1.05, 4.0, 4.3, 0.48, { fontSize: 20, bold: true });
  addText(s, 'Математика · информатика · AI · предпринимательство', 1.05, 4.62, 4.42, 0.42, { fontSize: 11.5, color: C.muted });

  box(s, 6.3, 3.12, 6.0, 2.35, '0D2B2D', 0.16, C.cyan);
  addText(s, 'ЧТО НУЖНО СЕЙЧАС', 6.62, 3.42, 2.4, 0.28, { fontSize: 9, color: C.cyan, bold: true, charSpacing: 1.2 });
  const asks = ['50–100 тестовых пользователей', 'Ментор по AI / cybersecurity', 'Экспертная проверка сценариев'];
  asks.forEach((a, i) => {
    dot(s, 6.65, 4.03 + i * 0.45, 0.13, C.cyan);
    addText(s, a, 6.95, 3.92 + i * 0.45, 4.75, 0.34, { fontSize: 13, bold: true });
  });
  addText(s, 'SAQ', 0.75, 6.15, 2.0, 0.58, { fontSize: 31, bold: true });
  addText(s, 'Think Before You Trust.', 2.45, 6.2, 4.6, 0.48, { fontSize: 20, color: C.cyan, bold: true });
  addText(s, 'Спасибо', 10.75, 6.24, 1.4, 0.34, { fontSize: 13, color: C.muted, align: 'right' });
  addNotes(s, 'Я ученик 11 класса из Казахстана и не утверждаю, что уже построил большую систему. Но я понимаю первый реалистичный шаг и готов его реализовать. Сейчас мне нужны 50–100 тестовых пользователей, ментор по AI и кибербезопасности и экспертная проверка сценариев. Видение SAQ простое: сделать паузу перед опасным действием новой цифровой привычкой. Think Before You Trust.');
}

pptx.writeFile({ fileName: '/workspace/outputs/SAQ_Pitch_Deck_RU.pptx', compression: true });
