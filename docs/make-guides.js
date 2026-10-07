const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, LevelFormat, Footer, PageNumber, HeightRule, VerticalAlign
} = require('docx');

const OUT = process.argv[2];
const FONT = { ascii: 'Yu Gothic', eastAsia: 'Yu Gothic', hAnsi: 'Yu Gothic', cs: 'Yu Gothic' };
const GREEN = '1F7A4D', GREEN_BG = 'E2F1E8', GRAY = '55645A', LINE = 'C9D3CB', WARN_BG = 'FCEFD9';
const PAGE_W = 11906, MARGIN = 1000, CONTENT_W = PAGE_W - MARGIN * 2; // A4, 約1.76cm余白

/* ---------- 部品 ---------- */
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 20, bold: o.bold, color: o.color });
// 「**太字**」記法を TextRun に分解
function rich(str, o = {}) {
  return str.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(s =>
    s.startsWith('**') ? t(s.slice(2, -2), { ...o, bold: true }) : t(s, o));
}
const p = (str, o = {}) => new Paragraph({ children: rich(str, o), spacing: { after: o.after ?? 80, line: 300 }, alignment: o.align });
const title = (main, sub) => [
  new Paragraph({ children: [t(main, { size: 34, bold: true, color: GREEN })], spacing: { after: 40 } }),
  new Paragraph({
    children: [t(sub, { size: 20, color: GRAY })], spacing: { after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: GREEN, space: 6 } }
  })
];
const h = (num, text) => new Paragraph({
  children: [t(num ? num + '　' : '', { size: 24, bold: true, color: GREEN }), t(text, { size: 24, bold: true })],
  spacing: { before: 180, after: 70 }, keepNext: true,
  border: { left: { style: BorderStyle.SINGLE, size: 24, color: GREEN, space: 8 } }
});
const bullet = (str) => new Paragraph({ children: rich(str), numbering: { reference: 'bul', level: 0 }, spacing: { after: 40, line: 290 } });
let stepRef = 0;
function steps(list) {
  const ref = 'step' + (stepRef++);
  numberingConfigs.push({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 300 } }, run: { font: FONT, bold: true, color: GREEN } } }] });
  return list.map(s => new Paragraph({ children: rich(s), numbering: { reference: ref, level: 0 }, spacing: { after: 50, line: 290 } }));
}
const numberingConfigs = [
  { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '・', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 300 } } } }] }
];

const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(content, width, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(c =>
    typeof c === 'string' ? new Paragraph({ children: rich(c, { bold: o.bold, size: o.size }), spacing: { after: 20, line: 280 }, alignment: o.align }) : c);
  return new TableCell({
    children: paras, width: { size: width, type: WidthType.DXA }, borders, verticalAlign: o.vAlign || VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    margins: { top: 70, bottom: 70, left: 110, right: 110 }, columnSpan: o.span
  });
}
// widths は比率で渡し、本文幅に合わせて DXA に変換
function table(ratios, header, rows, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const w = ratios.map(r => Math.floor(CONTENT_W * r / sum));
  w[w.length - 1] += CONTENT_W - w.reduce((a, b) => a + b, 0);
  const tr = [];
  if (header) tr.push(new TableRow({ tableHeader: true, children: header.map((c, i) => cell(c, w[i], { bold: true, fill: GREEN_BG })) }));
  rows.forEach(r => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, w[i], { fill: o.firstColFill && i === 0 ? 'F4F7F4' : undefined, bold: o.firstColBold && i === 0 })) })));
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: w, rows: tr }), spacer(60)];
}
// 書き込み欄・注意書きなどの囲み
function box(lines, fill, o = {}) {
  return [new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({
      height: o.height ? { value: o.height, rule: HeightRule.ATLEAST } : undefined,
      children: [new TableCell({
        children: lines.map(l => typeof l === 'string' ? new Paragraph({ children: rich(l), spacing: { after: 60, line: 300 } }) : l),
        width: { size: CONTENT_W, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
        borders: { top: { style: BorderStyle.SINGLE, size: 8, color: o.color || GREEN }, bottom: { style: BorderStyle.SINGLE, size: 8, color: o.color || GREEN }, left: { style: BorderStyle.SINGLE, size: 8, color: o.color || GREEN }, right: { style: BorderStyle.SINGLE, size: 8, color: o.color || GREEN } },
        shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined,
        margins: { top: 120, bottom: 120, left: 180, right: 180 }
      })]
    })]
  }), spacer(60)];
}
const spacer = (after = 120) => new Paragraph({ children: [], spacing: { after } });
// 記入欄（URL・QRコード）
function fillInRow(leftLines, rightLabel) {
  const wR = 2200, wL = CONTENT_W - wR;
  const b = { style: BorderStyle.SINGLE, size: 8, color: GREEN };
  return [new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [wL, wR],
    rows: [new TableRow({
      height: { value: 1900, rule: HeightRule.ATLEAST },
      children: [
        new TableCell({ children: leftLines.map(l => new Paragraph({ children: rich(l), spacing: { after: 140, line: 300 } })), width: { size: wL, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b }, margins: { top: 140, bottom: 100, left: 180, right: 180 }, verticalAlign: VerticalAlign.CENTER }),
        new TableCell({ children: [new Paragraph({ children: [t(rightLabel, { size: 16, color: GRAY })], alignment: AlignmentType.CENTER })], width: { size: wR, type: WidthType.DXA }, borders: { top: b, bottom: b, left: { style: BorderStyle.DASHED, size: 8, color: GREEN }, right: b }, verticalAlign: VerticalAlign.CENTER })
      ]
    })]
  }), spacer(60)];
}

function makeDoc(children, footerText) {
  return new Document({
    creator: 'マンホール商会',
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    numbering: { config: numberingConfigs },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 900, bottom: 900, left: MARGIN, right: MARGIN } } },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [t(footerText + '　', { size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GRAY })]
          })]
        })
      },
      children
    }]
  });
}

const CRIT_ROWS = [
  ['血圧', '上（収縮期）**130未満** かつ 下（拡張期）**85未満**（mmHg）'],
  ['BMI', '**18.5以上 25未満**　※BMI＝体重(kg)÷身長(m)÷身長(m)。アプリが自動で計算します'],
  ['体重', '前回の判定測定（半年前）から **3%以上減少**　※初めての判定のときは「判定なし」']
];

/* ================================================================== */
/* 従業員向け                                                          */
/* ================================================================== */
const employee = [
  ...title('健康記録アプリの使い方', '従業員のみなさんへ ｜ マンホール商会 健康経営の取り組み'),
  p('会社では、みなさんに長く健康で働いてもらうため、**血圧・BMI・体重**の記録に取り組んでいます。スマートフォンのアプリで測定値を入力するだけで、紙の提出は不要です。'),
  ...box([
    '**毎年4月と10月**に、血圧と体重を測ってアプリに記録してください。',
    '血圧・BMI・体重のうち**1つでも基準をクリア**すれば、**健康手当の対象**になります。'
  ], GREEN_BG),

  h('1', 'アプリを開く'),
  ...fillInRow([
    'アプリのURL：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿',
    'あなたの社員番号：＿＿＿＿＿＿＿＿　　最初の暗証番号：＿＿＿＿＿＿＿＿'
  ], 'QRコードを貼る'),
  ...steps([
    'スマートフォンで上のURLを開く（QRコードを読み取ってもOK）。',
    '**社員番号**と**暗証番号**を入力して「ログイン」を押す。',
    'はじめてログインしたら、「**暗証番号の変更**」タブで自分だけの番号（4〜8桁の数字）に変える。',
    'ホーム画面に追加しておくと、次から1タップで開けます。iPhone：Safariの共有ボタン →「ホーム画面に追加」／Android：Chromeの「︙」→「ホーム画面に追加」'
  ]),

  h('2', '測定値を入力する'),
  ...steps([
    '「自分の記録」タブの「**測定値を入力**」に、測定日・血圧（上と下）・体重を入れる。',
    '身長は最初の1回だけ入力すれば、次からは自動で入ります。',
    '入力中に、BMIと基準をクリアしているかがその場で表示されます。4月・10月の日付なら「**判定に使われます**」と表示されます。',
    '「**記録する**」を押して完了。'
  ]),
  ...table([1.1, 4], null, [
    ['血圧の測り方', '朝起きて1時間以内、トイレのあと・朝食や薬の前に、いすに座って1〜2分落ち着いてから測ります。'],
    ['体重の測り方', '朝、トイレのあと・朝食の前など、毎回できるだけ同じ条件で測ります。']
  ], { firstColFill: true, firstColBold: true }),
  p('4月・10月以外の月も自由に記録できます（判定には使いません）。記録を続けると、推移をグラフで見られます。', { color: GRAY }),

  h('3', '結果を見る'),
  p('「自分の記録」タブの上に、直近の判定結果が表示されます。'),
  ...table([1.3, 4], ['表示', '意味'], [
    ['○ 手当の対象です', '3つの項目のうち、1つ以上クリアしています'],
    ['クリア／未達', '各項目が基準を満たしているかどうか'],
    ['判定なし', '体重の比較に使う前回の判定測定がない（初めての判定など）'],
    ['受診をすすめる表示', '血圧が140/90以上のとき。日を変えて測り直し、続くようなら医療機関や産業医に相談してください（手当の判定とは関係ありません）']
  ], { firstColBold: true }),

  h('4', '手当の基準（どれか1つクリアで対象）'),
  ...table([1, 5], ['項目', 'クリアの条件'], CRIT_ROWS, { firstColBold: true }),
  p('基準は、厚生労働省の特定健診の判定値などを参考にしています。', { color: GRAY }),

  h('5', 'よくある質問'),
  ...table([2, 4], ['こんなとき', 'どうする'], [
    ['入力をまちがえた', '同じ測定日で正しい値をもう一度記録してください。あとから記録した値が使われます。'],
    ['判定月に何回か測った', 'その月で**いちばん新しい測定日**の値が判定に使われます。'],
    ['暗証番号を忘れた', '下の問い合わせ先に連絡してください。新しい番号を発行します。'],
    ['ログインできなくなった', '暗証番号を5回続けて間違えると、30分間ログインできません。時間をおいて試してください。'],
    ['自分のデータは誰が見るの？', '見られるのは本人と、管理を担当する経理担当者だけです。健康づくりと手当の判定のために使います。']
  ], { firstColBold: true }),

  ...box([
    '**困ったときの問い合わせ先**',
    '担当：＿＿＿＿＿＿＿＿＿＿＿＿　　内線・連絡先：＿＿＿＿＿＿＿＿＿＿＿＿'
  ], null),
  p('※ このアプリの数値は健康管理の目安で、診断ではありません。気になる数値が続くときは、医療機関や産業医に相談してください。', { size: 17, color: GRAY })
];

/* ================================================================== */
/* 経理担当者向け                                                      */
/* ================================================================== */
const admin = [
  ...title('健康記録アプリ 運用マニュアル', '経理担当者（管理者）向け ｜ マンホール商会 健康経営の取り組み'),
  p('従業員がスマートフォンで入力した血圧・体重は、Googleスプレッドシートに自動でたまります。経理担当者は、アプリの「**管理（全員）**」タブか、スプレッドシートで確認できます。'),
  ...box([
    'アプリのURL：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿',
    'スプレッドシートの場所：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿',
    '管理者の社員番号：＿＿＿＿＿＿＿＿　（暗証番号はここに書かないでください）'
  ], null),

  h('1', 'スプレッドシートの3つのシート'),
  ...table([1.2, 4.5], ['シート', '内容'], [
    ['従業員', '社員の名簿。**社員番号・氏名・身長・暗証番号・管理者・在籍**の6列。社員の登録・変更はここで行います。'],
    ['記録', '従業員がアプリで入力した測定値が、1回ごとに1行ずつ自動で追加されます。'],
    ['判定結果', '「判定結果を書き出す」を押したときに、その判定期の全員分の結果が保存されます。給与計算にはこのシートを使います。']
  ], { firstColBold: true }),
  ...box(['**注意：**シート名・1行目の見出し・列の順番は変えないでください。アプリが正しく動かなくなります。'], WARN_BG, { color: 'D9A441' }),

  h('2', '社員の登録・変更'),
  ...table([1.5, 4.5], ['こんなとき', '操作'], [
    ['新しく入社した', '「従業員」シートの一番下に行を追加します。例：0031｜山田 一郎｜168｜5821｜（空欄）｜在籍。決めた社員番号と暗証番号（4〜8桁の数字）を本人に渡してください。'],
    ['退職した', '「在籍」列を「**退職**」にします。行は**削除しないでください**（過去の記録との対応が分からなくなります）。ログインできなくなり、一覧からも外れます。'],
    ['暗証番号を忘れた', '「暗証番号」列のセルに新しい数字を上書きして、本人に伝えます。本人がログインすると、自動で暗号化されて数字は見えなくなります。'],
    ['管理者を追加・交代する', '「管理者」列に「管理者」と入力すると、その人のアプリに「管理（全員）」タブが出ます。外すときは空欄にします。']
  ], { firstColBold: true }),

  h('3', '年間スケジュール'),
  ...table([1.3, 4.5], ['時期', 'やること'], [
    ['3月下旬／9月下旬', '従業員へ「4月（10月）は判定月です。期間中に測って記録してください」と周知する。'],
    ['4月／10月 中旬', 'アプリの管理タブで提出状況を確認し、「未提出のみ表示」で未提出者に声をかける。'],
    ['4月／10月 月末', '提出状況を最終確認する。**判定は4月・10月中に測定した記録だけが対象**です。'],
    ['5月／11月 初め', '「判定結果」シートに書き出し、手当対象者を確定して給与計算へ回す。支給月は社内の規程に従ってください。']
  ], { firstColBold: true }),

  h('4', '判定期の作業手順'),
  ...steps([
    'アプリに管理者の社員番号でログインし、「**管理（全員）**」タブを開く。',
    '上の選択欄で判定期（例：2026年10月判定）を選ぶ。',
    '提出人数・手当対象人数・項目ごとのクリア人数を確認する。行をタップすると、その人の推移を見られます。',
    '「**『判定結果』シートに書き出す**」を押す（何度押しても最新の内容に置き換わります）。スプレッドシートのメニュー「健康記録アプリ」→「今期の判定結果を書き出す」でも同じです。',
    '必要に応じて「CSVをダウンロード」でファイルに保存する（Excelで開けます）。'
  ]),

  h('5', '判定のルール'),
  ...table([1, 5], ['項目', 'クリアの条件（どれか1つで手当の対象）'], CRIT_ROWS, { firstColBold: true }),
  ...steps([
    '同じ判定月に複数の記録があるときは、**測定日がいちばん新しいもの**で判定します。',
    '体重は前回の判定期（4月なら前年10月、10月なら同年4月）の記録と比べます。前回が未提出のときは「判定なし」です。',
    '一覧の「**要受診**」は、血圧が140/90以上の人です。手当とは別に、本人へ受診をすすめる声かけをお願いします。'
  ]),

  h('6', '困ったとき'),
  ...table([2, 4], ['こんなとき', 'どうする'], [
    ['従業員が値を間違えて入力した', '本人に、同じ測定日で正しい値をもう一度記録してもらいます。経理担当者が「記録」シートの数値を直接直しても構いません。'],
    ['判定月を過ぎてから記録した', '5月・11月の記録は任意の記録として扱われ、判定には使われません。特例で認める場合は、「記録」シートのその行の「判定期」列に「2026-04」のように入力します。'],
    ['アプリの画面が開かない・おかしい', 'アプリを導入した担当者に連絡してください（導入手順は health-app リポジトリの README.md にあります）。']
  ], { firstColBold: true }),

  h('7', '個人情報の取り扱い'),
  bullet('血圧・体重などの健康情報は、法律上「**要配慮個人情報**」です。健康づくりと手当の判定以外には使わないでください。'),
  bullet('スプレッドシートの共有は**必要な人だけ**に。印刷物やCSVファイルは保管に気をつけ、不要になったら確実に廃棄してください。')
];

fs.mkdirSync(OUT, { recursive: true });
Promise.all([
  Packer.toBuffer(makeDoc(employee, '健康記録アプリの使い方（従業員向け）')).then(b => fs.writeFileSync(path.join(OUT, '健康記録アプリの使い方_従業員向け.docx'), b)),
  Packer.toBuffer(makeDoc(admin, '健康記録アプリ 運用マニュアル（経理担当者向け）')).then(b => fs.writeFileSync(path.join(OUT, '健康記録アプリ_運用マニュアル_経理担当者向け.docx'), b))
]).then(() => console.log('done'));
