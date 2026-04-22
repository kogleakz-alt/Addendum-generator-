const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, UnderlineType } = require('docx');

const FONT = 'Times New Roman';
const SZ = 22; // 11pt

function noBorders() {
  const nil = { style: BorderStyle.NIL };
  return { top: nil, bottom: nil, left: nil, right: nil };
}

function makeCell(paragraphs, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: noBorders(),
    margins: { top: 0, bottom: 0, left: 108, right: 108 },
    children: paragraphs,
  });
}

function makePara(text, opts = {}) {
  const lines = (text || '').split('\n');
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: 0, after: 40 },
    children: lines.flatMap((line, i) => [
      ...(i > 0 ? [new TextRun({ break: 1 })] : []),
      new TextRun({
        text: line,
        font: FONT,
        size: SZ,
        bold: opts.bold || false,
        underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
      }),
    ]),
  });
}

function buildTable(d) {
  const rows = [];

  function addRow(enParas, ruParas) {
    rows.push(new TableRow({
      children: [
        makeCell(enParas, 4381),
        makeCell([makePara('')], 345),
        makeCell(ruParas, 4562),
      ],
    }));
  }

  // Title
  addRow(
    (d.title_en||'').split('\n').map(l => makePara(l, { bold: true })),
    (d.title_ru||'').split('\n').map(l => makePara(l, { bold: true, align: AlignmentType.CENTER }))
  );

  // Date
  addRow([makePara(d.date_en||'')], [makePara(d.date_ru||'')]);

  function secRow(en, ru) {
    addRow([makePara('  ' + en, { underline: true })], [makePara('  ' + ru, { underline: true })]);
  }
  function contentRow(en, ru) {
    addRow([makePara(en||'')], [makePara(ru||'')]);
  }

  secRow('1. GOODS', '1. ТОВАР');
  contentRow((d.goods_en||'') + (d.lot ? '\n' + d.lot : ''), d.goods_ru||'');
  contentRow(d.origin_en||'', d.origin_ru||'');
  secRow('2. QUANTITY', '2. КОЛИЧЕСТВО');
  contentRow(d.qty_en||'', d.qty_ru||'');
  secRow('3. PRICE', '3. ЦЕНА');
  contentRow(d.price_en||'', d.price_ru||'');
  secRow('4. TOTAL VALUE', '4. ОБЩАЯ СТОИМОСТЬ');
  contentRow(d.total_en||'', d.total_ru||'');
  secRow('5. PAYMENT TERMS', '5. УСЛОВИЯ ОПЛАТЫ');
  contentRow(d.payment_en||'', d.payment_ru||'');
  secRow('6. PACKAGING', '6. УПАКОВКА');
  contentRow(d.pack_en||'', d.pack_ru||'');
  secRow('7. DELIVERY PERIOD and TERMS', '7. СРОК и УСЛОВИЯ ПОСТАВКИ');
  contentRow(d.delivery_en||'', d.delivery_ru||'');

  return new Table({
    width: { size: 9288, type: WidthType.DXA },
    columnWidths: [4381, 345, 4562],
    rows,
  });
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  try {
    const data = JSON.parse(event.body);
    const table = buildTable(data);

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1134, right: 850, bottom: 1134, left: 1134 },
          },
        },
        children: [table],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    const num = data.num || 'X';
    const contract = (data.contract || '').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30);

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="Addendum_N${num}_${contract}.docx"`,
      },
      body: buffer.toString('base64'),
      isBase64Encoded: true,
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
