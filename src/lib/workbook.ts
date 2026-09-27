import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import catalog from '../data/catalog.json';
import sourceStyle from '../data/workbook-style.json';
import sourceTheme from '../data/workbook-theme.json';
import sourceCellStyles from '../data/workbook-cell-styles.json';
import type { Product } from '../types';

type SheetEntry =
  | { kind: 'category'; category: string }
  | { kind: 'product'; product: Product };

const workbookMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const originalHeights = sourceStyle.originalRowHeights as Record<string, number>;
const originalBlankPriceText = sourceStyle.originalBlankPriceText as Record<string, string>;
const originalCategoryRows = sourceStyle.originalCategoryRows as Record<string, number>;
const extraCategoryCells = sourceStyle.originalCategoryExtraCells as Record<string, Record<string, string>>;
const columnLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
const sourceRows = sourceCellStyles.rows as Record<string, string[]>;
const sourceStyles = sourceCellStyles.styles as unknown as Record<string, Partial<ExcelJS.Style>>;
const originalProducts = catalog.products as Product[];

function styleRowForNewProduct(category: string): number {
  const sameCategory = originalProducts.find(
    (item) =>
      item.category === category &&
      item.originalRow &&
      !sourceStyle.highlightedProductSourceRows.includes(item.originalRow),
  );
  return sameCategory?.originalRow ?? 8;
}

function applySourceStyle(cell: ExcelJS.Cell, sourceRow: number, columnNumber: number): void {
  const styleId = sourceRows[String(sourceRow)]?.[columnNumber - 1];
  const style = styleId && sourceStyles[styleId];
  if (!style) throw new Error(`Falta el estilo de origen para ${columnLetters[columnNumber - 1]}${sourceRow}.`);
  // Each cell gets its own object: ExcelJS can update a style while serializing the table.
  cell.style = structuredClone(style);
}

function inSourceRanges(row: number | undefined, ranges: Array<[number, number]>): boolean {
  return row !== undefined && ranges.some(([start, end]) => row >= start && row <= end);
}

function contiguousRanges(rows: number[], column: string): string[] {
  const sorted = [...new Set(rows)].sort((a, b) => a - b);
  const ranges: string[] = [];
  for (let index = 0; index < sorted.length;) {
    const start = sorted[index];
    let end = start;
    while (index + 1 < sorted.length && sorted[index + 1] === end + 1) {
      end = sorted[++index];
    }
    ranges.push(start === end ? `${column}${start}` : `${column}${start}:${column}${end}`);
    index++;
  }
  return ranges;
}

function parseDate(value: string) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) throw new Error('La fecha debe tener el formato AAAA-MM-DD.');
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const parsed = new Date(Date.UTC(year, month - 1, day));
  if (
    parsed.getUTCFullYear() !== year ||
    parsed.getUTCMonth() + 1 !== month ||
    parsed.getUTCDate() !== day
  ) {
    throw new Error('La fecha indicada no existe.');
  }
  return { day: match[3], month: match[2], year: match[1] };
}

function orderedEntries(products: Product[]): SheetEntry[] {
  const groups = new Map<string, Product[]>();
  for (const product of products) {
    const category = product.category || 'Sin categoría';
    const group = groups.get(category) ?? [];
    group.push(product);
    groups.set(category, group);
  }

  const categoryOrder = [...catalog.categories];
  for (const category of groups.keys()) {
    if (!categoryOrder.includes(category)) categoryOrder.push(category);
  }

  return categoryOrder.flatMap((category) => [
    { kind: 'category' as const, category },
    ...(groups.get(category) ?? []).map((product) => ({ kind: 'product' as const, product })),
  ]);
}

function productValues(product: Product): (string | number | null)[] {
  const blank = (value: string): string | null => value === '' ? null : value;
  const sourceBlankPrice = product.originalRow
    ? originalBlankPriceText[String(product.originalRow)]
    : undefined;
  return [
    blank(product.product),
    blank(product.format),
    blank(product.unit),
    product.price ?? sourceBlankPrice ?? null,
    blank(product.supplier),
    blank(product.address),
    blank(product.social),
    blank(product.notes),
  ];
}

function preferredProductHeight(product: Product): number {
  const original = product.originalRow
    ? originalHeights[String(product.originalRow)]
    : undefined;
  const sourceProduct = product.originalRow
    ? originalProducts.find((item) => item.originalRow === product.originalRow)
    : undefined;
  if (
    sourceProduct &&
    product.product === sourceProduct.product &&
    product.format === sourceProduct.format &&
    product.supplier === sourceProduct.supplier &&
    product.address === sourceProduct.address &&
    product.notes === sourceProduct.notes
  ) {
    return original ?? sourceStyle.rowHeights.product;
  }
  const textColumns: Array<[string, number]> = [
    [product.product, 43],
    [product.format, 17],
    [product.supplier, 42],
    [product.address, 44],
    [product.notes, 27],
  ];
  const lines = Math.max(
    1,
    ...textColumns.map(([text, width]) =>
      (text || '').split('\n').reduce((count, line) => count + Math.max(1, Math.ceil(line.length / width)), 0),
    ),
  );
  return Math.max(original ?? sourceStyle.rowHeights.product, lines * 15);
}

function addSourceComments(worksheet: ExcelJS.Worksheet, rowNumber: number, product: Product) {
  for (const [column, comment] of Object.entries(product.sourceComments ?? {})) {
    if (columnLetters.includes(column) && comment) {
      worksheet.getCell(`${column}${rowNumber}`).note = comment;
    }
  }
}

/** Build the dated price list while keeping the source workbook's column order and styling. */
export function createWorkbook(products: Product[], date: string): ExcelJS.Workbook {
  const { day, month, year } = parseDate(date);
  const workbook = new ExcelJS.Workbook();
  // ExcelJS 4.4 has no public theme setter. Its pinned serializer reads this field.
  // Reusing the source theme keeps TableStyleLight11's green bands unchanged.
  (workbook as ExcelJS.Workbook & { _themes?: Record<string, string> })._themes = sourceTheme;
  workbook.creator = 'Listado de precios';
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet(`${day}-${month}-${year}`, {
    properties: { defaultRowHeight: sourceStyle.rowHeights.product },
    pageSetup: { orientation: 'portrait' },
  });

  worksheet.columns = sourceStyle.columnWidths.map((width) => ({ width }));
  worksheet.mergeCells('A1:G1');
  worksheet.mergeCells('A2:G2');
  worksheet.getCell('A1').value = sourceStyle.title;
  worksheet.getCell('A2').value = sourceStyle.subtitle;
  worksheet.getRow(1).height = sourceStyle.rowHeights.title;
  worksheet.getRow(2).height = sourceStyle.rowHeights.subtitle;

  const entries = orderedEntries(products);
  const tableRows = entries.map((entry) =>
    entry.kind === 'category'
      ? [entry.category, null, null, null, null, extraCategoryCells[entry.category]?.F ?? null, null, null]
      : productValues(entry.product),
  );

  worksheet.addTable({
    name: 'ListadoPreciosYEYA4',
    ref: 'A3',
    headerRow: true,
    totalsRow: false,
    style: { theme: sourceStyle.tableStyle as 'TableStyleLight11', showRowStripes: true },
    columns: sourceStyle.headers.map((name) => ({ name, filterButton: true })),
    rows: tableRows,
  });

  const headerRow = worksheet.getRow(3);
  headerRow.height = sourceStyle.rowHeights.header;
  for (let rowNumber = 1; rowNumber <= 3; rowNumber++) {
    for (let columnNumber = 1; columnNumber <= 8; columnNumber++) {
      applySourceStyle(worksheet.getRow(rowNumber).getCell(columnNumber), rowNumber, columnNumber);
    }
  }

  entries.forEach((entry, index) => {
    const rowNumber = index + 4;
    const row = worksheet.getRow(rowNumber);
    const isCategory = entry.kind === 'category';
    const categoryOriginalRow = isCategory ? originalCategoryRows[entry.category] : undefined;
    row.height = isCategory
      ? originalHeights[String(categoryOriginalRow)] ?? sourceStyle.rowHeights.category
      : preferredProductHeight(entry.product);

    const sourceRow = isCategory
      ? categoryOriginalRow ?? 7
      : entry.product.originalRow ?? styleRowForNewProduct(entry.product.category);
    for (let columnNumber = 1; columnNumber <= 8; columnNumber++) {
      applySourceStyle(row.getCell(columnNumber), sourceRow, columnNumber);
    }

    if (isCategory) return;

    addSourceComments(worksheet, rowNumber, entry.product);
  });

  // Keep each original rule on the same logical row when products are inserted.
  // Newly added products receive the same dropdowns, with LITRO available to them.
  const originalUnitRows: number[] = [];
  const newUnitRows: number[] = [];
  const socialRows: number[] = [];
  const conditionalRows: number[] = [];
  entries.forEach((entry, index) => {
    const outputRow = index + 4;
    const sourceRow = entry.kind === 'category'
      ? originalCategoryRows[entry.category]
      : entry.product.originalRow;
    const isNewProduct = entry.kind === 'product' && !entry.product.originalRow;
    if (inSourceRanges(sourceRow, [[7, 87], [89, 104], [106, 121], [124, 124]])) {
      originalUnitRows.push(outputRow);
    } else if (isNewProduct) {
      newUnitRows.push(outputRow);
    }
    if (inSourceRanges(sourceRow, [[7, 87], [89, 121], [124, 124]]) || isNewProduct) {
      socialRows.push(outputRow);
    }
    if (inSourceRanges(sourceRow, [[7, 104], [106, 121], [124, 124]]) || isNewProduct) {
      conditionalRows.push(outputRow);
    }
  });
  const socialFormula = '"WHATSAPP,FACEBOOK,INSTAGRAM,TIKTOK,WEB,OTRO"';
  const unitFormula = '"CAJA,SACO,BOLSA,LATA,ESTUCHE,SET,UNIDAD,TETRA,VIDRIO,TONELADA,POR DEFINIR"';
  // ExcelJS exposes this registry at runtime but omits it from its Worksheet type.
  const validations = (worksheet as ExcelJS.Worksheet & {
    dataValidations: { add(address: string, rule: ExcelJS.DataValidation): void };
  }).dataValidations;
  for (const range of contiguousRanges(socialRows, 'G')) {
    validations.add(range, { type: 'list', allowBlank: false, formulae: [socialFormula] });
  }
  for (const range of contiguousRanges(originalUnitRows, 'C')) {
    validations.add(range, { type: 'list', allowBlank: false, formulae: [unitFormula] });
  }
  for (const range of contiguousRanges(newUnitRows, 'C')) {
    validations.add(range, { type: 'list', allowBlank: false, formulae: [unitFormula.replace('UNIDAD,', 'UNIDAD,LITRO,')] });
  }

  const conditionalRanges = contiguousRanges(conditionalRows, 'C');
  if (conditionalRanges.length) {
    worksheet.addConditionalFormatting({
      ref: conditionalRanges.join(' '),
      rules: [{
        type: 'expression',
        priority: 1,
        formulae: [`C${Math.min(...conditionalRows)}="POR DEFINIR"`],
        style: {
          fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '00000000' },
            bgColor: { argb: 'FFFEF3C7' },
          },
          font: { color: { argb: 'FF92400E' } },
        },
      }],
    });
  }

  return workbook;
}

function valueText(value: ExcelJS.CellValue): string {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number') return String(value);
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') {
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join('');
    }
    if ('text' in value && typeof value.text === 'string') return value.text;
    if ('result' in value) return valueText(value.result as ExcelJS.CellValue);
  }
  return String(value);
}

function noteText(cell: ExcelJS.Cell): string {
  const note = cell.note;
  if (!note) return '';
  if (typeof note === 'string') return note;
  return note.texts?.map((part) => part.text).join('') ?? '';
}

function groupValidationRanges(xml: string): string {
  return xml.replace(/<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>/, (section) => {
    const nodes = [...section.matchAll(/<dataValidation\b[^>]*>[\s\S]*?<\/dataValidation>/g)]
      .map((match) => match[0]);
    if (!nodes.length) return section;
    const groups = new Map<string, { template: string; ranges: string[]; column: string }>();
    for (const node of nodes) {
      const ref = /\bsqref="([^"]+)"/.exec(node)?.[1];
      if (!ref) throw new Error('Falta un rango de validación en el Excel.');
      const template = node.replace(/\bsqref="[^"]+"/, 'sqref=""');
      const column = /^[A-Z]+/.exec(ref)?.[0] ?? '';
      const group = groups.get(template) ?? { template, ranges: [], column };
      group.ranges.push(ref);
      groups.set(template, group);
    }
    const ordered = [...groups.values()].sort((a, b) => b.column.localeCompare(a.column));
    const combined = ordered.map(({ template, ranges }) => {
      ranges.sort((a, b) => {
        const rowA = Number(/\d+/.exec(a)?.[0] ?? 0);
        const rowB = Number(/\d+/.exec(b)?.[0] ?? 0);
        return rowA - rowB;
      });
      return template.replace('sqref=""', `sqref="${ranges.join(' ')}"`);
    });
    return `<dataValidations count="${combined.length}">${combined.join('')}</dataValidations>`;
  });
}

function readPrice(cell: ExcelJS.Cell): number | null {
  const raw = cell.value;
  if (raw == null) return null;
  if (typeof raw === 'number') {
    if (Number.isFinite(raw)) return raw;
    throw new Error(`Precio inválido en ${cell.address}.`);
  }
  const text = valueText(raw).trim();
  if (!text || /^(sin precio|por definir|n\/?a)$/i.test(text)) return null;
  const currency = text.replace(/[$\s]/g, '');
  if (!/^-?[\d.,]+$/.test(currency)) throw new Error(`Precio inválido en ${cell.address}.`);
  const lastComma = currency.lastIndexOf(',');
  const lastDot = currency.lastIndexOf('.');
  const decimalMark = lastComma > lastDot ? ',' : '.';
  const normalized = decimalMark === ','
    ? currency.replace(/\./g, '').replace(',', '.')
    : currency.replace(/,/g, '');
  const price = Number(normalized);
  if (!Number.isFinite(price)) throw new Error(`Precio inválido en ${cell.address}.`);
  return price;
}

function isCategoryFill(cell: ExcelJS.Cell): boolean {
  const cellFill = cell.fill;
  if (!cellFill || cellFill.type !== 'pattern') return false;
  const color = cellFill.fgColor;
  if (color?.argb?.toUpperCase().endsWith(sourceStyle.colors.lightGreen)) return true;
  const tint = (color as { tint?: number } | undefined)?.tint ?? 0;
  return color?.theme === 6 && Math.abs(tint - 0.8) < 0.01;
}

const originalProductsByName = new Map(
  (catalog.products as Product[]).map((product) => [
    `${product.category.trim().toLocaleLowerCase('es')}\u0000${product.product.trim().toLocaleLowerCase('es')}`,
    product,
  ]),
);

/** Read the source layout or a dated export, preserving categories and blank prices. */
export async function importWorkbook(file: File): Promise<{ categories: string[]; products: Product[] }> {
  const workbook = new ExcelJS.Workbook();
  const bytes = new Uint8Array(await file.arrayBuffer());
  await workbook.xlsx.load(bytes as unknown as Parameters<typeof workbook.xlsx.load>[0]);
  const worksheet = workbook.worksheets.find(
    (sheet) =>
      valueText(sheet.getCell('A3').value).trim().toLocaleLowerCase('es') === 'producto' &&
      valueText(sheet.getCell('D3').value).trim().toLocaleLowerCase('es') === 'precio',
  );
  if (!worksheet) throw new Error('El libro no tiene el formato del listado de precios.');

  const categories: string[] = [];
  const products: Product[] = [];
  let category = '';
  for (let rowNumber = 4; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    const fields = columnLetters.map((_, index) => valueText(row.getCell(index + 1).value));
    if (fields.every((field) => field.trim() === '')) continue;

    const first = fields[0];
    const separatorTailIsBlank = [1, 2, 3, 4, 6, 7].every((index) => fields[index].trim() === '');
    const knownCategory = catalog.categories.some((name) => name.trim() === first.trim());
    if (
      first.trim() &&
      separatorTailIsBlank &&
      (!fields[5].trim() || fields[5].trim() === 'Y') &&
      (isCategoryFill(row.getCell(1)) || knownCategory)
    ) {
      category = first;
      if (!categories.includes(category)) categories.push(category);
      continue;
    }

    if (!first.trim()) throw new Error(`Falta el nombre del producto en A${rowNumber}.`);
    if (!category) {
      category = 'Sin categoría';
      categories.push(category);
    }
    const original = originalProductsByName.get(
      `${category.trim().toLocaleLowerCase('es')}\u0000${first.trim().toLocaleLowerCase('es')}`,
    );
    const sourceComments: Record<string, string> = {};
    columnLetters.forEach((letter, index) => {
      const note = noteText(row.getCell(index + 1));
      if (note) sourceComments[letter] = note;
    });
    // ExcelJS does not read the three threaded comments in the supplied file.
    // Restore only those known comments while the product is still on its source row.
    if (original?.originalRow === rowNumber) {
      for (const [letter, comment] of Object.entries(original.sourceComments ?? {})) {
        if (!sourceComments[letter] && comment.startsWith('[Comentario encadenado]')) {
          sourceComments[letter] = comment;
        }
      }
    }
    const product: Product = {
      id: `import-${rowNumber}-${globalThis.crypto?.randomUUID?.() ?? Math.random().toString(36).slice(2)}`,
      category,
      product: fields[0],
      format: fields[1],
      unit: fields[2],
      price: readPrice(row.getCell(4)),
      supplier: fields[4],
      address: fields[5],
      social: fields[6],
      notes: fields[7],
    };
    if (original?.originalRow) product.originalRow = original.originalRow;
    if (Object.keys(sourceComments).length) product.sourceComments = sourceComments;
    products.push(product);
  }
  if (!products.length) throw new Error('El libro no contiene productos.');
  return { categories, products };
}

/** Download an .xlsx named for the selected local calendar date. */
export async function exportWorkbook(products: Product[], date: string): Promise<void> {
  const { day, month, year } = parseDate(date);
  const workbook = createWorkbook(products, date);
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(new Uint8Array(buffer));
  const sheet = zip.file('xl/worksheets/sheet1.xml');
  if (!sheet) throw new Error('No se encontró la hoja del listado en el Excel.');
  const xml = await sheet.async('string');
  // ExcelJS 4.4 serializes the legacy comment drawing after tableParts.
  // Microsoft Excel requires legacyDrawing before tableParts and rejects the file otherwise.
  const reordered = xml.replace(
    /(<tableParts\b[^>]*>[\s\S]*?<\/tableParts>)(\s*)(<legacyDrawing\b[^>]*\/>)/,
    (_match, table: string, whitespace: string, drawing: string) => `${drawing}${whitespace}${table}`,
  );
  if (xml.indexOf('<legacyDrawing') > xml.indexOf('<tableParts') && reordered === xml) {
    throw new Error('No se pudo corregir el orden de elementos de la hoja Excel.');
  }
  const corrected = groupValidationRanges(reordered);
  if (corrected !== xml) zip.file('xl/worksheets/sheet1.xml', corrected);

  // ExcelJS supplies Calibri as the default font. The source workbook uses
  // Carlito, which also applies to blank cells inside the merged title area.
  const stylesFile = zip.file('xl/styles.xml');
  if (!stylesFile) throw new Error('No se encontraron los estilos del libro Excel.');
  const stylesXml = await stylesFile.async('string');
  const sourceDefaultFont = '<font><sz val="11"/><name val="Carlito"/></font>';
  const correctedStyles = stylesXml.replace(
    /(<fonts\b[^>]*>)<font>[\s\S]*?<\/font>/,
    `$1${sourceDefaultFont}`,
  );
  if (correctedStyles === stylesXml) {
    throw new Error('No se pudo conservar la fuente predeterminada del libro original.');
  }
  zip.file('xl/styles.xml', correctedStyles);

  const fileBytes = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  const blob = new Blob([fileBytes], { type: workbookMime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `Listado de precios ${day}-${month}-${year.slice(-2)}.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
