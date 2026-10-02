import ExcelJS from 'exceljs';

const SHEETS = {
  members: 'Members',
  expenses: 'Expenses',
  receipts: 'receipts',
  documents: 'Member Documents',
};

function value(row, names) {
  const normalizedNames = new Set(names.map((name) => name.toLowerCase().replace(/[^a-z0-9]/g, '')));
  const entry = Object.entries(row).find(([name, result]) => normalizedNames.has(name.toLowerCase().replace(/[^a-z0-9]/g, ''))
    && result !== undefined && result !== null && result !== '');
  return entry ? entry[1] : '';
}

function getWorksheet(workbook, name) {
  return workbook.worksheets.find((worksheet) => worksheet.name.toLowerCase() === name.toLowerCase());
}

function numericValue(input) {
  if (typeof input === 'number') return input;
  const numeric = Number(String(input || '').replace(/^(?:INR|Rs\.?)\s*/i, '').replace(/[₹,\s]/g, ''));
  return Number.isFinite(numeric) ? numeric : 0;
}

function cellDate(input) {
  if (!input) return '';
  if (input instanceof Date && !Number.isNaN(input.getTime())) return input.toISOString().slice(0, 10);
  if (typeof input === 'number') {
    const date = new Date(Date.UTC(1899, 11, 30) + input * 86400000);
    return date.toISOString().slice(0, 10);
  }
  const text = String(input).trim();
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? text : parsed.toISOString().slice(0, 10);
}

function rowsAsObjects(worksheet) {
  if (!worksheet || worksheet.rowCount < 2) return [];
  const headers = worksheet.getRow(1).values.slice(1).map((header) => String(header || '').trim());
  const rows = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const result = {};
    headers.forEach((header, index) => {
      if (header) result[header] = row.getCell(index + 1).value;
    });
    if (Object.values(result).some((item) => item !== null && item !== undefined && item !== '')) rows.push(result);
  });
  return rows;
}

function asId(input, prefix, index) {
  return String(input || `${prefix}-${index + 1}`).trim();
}

function normalizeList(input) {
  if (Array.isArray(input)) return input.map(String);
  return String(input || '').split(',').map((part) => part.trim()).filter(Boolean);
}

function parseRecords(workbook, sheetName, transformer) {
  return rowsAsObjects(getWorksheet(workbook, sheetName)).map(transformer);
}

export async function parseWorkbook(input) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(input);
  if (!getWorksheet(workbook, SHEETS.members)) {
    throw new Error('The workbook is missing the required Members worksheet.');
  }

  const members = parseRecords(workbook, SHEETS.members, (row, index) => ({
    id: asId(value(row, ['ID', 'Id', 'Member ID']), 'player', index),
    name: String(value(row, ['Name', 'Player Name', 'Member Name']) || '').trim(),
    email: String(value(row, ['Email', 'Email Address']) || '').trim(),
    position: String(value(row, ['Position', 'Role']) || 'Player').trim(),
  })).filter((member) => member.name);

  const memberIdFor = (person) => {
    const match = members.find((member) => member.id === person
      || member.name.toLowerCase() === String(person).trim().toLowerCase()
      || member.email.toLowerCase() === String(person).trim().toLowerCase());
    return match ? match.id : String(person || '').trim();
  };

  const expenses = parseRecords(workbook, SHEETS.expenses, (row, index) => ({
    id: asId(value(row, ['ID', 'Id', 'Expense ID']), 'expense', index),
    title: String(value(row, ['Description', 'Title', 'Expense']) || '').trim(),
    amount: numericValue(value(row, ['Amount', 'Cost', 'Total'])),
    paidBy: memberIdFor(value(row, ['Paid By ID', 'Paid By', 'Person', 'Name'])),
    date: cellDate(value(row, ['Date', 'Expense Date'])),
    participantIds: normalizeList(value(row, ['Participant IDs', 'Participants'])).map(memberIdFor),
  })).filter((expense) => expense.title);

  const receipts = parseRecords(workbook, SHEETS.receipts, (row, index) => ({
    id: asId(value(row, ['ID', 'Receipt ID']), 'receipt', index),
    expenseId: String(value(row, ['Expense ID']) || ''),
    date: cellDate(value(row, ['Date', 'Receipt Date'])),
    amount: numericValue(value(row, ['Amount', 'Total', 'Cost'])),
    personId: memberIdFor(value(row, ['Person ID', 'Person', 'Paid By ID', 'Paid By'])),
    description: String(value(row, ['Description', 'Merchant', 'Title']) || ''),
    fileName: String(value(row, ['File Name', 'Document']) || ''),
    ocrText: String(value(row, ['OCR Text']) || ''),
  })).filter((receipt) => receipt.description || receipt.amount);

  const documents = rowsAsObjects(workbook.getWorksheet(SHEETS.documents)).map((row, index) => ({
    id: asId(value(row, ['ID', 'Document ID']), 'document', index),
    memberId: String(value(row, ['Member ID']) || ''),
    fileName: String(value(row, ['File Name', 'Name']) || ''),
    contentType: String(value(row, ['Content Type', 'Type']) || ''),
    kind: String(value(row, ['Kind']) || 'member'),
  })).filter((document) => document.fileName);

  return { members, expenses, receipts, documents };
}

function addSheet(workbook, name, headers, records) {
  const worksheet = workbook.addWorksheet(name);
  worksheet.addRow(headers);
  records.forEach((record) => worksheet.addRow(record));
  worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF363A35' } };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  worksheet.columns.forEach((column) => {
    column.width = Math.min(Math.max(column.header ? String(column.header).length + 4 : 16, 14), 42);
  });
  return worksheet;
}

export async function createWorkbook({ members, expenses, receipts, documents }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Thunder team admin';
  workbook.created = new Date();

  addSheet(workbook, SHEETS.members, ['ID', 'Name', 'Email', 'Position'], members.map((member) => [
    member.id, member.name, member.email, member.position,
  ]));

  addSheet(workbook, SHEETS.expenses, ['ID', 'Date', 'Description', 'Amount', 'Paid By ID', 'Participant IDs'], expenses.map((expense) => [
    expense.id, expense.date, expense.title, Number(expense.amount), expense.paidBy, (expense.participantIds || []).join(', '),
  ]));

  addSheet(workbook, SHEETS.receipts, ['ID', 'Expense ID', 'Date', 'Amount', 'Person ID', 'Description', 'File Name', 'OCR Text'], receipts.map((receipt) => [
    receipt.id, receipt.expenseId, receipt.date, Number(receipt.amount), receipt.personId, receipt.description, receipt.fileName, receipt.ocrText,
  ]));

  addSheet(workbook, SHEETS.documents, ['ID', 'Member ID', 'File Name', 'Content Type', 'Kind'], documents.map((document) => [
    document.id, document.memberId, document.fileName, document.contentType, document.kind || 'member',
  ]));

  return workbook.xlsx.writeBuffer();
}

export const PUBLIC_WORKBOOK_PATH = `${(process.env.PUBLIC_URL || '').replace(/\/$/, '')}/thunder.xlsx`;
