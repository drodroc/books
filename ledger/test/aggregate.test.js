const fs = require("fs");
const path = require("path");
const XLSX = require("../lib/xlsx.full.min.js");
const core = require("../core.js");

function matrixFrom(file) {
  const workbook = XLSX.read(fs.readFileSync(file), { type: "buffer", cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
}

function assert(cond, message) {
  if (!cond) {
    console.error("FAIL", message);
    process.exit(1);
  }
}

const yidianPath = "/home/ubuntu/.cursor/projects/workspace/uploads/____8_9_10____59ac.xlsx";
const financePath = "/home/ubuntu/.cursor/projects/workspace/uploads/____8_9_10__4af7.xlsx";
if (fs.existsSync(yidianPath) && fs.existsSync(financePath)) {
const yidian = matrixFrom(yidianPath);
const parsed = core.parseRevenueMatrix(yidian);
assert(parsed.ok, parsed.error);
assert(parsed.timeField === "订单创建时间", "time field " + parsed.timeField);
const report = core.aggregate(parsed.lines, "2026-10", "paid");
assert(report.totals.confirmed.cents === 12987347, "confirmed cents " + report.totals.confirmed.cents);
assert(report.totals.confirmed.qty === 2425, "confirmed qty " + report.totals.confirmed.qty);
assert(report.orderCounts.confirmed === 630, "confirmed orders " + report.orderCounts.confirmed);
assert(report.totals.refund.cents > 217890, "refund cents " + report.totals.refund.cents);
let touched = 0;
parsed.lines.forEach(function (line) {
  var confirmMonth = core.monthKey(line.confirmWall);
  var orderMonth = core.monthKey(line.orderWall);
  if (confirmMonth !== "2026-10" && orderMonth !== "2026-10") return;
  if (!line.orderWall && !line.confirmWall && line.paidCents === 0) return;
  touched += line.paidCents;
});
var bucketSum = report.totals.confirmed.cents + report.totals.pending.cents + report.totals.later.cents + report.totals.refund.cents;
assert(bucketSum === touched, "buckets " + bucketSum + " touched " + touched);

const finance = matrixFrom(financePath);
const financeParsed = core.parseRevenueMatrix(finance);
assert(financeParsed.ok, financeParsed.error);
assert(financeParsed.timeField === "订单付款时间", financeParsed.timeField);
const financeReport = core.aggregate(financeParsed.lines, "2026-10", "paid");
assert(financeReport.totals.confirmed.cents > 0, "finance confirm empty");
assert(financeReport.totals.pending.cents > 0, "finance pending empty");

const custom = financeParsed.lines.filter(function (line) {
  return line.attr.indexOf("尺寸定制不要随便拍") !== -1;
});
const customReport = core.aggregate(custom, "2026-10", "paid");
assert(customReport.totals.pending.qty === 46, "custom pending qty " + customReport.totals.pending.qty);
assert(customReport.totals.pending.cents === 46000, "custom pending cents " + customReport.totals.pending.cents);
assert(customReport.totals.confirmed.qty === 44, "custom confirm qty " + customReport.totals.confirmed.qty);
assert(customReport.totals.confirmed.cents === 44000, "custom confirm cents " + customReport.totals.confirmed.cents);

const expense = core.parseExpenseMatrix([
  ["日期", "支出项目", "金额", "备注"],
  ["2026-10-03", "PCB打样", "1,280.50", "新能源"],
  ["2026-09-01", "旧采购", 20, ""],
  ["", "", "", ""]
]);
assert(expense.ok && expense.expenses.length === 2, "expense count");
assert(expense.expenses[0].amountCents === 128050, "expense cents " + expense.expenses[0].amountCents);
assert(expense.expenses[0].month === "2026-10", expense.expenses[0].month);
assert(expense.incomes.length === 0, "plain expense sheet has no incomes");

const same = core.parseRevenueMatrix(yidian);
const seen = {};
let dup = 0;
parsed.lines.concat(same.lines).forEach(function (line) {
  if (seen[line.id]) dup++;
  seen[line.id] = true;
});
assert(dup === parsed.lines.length, "duplicate ids " + dup);
}

function matrixOf(file) {
  const workbook = XLSX.read(fs.readFileSync(file), { type: "buffer", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
}
function cents(items) {
  return items.reduce(function (sum, item) { return sum + item.amountCents; }, 0);
}
const books = [
  ["/home/ubuntu/.cursor/projects/workspace/uploads/10_____228b.xlsx", 21042475, 1574410],
  ["/home/ubuntu/.cursor/projects/workspace/uploads/10________38c4.xlsx", 7637935, 7240],
  ["/home/ubuntu/.cursor/projects/workspace/uploads/____-__1_6a7d.xlsx", 34333441, 5167800]
];
books.forEach(function (item) {
  const parsedBook = core.parseExpenseMatrix(matrixOf(item[0]));
  assert(parsedBook.ok && parsedBook.statement, item[0]);
  assert(cents(parsedBook.expenses) === item[1], item[0] + " expense " + cents(parsedBook.expenses));
  assert(cents(parsedBook.incomes) === item[2], item[0] + " income " + cents(parsedBook.incomes));
  assert(parsedBook.expenses.every(function (row) { return row.month === "2026-10"; }), item[0] + " month");
});

console.log("ok", books.map(function (item) {
  const parsedBook = core.parseExpenseMatrix(matrixOf(item[0]));
  return { file: item[0].split("/").pop(), expense: cents(parsedBook.expenses) / 100, income: cents(parsedBook.incomes) / 100 };
}));
