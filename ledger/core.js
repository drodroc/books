(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.LedgerCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  function normHeader(value) {
    return String(value == null ? "" : value)
      .replace(/\s+/g, "")
      .replace(/[\r\n]/g, "")
      .trim();
  }

  function parseMoneyCents(value) {
    if (value == null || value === "" || value === "无退款申请") return 0;
    if (typeof value === "number") {
      if (!isFinite(value)) return 0;
      return Math.round(value * 100);
    }
    var text = String(value).replace(/,/g, "").replace(/元/g, "").trim();
    if (!text || text === "无退款申请" || text === "-") return 0;
    var number = Number(text);
    if (!isFinite(number)) return 0;
    return Math.round(number * 100);
  }

  function parseQty(value) {
    if (value == null || value === "") return 0;
    var number = Number(String(value).replace(/,/g, "").trim());
    return isFinite(number) ? number : 0;
  }

  function pad(number) {
    return String(number).padStart(2, "0");
  }

  function wallFromDateUTC(date) {
    return {
      y: date.getUTCFullYear(),
      m: date.getUTCMonth() + 1,
      d: date.getUTCDate(),
      hh: date.getUTCHours(),
      mm: date.getUTCMinutes(),
      ss: date.getUTCSeconds()
    };
  }

  function wallFromDateLocal(date) {
    return {
      y: date.getFullYear(),
      m: date.getMonth() + 1,
      d: date.getDate(),
      hh: date.getHours(),
      mm: date.getMinutes(),
      ss: date.getSeconds()
    };
  }

  // Excel serials and date-only strings stay on the clock written in the cell.
  // JavaScript Date objects from SheetJS use the computer's local clock, which
  // is what Windows shows when the workbook is opened in Excel.
  function parseWall(value) {
    if (value == null || value === "") return null;
    if (value instanceof Date) {
      if (isNaN(value.getTime())) return null;
      return wallFromDateLocal(value);
    }
    if (typeof value === "number" && value > 20000 && value < 90000) {
      var ms = Date.UTC(1899, 11, 30) + Math.round(value * 86400000);
      return wallFromDateUTC(new Date(ms));
    }
    var text = String(value).trim();
    if (!text) return null;
    var match = text.match(
      /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/
    );
    if (!match) return null;
    return {
      y: +match[1],
      m: +match[2],
      d: +match[3],
      hh: +(match[4] || 0),
      mm: +(match[5] || 0),
      ss: +(match[6] || 0)
    };
  }

  function monthKey(wall) {
    if (!wall) return "";
    return wall.y + "-" + pad(wall.m);
  }

  function stamp(wall) {
    if (!wall) return "";
    return (
      wall.y +
      "-" +
      pad(wall.m) +
      "-" +
      pad(wall.d) +
      " " +
      pad(wall.hh) +
      ":" +
      pad(wall.mm) +
      ":" +
      pad(wall.ss)
    );
  }

  function findHeaderRow(matrix, required) {
    var limit = Math.min(matrix.length, 30);
    for (var r = 0; r < limit; r++) {
      var cells = (matrix[r] || []).map(normHeader);
      var ok = required.every(function (name) {
        return cells.indexOf(name) !== -1;
      });
      if (ok) return r;
    }
    return -1;
  }

  function indexHeaders(row) {
    var map = {};
    (row || []).forEach(function (cell, index) {
      var key = normHeader(cell);
      if (key && map[key] == null) map[key] = index;
    });
    return map;
  }

  function pick(map, names) {
    for (var i = 0; i < names.length; i++) {
      if (map[names[i]] != null) return map[names[i]];
    }
    return -1;
  }

  function cell(row, index) {
    if (index < 0) return "";
    return row[index];
  }

  function lineId(line) {
    return [
      line.orderId,
      line.title,
      line.attr,
      line.qty,
      line.paidCents,
      line.payableCents,
      line.orderStamp,
      line.confirmStamp,
      line.status,
      line.refundCents
    ].join("|");
  }

  function parseRevenueMatrix(matrix) {
    var headerAt = findHeaderRow(matrix, ["商品标题", "购买数量"]);
    if (headerAt < 0) {
      return { ok: false, error: "没有找到「商品标题」和「购买数量」这两列表头。" };
    }
    var map = indexHeaders(matrix[headerAt]);
    var paidCol = pick(map, ["买家实付金额", "买家实付"]);
    if (paidCol < 0) {
      return { ok: false, error: "没有找到「买家实付金额」这一列。" };
    }
    var titleCol = map["商品标题"];
    var attrCol = pick(map, ["商品属性"]);
    var qtyCol = map["购买数量"];
    var priceCol = pick(map, ["商品价格"]);
    var payableCol = pick(map, ["买家应付货款", "买家应付"]);
    var statusCol = pick(map, ["退款状态"]);
    var refundCol = pick(map, ["退款金额"]);
    var payCol = pick(map, ["订单付款时间"]);
    var createCol = pick(map, ["订单创建时间", "下单时间"]);
    var confirmCol = pick(map, ["确认收货时间"]);
    var orderCol = pick(map, ["主订单编号", "订单编号"]);
    var timeField = payCol >= 0 ? "订单付款时间" : createCol >= 0 ? "订单创建时间" : "";
    var timeCol = payCol >= 0 ? payCol : createCol;

    var lines = [];
    for (var r = headerAt + 1; r < matrix.length; r++) {
      var row = matrix[r] || [];
      var title = String(cell(row, titleCol) || "").trim();
      if (!title || title === "合计") continue;
      var attr = String(cell(row, attrCol) || "").trim();
      var orderWall = parseWall(cell(row, timeCol));
      var confirmWall = parseWall(cell(row, confirmCol));
      var line = {
        orderId: String(cell(row, orderCol) || "").trim(),
        title: title,
        attr: attr,
        qty: parseQty(cell(row, qtyCol)),
        priceCents: parseMoneyCents(cell(row, priceCol)),
        payableCents: parseMoneyCents(cell(row, payableCol)),
        paidCents: parseMoneyCents(cell(row, paidCol)),
        status: String(cell(row, statusCol) || "").trim(),
        refundCents: parseMoneyCents(cell(row, refundCol)),
        orderWall: orderWall,
        confirmWall: confirmWall,
        orderStamp: stamp(orderWall),
        confirmStamp: stamp(confirmWall),
        timeField: timeField
      };
      line.id = lineId(line);
      lines.push(line);
    }
    return {
      ok: true,
      lines: lines,
      timeField: timeField,
      columns: Object.keys(map)
    };
  }

  var EXPENSE_NAME_HEADERS = [
    "支出项目",
    "费用项目",
    "费用名称",
    "项目名称",
    "项目",
    "物料名称",
    "品名",
    "商品名称",
    "业务描述",
    "摘要",
    "用途",
    "内容",
    "名称"
  ];
  var EXPENSE_AMOUNT_HEADERS = [
    "支出金额（元）",
    "支出金额(元)",
    "支出金额",
    "金额（元）",
    "金额(元)",
    "金额",
    "成本",
    "费用",
    "付款金额",
    "实付金额"
  ];
  var EXPENSE_DATE_HEADERS = ["支出日期", "日期", "入账时间", "付款时间", "发生时间", "时间"];
  var EXPENSE_NOTE_HEADERS = ["备注", "说明", "附言"];
  var EXPENSE_QTY_HEADERS = ["数量", "购买数量"];

  function guessExpenseMapping(headers) {
    var map = indexHeaders(headers);
    return {
      name: pick(map, EXPENSE_NAME_HEADERS.map(normHeader)),
      amount: pick(map, EXPENSE_AMOUNT_HEADERS.map(normHeader)),
      date: pick(map, EXPENSE_DATE_HEADERS.map(normHeader)),
      note: pick(map, EXPENSE_NOTE_HEADERS.map(normHeader)),
      qty: pick(map, EXPENSE_QTY_HEADERS.map(normHeader)),
      income: pick(map, ["收入金额（元）", "收入金额(元)", "收入金额"].map(normHeader))
    };
  }

  function parseExpenseMatrix(matrix, mapping) {
    var headerAt = -1;
    var guessed = null;
    var limit = Math.min(matrix.length, 30);
    for (var r = 0; r < limit; r++) {
      var trial = guessExpenseMapping(matrix[r] || []);
      if (trial.amount >= 0 && (trial.name >= 0 || trial.note >= 0)) {
        headerAt = r;
        guessed = trial;
        break;
      }
    }
    if (headerAt < 0) {
      var previewHeaders = (matrix[0] || []).map(function (cellValue) {
        return String(cellValue == null ? "" : cellValue);
      });
      return {
        ok: false,
        needsMapping: previewHeaders.some(function (item) { return item.trim(); }),
        headers: previewHeaders,
        headerAt: 0,
        error: "没有认出支出金额列。请指定哪一列是项目、哪一列是金额。"
      };
    }
    var used = mapping || guessed;
    if (used.amount == null || used.amount < 0) {
      return { ok: false, needsMapping: true, headers: (matrix[headerAt] || []).map(String), headerAt: headerAt, error: "还需要指定金额列。" };
    }
    var expenses = [];
    var skippedIncome = 0;
    for (var i = headerAt + 1; i < matrix.length; i++) {
      var row = matrix[i] || [];
      var amountCents = parseMoneyCents(cell(row, used.amount));
      var incomeCents = used.income >= 0 ? parseMoneyCents(cell(row, used.income)) : 0;
      if (amountCents === 0 && incomeCents > 0) {
        skippedIncome++;
        continue;
      }
      if (amountCents === 0) continue;
      var name = String(cell(row, used.name) || "").trim();
      var note = String(cell(row, used.note) || "").trim();
      if (!name) name = note || "未命名支出";
      var dateWall = parseWall(cell(row, used.date));
      expenses.push({
        name: name,
        note: name === note ? "" : note,
        amountCents: Math.abs(amountCents),
        qty: parseQty(cell(row, used.qty)),
        dateWall: dateWall,
        dateStamp: stamp(dateWall),
        month: monthKey(dateWall)
      });
    }
    return {
      ok: true,
      expenses: expenses,
      skippedIncome: skippedIncome,
      headerAt: headerAt,
      mapping: used
    };
  }

  function emptyBucket() {
    return { qty: 0, cents: 0, orders: 0 };
  }

  function addBucket(bucket, line, cents) {
    bucket.qty += line.qty;
    bucket.cents += cents;
    bucket.orders += 1;
  }

  function amountOf(line, basis) {
    return basis === "payable" ? line.payableCents : line.paidCents;
  }

  function isIgnoredUnpaid(line) {
    return !line.orderWall && !line.confirmWall && line.paidCents === 0 && line.payableCents === 0;
  }

  function aggregate(lines, month, basis) {
    var products = {};
    var totals = {
      confirmed: emptyBucket(),
      pending: emptyBucket(),
      later: emptyBucket(),
      refund: emptyBucket()
    };
    var ignoredUnpaid = 0;
    var orderIds = { confirmed: {}, pending: {}, later: {}, refund: {} };

    function touch(bucketName, orderId) {
      if (!orderId) return;
      if (!orderIds[bucketName][orderId]) {
        orderIds[bucketName][orderId] = true;
      }
    }

    lines.forEach(function (line) {
      if (isIgnoredUnpaid(line)) {
        ignoredUnpaid++;
        return;
      }
      var confirmMonth = monthKey(line.confirmWall);
      var orderMonth = monthKey(line.orderWall);
      var inConfirm = confirmMonth === month;
      var inOrder = orderMonth === month;
      if (!inConfirm && !inOrder) return;

      var product = products[line.title] || {
        title: line.title,
        attrs: {},
        confirmed: emptyBucket(),
        pending: emptyBucket(),
        later: emptyBucket(),
        refund: emptyBucket()
      };
      var attrName = line.attr || "";
      var attr = product.attrs[attrName] || {
        attr: attrName,
        confirmed: emptyBucket(),
        pending: emptyBucket(),
        later: emptyBucket(),
        refund: emptyBucket()
      };
      var cents = amountOf(line, basis);
      var bucketName = "";
      if (line.status === "退款成功") {
        if (inConfirm || inOrder) bucketName = "refund";
      } else if (inConfirm) {
        bucketName = "confirmed";
      } else if (inOrder && !line.confirmWall) {
        bucketName = "pending";
      } else if (inOrder && line.confirmWall) {
        bucketName = "later";
      }
      if (!bucketName) {
        products[line.title] = product;
        return;
      }
      addBucket(product[bucketName], line, cents);
      addBucket(attr[bucketName], line, cents);
      addBucket(totals[bucketName], line, cents);
      touch(bucketName, line.orderId);
      product.attrs[attrName] = attr;
      products[line.title] = product;
    });

    var list = Object.keys(products).map(function (title) {
      var product = products[title];
      product.attrs = Object.keys(product.attrs)
        .map(function (key) { return product.attrs[key]; })
        .sort(function (a, b) { return b.confirmed.cents - a.confirmed.cents || b.pending.cents - a.pending.cents; });
      return product;
    });
    list.sort(function (a, b) {
      return b.confirmed.cents - a.confirmed.cents || b.pending.cents - a.pending.cents || a.title.localeCompare(b.title, "zh");
    });

    function uniqueOrders(name) {
      return Object.keys(orderIds[name]).length;
    }

    return {
      products: list,
      totals: totals,
      orderCounts: {
        confirmed: uniqueOrders("confirmed"),
        pending: uniqueOrders("pending"),
        later: uniqueOrders("later"),
        refund: uniqueOrders("refund")
      },
      ignoredUnpaid: ignoredUnpaid
    };
  }

  function titleKey(title) {
    return "T:" + title;
  }

  function attrKey(title, attr) {
    return "A:" + title + "\n" + (attr || "");
  }

  return {
    parseRevenueMatrix: parseRevenueMatrix,
    parseExpenseMatrix: parseExpenseMatrix,
    aggregate: aggregate,
    parseWall: parseWall,
    monthKey: monthKey,
    stamp: stamp,
    titleKey: titleKey,
    attrKey: attrKey,
    parseMoneyCents: parseMoneyCents
  };
});
