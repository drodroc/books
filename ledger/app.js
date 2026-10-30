(function () {
  var core = window.LedgerCore;
  var app = document.getElementById("app");
  var state = {
    lines: [],
    expenses: [],
    files: [],
    month: "",
    query: "",
    expenseQuery: "",
    filter: "all",
    sort: "confirmed",
    expanded: {},
    selectedExpenseId: "",
    selectedProductKey: "",
    detail: null,
    mapper: null,
    help: false,
    toast: ""
  };
  var saveTimer = 0;

  function uid() {
    return Math.random().toString(36).slice(2, 8) + Date.now().toString(36);
  }

  function esc(value) {
    return String(value == null ? "" : value).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }

  function fmtMoney(cents) {
    var sign = cents < 0 ? "-" : "";
    var abs = Math.abs(cents) / 100;
    return sign + abs.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fmtDate(stamp) {
  if (!stamp) return "未填日期";
  return String(stamp).replace(/ 00:00:00$/, "");
}

function fmtQty(qty) {
    if (!isFinite(qty)) return "0";
    if (Math.abs(qty - Math.round(qty)) < 1e-6) return Math.round(qty).toLocaleString("zh-CN");
    return qty.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
  }

  function monthLabel(month) {
    if (!month) return "未选择月份";
    var parts = month.split("-");
    return parts[0] + "年" + Number(parts[1]) + "月";
  }

  function openDb() {
    return new Promise(function (resolve, reject) {
      var request = indexedDB.open("yidian-ledger", 1);
      request.onupgradeneeded = function () {
        request.result.createObjectStore("kv");
      };
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    });
  }

  function loadState() {
    return openDb().then(function (db) {
      return new Promise(function (resolve) {
        var tx = db.transaction("kv", "readonly");
        var request = tx.objectStore("kv").get("state");
        request.onsuccess = function () { resolve(request.result || null); };
        request.onerror = function () { resolve(null); };
      });
    }).catch(function () { return null; });
  }

  function saveState() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      openDb().then(function (db) {
        var tx = db.transaction("kv", "readwrite");
        tx.objectStore("kv").put({
          lines: state.lines,
          expenses: state.expenses,
          files: state.files,
          month: state.month,
          expanded: state.expanded
        }, "state");
      }).catch(function () {});
    }, 200);
  }

  function toast(message) {
    state.toast = message;
    render();
    setTimeout(function () {
      if (state.toast === message) {
        state.toast = "";
        var node = document.querySelector(".toast");
        if (node) node.remove();
      }
    }, 2800);
  }

  function monthsInData() {
    var set = {};
    state.lines.forEach(function (line) {
      var confirmMonth = core.monthKey(line.confirmWall);
      var orderMonth = core.monthKey(line.orderWall);
      if (confirmMonth) set[confirmMonth] = true;
      if (orderMonth) set[orderMonth] = true;
    });
    state.expenses.forEach(function (expense) {
      if (expense.month) set[expense.month] = true;
    });
    return Object.keys(set).sort();
  }

  function ensureMonth() {
    var months = monthsInData();
    if (!state.month || months.indexOf(state.month) === -1) {
      state.month = months.length ? months[months.length - 1] : "";
    }
  }

  function sheetMatrix(workbook) {
    var matrices = workbook.SheetNames.map(function (name) {
      return XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: "" });
    });
    return matrices;
  }

  function readWorkbook(file) {
    return file.arrayBuffer().then(function (buffer) {
      return XLSX.read(buffer, { type: "array", cellDates: true });
    });
  }

  function mergeLines(parsed, fileName) {
    var seen = {};
    state.lines.forEach(function (line) { seen[line.id] = true; });
    var added = 0;
    var skipped = 0;
    parsed.lines.forEach(function (line) {
      if (seen[line.id]) {
        skipped++;
        return;
      }
      seen[line.id] = true;
      line.sourceFile = fileName;
      state.lines.push(line);
      added++;
    });
    state.files.push({ name: fileName, kind: "revenue", added: added, skipped: skipped, at: Date.now() });
    return { added: added, skipped: skipped, timeField: parsed.timeField };
  }

  function importRevenueFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    var notes = [];
    var chain = Promise.resolve();
    files.forEach(function (file) {
      chain = chain.then(function () {
        return readWorkbook(file).then(function (workbook) {
          var matrices = sheetMatrix(workbook);
          var parsed = null;
          var error = "这不是收益订单表。";
          for (var i = 0; i < matrices.length; i++) {
            var trial = core.parseRevenueMatrix(matrices[i]);
            if (trial.ok) {
              parsed = trial;
              break;
            }
            error = trial.error;
          }
          if (!parsed) throw new Error(file.name + "：" + error);
          var result = mergeLines(parsed, file.name);
          notes.push(file.name + " 新增 " + result.added + " 行" + (result.skipped ? "，重复 " + result.skipped + " 行已跳过" : "") + "。下单时间用的是" + (result.timeField || "空") + "。");
        });
      });
    });
    chain.then(function () {
      ensureMonth();
      saveState();
      toast(notes.join(" "));
    }).catch(function (error) {
      toast(error.message || "收益表没有导入成功");
    });
  }

  function addExpenses(parsed, fileName) {
    parsed.expenses.forEach(function (expense) {
      state.expenses.push({
        id: uid(),
        name: expense.name,
        note: expense.note,
        amountCents: expense.amountCents,
        qty: expense.qty,
        dateStamp: expense.dateStamp,
        month: expense.month,
        sourceFile: fileName,
        allocs: []
      });
    });
    state.files.push({
      name: fileName,
      kind: "expense",
      added: parsed.expenses.length,
      at: Date.now()
    });
  }

  function importExpenseFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    var chain = Promise.resolve();
    files.forEach(function (file) {
      chain = chain.then(function () {
        return readWorkbook(file).then(function (workbook) {
          var matrices = sheetMatrix(workbook);
          var parsed = null;
          var fallback = null;
          for (var i = 0; i < matrices.length; i++) {
            var trial = core.parseExpenseMatrix(matrices[i]);
            if (trial.ok && trial.expenses.length) {
              parsed = trial;
              break;
            }
            if (!fallback) fallback = trial;
          }
          if (!parsed) {
            state.mapper = {
              fileName: file.name,
              matrix: matrices[0],
              headers: (fallback && fallback.headers) || (matrices[0][0] || []).map(String),
              headerAt: fallback && fallback.headerAt ? fallback.headerAt : 0,
              error: (fallback && fallback.error) || "没有认出支出表"
            };
            render();
            throw new Error("MAP");
          }
          addExpenses(parsed, file.name);
          var extra = parsed.skippedIncome ? " 收入行 " + parsed.skippedIncome + " 条没有当作支出。" : "";
          toast(file.name + " 导入支出 " + parsed.expenses.length + " 笔。" + extra);
        });
      });
    });
    chain.then(function () {
      ensureMonth();
      saveState();
      render();
    }).catch(function (error) {
      if (error.message !== "MAP") toast(error.message || "支出表没有导入成功");
    });
  }

  function confirmMapping() {
    var mapper = state.mapper;
    if (!mapper) return;
    var name = Number(document.getElementById("map-name").value);
    var amount = Number(document.getElementById("map-amount").value);
    var date = Number(document.getElementById("map-date").value);
    var note = Number(document.getElementById("map-note").value);
    var parsed = core.parseExpenseMatrix(mapper.matrix, {
      name: name,
      amount: amount,
      date: date,
      note: note,
      qty: -1,
      income: -1
    });
    if (!parsed.ok || !parsed.expenses.length) {
      toast(parsed.error || "按这个列对应关系，没有读到支出");
      return;
    }
    addExpenses(parsed, mapper.fileName);
    state.mapper = null;
    ensureMonth();
    saveState();
    toast("导入支出 " + parsed.expenses.length + " 笔");
  }

  function zeroBucket() {
    return { qty: 0, cents: 0, orders: 0 };
  }

  function parseTarget(key) {
    if (!key) return null;
    if (key.indexOf("T:") === 0) return { title: key.slice(2), attr: null };
    if (key.indexOf("A:") === 0) {
      var rest = key.slice(2);
      var splitAt = rest.indexOf("\n");
      if (splitAt < 0) return { title: rest, attr: "" };
      return { title: rest.slice(0, splitAt), attr: rest.slice(splitAt + 1) };
    }
    return null;
  }

  function viewModel() {
    ensureMonth();
    var report = core.aggregate(state.lines, state.month, "paid");
    var expenses = state.expenses.filter(function (expense) {
      return !expense.month || expense.month === state.month;
    });
    var costByKey = {};
    expenses.forEach(function (expense) {
      expense.allocs.forEach(function (alloc) {
        costByKey[alloc.targetKey] = (costByKey[alloc.targetKey] || 0) + alloc.cents;
        var target = parseTarget(alloc.targetKey);
        if (!target || !target.title) return;
        var product = null;
        report.products.forEach(function (item) {
          if (item.title === target.title) product = item;
        });
        if (!product) {
          product = {
            title: target.title,
            attrs: [],
            confirmed: zeroBucket(),
            pending: zeroBucket(),
            later: zeroBucket(),
            refund: zeroBucket()
          };
          report.products.push(product);
        }
        if (target.attr != null) {
          var found = false;
          product.attrs.forEach(function (attr) {
            if ((attr.attr || "") === target.attr) found = true;
          });
          if (!found) {
            product.attrs.push({
              attr: target.attr,
              confirmed: zeroBucket(),
              pending: zeroBucket(),
              later: zeroBucket(),
              refund: zeroBucket()
            });
          }
        }
      });
    });
    function directCost(key) {
      return costByKey[key] || 0;
    }
    function titleCost(product) {
      var cents = directCost(core.titleKey(product.title));
      product.attrs.forEach(function (attr) {
        cents += directCost(core.attrKey(product.title, attr.attr));
      });
      return cents;
    }
    report.products.forEach(function (product) {
      product.costCents = titleCost(product);
      product.profitCents = product.confirmed.cents - product.costCents;
      product.attrs.forEach(function (attr) {
        attr.costCents = directCost(core.attrKey(product.title, attr.attr));
        attr.profitCents = attr.confirmed.cents - attr.costCents;
      });
    });
    var allocated = expenses.reduce(function (sum, expense) {
      return sum + expense.allocs.reduce(function (inner, alloc) { return inner + alloc.cents; }, 0);
    }, 0);
    var expenseTotal = expenses.reduce(function (sum, expense) { return sum + expense.amountCents; }, 0);
    var unallocated = expenseTotal - allocated;
    report.expenses = expenses;
    report.expenseTotal = expenseTotal;
    report.allocated = allocated;
    report.unallocated = unallocated;
    report.monthProfit = report.totals.confirmed.cents - expenseTotal;
    report.matchedProfit = report.totals.confirmed.cents - allocated;
    return report;
  }

  function productMatches(product, query) {
    if (!query) return true;
    var text = product.title;
    product.attrs.forEach(function (attr) { text += " " + attr.attr; });
    return text.toLowerCase().indexOf(query.toLowerCase()) !== -1;
  }

  function statusOf(cost, profit, confirmedCents) {
    if (cost <= 0 && confirmedCents <= 0) return { cls: "flat", text: "无确认收入" };
    if (cost <= 0) return { cls: "wait", text: "未挂成本" };
    if (profit > 0) return { cls: "profit", text: "赚钱" };
    if (profit < 0) return { cls: "loss", text: "亏损" };
    return { cls: "flat", text: "打平" };
  }

  function visibleProducts(report) {
    var query = state.query.trim();
    var list = report.products.filter(function (product) {
      if (!productMatches(product, query)) return false;
      if (state.filter === "profit") return product.costCents > 0 && product.profitCents > 0;
      if (state.filter === "loss") return product.profitCents < 0;
      if (state.filter === "nocost") return product.confirmed.cents > 0 && product.costCents === 0;
      if (state.filter === "costed") return product.costCents > 0;
      return product.confirmed.cents || product.pending.cents || product.later.cents || product.refund.cents || product.costCents;
    });
    list.sort(function (a, b) {
      if (state.sort === "profit") return a.profitCents - b.profitCents;
      if (state.sort === "name") return a.title.localeCompare(b.title, "zh");
      return b.confirmed.cents - a.confirmed.cents;
    });
    return list;
  }

  function expenseInMonth(expense) {
    return !expense.month || expense.month === state.month;
  }

  function assignExpense(expenseId, targetKey, targetLabel) {
    var expense = state.expenses.filter(function (item) { return item.id === expenseId; })[0];
    if (!expense) return;
    expense.allocs = [{ targetKey: targetKey, targetLabel: targetLabel, cents: expense.amountCents }];
    state.selectedExpenseId = expenseId;
    state.selectedProductKey = targetKey;
    saveState();
    toast("已把「" + expense.name + "」分到「" + targetLabel + "」");
  }

  function clearAlloc(expenseId) {
    var expense = state.expenses.filter(function (item) { return item.id === expenseId; })[0];
    if (!expense) return;
    expense.allocs = [];
    saveState();
    render();
  }

  function splitExpense(expenseId) {
    var expense = state.expenses.filter(function (item) { return item.id === expenseId; })[0];
    var input = document.getElementById("split-" + expenseId);
    if (!expense || !input) return;
    var cents = core.parseMoneyCents(input.value);
    if (cents <= 0 || cents >= expense.amountCents) {
      toast("拆出的金额要大于 0，并且小于这笔支出");
      return;
    }
    var allocated = expense.allocs.reduce(function (sum, alloc) { return sum + alloc.cents; }, 0);
    if (allocated > expense.amountCents - cents) {
      expense.allocs = [];
    } else if (expense.allocs.length === 1) {
      expense.allocs[0].cents = expense.amountCents - cents;
    }
    expense.amountCents -= cents;
    state.expenses.push({
      id: uid(),
      name: expense.name,
      note: expense.note,
      amountCents: cents,
      qty: 0,
      dateStamp: expense.dateStamp,
      month: expense.month,
      sourceFile: expense.sourceFile,
      allocs: []
    });
    saveState();
    toast("已拆出 " + fmtMoney(cents) + " 元，拖到另一个产品上");
  }

  function removeExpense(expenseId) {
    state.expenses = state.expenses.filter(function (expense) { return expense.id !== expenseId; });
    saveState();
    render();
  }

  function linesFor(title, attr) {
    return state.lines.filter(function (line) {
      if (line.title !== title) return false;
      if (attr != null && (line.attr || "") !== attr) return false;
      var confirmMonth = core.monthKey(line.confirmWall);
      var orderMonth = core.monthKey(line.orderWall);
      return confirmMonth === state.month || orderMonth === state.month;
    });
  }

  function exportReport() {
    var report = viewModel();
    var rows = [["商品标题", "商品属性", "当月确认收货数量", "当月确认收货金额", "尚未确认收货数量", "尚未确认收货金额", "本月下单其他月确认数量", "本月下单其他月确认金额", "退款成功数量", "退款成功金额", "分到这里的成本", "利润（确认收货实付-成本）"]];
    report.products.forEach(function (product) {
      rows.push([
        product.title, "（整个商品）",
        product.confirmed.qty, product.confirmed.cents / 100,
        product.pending.qty, product.pending.cents / 100,
        product.later.qty, product.later.cents / 100,
        product.refund.qty, product.refund.cents / 100,
        product.costCents / 100, product.profitCents / 100
      ]);
      product.attrs.forEach(function (attr) {
        rows.push([
          product.title, attr.attr || "（无商品属性）",
          attr.confirmed.qty, attr.confirmed.cents / 100,
          attr.pending.qty, attr.pending.cents / 100,
          attr.later.qty, attr.later.cents / 100,
          attr.refund.qty, attr.refund.cents / 100,
          attr.costCents / 100, attr.profitCents / 100
        ]);
      });
    });
    var expenses = [["日期", "支出项目", "金额", "分到", "备注", "来源文件"]];
    report.expenses.forEach(function (expense) {
      expenses.push([
        expense.dateStamp,
        expense.name,
        expense.amountCents / 100,
        expense.allocs.map(function (alloc) { return alloc.targetLabel; }).join("、"),
        expense.note || "",
        expense.sourceFile || ""
      ]);
    });
    var summary = [
      ["项目", "金额"],
      ["月份", state.month],
      ["当月确认收货买家实付", report.totals.confirmed.cents / 100],
      ["尚未确认收货买家实付", report.totals.pending.cents / 100],
      ["本月支出", report.expenseTotal / 100],
      ["已分到产品的成本", report.allocated / 100],
      ["还没分到产品的支出", report.unallocated / 100],
      ["确认收货实付减全部支出", report.monthProfit / 100]
    ];
    var book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(summary), "汇总");
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "产品利润");
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(expenses), "支出分配");
    XLSX.writeFile(book, "月度算账-" + (state.month || "未选月") + ".xlsx");
  }

  function downloadTemplate() {
    var rows = [
      ["日期", "支出项目", "金额", "数量", "备注"],
      ["2026-10-03", "示例：PCB打样，导入前删掉这一行", 500, 10, "写给自己看的说明"]
    ];
    var book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "支出");
    XLSX.writeFile(book, "支出导入模板.xlsx");
  }

  function bucketCells(bucket) {
    return "<td class='num'>" + fmtQty(bucket.qty) + "</td><td class='num'>" + fmtMoney(bucket.cents) + "</td>";
  }

  function rowHtml(kind, title, attr, bucketConfirmed, bucketPending, cost, profit, key, label) {
    var selected = state.selectedProductKey === key ? " selected" : "";
    var placeholder = title.indexOf("补差价") !== -1 ? "<div class='hint'>这个数量是按差价记的，不是实物件数</div>" : "";
    var name = kind === "title"
      ? "<button class='link' data-act='toggle' data-title='" + esc(title) + "'>" + (state.expanded[title] ? "▾ " : "▸ ") + esc(title) + "</button>"
      : esc(attr || "（无商品属性）");
    var badge = statusOf(cost, profit, bucketConfirmed.cents);
    return "<tr class='" + kind + "-row droppable" + selected + "' data-drop='" + encodeURIComponent(key) + "' data-label='" + esc(label) + "'>" +
      "<td class='product'><div class='product-name'>" + name + placeholder + "</div></td>" +
      bucketCells(bucketConfirmed) +
      bucketCells(bucketPending) +
      "<td class='num'>" + fmtMoney(cost) + "</td>" +
      "<td class='num'>" + fmtMoney(profit) + " <span class='tag " + badge.cls + "'>" + badge.text + "</span></td>" +
      "</tr>";
  }

  function render() {
    var active = document.activeElement;
    var activeId = active && active.id;
    var selection = active && typeof active.selectionStart === "number" ? active.selectionStart : null;
    var report = viewModel();
    var months = monthsInData();
    var products = visibleProducts(report);
    var query = state.expenseQuery.trim();
    var expenses = report.expenses.filter(function (expense) {
      if (!query) return true;
      return (expense.name + " " + (expense.note || "")).toLowerCase().indexOf(query.toLowerCase()) !== -1;
    }).sort(function (a, b) {
      var aOpen = a.allocs.length ? 1 : 0;
      var bOpen = b.allocs.length ? 1 : 0;
      if (aOpen !== bOpen) return aOpen - bOpen;
      return (b.dateStamp || "").localeCompare(a.dateStamp || "");
    });

    var verdict = "先导入收益表。";
    var verdictClass = "";
    if (state.lines.length && !state.expenses.length) {
      verdict = monthLabel(state.month) + "确认收货 " + fmtMoney(report.totals.confirmed.cents) + " 元。支出还没导入，现在还看不出赚没赚钱。";
    } else if (state.lines.length) {
      verdictClass = report.monthProfit > 0 ? "profit" : report.monthProfit < 0 ? "loss" : "";
      var word = report.monthProfit > 0 ? "赚了" : report.monthProfit < 0 ? "亏了" : "打平";
      verdict = "把这个月全部支出算上，" + word + " " + fmtMoney(Math.abs(report.monthProfit)) + " 元。";
    }

    var monthOptions = months.map(function (month) {
      return "<option value='" + esc(month) + "'" + (month === state.month ? " selected" : "") + ">" + esc(monthLabel(month)) + "</option>";
    }).join("");
    if (!monthOptions) monthOptions = "<option value=''>导入后出现月份</option>";

    var filePills = state.files.slice(-8).map(function (file) {
      return "<span class='file-pill'>" + (file.kind === "expense" ? "支出 " : "收益 ") + esc(file.name) + "</span>";
    }).join("");

    var body = "";
    if (!products.length) {
      body = "<tr><td colspan='7'><div class='empty'>这个月还没有对得上的商品。导入淘宝商品订单报表后，同一商品标题、同一商品属性会加在一起。</div></td></tr>";
    } else {
      products.forEach(function (product) {
        body += rowHtml("title", product.title, "", product.confirmed, product.pending, product.costCents, product.profitCents, core.titleKey(product.title), product.title);
        if (state.expanded[product.title]) {
          product.attrs.forEach(function (attr) {
            var label = product.title + " / " + (attr.attr || "（无商品属性）");
            body += rowHtml("attr", product.title, attr.attr, attr.confirmed, attr.pending, attr.costCents, attr.profitCents, core.attrKey(product.title, attr.attr), label);
          });
        }
      });
    }

    var cards = expenses.map(function (expense) {
      var assigned = expense.allocs[0];
      var cls = "card" + (assigned ? " assigned" : "") + (state.selectedExpenseId === expense.id ? " selected" : "");
      var where = assigned
        ? "<div class='meta'>已分到「" + esc(assigned.targetLabel) + "」</div>"
        : "<div class='meta'>还没分到产品。拖到左边，或先点一个产品再点分配。</div>";
      return "<article class='" + cls + "' draggable='true' data-expense='" + esc(expense.id) + "'>" +
        "<div class='card-top'><strong>" + esc(expense.name) + "</strong><span class='money'>" + fmtMoney(expense.amountCents) + "</span></div>" +
        "<div class='meta'>" + esc(fmtDate(expense.dateStamp)) + (expense.note ? " · " + esc(expense.note) : "") + "</div>" +
        where +
        "<div class='card-actions'>" +
          "<button class='btn tiny' data-act='assign-selected' data-id='" + esc(expense.id) + "'>分到选中产品</button>" +
          (assigned ? "<button class='btn tiny' data-act='clear-alloc' data-id='" + esc(expense.id) + "'>取消分配</button>" : "") +
          "<span class='split-box'><input id='split-" + esc(expense.id) + "' placeholder='拆出金额'><button class='btn tiny' data-act='split' data-id='" + esc(expense.id) + "'>拆出</button></span>" +
          "<button class='btn tiny danger' data-act='remove-expense' data-id='" + esc(expense.id) + "'>删除</button>" +
        "</div></article>";
    }).join("");
    if (!cards) {
      cards = "<div class='empty'>支出会出现在这里。每一笔都能拖到左边的商品上。一份采购如果对应几个产品，先拆成几笔再拖。</div>";
    }

    var filters = [
      ["all", "全部"],
      ["profit", "赚钱"],
      ["loss", "亏损"],
      ["nocost", "未挂成本"],
      ["costed", "已挂成本"]
    ].map(function (item) {
      return "<button class='chip" + (state.filter === item[0] ? " on" : "") + "' data-act='filter' data-filter='" + item[0] + "'>" + item[1] + "</button>";
    }).join("");

    app.innerHTML =
      "<header class='top'>" +
        "<div class='brand'><h1>月度算账</h1><p>同一个商品、同一个属性加在一起。利润用当月确认收货的买家实付，减去拖过去的支出。</p></div>" +
        "<div class='top-actions'>" +
          "<span class='month-label'>月份</span>" +
          "<select id='month' class='month'>" + monthOptions + "</select>" +
          "<button class='btn' data-act='pick-revenue'>导入收益表</button>" +
          "<button class='btn' data-act='pick-expense'>导入支出表</button>" +
          "<button class='btn ghost' data-act='template'>支出模板</button>" +
          "<button class='btn ghost' data-act='export'>导出本月</button>" +
          "<button class='btn ghost' data-act='help'>怎么算</button>" +
          "<button class='btn ghost' data-act='reset'>清空</button>" +
        "</div>" +
      "</header>" +
      "<section class='hero'>" +
        "<div><h2 class='" + verdictClass + "'>" + esc(verdict) + "</h2>" +
          "<div class='sub'>尚未确认收货 " + fmtMoney(report.totals.pending.cents) + " 元先不计入利润。已分到产品的成本 " + fmtMoney(report.allocated) + " 元，还没拖过去的支出 " + fmtMoney(report.unallocated) + " 元。" +
          (report.totals.refund.cents ? " 退款成功 " + fmtMoney(report.totals.refund.cents) + " 元已从收入里拿开。" : "") +
          (report.totals.later.cents ? " 另有本月下单、其他月份确认 " + fmtMoney(report.totals.later.cents) + " 元，不算进本月确认收货。" : "") +
          "</div></div>" +
        "<div class='stat'><div class='k'>当月确认收货</div><div class='v'>" + fmtMoney(report.totals.confirmed.cents) + "<small>" + fmtQty(report.totals.confirmed.qty) + " 件</small></div></div>" +
        "<div class='stat'><div class='k'>尚未确认收货</div><div class='v'>" + fmtMoney(report.totals.pending.cents) + "<small>" + fmtQty(report.totals.pending.qty) + " 件</small></div></div>" +
        "<div class='stat'><div class='k'>本月支出</div><div class='v'>" + fmtMoney(report.expenseTotal) + "</div></div>" +
        "<div class='stat'><div class='k'>确认收货减全部支出</div><div class='v'>" + fmtMoney(report.monthProfit) + "</div></div>" +
      "</section>" +
      "<main class='layout'>" +
        "<section class='panel'>" +
          "<div class='panel-head'><h3>收入 · 按商品</h3><div class='tools'>" +
            "<input id='query' type='search' placeholder='搜商品标题或属性' value='" + esc(state.query) + "'>" +
            "<select id='sort'><option value='confirmed'" + (state.sort === "confirmed" ? " selected" : "") + ">按确认收入</option><option value='profit'" + (state.sort === "profit" ? " selected" : "") + ">按利润从低到高</option><option value='name'" + (state.sort === "name" ? " selected" : "") + ">按名称</option></select>" +
          "</div></div>" +
          "<div class='chips'>" + filters + "</div>" +
          (filePills ? "<div class='files'>" + filePills + "</div>" : "") +
          "<div class='table-wrap'><table><thead><tr>" +
            "<th>商品</th><th class='num'>确认数量</th><th class='num'>确认金额</th><th class='num'>未确认数量</th><th class='num'>未确认金额</th><th class='num'>成本</th><th class='num'>利润</th>" +
          "</tr></thead><tbody>" + body + "</tbody></table></div>" +
        "</section>" +
        "<aside class='panel expense-panel'>" +
          "<div class='panel-head'><h3>支出 · " + expenses.length + " 笔</h3><div class='tools'><input id='expense-query' type='search' placeholder='搜支出' value='" + esc(state.expenseQuery) + "'></div></div>" +
          "<div class='expense-list'>" + cards + "</div>" +
        "</aside>" +
      "</main>" +
      "<input id='file-revenue' type='file' accept='.xlsx,.xls' multiple hidden>" +
      "<input id='file-expense' type='file' accept='.xlsx,.xls' multiple hidden>" +
      (state.toast ? "<div class='toast'>" + esc(state.toast) + "</div>" : "") +
      modalHtml(report);

    if (activeId) {
      var next = document.getElementById(activeId);
      if (next) {
        next.focus();
        if (selection != null && next.setSelectionRange) {
          var pos = Math.min(selection, next.value.length);
          next.setSelectionRange(pos, pos);
        }
      }
    }
  }

  function modalHtml() {
    if (state.help) {
      return "<div class='modal-back' data-act='close-modal'><div class='modal help' data-stop='1'><h3>这页怎么算</h3><ol>" +
        "<li>收益表用淘宝商品订单报表。同一商品标题、同一商品属性的购买数量和买家实付会加在一起。</li>" +
        "<li>当月确认收货：确认收货时间落在所选月份。尚未确认收货：付款时间在这个月，没有付款时间就看下单时间，并且确认收货时间为空。</li>" +
        "<li>退款成功不进这两栏。退款关闭留在原栏里。</li>" +
        "<li>没有付款时间、也没有确认收货、实付为 0 的行，当作没付成，不计入。</li>" +
        "<li>支出表至少要有项目和金额。日期写了就会归到对应月份。一份支出对应几个产品时，先拆开再拖。</li>" +
        "<li>拖到商品标题上，成本算整个商品的。拖到展开后的某个属性上，成本只算这个属性。</li>" +
        "<li>单个产品的利润 = 这个月确认收货的买家实付 − 分到它上面的成本。还没确认收货的钱先不算利润。顶部的结余会把还没拖过去的支出也减掉。</li>" +
        "<li>数据只留在这台电脑上，不会上传。用月度算账软件时，记在本机的月度算账数据目录里；用浏览器打开同一页面时，记在这个浏览器里。换浏览器或清掉这些数据后需要重新导入。</li>" +
        "</ol><button class='btn' data-act='close-modal'>知道了</button></div></div>";
    }
    if (state.mapper) {
      var headers = state.mapper.headers;
      function options(selected) {
        return "<option value='-1'>不用</option>" + headers.map(function (header, index) {
          var label = header || "第" + (index + 1) + "列";
          return "<option value='" + index + "'" + (index === selected ? " selected" : "") + ">" + esc(label) + "</option>";
        }).join("");
      }
      return "<div class='modal-back'><div class='modal' data-stop='1'><h3>指定支出表的列</h3><p class='muted'>" + esc(state.mapper.fileName) + " 的表头没有自动对上。</p>" +
        "<div class='mapper'><span>项目</span><select id='map-name'>" + options(0) + "</select><span>金额</span><select id='map-amount'>" + options(1) + "</select><span>日期</span><select id='map-date'>" + options(2) + "</select><span>备注</span><select id='map-note'>" + options(-1) + "</select></div>" +
        "<button class='btn' data-act='confirm-map'>按这个导入</button> <button class='btn tiny' data-act='close-modal'>取消</button></div></div>";
    }
    if (state.detail) {
      var detailLines = linesFor(state.detail.title, state.detail.attr);
      var rows = detailLines.map(function (line) {
        var kind = line.status === "退款成功" ? "退款成功" : (core.monthKey(line.confirmWall) === state.month ? "本月确认" : (!line.confirmWall ? "未确认" : "其他月确认"));
        return "<tr><td>" + esc(line.orderId) + "</td><td>" + esc(line.attr || "（无商品属性）") + "</td><td class='num'>" + fmtQty(line.qty) + "</td><td class='num'>" + fmtMoney(line.paidCents) + "</td><td>" + esc(kind) + "</td><td>" + esc(line.orderStamp) + "</td><td>" + esc(line.confirmStamp || "") + "</td></tr>";
      }).join("");
      return "<div class='modal-back' data-act='close-modal'><div class='modal' data-stop='1'><h3>" + esc(state.detail.title) + "</h3><p class='muted'>这个月相关的订单行</p><table><thead><tr><th>订单</th><th>属性</th><th class='num'>数量</th><th class='num'>实付</th><th>归类</th><th>下单/付款</th><th>确认收货</th></tr></thead><tbody>" + rows + "</tbody></table><button class='btn' data-act='close-modal'>关闭</button></div></div>";
    }
    return "";
  }

  app.addEventListener("click", function (event) {
    if (event.target && event.target.dataset && event.target.dataset.stop) {
      event.stopPropagation();
    }
    var node = event.target.closest("[data-act]");
    if (!node) {
      var row = event.target.closest("[data-drop]");
      if (row && !event.target.closest("button")) {
        state.selectedProductKey = decodeURIComponent(row.dataset.drop);
        render();
      }
      return;
    }
    var act = node.dataset.act;
    if (act === "pick-revenue") document.getElementById("file-revenue").click();
    if (act === "pick-expense") document.getElementById("file-expense").click();
    if (act === "template") downloadTemplate();
    if (act === "export") exportReport();
    if (act === "help") { state.help = true; render(); }
    if (act === "close-modal") {
      if (event.target.closest(".modal") && !event.target.closest("button")) return;
      state.help = false;
      state.mapper = null;
      state.detail = null;
      render();
    }
    if (act === "confirm-map") confirmMapping();
    if (act === "reset") {
      if (window.confirm("清空这台电脑上保存的收益、支出和分配？")) {
        state.lines = [];
        state.expenses = [];
        state.files = [];
        state.month = "";
        state.expanded = {};
        saveState();
        render();
      }
    }
    if (act === "filter") { state.filter = node.dataset.filter; render(); }
    if (act === "toggle") {
      var title = node.dataset.title;
      state.expanded[title] = !state.expanded[title];
      state.selectedProductKey = core.titleKey(title);
      saveState();
      render();
    }
    if (act === "assign-selected") {
      if (!state.selectedProductKey) {
        toast("先在左边点一下要计入的商品");
        return;
      }
      var target = document.querySelector("[data-drop='" + encodeURIComponent(state.selectedProductKey) + "']");
      var label = target ? target.dataset.label : state.selectedProductKey.replace(/^T:/, "").replace(/^A:/, "").replace("\n", " / ");
      assignExpense(node.dataset.id, state.selectedProductKey, label);
    }
    if (act === "clear-alloc") clearAlloc(node.dataset.id);
    if (act === "split") splitExpense(node.dataset.id);
    if (act === "remove-expense") removeExpense(node.dataset.id);
  });

  app.addEventListener("change", function (event) {
    if (event.target.id === "month") {
      state.month = event.target.value;
      saveState();
      render();
    }
    if (event.target.id === "sort") {
      state.sort = event.target.value;
      render();
    }
    if (event.target.id === "file-revenue") {
      importRevenueFiles(event.target.files);
      event.target.value = "";
    }
    if (event.target.id === "file-expense") {
      importExpenseFiles(event.target.files);
      event.target.value = "";
    }
  });

  function rememberSearch(event) {
    if (event.target.id !== "query" && event.target.id !== "expense-query") return;
    if (event.target.id === "query") state.query = event.target.value;
    if (event.target.id === "expense-query") state.expenseQuery = event.target.value;
    if (event.isComposing) return;
    render();
  }
  app.addEventListener("input", rememberSearch);
  app.addEventListener("compositionend", rememberSearch);

  app.addEventListener("dragstart", function (event) {
    var card = event.target.closest("[data-expense]");
    if (!card) return;
    event.dataTransfer.setData("text/plain", card.dataset.expense);
    event.dataTransfer.effectAllowed = "move";
    state.selectedExpenseId = card.dataset.expense;
  });

  app.addEventListener("dragover", function (event) {
    var row = event.target.closest("[data-drop]");
    if (!row) return;
    event.preventDefault();
    document.querySelectorAll(".drop-hover").forEach(function (node) {
      if (node !== row) node.classList.remove("drop-hover");
    });
    row.classList.add("drop-hover");
  });

  app.addEventListener("drop", function (event) {
    var row = event.target.closest("[data-drop]");
    if (!row) return;
    event.preventDefault();
    row.classList.remove("drop-hover");
    var expenseId = event.dataTransfer.getData("text/plain") || state.selectedExpenseId;
    assignExpense(expenseId, decodeURIComponent(row.dataset.drop), row.dataset.label);
  });

  app.addEventListener("dragend", function () {
    document.querySelectorAll(".drop-hover").forEach(function (node) { node.classList.remove("drop-hover"); });
  });

  function cssEscape(value) {
    if (window.CSS && CSS.escape) return CSS.escape(value);
    return String(value).replace(/"/g, '\\"');
  }

  loadState().then(function (saved) {
    if (saved) {
      state.lines = saved.lines || [];
      state.expenses = saved.expenses || [];
      state.files = saved.files || [];
      state.month = saved.month || "";
      state.expanded = saved.expanded || {};
    }
    ensureMonth();
    render();
  });
})();
