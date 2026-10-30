const FIELDS = {
  orders: [
    ["title", "商品标题"],
    ["attr", "商品属性"],
    ["qty", "数量"],
    ["paid", "实付金额"],
    ["pay_time", "付款时间"],
    ["confirm_time", "确认收货时间"],
    ["refund_status", "退款状态"],
    ["refund_amount", "退款金额"],
    ["order_no", "订单编号"],
    ["product_id", "商品ID"],
  ],
  aggregated: [
    ["title", "商品标题"],
    ["attr", "商品属性"],
    ["qty", "数量"],
    ["paid", "实付金额"],
    ["detail", "成本明细"],
    ["line_cost", "成本金额"],
  ],
  expenses: [
    ["name", "项目"],
    ["amount", "金额"],
    ["spent_on", "日期"],
    ["category", "分类"],
    ["note", "备注"],
    ["link_title", "对应商品"],
  ],
};

const KIND_NAME = {
  orders: "订单明细（收益）",
  aggregated: "已经合并过的收益表",
  expenses: "支出表",
  unknown: "还没认出来",
};

let state = null;
let basis = localStorage.getItem("ledger-basis") || "confirmed";
let month = localStorage.getItem("ledger-month") || "";
let openTitles = new Set();
let toastTimer = 0;

const $ = (sel) => document.querySelector(sel);

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[ch]));
}

function enc(value) {
  return encodeURIComponent(value ?? "");
}

function yuan(value) {
  if (value == null || value === "") return "—";
  const number = Math.round(Number(value) * 100) / 100;
  if (!Number.isFinite(number)) return "—";
  const text = Math.abs(number).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return (number < 0 ? "−" : "") + "¥" + text;
}

function qty(value) {
  const number = Number(value) || 0;
  if (Math.abs(number - Math.round(number)) < 0.001) return Math.round(number).toLocaleString("zh-CN");
  return number.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
}

function monthLabel(value) {
  if (!value || !/^\d{4}-\d{2}$/.test(value)) return value || "";
  const [year, mon] = value.split("-");
  return `${year}年${Number(mon)}月`;
}

function parseMoney(text) {
  const number = Number(String(text).replace(/[,，¥￥\s]/g, ""));
  if (!Number.isFinite(number)) throw new Error("金额不是数字");
  return Math.round(number * 100) / 100;
}

function toast(message, bad) {
  const node = $("#toast");
  node.textContent = message;
  node.classList.toggle("bad", !!bad);
  node.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => node.classList.add("hidden"), 4200);
}

async function api(path, options = {}) {
  const response = await fetch(path, options);
  const type = response.headers.get("content-type") || "";
  if (type.includes("application/json")) {
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败");
    return data;
  }
  if (!response.ok) throw new Error("请求失败");
  return response;
}

function jsonCall(path, method, body) {
  return api(path, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
}

async function reload() {
  const y = window.scrollY;
  const box = $("#expenses");
  const sideY = box ? box.scrollTop : 0;
  state = await api(`/api/state?month=${encodeURIComponent(month)}&basis=${encodeURIComponent(basis)}`);
  month = state.month;
  localStorage.setItem("ledger-month", month);
  localStorage.setItem("ledger-basis", basis);
  render();
  window.scrollTo(0, y);
  const next = $("#expenses");
  if (next) next.scrollTop = sideY;
}

function closeModal() {
  $("#modal").classList.add("hidden");
  $("#modal").innerHTML = "";
}

function openModal(html, ready) {
  const root = $("#modal");
  root.classList.remove("hidden");
  root.innerHTML = `<div class="dialog" role="dialog">${html}</div>`;
  root.onclick = (event) => {
    if (event.target === root) closeModal();
  };
  if (ready) ready(root.querySelector(".dialog"));
}

function statusText(status) {
  return { profit: "赚钱", loss: "亏损", even: "打平", unknown: "成本未挂", none: "无收入" }[status] || "";
}

function filteredProducts() {
  const q = $("#q").value.trim().toLowerCase();
  const track = $("#track").value;
  const status = $("#status").value;
  let rows = state.products.slice();
  if (track) rows = rows.filter((row) => row.track === track);
  if (status === "loss") rows = rows.filter((row) => row.status === "loss");
  if (status === "profit") rows = rows.filter((row) => row.status === "profit");
  if (status === "unknown") rows = rows.filter((row) => row.status === "unknown");
  if (status === "open") rows = rows.filter((row) => row.unconfirmed_amt > 0);
  if (q) {
    rows = rows.filter((row) => {
      const blob = [
        row.title,
        row.track,
        ...row.attrs.map((attr) => attr.attr),
        ...row.allocations.map((alloc) => alloc.name),
      ].join("\n").toLowerCase();
      return blob.includes(q);
    });
  }
  const sort = $("#sort").value;
  const profitKey = (row) => (row.profit == null ? null : row.profit);
  rows.sort((a, b) => {
    if (sort === "name") return a.title.localeCompare(b.title, "zh");
    if (sort === "unknown") return Number(b.status === "unknown") - Number(a.status === "unknown") || b.confirmed_amt - a.confirmed_amt;
    if (sort === "profit" || sort === "loss") {
      const av = profitKey(a);
      const bv = profitKey(b);
      if (av == null && bv == null) return b.confirmed_amt - a.confirmed_amt;
      if (av == null) return 1;
      if (bv == null) return -1;
      return sort === "profit" ? bv - av : av - bv;
    }
    return b.confirmed_amt - a.confirmed_amt;
  });
  return rows;
}

function renderMonths() {
  const months = state.months.length ? state.months : [state.month];
  $("#months").innerHTML = months.map((item) => (
    `<button type="button" data-month="${esc(item)}" class="${item === state.month ? "on" : ""}">${esc(monthLabel(item))}</button>`
  )).join("");
}

function renderSummary() {
  const summary = state.summary;
  let tone = "unknown";
  let seal = "未算";
  let title = "还看不出赚没赚钱";
  const profitText = yuan(Math.abs(summary.month_profit || 0));
  if (!summary.has_expenses) {
    tone = "unknown";
    seal = "未算";
    title = summary.revenue_for_profit ? "还没有支出" : "先导入这个月的表";
  } else if (summary.unknown_count > 0) {
    tone = summary.month_profit < -0.004 ? "loss" : "unknown";
    seal = "待核";
    title = `按已导入的支出，${summary.month_profit < -0.004 ? "亏" : "盈"} ${profitText}`;
  } else if (summary.month_profit > 0.004) {
    tone = "profit";
    seal = "赚钱";
    title = `这个月赚钱 ${yuan(summary.month_profit)}`;
  } else if (summary.month_profit < -0.004) {
    tone = "loss";
    seal = "亏损";
    title = `这个月亏损 ${profitText}`;
  } else {
    tone = "even";
    seal = "打平";
    title = "这个月打平";
  }
  $("#summary").innerHTML = `
    <div class="verdict ${tone}">
      <div class="seal ${tone}">${seal}</div>
      <div>
        <h2>${esc(monthLabel(summary.month))} · ${esc(title)}</h2>
        <p class="sentence">${esc(summary.sentence)}</p>
        <div class="metrics">
          <div class="metric"><span>确认收货</span><strong>${yuan(summary.confirmed_amt)}</strong><em>${qty(summary.confirmed_qty)} 件</em></div>
          <div class="metric"><span>未确认收货</span><strong>${yuan(summary.unconfirmed_amt)}</strong><em>${qty(summary.unconfirmed_qty)} 件</em></div>
          <div class="metric"><span>退款成功，已剔除</span><strong>${yuan(summary.refund_amt)}</strong><em>${qty(summary.refund_qty)} 件</em></div>
          <div class="metric"><span>支出合计</span><strong>${summary.has_expenses ? yuan(summary.expense_total) : "—"}</strong><em>已挂到商品 ${yuan(summary.allocated_total)}</em></div>
          <div class="metric ${summary.unallocated_total > 0 ? "warn" : ""}"><span>还没挂到商品</span><strong>${summary.has_expenses ? yuan(summary.unallocated_total) : "—"}</strong><em>整月盈亏已经扣过</em></div>
          <div class="metric"><span>商品</span><strong>${summary.product_count}</strong><em>赚 ${summary.profit_count} · 亏 ${summary.loss_count} · 未挂 ${summary.unknown_count}</em></div>
        </div>
        ${summary.backlog_amt > 0 ? `<p class="note">以前付款、到这个月还没确认收货的还有 ${yuan(summary.backlog_amt)}（${qty(summary.backlog_qty)} 件）。它们留在付款那个月的「未确认」里，没有算进这个月。</p>` : ""}
      </div>
    </div>`;
}

function renderTrackOptions() {
  const select = $("#track");
  const current = select.value;
  const tracks = state.tracks;
  select.innerHTML = `<option value="">全部赛道</option>` + tracks.map((track) => (
    `<option value="${esc(track)}">${esc(track)}</option>`
  )).join("");
  if ([...select.options].some((option) => option.value === current)) select.value = current;
}

function renderProducts() {
  renderTrackOptions();
  document.querySelectorAll(".basis button").forEach((button) => {
    button.classList.toggle("on", button.dataset.basis === basis);
  });
  const rows = filteredProducts();
  if (!state.products.length) {
    $("#products").innerHTML = `
      <div class="empty">
        <strong>这个月还没有收益。</strong><br>
        点上面的「导入收益」，选择淘宝订单导出的财务报表。软件会按商品标题和商品属性把数量、金额加在一起，并分成确认收货和未确认收货。<br>
        然后点「导入支出」，或先<button type="button" id="empty-demo" class="linkish">看一个小例子</button>熟悉怎么拖。
      </div>`;
    return;
  }
  if (!rows.length) {
    $("#products").innerHTML = `<div class="empty">没有符合筛选的商品。</div>`;
    return;
  }
  $("#products").innerHTML = rows.map((product) => productCard(product)).join("");
}

function productCard(product) {
  const opened = openTitles.has(product.title);
  const tone = product.status === "profit" ? "profit" : product.status === "loss" ? "loss" : product.status === "unknown" ? "unknown" : "";
  const profit = product.profit == null ? "成本未挂" : yuan(product.profit);
  const names = product.allocations.slice(0, 3).map((alloc) => alloc.name).join("、");
  return `
    <article class="product" data-drop="${enc(product.title)}">
      <button type="button" class="prow" data-toggle="${enc(product.title)}">
        <span class="twist">${opened ? "▾" : "▸"}</span>
        <span class="pname">
          <b>${esc(product.title)}</b>
          <small>${product.track ? esc(product.track) + " · " : ""}${product.attr_count} 个属性${names ? " · " + esc(names) : ""}</small>
        </span>
        <span class="num"><span>确认收货</span><strong>${yuan(product.confirmed_amt)}</strong><em>${qty(product.confirmed_qty)} 件</em></span>
        <span class="num"><span>未确认</span><strong>${yuan(product.unconfirmed_amt)}</strong><em>${qty(product.unconfirmed_qty)} 件</em></span>
        <span class="num"><span>已挂成本</span><strong>${product.cost ? yuan(product.cost) : "—"}</strong><em>${product.refund_amt ? "退款 " + yuan(product.refund_amt) : ""}</em></span>
        <span class="num ${tone}"><span>${esc(statusText(product.status))}</span><strong>${esc(profit)}</strong></span>
      </button>
      ${opened ? productBody(product) : ""}
    </article>`;
}

function productBody(product) {
  const revenue = basis === "all" ? product.confirmed_amt + product.unconfirmed_amt : product.confirmed_amt;
  const chips = product.allocations.map((alloc) => `
    <div class="chip" draggable="true" data-alloc="${alloc.id}">
      <span>${esc(alloc.name)}${alloc.attr ? " · " + esc(alloc.attr) : ""}</span>
      <input data-alloc-amt="${alloc.id}" value="${alloc.amount.toFixed(2)}" aria-label="分摊金额">
      <button type="button" data-unalloc="${alloc.id}">拿下</button>
    </div>`).join("");
  const attrs = product.attrs.map((attr) => `
    <tr data-drop="${enc(product.title)}" data-attr="${enc(attr.attr_raw)}" data-attr-row="1">
      <td>${esc(attr.attr)}</td>
      <td>${qty(attr.confirmed_qty)}</td>
      <td>${yuan(attr.confirmed_amt)}</td>
      <td>${qty(attr.unconfirmed_qty)}</td>
      <td>${yuan(attr.unconfirmed_amt)}</td>
      <td>${(attr.share * 100).toFixed(1)}%</td>
      <td>${attr.estimated_cost == null ? "—" : yuan(attr.estimated_cost)}</td>
      <td>${attr.estimated_profit == null ? "—" : yuan(attr.estimated_profit)}</td>
    </tr>`).join("");
  return `
    <div class="pbody">
      <p class="note">这个商品按当前口径的收入是 ${yuan(revenue)}。${product.profit == null ? "还没有成本，先不要当成赚钱。" : `利润 = 收入 ${yuan(revenue)} − 成本 ${yuan(product.cost)}。`}</p>
      <div class="chips">${chips || `<span class="note">把右边的支出拖到这里。</span>`}</div>
      <label class="flag">
        <input type="checkbox" data-nocost="${enc(product.title)}" ${product.no_cost ? "checked" : ""} ${product.cost > 0 ? "disabled" : ""}>
        这个商品本月没有支出
      </label>
      <div class="track-edit">
        <span>赛道</span>
        <input data-track="${enc(product.title)}" value="${esc(product.track || "")}" placeholder="例如 新能源赛道">
      </div>
      <table class="attrs">
        <thead>
          <tr>
            <th>商品属性</th><th>确认数量</th><th>确认金额</th><th>未确认数量</th><th>未确认金额</th><th>占该商品</th><th>估算成本</th><th>估算利润</th>
          </tr>
        </thead>
        <tbody>${attrs}</tbody>
      </table>
      <p class="note">估算成本把挂在整个商品上的支出，按金额占比分到每个属性，方便看哪个规格在拖后腿。把支出拖到某一行属性上，那一笔就只算这个属性的。</p>
    </div>`;
}

function renderSide() {
  const q = $("#eq").value.trim().toLowerCase();
  let rows = state.expenses.slice();
  if (q) {
    rows = rows.filter((row) => `${row.name}\n${row.category}\n${row.note}`.toLowerCase().includes(q));
  }
  const open = rows.filter((row) => row.remaining > 0.004);
  const done = rows.filter((row) => row.remaining <= 0.004);
  if (!state.expenses.length) {
    $("#expenses").innerHTML = `<div class="empty">还没有这个月的支出。导入支出表，或记一笔。每一笔都会出现在这里，拖到对应的商品上。</div>`;
    return;
  }
  $("#expenses").innerHTML = `
    <div class="expense-group">
      <h3>还没挂完 ${open.length}</h3>
      ${open.map(expenseCard).join("") || `<div class="empty">都挂到商品上了。</div>`}
    </div>
    <div class="expense-group">
      <h3>已经挂完 ${done.length}</h3>
      ${done.map((row) => expenseCard(row, true)).join("")}
    </div>`;
}

function expenseCard(expense, done) {
  const where = expense.allocations.map((alloc) => `${alloc.title} ${yuan(alloc.amount)}`).join("；");
  return `
    <article class="expense ${done ? "done" : ""}" draggable="${done ? "false" : "true"}" data-expense="${expense.id}">
      <header>
        <strong>${esc(expense.name)}</strong>
        <b>${yuan(expense.amount)}</b>
      </header>
      <div class="rest">${done ? "已全部挂上" : `还剩 ${yuan(expense.remaining)}`}</div>
      <div class="meta">
        ${expense.category ? esc(expense.category) + " · " : ""}${expense.spent_on ? esc(expense.spent_on) : monthLabel(expense.month)}
        ${expense.note ? "<br>" + esc(expense.note) : ""}
        ${where ? "<br>已挂：" + esc(where) : ""}
      </div>
      <div class="ops">
        ${expense.suggestion && expense.remaining > 0 ? `<button type="button" class="linkish" data-suggest="${expense.id}" data-title="${enc(expense.suggestion)}">挂到 ${esc(shortTitle(expense.suggestion))}</button>` : ""}
        <button type="button" class="linkish" data-assign="${expense.id}">挂到…</button>
        <button type="button" class="linkish" data-edit="${expense.id}">修改</button>
        <button type="button" class="linkish" data-del-exp="${expense.id}">删除</button>
      </div>
    </article>`;
}

function shortTitle(title) {
  const code = (title.match(/^[A-Za-z]{2,}(?:-[A-Za-z]*)?\d+/) || [])[0];
  return code || (title.length > 18 ? title.slice(0, 18) + "…" : title);
}

function render() {
  renderMonths();
  renderSummary();
  renderProducts();
  renderSide();
}

function zoneFrom(target) {
  const attrRow = target.closest("[data-attr-row]");
  const product = target.closest("[data-drop]");
  const node = attrRow || product;
  if (!node) return null;
  return {
    title: decodeURIComponent(node.dataset.drop || ""),
    attr: attrRow ? decodeURIComponent(attrRow.dataset.attr || "") : "",
    node,
  };
}

async function previewFile(file, intent) {
  if (!file) return;
  toast("正在读表格…");
  const form = new FormData();
  form.append("file", file);
  try {
    const info = await api("/api/preview", { method: "POST", body: form });
    openPreview(info, intent);
  } catch (error) {
    toast(error.message, true);
  }
}

function openPreview(info, intent) {
  let kind = info.kind;
  if (kind === "unknown") kind = intent === "expense" ? "expenses" : "orders";
  const pending = {
    info,
    intent,
    kind,
    includeCosts: info.stats.cost_rows > 0,
    allowOverlap: false,
    month: info.suggested_month || state.month || "",
    track: info.stats.track || "",
  };
  drawPreview(pending);
}

function drawPreview(pending) {
  const info = pending.info;
  const warnings = info.warnings.slice();
  if (pending.intent === "revenue" && pending.kind === "expenses") {
    warnings.unshift("你点的是导入收益，但这张表更像支出。如果确实是支出，把类型改成支出表。");
  }
  if (pending.intent === "expense" && (pending.kind === "orders" || pending.kind === "aggregated")) {
    warnings.unshift("你点的是导入支出，但这张表更像收益。");
  }
  const months = Object.entries(info.stats.months || {}).map(([key, count]) => `${key} ${count} 行`).join("，");
  const fields = FIELDS[pending.kind] || FIELDS.orders;
  const mapping = info.mapping || {};
  openModal(`
    <h3>${esc(info.filename || "导入")}</h3>
    <p>工作表「${esc(info.sheet)}」，认出 ${info.stats.rows} 行${info.stats.cost_rows ? `，另外有 ${info.stats.cost_rows} 笔成本` : ""}。${months ? "日期分布：" + esc(months) + "。" : ""}</p>
    ${warnings.map((text) => `<div class="warn">${esc(text)}</div>`).join("")}
    <div class="grid-2">
      <label class="field">表格类型
        <select id="imp-kind">
          <option value="orders" ${pending.kind === "orders" ? "selected" : ""}>订单明细（收益）</option>
          <option value="aggregated" ${pending.kind === "aggregated" ? "selected" : ""}>已经合并的收益表</option>
          <option value="expenses" ${pending.kind === "expenses" ? "selected" : ""}>支出表</option>
        </select>
      </label>
      <label class="field">${pending.kind === "aggregated" ? "这张表代表哪个月" : pending.kind === "expenses" ? "没有日期的支出算进哪个月" : "没有日期的行归入（默认不导入）"}
        <input id="imp-month" type="month" value="${esc(pending.month)}">
      </label>
      <label class="field full">赛道（可以空着）
        <input id="imp-track" value="${esc(pending.track)}" placeholder="例如 新能源赛道">
      </label>
    </div>
    ${pending.kind === "orders" && (info.stats.months || {})["（没有日期）"] ? `<label class="flag"><input id="imp-undated" type="checkbox"> 仍然导入这 ${(info.stats.months || {})["（没有日期）"]} 行没有日期的记录，并算进上面的月份（一般不要勾）</label>` : ""}
    ${pending.kind === "aggregated" ? `<label class="flag"><input id="imp-costs" type="checkbox" ${pending.includeCosts ? "checked" : ""}> 同时导入表里「明细」旁边的成本，并挂到对应商品</label>` : ""}
    ${pending.kind === "aggregated" ? `<label class="flag"><input id="imp-overlap" type="checkbox"> 这个月可能已经有订单明细，仍然导入</label>` : ""}
    <details ${info.needs_mapping ? "open" : ""}>
      <summary>列对得上吗</summary>
      <div class="grid-2" style="margin-top:8px">
        ${fields.map(([key, label]) => `
          <label class="field">${esc(label)}
            <select data-map="${esc(key)}">
              <option value="">不使用</option>
              ${headerOptions(info.header_row, mapping[key])}
            </select>
          </label>`).join("")}
      </div>
    </details>
    ${previewTable(info.preview)}
    <div class="actions">
      <button type="button" class="btn" id="imp-cancel">取消</button>
      <button type="button" class="btn solid" id="imp-ok">导入</button>
    </div>
  `, (dialog) => {
    dialog.querySelector("#imp-cancel").onclick = closeModal;
    dialog.querySelector("#imp-kind").onchange = (event) => {
      pending.kind = event.target.value;
      pending.month = dialog.querySelector("#imp-month").value;
      pending.track = dialog.querySelector("#imp-track").value;
      drawPreview(pending);
    };
    dialog.querySelector("#imp-ok").onclick = () => commitPreview(dialog, pending);
  });
}

function headerOptions(headers, selected) {
  const cells = (headers || []).slice();
  let last = cells.length - 1;
  while (last > 0 && !cells[last]) last -= 1;
  return cells.slice(0, last + 1).map((header, index) => (
    `<option value="${index}" ${String(selected) === String(index) ? "selected" : ""}>${index + 1}. ${esc(header || "（空）")}</option>`
  )).join("");
}

function previewTable(rows) {
  if (!rows || !rows.length) return "";
  const keys = Object.keys(rows[0]);
  return `<table class="attrs" style="margin-top:10px"><thead><tr>${keys.map((key) => `<th>${esc(key)}</th>`).join("")}</tr></thead><tbody>
    ${rows.map((row) => `<tr>${keys.map((key) => `<td>${esc(row[key])}</td>`).join("")}</tr>`).join("")}
  </tbody></table>`;
}

async function commitPreview(dialog, pending) {
  const mapping = {};
  dialog.querySelectorAll("[data-map]").forEach((select) => {
    if (select.value !== "") mapping[select.dataset.map] = Number(select.value);
  });
  const chosenMonth = dialog.querySelector("#imp-month").value;
  const button = dialog.querySelector("#imp-ok");
  button.disabled = true;
  try {
    const result = await jsonCall("/api/commit", "POST", {
      token: pending.info.token,
      kind: pending.kind,
      month: chosenMonth,
      include_costs: !!dialog.querySelector("#imp-costs")?.checked,
      allow_overlap: !!dialog.querySelector("#imp-overlap")?.checked,
      assign_undated: !!dialog.querySelector("#imp-undated")?.checked,
      track: dialog.querySelector("#imp-track").value.trim(),
      mapping,
    });
    closeModal();
    if (/^\d{4}-\d{2}$/.test(chosenMonth)) month = chosenMonth;
    else if (pending.info.suggested_month) month = pending.info.suggested_month;
    const bits = [`导入 ${result.inserted} 行`];
    if (result.updated) bits.push(`更新 ${result.updated} 行`);
    if (result.skipped) bits.push(`跳过 ${result.skipped} 行`);
    if (result.costs) bits.push(`成本 ${result.costs} 笔`);
    if (result.linked) bits.push(`自动挂上 ${result.linked} 笔`);
    toast(bits.join("，"));
    await reload();
  } catch (error) {
    button.disabled = false;
    toast(error.message, true);
  }
}

function openAddExpense() {
  const titles = state.products.map((product) => product.title);
  openModal(`
    <h3>记一笔支出</h3>
    <div class="grid-2">
      <label class="field">项目<input id="ex-name" placeholder="例如 PCB板加工"></label>
      <label class="field">金额<input id="ex-amt" inputmode="decimal" placeholder="0.00"></label>
      <label class="field">日期<input id="ex-date" type="date" value="${state.month}-01"></label>
      <label class="field">月份<input id="ex-month" type="month" value="${esc(state.month)}"></label>
      <label class="field">分类<input id="ex-cat" placeholder="加工 / 物料 / 物流" list="cats"></label>
      <label class="field">立刻挂到
        <select id="ex-title"><option value="">先不挂，放在右边</option>${titles.map((title) => `<option value="${esc(title)}">${esc(title)}</option>`).join("")}</select>
      </label>
      <label class="field full">备注<textarea id="ex-note" rows="2"></textarea></label>
    </div>
    <datalist id="cats">${state.categories.map((cat) => `<option value="${esc(cat)}">`).join("")}</datalist>
    <div class="actions">
      <button type="button" class="btn" id="ex-cancel">取消</button>
      <button type="button" class="btn solid" id="ex-ok">记下</button>
    </div>
  `, (dialog) => {
    dialog.querySelector("#ex-cancel").onclick = closeModal;
    dialog.querySelector("#ex-name").focus();
    dialog.querySelector("#ex-ok").onclick = async () => {
      try {
        const spent = dialog.querySelector("#ex-date").value;
        await jsonCall("/api/expenses", "POST", {
          name: dialog.querySelector("#ex-name").value,
          amount: parseMoney(dialog.querySelector("#ex-amt").value),
          month: dialog.querySelector("#ex-month").value || (spent ? spent.slice(0, 7) : state.month),
          spent_on: spent,
          category: dialog.querySelector("#ex-cat").value,
          note: dialog.querySelector("#ex-note").value,
          title: dialog.querySelector("#ex-title").value,
        });
        closeModal();
        toast("已记下");
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    };
  });
}

function openEditExpense(expense) {
  openModal(`
    <h3>修改支出</h3>
    <div class="grid-2">
      <label class="field">项目<input id="ex-name" value="${esc(expense.name)}"></label>
      <label class="field">金额<input id="ex-amt" value="${expense.amount.toFixed(2)}"></label>
      <label class="field">日期<input id="ex-date" type="date" value="${esc((expense.spent_on || "").slice(0, 10))}"></label>
      <label class="field">月份<input id="ex-month" type="month" value="${esc(expense.month)}"></label>
      <label class="field">分类<input id="ex-cat" value="${esc(expense.category)}"></label>
      <label class="field full">备注<textarea id="ex-note" rows="2">${esc(expense.note)}</textarea></label>
    </div>
    <div class="actions">
      <button type="button" class="btn" id="ex-cancel">取消</button>
      <button type="button" class="btn solid" id="ex-ok">保存</button>
    </div>
  `, (dialog) => {
    dialog.querySelector("#ex-cancel").onclick = closeModal;
    dialog.querySelector("#ex-ok").onclick = async () => {
      try {
        await jsonCall(`/api/expenses/${expense.id}`, "PATCH", {
          name: dialog.querySelector("#ex-name").value,
          amount: parseMoney(dialog.querySelector("#ex-amt").value),
          month: dialog.querySelector("#ex-month").value,
          spent_on: dialog.querySelector("#ex-date").value,
          category: dialog.querySelector("#ex-cat").value,
          note: dialog.querySelector("#ex-note").value,
        });
        closeModal();
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    };
  });
}

function openAssign(expense) {
  let picked = "";
  const draw = (query) => {
    const rows = state.products.filter((product) => !query || product.title.toLowerCase().includes(query.toLowerCase()));
    return rows.slice(0, 80).map((product) => (
      `<button type="button" data-pick="${enc(product.title)}" class="${picked === product.title ? "on" : ""}">${esc(product.title)}</button>`
    )).join("") || `<div class="empty">没有这个商品。可以先导入收益。</div>`;
  };
  openModal(`
    <h3>把「${esc(expense.name)}」挂到商品</h3>
    <p class="note">还剩 ${yuan(expense.remaining)}。可以只挂一部分。</p>
    <label class="field">金额<input id="as-amt" value="${expense.remaining.toFixed(2)}"></label>
    <input id="as-q" placeholder="搜商品" style="width:100%;margin-bottom:8px">
    <div class="pick-list" id="as-list">${draw("")}</div>
    <div class="actions">
      <button type="button" class="btn" id="as-cancel">取消</button>
      <button type="button" class="btn solid" id="as-ok">挂上</button>
    </div>
  `, (dialog) => {
    dialog.querySelector("#as-cancel").onclick = closeModal;
    dialog.querySelector("#as-q").oninput = (event) => {
      dialog.querySelector("#as-list").innerHTML = draw(event.target.value.trim());
    };
    dialog.querySelector("#as-list").onclick = (event) => {
      const button = event.target.closest("[data-pick]");
      if (!button) return;
      picked = decodeURIComponent(button.dataset.pick);
      dialog.querySelectorAll("[data-pick]").forEach((node) => node.classList.toggle("on", node === button));
    };
    dialog.querySelector("#as-ok").onclick = async () => {
      if (!picked) return toast("先选一个商品", true);
      try {
        await jsonCall("/api/allocations", "POST", {
          expense_id: expense.id,
          title: picked,
          amount: parseMoney(dialog.querySelector("#as-amt").value),
        });
        closeModal();
        openTitles.add(picked);
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    };
  });
}

function openImports() {
  const kindName = { orders: "收益明细", aggregated: "合并收益", expenses: "支出", demo: "示例" };
  const rows = state.imports.length ? state.imports.map((item) => `
    <tr>
      <td>${esc(item.filename || "")}<br><span class="note">${esc(kindName[item.kind] || item.kind)} · ${esc(item.created_at || "").replace("T", " ")} · ${item.row_count} 行</span></td>
      <td><button type="button" class="linkish" data-del-imp="${item.id}">删除这次导入</button></td>
    </tr>`).join("") : `<tr><td>还没有导入记录。</td></tr>`;
  openModal(`
    <h3>导入记录</h3>
    <p class="note">删掉某一次导入，会把它带进来的收益或支出一起删掉，已经挂上去的成本也会消失。重复导入同一份订单时，软件会更新确认收货和退款，不会把收入加两次。</p>
    <table class="imports">${rows}</table>
    <div class="actions"><button type="button" class="btn" id="imp-close">关闭</button></div>
  `, (dialog) => {
    dialog.querySelector("#imp-close").onclick = closeModal;
    dialog.onclick = async (event) => {
      const button = event.target.closest("[data-del-imp]");
      if (!button) return;
      if (!window.confirm("删除这次导入？相关的收益或支出都会去掉。")) return;
      try {
        await api(`/api/imports/${button.dataset.delImp}`, { method: "DELETE" });
        closeModal();
        toast("已删除这次导入");
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    };
  });
}

function download(path) {
  const link = document.createElement("a");
  link.href = path;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function bind() {
  $("#btn-revenue").onclick = () => $("#file-revenue").click();
  $("#btn-expense").onclick = () => $("#file-expense").click();
  $("#file-revenue").onchange = (event) => {
    previewFile(event.target.files[0], "revenue");
    event.target.value = "";
  };
  $("#file-expense").onchange = (event) => {
    previewFile(event.target.files[0], "expense");
    event.target.value = "";
  };
  $("#btn-add").onclick = openAddExpense;
  $("#btn-export").onclick = () => download(`/api/export?month=${encodeURIComponent(state.month)}&basis=${encodeURIComponent(basis)}`);
  $("#btn-template").onclick = () => download("/api/template");
  $("#btn-backup").onclick = () => download("/api/backup");
  $("#btn-restore").onclick = () => $("#file-restore").click();
  $("#file-restore").onchange = async (event) => {
    const file = event.target.files[0];
    event.target.value = "";
    if (!file) return;
    if (!window.confirm("用这份备份覆盖现在的账？当前数据会被换掉。")) return;
    const form = new FormData();
    form.append("file", file);
    try {
      await api("/api/restore", { method: "POST", body: form });
      toast("已恢复备份");
      await reload();
    } catch (error) {
      toast(error.message, true);
    }
  };
  $("#btn-demo").onclick = async () => {
    try {
      const result = await api("/api/demo", { method: "POST" });
      month = result.month;
      toast("已放进 2099年1月，可以试着把「示例包装」拖到商品上");
      await reload();
    } catch (error) {
      toast(error.message, true);
    }
  };
  $("#btn-imports").onclick = openImports;
  $("#months").onclick = async (event) => {
    const button = event.target.closest("[data-month]");
    if (!button) return;
    month = button.dataset.month;
    await reload();
  };
  $("#q").oninput = renderProducts;
  $("#eq").oninput = renderSide;
  $("#track").onchange = renderProducts;
  $("#status").onchange = renderProducts;
  $("#sort").onchange = renderProducts;
  document.querySelector(".basis").onclick = async (event) => {
    const button = event.target.closest("[data-basis]");
    if (!button || button.dataset.basis === basis) return;
    basis = button.dataset.basis;
    await reload();
  };
  $("#products").onclick = async (event) => {
    if (event.target.closest("#empty-demo")) {
      $("#btn-demo").click();
      return;
    }
    const toggle = event.target.closest("[data-toggle]");
    if (toggle && !event.target.closest("input, .chip")) {
      const title = decodeURIComponent(toggle.dataset.toggle);
      if (openTitles.has(title)) openTitles.delete(title);
      else openTitles.add(title);
      renderProducts();
      return;
    }
    const unalloc = event.target.closest("[data-unalloc]");
    if (unalloc) {
      try {
        await api(`/api/allocations/${unalloc.dataset.unalloc}`, { method: "DELETE" });
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    }
  };
  $("#products").change = null;
  $("#products").addEventListener("change", async (event) => {
    const amount = event.target.closest("[data-alloc-amt]");
    if (amount) {
      try {
        const value = parseMoney(amount.value);
        if (value <= 0) {
          await api(`/api/allocations/${amount.dataset.allocAmt}`, { method: "DELETE" });
        } else {
          await jsonCall(`/api/allocations/${amount.dataset.allocAmt}`, "PATCH", { amount: value });
        }
        await reload();
      } catch (error) {
        toast(error.message, true);
        await reload();
      }
      return;
    }
    const noCost = event.target.closest("[data-nocost]");
    if (noCost) {
      try {
        await jsonCall("/api/flags", "POST", {
          title: decodeURIComponent(noCost.dataset.nocost),
          month: state.month,
          no_cost: noCost.checked,
        });
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
      return;
    }
    const track = event.target.closest("[data-track]");
    if (track) {
      try {
        await jsonCall("/api/meta", "POST", {
          title: decodeURIComponent(track.dataset.track),
          track: track.value.trim(),
        });
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    }
  });
  $("#expenses").onclick = async (event) => {
    const suggest = event.target.closest("[data-suggest]");
    if (suggest) {
      event.stopPropagation();
      try {
        await jsonCall("/api/allocations", "POST", {
          expense_id: Number(suggest.dataset.suggest),
          title: decodeURIComponent(suggest.dataset.title),
        });
        openTitles.add(decodeURIComponent(suggest.dataset.title));
        toast("已挂上，金额可以再改小");
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
      return;
    }
    const assign = event.target.closest("[data-assign]");
    if (assign) {
      openAssign(state.expenses.find((row) => String(row.id) === assign.dataset.assign));
      return;
    }
    const edit = event.target.closest("[data-edit]");
    if (edit) {
      openEditExpense(state.expenses.find((row) => String(row.id) === edit.dataset.edit));
      return;
    }
    const del = event.target.closest("[data-del-exp]");
    if (del) {
      if (!window.confirm("删除这笔支出？挂到商品上的金额也会去掉。")) return;
      try {
        await api(`/api/expenses/${del.dataset.delExp}`, { method: "DELETE" });
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    }
  };

  document.addEventListener("dragstart", (event) => {
    if (event.target.closest("input, textarea, button, select, a")) {
      event.preventDefault();
      return;
    }
    const expense = event.target.closest("[data-expense]");
    const alloc = event.target.closest("[data-alloc]");
    if (expense && expense.getAttribute("draggable") === "true") {
      event.dataTransfer.setData("text/plain", `expense:${expense.dataset.expense}`);
      event.dataTransfer.effectAllowed = "copy";
      document.body.classList.add("is-dragging");
      return;
    }
    if (alloc) {
      event.dataTransfer.setData("text/plain", `alloc:${alloc.dataset.alloc}`);
      event.dataTransfer.effectAllowed = "move";
      document.body.classList.add("is-dragging");
    }
  });
  document.addEventListener("dragend", () => {
    document.body.classList.remove("is-dragging");
    document.querySelectorAll(".drop").forEach((node) => node.classList.remove("drop"));
  });
  document.addEventListener("dragover", (event) => {
    const types = [...event.dataTransfer.types];
    if (types.includes("Files")) event.preventDefault();
    const zone = zoneFrom(event.target);
    if (zone && document.body.classList.contains("is-dragging")) {
      event.preventDefault();
      zone.node.classList.add("drop");
    }
    if (event.target.closest("#side") && document.body.classList.contains("is-dragging")) {
      event.preventDefault();
    }
  });
  document.addEventListener("dragleave", (event) => {
    const zone = zoneFrom(event.target);
    if (zone) zone.node.classList.remove("drop");
  });
  document.addEventListener("drop", async (event) => {
    const text = event.dataTransfer.getData("text/plain");
    const file = event.dataTransfer.files?.[0];
    document.body.classList.remove("is-dragging");
    if (file && !text.startsWith("expense:") && !text.startsWith("alloc:")) {
      event.preventDefault();
      previewFile(file, "auto");
      return;
    }
    if (!text) return;
    const zone = zoneFrom(event.target);
    if (text.startsWith("expense:") && zone) {
      event.preventDefault();
      try {
        await jsonCall("/api/allocations", "POST", {
          expense_id: Number(text.slice(8)),
          title: zone.title,
          attr: zone.attr,
        });
        openTitles.add(zone.title);
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
      return;
    }
    if (text.startsWith("alloc:")) {
      event.preventDefault();
      const id = Number(text.slice(6));
      try {
        if (zone) {
          await jsonCall(`/api/allocations/${id}`, "PATCH", { title: zone.title, attr: zone.attr });
          openTitles.add(zone.title);
        } else if (event.target.closest("#side")) {
          await api(`/api/allocations/${id}`, { method: "DELETE" });
        } else {
          return;
        }
        await reload();
      } catch (error) {
        toast(error.message, true);
      }
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeModal();
  });
}

bind();
reload().catch((error) => toast(error.message, true));
