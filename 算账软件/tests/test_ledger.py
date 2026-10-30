# -*- coding: utf-8 -*-
import io
import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

import ledger


ORDER_HEADERS = [
    "主订单编号", "商品标题", "商品价格", "购买数量", "商品属性", "套餐信息",
    "买家应付货款", "买家实付金额", "退款状态", "退款金额", "订单付款时间",
    "商品ID", "确认收货时间",
]


def make_xlsx(sheets):
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets:
        ws = wb.active if first else wb.create_sheet(name)
        if first:
            ws.title = name
            first = False
        for row in rows:
            ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class LedgerTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        ledger.DATA = root / "data"
        ledger.TMP = ledger.DATA / "tmp"
        ledger.DB_PATH = ledger.DATA / "ledger.db"
        self.conn = ledger.connect()

    def tearDown(self):
        self.conn.close()
        self.tmp.cleanup()

    def import_bytes(self, filename, data, kind, month="", include_costs=True, allow_overlap=False):
        token, _path = ledger.save_upload(filename, data)
        return ledger.commit_import(
            self.conn, token, kind, month, include_costs, None, "", allow_overlap
        )

    def test_orders_split_confirmed_unconfirmed_and_refund(self):
        data = make_xlsx([("export", [
            ORDER_HEADERS,
            ["A1", "夹爪", 10, 2, "颜色:红", "", 20, 20, "没有申请退款", "无退款申请", "2026-10-02 10:00:00", "1", "2026-10-20 10:00:00"],
            ["A2", "夹爪", 10, 1, "颜色:红", "", 10, 10, "没有申请退款", "无退款申请", "2026-10-03 10:00:00", "1", ""],
            ["A3", "夹爪", 10, 4, "颜色:蓝", "", 40, 40, "没有申请退款", "无退款申请", "2026-09-01 10:00:00", "1", "2026-10-08 10:00:00"],
            ["A4", "夹爪", 10, 3, "颜色:红", "", 30, 30, "退款成功", 30, "2026-10-04 10:00:00", "1", "2026-10-09 10:00:00"],
            ["A5", "夹爪", 10, 1, "颜色:红", "", 15, 15, "买家已经申请退款，等待卖家同意", 0, "2026-10-05 10:00:00", "1", "2026-10-11 10:00:00"],
        ])])
        result = self.import_bytes("收益.xlsx", data, "orders")
        self.assertEqual(result["inserted"], 5)
        state = ledger.build_state(self.conn, "2026-10", "confirmed")
        product = state["products"][0]
        self.assertEqual(product["title"], "夹爪")
        # 红确认 2 件 20 元 + 退款进行中 1 件 15 元；蓝是 9 月付款、10 月确认 4 件 40 元
        self.assertEqual(product["confirmed_qty"], 7)
        self.assertEqual(product["confirmed_amt"], 75)
        self.assertEqual(product["unconfirmed_qty"], 1)
        self.assertEqual(product["unconfirmed_amt"], 10)
        self.assertEqual(product["refund_amt"], 30)
        self.assertEqual(product["pending_count"], 1)
        red = next(attr for attr in product["attrs"] if attr["attr"] == "颜色:红")
        blue = next(attr for attr in product["attrs"] if attr["attr"] == "颜色:蓝")
        self.assertEqual(red["confirmed_qty"], 3)
        self.assertEqual(red["confirmed_amt"], 35)
        self.assertEqual(red["unconfirmed_amt"], 10)
        self.assertEqual(blue["confirmed_amt"], 40)
        self.assertIsNone(product["profit"])
        self.assertIn("还没有支出", state["summary"]["sentence"])

    def test_reimport_updates_confirmation_without_doubling(self):
        first = make_xlsx([("export", [
            ORDER_HEADERS,
            ["A1", "夹爪", 10, 2, "颜色:红", "", 20, 20, "没有申请退款", "无退款申请", "2026-10-02 10:00:00", "1", ""],
        ])])
        self.import_bytes("a.xlsx", first, "orders")
        second = make_xlsx([("export", [
            ORDER_HEADERS,
            ["A1", "夹爪", 10, 2, "颜色:红", "", 20, 20, "没有申请退款", "无退款申请", "2026-10-02 10:00:00", "1", "2026-10-21 12:00:00"],
        ])])
        result = self.import_bytes("b.xlsx", second, "orders")
        self.assertEqual(result["inserted"], 0)
        self.assertEqual(result["updated"], 1)
        state = ledger.build_state(self.conn, "2026-10")
        product = state["products"][0]
        self.assertEqual(product["confirmed_amt"], 20)
        self.assertEqual(product["unconfirmed_amt"], 0)
        self.assertEqual(state["summary"]["confirmed_amt"], 20)

    def test_aggregated_costs_forward_fill_and_profit(self):
        data = make_xlsx([
            ("说明", [["项目", "说明"], ["赛道", "新能源赛道"]]),
            ("按商品属性", [
                ["商品标题", "商品属性", "成本", "销售数量", "买家实付金额", "占该标题金额", "净利润", "标题销售数量合计", "标题实付合计", "成本", "明细", "利润"],
                ["示例套件", "颜色:甲", None, 10, 1000, 0.8, None, 12, 1200, 700, "板子", -100],
                [None, "颜色:乙", None, 2, 200, 0.2, None, None, None, 400, "卡片", None],
                ["合计", "2 个属性", None, 12, 1200, 1, None, 12, 1200, None, None, None],
            ]),
        ])
        result = self.import_bytes("合并.xlsx", data, "aggregated", "2026-10", True)
        self.assertEqual(result["inserted"], 2)
        self.assertEqual(result["costs"], 2)
        state = ledger.build_state(self.conn, "2026-10")
        product = state["products"][0]
        self.assertEqual(product["track"], "新能源赛道")
        self.assertEqual(product["confirmed_qty"], 12)
        self.assertEqual(product["confirmed_amt"], 1200)
        self.assertEqual(product["unconfirmed_amt"], 0)
        self.assertEqual(product["cost"], 1100)
        self.assertEqual(product["profit"], 100)
        self.assertEqual(product["status"], "profit")
        self.assertEqual(state["summary"]["month_profit"], 100)
        self.assertEqual(state["summary"]["unallocated_total"], 0)
        # 再导一次不会加倍
        again = make_xlsx([
            ("按商品属性", [
                ["商品标题", "商品属性", "销售数量", "买家实付金额", "成本", "明细"],
                ["示例套件", "颜色:甲", 10, 1000, 700, "板子"],
                [None, "颜色:乙", 2, 200, 400, "卡片"],
            ]),
        ])
        result = self.import_bytes("合并2.xlsx", again, "aggregated", "2026-10", True)
        self.assertEqual(result["inserted"], 0)
        state = ledger.build_state(self.conn, "2026-10")
        self.assertEqual(state["summary"]["confirmed_amt"], 1200)
        self.assertEqual(state["summary"]["expense_total"], 1100)

    def test_costs_only_does_not_double_revenue(self):
        self.import_bytes("收益.xlsx", make_xlsx([("export", [
            ORDER_HEADERS,
            ["A1", "示例套件", 1, 10, "颜色:甲", "", 1000, 1000, "没有申请退款", "", "2026-10-01", "1", "2026-10-08"],
        ])]), "orders")
        merged = make_xlsx([("按商品属性", [
            ["商品标题", "商品属性", "销售数量", "买家实付金额", "成本", "明细"],
            ["示例套件", "颜色:甲", 10, 1000, 700, "板子"],
        ])])
        token, _path = ledger.save_upload("合并.xlsx", merged)
        with self.assertRaises(ValueError):
            ledger.commit_import(self.conn, token, "aggregated", "2026-10", True, None, "", False, False, False)
        token, _path = ledger.save_upload("合并.xlsx", merged)
        result = ledger.commit_import(self.conn, token, "aggregated", "2026-10", True, None, "", False, False, True)
        self.assertEqual(result["inserted"], 0)
        self.assertEqual(result["costs"], 1)
        state = ledger.build_state(self.conn, "2026-10")
        product = state["products"][0]
        self.assertEqual(product["confirmed_amt"], 1000)
        self.assertEqual(product["cost"], 700)
        self.assertEqual(product["profit"], 300)

    def test_drag_split_move_and_month_profit_uses_all_expenses(self):
        self.import_bytes("收益.xlsx", make_xlsx([("export", [
            ORDER_HEADERS,
            ["A1", "甲", 1, 1, "红", "", 800, 800, "没有申请退款", "无退款申请", "2026-10-01", "1", "2026-10-02"],
            ["A2", "乙", 1, 1, "红", "", 200, 200, "没有申请退款", "无退款申请", "2026-10-01", "2", "2026-10-02"],
        ])]), "orders")
        expense_id = ledger.add_expense(self.conn, "木板", 500, "2026-10")
        ledger.allocate(self.conn, expense_id, "甲", 300)
        with self.assertRaises(ValueError):
            ledger.allocate(self.conn, expense_id, "乙", 300)
        ledger.allocate(self.conn, expense_id, "乙", 200)
        state = ledger.build_state(self.conn, "2026-10")
        by_title = {item["title"]: item for item in state["products"]}
        self.assertEqual(by_title["甲"]["profit"], 500)
        self.assertEqual(by_title["乙"]["profit"], 0)
        self.assertEqual(by_title["乙"]["status"], "even")
        self.assertEqual(state["summary"]["month_profit"], 500)
        self.assertEqual(state["summary"]["unallocated_total"], 0)
        alloc = by_title["甲"]["allocations"][0]
        ledger.update_allocation(self.conn, alloc["id"], amount=100)
        state = ledger.build_state(self.conn, "2026-10")
        self.assertEqual(state["summary"]["unallocated_total"], 200)
        # 整月盈亏仍把没挂出去的成本算进去：1000 - 500 = 500
        self.assertEqual(state["summary"]["month_profit"], 500)
        self.assertEqual(state["expenses"][0]["remaining"], 200)
        ledger.update_allocation(self.conn, alloc["id"], title="乙")
        state = ledger.build_state(self.conn, "2026-10")
        by_title = {item["title"]: item for item in state["products"]}
        self.assertEqual(by_title["乙"]["cost"], 300)
        self.assertEqual(by_title["甲"]["status"], "unknown")

    def test_expense_links_by_product_code_and_no_cost_flag(self):
        self.import_bytes("收益.xlsx", make_xlsx([("export", [
            ORDER_HEADERS,
            ["A1", "YDI-D14工创套件", 1, 2, "颜色:甲", "", 100, 100, "没有申请退款", "", "2026-10-01", "9", "2026-10-03"],
        ])]), "orders")
        table = make_xlsx([("支出", [
            ["日期", "项目", "金额", "分类", "备注", "对应商品"],
            ["2026-10-06", "板子", 40, "加工", "", "YDI-D14"],
            ["2026-10-07", "胶带", 5, "物料", "", ""],
            ["2026-10-08", "YDI-D14 补发运费", 8, "物流", "", ""],
        ])])
        info_token, path = ledger.save_upload("支出.xlsx", table)
        info = ledger.analyze_file(path)
        self.assertEqual(info["kind"], "expenses")
        result = ledger.commit_import(self.conn, info_token, "expenses", "2026-10", True, None, "", False)
        self.assertEqual(result["linked"], 1)
        state = ledger.build_state(self.conn, "2026-10")
        product = state["products"][0]
        self.assertEqual(product["cost"], 40)
        self.assertEqual(product["profit"], 60)
        self.assertEqual(state["summary"]["unallocated_total"], 13)
        self.assertEqual(state["summary"]["month_profit"], 47)
        tape = next(item for item in state["expenses"] if item["name"] == "胶带")
        self.assertEqual(tape["suggestion"], "")
        board = next(item for item in state["expenses"] if item["name"] == "板子")
        self.assertEqual(board["remaining"], 0)
        freight = next(item for item in state["expenses"] if "补发" in item["name"])
        self.assertEqual(freight["suggestion"], "YDI-D14工创套件")
        self.assertEqual(freight["remaining"], 8)

    def test_real_october_files_match_known_totals(self):
        orders = Path("/home/ubuntu/.cursor/projects/workspace/uploads/____8_9_10__14d7.xlsx")
        merged = Path("/home/ubuntu/.cursor/projects/workspace/uploads/10________-______a88f.xlsx")
        if not orders.exists() or not merged.exists():
            self.skipTest("样例表不在这台机器上")
        token, path = ledger.save_upload(orders.name, orders.read_bytes())
        info = ledger.analyze_file(path)
        self.assertEqual(info["kind"], "orders")
        self.assertGreater(info["stats"]["rows"], 9000)
        self.import_bytes(orders.name, orders.read_bytes(), "orders")
        state = ledger.build_state(self.conn, "2026-10")
        d14 = next(item for item in state["products"] if item["title"].startswith("YDI-D14"))
        self.assertEqual(d14["confirmed_qty"], 279)
        self.assertEqual(d14["confirmed_amt"], 84321.19)
        confirmed_attrs = [item for item in d14["attrs"] if item["confirmed_amt"] or item["confirmed_qty"]]
        self.assertEqual(len(confirmed_attrs), 17)
        self.assertGreater(state["summary"]["unconfirmed_amt"], 0)
        self.assertGreater(state["summary"]["refund_amt"], 0)
        # 同一份再导入只更新，不加倍
        again = self.import_bytes(orders.name, orders.read_bytes(), "orders")
        self.assertEqual(again["inserted"], 0)
        self.assertGreater(again["updated"], 0)
        state = ledger.build_state(self.conn, "2026-10")
        d14 = next(item for item in state["products"] if item["title"].startswith("YDI-D14"))
        self.assertEqual(d14["confirmed_amt"], 84321.19)

        # 合并表单独验证成本和利润，换一个库，避免和订单明细加在一起
        self.conn.close()
        shutil.rmtree(ledger.DATA)
        self.conn = ledger.connect()
        merged_info_token, merged_path = ledger.save_upload(merged.name, merged.read_bytes())
        merged_info = ledger.analyze_file(merged_path)
        self.assertEqual(merged_info["kind"], "aggregated")
        self.assertEqual(merged_info["stats"]["track"], "新能源赛道")
        self.assertGreater(merged_info["stats"]["cost_rows"], 0)
        self.import_bytes(merged.name, merged.read_bytes(), "aggregated", "2026-10", True)
        state = ledger.build_state(self.conn, "2026-10")
        self.assertEqual(state["summary"]["confirmed_amt"], 282891.63)
        self.assertEqual(state["summary"]["confirmed_qty"], 6207)
        d14 = next(item for item in state["products"] if item["title"].startswith("YDI-D14"))
        self.assertEqual(d14["confirmed_amt"], 84321.19)
        self.assertEqual(d14["cost"], 90349.86)
        self.assertEqual(d14["profit"], -6028.67)
        self.assertEqual(d14["status"], "loss")
        d12 = next(item for item in state["products"] if item["title"].startswith("YDI-D12"))
        self.assertEqual(d12["profit"], 14703.25)
        lamp = next(item for item in state["products"] if item["title"].startswith("YDI-D69"))
        self.assertEqual(lamp["profit"], -700)
        self.assertEqual(state["summary"]["unallocated_total"], 0)
        book = ledger.export_workbook(self.conn, "2026-10")
        self.assertTrue(book[:2] == b"PK")

    def test_backup_roundtrip(self):
        ledger.load_demo(self.conn)
        payload = ledger.backup_data(self.conn)
        ledger.delete_import(self.conn, payload["imports"][0]["id"])
        self.assertEqual(ledger.build_state(self.conn, "2099-01")["summary"]["product_count"], 0)
        ledger.restore_data(self.conn, payload)
        state = ledger.build_state(self.conn, "2099-01")
        self.assertEqual(state["summary"]["confirmed_amt"], 1500)
        self.assertEqual(state["summary"]["month_profit"], 1080)
        self.assertEqual(state["summary"]["unconfirmed_amt"], 400)


if __name__ == "__main__":
    unittest.main()
