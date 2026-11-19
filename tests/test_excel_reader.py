from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from permission_reminder.core.excel_reader import ExcelFormatError, read_clients


HEADERS = [
    "客户号", "客户姓名", "专业投资者到期日期", "私募基金到期日期", "资管权限到期日期",
]


class ExcelReaderTests(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def write_workbook(self, rows, filename="clients.xlsx"):
        path = self.root / filename
        workbook = Workbook()
        for row in rows:
            workbook.active.append(row)
        workbook.save(path)
        workbook.close()
        return path

    def test_finds_headers_on_first_second_and_later_rows(self):
        for title_count in (0, 1, 4):
            with self.subTest(title_count=title_count):
                rows = [["客户权限清单"] for _ in range(title_count)]
                rows += [HEADERS, ["A001", "张三", 20261021, None, None]]
                records = read_clients(self.write_workbook(rows))
                self.assertEqual(len(records), 1)
                self.assertEqual(records[0].customer_id, "A001")
                self.assertEqual(records[0].expiry_dates["专业投资者"], date(2026, 10, 21))

    def test_accepts_customer_aliases_and_whitespace_in_headers(self):
        for name_alias in ("客户名称", "客户名"):
            with self.subTest(name_alias=name_alias):
                headers = [" 客 户 编号\n", name_alias, "专业投资者\n到期日期",
                           " 私募基金到期日期 ", "资管权限\t到期日期"]
                records = read_clients(self.write_workbook([
                    headers, ["00123", "张三", 20261021, None, None],
                ]))
                self.assertEqual(records[0].customer_id, "00123")
                self.assertEqual(records[0].customer_name, "张三")

    def test_accepts_xlsm_without_changing_date_parsing(self):
        path = self.write_workbook([
            HEADERS, ["A001", "张三", 46316, datetime(2026, 11, 2), "2026年12月3日"],
        ], "clients.xlsm")
        record = read_clients(path)[0]
        self.assertEqual(record.expiry_dates, {
            "专业投资者": date(2026, 10, 21),
            "私募基金": date(2026, 11, 2),
            "资管权限": date(2026, 12, 3),
        })

    def test_blank_dates_and_invalid_dates_still_skip_only_that_permission(self):
        records = read_clients(self.write_workbook([
            HEADERS,
            ["A001", "张三", None, "不是日期", 20261021],
            [None, None, None, None, None],
            [None, "缺客户号", 20261021, None, None],
        ]))
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].expiry_dates, {
            "专业投资者": None, "私募基金": None, "资管权限": date(2026, 10, 21),
        })

    def test_missing_column_has_explicit_error(self):
        path = self.write_workbook([
            ["客户权限清单"], HEADERS[:-1], ["A001", "张三", 20261021, None],
        ])
        with self.assertRaisesRegex(ExcelFormatError, "缺少关键列：资管权限到期日期"):
            read_clients(path)

    def test_duplicate_columns_after_normalization_are_rejected(self):
        for duplicate in ("客户名称", "客户 姓名\n", "专业投资者到期日期"):
            with self.subTest(duplicate=duplicate):
                path = self.write_workbook([
                    HEADERS + [duplicate], ["A001", "张三", 20261021, None, None, "重复"],
                ])
                with self.assertRaisesRegex(ExcelFormatError, "表头包含重复列"):
                    read_clients(path)

    def test_empty_unnamed_columns_do_not_create_duplicate_header_errors(self):
        records = read_clients(self.write_workbook([
            [None, *HEADERS, None], [None, "A001", "张三", 20261021, None, None, None],
        ]))
        self.assertEqual(len(records), 1)

    def test_duplicate_extra_note_columns_do_not_block_business_columns(self):
        records = read_clients(self.write_workbook([
            HEADERS + ["备注", "备注"],
            ["A001", "张三", 20261021, None, None, "业务备注", "审核备注"],
        ]))
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].customer_id, "A001")
        self.assertEqual(records[0].expiry_dates["专业投资者"], date(2026, 10, 21))

    def test_unrecognizable_and_empty_workbooks_have_header_error(self):
        for rows in ([], [["未知台账"], ["一", "二", "三"]]):
            with self.subTest(rows=rows):
                with self.assertRaisesRegex(ExcelFormatError, "未找到.*表头行"):
                    read_clients(self.write_workbook(rows))

    def test_missing_unsupported_and_corrupt_files_have_business_errors(self):
        with self.assertRaisesRegex(ExcelFormatError, "文件不存在"):
            read_clients(self.root / "missing.xlsx")
        for filename, message in (("clients.csv", "不是有效的 Excel 文件"),
                                  ("broken.xlsx", "Excel 读取失败")):
            with self.subTest(filename=filename):
                path = self.root / filename
                path.write_text("invalid workbook", encoding="utf-8")
                with self.assertRaisesRegex(ExcelFormatError, message):
                    read_clients(path)


if __name__ == "__main__":
    unittest.main()
