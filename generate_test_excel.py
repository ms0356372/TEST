"""建立需求文件指定的 test_input.xlsx 測試活頁簿。

Excel 檔是 ZIP 格式的二進位檔，部分 PR 系統不接受直接附加；因此專案
保存可重現的產生器，而不將產物提交至 Git。
"""
from pathlib import Path

from openpyxl import Workbook


def create_test_workbook(destination: Path = Path("test_input.xlsx")) -> Path:
    """建立包含來源 E 欄及待驗證 F 欄資料的固定測試檔。"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "資料"
    worksheet["E3"] = "特殊作業"
    values = ["99;03;02", "02;99", "03", "99 ; 03 ; 02", "99;99;02", "99;88;02"]
    for row, value in enumerate(values, start=4):
        worksheet.cell(row=row, column=5, value=value)

    # 原 F 欄資料用來確認 insert_cols() 後會完整右移至 G 欄。
    worksheet["F3"] = "部門"
    worksheet["F4"] = "A部門"
    workbook.save(destination)
    return destination


if __name__ == "__main__":
    output = create_test_workbook()
    print(f"已建立：{output.resolve()}")
