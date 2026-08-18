from pathlib import Path
from typing import Any,Iterable,Sequence
def write_analysis_workbook(path:str|Path,headers:Sequence[str],rows:Iterable[Sequence[Any]])->Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    base_font=Font(name="標楷體",size=10); center=Alignment(horizontal="center",vertical="center")
    header_fill=PatternFill("solid",fgColor="F2F2F2")
    risk_fills={"中度":PatternFill("solid",fgColor="D6DCE4"),"高度":PatternFill("solid",fgColor="FFFF00"),"極高":PatternFill("solid",fgColor="FF0000")}
    higher_fill=PatternFill("solid",fgColor="FFFF00")
    target=Path(path); wb=Workbook(); ws=wb.active; ws.title="分析結果"; ws.append(list(headers))
    for row in rows:ws.append(list(row))
    for row in ws.iter_rows():
        for cell in row:cell.font=base_font;cell.alignment=center
    for cell in ws[1]:cell.fill=header_fill
    for row in range(2,ws.max_row+1):
        if ws.cell(row,22).value in risk_fills:ws.cell(row,22).fill=risk_fills[ws.cell(row,22).value]
        if ws.cell(row,23).value=="較高":ws.cell(row,23).fill=higher_fill
    for col in range(1,ws.max_column+1):
        width=max((len(str(ws.cell(row,col).value or "")) for row in range(1,ws.max_row+1)),default=8)
        ws.column_dimensions[get_column_letter(col)].width=min(max(width+2,10),40)
    ws.freeze_panes="A2"; target.parent.mkdir(parents=True,exist_ok=True); wb.save(target); return target
