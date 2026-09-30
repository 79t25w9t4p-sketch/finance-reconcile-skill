# coding: utf-8
"""Read one XLSX sheet without modifying it; emit CSV to stdout."""
import sys, csv, io, zipfile, json
from pathlib import Path
try:
    import openpyxl
except ImportError:
    sys.exit("XLSX requires openpyxl. Use an existing Python environment with it, or export CSV.")
try:
    path=Path(sys.argv[1])
    if path.stat().st_size > 10*1024*1024: raise ValueError("XLSX exceeds 10 MB")
    with zipfile.ZipFile(path) as z:
        if sum(x.file_size for x in z.infolist()) > 40*1024*1024: raise ValueError("XLSX expanded size exceeds 40 MB")
    wb=openpyxl.load_workbook(path,read_only=True,data_only=False,keep_links=False)
    sheet=sys.argv[2] if len(sys.argv)>2 else ''
    if not sheet and len(wb.sheetnames)!=1:
        raise ValueError("Choose --left-sheet/--right-sheet; available sheets: " + ', '.join(wb.sheetnames))
    ws=wb[sheet or wb.sheetnames[0]]
    header_row=int(sys.argv[3]) if len(sys.argv)>3 else 1
    if header_row<1 or header_row>ws.max_row: raise ValueError('Header row outside sheet')
    if ws.max_row-header_row>5000 or ws.max_column>50: raise ValueError("At most 5000 data rows and 50 columns")
    out=io.StringIO();writer=csv.writer(out);source_lines=[]
    for row_number,row in enumerate(ws.iter_rows(min_row=header_row),header_row):
        values=[]
        for c in row:
            if c.data_type=='f': raise ValueError("Formula at " + c.coordinate + ": recalculate and export values to CSV first")
            if c.data_type=='e': raise ValueError("Excel error at " + c.coordinate)
            v=c.value
            if isinstance(v,(int,float)) and ('000' in c.number_format or abs(v)>=10**15):
                raise ValueError("Ambiguous numeric identifier/large number at " + c.coordinate + ": provide text identifiers and exact decimal amounts")
            if isinstance(v,float): v=str(v)
            values.append('' if v is None else str(v))
        writer.writerow(values)
        if row_number>header_row and any(x.strip() for x in values): source_lines.append(row_number)
    sheet_name=ws.title
    wb.close();sys.stdout.write(json.dumps({'csv':out.getvalue(),'sourceLines':source_lines,'sheet':sheet_name},ensure_ascii=False))
except Exception as e:
    sys.exit(str(e))
