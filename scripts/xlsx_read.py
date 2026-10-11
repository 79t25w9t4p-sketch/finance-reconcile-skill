# coding: utf-8
"""Read one XLSX sheet without modifying it; emit CSV to stdout."""
import sys, csv, io, zipfile, json
from pathlib import Path
try:
    import openpyxl
except ImportError:
    sys.exit("XLSX requires openpyxl. Use an existing Python environment with it, or export CSV.")
def open_checked(path):
    path=Path(path)
    if path.stat().st_size > 10*1024*1024: raise ValueError("XLSX exceeds 10 MB")
    with zipfile.ZipFile(path) as z:
        if sum(x.file_size for x in z.infolist()) > 40*1024*1024: raise ValueError("XLSX expanded size exceeds 40 MB")
    return openpyxl.load_workbook(path,read_only=True,data_only=False,keep_links=False)


def inspect_layout(wb, options):
    """Expose candidate headers only, never the rows examined to find them."""
    sheet=options.get('sheet')
    names=[sheet] if sheet else wb.sheetnames
    if len(names)>20: raise ValueError('超过20个工作表，请先用 --left-sheet/--right-sheet 指定明细页')
    header=options.get('headerRow')
    if header and (not str(header).isascii() or not str(header).isdigit() or int(header)<1):
        raise ValueError('表头行必须是正整数')
    reference=set(options['reference']);amount=set(options['amount'])
    candidates=[];sheets=[];output_bytes=0
    for name in names:
        ws=wb[name]
        if ws.max_column>50: raise ValueError('工作表超过50列，请指定符合限制的明细页：'+name)
        start=int(header) if header else 1
        end=int(header) if header else min(20,ws.max_row)
        sheets.append({'sheet':name,'state':ws.sheet_state})
        if start>ws.max_row:
            if sheet: raise ValueError('Header row outside sheet: '+name)
            continue
        for number,row in enumerate(ws.iter_rows(min_row=start,max_row=end),start):
            # Formula/error cells cannot be treated as a recognized header.
            if any(c.data_type in {'f','e'} for c in row): continue
            headers=['' if c.value is None else str(c.value).strip() for c in row]
            while headers and not headers[-1]: headers.pop()
            normalized={value.lower() for value in headers}
            if normalized.intersection(reference) and normalized.intersection(amount):
                candidate={'sheet':name,'state':ws.sheet_state,'headerRow':number,'headers':headers}
                output_bytes+=len(json.dumps(candidate,ensure_ascii=False).encode('utf-8'))
                if output_bytes>2*1024*1024:
                    raise ValueError('候选表头超过2MB，请直接指定明细页和表头行')
                candidates.append(candidate)
    return {'candidates':candidates,'sheets':sheets,'scanRows':20 if not header else None}


def read_sheet(wb, sheet, header_row):
    if not sheet and len(wb.sheetnames)!=1:
        raise ValueError("Choose --left-sheet/--right-sheet; available sheets: " + ', '.join(wb.sheetnames))
    ws=wb[sheet or wb.sheetnames[0]]
    if header_row<1 or header_row>ws.max_row: raise ValueError('Header row outside sheet')
    if ws.max_row-header_row>5000 or ws.max_column>50: raise ValueError("At most 5000 data rows and 50 columns")
    out=io.StringIO();writer=csv.writer(out);source_lines=[];width=None
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
        if width is None:
            width=len(values)
            while width and not values[width-1].strip(): width-=1
        elif any(value.strip() for value in values[width:]):
            raise ValueError('表头之外存在数据：第'+str(row_number)+'行；请补齐列名，不能直接丢弃该列')
        values=values[:width]
        writer.writerow(values)
        if row_number>header_row and any(x.strip() for x in values): source_lines.append(row_number)
    return {'csv':out.getvalue(),'sourceLines':source_lines,'sheet':ws.title}


def main():
    layout=len(sys.argv)>1 and sys.argv[1]=='--layout'
    wb=open_checked(sys.argv[2] if layout else sys.argv[1])
    try:
        if layout:
            result=inspect_layout(wb,json.loads(sys.argv[3]))
        else:
            result=read_sheet(wb,sys.argv[2] if len(sys.argv)>2 else '',int(sys.argv[3]) if len(sys.argv)>3 else 1)
        sys.stdout.write(json.dumps(result,ensure_ascii=False))
    finally:
        wb.close()


if __name__=='__main__':
    try: main()
    except Exception as e: sys.exit(str(e))
