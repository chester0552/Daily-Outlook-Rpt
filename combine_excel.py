from pathlib import Path 


OUTLOOK_RPT = "Output"
COMBINED_RPT = "Report\Combined"

excel_files = list(Path(OUTLOOK_RPT).glob('*xlsx'))
combined_wb = xw.Book()



