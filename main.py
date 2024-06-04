import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# 确保目标文件夹存在
Path('PDFs').mkdir(parents=True, exist_ok=True)

filepaths = glob.glob('bills/*.xlsx')
# print(filepaths)
# ['bills/10001-2023.1.18.xlsx', 'bills/10002-2023.1.18.xlsx',
# 'bills/10003-2023.1.18.xlsx']

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    # print(df)  # 读取几个表的内容
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    # print('filename',filename) # filename 10003-2023.1.18
    bill_num = filename.split('-')[0]
    # print('data_num', data_num) # data_num 10003
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'bill_num. {bill_num}')
    pdf.output(f"PDFs/{filename}.pdf")
