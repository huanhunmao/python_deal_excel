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
    # print(df)  # 读取几个表的内容
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    # print('filename',filename) # filename 10003-2023.1.18
    bill_num, date = filename.split('-')
    # print('data_num', data_num) # data_num 10003
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Bill_num. {bill_num}', ln=1)

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Date. {date}', ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # Add the header
    columns = list(df.columns)
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows in table
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
