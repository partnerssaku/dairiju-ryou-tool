import openpyxl
import glob
import os

files = glob.glob(r"C:\Users\崎久保秀一\Desktop\ClaudeWork\代理受領フォルダ\代理受領通知書_一括出力_*.xlsx")
files.sort(key=os.path.getctime)
filepath = files[-1]
print(f"Reading from: {os.path.basename(filepath)}")

try:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.worksheets[0]
    print(f"Sheet name: {sheet.title}")
    print(f"H4 (発行日): {sheet['H4'].value}")
    print(f"H15 (事業者名): {sheet['H15'].value}")
    print(f"H16 (事業所名): {sheet['H16'].value}")
    print(f"H18 (代表者): {sheet['H18'].value}")
    print(f"E29 (受給日): {sheet['E29'].value}")
    print(f"C25 (月): {sheet['C25'].value}")
    print(f"D25 (text): {sheet['D25'].value}")
    print(f"F25 (代理受領額): {sheet['F25'].value}")
except Exception as e:
    print(f"Error reading excel: {e}")
