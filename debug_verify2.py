import openpyxl

filepath = r"C:\Users\崎久保秀一\Desktop\ClaudeWork\代理受領フォルダ\代理受領通知書_一括出力_20260224_181136.xlsx"

try:
    with open('debug_verify_out3.txt', 'w', encoding='utf-8') as f:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.worksheets[0]
        f.write(f"Sheet name: {sheet.title}\n")
        f.write(f"D7: {repr(sheet['D7'].value)}\n")
        f.write(f"D8: {repr(sheet['D8'].value)}\n")
        f.write(f"C25: {repr(sheet['C25'].value)}\n")
        f.write(f"D25: {repr(sheet['D25'].value)}\n")
        f.write(f"E25: {repr(sheet['E25'].value)}\n")  # Maybe service type goes here?
        f.write(f"E26: {repr(sheet['E26'].value)}\n")
        f.write(f"E27: {repr(sheet['E27'].value)}\n")
        f.write(f"E29: {repr(sheet['E29'].value)}\n")
        f.write(f"F25: {repr(sheet['F25'].value)}\n")
        f.write(f"H4: {repr(sheet['H4'].value)}\n")
        f.write(f"H15: {repr(sheet['H15'].value)}\n")
        f.write(f"H16: {repr(sheet['H16'].value)}\n")
        f.write(f"H17: {repr(sheet['H17'].value)}\n")
        f.write(f"H18: {repr(sheet['H18'].value)}\n")
        f.write(f"H19: {repr(sheet['H19'].value)}\n")
        
        for r in range(30, 36):
            f.write(f"H{r}: {repr(sheet[f'H{r}'].value)}\n")
except Exception as e:
    print(f"Error: {e}")
