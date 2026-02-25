import csv
filename = r"C:\Users\崎久保秀一\Desktop\ClaudeWork\代理受領フォルダ\TH01_202602_2323300125_002_20260224_23_202602241534_2.CSV"

with open(filename, 'r', encoding='shift_jis', errors='replace') as f:
    reader = csv.reader(f)
    count = 0
    for row in reader:
        if len(row) > 2 and row[2].startswith('J'):
            print(f"Record {row[2]}: {row}")
            count += 1
            if count > 20: break
