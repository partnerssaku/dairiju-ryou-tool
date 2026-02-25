import csv

filename = r"C:\Users\崎久保秀一\Desktop\ClaudeWork\代理受領フォルダ\TH01_202602_2323300125_002_20260224_23_202602241534_2.CSV"
outname = r"C:\Users\崎久保秀一\Desktop\ClaudeWork\代理受領フォルダ\debug_csv_out.txt"

with open(filename, 'r', encoding='shift_jis', errors='replace') as f:
    with open(outname, 'w', encoding='utf-8') as out:
        reader = csv.reader(f)
        for row in reader:
            if len(row) > 7 and row[7] == '2000006458': # 岡田英樹
                out.write(f"Record {row[2]}: {row}\n")
