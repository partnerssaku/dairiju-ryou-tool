import csv
import json

filename = r"C:\Users\崎久保秀一\Desktop\ClaudeWork\代理受領フォルダ\TH01_202602_2323300125_002_20260224_23_202602241534_2.CSV"

users = {}

with open(filename, 'r', encoding='shift_jis', errors='replace') as f:
    reader = csv.reader(f)
    for row in reader:
        if len(row) < 3: continue
        
        record_id = row[2]
        if record_id == 'J131':
            # 基本情報レコード
            user_id = row[7]
            kana_name = row[9]
            target_month = row[4] # YYYYMM
            users[user_id] = {
                'user_id': user_id,
                'name': kana_name,
                'month': target_month,
                'service_cost': 0, # 総費用額
                'user_burden': 0,  # 利用者負担額
                'proxy_amount': 0  # 代理受領額（給付費）
            }
        elif record_id == 'J141':
            # 請求情報レコード (if available)
            user_id = row[7]
            if user_id in users:
                # Based on standard UKE J141
                try:
                    users[user_id]['service_cost'] = int(row[11]) if len(row) > 11 else 0
                    users[user_id]['user_burden'] = int(row[13]) if len(row) > 13 else 0
                    users[user_id]['proxy_amount'] = int(row[12]) if len(row) > 12 else 0 # 給付費
                except:
                    pass

with open('extracted_data.txt', 'w', encoding='utf-8') as out:
    for u in users.values():
        out.write(json.dumps(u, ensure_ascii=False) + '\n')
