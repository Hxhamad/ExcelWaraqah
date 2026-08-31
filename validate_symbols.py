# One-off data prep (not project code): validate 212 Tadawul codes live via yfinance,
# keep those with price history, fill EN name + EN sector. Output: v2/symbols.csv
import csv, json, os, sys, time

sys.path.insert(0, 'C:/Users/Hamad/portfolio_excel/v2')
import yfinance as yf

SRC = 'C:/Users/Hamad/portfolio_excel/v2/symbols_full.csv'
OUT = 'C:/Users/Hamad/portfolio_excel/v2/symbols.csv'
STATE = 'C:/Users/Hamad/portfolio_excel/v2/data/validate_state.json'

SECTOR_MAP = {
    'الطاقة': 'Energy', 'المواد الأساسية': 'Materials', 'التأمين': 'Insurance',
    'الصناديق العقارية المتداولة': 'REIT', 'السلع الرأسمالية': 'Capital Goods',
    'إنتاج الأغذية': 'Food Producers', 'البنوك': 'Banks',
    'الخدمات الإستهلاكية': 'Consumer Services', 'إدارة وتطوير العقارات': 'Real Estate Development',
    'تجزئة السلع الكمالية': 'Consumer Durables Retail', 'الرعاية الصحية': 'Health Care',
    'النقل': 'Transportation', 'السلع طويلة الاجل': 'Consumer Durables',
    'الإستثمار والتمويل': 'Investment & Finance', 'تجزئة الأغذية': 'Food Retail',
    'الاتصالات': 'Telecommunications', 'الخدمات اللوجستية': 'Logistics',
    'السيارات والمكونات': 'Auto & Components', 'البنية التحتية': 'Infrastructure',
    'الأغذية المتنوعة': 'Diversified Food', 'الأدوية': 'Pharmaceuticals',
}

rows = list(csv.DictReader(open(SRC, encoding='utf-8-sig')))
done = {}
if os.path.exists(STATE):
    done = json.load(open(STATE, encoding='utf-8'))

out_rows, failed = [], []
for i, r in enumerate(rows, 1):
    code = r['code']
    if code in done:
        res = done[code]
    else:
        try:
            t = yf.Ticker(f'{code}.SR')
            h = t.history(period='5d')
            ok = h is not None and len(h) > 0
            name_en, sec_en = '', ''
            if ok:
                try:
                    info = t.info or {}
                    name_en = (info.get('shortName') or info.get('longName') or '').strip()
                    sec_en = (info.get('sector') or '').strip()
                except Exception:
                    pass
            res = {'ok': ok, 'name_en': name_en, 'sec_en': sec_en}
        except Exception as e:
            res = {'ok': False, 'err': str(e)[:80]}
        done[code] = res
        json.dump(done, open(STATE, 'w', encoding='utf-8'), ensure_ascii=False)
        time.sleep(1.0)
    if res.get('ok'):
        out_rows.append({
            'code': code,
            'name_ar': r['name_ar'],
            'name_en': res.get('name_en', ''),
            'sector': res.get('sec_en') or SECTOR_MAP.get(r['sector_ar'], 'Other'),
            'sector_ar': r['sector_ar'],
        })
    else:
        failed.append(code)
    if i % 25 == 0:
        print(f'{i}/{len(rows)} checked, ok={len(out_rows)}, fail={len(failed)}', flush=True)

with open(OUT, 'w', newline='', encoding='utf-8-sig') as f:
    wr = csv.DictWriter(f, fieldnames=['code', 'name_ar', 'name_en', 'sector', 'sector_ar'])
    wr.writeheader(); wr.writerows(out_rows)

print(f'DONE: validated={len(out_rows)} failed={len(failed)}')
print('failed codes:', failed)
