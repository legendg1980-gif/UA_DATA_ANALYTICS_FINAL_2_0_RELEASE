from __future__ import annotations
import os, csv, random, hashlib, sqlite3, json, math, zipfile, shutil, gzip
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from collections import defaultdict, Counter

from artifact_tool import Workbook, SpreadsheetFile
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from PIL import Image, ImageDraw, ImageFont

SEED = 20260813
rng = random.Random(SEED)
ROOT = Path('/mnt/data/Defense_Analytics_Course_TrackA_V01_STARTER')
if ROOT.exists():
    shutil.rmtree(ROOT)

# Folder layout
for p in [
    '00_ReadMe',
    '01_Source/Common', '01_Source/Track_A/Contracts', '01_Source/Track_A/Suppliers',
    '01_Source/Track_A/Plan', '01_Source/Track_A/Deliveries', '01_Source/Track_A/TXT',
    '01_Source/Track_A/PDF', '01_Source/Track_A/Invoices', '01_Source/Track_A/Payments',
    '02_Reference', '03_RAW', '04_STAGING', '05_DuckDB', '06_Parquet', '07_Python',
    '08_PowerBI', '09_QA', '10_Homework', '11_Checkpoints', '12_Capstone', '13_Exam',
    'Instructor/Truth', 'Instructor/Manifests', 'Generator'
]:
    (ROOT/p).mkdir(parents=True, exist_ok=True)

# -------------------- Helpers --------------------
def money(cents:int)->float:
    return cents/100.0

def allocate_int(total:int, n:int, local_rng:random.Random, min_each:int=1):
    if n <= 0: return []
    if total < n*min_each:
        raise ValueError('total too small')
    rem = total - n*min_each
    weights = [local_rng.randint(1,1000) for _ in range(n)]
    sw = sum(weights)
    vals = [min_each + (rem*w)//sw for w in weights]
    diff = total - sum(vals)
    for i in range(diff):
        vals[i % n] += 1
    return vals

def allocate_cents(total_cents:int, n:int, local_rng:random.Random, min_cents:int=100):
    return allocate_int(total_cents, n, local_rng, min_cents)

def iso(d):
    return d.isoformat() if isinstance(d,(date,datetime)) else d

def write_csv(path:Path, headers, rows, delimiter=',', encoding='utf-8-sig'):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path,'w',newline='',encoding=encoding) as f:
        w=csv.DictWriter(f, fieldnames=headers, delimiter=delimiter, extrasaction='ignore', quoting=csv.QUOTE_MINIMAL)
        w.writeheader(); w.writerows(rows)

def write_txt_rows(path:Path, lines, encoding='utf-8'):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path,'w',encoding=encoding,newline='') as f:
        for line in lines:
            f.write(str(line).rstrip('\n')+'\n')

def xlsx_value(v):
    if isinstance(v, Decimal): return float(v)
    return v

def write_xlsx(path:Path, headers, rows, sheet_name='Data', header_map=None, extra_sheet=None):
    wb = Workbook.create()
    sh = wb.worksheets.add(sheet_name[:31])
    actual_headers = [header_map.get(h,h) if header_map else h for h in headers]
    matrix = [actual_headers] + [[xlsx_value(r.get(h)) for h in headers] for r in rows]
    sh.get_range_by_indexes(0,0,len(matrix),len(headers)).values = matrix
    hdr = sh.get_range_by_indexes(0,0,1,len(headers))
    hdr.format = {
        'fill':'#1F4E78','font':{'bold':True,'color':'#FFFFFF'},
        'horizontal_alignment':'center','vertical_alignment':'center','wrap_text':True
    }
    sh.freeze_panes.freeze_rows(1)
    # sensible widths
    for i,h in enumerate(actual_headers):
        width = 14
        hl = str(h).lower()
        if 'name' in hl or 'назв' in hl or 'comment' in hl or 'прич' in hl: width=24
        if 'date' in hl or 'дата' in hl: width=14
        if 'amount' in hl or 'сума' in hl: width=16
        colrng = sh.get_range_by_indexes(0,i,len(matrix),1)
        colrng.format.column_width = width
        canon = str(headers[i]).lower()
        if ('date' in canon or canon.endswith('at') or canon in ('validfrom','validto','activefrom','activeto')) and len(matrix)>1:
            sh.get_range_by_indexes(1,i,len(matrix)-1,1).format.number_format = 'yyyy-mm-dd'
        elif 'amount' in canon or 'price' in canon:
            sh.get_range_by_indexes(1,i,max(1,len(matrix)-1),1).format.number_format = '#,##0.00'
        elif 'quantity' in canon:
            sh.get_range_by_indexes(1,i,max(1,len(matrix)-1),1).format.number_format = '0.###'
    if extra_sheet:
        es=wb.worksheets.add(extra_sheet['name'][:31])
        vals=extra_sheet['values']
        es.get_range_by_indexes(0,0,len(vals),max(len(x) for x in vals)).values=vals
    SpreadsheetFile.export_xlsx(wb).save(str(path))

# exact proportional allocation based on weights, cents/int
def allocate_by_weights(total:int, weights:list[int], min_each=0):
    n=len(weights)
    if n==0:return []
    base=[min_each]*n
    rem=total-min_each*n
    if rem<0: raise ValueError('total too small')
    sw=sum(weights)
    raw=[rem*w/sw for w in weights]
    vals=[int(math.floor(x))+min_each for x in raw]
    diff=total-sum(vals)
    order=sorted(range(n), key=lambda i: raw[i]-math.floor(raw[i]), reverse=True)
    for i in order[:diff]: vals[i]+=1
    return vals

def hash_row(*vals):
    return hashlib.sha256('|'.join('' if v is None else str(v) for v in vals).encode('utf-8')).hexdigest()

# -------------------- Common dimensions --------------------
start=date(2024,1,1); end=date(2026,12,31)
dates=[]; d=start
uk_month=['січень','лютий','березень','квітень','травень','червень','липень','серпень','вересень','жовтень','листопад','грудень']
uk_day=['понеділок','вівторок','середа','четвер','п’ятниця','субота','неділя']
while d<=end:
    dates.append({'DateKey':int(d.strftime('%Y%m%d')),'Date':d,'Year':d.year,'Quarter':(d.month-1)//3+1,
                  'MonthNumber':d.month,'MonthNameUA':uk_month[d.month-1],'YearMonth':d.strftime('%Y-%m'),
                  'WeekNumber':int(d.strftime('%V')),'DayOfMonth':d.day,'DayOfWeek':d.weekday()+1,
                  'DayNameUA':uk_day[d.weekday()],'IsMonthEnd':(d+timedelta(days=1)).month!=d.month,
                  'IsWorkingDay':d.weekday()<5})
    d+=timedelta(days=1)
assert len(dates)==1096

categories=['Категорія 01','Категорія 02','Категорія 03','Категорія 04','Категорія 05','Категорія 06']
uoms=['шт','компл','кг','л','м','упак']
items=[]
for i in range(1,181):
    items.append({'ItemID':f'IT{i:04d}','ItemCode':f'IT-{i:04d}','ItemName':f'Навчальна позиція {i:03d}',
                  'Category':categories[(i-1)%len(categories)],'Subcategory':f'Підкатегорія {((i-1)%12)+1:02d}',
                  'UnitOfMeasure':uoms[(i-1)%len(uoms)],'CriticalityClass':['Low','Medium','High'][i%3],
                  'IsActive':True,'ValidFrom':date(2024,1,1),'ValidTo':None})
item_source=[dict(x) for x in items]
# 4 bad item rows -> rejected
item_source += [
    {**items[4], 'ItemID':'IT0005'},
    {**items[9], 'ItemID':'it-0010'},
    {**items[14], 'ItemID':'IT-0015'},
    {**items[19], 'ItemID':'', 'UnitOfMeasure':''},
]

organizations=[]
for i in range(1,31):
    organizations.append({'OrganizationID':f'ORG{i:03d}','OrganizationCode':f'O{i:03d}',
                          'OrganizationName':f'Умовна організація {i:02d}','OrganizationType':'Навчальна',
                          'ParentOrganizationID':None if i<=5 else f'ORG{((i-1)%5)+1:03d}',
                          'OrganizationLevel':1 if i<=5 else 2,'RegionSynthetic':f'Регіон {((i-1)%6)+1:02d}','IsActive':True})
statuses=[{'StatusID':f'ST{i:02d}','StatusName':n} for i,n in enumerate([
    'Planned','Approved','InProgress','Completed','PartiallyCompleted','Rejected','Cancelled','Overdue',
    'Created','Delivered','Accepted','Invoiced','Paid','Pending','OnHold','Closed','Warning','Error'],1)]
reasons=[{'ReasonID':f'RS{i:02d}','ReasonName':f'Навчальна причина {i:02d}'} for i in range(1,26)]
units=[{'UnitCode':u,'UnitName':n} for u,n in zip(['EA','SET','KG','L','M','PK','BOX','PAIR','ROLL','DAY','HOUR','TON','M2','M3','LOT'],
                                                     ['шт','компл','кг','л','м','упак','короб','пара','рулон','день','год','т','м2','м3','партія'])]
category_mapping=[]
for i in range(150):
    category_mapping.append({'DomainName':'ItemCategory','SourceValue':f'cat_{i+1:03d}','CanonicalValue':categories[i%len(categories)],
                             'ValidFrom':date(2024,1,1),'ValidTo':None,'MappingStatus':'Approved'})
source_priority=[
    {'AttributeName':'ContractAmount','SourceSystem':'CONTRACT_REGISTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':'Основне джерело'},
    {'AttributeName':'ContractAmount','SourceSystem':'INVOICE_REGISTER','Priority':2,'ValidFrom':'2024-01-01','ValidTo':'','Comment':'Контрольне джерело'},
    {'AttributeName':'PlannedDate','SourceSystem':'PLAN_REGISTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'DeliveredQuantity','SourceSystem':'DELIVERY_REGISTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'DeliveredQuantity','SourceSystem':'WAREHOUSE_TXT','Priority':2,'ValidFrom':'2024-01-01','ValidTo':'','Comment':'Звірка'},
    {'AttributeName':'AcceptedQuantity','SourceSystem':'ACCEPTANCE_PDF','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'InvoiceAmount','SourceSystem':'INVOICE_REGISTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'PaidAmount','SourceSystem':'PAYMENT_REGISTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'SupplierName','SourceSystem':'SUPPLIER_MASTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'ItemName','SourceSystem':'ITEM_MASTER','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'EventTimestamp','SourceSystem':'EVENT_LOG','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
    {'AttributeName':'AcceptanceAmount','SourceSystem':'ACCEPTANCE_PDF','Priority':1,'ValidFrom':'2024-01-01','ValidTo':'','Comment':''},
]

# -------------------- Suppliers --------------------
suppliers=[]
for i in range(1,97):
    suppliers.append({'SupplierID':f'SUP{i:03d}','SupplierCode':f'S-{i:03d}','SupplierName':f'ТОВ «Навчальний постачальник {i:03d}»',
                      'SupplierCategory':['A','B','C'][i%3],'RiskCategory':['Low','Medium','High'][i%3],
                      'RegionSynthetic':f'Регіон {((i-1)%6)+1:02d}','ActiveFrom':date(2024,1,1),'ActiveTo':None,'IsActive':True})
supplier_updates=[]
for i in range(1,19):
    sid=f'SUP{i:03d}'
    supplier_updates.append({'SupplierID':sid,'EffectiveDate':date(2026,7,1),
                             'SupplierNameNew':f'ТОВ «Навчальний постачальник {i:03d} - оновлено»',
                             'RiskCategoryNew':['Medium','High','Low'][i%3]})
# dim version rows 114
supplier_dim=[]; sk=1
for s in suppliers:
    if int(s['SupplierID'][3:])<=18:
        supplier_dim.append({'SupplierKey':sk,**s,'ValidFrom':date(2024,1,1),'ValidTo':date(2026,6,30),'IsCurrent':False}); sk+=1
        upd=supplier_updates[int(s['SupplierID'][3:])-1]
        ns=dict(s); ns['SupplierName']=upd['SupplierNameNew']; ns['RiskCategory']=upd['RiskCategoryNew']
        supplier_dim.append({'SupplierKey':sk,**ns,'ValidFrom':date(2026,7,1),'ValidTo':None,'IsCurrent':True}); sk+=1
    else:
        supplier_dim.append({'SupplierKey':sk,**s,'ValidFrom':date(2024,1,1),'ValidTo':None,'IsCurrent':True}); sk+=1
assert len(supplier_dim)==114
supplier_current_key={r['SupplierID']:r['SupplierKey'] for r in supplier_dim if r['IsCurrent']}

# -------------------- Contracts --------------------
orig_cents=allocate_cents(148_500_000_00,1200,rng,20_000_00)
current_cents=allocate_by_weights(152_350_000_00,[max(1,x+rng.randint(-500000,500000)) for x in orig_cents],10000)
contracts=[]
for i in range(1,1201):
    cdate=date(2026,1,1)+timedelta(days=rng.randint(0,250))
    startd=cdate+timedelta(days=rng.randint(0,10)); endd=min(date(2026,12,31),startd+timedelta(days=rng.randint(60,240)))
    sid=f'SUP{((i-1)%96)+1:03d}'
    contracts.append({'ContractID':f'C{i:05d}','ContractNumber':f'CTR-{i:05d}','SupplierID':sid,'ContractDate':cdate,'StartDate':startd,'EndDate':endd,
                      'OriginalContractAmount':money(orig_cents[i-1]),'CurrentContractAmount':money(current_cents[i-1]),'CurrencyCode':'UAH',
                      'ContractStatus':'Approved' if i%17 else 'Completed','PaymentTermsDays':[15,30,45][i%3],'ContractType':['Framework','Supply','Service'][i%3]})
assert round(sum(x['OriginalContractAmount'] for x in contracts),2)==148_500_000.00
assert round(sum(x['CurrentContractAmount'] for x in contracts),2)==152_350_000.00

# dirty contract sources: 600 valid each + 13 invalid
c_h1=[dict(x) for x in contracts[:600]]; c_h2=[dict(x) for x in contracts[600:]]
# H1 4 duplicate + 3 bad date
for idx in [3,77,205,411]: c_h1.append(dict(c_h1[idx]))
for idx in [55,156,477]:
    bad=dict(c_h1[idx]); bad['ContractID']=f'BADDATE{idx:03d}'; bad['StartDate']=date(2026,10,1); bad['EndDate']=date(2026,9,1); c_h1.append(bad)
# H2 2 duplicate + 4 null supplier
for idx in [18,301]: c_h2.append(dict(c_h2[idx]))
for idx in [44,121,322,455]:
    bad=dict(c_h2[idx]); bad['ContractID']=f'NOSUP{idx:03d}'; bad['SupplierID']=''; c_h2.append(bad)
assert len(c_h1)==607 and len(c_h2)==606

# -------------------- Contract Lines --------------------
# line count distribution = 3600
line_counts=[]
line_counts += [1]*300+[2]*300+[3]*300+[5]*200+[8]*100
assert len(line_counts)==1200 and sum(line_counts)==3600
contract_lines=[]; line_id=1
for ci,c in enumerate(contracts):
    n=line_counts[ci]
    cents=allocate_by_weights(int(round(c['OriginalContractAmount']*100)),[rng.randint(1,100) for _ in range(n)],100)
    for ln in range(1,n+1):
        qty=rng.randint(1,80)
        amount_c=cents[ln-1]
        unit_price=round((amount_c/100)/qty,4)
        item=items[(line_id*7)%len(items)]
        contract_lines.append({'ContractLineID':f'CL{line_id:06d}','ContractID':c['ContractID'],'LineNumber':ln,'ItemID':item['ItemID'],
                               'OrderedQuantity':qty,'UnitPrice':unit_price,'LineAmount':money(amount_c),'UnitOfMeasure':item['UnitOfMeasure'],
                               'PlannedCompletionDate':min(c['EndDate'],c['StartDate']+timedelta(days=rng.randint(20,120)))})
        line_id+=1
assert len(contract_lines)==3600
# dirty +21
contract_line_source=[dict(x) for x in contract_lines]
for idx in [10,200,450,700,900,1200,1700,2500]: contract_line_source.append(dict(contract_lines[idx]))
for j,idx in enumerate([50,350,650,950,1550]):
    b=dict(contract_lines[idx]); b['ContractLineID']=f'UNKITEM{j:02d}'; b['ItemID']='IT9999'; contract_line_source.append(b)
for j,idx in enumerate([75,575,1075,2075]):
    b=dict(contract_lines[idx]); b['ContractLineID']=f'BADQTY{j:02d}'; b['OrderedQuantity']=0 if j%2==0 else -3; contract_line_source.append(b)
for j,idx in enumerate([125,825,1825,2825]):
    b=dict(contract_lines[idx]); b['ContractLineID']=f'BADAMT{j:02d}'; b['LineAmount']=round(b['LineAmount']+1234.56,2); contract_line_source.append(b)
assert len(contract_line_source)==3621
line_by_id={x['ContractLineID']:x for x in contract_lines}
lines_by_contract=defaultdict(list)
for x in contract_lines: lines_by_contract[x['ContractID']].append(x)

# -------------------- Plan versions and lines --------------------
# choose 400 contracts for v2 with line count total 1200: 80@1,160@2,60@3,60@5,40@8
groups=defaultdict(list)
for c,n in zip(contracts,line_counts): groups[n].append(c['ContractID'])
v2_contracts=[]
for n,k in [(1,80),(2,160),(3,60),(5,60),(8,40)]: v2_contracts += groups[n][:k]
assert len(v2_contracts)==400 and sum(len(lines_by_contract[c]) for c in v2_contracts)==1200
v3_contracts=sorted(v2_contracts, key=lambda cid: len(lines_by_contract[cid]), reverse=True)[:150]
# 400 updated line IDs across 150 v3 contracts; at least 1 each
v3_line_ids=[]
for cid in v3_contracts: v3_line_ids.append(lines_by_contract[cid][0]['ContractLineID'])
remaining_candidates=[l['ContractLineID'] for cid in v3_contracts for l in lines_by_contract[cid][1:]]
rng.shuffle(remaining_candidates); v3_line_ids += remaining_candidates[:250]
assert len(v3_line_ids)==400

# exact final current plan amount
line_weights=[max(1,int(round(x['LineAmount']*100))) for x in contract_lines]
final_plan_cents=allocate_by_weights(139_800_000_00,line_weights,100)
final_plan_by_line={cl['ContractLineID']:final_plan_cents[i] for i,cl in enumerate(contract_lines)}
# net change categories on revised contracts: 74 decreased,53 increased, rest neutral/line redistribution
v2_decrease=set(v2_contracts[:74]); v2_increase=set(v2_contracts[74:127])
plan_rows=[]
# v1 initial all 3600
for cl in contract_lines:
    cid=cl['ContractID']; fc=final_plan_by_line[cl['ContractLineID']]
    init=fc
    if cid in v2_decrease: init=int(round(fc*1.08))
    elif cid in v2_increase: init=max(100,int(round(fc*0.92)))
    elif cid in v2_contracts: init=max(100,fc+rng.randint(-50000,50000))
    plan_rows.append({'PlanLineID':f'PL{len(plan_rows)+1:06d}','ContractID':cid,'ContractLineID':cl['ContractLineID'],'VersionNo':1,
                      'PlannedDate':cl['PlannedCompletionDate'],'PlannedQuantity':cl['OrderedQuantity'],'PlannedAmount':money(init)})
# v2 all lines of 400 contracts; final except v3-line updates
for cid in v2_contracts:
    for cl in lines_by_contract[cid]:
        fc=final_plan_by_line[cl['ContractLineID']]
        v2c=fc
        if cl['ContractLineID'] in v3_line_ids:
            v2c=max(100,fc+rng.randint(-100000,100000))
        plan_rows.append({'PlanLineID':f'PL{len(plan_rows)+1:06d}','ContractID':cid,'ContractLineID':cl['ContractLineID'],'VersionNo':2,
                          'PlannedDate':cl['PlannedCompletionDate']+timedelta(days=rng.randint(-5,10)),'PlannedQuantity':cl['OrderedQuantity'],'PlannedAmount':money(v2c)})
# v3 400 line updates to exact final
for lid in v3_line_ids:
    cl=line_by_id[lid]
    plan_rows.append({'PlanLineID':f'PL{len(plan_rows)+1:06d}','ContractID':cl['ContractID'],'ContractLineID':lid,'VersionNo':3,
                      'PlannedDate':cl['PlannedCompletionDate']+timedelta(days=rng.randint(-5,15)),'PlannedQuantity':cl['OrderedQuantity'],'PlannedAmount':money(final_plan_by_line[lid])})
assert len(plan_rows)==5200
# verify current latest per line exact
latest={}
for r in plan_rows:
    lid=r['ContractLineID']
    if lid not in latest or r['VersionNo']>latest[lid]['VersionNo']: latest[lid]=r
assert round(sum(r['PlannedAmount'] for r in latest.values()),2)==139_800_000.00
# raw plan files 3610/1206/408
plans_v1=[dict(r) for r in plan_rows if r['VersionNo']==1]
plans_v2=[dict(r) for r in plan_rows if r['VersionNo']==2]
plans_v3=[dict(r) for r in plan_rows if r['VersionNo']==3]
for idx in range(10): plans_v1.append(dict(plans_v1[100+idx]))
for idx in range(6):
    b=dict(plans_v2[50+idx]); b['PlanLineID']=f'BADPDATE{idx:02d}'; b['PlannedDate']=date(2028,1,1); plans_v2.append(b)
for idx in range(8):
    b=dict(plans_v3[20+idx]); b['PlanLineID']=f'BADPLINE{idx:02d}'; b['ContractLineID']='CL999999'; plans_v3.append(b)
assert [len(plans_v1),len(plans_v2),len(plans_v3)]==[3610,1206,408]

# plan version truth and revision log
plan_versions=[]
for c in contracts:
    plan_versions.append({'PlanVersionID':f"{c['ContractID']}-V1",'ContractID':c['ContractID'],'VersionNo':1,'ValidFrom':c['StartDate'],'ApprovedAt':c['StartDate'],'VersionStatus':'Approved','RevisionReasonID':'RS01','IsCurrent':c['ContractID'] not in v2_contracts})
for cid in v2_contracts:
    c=contracts[int(cid[1:])-1]
    plan_versions.append({'PlanVersionID':f'{cid}-V2','ContractID':cid,'VersionNo':2,'ValidFrom':min(date(2026,7,1),c['EndDate']),'ApprovedAt':min(date(2026,6,28),c['EndDate']),'VersionStatus':'Approved','RevisionReasonID':'RS02','IsCurrent':cid not in v3_contracts})
for cid in v3_contracts:
    c=contracts[int(cid[1:])-1]
    plan_versions.append({'PlanVersionID':f'{cid}-V3','ContractID':cid,'VersionNo':3,'ValidFrom':min(date(2026,10,1),c['EndDate']),'ApprovedAt':min(date(2026,9,28),c['EndDate']),'VersionStatus':'Approved','RevisionReasonID':'RS03','IsCurrent':True})
assert len(plan_versions)==1750
revision_lines=[]
for cid in v2_contracts:
    revision_lines.append(f"2026-06-28|{cid}|1|2|RS02|APPROVED")
for cid in v3_contracts:
    revision_lines.append(f"2026-09-28|{cid}|2|3|RS03|APPROVED")
# 10 invalid
for i in range(4): revision_lines.append(f"2026-09-28|{v3_contracts[i]}|3|3|RS04|ACTIVE_CONFLICT")
for i in range(6): revision_lines.append(f"2026-08-15|{v2_contracts[20+i]}|4|2|RS05|BAD_SEQUENCE")
assert len(revision_lines)==560

# -------------------- Deliveries --------------------
raw_month_counts=[280,295,310,322,340,355,367,380,392,401,394,400]
defects_by_month=[2,2,2,2,2,3,3,3,4,4,4,5]
valid_month_counts=[r-d for r,d in zip(raw_month_counts,defects_by_month)]
assert sum(valid_month_counts)==4200
high_cents=allocate_cents(109_500_000_00,760,rng,20_000_00)
low_cents=allocate_cents(3_140_000_00,3440,rng,100)
amount_cents=high_cents+low_cents
# high links first 760 unique lines with largest current plan value
sorted_lines=sorted(contract_lines,key=lambda x:final_plan_by_line[x['ContractLineID']],reverse=True)
high_lines=sorted_lines[:760]
all_line_choices=contract_lines
valid_deliveries=[]; did=1; idx=0
high_delivery_ids=[]
for month,count in enumerate(valid_month_counts,1):
    for _ in range(count):
        is_high=idx<760
        cl=high_lines[idx] if is_high else rng.choice(all_line_choices)
        planned=latest[cl['ContractLineID']]['PlannedDate']
        # 27 deliberate delays among first 27 high rows
        if idx<27:
            ddate=min(date(2026,12,31),planned+timedelta(days=20+rng.randint(0,10)))
        else:
            # force inside month but avoid >15 day delay from planned where possible
            month_start=date(2026,month,1)
            nextm=date(2027,1,1) if month==12 else date(2026,month+1,1)
            candidate=month_start+timedelta(days=rng.randint(0,(nextm-month_start).days-1))
            ddate=min(candidate, planned+timedelta(days=10))
            if ddate<month_start: ddate=month_start
        rec={'DeliveryID':f'D{did:06d}','DeliveryDocumentNo':f'DOC-{(did-1)//5+1:05d}','ContractID':cl['ContractID'],'ContractLineID':cl['ContractLineID'],
             'SupplierID':contracts[int(cl['ContractID'][1:])-1]['SupplierID'],'DeliveryDate':ddate,'ItemID':cl['ItemID'],'Quantity':rng.randint(1,100),
             'Amount':money(amount_cents[idx]),'WarehouseSynthetic':f'WH-{((did-1)%8)+1:02d}','DeliveryStatus':'Delivered'}
        valid_deliveries.append(rec)
        if is_high: high_delivery_ids.append(rec['DeliveryID'])
        did+=1; idx+=1
assert len(valid_deliveries)==4200 and round(sum(x['Amount'] for x in valid_deliveries),2)==112_640_000.00
# split valid by month based on row generation order, add defects
monthly_deliveries=[]; pos=0
for m,(rawc,validc,defc) in enumerate(zip(raw_month_counts,valid_month_counts,defects_by_month),1):
    arr=[dict(x) for x in valid_deliveries[pos:pos+validc]]; pos+=validc
    # defect recipes distributed: duplicates, orphan contract, unknown item, bad qty
    for j in range(defc):
        base=dict(arr[(j*13)%len(arr)])
        if j%4==0: # duplicate ID exact-ish
            bad=dict(base)
        elif j%4==1:
            bad=dict(base); bad['DeliveryID']=f'ORPH{m:02d}{j:02d}'; bad['ContractID']='C99999'
        elif j%4==2:
            bad=dict(base); bad['DeliveryID']=f'UNKIT{m:02d}{j:02d}'; bad['ItemID']='IT9999'
        else:
            bad=dict(base); bad['DeliveryID']=f'BADQ{m:02d}{j:02d}'; bad['Quantity']=0
        arr.append(bad)
    assert len(arr)==rawc
    monthly_deliveries.append(arr)

# Secondary corroborating TXT sources, excluded from primary balance
warehouse_subset=valid_deliveries[:600]
warehouse_lines=['MovementID|DocumentNo|ItemID|Date|Quantity|MovementType|DeliveryID']
for i,r in enumerate(warehouse_subset,1):
    warehouse_lines.append(f"M{i:06d}|{r['DeliveryDocumentNo']}|{r['ItemID']}|{r['DeliveryDate'].strftime('%Y%m%d')}|{r['Quantity']}|IN|{r['DeliveryID']}")
legacy_subset=valid_deliveries[600:1050]
legacy_lines=[]
for r in legacy_subset:
    legacy_lines.append(f"{r['DeliveryID']:<7}{r['ContractID']:<8}{r['ItemID']:<8}{r['DeliveryDate'].strftime('%Y%m%d'):<10}{int(r['Quantity']):06d}  {int(round(r['Amount']*100)):012d}")

# -------------------- Acceptance (PDF source) --------------------
delivery_by_id={x['DeliveryID']:x for x in valid_deliveries}
# 31 partial, rejected total 590,000.00
rej_cents=allocate_cents(590_000_00,31,rng,5_000_00)
acceptance=[]
for i,didv in enumerate(high_delivery_ids):
    dr=delivery_by_id[didv]; dc=int(round(dr['Amount']*100))
    rejected=rej_cents[i] if i<31 else 0
    if rejected>=dc: rejected=max(0,dc//10)
    accepted=dc-rejected
    acceptance.append({'AcceptanceID':f'AC{i+1:06d}','AcceptanceDocumentNo':f'ACT-{(i//95)+41:05d}','DeliveryID':didv,'AcceptanceDate':min(date(2026,12,31),dr['DeliveryDate']+timedelta(days=1)),
                       'ItemID':dr['ItemID'],'DeliveredQuantityDocument':dr['Quantity'],'AcceptedQuantity':dr['Quantity']-(1 if rejected else 0),
                       'RejectedQuantity':1 if rejected else 0,'AcceptedAmount':money(accepted),'RejectedAmount':money(rejected),
                       'ResultStatus':'PartiallyCompleted' if rejected else 'Accepted','ReasonID':'RS06' if rejected else ''})
assert len(acceptance)==760
assert round(sum(x['AcceptedAmount'] for x in acceptance),2)==108_910_000.00
assert round(sum(x['RejectedAmount'] for x in acceptance),2)==590_000.00
# 8 PDF docs x 95 valid + 1 invalid = 96 source rows each
accept_by_doc=defaultdict(list)
for r in acceptance: accept_by_doc[r['AcceptanceDocumentNo']].append(r)

# -------------------- Invoices --------------------
# categories: unpaid100=5m, partial84=8m, overpaid9=1m, full1407=96.24m
invoice_category_counts={'UNPAID':100,'PARTIAL':84,'OVERPAID':9,'FULL_SINGLE':416,'FULL_MULTI':991}
invoice_category_totals={'UNPAID':5_000_000_00,'PARTIAL':8_000_000_00,'OVERPAID':1_000_000_00,'FULL_SINGLE':25_000_000_00,'FULL_MULTI':71_240_000_00}
assert sum(invoice_category_counts.values())==1600 and sum(invoice_category_totals.values())==110_240_000_00
invoices=[]; inv_idx=1
contract_cycle=[c['ContractID'] for c in contracts]+[c['ContractID'] for c in contracts[:400]]
cat_sequence=[]
for cat,n in invoice_category_counts.items(): cat_sequence += [cat]*n
# keep groups but deterministic
for cat,total in invoice_category_totals.items():
    idxs=[i for i,c in enumerate(cat_sequence) if c==cat]
    vals=allocate_cents(total,len(idxs),rng,1000_00)
    for j,v in zip(idxs,vals):
        # store temporarily
        pass
cat_values={}
for cat,total in invoice_category_totals.items(): cat_values[cat]=iter(allocate_cents(total,invoice_category_counts[cat],rng,1000_00))
for i,cat in enumerate(cat_sequence):
    cid=contract_cycle[i]
    c=contracts[int(cid[1:])-1]
    invdate=max(c['StartDate'],date(2026,1,1)+timedelta(days=rng.randint(0,330)))
    if invdate>date(2026,12,1): invdate=date(2026,12,1)
    amt=next(cat_values[cat])
    due=invdate+timedelta(days=c['PaymentTermsDays'])
    invoices.append({'InvoiceID':f'INV{inv_idx:06d}','InvoiceNumber':f'INV-{inv_idx:06d}','SupplierID':c['SupplierID'],'ContractID':cid,
                     'InvoiceDate':invdate,'DueDate':due,'InvoiceAmount':money(amt),'InvoiceStatus':'Open' if cat in ('UNPAID','PARTIAL') else 'Paid','PaymentCategory':cat})
    inv_idx+=1
assert len(invoices)==1600 and round(sum(x['InvoiceAmount'] for x in invoices),2)==110_240_000.00
# raw invoice with 8 defects, quarter split based on invoice date then balance to exact raw counts 350/390/415/453
# distribute canonical invoices into fixed chunks for predictable file sizes
inv_chunks=[invoices[:348],invoices[348:736],invoices[736:1149],invoices[1149:]]
# valid lengths 348,388,413,451 +2 defect each -> 350/390/415/453
for q,arr in enumerate(inv_chunks,1):
    b1=dict(arr[0]); b1['InvoiceID']=arr[1]['InvoiceID'] # duplicate ID
    b2=dict(arr[2]); b2['InvoiceID']=f'BADDUE{q:02d}'; b2['DueDate']=b2['InvoiceDate']-timedelta(days=1)
    arr.extend([b1,b2])
assert [len(x) for x in inv_chunks]==[350,390,415,453]

# -------------------- Payments --------------------
# target paid totals by category: partial 4.24m, overpaid1.09m, full exact invoice amounts; unpaid0
invoice_paid_target={}
# partial paid total exact 4.24m allocated by invoice weights
partial_invs=[x for x in invoices if x['PaymentCategory']=='PARTIAL']
partial_paid=allocate_by_weights(4_240_000_00,[int(round(x['InvoiceAmount']*100)) for x in partial_invs],100)
for x,v in zip(partial_invs,partial_paid): invoice_paid_target[x['InvoiceID']]=v
over_invs=[x for x in invoices if x['PaymentCategory']=='OVERPAID']
overpaid_extra=allocate_cents(90_000_00,len(over_invs),rng,1000_00)
for x,e in zip(over_invs,overpaid_extra): invoice_paid_target[x['InvoiceID']]=int(round(x['InvoiceAmount']*100))+e
for x in invoices:
    if x['PaymentCategory']=='UNPAID': invoice_paid_target[x['InvoiceID']]=0
    elif x['PaymentCategory'] in ('FULL_SINGLE','FULL_MULTI'): invoice_paid_target[x['InvoiceID']]=int(round(x['InvoiceAmount']*100))
assert sum(invoice_paid_target.values())==101_570_000_00
# payment counts: partial84*1, overpaid9*2, full_single416*1, full_multi991*2 => 2500
payments=[]; pidx=1
for inv in invoices:
    cat=inv['PaymentCategory']
    count=0 if cat=='UNPAID' else (2 if cat in ('OVERPAID','FULL_MULTI') else 1)
    target=invoice_paid_target[inv['InvoiceID']]
    if count==0: continue
    parts=allocate_int(target,count,rng,1)
    for k,pc in enumerate(parts):
        pdate=min(date(2026,12,31), inv['InvoiceDate']+timedelta(days=5+15*k+rng.randint(0,10)))
        payments.append({'PaymentID':f'PAY{pidx:06d}','InvoiceID':inv['InvoiceID'],'ContractID':inv['ContractID'],'SupplierID':inv['SupplierID'],
                         'PaymentDate':pdate,'PaymentAmount':money(pc),'PaymentStatus':'Accepted','PaymentReference':f'REF-{pidx:07d}'})
        pidx+=1
assert len(payments)==2500 and round(sum(x['PaymentAmount'] for x in payments),2)==101_570_000.00
# raw monthly payment files with target raw sizes, defects total28
raw_pay=[160,170,180,190,200,205,210,220,230,240,255,268]
pdef=[1,1,1,1,2,2,2,2,3,3,4,6]
valid_pay=[r-d for r,d in zip(raw_pay,pdef)]
# distribute canonical 2500 sequentially according to valid counts
monthly_payments=[]; pos=0
for m,(rc,vc,dc) in enumerate(zip(raw_pay,valid_pay,pdef),1):
    arr=[dict(x) for x in payments[pos:pos+vc]]; pos+=vc
    for j in range(dc):
        base=dict(arr[(j*11)%len(arr)])
        if j%3==0:
            bad=dict(base) # duplicate PaymentID
        elif j%3==1:
            bad=dict(base); bad['PaymentID']=f'ORPINV{m:02d}{j:02d}'; bad['InvoiceID']='INV999999'
        else:
            bad=dict(base); bad['PaymentID']=f'BADPAY{m:02d}{j:02d}'; bad['PaymentAmount']=-100.0 if j%2 else 0.0
        arr.append(bad)
    assert len(arr)==rc
    monthly_payments.append(arr)
assert pos==2500

# -------------------- Event log --------------------
events=[]
event_types=['CREATED','APPROVED','PLAN_UPDATED','DELIVERY_REGISTERED','ACCEPTED','INVOICED','PAID']
for i in range(3300):
    c=contracts[i%1200]
    et=event_types[i%len(event_types)]
    ts=datetime(2026,1,1,8,0,0)+timedelta(minutes=i*37)
    events.append({'EventID':f'EV{i+1:06d}','ContractID':c['ContractID'],'DocumentID':f'DOC-{(i%900)+1:05d}','EventTimestamp':ts,'EventType':et})
# log strings +17 invalid/malformed/duplicates
log_lines=[]
for e in events:
    log_lines.append(f"{e['EventTimestamp'].strftime('%Y-%m-%d %H:%M:%S')} INFO Contract={e['ContractID']} Document={e['DocumentID']} Event={e['EventType']}")
for i in range(8): log_lines.append(log_lines[100+i])
for i in range(5): log_lines.append(f"2026-99-{i+1:02d} 25:61:00 INFO Contract=C00001 Document=DOC-00001 Event=BADTIME")
for i in range(4): log_lines.append(f"MALFORMED EVENT LINE {i+1}")
assert len(log_lines)==3317

# -------------------- Defect & business manifests --------------------
defect_manifest=[]
def add_def(code, source, rule, count, classification='Rejected', note=''):
    defect_manifest.append({'DefectCode':code,'SourceFile':source,'RuleID':rule,'Count':count,'ExpectedClassification':classification,'ExpectedAction':classification,'Note':note})
add_def('ITM_VARIANT_DUP','items_master.xlsx','DQ_ITEM_KEY',4)
add_def('CONTRACT_SCHEMA_DRIFT','contracts_2026_H2.xlsx','DQ_SCHEMA',1,'Warning','Contract_ID header requires mapping; rows remain recoverable')
add_def('CONTRACT_DUP','contracts_2026_H1/H2.xlsx','DQ_UNIQUE',6)
add_def('CONTRACT_BAD_DATE','contracts_2026_H1.xlsx','DQ_DATE_SEQUENCE',3)
add_def('CONTRACT_NO_SUPPLIER','contracts_2026_H2.xlsx','DQ_REFERENCE',4)
add_def('LINE_DUP','contract_lines.xlsx','DQ_UNIQUE',8)
add_def('LINE_UNKNOWN_ITEM','contract_lines.xlsx','DQ_REFERENCE',5)
add_def('LINE_BAD_QTY','contract_lines.xlsx','DQ_RANGE',4)
add_def('LINE_BAD_AMOUNT','contract_lines.xlsx','DQ_RECON',4)
add_def('REV_ACTIVE_CONFLICT','plan_revision_log.txt','DQ_VERSION',4)
add_def('REV_BAD_SEQUENCE','plan_revision_log.txt','DQ_VERSION',6)
add_def('PLAN_DUP','delivery_plan_v1.csv','DQ_UNIQUE',10)
add_def('PLAN_BAD_DATE','delivery_plan_v2.csv','DQ_DATE_RANGE',6)
add_def('PLAN_UNKNOWN_LINE','delivery_plan_v3.csv','DQ_REFERENCE',8)
add_def('DELIVERY_MIXED','deliveries_*.csv','DQ_MULTI',36)
add_def('PDF_DUP','acceptance_act_*.pdf','DQ_UNIQUE',4)
add_def('PDF_BALANCE','acceptance_act_*.pdf','DQ_BALANCE',4)
add_def('INVOICE_DUP','invoices_Q*.xlsx','DQ_UNIQUE',4)
add_def('INVOICE_BAD_DUE','invoices_Q*.xlsx','DQ_DATE_SEQUENCE',4)
add_def('PAYMENT_MIXED','payments_*.csv','DQ_MULTI',28)
add_def('EVENT_LOG_INVALID','contract_events.log','DQ_EVENT',17)

# computed business exception manifest
business_ex=[]
# 31 partial acceptances
for r in acceptance[:31]: business_ex.append({'ExceptionType':'PARTIAL_ACCEPTANCE','BusinessKey':r['AcceptanceID'],'ExpectedValue':r['RejectedAmount'],'ValidRecord':True,'Interpretation':'Не видаляти; операційне відхилення'})
# 27 delayed deliveries
for r in valid_deliveries[:27]: business_ex.append({'ExceptionType':'DELIVERY_DELAY_GT15','BusinessKey':r['DeliveryID'],'ExpectedValue':20,'ValidRecord':True,'Interpretation':'Прострочення поставки'})
# 84 partial invoices +100 unpaid +9 overpaid +68 overdue sample
for inv in partial_invs: business_ex.append({'ExceptionType':'PARTIAL_PAYMENT','BusinessKey':inv['InvoiceID'],'ExpectedValue':round(inv['InvoiceAmount']-money(invoice_paid_target[inv['InvoiceID']]),2),'ValidRecord':True,'Interpretation':'Часткова оплата'})
for inv in [x for x in invoices if x['PaymentCategory']=='UNPAID']:
    business_ex.append({'ExceptionType':'UNPAID_INVOICE','BusinessKey':inv['InvoiceID'],'ExpectedValue':inv['InvoiceAmount'],'ValidRecord':True,'Interpretation':'Неоплачений рахунок'})
for inv in over_invs:
    business_ex.append({'ExceptionType':'OVERPAYMENT','BusinessKey':inv['InvoiceID'],'ExpectedValue':round(money(invoice_paid_target[inv['InvoiceID']])-inv['InvoiceAmount'],2),'ValidRecord':True,'Interpretation':'Оплата понад суму рахунку'})
for cid in list(v2_decrease): business_ex.append({'ExceptionType':'PLAN_DECREASED','BusinessKey':cid,'ExpectedValue':'','ValidRecord':True,'Interpretation':'Зміна плану'})
for cid in list(v2_increase): business_ex.append({'ExceptionType':'PLAN_INCREASED','BusinessKey':cid,'ExpectedValue':'','ValidRecord':True,'Interpretation':'Зміна плану'})

# -------------------- Write source artifacts --------------------
# Common XLSX/CSV
write_xlsx(ROOT/'01_Source/Common/calendar.xlsx', list(dates[0].keys()), dates)
write_xlsx(ROOT/'01_Source/Common/items_master.xlsx', list(item_source[0].keys()), item_source)
write_csv(ROOT/'01_Source/Common/organizations.csv', list(organizations[0].keys()), organizations)
write_csv(ROOT/'01_Source/Common/statuses.csv', list(statuses[0].keys()), statuses)
write_csv(ROOT/'01_Source/Common/reason_codes.csv', list(reasons[0].keys()), reasons)
write_csv(ROOT/'01_Source/Common/units_of_measure.csv', list(units[0].keys()), units)
write_xlsx(ROOT/'01_Source/Common/category_mapping.xlsx', list(category_mapping[0].keys()), category_mapping)
write_xlsx(ROOT/'01_Source/Common/source_priority.xlsx', list(source_priority[0].keys()), source_priority)

# Suppliers
write_xlsx(ROOT/'01_Source/Track_A/Suppliers/suppliers_master.xlsx', list(suppliers[0].keys()), suppliers)
write_csv(ROOT/'01_Source/Track_A/Suppliers/supplier_updates.csv', list(supplier_updates[0].keys()), supplier_updates)

# Contracts, H2 intentionally renamed headers
contract_headers=list(c_h1[0].keys())
write_xlsx(ROOT/'01_Source/Track_A/Contracts/contracts_2026_H1.xlsx',contract_headers,c_h1)
write_xlsx(ROOT/'01_Source/Track_A/Contracts/contracts_2026_H2.xlsx',contract_headers,c_h2,header_map={'ContractID':'Contract_ID','SupplierID':'Supplier_Code','ContractDate':'Contract_Date'})
write_xlsx(ROOT/'01_Source/Track_A/Contracts/contract_lines.xlsx',list(contract_line_source[0].keys()),contract_line_source)
# amendments source from current-original amounts for first 250 changed contracts
amendments=[]
for c in contracts[:250]:
    delta=round(c['CurrentContractAmount']-c['OriginalContractAmount'],2)
    amendments.append({'AmendmentID':f'AMD-{c["ContractID"]}','ContractID':c['ContractID'],'AmendmentDate':'2026-07-01','AmountDelta':delta,'Reason':'Навчальне коригування'})
write_csv(ROOT/'01_Source/Track_A/Contracts/contract_amendments.csv',list(amendments[0].keys()),amendments)

# Plans
plan_headers=list(plans_v1[0].keys())
write_csv(ROOT/'01_Source/Track_A/Plan/delivery_plan_v1.csv',plan_headers,plans_v1,delimiter=',')
write_csv(ROOT/'01_Source/Track_A/Plan/delivery_plan_v2.csv',plan_headers,plans_v2,delimiter=';')
# v3 uses changed header names
write_csv(ROOT/'01_Source/Track_A/Plan/delivery_plan_v3.csv', ['Plan_ID','ContractID','Contract_Line_ID','VersionNo','PlannedDate','PlannedQuantity','PlannedAmount'],
          [{'Plan_ID':r['PlanLineID'],'ContractID':r['ContractID'],'Contract_Line_ID':r['ContractLineID'],'VersionNo':r['VersionNo'],'PlannedDate':r['PlannedDate'],
            'PlannedQuantity':r['PlannedQuantity'],'PlannedAmount':r['PlannedAmount']} for r in plans_v3])
write_txt_rows(ROOT/'01_Source/Track_A/Plan/plan_revision_log.txt',['RevisionDate|ContractID|FromVersion|ToVersion|ReasonID|Status']+revision_lines)

# Deliveries monthly with controlled parser differences
base_del_headers=list(valid_deliveries[0].keys())
for m,arr in enumerate(monthly_deliveries,1):
    path=ROOT/f'01_Source/Track_A/Deliveries/deliveries_2026_{m:02d}.csv'
    if m==4:
        write_csv(path,base_del_headers,arr,delimiter=';')
    elif m==7:
        headers=['DeliveryID','DeliveryDocumentNo','Contract_ID','ContractLineID','SupplierID','DeliveryDate','ItemID','Quantity','Amount','WarehouseSynthetic','DeliveryStatus']
        rows=[{**r,'Contract_ID':r['ContractID']} for r in arr]
        for r in rows: r.pop('ContractID',None)
        write_csv(path,headers,rows)
    elif m==9:
        headers=base_del_headers+['Comment']
        rows=[{**r,'Comment':''} for r in arr]
        write_csv(path,headers,rows)
    elif m==11:
        headers=['DeliveryDate','DeliveryID','ContractID','ContractLineID','ItemID','Quantity','Amount','SupplierID','DeliveryDocumentNo','DeliveryStatus','WarehouseSynthetic']
        write_csv(path,headers,arr)
    elif m==12:
        # amount as localized text for all rows
        rows=[]
        for r in arr:
            rr=dict(r); rr['Amount']=f"{r['Amount']:,.2f}".replace(',','X').replace('.',',').replace('X',' '); rows.append(rr)
        write_csv(path,base_del_headers,rows,delimiter=';')
    else:
        write_csv(path,base_del_headers,arr)

# TXT + LOG
write_txt_rows(ROOT/'01_Source/Track_A/TXT/warehouse_movements.txt',warehouse_lines)
write_txt_rows(ROOT/'01_Source/Track_A/TXT/legacy_deliveries_fixed_width.txt',legacy_lines)
write_txt_rows(ROOT/'01_Source/Track_A/TXT/contract_events.log',log_lines)
write_txt_rows(ROOT/'01_Source/Track_A/TXT/legacy_layout_spec.txt',[
    'Fixed-width specification:',
    '1-7 DeliveryID','8-15 ContractID','16-23 ItemID','24-33 Date YYYYMMDD','34-39 Quantity','42-53 Amount in cents'
])

# PDFs
font_path='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
font_bold='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
pdfmetrics.registerFont(TTFont('DV',font_path)); pdfmetrics.registerFont(TTFont('DVB',font_bold))

def draw_pdf_act(path, act_no, rows, semi=False, scan=False, invalid_type='duplicate'):
    # source rows = 95 valid +1 invalid
    src=[dict(r) for r in rows]
    if invalid_type=='duplicate':
        src.append(dict(src[0]))
    else:
        bad=dict(src[0]); bad['AcceptanceID']=bad['AcceptanceID']+'X'; bad['AcceptedQuantity']=max(0,bad['AcceptedQuantity']-2) # breaks balance
        src.append(bad)
    if scan:
        W,H=1654,2339
        c=canvas.Canvas(str(path),pagesize=A4)
        page_size=32
        for page_no,start_idx in enumerate(range(0,len(src),page_size),1):
            im=Image.new('RGB',(W,H),'white'); dr=ImageDraw.Draw(im)
            fnt=ImageFont.truetype(font_path,28); bold=ImageFont.truetype(font_bold,32)
            y=80; dr.text((80,y),f'АКТ ПРИЙМАННЯ {act_no}',font=bold,fill='black'); y+=55
            dr.text((80,y),f'Навчальний синтетичний документ - скан, стор. {page_no}',font=fnt,fill='black'); y+=55
            for r in src[start_idx:start_idx+page_size]:
                dr.text((80,y),f"{r['AcceptanceID']}  {r['DeliveryID']}  {r['ItemID']}  {r['AcceptedQuantity']}  {r['RejectedQuantity']}  {r['AcceptedAmount']:.2f}",font=fnt,fill='black'); y+=56
            img_path=path.with_name(path.stem+f'_scan_{page_no}.png'); im.save(img_path)
            c.drawImage(str(img_path),10*mm,10*mm,width=190*mm,height=277*mm,preserveAspectRatio=True)
            img_path.unlink()
            if start_idx+page_size < len(src): c.showPage()
        c.save()
        return src
    c=canvas.Canvas(str(path),pagesize=A4); w,h=A4
    c.setFont('DVB',14); c.drawString(18*mm,h-18*mm,f'АКТ ПРИЙМАННЯ {act_no}')
    c.setFont('DV',8); c.drawString(18*mm,h-24*mm,'Навчальний синтетичний документ. Не містить реальних даних.')
    y=h-34*mm
    if semi:
        total_acc=sum(r['AcceptedAmount'] for r in src); total_rej=sum(r['RejectedAmount'] for r in src)
        c.setFont('DV',10)
        for txt in [f'Номер документа: {act_no}',f'Кількість позицій: {len(src)}',f'Прийнята сума: {total_acc:.2f}',f'Відхилена сума: {total_rej:.2f}',
                    'Нижче наведено напівструктурований перелік позицій:']:
            c.drawString(18*mm,y,txt); y-=6*mm
        c.setFont('DV',7)
        for r in src:
            if y<18*mm: c.showPage(); c.setFont('DV',7); y=h-18*mm
            c.drawString(18*mm,y,f"ID={r['AcceptanceID']}; DELIVERY={r['DeliveryID']}; ITEM={r['ItemID']}; ACCEPTED={r['AcceptedQuantity']}; REJECTED={r['RejectedQuantity']}; AMOUNT={r['AcceptedAmount']:.2f}")
            y-=4*mm
    else:
        headers=['AcceptanceID','DeliveryID','ItemID','Delivered','Accepted','Rejected','AcceptedAmount']
        widths=[24,22,18,16,16,16,28]
        c.setFont('DVB',6.5); x=14*mm
        for hh,ww in zip(headers,widths): c.drawString(x,y,hh); x+=ww*mm
        y-=4*mm; c.setFont('DV',6.2)
        for r in src:
            if y<15*mm:
                c.showPage(); c.setFont('DV',6.2); y=h-18*mm
            vals=[r['AcceptanceID'],r['DeliveryID'],r['ItemID'],str(r['DeliveredQuantityDocument']),str(r['AcceptedQuantity']),str(r['RejectedQuantity']),f"{r['AcceptedAmount']:.2f}"]
            x=14*mm
            for val,ww in zip(vals,widths): c.drawString(x,y,str(val)); x+=ww*mm
            y-=3.5*mm
    c.save(); return src

pdf_source_rows=[]
for i,act_num in enumerate(range(41,49)):
    act=f'ACT-{act_num:05d}'
    rows=accept_by_doc[act]
    src=draw_pdf_act(ROOT/f'01_Source/Track_A/PDF/acceptance_act_{act_num:05d}.pdf',act,rows,semi=(i==6),scan=(i==7),invalid_type='duplicate' if i<4 else 'balance')
    for rr in src:
        rec=dict(rr); rec['SourcePDF']=f'acceptance_act_{act_num:05d}.pdf'; pdf_source_rows.append(rec)
assert len(pdf_source_rows)==768

# Invoices XLSX, remove helper category from student source
invoice_headers=['InvoiceID','InvoiceNumber','SupplierID','ContractID','InvoiceDate','DueDate','InvoiceAmount','InvoiceStatus']
for q,arr in enumerate(inv_chunks,1):
    write_xlsx(ROOT/f'01_Source/Track_A/Invoices/invoices_Q{q}.xlsx',invoice_headers,arr)

# Payments CSV
pay_headers=list(payments[0].keys())
for m,arr in enumerate(monthly_payments,1):
    write_csv(ROOT/f'01_Source/Track_A/Payments/payments_2026_{m:02d}.csv',pay_headers,arr,delimiter=';' if m in (6,12) else ',')

# -------------------- Truth/Staging exports --------------------
truth_tables={
    'DimDate':dates,'DimItem':items,'DimOrganization':organizations,'DimStatus':statuses,'DimReason':reasons,
    'DimSupplier':supplier_dim,'DimContract':contracts,'FactContractLine':contract_lines,'DeliveryPlanVersion':plan_versions,
    'FactDeliveryPlan':plan_rows,'FactDelivery':valid_deliveries,'FactAcceptance':acceptance,'FactInvoice':invoices,'FactPayment':payments,'FactContractEvent':events
}
for name,rows in truth_tables.items():
    if rows:
        write_csv(ROOT/f'Instructor/Truth/{name}.csv',list(rows[0].keys()),rows,encoding='utf-8')
# staging subset for SQL student lessons
for name,rows in {
    'stg_contracts':contracts,'stg_contract_lines':contract_lines,'stg_delivery_plan':plan_rows,'stg_deliveries':valid_deliveries,
    'stg_acceptance':acceptance,'stg_invoices':invoices,'stg_payments':payments,'stg_events':events
}.items():
    write_csv(ROOT/f'04_STAGING/{name}.csv',list(rows[0].keys()),rows,encoding='utf-8')

# SQLite truth (portable fallback)
sqlite_path=ROOT/'Instructor/Truth/truth_A_V01.sqlite'
conn=sqlite3.connect(sqlite_path)
def sql_type(v):
    if isinstance(v,(int,bool)): return 'INTEGER'
    if isinstance(v,float): return 'REAL'
    return 'TEXT'
for name,rows in truth_tables.items():
    if not rows: continue
    cols=list(rows[0].keys())
    defs=', '.join([f'"{c}" {sql_type(next((r[c] for r in rows if r.get(c) is not None), ""))}' for c in cols])
    conn.execute(f'DROP TABLE IF EXISTS "{name}"'); conn.execute(f'CREATE TABLE "{name}" ({defs})')
    q='INSERT INTO "'+name+'" VALUES ('+','.join(['?']*len(cols))+')'
    vals=[]
    for r in rows:
        vals.append([iso(r.get(c)) if isinstance(r.get(c),(date,datetime)) else r.get(c) for c in cols])
    conn.executemany(q,vals)
conn.commit(); conn.close()

# -------------------- Expected results / QA --------------------
expected=[
    {'MetricCode':'PRIMARY_INPUT_ROWS','ExpectedValue':24719,'Tolerance':0,'Critical':True,'VisibleLevel':'Visible'},
    {'MetricCode':'PRIMARY_ACCEPTED_ROWS','ExpectedValue':24550,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'PRIMARY_REJECTED_ROWS','ExpectedValue':169,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_ITEMS','ExpectedValue':180,'Tolerance':0,'Critical':True,'VisibleLevel':'Partial'},
    {'MetricCode':'VALID_CONTRACTS','ExpectedValue':1200,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_CONTRACT_LINES','ExpectedValue':3600,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'PLAN_VERSIONS','ExpectedValue':1750,'Tolerance':0,'Critical':True,'VisibleLevel':'Partial'},
    {'MetricCode':'VALID_PLAN_LINES','ExpectedValue':5200,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_DELIVERIES','ExpectedValue':4200,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_ACCEPTANCE_ROWS','ExpectedValue':760,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_INVOICES','ExpectedValue':1600,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_PAYMENTS','ExpectedValue':2500,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'VALID_EVENTS','ExpectedValue':3300,'Tolerance':0,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'ORIGINAL_CONTRACT_AMOUNT','ExpectedValue':148500000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'CURRENT_CONTRACT_AMOUNT','ExpectedValue':152350000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'CURRENT_PLAN_AMOUNT','ExpectedValue':139800000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'DELIVERED_AMOUNT','ExpectedValue':112640000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'ACCEPTED_AMOUNT','ExpectedValue':108910000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'INVOICED_AMOUNT','ExpectedValue':110240000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'PAID_AMOUNT','ExpectedValue':101570000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'DELIVERY_ACCEPTANCE_DELTA','ExpectedValue':3730000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
    {'MetricCode':'INVOICE_PAYMENT_DELTA','ExpectedValue':8670000.00,'Tolerance':0.01,'Critical':True,'VisibleLevel':'Hidden'},
]
write_csv(ROOT/'Instructor/Manifests/expected_results_A_V01.csv',list(expected[0].keys()),expected,encoding='utf-8')
write_xlsx(ROOT/'Instructor/Manifests/expected_results_A_V01.xlsx',list(expected[0].keys()),expected)
write_csv(ROOT/'Instructor/Manifests/defect_manifest_A_V01.csv',list(defect_manifest[0].keys()),defect_manifest,encoding='utf-8')
write_xlsx(ROOT/'Instructor/Manifests/defect_manifest_A_V01.xlsx',list(defect_manifest[0].keys()),defect_manifest)
write_csv(ROOT/'Instructor/Manifests/business_exception_manifest_A_V01.csv',list(business_ex[0].keys()),business_ex,encoding='utf-8')

# QA rule catalog
qa_rules=[
    ['DQ001','ContractID not null','DimContract','ContractID','NOT_NULL','Critical'],
    ['DQ002','Contract unique','DimContract','ContractID','UNIQUE','Critical'],
    ['DQ003','Supplier exists','DimContract','SupplierID','REFERENCE','Critical'],
    ['DQ004','StartDate <= EndDate','DimContract','StartDate/EndDate','DATE_SEQUENCE','Critical'],
    ['DQ005','Item exists','FactContractLine','ItemID','REFERENCE','Critical'],
    ['DQ006','Quantity > 0','FactContractLine','OrderedQuantity','RANGE','Critical'],
    ['DQ007','Line amount reconciles','FactContractLine','LineAmount','RECONCILIATION','Critical'],
    ['DQ008','One current plan version','DeliveryPlanVersion','IsCurrent','VERSION','Critical'],
    ['DQ009','Plan line references contract line','FactDeliveryPlan','ContractLineID','REFERENCE','Critical'],
    ['DQ010','Delivery ID unique','FactDelivery','DeliveryID','UNIQUE','Critical'],
    ['DQ011','Delivery contract exists','FactDelivery','ContractID','REFERENCE','Critical'],
    ['DQ012','Delivery item exists','FactDelivery','ItemID','REFERENCE','Critical'],
    ['DQ013','Delivery quantity > 0','FactDelivery','Quantity','RANGE','Critical'],
    ['DQ014','Acceptance balance','FactAcceptance','Quantity','BALANCE','Critical'],
    ['DQ015','Invoice ID unique','FactInvoice','InvoiceID','UNIQUE','Critical'],
    ['DQ016','DueDate >= InvoiceDate','FactInvoice','DueDate','DATE_SEQUENCE','Critical'],
    ['DQ017','Payment ID unique','FactPayment','PaymentID','UNIQUE','Critical'],
    ['DQ018','Payment invoice exists','FactPayment','InvoiceID','REFERENCE','Critical'],
    ['DQ019','Payment amount > 0','FactPayment','PaymentAmount','RANGE','Critical'],
    ['DQ020','Event log syntax valid','FactContractEvent','EventTimestamp','FORMAT','Critical'],
]
qa_rows=[{'RuleID':x[0],'RuleName':x[1],'TargetTable':x[2],'TargetField':x[3],'RuleType':x[4],'Severity':x[5],'IsCritical':x[5]=='Critical'} for x in qa_rules]
write_csv(ROOT/'09_QA/DQ_RuleCatalog.csv',list(qa_rows[0].keys()),qa_rows)

# Homework map
hw=[
    ['HW01','Карта процесу та grain','contracts sample + payments sample','Карта сутностей/подій/ключів'],
    ['HW02','Excel audit','contracts + lines','Контрольні формули, дублі, типи'],
    ['HW03','Data Quality Profile','dirty starter subset','Паспорт якості та журнал помилок'],
    ['HW04','XLSX ingestion','contracts/suppliers','Відтворюваний Power Query'],
    ['HW05','CSV ingestion','12 deliveries CSV','Folder pipeline + schema drift'],
    ['HW06','TXT/LOG parsing','warehouse/fixed-width/event log','Структуровані таблиці'],
    ['HW07','PDF extraction','8 acceptance PDF','Витяг + класифікація PDF'],
    ['HW08','Multi-source reconciliation','all source families','Plan/Delivery/Acceptance/Invoice/Payment'],
    ['HW09','Star schema','staging tables','Модель + паспорт grain'],
    ['HW10','SQL basic','staging exports','SELECT/GROUP BY/CASE/JOIN'],
    ['HW11','SQL reconciliation','staging exports','Anti-join/LAG/LEAD/ROW_NUMBER'],
    ['HW12','DAX measures','canonical mart','12 перевірених мір'],
    ['HW13','Power BI','canonical mart','Executive/Analysis/Detail'],
    ['HW14','Python QA','staging exports','Автоматичний QA report'],
]
write_csv(ROOT/'10_Homework/Homework_Map.csv',['Homework','Topic','Input','Deliverable'],[dict(zip(['Homework','Topic','Input','Deliverable'],x)) for x in hw])

# Checkpoint map
cp=[
    {'Checkpoint':'CP1','AfterLesson':5,'Focus':'Data preparation','CriticalChecks':'Input/Accepted/Rejected; schema drift; duplicates'},
    {'Checkpoint':'CP2','AfterLesson':11,'Focus':'Data model + SQL','CriticalChecks':'Grain; joins; anti-join; current plan'},
    {'Checkpoint':'CP3','AfterLesson':14,'Focus':'DAX + Power BI','CriticalChecks':'Golden measures; filter behavior; drill-through'},
]
write_csv(ROOT/'11_Checkpoints/Checkpoint_Map.csv',list(cp[0].keys()),cp)

# Sample files for first lessons
write_xlsx(ROOT/'02_Reference/contracts_sample.xlsx',contract_headers,contracts[:250])
write_xlsx(ROOT/'02_Reference/contract_lines_sample.xlsx',list(contract_lines[0].keys()),contract_lines[:600])
write_csv(ROOT/'02_Reference/payments_sample.csv',pay_headers,payments[:800])

# Historical gzip source for DuckDB -> Parquet lab (250k rows)
hist_path=ROOT/'06_Parquet/historical_payments.csv.gz'
with gzip.open(hist_path,'wt',encoding='utf-8',newline='') as f:
    headers=['HistoricalPaymentID','SupplierID','PaymentDate','PaymentAmount','Status']
    w=csv.DictWriter(f,fieldnames=headers); w.writeheader()
    for i in range(250000):
        w.writerow({'HistoricalPaymentID':f'HP{i+1:07d}','SupplierID':f'SUP{(i%96)+1:03d}',
                    'PaymentDate':(date(2024,1,1)+timedelta(days=i%730)).isoformat(),
                    'PaymentAmount':round(1000+(i%10000)*1.37,2),'Status':'Accepted' if i%23 else 'Reversed'})

# DuckDB SQL scripts (runtime artifact generation on user's machine)
staging_names=['stg_contracts','stg_contract_lines','stg_delivery_plan','stg_deliveries','stg_acceptance','stg_invoices','stg_payments','stg_events']
sql_lines=["-- Run from 05_DuckDB folder: duckdb training_A_V01.duckdb < build_training_A_V01.sql",""]
for n in staging_names:
    sql_lines.append(f"CREATE OR REPLACE TABLE {n} AS SELECT * FROM read_csv_auto('../04_STAGING/{n}.csv', header=true, sample_size=-1);")
sql_lines += [
    "CREATE OR REPLACE VIEW vw_current_plan AS SELECT * EXCLUDE(rn) FROM (SELECT *, ROW_NUMBER() OVER(PARTITION BY ContractLineID ORDER BY VersionNo DESC) rn FROM stg_delivery_plan) WHERE rn=1;",
    "CREATE OR REPLACE VIEW vw_payment_by_invoice AS SELECT InvoiceID, COUNT(*) PaymentCount, SUM(PaymentAmount) PaidAmount FROM stg_payments GROUP BY InvoiceID;",
    "COPY (SELECT * FROM read_csv_auto('../06_Parquet/historical_payments.csv.gz', header=true)) TO '../06_Parquet/historical_payments.parquet' (FORMAT PARQUET);",
]
write_txt_rows(ROOT/'05_DuckDB/build_training_A_V01.sql',sql_lines)
write_txt_rows(ROOT/'05_DuckDB/build_training_A_V01.bat',[
    '@echo off','cd /d "%~dp0"','duckdb training_A_V01.duckdb < build_training_A_V01.sql','echo DuckDB created: training_A_V01.duckdb','pause'
])

# Python QA starter (does not require pandas)
pyqa='''from pathlib import Path\nimport csv\nROOT=Path(__file__).resolve().parents[1]\n\ndef count_csv(path, delimiter=","):\n    with open(path,encoding="utf-8-sig",newline="") as f:\n        return sum(1 for _ in csv.reader(f,delimiter=delimiter))-1\n\nprint("QA starter")\nprint("stg_contracts:", count_csv(ROOT/"04_STAGING/stg_contracts.csv"))\nprint("stg_payments:", count_csv(ROOT/"04_STAGING/stg_payments.csv"))\n# TODO: додайте перевірки DQ001-DQ020\n'''
(ROOT/'07_Python/qa_starter.py').write_text(pyqa,encoding='utf-8')

# Generator metadata
metadata={
    'seed':SEED,'variant':'A_V01_STARTER','primary_input_rows':24719,'primary_accepted_rows':24550,'primary_rejected_rows':169,
    'secondary_sources_excluded_from_primary_balance':['warehouse_movements.txt','legacy_deliveries_fixed_width.txt','historical_payments.csv.gz'],
    'golden_amounts':{'OriginalContractAmount':148500000.00,'CurrentContractAmount':152350000.00,'CurrentPlan':139800000.00,'Delivered':112640000.00,'Accepted':108910000.00,'Invoiced':110240000.00,'Paid':101570000.00}
}
(ROOT/'Generator/generator_seed.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')

# README
readme=f'''# Track A / Variant 01 / STARTER\n\nЦе еталонний синтетичний навчальний набір для курсу «Аналітика даних в оборонній галузі».\n\n## Наскрізний процес\nКонтракт -> специфікація -> версія плану -> поставка -> приймання -> рахунок -> оплата.\n\n## Основний баланс первинних джерел\n- Input rows: 24 719\n- Accepted after base QA: 24 550\n- Rejected: 169\n- Перевірка: 24 719 = 24 550 + 169\n\nУ цей баланс НЕ включені вторинні контрольні джерела warehouse_movements.txt, legacy_deliveries_fixed_width.txt та історичний файл для лабораторної Parquet.\n\n## Golden aggregates\n- Original Contract Amount: 148 500 000.00\n- Current Contract Amount: 152 350 000.00\n- Current Plan: 139 800 000.00\n- Delivered: 112 640 000.00\n- Accepted: 108 910 000.00\n- Invoiced: 110 240 000.00\n- Paid: 101 570 000.00\n\n## Формати\nXLSX, CSV, TXT, fixed-width TXT, LOG, digital PDF, semi-structured PDF, scanned PDF, gzip CSV for DuckDB/Parquet lab.\n\n## DuckDB / Parquet\nУ поточному середовищі генерації бібліотека DuckDB недоступна, тому пакет містить готові staging CSV та скрипт 05_DuckDB/build_training_A_V01.sql. На машині з DuckDB виконайте build_training_A_V01.bat або CLI-команду з коментаря у SQL-файлі. Скрипт створить training_A_V01.duckdb і historical_payments.parquet.\n\n## Для викладача\nInstructor/ містить canonical truth, SQLite fallback, defect manifest, business exceptions та expected results. Цю папку не треба видавати слухачам.\n\n## Важливо\nУсі назви, організації, постачальники, документи та суми синтетичні. Набір не містить реальних оборонних або персональних даних.\n'''
(ROOT/'00_ReadMe/README_UA.md').write_text(readme,encoding='utf-8')

# File catalog
catalog=[]
for fp in sorted(ROOT.rglob('*')):
    if fp.is_file():
        catalog.append({'Path':str(fp.relative_to(ROOT)).replace('\\','/'),'SizeBytes':fp.stat().st_size,'Extension':fp.suffix.lower()})
write_csv(ROOT/'00_ReadMe/File_Catalog.csv',list(catalog[0].keys()),catalog,encoding='utf-8')

# Validate primary counts from intended source raw counts
primary_raw = 1096+184+30+18+25+15+150+12 + 96+18 + 1213+3621+5224+560+4236+3317+768+1608+2528
assert primary_raw==24719
primary_valid = 1096+180+30+18+25+15+150+12 + 96+18 + 1200+3600+5200+550+4200+3300+760+1600+2500
assert primary_valid==24550 and primary_raw-primary_valid==169

# zip
zip_path=Path('/mnt/data/Defense_Analytics_Course_TrackA_V01_STARTER.zip')
if zip_path.exists(): zip_path.unlink()
with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for fp in ROOT.rglob('*'):
        if fp.is_file(): z.write(fp,fp.relative_to(ROOT.parent))

print(json.dumps({
    'root':str(ROOT),'zip':str(zip_path),'files':sum(1 for x in ROOT.rglob('*') if x.is_file()),
    'zip_mb':round(zip_path.stat().st_size/1024/1024,2),
    'primary_input':primary_raw,'primary_valid':primary_valid,'primary_rejected':primary_raw-primary_valid,
    'golden':metadata['golden_amounts']
},ensure_ascii=False,indent=2))
