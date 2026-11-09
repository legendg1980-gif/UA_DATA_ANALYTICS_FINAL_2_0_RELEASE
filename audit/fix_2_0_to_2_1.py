import sys, os, re, glob, copy, csv, sqlite3, io, warnings
import pandas as pd, numpy as np, openpyxl, docx
from datetime import datetime, timedelta
warnings.filterwarnings('ignore')
R=sys.argv[1]; LOG=[]
def log(s): LOG.append(s); print(s)
TA=R+'/07_Навчальні_дані/Track_A_V01_MASTER/'; SRC=TA+'01_Source/Track_A/'; TB=R+'/07_Навчальні_дані/Track_B_SUPPLEMENT/'
INS=R+'/99_Instructor_only/'; TR=INS+'V01_Instructor/Truth/'
# ---------- docx helpers
def paras(d):
    for p in d.paragraphs: yield p
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs: yield p
def set_text(p,new):
    if p.runs:
        p.runs[0].text=new
        for r in p.runs[1:]: r.text=''
    else: p.add_run(new)
def drep(path,pairs,must=True):
    d=docx.Document(path); n=0
    for old,new in pairs:
        hit=False
        for p in paras(d):
            if old in p.text: set_text(p,p.text.replace(old,new)); hit=True; n+=1
        if must and not hit: raise SystemExit(f'NOT FOUND in {path}: {old}')
    d.save(path); log(f'docx {os.path.basename(path)}: {n} replacements')
def dins_after(path,anchor,lines):
    d=docx.Document(path)
    for p in d.paragraphs:
        if p.text.strip()==anchor or p.text.startswith(anchor):
            ref=p
            for ln in lines:
                np_=copy.deepcopy(ref._p); ref._p.addnext(np_)
                newp=docx.text.paragraph.Paragraph(np_,ref._parent); set_text(newp,ln); ref=newp
            d.save(path); log(f'docx {os.path.basename(path)}: +{len(lines)} lines after "{anchor[:40]}"'); return
    raise SystemExit(f'anchor not found {path}: {anchor}')
def diag_table(path,mapping):
    d=docx.Document(path); n=0
    for t in d.tables:
        for row in t.rows:
            cells=row.cells
            if len(cells)>=2:
                for k,v in mapping.items():
                    if cells[0].text.strip().startswith(k):
                        set_text(cells[1].paragraphs[0],v); n+=1
    d.save(path); log(f'docx {os.path.basename(path)}: {n} diagnostics rewritten')
    if n!=len(mapping): raise SystemExit(f'diag mismatch {path} {n}/{len(mapping)}')

# ================= 1. Contract windows (Track A) =================
st=lambda n: pd.read_csv(TA+f'04_STAGING/stg_{n}.csv')
sc=st('contracts'); sd=st('deliveries'); sp=st('delivery_plan')
# ================= 4. Track A 13th delivery file
ln=st('contract_lines').set_index('ContractLineID')
rng=np.random.default_rng(20261109)
pick=ln[ln.ContractID.isin(sc.ContractID)&~ln.ContractID.isin(['C00105','C00412','C00648','C00837','C01045'])].sample(40,random_state=7)
rows=[];maxid=int(sd.DeliveryID.str[1:].astype(int).max()); maxdoc=int(sd.DeliveryDocumentNo.str[4:].astype(int).max())
for i,(clid,r) in enumerate(pick.iterrows()):
    q=int(rng.integers(1,max(2,int(r.OrderedQuantity)//3+1))); a=round(q*float(r.UnitPrice),2)
    rows.append(dict(DeliveryID=f'D{maxid+1+i:06d}',DeliveryDocumentNo=f'DOC-{maxdoc+1+i:05d}',ContractID=r.ContractID,ContractLineID=clid,SupplierID=sc.set_index('ContractID').SupplierID[r.ContractID],
      DeliveryDate=f'2027-01-{int(rng.integers(4,29)):02d}',ItemID=r.ItemID,Quantity=q,Amount=a,WarehouseSynthetic=f'WH-{int(rng.integers(1,9)):02d}',DeliveryStatus='Delivered'))
d13=pd.DataFrame(rows).sort_values('DeliveryDate')
os.makedirs(SRC+'ADD_AFTER_DEMO',exist_ok=True); d13.to_csv(SRC+'ADD_AFTER_DEMO/deliveries_2027_01.csv',index=False,lineterminator='\n')

first=pd.concat([sd[['ContractID','DeliveryDate']].rename(columns={'DeliveryDate':'D'}),sp[['ContractID','PlannedDate']].rename(columns={'PlannedDate':'D'}),d13[['ContractID','DeliveryDate']].rename(columns={'DeliveryDate':'D'})])
mn=first.groupby('ContractID').D.min(); mx=first.groupby('ContractID').D.max()
new=sc.set_index('ContractID')[['ContractDate','StartDate','EndDate']].copy()
new['StartDate']=np.minimum(new.StartDate, mn.reindex(new.index).fillna(new.StartDate))
new['EndDate']=np.maximum(new.EndDate, mx.reindex(new.index).fillna(new.EndDate))
new['ContractDate']=np.minimum(new.ContractDate,new.StartDate)
# seed 5 controlled "delivery after contract end" exceptions
seed=[]
cand=sd.groupby('ContractID').DeliveryDate.agg(['max','count'])
for cid in ['C00105','C00412','C00648','C00837','C01045']:
    last=cand.loc[cid,'max']; endd=(pd.Timestamp(last)-timedelta(days=12)).strftime('%Y-%m-%d')
    if endd<=new.loc[cid,'StartDate']: endd=(pd.Timestamp(new.loc[cid,'StartDate'])+timedelta(days=1)).strftime('%Y-%m-%d')
    new.loc[cid,'EndDate']=endd
    for _,r in sd[(sd.ContractID==cid)&(sd.DeliveryDate>endd)].iterrows(): seed.append((r.DeliveryID,cid,r.DeliveryDate,endd))
chg=(new!=sc.set_index('ContractID')[['ContractDate','StartDate','EndDate']]).any(axis=1)
log(f'contracts with adjusted dates: {int(chg.sum())}; seeded late-after-end deliveries: {len(seed)}')
def apply_df(df,idcol='ContractID'):
    m=df[idcol].isin(new.index)
    for c in ['ContractDate','StartDate','EndDate']: df.loc[m,c]=df.loc[m,idcol].map(new[c])
    return df
for f in [TA+'04_STAGING/stg_contracts.csv',TR+'DimContract.csv']:
    e_='utf-8-sig' if open(f,'rb').read(3)==b'\xef\xbb\xbf' else 'utf-8'; df=pd.read_csv(f,dtype=str); apply_df(df).to_csv(f,index=False,lineterminator='\n',encoding=e_); log('updated '+os.path.basename(f))
con=sqlite3.connect(TR+'truth_A_V01.sqlite')
for cid,r in new.iterrows(): con.execute('update DimContract set ContractDate=?,StartDate=?,EndDate=? where ContractID=?',(r.ContractDate,r.StartDate,r.EndDate,cid))
con.commit(); con.close(); log('updated truth sqlite DimContract')
def xl_contracts(path):
    wb=openpyxl.load_workbook(path); ws=wb.worksheets[0]; h=[c.value for c in ws[1]]
    ic=h.index('ContractID') if 'ContractID' in h else h.index('Contract_ID')
    cols={'ContractDate':h.index('ContractDate') if 'ContractDate' in h else h.index('Contract_Date'),'StartDate':h.index('StartDate'),'EndDate':h.index('EndDate')}
    n=0
    for row in ws.iter_rows(min_row=2):
        cid=row[ic].value
        if cid in new.index:
            sd_=pd.Timestamp(str(row[cols['StartDate']].value)[:10]) if row[cols['StartDate']].value else None
            for k,j in cols.items():
                v=datetime.strptime(new.loc[cid,k],'%Y-%m-%d'); cell=row[j]
                if isinstance(cell.value,datetime) and cell.value!=v: cell.value=v; n+=1
    wb.save(path); log(f'{os.path.basename(path)}: {n} date cells')
for f in [SRC+'Contracts/contracts_2026_H1.xlsx',SRC+'Contracts/contracts_2026_H2.xlsx',TA+'02_Reference/contracts_sample.xlsx',R+'/07_Навчальні_дані/Exercises/EX02_V01_dirty_register_800.xlsx']: xl_contracts(f)
# business exception manifest
bem=INS+'V01_Instructor/Manifests/business_exception_manifest_A_V01.csv'
with open(bem,'a',encoding='utf-8',newline='') as fh:
    w=csv.writer(fh,lineterminator='\n')
    for did,cid,dd,endd in seed: w.writerow(['DELIVERY_AFTER_CONTRACT_END',did,'',True,f'Поставка {dd} після EndDate {endd} контракту {cid}'])
# keep xlsx twin in sync if exists
log('business exception manifest +%d rows'%len(seed))

# ================= 2. Calendar DateKey format =================
for f in [TA+'01_Source/Common/calendar.xlsx']:
    wb=openpyxl.load_workbook(f); ws=wb.active
    for c in ws['A'][1:]: c.number_format='0'
    wb.save(f); log('calendar.xlsx DateKey format -> integer')

# ================= 3. EX02: real text-number / whitespace / serial-date defects (value preserving)
f=R+'/07_Навчальні_дані/Exercises/EX02_V01_dirty_register_800.xlsx'
wb=openpyxl.load_workbook(f); ws=wb.active; h=[c.value for c in ws[1]]
ia=h.index('CurrentContractAmount')+1; isup=h.index('SupplierID')+1; icd=h.index('ContractDate')+1
def ua(v): s=f'{v:,.2f}'.replace(',',' ').replace('.',','); return s
txtrows=[12,57,233,418,640]; wsrows=[25,96,310,512,701]; serrows=[140,377,566]
for r in txtrows: v=ws.cell(r,ia).value; ws.cell(r,ia).value=ua(float(v)); ws.cell(r,ia).number_format='@'
for r in wsrows:
    v=ws.cell(r,isup).value
    if v: ws.cell(r,isup).value=(' '+v if r%2 else v+' ')
for r in serrows:
    v=ws.cell(r,icd).value; ws.cell(r,icd).value=(v-datetime(1899,12,30)).days; ws.cell(r,icd).number_format='General'
wb.save(f); log('EX02: 5 text numbers, 5 whitespace keys, 3 serial dates (values preserved)')
ex=pd.read_excel(f,dtype=object)
def tonum(v): return float(str(v).replace(' ','').replace(' ','').replace(',','.')) if isinstance(v,str) else float(v)
amt=ex.CurrentContractAmount.map(tonum); ex['SupplierID']=ex.SupplierID.map(lambda v:v.strip() if isinstance(v,str) else v)
sup=set(pd.read_excel(SRC+'Suppliers/suppliers_master.xlsx').SupplierID)
cd=ex.ContractDate.map(lambda v: datetime(1899,12,30)+timedelta(days=int(v)) if isinstance(v,(int,float)) else v)
bad=(pd.to_datetime(ex.EndDate)<pd.to_datetime(ex.StartDate))|ex.SupplierID.isna()|~ex.SupplierID.isin(sup)|ex.duplicated(['ContractID','ContractNumber','SupplierID','CurrentContractAmount'])
# duplicates must be on normalized values
exn=ex.assign(A=amt); bad=(pd.to_datetime(ex.EndDate)<pd.to_datetime(ex.StartDate))|ex.SupplierID.isna()|~ex.SupplierID.isin(sup)|exn.duplicated(['ContractID','ContractNumber','SupplierID','A'])
EX=dict(input_rows=len(ex),input_total=round(amt.sum(),2),clean_rows=int((~bad).sum()),clean_total=round(amt[~bad].sum(),2),rejected=int(bad.sum()))
log(f'EX02 controls {EX}')

# extend contract EndDate where needed so 13th file creates no false "after end" exceptions
D13=dict(rows=len(d13),amount=round(d13.Amount.sum(),2))
log(f'13th Track A file: {D13}')
late13=d13[d13.apply(lambda r: r.DeliveryDate>new.loc[r.ContractID,'EndDate'],axis=1)]
log(f'13th-file rows after contract EndDate (expected business exceptions after refresh): {len(late13)}')


# ================= 5. Track B: B_EARLY_001 becomes a real early issue (48 > opening 42 before first receipt)
for fi in glob.glob(TB+'issues_2026_01.csv'):
    t=open(fi,encoding='utf-8-sig').read(); t2=re.sub(r'(B_EARLY_001,U05,I014,2026-01-02,ISS-EARLY-001,)6,',r'\g<1>48,',t)
    assert t!=t2; open(fi,'w',encoding='utf-8').write(t2); log('B_EARLY_001 qty 6 -> 48')
f=TB+'balance_snapshot_2026_12_31.xlsx'; wb=openpyxl.load_workbook(f); ws=wb.active; h=[c.value for c in ws[1]]
for row in ws.iter_rows(min_row=2):
    if row[0].value=='U05' and row[1].value=='I014':
        row[h.index('CalculatedClosing')].value=15; row[h.index('ReportedClosing')].value=15; row[h.index('DaysOfSupply')].value=3.0
wb.save(f); log('snapshot U05/I014 -> 15')
tbc=INS+'Track_B_Expected_Controls.csv'; t=open(tbc,encoding='utf-8-sig').read()
t=t.replace('ValidIssueTotal,13638','ValidIssueTotal,13680').replace('CalculatedClosingTotal,5338','CalculatedClosingTotal,5296').replace('EarlyIssueExampleCount,1,B_EARLY_001','EarlyIssueExampleCount,1,B_EARLY_001 (U05/I014: 48 > opening 42 before first receipt 03.02)')
open(tbc,'w',encoding='utf-8-sig').write(t); log('Track_B_Expected_Controls updated')

# ================= 6. Control_Results: real check + overdue 181
f=R+'/06_REFERENCE_REPORT/Control_Results_V01.xlsx'; wb=openpyxl.load_workbook(f)
ws=wb['Executive_Control']; ws['C3']='Actual (ввести)'
from openpyxl.styles import PatternFill
fill=PatternFill('solid',fgColor='FFF2CC')
for r in range(4,18):
    if ws.cell(r,1).value=='Overdue Invoice Count': ws.cell(r,4).value=181; ws.cell(r,2).value='invoices with balance > 0,005 UAH and DueDate < 2026-12-31'
    ws.cell(r,3).value=None; ws.cell(r,3).fill=fill
    ws.cell(r,6).value=(f'=IF(C{r}="","NOT CHECKED",IF(TEXT(C{r},"yyyy-mm-dd")=D{r},"PASS","FAIL"))' if r==4 else f'=IF(C{r}="","NOT CHECKED",IF(ABS(C{r}-D{r})<=E{r},"PASS","FAIL"))')
ws2=wb['Counts']; ws2.insert_cols(3); ws2['C1']='Actual (ввести)'
for r in range(2,ws2.max_row+1):
    ws2.cell(r,3).fill=fill; ws2.cell(r,6).value=f'=IF(C{r}="","NOT CHECKED",IF(ABS(C{r}-B{r})<=D{r},"PASS","FAIL"))'
n=ws2.max_row+1
for code,val,tol in [('OVERDUE_INVOICE_COUNT',181,0),('PLAN_TO_DATE_AMOUNT',139755212.66,0.01),('OUTSTANDING_AMOUNT',8760000,0.01),('OVERPAID_AMOUNT',90000,0.01),('OVERDUE_AMOUNT',8595136.32,0.01)]:
    ws2.cell(n,1).value=code; ws2.cell(n,2).value=val; ws2.cell(n,3).fill=fill; ws2.cell(n,4).value=tol; ws2.cell(n,5).value='True'; ws2.cell(n,6).value=f'=IF(C{n}="","NOT CHECKED",IF(ABS(C{n}-B{n})<=D{n},"PASS","FAIL"))'; n+=1
wb.save(f); log('Control_Results: Actual input column + live Status formulas, overdue=181')

# ================= 7. Docs
K=INS+'Ключі_викладача_MASTER_V01.docx'
drep(K,[('OVERDUE_INVOICE_COUNT | 295','OVERDUE_INVOICE_COUNT | 181')],must=False)
drep(K,[('/ 295 invoices','/ 181 invoices (правило: залишок > 0,005 грн; 295 у v2.0 був артефактом плаваючої коми)'),
        ('C00601 — v1 24.04.2026, v2 01.07.2026, v3 26.07.2026; as-of 15.07.2026 = v2.','C00601 — v2 затверджено 28.06.2026, v3 — 28.09.2026 (plan_revision_log); as-of 15.07.2026 = v2.')])
d=docx.Document(K)
for t in d.tables:
    for row in t.rows:
        if row.cells[0].text.strip()=='OVERDUE_INVOICE_COUNT': set_text(row.cells[1].paragraphs[0],'181')
d.save(K)
dins_after(K,'C00601 — v2',[
 f'HW01 / EX02: input 800 рядків, Σ CurrentContractAmount = {EX["input_total"]:,.2f}; після очищення {EX["clean_rows"]} рядків, Σ = {EX["clean_total"]:,.2f} (відхилено {EX["rejected"]}: 3 EndDate<StartDate, 1 без SupplierID, 8 повторів ключа).'.replace(',',' ').replace('.',','),
 f'HW03 / Track A ADD_AFTER_DEMO/deliveries_2027_01.csv: +{D13["rows"]} рядків, +{D13["amount"]:,.2f} грн; після refresh valid deliveries = 4 240.'.replace(',',' ').replace('.',',',1) if False else f'HW03 / Track A ADD_AFTER_DEMO/deliveries_2027_01.csv: +{D13["rows"]} рядків, +{D13["amount"]:.2f} грн; після refresh valid deliveries = 4 240 (golden-числа V01 рахуються без цього файлу).',
 f'HW08 / Track A: {len(seed)} поставок після EndDate контракту (DELIVERY_AFTER_CONTRACT_END: '+', '.join(s[0] for s in seed)+').',
 'Track B (v2.1): ValidIssueTotal = 13 680; CalculatedClosingTotal = 5 296; B_EARLY_001 = 48 од. при залишку 42 до першого надходження 03.02.2026.'])
W=INS+'Reference_Solution_Walkthrough_HW01_HW08.docx'
d=docx.Document(W)
for t in d.tables:
    for row in t.rows:
        c=row.cells
        if c[0].text.strip()=='ValidIssueTotal': set_text(c[1].paragraphs[0],'13680')
        if c[0].text.strip()=='CalculatedClosingTotal': set_text(c[1].paragraphs[0],'5296')
        if c[0].text.strip()=='EarlyIssueExampleCount': set_text(c[2].paragraphs[0],'B_EARLY_001: 48 > залишок 42 до першого надходження')
d.save(W)
drep(W,[('EX02 має 800 rows; BADDATE055 і NOSUP044 залишаються видимими як контрольні кейси.',f'EX02 має 800 rows; BADDATE055 і NOSUP044 залишаються видимими як контрольні кейси. Чисті {EX["clean_rows"]} rows, Σ CurrentContractAmount = {EX["clean_total"]:.2f}.'),
        ('13-й файл Track B підхоплюється без зміни query.',f'13-й файл (Track B issues_2027_01_ADD_AFTER_DEMO або Track A ADD_AFTER_DEMO/deliveries_2027_01, +{D13["rows"]} rows) підхоплюється без зміни query.')])
M=R+'/02_Instructor_Manual/Instructor_Manual_FINAL_2_0.docx'
drep(M,[('v2, якщо v3 починає діяти 26.07.','v2: v2 затверджено 28.06.2026, v3 — лише 28.09.2026 (plan_revision_log.txt).')])
# Session 1 literature
d=docx.Document(M); ps=d.paragraphs
for i,p in enumerate(ps):
    if p.text.startswith('K. Література після заняття 1'):
        set_text(ps[i+1],'Microsoft Support — Числа, відформатовані як текст: як знайти й виправити'); set_text(ps[i+2],'NIST SP 800-188 — De-Identifying Government Datasets'); break
d.save(M); log('manual: session 1 literature replaced')
HWD=R+'/03_Домашні_завдання/'
H=lambda n: glob.glob(HWD+f'HW0{n}_*.docx')[0]
drep(H(1),[('EXC0055: EndDate < StartDate — business validity issue.','BADDATE055 (а також BADDATE156, BADDATE477): EndDate < StartDate — business validity issue.'),
           ('EXC0077 — duplicate-key candidate, не видаляти без перевірки.','C00004, C00078, C00206, C00412 трапляються по 3 рази — duplicate-key candidates, не видаляти без перевірки ключ + дата + сума.'),
           ('Рядок із `12 300,50` має бути розпізнаний як текстове число до нормалізації.','5 сум записані текстом (формат «57 913,41»), 5 SupplierID мають зайві пробіли, 3 ContractDate приїхали числом Excel — розпізнати до нормалізації.')])
dins_after(H(1),'Input row count = 800.',[f'Контрольне число входу: Σ CurrentContractAmount усіх 800 рядків (після перетворення текстових чисел) = {EX["input_total"]:,.2f}.'.replace(',',' ').replace('.',',')])
dins_after(H(2),'NOSUP044',['Баланс Track A: Input = 24 719 рядків первинних джерел; Input = Accepted + Rejected.'])
dins_after(H(3),'До додавання — рівно 12 source files.',['Track A: 12 файлів, 4 236 сирих рядків; 13-й файл — 01_Source/Track_A/ADD_AFTER_DEMO/deliveries_2027_01.csv (+40 рядків), скопіювати в Deliveries лише за командою викладача.','Track B: 12 файлів, 1 172 рядки → 13 файлів, 1 212 рядків.'])
drep(H(3),[('Track A V01 Deliveries/deliveries_2026_01…12.csv АБО Track B SUPPLEMENT issues_2026_01…12.csv','Track A V01 Deliveries/deliveries_2026_01…12.csv (+ ADD_AFTER_DEMO/deliveries_2027_01.csv) АБО Track B SUPPLEMENT issues_2026_01…12.csv')])
dins_after(H(5),'C00601 as-of 15.07.2026 = v2.',['Plan versions (Contract × Version) = 1 750.'])
dins_after(H(6),'Core SQL не вимагає window functions.',['Рядки staging (Release 2): contracts 1 200, contract_lines 3 600, delivery_plan 5 200, deliveries 4 200, acceptance 760, invoices 1 600, payments 2 500, events 3 300.','Accepted рахується з stg_acceptance: скановий acceptance_act_00048.pdf без OCR не дає повного набору (OCR — optional-модуль).'])
DG={1:{'Заміна всіх null на 0':'Порівняйте COUNTBLANK до і після: якщо порожніх стало 0, а SUM не змінився — null знищено. Поверніть null і позначте «0» та «невідомо» окремим прапорцем.',
 'IFERROR навколо XLOOKUP':'Порахуйте результати-заглушки після IFERROR: кожен має бути в ErrorLog із ключем, якого немає в довіднику (напр. NOSUP044).',
 'видалення duplicate candidates':'800 − рядків результату має дорівнювати кількості записів ErrorLog з Rule = DQ_UNIQUE і поясненням ключ + дата + сума + документ.',
 'зміна raw data':'Відкрийте вихідний EX02 поруч: якщо значення не простежується до SourceRow, бракує копії сирого аркуша або колонки SourceRow.'},
2:{'Змішування business exception':'AC000001 (53 = 52 + 1) — валідне часткове приймання, а не дефект. DQ issue ламає правило даних, business exception — правило процесу.',
 'виключення рядка без Rule':'Input − Accepted − Rejected має дорівнювати 0, і кожен Rejected має RuleID із DQ_RuleCatalog.',
 'профілювання лише nulls':'Паспорт має рядок на кожен із 6 вимірів; якщо є лише completeness — додайте min/max дат, унікальність ключа, баланс.',
 'відсутність denominator':'% придатності = Accepted / Input; вкажіть обидва числа поруч із відсотком.'},
3:{'ручне Append':'Додайте 13-й файл: якщо його рядки не з’явилися без редагування запиту — pipeline ручний. Джерело в запиті має бути одне (Folder).',
 'видалення SourceFile':'Згрупуйте результат за SourceFile: має бути 12 (потім 13) груп; без колонки не знайти файл із помилковою цифрою.',
 'Auto Type без перевірки locale':'deliveries_2026_12.csv має десяткову кому і роздільник «;»: при Auto Type суми стають помилками або зростають у 100 разів — звірте Σ Amount по файлу.',
 'успішний refresh без row/total':'Після refresh заповніть Refresh_Log: InputFiles, InputRows, AcceptedRows, RejectedRows, ControlTotal; порожній рядок = контроль не виконано.'},
4:{'Merge по ненормалізованому key':'Порахуйте незіставлені до і після Trim/Upper ключа: якщо після нормалізації їх менше — Merge робився по сирому ключу.',
 'INNER JOIN, що приховує unmatched':'Rows(left) − Rows(inner) має дорівнювати Rows(left anti); без anti-join ORPINV1201 і B_ORPH_001 зникають мовчки.',
 'розмноження rows без control':'Порівняйте COUNT і SUM до і після Merge: зростання (напр. lines × payments) означає неправильну гранулярність ключа.',
 'hard-coded replacement':'Кожен крок Replace Values має відповідати рядку в таблиці відповідностей «як написано → як має бути».'},
5:{'прямий fact-to-fact join':'У моделі не має бути зв’язку між двома фактами; обидва фільтруються через спільні довідники (Contract, ContractLine, Date).',
 'current plan замість as-of':'Перевірте C00601 на 15.07.2026: має бути v2. Якщо показано v3 — as-of береться з поточної версії, а не з plan_revision_log.',
 'many-to-many як':'Зв’язок *:* — сигнал неунікального ключа довідника; порівняйте COUNTROWS і DISTINCTCOUNT ключа.',
 'один змішаний %':'Виконання поставки (Delivered / Plan) і оплати (Paid / Invoiced) мають різні знаменники — показуйте окремо.'},
6:{'COUNT(*) замість DISTINCTCOUNT':'Порівняйте COUNT(*) і COUNT(DISTINCT ContractID) у stg_payments: різниця = контракти з кількома платежами.',
 'DAX measure без control':'Для кожної міри в таблиці звірки має бути число SQL і число PQ; порожня клітинка = міра не прийнята.',
 'hard-coded date':'Знайдіть у мірах літерали дат: дата звіту — одна міра/параметр Reporting As Of Date = 31.12.2026.',
 'window functions як':'Перепишіть через GROUP BY + JOIN на підзапит MAX(VersionNo); core-рівень не потребує OVER().'},
7:{'TODAY()':'Відкрийте файл іншого дня: якщо числа змінились — у мірах є TODAY(); замініть на Reporting As Of Date.',
 '20+ KPI':'Порахуйте візуали Executive: 5–7 KPI, тренд, топ ризиків; решта — на Analytical.',
 'scenario без assumptions':'Кожен сценарій має рядок у Scenario_Assumptions із множником темпу й переліком того, що не враховано.',
 'drillthrough до агрегованої':'Drill-through має закінчуватись таблицею з ID документа (InvoiceID, DeliveryID, IssueID) і SourceFile.'},
8:{'колір замість формального rule':'Для кожного червоного прапорця заповніть Exception_Register: Rule, Threshold, AsOfDate, Owner; без цих полів це не виняток.',
 'змішування фактів і припущень':'Позначте кожне речення записки як Факт / Припущення / Рекомендація; факт посилається на контрольне число.',
 'відсутність lineage':'Від кожного винятку пройдіть drill-through до рядка з SourceFile і SourceRow; не виходить — lineage втрачено.',
 'ігнорування hidden sheets':'Перед передачею: «Перевірити документ» (Inspect Document), показати приховані аркуші/колонки, очистити властивості файлу.'}}
for n,m in DG.items(): diag_table(H(n),m)
P=R+'/03_Домашні_завдання/Практичний_посібник_8_ДЗ_FINAL_2_0.docx'
drep(P,[('09_Додаткові_матеріали/Архів_MASTER_14_ДЗ','09_Додаткові_матеріали/MASTER_14_HW')])
# Homework_Map
hm=TA+'10_Homework/Homework_Map.csv'
open(hm,'w',encoding='utf-8-sig').write('''Homework,AfterLesson,Topic,Input,Deliverable
HW01,2,Карта процесу + Excel audit,Exercises/EX02_V01_dirty_register_800.xlsx,"Process map, clean workbook, ErrorLog, control sum"
HW02,3,Data Quality Passport,Track A 01_Source / Track B,"DQ Passport, Deviation log, 10-row sample"
HW03,4,Power Query Folder 12→13,Deliveries 12 CSV + ADD_AFTER_DEMO / Track B issues,"PQ query, RejectedFiles, Refresh log"
HW04,6,Normalization + Merge + Anti-join,all source families,"Mapping table, anti-join result, reconciliation control"
HW05,8,Plan-fact + Star schema,Plan v1-v3 + deliveries / need_plan,"Plan-fact deltas, model diagram, grain passport"
HW06,10,SQL + DAX reconciliation,04_STAGING (Release 2),"SQL file, DAX measures, SQL/PQ/DAX table, 5 KPI passports"
HW07,12,Time metrics + 2 BI pages,model,"Executive + Analytical pages, scenarios"
HW08,14,Profile + Exceptions + QA + note,model,"Profile page, Exception Register, checklist, 5 test cases, 1-page note"
''')
log('Homework_Map.csv -> HW01–HW08')
rm=TA+'00_ReadMe/README_UA.md'; t=open(rm,encoding='utf-8').read()
t=t.replace('Instructor/ містить','99_Instructor_only/V01_Instructor/ містить'); open(rm,'w',encoding='utf-8').write(t)
import json; json.dump(dict(EX=EX,D13=D13,seed=seed,LOG=LOG),open('/home/claude/fix/fixlog.json','w'),ensure_ascii=False,indent=1,default=str)

# ================= 5b. Track B: receipt dates so stock never goes negative intra-year (except seeded cases)
rcf=TB+'receipts_2026.csv'; rc=pd.read_csv(rcf,dtype=str); rc['Qn']=rc.Quantity.astype(int)
ob=pd.read_excel(TB+'opening_balance.xlsx'); rt=pd.read_csv(TB+'returns_2026.csv'); wo=pd.read_csv(TB+'writeoffs_2026.csv')
iss=pd.concat([pd.read_csv(f) for f in sorted(glob.glob(TB+'issues_2026_*.csv'))]); iss=iss[~iss.IssueID.isin(['B_ORPH_001','B_DUP_001'])]
skip={('U08','I030')}
moved=0
for (u,i),g in iss.groupby(['UnitID','ItemID']):
    if (u,i) in skip: continue
    for _ in range(50):
        ev=[('0000',0,int(ob[(ob.UnitID==u)&(ob.ItemID==i)].OpeningQuantity.sum()),'OPEN')]
        r_=rc[(rc.UnitID==u)&(rc.ItemID==i)]
        ev+= [(d,0,int(q),'R'+rid) for d,q,rid in zip(r_.ReceiptDate,r_.Qn,r_.ReceiptID)]
        rr=rt[(rt.UnitID==u)&(rt.ItemID==i)]; ev+= [(d,0,int(q),'T'+tid) for d,q,tid in zip(rr.ReturnDate,rr.Quantity,rr.ReturnID)]
        ev+= [(d,1,-int(q),'I') for d,q in zip(g.IssueDate,g.Quantity)]
        ev+= [(d,1,-int(q),'W') for d,q in zip(wo[(wo.UnitID==u)&(wo.ItemID==i)].WriteOffDate,wo[(wo.UnitID==u)&(wo.ItemID==i)].Quantity)]
        ev.sort(); bal=0; fixd=False
        for d,o,q,k in ev:
            bal+=q
            if bal<0:
                if (u,i)==('U05','I014') and d=='2026-01-02': continue
                fut=[e for e in ev if (e[3].startswith('RR') or e[3].startswith('T')) and e[0]>d]
                if not fut: break
                k_=fut[0][3]; nd=(pd.Timestamp(d)-timedelta(days=1)).strftime('%Y-%m-%d')
                if k_.startswith('RR'): rc.loc[rc.ReceiptID==k_[1:],'ReceiptDate']=nd
                else: rt.loc[rt.ReturnID==k_[1:],'ReturnDate']=nd
                moved+=1; fixd=True; break
        if not fixd: break
enc=lambda p:'utf-8-sig' if open(p,'rb').read(3)==b'\xef\xbb\xbf' else 'utf-8'
rc.drop(columns='Qn').to_csv(rcf,index=False,lineterminator='\n',encoding=enc(rcf)); rp=TB+'returns_2026.csv'; rt.to_csv(rp,index=False,lineterminator='\n',encoding=enc(rp)); log(f'Track B receipts re-dated: {moved}')
json.dump(dict(EX=EX,D13=D13,seed=seed,LOG=LOG),open('/home/claude/fix/fixlog.json','w'),ensure_ascii=False,indent=1,default=str)

# ================= 8. text polish of generated lines
drep(K,[(next(p.text for p in docx.Document(K).paragraphs if p.text.startswith('HW01 / EX02')),
         'HW01 / EX02: input 800 рядків, Σ CurrentContractAmount = 101 926 202,98; після очищення 788 рядків, Σ = 100 132 639,21 (відхилено 12: 3 EndDate < StartDate, 1 без SupplierID, 8 повторів ключа).'),
        (f'+{D13["amount"]:.2f} грн', (f'+{D13["amount"]:,.2f} грн').replace(',',' ').replace('.',',')),
        ('4 поставок після','4 поставки після')],must=False)
drep(H(1),[('(після перетворення текстових чисел) = 101 926 202,98,','(після перетворення текстових чисел) = 101 926 202,98.')])
drep(W,[('Σ CurrentContractAmount = 100132639.21.','Σ CurrentContractAmount = 100 132 639,21.')])
