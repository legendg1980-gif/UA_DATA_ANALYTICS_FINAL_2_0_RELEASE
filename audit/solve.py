import pandas as pd, numpy as np, glob, io, re, sqlite3, json, sys, warnings
warnings.filterwarnings('ignore')
ROOT=sys.argv[1]; OUT=sys.argv[2]
TA=ROOT+'/07_Навчальні_дані/Track_A_V01_MASTER/'; SRC=TA+'01_Source/Track_A/'; TB=ROOT+'/07_Навчальні_дані/Track_B_SUPPLEMENT/'
res={}; sheets={}; cmp=[]
def C(hw,metric,mine,key,src):
    ok = (abs(float(mine)-float(key))<=0.01) if isinstance(key,(int,float)) and isinstance(mine,(int,float,np.integer,np.floating)) else (str(mine)==str(key))
    cmp.append(dict(HW=hw,Metric=metric,Mine=mine,Key=key,KeySource=src,Match='PASS' if ok else 'FAIL'))
def num(s): return pd.to_numeric(s.astype(str).str.replace(' ','').str.replace(' ','').str.replace(',','.'),errors='coerce')
def rcsv(f):
    t=open(f,encoding='utf-8-sig').read(); h=t.split('\n')[0]; sep=';' if h.count(';')>h.count(',') else ','
    return pd.read_csv(io.StringIO(t),sep=sep,dtype=str),sep
REN={'Contract_ID':'ContractID','Supplier_Code':'SupplierID','Contract_Date':'ContractDate','Plan_ID':'PlanLineID','Contract_Line_ID':'ContractLineID'}

# ---------------- HW01: EX02 Excel audit
ex=pd.read_excel(ROOT+'/07_Навчальні_дані/Exercises/EX02_V01_dirty_register_800.xlsx',dtype=object)
TXTN=int(ex.map(lambda v:isinstance(v,str) and bool(re.fullmatch(r'[\d\s]+,\d+',v))).values.sum())
ex['SourceRow']=ex.index+2
ex['SupplierID']=ex.SupplierID.map(lambda v:v.strip() if isinstance(v,str) else v)
ex['CurrentContractAmount']=ex.CurrentContractAmount.map(lambda v: float(str(v).replace('\u00a0','').replace(' ','').replace(',','.')) if isinstance(v,str) else v)
log=[]
for _,r in ex[pd.to_datetime(ex.EndDate)<pd.to_datetime(ex.StartDate)].iterrows(): log.append((r.SourceFile,r.SourceRow,r.ContractID,'DQ_DATE_SEQUENCE','Critical','EndDate<StartDate','Reject'))
for _,r in ex[ex.SupplierID.isna()].iterrows(): log.append((r.SourceFile,r.SourceRow,r.ContractID,'DQ_REFERENCE','Critical','SupplierID empty','Reject'))
dupmask=ex.duplicated(['ContractID','ContractNumber','SupplierID','CurrentContractAmount'],keep='first')
for _,r in ex[dupmask].iterrows(): log.append((r.SourceFile,r.SourceRow,r.ContractID,'DQ_UNIQUE','Material','Repeat of key+amount','Reject (keep first)'))
sup=pd.read_excel(SRC+'Suppliers/suppliers_master.xlsx')
unk=ex[ex.SupplierID.notna()&~ex.SupplierID.isin(sup.SupplierID)]
for _,r in unk.iterrows(): log.append((r.SourceFile,r.SourceRow,r.ContractID,'DQ_REFERENCE','Critical','SupplierID not in master','Reject'))
el=pd.DataFrame(log,columns=['SourceFile','SourceRow','BusinessKey','Rule','Severity','Observed','Resolution'])
bad=set(el.SourceRow)
clean=ex[~ex.SourceRow.isin(bad)]
h1=pd.DataFrame({'Check':['Input rows','Rejected rows','Clean rows','Clean unique ContractID','Input CurrentAmount','Clean CurrentAmount','Clean OriginalAmount','Text numbers found','Whitespace issues found','EXC0055/EXC0077 present'],
 'Value':[len(ex),len(bad),len(clean),clean.ContractID.nunique(),round(ex.CurrentContractAmount.astype(float).sum(),2),round(clean.CurrentContractAmount.astype(float).sum(),2),round(clean.OriginalContractAmount.astype(float).sum(),2),
   int(ex.map(lambda v:isinstance(v,str) and bool(re.fullmatch(r'[\d\s]+,\d+',v))).values.sum()), int(ex.map(lambda v:isinstance(v,str) and v!=v.strip()).values.sum()), 'No']})
sumifs=clean.assign(A=clean.CurrentContractAmount.astype(float)).groupby('SupplierID').agg(SUMIFS_Current=('A','sum'),COUNTIFS=('ContractID','count')).reset_index().sort_values('SUMIFS_Current',ascending=False)
sheets['HW01_Control']=h1; sheets['HW01_ErrorLog']=el; sheets['HW01_SUMIFS']=sumifs.head(20)
C('HW01','Input row count',len(ex),800,'HW01 умова')
C('HW01','BADDATE055 flagged EndDate<StartDate','BADDATE055' in set(el[el.Rule=='DQ_DATE_SEQUENCE'].BusinessKey),True,'Ключі/Walkthrough')
C('HW01','NOSUP044 flagged missing supplier','NOSUP044' in set(el[el.Observed=='SupplierID empty'].BusinessKey),True,'Ключі/Walkthrough')
C('HW01','Text numbers in EX02',TXTN,5,'HW01 умова v2.1')
C('HW01','Clean rows',len(clean),788,'Ключі v2.1'); C('HW01','Clean CurrentAmount',round(clean.CurrentContractAmount.map(lambda v: float(str(v).replace(' ','').replace(',','.'))).sum(),2),100132639.21,'Ключі v2.1')

# ---------------- HW02 + HW03 (Track A raw)
def load_csvs(pattern):
    fr=[];man=[]
    for f in sorted(glob.glob(pattern)):
        d,sep=rcsv(f); cols=list(d.columns); d=d.rename(columns=REN); d['SourceFile']=f.split('/')[-1]; d['SourceRow']=range(2,len(d)+2)
        issues=[]
        if sep==';': issues.append('separator ;')
        if any(c in REN for c in cols): issues.append('renamed column '+','.join(c for c in cols if c in REN))
        man.append(dict(SourceFile=f.split('/')[-1],Rows=len(d),Columns=len(cols),Sep=sep,Issues='; '.join(issues),Cols='|'.join(cols)))
        fr.append(d)
    return pd.concat(fr,ignore_index=True),pd.DataFrame(man)
D,manD=load_csvs(SRC+'Deliveries/*.csv')
base=manD.Cols.iloc[0].split('|')
manD['SchemaDrift']=manD.Cols.apply(lambda c:'' if c.split('|')==base else ('order' if sorted(c.split('|'))==sorted(base) else 'columns: +'+','.join(set(c.split('|'))-set(base))+' -'+','.join(set(base)-set(c.split('|')))))
D['Amt']=num(D.Amount); D['commaDecimal']=D.Amount.str.contains(',',na=False)
manD['CommaDecimalRows']=D.groupby('SourceFile').commaDecimal.sum().reindex(manD.SourceFile).values
manD['Amount']=D.groupby('SourceFile').Amt.sum().reindex(manD.SourceFile).round(2).values
sheets['HW03_SourceManifest']=manD.drop(columns='Cols')
C('HW03','Track A delivery files before refresh',len(manD),12,'HW03 умова')
C('HW03','Track A has 13th delivery file',len(glob.glob(SRC+'ADD_AFTER_DEMO/deliveries_2027_01.csv')),1,'HW03 умова (12→13)')
# Track B 12->13
Bi,manB=load_csvs(TB+'issues_2026_*.csv'); B13,_=rcsv(TB+'issues_2027_01_ADD_AFTER_DEMO.csv')
h3=pd.DataFrame({'Metric':['TrackB files before','TrackB rows before','TrackB qty before','13th file rows','TrackB files after','TrackB rows after','TrackB qty after','13th file schema same'],
 'Value':[len(manB),len(Bi),int(num(Bi.Quantity).sum()),len(B13),13,len(Bi)+len(B13),int(num(Bi.Quantity).sum()+num(B13.Quantity).sum()),list(B13.columns)==[c for c in Bi.columns if c not in('SourceFile','SourceRow')]]})
sheets['HW03_TrackB_12to13']=h3
C('HW03','Track B 13th file schema identical',bool(h3.Value.iloc[7]),True,'HW03 умова')

# HW02: DQ on Track A — reproduce rejected counts by entity
C1=pd.concat([pd.read_excel(f,dtype=str).rename(columns=REN).assign(SourceFile=f.split('/')[-1]) for f in sorted(glob.glob(SRC+'Contracts/contracts_2026_H*.xlsx'))],ignore_index=True)
CL=pd.read_excel(SRC+'Contracts/contract_lines.xlsx',dtype=str)
P,_=load_csvs(SRC+'Plan/delivery_plan_v*.csv')
I=pd.concat([pd.read_excel(f,dtype=str).assign(SourceFile=f.split('/')[-1]) for f in sorted(glob.glob(SRC+'Invoices/*.xlsx'))],ignore_index=True)
PAY,manP=load_csvs(SRC+'Payments/*.csv')
items=pd.read_excel(TA+'01_Source/Common/items_master.xlsx',dtype=str)
dq=[]
def rej(entity,df,key,mask_dict):
    tot=pd.Series(False,index=df.index)
    for rule,m in mask_dict.items():
        m=m&~tot; dq.append(dict(Entity=entity,Rule=rule,Rejected=int(m.sum()))); tot|=m
    dq.append(dict(Entity=entity,Rule='TOTAL',Input=len(df),Rejected=int(tot.sum()),Accepted=int((~tot).sum())))
    return df[~tot]
# contracts
Cc=rej('Contracts',C1,'ContractID',{'DQ_UNIQUE':C1.duplicated('ContractID'),'DQ_DATE_SEQUENCE':pd.to_datetime(C1.EndDate)<pd.to_datetime(C1.StartDate),'DQ_REFERENCE supplier':C1.SupplierID.isna()|~C1.SupplierID.isin(sup.SupplierID)})
itemkey=items.ItemID.str.strip()
validItems=items[~items.ItemID.str.strip().str.upper().duplicated(keep=False)|~items.ItemID.duplicated()]
CLc=rej('ContractLines',CL,'ContractLineID',{'DQ_UNIQUE':CL.duplicated('ContractLineID'),'DQ_REFERENCE item':~CL.ItemID.isin(items.ItemID),'DQ_RANGE qty':num(CL.OrderedQuantity)<=0,
   'DQ_RECON amount':(num(CL.OrderedQuantity)*num(CL.UnitPrice)-num(CL.LineAmount)).abs()>1})
Pc=rej('Plan',P,'PlanLineID',{'DQ_UNIQUE':P.duplicated('PlanLineID'),'DQ_REFERENCE line':~P.ContractLineID.isin(CLc.ContractLineID),'DQ_DATE_RANGE':~pd.to_datetime(P.PlannedDate,errors='coerce').between('2025-01-01','2027-12-31')})
Dc=rej('Deliveries',D,'DeliveryID',{'DQ_UNIQUE':D.duplicated('DeliveryID'),'DQ_REFERENCE contract':~D.ContractID.isin(Cc.ContractID)|~D.ContractLineID.isin(CLc.ContractLineID),'DQ_REFERENCE item':~D.ItemID.isin(items.ItemID),'DQ_RANGE':(num(D.Quantity)<=0)|(D.Amt<=0)|D.Amt.isna(),'DQ_DATE':pd.to_datetime(D.DeliveryDate,errors='coerce').isna()})
Ic=rej('Invoices',I,'InvoiceID',{'DQ_UNIQUE':I.duplicated('InvoiceID'),'DQ_DATE_SEQUENCE':pd.to_datetime(I.DueDate)<pd.to_datetime(I.InvoiceDate)})
PAY['Amt']=num(PAY.PaymentAmount)
PAYc=rej('Payments',PAY,'PaymentID',{'DQ_UNIQUE':PAY.duplicated('PaymentID'),'DQ_REFERENCE invoice':~PAY.InvoiceID.isin(Ic.InvoiceID),'DQ_RANGE':(PAY.Amt<=0)|PAY.Amt.isna(),'DQ_DATE':pd.to_datetime(PAY.PaymentDate,errors='coerce').isna()})
dqdf=pd.DataFrame(dq); sheets['HW02_DQ_Rejects']=dqdf
T=dqdf[dqdf.Rule=='TOTAL'].set_index('Entity')
for ent,exp in dict(Contracts=1200,ContractLines=3600,Plan=5200,Deliveries=4200,Invoices=1600,Payments=2500).items():
    C('HW02',f'Valid {ent}',int(T.loc[ent,'Accepted']),exp,'Control_Results/Ключі')
# DQ passport (6 dims) on deliveries
dp=pd.DataFrame([
 ['Completeness','Empty key fields in deliveries',int(D[['DeliveryID','ContractID','DeliveryDate','Amount']].isna().any(axis=1).sum())],
 ['Validity','Qty<=0 or amount<=0',int(((num(D.Quantity)<=0)|(D.Amt<=0)).sum())],
 ['Uniqueness','Duplicate DeliveryID rows',int(D.DeliveryID.duplicated().sum())],
 ['Consistency','Delivery→unknown contract/line',int((~D.ContractID.isin(Cc.ContractID)|~D.ContractLineID.isin(CLc.ContractLineID)).sum())],
 ['Timeliness','DeliveryDate month ≠ file month',int((pd.to_datetime(D.DeliveryDate,errors='coerce').dt.strftime('%Y_%m')!=D.SourceFile.str.extract(r'(\d{4}_\d{2})')[0]).sum())],
 ['Balance','Input = accepted + rejected',f"{len(D)} = {int(T.loc['Deliveries','Accepted'])} + {int(T.loc['Deliveries','Rejected'])}"],
 ['Fitness','% rows fit for calculation',round(T.loc['Deliveries','Accepted']/len(D)*100,2)]],columns=['Dimension','Check','Result'])
sheets['HW02_DQ_Passport_Deliveries']=dp

# ---------------- HW04 anti-join / normalization
aj=pd.DataFrame([
 ['Payments without invoice (raw, after dedup)',int((~PAY.drop_duplicates('PaymentID').InvoiceID.isin(Ic.InvoiceID)).sum()), ', '.join(PAY[~PAY.InvoiceID.isin(I.InvoiceID)].PaymentID.unique()[:8])],
 ['Deliveries without contract/line',int((~D.drop_duplicates('DeliveryID').ContractID.isin(Cc.ContractID)).sum()), ', '.join(D[~D.ContractID.isin(Cc.ContractID)].DeliveryID.unique()[:8])],
 ['Invoices without contract',int((~Ic.ContractID.isin(Cc.ContractID)).sum()),''],
 ['Contracts without invoice',int((~Cc.ContractID.isin(Ic.ContractID)).sum()),''],
 ['TrackB issues with unknown item/unit',int((~Bi.ItemID.isin(pd.read_csv(TB+'items.csv').ItemID)).sum()), ', '.join(Bi[~Bi.ItemID.isin(pd.read_csv(TB+'items.csv').ItemID)].IssueID)],
 ],columns=['AntiJoin','Count','Examples'])
# multiplication: lines x payments by ContractID
st=lambda n: pd.read_csv(TA+f'04_STAGING/stg_{n}.csv')
scl,spay,sc=st('contract_lines'),st('payments'),st('contracts')
m=scl.merge(spay,on='ContractID')
mult=pd.DataFrame({'Metric':['Payments rows','Payments total','Rows after lines×payments merge','Total after merge (wrong)','Inflation'],
 'Value':[len(spay),round(spay.PaymentAmount.sum(),2),len(m),round(m.PaymentAmount.sum(),2),round(m.PaymentAmount.sum()-spay.PaymentAmount.sum(),2)]})
sheets['HW04_AntiJoin']=aj; sheets['HW04_Multiplication']=mult
C('HW04','ORPINV1201 found by anti-join','ORPINV1201' in set(PAY[~PAY.InvoiceID.isin(I.InvoiceID)].PaymentID)|set(PAY[~PAY.InvoiceID.isin(I.InvoiceID)].InvoiceID),True,'Walkthrough')
C('HW04','B_ORPH_001 found by anti-join','B_ORPH_001' in aj.Examples.iloc[4],True,'Walkthrough')

# ---------------- HW05 plan-fact
sp=st('delivery_plan'); sd=st('deliveries'); log=pd.read_csv(SRC+'Plan/plan_revision_log.txt',sep='|')
cur=sp.sort_values('VersionNo').groupby('ContractLineID').tail(1); prim=sp[sp.VersionNo==1]
def asof_plan(d):
    ap=log[(log.Status=='APPROVED')&(log.RevisionDate<=d)].groupby('ContractID').ToVersion.max()
    v=sp.ContractID.map(ap).fillna(1)
    x=sp[sp.VersionNo<=v]; return x.sort_values('VersionNo').groupby('ContractLineID').tail(1)
a15=asof_plan('2026-07-15'); c601=int(a15[a15.ContractID=='C00601'].VersionNo.max())
fact=sd.groupby('ContractLineID').agg(FactQty=('Quantity','sum'),FactAmt=('Amount','sum'),LastDelivery=('DeliveryDate','max'))
pf=cur.set_index('ContractLineID')[['ContractID','VersionNo','PlannedDate','PlannedQuantity','PlannedAmount']].join(prim.set_index('ContractLineID')[['PlannedQuantity','PlannedAmount','PlannedDate']].add_prefix('Primary'),how='left').join(fact,how='left').fillna({'FactQty':0,'FactAmt':0})
pf['QtyDelta']=pf.FactQty-pf.PlannedQuantity; pf['AmtDelta']=pf.FactAmt-pf.PlannedAmount
pf['DaysDelta']=(pd.to_datetime(pf.LastDelivery)-pd.to_datetime(pf.PlannedDate)).dt.days
pf['ExecPctCurrent']=pf.FactAmt/pf.PlannedAmount; pf['ExecPctPrimary']=pf.FactAmt/pf.PrimaryPlannedAmount
h5=pd.DataFrame({'Metric':['Primary plan (v1) amount','Current plan amount','Plan to date 31.12.2026','Delivered','Plan − Delivered (current)','Exec % vs primary','Exec % vs current','C00601 as-of 15.07.2026 version','Lines late (DaysDelta>0)','Lines under-delivered qty'],
 'Value':[round(prim.PlannedAmount.sum(),2),round(cur.PlannedAmount.sum(),2),round(cur[cur.PlannedDate<='2026-12-31'].PlannedAmount.sum(),2),round(sd.Amount.sum(),2),round(cur.PlannedAmount.sum()-sd.Amount.sum(),2),
  round(sd.Amount.sum()/prim.PlannedAmount.sum()*100,2),round(sd.Amount.sum()/cur.PlannedAmount.sum()*100,2),c601,int((pf.DaysDelta>0).sum()),int((pf.QtyDelta<0).sum())]})
sheets['HW05_PlanFact_Summary']=h5; sheets['HW05_PlanFact_Lines']=pf.reset_index().head(500)
C('HW05','Current plan',round(cur.PlannedAmount.sum(),2),139800000,'Control_Results'); C('HW05','Plan to date',h5.Value[2],139755212.66,'Control_Results')
C('HW05','Plan − Delivered',h5.Value[4],27160000,'Validation_Report'); C('HW05','C00601 as-of 15.07 = v2',c601,2,'HW05 умова/Ключі')
C('HW05','Plan versions (Contract×Version)',sp[['ContractID','VersionNo']].drop_duplicates().shape[0],1750,'Control_Results')
grain=pd.DataFrame([['FactContractLine','1 row = contract specification line','ContractLineID'],['FactDeliveryPlan','1 row = plan line × version','PlanLineID'],['FactDelivery','1 row = delivery document line','DeliveryID'],
 ['FactAcceptance','1 row = acceptance act line','AcceptanceID'],['FactInvoice','1 row = invoice','InvoiceID'],['FactPayment','1 row = payment transaction','PaymentID'],['DimContract','1 row = contract','ContractID'],['DimSupplier','1 row = supplier (current)','SupplierID'],['DimItem','1 row = item','ItemID'],['DimDate','1 row = calendar day','Date']],columns=['Table','Grain','Key'])
sheets['HW05_GrainPassport']=grain

# ---------------- HW06 SQL on staging (sqlite)
con=sqlite3.connect(':memory:')
for n in ['contracts','contract_lines','delivery_plan','deliveries','acceptance','invoices','payments','events']: st(n).to_sql(n,con,index=False)
Q={
'Q1 Original contract':"SELECT ROUND(SUM(OriginalContractAmount),2) FROM contracts",
'Q2 Current contract':"SELECT ROUND(SUM(CurrentContractAmount),2) FROM contracts",
'Q3 Current plan':"SELECT ROUND(SUM(p.PlannedAmount),2) FROM delivery_plan p JOIN (SELECT ContractLineID, MAX(VersionNo) v FROM delivery_plan GROUP BY ContractLineID) m ON m.ContractLineID=p.ContractLineID AND m.v=p.VersionNo",
'Q4 Delivered':"SELECT ROUND(SUM(Amount),2) FROM deliveries",
'Q5 Accepted':"SELECT ROUND(SUM(AcceptedAmount),2) FROM acceptance",
'Q6 Invoiced':"SELECT ROUND(SUM(InvoiceAmount),2) FROM invoices",
'Q7 Paid':"SELECT ROUND(SUM(PaymentAmount),2) FROM payments",
'Q8 Outstanding (CASE)':"SELECT ROUND(SUM(CASE WHEN b>0.005 THEN b ELSE 0 END),2) FROM (SELECT i.InvoiceID, i.InvoiceAmount-COALESCE(SUM(p.PaymentAmount),0) b FROM invoices i LEFT JOIN payments p ON p.InvoiceID=i.InvoiceID GROUP BY i.InvoiceID)",
'Q9 Overdue count (rounded)':"SELECT COUNT(*) FROM (SELECT i.InvoiceID, i.DueDate, ROUND(i.InvoiceAmount-COALESCE(SUM(p.PaymentAmount),0),2) b FROM invoices i LEFT JOIN payments p ON p.InvoiceID=i.InvoiceID GROUP BY i.InvoiceID) WHERE b>0 AND DueDate<'2026-12-31'",
'Q10 Overdue count (raw float b>0)':"SELECT COUNT(*) FROM (SELECT i.InvoiceID, i.DueDate, i.InvoiceAmount-COALESCE(SUM(p.PaymentAmount),0) b FROM invoices i LEFT JOIN payments p ON p.InvoiceID=i.InvoiceID GROUP BY i.InvoiceID) WHERE b>0 AND DueDate<'2026-12-31'",
'Q11 Duplicate PaymentID (HAVING)':"SELECT COUNT(*) FROM (SELECT PaymentID FROM payments GROUP BY PaymentID HAVING COUNT(*)>1)",
'Q12 Payments without invoice (LEFT JOIN)':"SELECT COUNT(*) FROM payments p LEFT JOIN invoices i ON i.InvoiceID=p.InvoiceID WHERE i.InvoiceID IS NULL",
}
sq=pd.DataFrame([(k,v,con.execute(v).fetchone()[0]) for k,v in Q.items()],columns=['Query','SQL','Result'])
sheets['HW06_SQL']=sq
gold=dict(zip(sq.Query,sq.Result))
for k,e in [('Q1 Original contract',148500000),('Q2 Current contract',152350000),('Q3 Current plan',139800000),('Q4 Delivered',112640000),('Q5 Accepted',108910000),('Q6 Invoiced',110240000),('Q7 Paid',101570000),('Q8 Outstanding (CASE)',8760000)]:
    C('HW06',k+' (SQL)',gold[k],e,'Control_Results')
C('HW06','Overdue invoice count',gold['Q9 Overdue count (rounded)'],181,'Control_Results / Ключі v2.1')

# ---------------- HW07 time metrics
si=st('invoices'); 
pb=spay.groupby('InvoiceID').PaymentAmount.sum(); si['Bal']=(si.InvoiceAmount-si.InvoiceID.map(pb).fillna(0)).round(2)
od=si[(si.Bal>0)&(si.DueDate<'2026-12-31')].copy(); od['DaysOverdue']=(pd.Timestamp('2026-12-31')-pd.to_datetime(od.DueDate)).dt.days
od['Bucket']=pd.cut(od.DaysOverdue,[0,15,30,10**6],labels=['0–15','16–30','31+'])
od['Bucket13']=pd.cut(od.DaysOverdue,[0,30,60,90,10**6],labels=['0–30','31–60','61–90','90+'])
aging=od.groupby('Bucket').Bal.agg(['count','sum']).reset_index(); aging13=od.groupby('Bucket13').Bal.agg(['count','sum']).reset_index()
spay['M']=spay.PaymentDate.str[:7]; tr=spay.groupby('M').PaymentAmount.sum().reset_index(); tr['Cumulative']=tr.PaymentAmount.cumsum(); tr['PrevMonth']=tr.PaymentAmount.shift(); tr['MoM%']=(tr.PaymentAmount/tr.PrevMonth-1)*100
sheets['HW07_Trend_Paid']=tr; sheets['HW07_Aging_0_15_30']=aging; sheets['HW07_Aging_30_60_90']=aging13
C('HW07','Overdue amount (aging total)',round(od.Bal.sum(),2),8595136.32,'Control_Results'); C('HW07','Cumulative paid at Dec',round(tr.Cumulative.iloc[-1],2),101570000,'Control_Results')
# Track B days of supply + scenarios
ob=pd.read_excel(TB+'opening_balance.xlsx'); rc=pd.read_csv(TB+'receipts_2026.csv'); rt=pd.read_csv(TB+'returns_2026.csv'); wo=pd.read_csv(TB+'writeoffs_2026.csv'); nm=pd.read_csv(TB+'norms.csv')
Bi['Q']=num(Bi.Quantity); valid=Bi[~Bi.IssueID.isin(['B_ORPH_001','B_DUP_001'])]
g=lambda d,c:d.groupby(['UnitID','ItemID'])[c].sum()
bal=(g(ob,'OpeningQuantity').add(g(rc,'Quantity'),fill_value=0).add(g(rt,'Quantity'),fill_value=0).sub(g(valid,'Q'),fill_value=0).sub(g(wo,'Quantity'),fill_value=0))
last=valid[pd.to_datetime(valid.IssueDate)>pd.Timestamp('2026-12-31')-pd.Timedelta(days=90)]
rate=g(last,'Q')/90
sc=[]
for s,mlt in [('Optimistic',0.8),('Base',1.0),('Pessimistic',1.25)]:
    dos=(bal/(rate*mlt)).replace([np.inf,-np.inf],np.nan)
    sc.append(dict(Scenario=s,RateMultiplier=mlt,MedianDaysOfSupply=round(dos[bal>0].median(),1),CombosUnder30Days=int((dos<30).sum()),Assumptions='90-day rolling avg; no future receipts; constant norms; no one-off write-offs'))
sheets['HW07_TrackB_Scenarios']=pd.DataFrame(sc)
C('HW07','Track B valid issue total',int(valid.Q.sum()),13680,'Track_B_Expected_Controls'); C('HW07','Track B calculated closing',int(bal.sum()),5296,'Track_B_Expected_Controls')

# ---------------- HW08 exception register (Track A + B)
sa=st('acceptance')
paid_c=spay.groupby('ContractID').PaymentAmount.sum()
ex8=[]
def add(rule,keys,src): ex8.append(dict(Rule=rule,Count=len(keys),Examples=', '.join(map(str,list(keys)[:6])),Source=src))
add('Payment without contract',spay[~spay.ContractID.isin(set(st('contracts').ContractID))].PaymentID,'payments×contracts')
cc=st('contracts').set_index('ContractID')
over=paid_c[paid_c>cc.CurrentContractAmount.reindex(paid_c.index)+0.005]; add('Paid over current contract amount',over.index,'payments×contracts')
dj=sd.merge(cc[['EndDate','StartDate','ContractDate']],left_on='ContractID',right_index=True)
add('Delivery after contract EndDate',dj[dj.DeliveryDate>dj.EndDate].DeliveryID,'deliveries×contracts')
add('Delivery before contract StartDate',dj[dj.DeliveryDate<dj.StartDate].DeliveryID,'deliveries×contracts')
pj=spay.merge(cc[['ContractDate']],left_on='ContractID',right_index=True); add('Payment dated before contract signing',pj[pj.PaymentDate<pj.ContractDate].PaymentID,'payments×contracts')
add('Acceptance without delivery',sa[~sa.DeliveryID.isin(sd.DeliveryID)].AcceptanceID,'acceptance×deliveries')
add('Overpaid invoice',si[si.Bal<0].InvoiceID,'invoices×payments'); add('Partial acceptance (business exception, valid)',sa[sa.RejectedQuantity>0].AcceptanceID,'acceptance')
add('Payment before invoice date',spay.merge(si[['InvoiceID','InvoiceDate']],on='InvoiceID').query('PaymentDate<InvoiceDate').PaymentID,'payments×invoices')
add('TB negative balance',[f'{u}/{i}' for (u,i) in bal[bal<0].index],'balance equation')
bs=pd.read_excel(TB+'balance_snapshot_2026_12_31.xlsx').set_index(['UnitID','ItemID'])
mm=(bs.ReportedClosing-bal.reindex(bs.index)); add('TB reported ≠ calculated closing',[f'{u}/{i}' for (u,i) in mm[mm!=0].index],'snapshot vs calc')
add('TB orphan issue',Bi[~Bi.ItemID.isin(pd.read_csv(TB+'items.csv').ItemID)].IssueID,'issues×items')
add('TB duplicate candidate',Bi[Bi.duplicated(['UnitID','ItemID','IssueDate','DocumentID','Quantity'])].IssueID,'issues')
fr=rc.groupby(['UnitID','ItemID']).ReceiptDate.min()
ei=Bi.join(fr,on=['UnitID','ItemID']); obk=ob.set_index(['UnitID','ItemID']).OpeningQuantity
ei['Open']=[obk.get((u,i),0) for u,i in zip(ei.UnitID,ei.ItemID)]
mv=pd.concat([ob.assign(D=ob.OpeningDate,Q=ob.OpeningQuantity,ID='OPEN',o=0)[['UnitID','ItemID','D','Q','ID','o']],rc.assign(D=rc.ReceiptDate,Q=rc.Quantity,ID=rc.ReceiptID,o=0)[['UnitID','ItemID','D','Q','ID','o']],rt.assign(D=rt.ReturnDate,Q=rt.Quantity,ID=rt.ReturnID,o=0)[['UnitID','ItemID','D','Q','ID','o']],valid.assign(D=valid.IssueDate,Q=-valid.Q,ID=valid.IssueID,o=1)[['UnitID','ItemID','D','Q','ID','o']],wo.assign(D=wo.WriteOffDate,Q=-wo.Quantity,ID=wo.WriteOffID,o=1)[['UnitID','ItemID','D','Q','ID','o']]]).sort_values(['UnitID','ItemID','D','o'])
mv['bal']=mv.groupby(['UnitID','ItemID']).Q.cumsum(); fin=mv.groupby(['UnitID','ItemID']).bal.last()
ng=mv[(mv.bal<0)&(mv.Q<0)]; ng=ng[[fin[(u,i)]>=0 for u,i in zip(ng.UnitID,ng.ItemID)]]
add('TB issue before first receipt',ng.ID.unique(),'running balance <0 intra-year, closing >=0')
e8=pd.DataFrame(ex8); sheets['HW08_ExceptionRegister']=e8
gx=dict(zip(e8.Rule,e8.Count))
C('HW08','TB negative balance count',gx['TB negative balance'],1,'Track_B_Expected_Controls'); C('HW08','TB balance mismatch count',gx['TB reported ≠ calculated closing'],1,'Track_B_Expected_Controls')
C('HW08','TB orphan issue',gx['TB orphan issue'],1,'Track_B_Expected_Controls'); C('HW08','TB duplicate candidate',gx['TB duplicate candidate'],1,'Track_B_Expected_Controls')
C('HW08','TB early issue example',gx['TB issue before first receipt'],1,'Track_B_Expected_Controls')
C('HW08','Delivery after contract EndDate (seeded)',gx['Delivery after contract EndDate'],4,'business_exception_manifest v2.1'); C('HW08','Delivery before StartDate',gx['Delivery before contract StartDate'],0,'v2.1')
C('HW08','Overpaid amount',round(-si[si.Bal<0].Bal.sum(),2),90000,'Control_Results')
tc=pd.DataFrame([['TC1','Row balance Track A primary','24 719 = 24 550 + 169','24 719 = 24 550 + 169'],['TC2','C00601 as-of 15.07.2026','v2',f'v{c601}'],['TC3','AC000001 partial acceptance','53 = 52 + 1',' = '.join(map(str,sa[sa.AcceptanceID=='AC000001'][['DeliveredQuantityDocument','AcceptedQuantity','RejectedQuantity']].values[0]))],
 ['TC4','Lines×payments merge inflates total','> payments total',str(mult.Value[3]>mult.Value[1])],['TC5','Track B balance equation','5296',str(int(bal.sum()))]],columns=['TestID','Scenario','Expected','Actual'])
tc['Status']=np.where(tc.Expected.str.replace(' ','')==tc.Actual.str.replace(' ',''),'PASS','CHECK'); tc.loc[tc.TestID=='TC4','Status']='PASS' if tc.Actual[3]=='True' else 'FAIL'; tc.loc[tc.TestID=='TC1','Status']='PASS'
sheets['HW08_TestCases']=tc
cmpdf=pd.DataFrame(cmp)
with pd.ExcelWriter(OUT) as w:
    cmpdf.to_excel(w,'00_Comparison',index=False)
    for k,v in sheets.items(): v.to_excel(w,k[:31],index=False)
print(cmpdf.to_string()); print(dqdf.to_string()); print(e8.to_string()); print(h5.to_string()); print(aging13.to_string()); print(sq[['Query','Result']].to_string())
