import sys; sys.path.insert(0,'.')
from content import L
lim=dict(title=60,plan=(6,40),result=230,d1=26,d2=26)
bad=[]
def chk(k,f,s,n):
    if len(s)>n: bad.append(f'{k}.{f}: {len(s)}>{n}: {s}')
for k,d in L.items():
    chk(k,'title',d['title'],60); chk(k,'result',d['result'],230); chk(k,'d1',d['d1'],26); chk(k,'d2',d['d2'],26)
    assert len(d['plan'])<=6; [chk(k,'plan',x,40) for x in d['plan']]
    for f in ('c1','c2','hw','res'):
        t,b=d[f]; chk(k,f+'.t',t,40); assert len(b)<=6,(k,f); [chk(k,f,x,82) for x in b]
    t,steps=d['proc']; chk(k,'proc.t',t,40); assert len(steps)==6
    for a,b in steps: chk(k,'proc.lbl',a,16); chk(k,'proc.txt',b,62)
    t,code=d['code']; chk(k,'code.t',t,40); assert len(code)<=11,(k,len(code)); [chk(k,'code',x,84) for x in code]
    assert len(d['three'])==3; [ (chk(k,'three.h',a,12),chk(k,'three.t',b,150)) for a,b in d['three']]
    assert len(d['four'])==4; [ (chk(k,'four.h',a,13),chk(k,'four.t',b,110)) for a,b in d['four']]
    t,cols=d['cols']; chk(k,'cols.t',t,40); assert len(cols)==3
    for h,ls in cols: chk(k,'cols.h',h,12); assert len(ls)<=4; [chk(k,'cols.l',x,22) for x in ls]
    t,q=d['quiz']; chk(k,'quiz.t',t,40); assert len(q)<=5; [chk(k,'quiz',x,88) for x in q]
print('\n'.join(bad)); print(len(bad),'issues')
