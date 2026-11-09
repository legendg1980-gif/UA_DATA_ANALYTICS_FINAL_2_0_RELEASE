import json, sys; sys.path.insert(0,'.')
from content import L
P='g32a47ba9a3d_0_'
ORDER=[('S01','5'),('S02','12'),('S03','27'),('S04','183'),('S05','71'),('S06','83'),('S07','90'),('S08','83'),('S09','160'),('S10','71'),('S11','244'),('S12','266'),('S13','167'),('S14','77'),('S15','83'),('S16','83')]
TEMPLATE_SLIDES=['0','5','12','22','27','32','37','65','71','77','83','90','112','134','152','160','167','183','189','205','219','229','234','239','244','259','266','281','289','465','417','330','347','362']
# element role maps: template element suffix -> role
ROLES={'5':{'6':'t','7':'n','8':'c'},'27':{'29':'b'},'183':{'185':'h','186':'b'},'71':{'73':'h','74':'s'},
 '83':{'86':'t','85':'b'},'90':{'99':'t','98':'L1','101':'L2','103':'L3','105':'L4','107':'L5','109':'L6','97':'X1','100':'X2','102':'X3','104':'X4','106':'X5','108':'X6'},
 '160':{'164':'t','163':'b','162':'img'},'244':{'249':'H1','250':'H2','251':'H3','248':'X1','252':'X2','253':'X3'},
 '266':{'270':'H1','276':'H2','274':'H3','278':'H4','269':'X1','275':'X2','273':'X3','277':'X4'},
 '167':{'180':'t','174':'H1','175':'H2','176':'H3','177':'X1','178':'X2','179':'X3'},'77':{'79':'t','78':'b'}}
FIXED={('S04','h'):'РЕЗУЛЬТАТ ЗАНЯТТЯ',('S05','h'):'ЧАСТИНА 1',('S10','h'):'ЧАСТИНА 2',('S01','c'):'UA DATA ANALYTICS'}
BULLETS=['S03_b','S06_b','S08_b','S14_b','S15_b','S16_b']
def skeleton():
    req=[]
    for sid,src in reversed(ORDER):  # reverse so copies end in order after... we move later anyway
        ids={P+src:sid+'_pg'}  # slide IDs must be >=5 chars
        for e,r in ROLES.get(src,{}).items(): ids[P+e]=f'{sid}_{r}'
        req.append({'duplicateObject':{'objectId':P+src,'objectIds':ids}})
    for sid,src in ORDER:
        for e,r in ROLES.get(src,{}).items():
            oid=f'{sid}_{r}'
            if r=='img': req.append({'deleteObject':{'objectId':oid}}); continue
            txt=FIXED.get((sid,r),'{{'+oid+'}}')
            req.append({'deleteText':{'objectId':oid,'textRange':{'type':'ALL'}}})
            req.append({'insertText':{'objectId':oid,'text':txt,'insertionIndex':0}})
    for b in BULLETS:
        req.append({'createParagraphBullets':{'objectId':b,'textRange':{'type':'ALL'},'bulletPreset':'NUMBERED_DIGIT_ALPHA_ROMAN' if b=='S14_b' else 'BULLET_DISC_CIRCLE_SQUARE'}})
    pt=lambda n:{'magnitude':n,'unit':'PT'}
    req.append({'updateTextStyle':{'objectId':'S01_t','textRange':{'type':'ALL'},'style':{'fontSize':pt(36)},'fields':'fontSize'}})
    req.append({'updateTextStyle':{'objectId':'S04_b','textRange':{'type':'ALL'},'style':{'fontSize':pt(18)},'fields':'fontSize'}})
    req+=template_fixes()
    for s in TEMPLATE_SLIDES: req.append({'deleteObject':{'objectId':P+s}})
    for i,(sid,_) in enumerate(ORDER): req.append({'updateSlidesPosition':{'slideObjectIds':[sid+'_pg'],'insertionIndex':i}})
    return req
pt=lambda n:{'magnitude':n,'unit':'PT'}
def _size(oid,n): return {'updateTextStyle':{'objectId':oid,'textRange':{'type':'ALL'},'style':{'fontSize':pt(n)},'fields':'fontSize'}}
def template_fixes():
    # template inconsistencies found on render check (2026-11-09):
    # S07 first step label is 14pt regular while the other five are 16pt bold;
    # S13 only the first column has bullets, but cols lines are wrapped phrases -> no bullets anywhere.
    r=[{'updateTextStyle':{'objectId':'S07_L1','textRange':{'type':'ALL'},'style':{'fontSize':pt(16),'bold':True},'fields':'fontSize,bold'}},
       {'deleteParagraphBullets':{'objectId':'S13_X1','textRange':{'type':'ALL'}}}]
    return r
def _lines(t,w):
    n,cur=1,0
    for word in t.split():
        add=len(word)+(1 if cur else 0)
        if cur+add>w: n+=1; cur=len(word)
        else: cur+=add
    return n
def fit_styles(n):
    # IBM Plex Mono: ~0.6em per char. Title boxes ~675-683pt wide -> 37 chars at 30pt, 43 at 26pt.
    d=L[n]; r=[]
    # cover title: 596pt box holds 2 lines; 27 chars/line at 36pt, 33 at 30pt
    if _lines(d['title'],27)>2: r.append(_size('S01_t',30 if _lines(d['title'],33)<=2 else 26))
    for sid,f in [('S06','c1'),('S08','c2'),('S09','code'),('S13','cols'),('S14','quiz'),('S15','hw'),('S16','res'),('S07','proc')]:
        t=d[f][0]
        if sid=='S07' and len(t)<=25: continue          # keep template 44pt when it fits
        r.append(_size(sid+'_t',30 if len(t)<=36 else 26))
    for i,(a,_) in enumerate(d['three'],1):              # 210pt box: 11 chars at 30pt
        if len(a)>11: r.append(_size(f'S11_H{i}',26))
    for i,(a,_) in enumerate(d['four'],1):               # 150pt box: 12 chars at 19pt
        if len(a)>12: r.append(_size(f'S12_H{i}',17))
    return r
def deck(n):
    d=L[n]; m={}
    m['S01_t']=d['title']; m['S01_n']=f'Заняття {n}'
    m['S03_b']='\n'.join(d['plan']); m['S04_b']=d['result']; m['S05_s']=d['d1']; m['S10_s']=d['d2']
    for sid,f in [('S06','c1'),('S08','c2'),('S15','hw'),('S16','res'),('S14','quiz')]:
        t,b=d[f]; m[sid+'_t']=t; m[sid+'_b']='\n'.join(b)
    t,st=d['proc']; m['S07_t']=t
    for i,(a,b) in enumerate(st,1): m[f'S07_L{i}']=a; m[f'S07_X{i}']=b
    t,c=d['code']; m['S09_t']=t; m['S09_b']='\n'.join(c)
    for i,(a,b) in enumerate(d['three'],1): m[f'S11_H{i}']=a; m[f'S11_X{i}']=b
    for i,(a,b) in enumerate(d['four'],1): m[f'S12_H{i}']=a; m[f'S12_X{i}']=b
    t,cols=d['cols']; m['S13_t']=t
    for i,(a,ls) in enumerate(cols,1): m[f'S13_H{i}']=a; m[f'S13_X{i}']='\n'.join(ls)
    return [{'replaceAllText':{'containsText':{'text':'{{'+k+'}}'},'replaceText':v}} for k,v in m.items()]+fit_styles(n)
if __name__=='__main__':
    if sys.argv[1]=='skel': print(json.dumps(skeleton(),ensure_ascii=False))
    elif sys.argv[1]=='fix': print(json.dumps(template_fixes(),ensure_ascii=False))  # for decks copied before the fix
    else: print(json.dumps(deck(int(sys.argv[1])),ensure_ascii=False,separators=(',',':')))
