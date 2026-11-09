# Length checks for v2 slides (IBM Plex Mono ~0.6em/char). Run: python check2.py
import sys; sys.path.insert(0,'.')
from gen2 import texts, SPEC, _lines
bad=[]
def chk(c,msg):
    if not c: bad.append(msg)
for n in range(1,16):
    T=texts(n)
    for vid,kind,f,_ in SPEC:
        m=T[vid]
        for r,t in m.items():
            tag=f'L{n} {vid} {kind}/{r}'
            if r=='t' and kind!='cover': chk(len(t)<=43,f'{tag} title {len(t)}: {t}')
            if kind=='bul' and r=='b':
                ls=t.split('\n'); w=sum(_lines(x,66) for x in ls); chk(len(ls)<=6 and w<=8,f'{tag} {len(ls)} lines, {w} wrapped')
            if kind=='proc' and r[0]=='L': chk(len(t)<=16,f'{tag} label {len(t)}: {t}')
            if kind=='proc' and r[0]=='X': chk(_lines(t,19)<=4,f'{tag} step {len(t)}: {t}')
            if kind=='code' and r=='b':
                ls=t.split('\n'); chk(len(ls)<=12,f'{tag} {len(ls)} lines')
                for x in ls: chk(len(x)<=86,f'{tag} code {len(x)}: {x}')
            if kind=='div' and r=='s': chk(_lines(t,27)<=2,f'{tag} sub: {t}')
print('\n'.join(bad) or 'OK'); print(len(bad))
