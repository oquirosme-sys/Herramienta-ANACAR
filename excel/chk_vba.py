import re,glob,sys
for f in glob.glob('vba/*.bas'):
    L=open(f,encoding='utf8').read().split('\n')
    proc=None; names={}
    mod=set()
    for i,l in enumerate(L,1):
        s=l.strip()
        m=re.match(r'(Public |Private )?(Sub|Function) (\w+)\((.*?)\)',s)
        if m:
            proc=m.group(3); names={}
            for p in m.group(4).split(','):
                p=re.sub(r'^(Optional |ByVal |ByRef )+','',p.strip())
                n=re.match(r'(\w+)',p)
                if n: names[n.group(1).lower()]=i
            continue
        if re.match(r'End (Sub|Function)',s): proc=None; continue
        m=re.match(r'(Dim|Private|Public|Const)\s+(.*)',s)
        if m and not s.startswith(('Private Sub','Public Sub','Private Function','Public Function','Private Type','Private Const','Public Const')) or (m and 'Const' in s):
            body=m.group(2)
            body=re.sub(r'\(.*?\)','',body)
            for part in body.split(','):
                n=re.match(r'\s*(?:Const\s+)?(\w+)',part)
                if not n: continue
                k=n.group(1).lower()
                tgt=names if proc else mod
                if k in tgt: print(f, i, proc, 'DUP', k)
                if proc: names[k]=i
                else: mod.add(k)
        # name used as proc while same name as function? skip
