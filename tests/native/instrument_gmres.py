"""Instrument disposable GMRES C source; never enable QPC in release binaries."""
import argparse
from pathlib import Path
import re
import json

PROFILER=r'''
#include <windows.h>
static __declspec(thread) LARGE_INTEGER qp_freq;
static __declspec(thread) LONGLONG qp_in[6], qp_ex[6], qp_calls[6];
static __declspec(thread) LONGLONG qp_start[16], qp_child[16];
static __declspec(thread) int qp_depth, qp_id[16], qp_bad;
static void qp_enter(int id) {
    LARGE_INTEGER q;
    if (qp_depth >= 16) { qp_bad=1; return; }
    QueryPerformanceCounter(&q);
    qp_id[qp_depth]=id; qp_start[qp_depth]=q.QuadPart; qp_child[qp_depth]=0;
    qp_depth++; qp_calls[id]++;
}
static void qp_leave(int id) {
    LARGE_INTEGER q; LONGLONG elapsed;
    QueryPerformanceCounter(&q);
    if (qp_depth<1 || qp_id[qp_depth-1]!=id) { qp_bad=1; return; }
    qp_depth--; elapsed=q.QuadPart-qp_start[qp_depth];
    qp_in[id]+=elapsed; qp_ex[id]+=elapsed-qp_child[qp_depth];
    if (qp_depth) qp_child[qp_depth-1]+=elapsed;
}
__declspec(dllexport) void __stdcall sfn_gmres_profile_reset(void) {
    int i; qp_depth=0; qp_bad=0; QueryPerformanceFrequency(&qp_freq);
    for(i=0;i<6;i++) qp_in[i]=qp_ex[i]=qp_calls[i]=0;
}
__declspec(dllexport) int32_t __stdcall sfn_gmres_profile_read(double *out, int32_t n) {
    int i;
    if (!out || n<20 || qp_freq.QuadPart<=0 || qp_depth || qp_bad) return -1;
    out[0]=(double)qp_freq.QuadPart; out[1]=(double)qp_bad;
    for(i=0;i<6;i++) {
        out[2+3*i]=(double)qp_calls[i];
        out[3+3*i]=(double)qp_in[i]/(double)qp_freq.QuadPart;
        out[4+3*i]=(double)qp_ex[i]/(double)qp_freq.QuadPart;
    }
    return 0;
}
'''

def instrument(source):
    # Process from the last function to the first to keep original offsets valid.
    names=[('dot','double'),('axpy','void'),('csr_matvec','void'),('csr_resnorm','double'),('ilu_apply','int')]
    if 'static int32_t gmres_impl(' in source:
        names += [('csr_matvec_tuned','void'),('ilu_apply_tuned','int')]
        if 'static int ilu_apply_prepared(' in source: names += [('ilu_apply_prepared','int')]
        names += [('gmres_impl','int32_t')]
    else:
        names += [('sfn_gmres','int32_t')]
    spans=[]
    for id,(name,kind) in enumerate(names):
        match=re.search(r'\b'+name+r'\([^;{}]*\)\s*\{',source)
        if not match: raise ValueError(name)
        start=match.end(); depth=1; end=start
        while depth:
            if source[end]=='{': depth+=1
            elif source[end]=='}': depth-=1
            end+=1
        spans.append((start,end-1,id,kind))
    for start,end,id,kind in sorted(spans,reverse=True):
        body=source[start:end]
        if kind=='void':
            body=re.sub(r'\breturn\s*;',f'{{ qp_leave({id}); return; }}',body)
            body+=f'\n qp_leave({id});\n'
        else:
            body=re.sub(r'\breturn\s+([^;]+);',lambda m:'{ '+kind+' qp_result=('+m[1]+f'); qp_leave({id}); return qp_result; }}',body)
        source=source[:start]+f'\n qp_enter({id});\n'+body+source[end:]
    profiler=PROFILER.replace('[6]',f'[{len(names)}]').replace('i<6',f'i<{len(names)}').replace('n<20',f'n<{2+3*len(names)}')
    return source.replace('#pragma fp_contract(off)','#pragma fp_contract(off)\n'+profiler),[n for n,k in names]

def diagnose_zero_rows(source):
    """Three spare doubles report a possible exact-zero-row failure bound."""
    marker='static __declspec(thread) LARGE_INTEGER qp_freq;'
    assert marker in source
    extra=r'''
static __declspec(thread) double qp_zero_calls,qp_zero_rows,qp_zero_ratio;
static void qp_zero_diagnose(int32_t n,const int32_t *rp,const double *A,const double *F,double tol) {
    int32_t r,p;double total=0.0,lower=0.0,ratio;int zeros=0;
    for(r=0;r<n;r++) {
        int zero=1;
        if(!isfinite(F[r])) return;
        total+=F[r]*F[r];
        for(p=rp[r];p<rp[r+1];p++) if(A[p]!=0.0) {zero=0;break;}
        if(zero) {lower+=F[r]*F[r];zeros++;}
    }
    qp_zero_rows+=zeros;
    if(total>0.0 && isfinite(total)) {
        ratio=sqrt(lower)/sqrt(total);
        if(ratio>qp_zero_ratio) qp_zero_ratio=ratio;
        if(ratio>tol) qp_zero_calls++;
    }
}
'''
    source=source.replace(marker,marker+extra,1)
    source=source.replace('int i; qp_depth=0; qp_bad=0;', 'int i; qp_depth=0; qp_bad=0; qp_zero_calls=qp_zero_rows=qp_zero_ratio=0.0;',1)
    source=source.replace('  if(inverse) for(i=0;i<n;i++) {','  qp_zero_diagnose(n,rp,A,F,dtol);\n  if(inverse) for(i=0;i<n;i++) {',1)
    marker='    return 0;\n}\n'
    assert marker in source
    source=source.replace(marker,'    out[29]=qp_zero_calls;out[30]=qp_zero_rows;out[31]=qp_zero_ratio;\n'+marker,1)
    source=source.replace('n<29','n<32',1)
    return source

if __name__=='__main__':
    ap=argparse.ArgumentParser(); ap.add_argument('--src',default='native/sfn_gmres.c'); ap.add_argument('--out',required=True);ap.add_argument('--diagnose',action='store_true');a=ap.parse_args()
    p=Path(a.out)
    if p.exists(): raise ValueError('Preserve existing source')
    p.parent.mkdir(parents=True,exist_ok=True)
    source,names=instrument(Path(a.src).read_text(encoding='utf-8-sig'))
    if a.diagnose:
        assert len(names)==9,'Diagnosis needs three spare doubles in the 32-double harness'
        source=diagnose_zero_rows(source)
    p.write_text(source,encoding='utf-8',newline='\n')
    p.with_suffix('.profile.json').write_text(json.dumps(names,indent=2)+'\n',encoding='utf-8')
