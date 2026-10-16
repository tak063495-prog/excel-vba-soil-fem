; x87 kernels that reproduce VBA (x64) Double expression evaluation:
; the expression is evaluated on the x87 stack with the caller's control word
; and rounded to Double only when stored, exactly like the VBA statement.
; Leaf functions: no stack or nonvolatile register use.
; Strides are in bytes. SPDX-License-Identifier: Apache-2.0
.code
; void sfa_axpy_sub(int64 n, double* y, int64 ys, const double* x, int64 xs, const double* s)
;   y[t] = y[t] - s * x[t]
sfa_axpy_sub PROC
    test rcx,rcx
    jle axs_done
    mov r10,qword ptr [rsp+40]
    mov r11,qword ptr [rsp+48]
axs_next:
    fld qword ptr [r11]
    fmul qword ptr [r9]
    fsubr qword ptr [rdx]
    fstp qword ptr [rdx]
    add rdx,r8
    add r9,r10
    dec rcx
    jnz axs_next
axs_done:
    ret
sfa_axpy_sub ENDP
; void sfa_dot_add(int64 n, const double* a, int64 as, const double* b, int64 bs, double* s)
;   s = s + a[t] * b[t]   (stored every step)
sfa_dot_add PROC
    test rcx,rcx
    jle dta_done
    mov r10,qword ptr [rsp+40]
    mov r11,qword ptr [rsp+48]
dta_next:
    fld qword ptr [rdx]
    fmul qword ptr [r9]
    fadd qword ptr [r11]
    fstp qword ptr [r11]
    add rdx,r8
    add r9,r10
    dec rcx
    jnz dta_next
dta_done:
    ret
sfa_dot_add ENDP
; void sfa_dot_sub(...) : s = s - a[t] * b[t]
sfa_dot_sub PROC
    test rcx,rcx
    jle dts_done
    mov r10,qword ptr [rsp+40]
    mov r11,qword ptr [rsp+48]
dts_next:
    fld qword ptr [rdx]
    fmul qword ptr [r9]
    fsubr qword ptr [r11]
    fstp qword ptr [r11]
    add rdx,r8
    add r9,r10
    dec rcx
    jnz dts_next
dts_done:
    ret
sfa_dot_sub ENDP
; void sfa_div(const double* a, const double* b, double* out) : out = a / b
sfa_div PROC
    fld qword ptr [rcx]
    fdiv qword ptr [rdx]
    fstp qword ptr [r8]
    ret
sfa_div ENDP
; void sfa_subdiv(const double* r, const double* s, const double* p, double* out) : out = r - s / p
sfa_subdiv PROC
    fld qword ptr [rdx]
    fdiv qword ptr [r8]
    fsubr qword ptr [rcx]
    fstp qword ptr [r9]
    ret
sfa_subdiv ENDP
; void sfa_mulsub(const double* y, const double* s, const double* x, double* out) : out = y - s * x
sfa_mulsub PROC
    fld qword ptr [rdx]
    fmul qword ptr [r8]
    fsubr qword ptr [rcx]
    fstp qword ptr [r9]
    ret
sfa_mulsub ENDP
END
