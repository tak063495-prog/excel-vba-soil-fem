#!/bin/sh
# Builds soilfem_mat_x64.dll with mingw-w64 GCC (x87 arithmetic, no CRT).
set -e
x86_64-w64-mingw32-gcc -O2 -std=c11 -Wall -Wextra -Werror -Wno-unused-parameter \
  -mfpmath=387 -fexcess-precision=standard -ffp-contract=off -fno-fast-math \
  -fno-builtin -fno-stack-protector -mno-stack-arg-probe \
  -nostdlib -nostartfiles -shared -Wl,--entry=DllMain -Wl,--kill-at -s \
  -o soilfem_mat_x64.dll soilfem_mat.c -lkernel32 -lgcc
x86_64-w64-mingw32-objdump -p soilfem_mat_x64.dll | grep -E "DLL Name|sfm_" || true
