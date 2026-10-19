#ifndef SFN_ASSEMBLY_H
#define SFN_ASSEMBLY_H
#include <stdint.h>

/* Build the CSR scatter list used by the VBA assembly path.
 * P: [0] element base [1] element stride in bytes [2] nElem [3] kmat offset
 * [4] Long element positions(nElem,16,16), VBA column-major [5] CSR nnz
 * [6] Long start(nElem+1) [7] output capacity, or -1 for count-only
 * [8..11] Long row/col/position/valueIndex(capacity) [12] Long count output.
 * Element kmat is Double(16,16), VBA column-major; iterate element/row/col.
 * Count-only leaves arrays 8..11 untouched (may be NULL).
 * Caller owns valid, disjoint spans. No allocation or floating arithmetic.
 * Returns 0 on success and -1 with all outputs untouched on rejection. */
__declspec(dllexport) int32_t __stdcall sfn_csr_scatter_list(const int64_t *P);

#endif
