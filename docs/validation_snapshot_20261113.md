# Validation snapshot, 2026-11-13

The user requested stopping the long-running case analyses and preserving
analysis-ready workbooks. The calculations were stopped explicitly; their
interruption is not a nonconvergence bound or a completed SRM search.

The standard large mixed-element case completed self weight, applied loading,
and the complete load history at reduction factors 1.5 and 2.0. The 2.5 trial
was unfinished when stopped. A local T6 mesh comparison completed its initial
self-weight and applied-load history; its SRM search was also unfinished.
These observations do not establish a critical safety factor.

The mechanical solver used for those runs is the `0fc6f01` implementation.
The later `2f02dd7` change affects Q8/T6 size diagnostics only. It preserves
material forces, geometry validity, shape targets and raw Jacobian reporting.
The real Excel equal-area size checks cover nine threshold/scale/translation
conditions in both old and corrected modes. Existing shape, invalid-geometry,
diagonal-selection and edge-flip controls pass in both modes.

The mechanical validation also includes 24 issue fixtures with checkpoint
resume, long-history prediction controls, predicted-state pivot retry,
runtime/input/capacity/cancellation propagation, and the settings migration
guard. The large case's accepted initial checkpoint was independently checked
for displacement/stress consistency, reactions and moment balance.

Saved rerun cases have cleared result tables, an empty restart selection,
current canonical VBA and the paired native libraries. A fixed-Fs verification
case must be read as a prescribed-factor check, not a critical-factor search.
Experimental solver corrections and observation/interruption hooks are not
part of the saved production workbooks or committed VBA.

User attachments and their geometry are retained locally, not included in
this repository. PR #68 contains the code changes; full large-case SRM
verification remains incomplete.
