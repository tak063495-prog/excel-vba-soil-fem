# Optional V1 history after long converged increments

`ACCEL_PREDICT_LONG_HISTORY=1` lets the existing V1 displacement predictor use
a normally accepted static increment even when it required more than five
iterations. The default is 0, which retains the existing five-iteration limit.
V1 itself must also be enabled. The new long-history path excludes retried/
cutback, zero-step, baseline-retry, JOINT, mixed and coupled/transient histories.
The existing short-history policy is unchanged. Stage reset,
step-ratio limits and predictor beta retain their existing behavior.

The history supplies an initial displacement guess. The original material
update, force residual, prescribed-displacement check and convergence
tolerances determine acceptance; this does not skip a load increment or
commit an unchecked prediction.

A 3303/3304 singular-pivot result reached through an uncommitted predicted
increment now uses the existing same-target baseline retry before reducing
the load step. The accepted material/displacement history is restored, V1 is
disabled for that retry, and the previous SRM failure note is retained. An
ordinary baseline singular system and input/runtime/capacity/cancellation
errors keep their existing behavior.

`tests/check_long_predictor.py` uses deliberately damped, real FE corrections
to produce accepted increments longer than five iterations. It compares
AUTO/OFF solves, complete replay fields, eligible-history rejection controls,
same-target retry after a predicted-state pivot, and fatal error propagation.
Its old-code mode reproduces an unnecessary cutback and stale failure note.
