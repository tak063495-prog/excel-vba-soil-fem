# Singular operators in rejected associated look-ahead trials

Associated potential recovery may visit intermediate states with a larger
force residual before accepting a combined correction. The stiffness of one
of those uncommitted states can become singular. That failed trial operator
must not invalidate the accepted state from which the recovery started.

A `GLOBAL_SINGULAR` result from an inner look-ahead linear solve now rejects
that recovery trial. The code restores the starting displacement, reevaluates
its material state and internal force, rebuilds its tangent, and restores its
previous SRM failure note. The caller can then try its other correction paths.
No material state or load increment is committed by this recovery procedure.

Input, runtime, capacity and user-cancellation errors still propagate. Failure
to restore the starting operator also remains fatal. This does not classify
an arbitrary singular global system as a physical SRM limit or change the
material force, strength parameters or convergence tolerances.

`tests/check_associated_linear_reject.py` injects failure at a real uncommitted
FE look-ahead point. Its old-code mode reproduces a fatal status and a stale
internal force. The corrected AUTO/OFF paths restore force and tangent before
accepting a genuinely smaller residual through the ordinary fallback. The
test also checks fatal error categories and failed base restoration.
