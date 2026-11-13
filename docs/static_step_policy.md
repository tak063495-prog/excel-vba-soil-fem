# Static load-step options

`STATIC_KEEP_ACCEPTED_STEP=1` retains the width of a converged normal-Newton
increment when it needed more than five iterations. The default, 0, retains
the previous iteration-count adaptation (0.7 for six to eight iterations,
0.5 above eight). Fast convergence can still grow the step. Elastic-correction
iteration scheduling retains its existing rules.

This option changes only the proposal after acceptance. The equilibrium,
displacement and material checks, failed-trial rollback, retry cutback,
minimum/maximum step bounds, remaining-load cap and cancellation are unchanged.
A slowly converged accepted step therefore does not by itself trigger another
shrinkage. Failed attempts can still reduce the step.

`STATIC_INITIAL_STEPS=0` retains the previous initial widths: 0.1 for gravity,
0.05 for applied loading. An integer from 1 to 100000 overrides the initial
width with its reciprocal for each static stage, subject to the existing
maximum of 0.2 and adaptive bounds. Invalid numeric counts raise an input
error; they are not rounded. This is an initial width, not a promise of a
fixed number of increments. SRM uses the same static load solver.

Both settings appear on the analysis settings page. Use separate cold runs
with smaller increments to assess load-path sensitivity for a particular
elastoplastic model. Passing a scheduling unit test does not establish full
gravity/load/SRM completion or stability of the user's model.
