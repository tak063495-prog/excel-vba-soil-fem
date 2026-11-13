# Associated-flow residual recovery

`ASSOCIATED_ENERGY_RECOVERY=1` (default) enables a bounded potential-guided
look-ahead when ordinary residual line search fails. It applies only to actual
associated continuum materials (`phi=psi`), single-substep return mapping,
satisfied prescribed displacements, and uncoupled static displacement analysis.
JOINT, fixed-stress, mixed and consolidation operators retain their existing
recovery paths. The selected flow-policy label alone does not establish eligibility.

For this restricted associated problem, the negative free residual is a convex
incremental-potential gradient. A positive residual-direction product at both
endpoints bounds the directional derivative below zero along that segment.
The helper may therefore cross a ridge in the Euclidean force-residual norm
without treating that intermediate state as converged. It checks eligibility
again at each intermediate state and requires finite values and bounded work.

At most six look-ahead steps and sixteen halved trial lengths per step are
allowed. Inner Newton solves consume the normal global iteration budget.
Only a combined correction that decreases the **original actual free residual**
is accepted. The existing equilibrium, displacement and material convergence
criteria remain in effect. This helper never commits a material state or a load
increment. Rejection restores the displacement, material state and current
operator; fatal/capacity errors and user cancellation retain their normal meaning.

Set the option to 0 to retain the previous recovery selection. It does not
change cohesion, friction, dilation, the force law, or convergence tolerances.

`tests/check_associated_watchdog.py` exercises a real two-element elastic FE
problem with a deliberately chosen descent direction whose first trial increases
the force-residual norm. It checks eventual residual reduction, unchanged
committed displacement/increment count, and exact rollback of a negative-work
trial. Separate full-model testing is required to establish completion of a
particular gravity/load/SRM analysis.

Constitutive background: [Sysala and Cermak, implicit Mohr-Coulomb return mapping](https://arxiv.org/abs/1508.07435).
