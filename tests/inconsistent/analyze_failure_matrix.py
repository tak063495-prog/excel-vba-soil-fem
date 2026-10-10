"""Independently inspect a saved constrained CSR failure system using NumPy."""
import argparse
import hashlib
import json
import time
from pathlib import Path

import numpy as np


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--matrix', type=Path, required=True)
    parser.add_argument('--vectors', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    entries = np.genfromtxt(args.matrix, delimiter=',', names=True)
    vectors = np.genfromtxt(args.vectors, delimiter=',', names=True)
    n = len(vectors)
    if not np.array_equal(vectors['dof'], np.arange(n)):
        raise ValueError('Expected consecutive vector DOF indices')
    rows, columns = entries['row'].astype(int), entries['column'].astype(int)
    if (np.any(rows < 0) or np.any(columns < 0) or np.any(rows >= n) or np.any(columns >= n)
            or not np.isfinite(entries['value']).all() or not np.isfinite(vectors['b']).all()):
        raise ValueError('Invalid matrix/vector indices or nonfinite values')
    matrix = np.zeros((n, n))
    np.add.at(matrix, (rows, columns), entries['value'])
    rhs, original = vectors['b'], vectors['gmres_x']
    free = vectors['constraint'] == 0
    rhs_norm = float(np.linalg.norm(rhs))
    if rhs_norm == 0:
        raise ValueError('Expected a nonzero diagnostic right-hand side')
    record = dict(ndof=n, nnz=len(entries), free_dof=int(free.sum()),
                  matrix_sha256=hashlib.sha256(args.matrix.read_bytes()).hexdigest(),
                  vectors_sha256=hashlib.sha256(args.vectors.read_bytes()).hexdigest(),
                  rhs_norm=rhs_norm,
                  gmres_true_residual=float(np.linalg.norm(matrix @ original - rhs) / rhs_norm),
                  symmetry_relative=float(np.linalg.norm(matrix - matrix.T) / np.linalg.norm(matrix)),
                  ilu_min_diagonal=float(np.min(np.abs(vectors['ilu_diagonal']))),
                  trial_u_max=float(np.max(np.abs(vectors['trial_u']))),
                  committed_u_max=float(np.max(np.abs(vectors['committed_u']))))
    started = time.monotonic()
    try:
        solved = np.linalg.solve(matrix, rhs)
        record.update(dense_true_residual=float(np.linalg.norm(matrix @ solved - rhs) / rhs_norm),
                      dense_max_correction=float(np.max(np.abs(solved))),
                      dense_seconds=time.monotonic() - started)
    except np.linalg.LinAlgError as error:
        record['dense_error'] = str(error)
    free_matrix = matrix[np.ix_(free, free)]
    eigenvalues = np.linalg.eigvalsh((free_matrix + free_matrix.T) / 2)
    minimum = float(np.min(np.abs(eigenvalues)))
    record.update(min_eigenvalue=float(eigenvalues[0]), max_eigenvalue=float(eigenvalues[-1]),
                  negative_eigenvalues=int((eigenvalues < 0).sum()),
                  small_eigenvalues=int((np.abs(eigenvalues) < abs(eigenvalues[-1]) * 1e-12).sum()),
                  condition_estimate=float(np.max(np.abs(eigenvalues)) / minimum) if minimum else None,
                  eigenanalysis_basis='symmetric part of the free-DOF matrix; check symmetry_relative first')
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    print(f"Diagnostic system: {n} DOF, original true residual {record['gmres_true_residual']:.6g}, "
          f"dense true residual {record.get('dense_true_residual')}, condition {record['condition_estimate']}")


if __name__ == '__main__':
    main()
