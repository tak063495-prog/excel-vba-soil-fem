"""Independently integrate material stress forces for gravity-only SRM fixtures.

No workbook writes. The residual excludes every artificial hourglass force.
Only complete accepted INCONSISTENT states are judged against full loading.
"""
import argparse
import json
import math
from pathlib import Path

import numpy as np
from analyze_practical_results import displacement_geometry


def audit(fixture, result):
    nodes = np.asarray(fixture['nodes'], dtype=float)
    coordinates = nodes[:, 1:3]
    constrained = nodes[:, 3:5] != 0
    stress_rows = np.asarray(result['stresses_flat'], dtype=float).reshape(-1, 5)
    if len(stress_rows) != 4*len(fixture['elements']) or not np.isfinite(stress_rows).all():
        raise ValueError('Invalid stress coverage: ' + result['case'])
    stress = {tuple(map(int, row[:2])): row[2:] for row in stress_rows}
    expected_points = {(int(e[0]), g) for e in fixture['elements'] for g in range(1, 5)}
    if len(stress) != len(stress_rows) or set(stress) != expected_points:
        raise ValueError('Duplicate or invalid stress identifiers: ' + result['case'])
    body = np.zeros((len(nodes), 2))
    internal = np.zeros_like(body)
    metrics = dict(zip(result['metric_names'], result['metrics']))
    fs = metrics['fss']
    if not fs > 0:
        raise ValueError('Accepted state has no valid strength factor')
    yield_margins = []
    full_yield_margins = []
    columns = result.get('gauss_state_columns', [])
    gauss_rows = np.asarray(result.get('gauss_state_flat', []), dtype=float)
    if not columns or 'sigma_z' not in columns or gauss_rows.size % len(columns):
        raise ValueError('Missing complete Gauss sigma_z records: ' + result['case'])
    gauss_rows = gauss_rows.reshape(-1, len(columns))
    if len(gauss_rows) != 4*len(fixture['elements']) or not np.isfinite(gauss_rows).all():
        raise ValueError('Invalid Gauss-state coverage: ' + result['case'])
    sigma_z = {(int(row[columns.index('element')]), int(row[columns.index('gauss')])):
               row[columns.index('sigma_z')] for row in gauss_rows}
    if len(sigma_z) != len(gauss_rows) or set(sigma_z) != expected_points:
        raise ValueError('Duplicate or invalid Gauss-state identifiers: ' + result['case'])
    g = 1 / math.sqrt(3)
    for element in fixture['elements']:
        indices = np.asarray(element[1:9], dtype=int) - 1
        xy = coordinates[indices]
        material = fixture['materials'][element[9]-1]
        for gp, (r, s) in enumerate([(-g, -g), (g, -g), (g, g), (-g, g)], 1):
            shape = np.array([-.25*(1-r)*(1-s)*(1+r+s), -.25*(1+r)*(1-s)*(1-r+s),
                              -.25*(1+r)*(1+s)*(1-r-s), -.25*(1-r)*(1+s)*(1+r-s),
                              .5*(1-r*r)*(1-s), .5*(1+r)*(1-s*s),
                              .5*(1-r*r)*(1+s), .5*(1-r)*(1-s*s)])
            dr = np.array([.25*(1-s)*(2*r+s), .25*(1-s)*(2*r-s),
                           .25*(1+s)*(2*r+s), .25*(1+s)*(2*r-s),
                           -r*(1-s), .5*(1-s*s), -r*(1+s), -.5*(1-s*s)])
            ds = np.array([.25*(1-r)*(r+2*s), .25*(1+r)*(-r+2*s),
                           .25*(1+r)*(r+2*s), .25*(1-r)*(-r+2*s),
                           -.5*(1-r*r), -(1+r)*s, .5*(1-r*r), -(1-r)*s])
            natural = np.vstack([dr, ds])
            jacobian = natural @ xy
            derivative = np.linalg.solve(jacobian, natural)
            measure = np.linalg.det(jacobian) * material[3]
            sx, sy, tau = stress[element[0], gp]
            # Reconstruct all three principal stresses independently. This
            # uses sigma_z, not the return mapper's exported principal values.
            reduced_phi = math.atan(math.tan(math.radians(material[5])) / fs)
            reduced_c = material[6] / fs
            radius = math.hypot(.5*(sx-sy), tau)
            pressure_term = .5*(sx+sy)*math.sin(reduced_phi)
            cohesion_term = reduced_c*math.cos(reduced_phi)
            yield_margins.append((radius-pressure_term-cohesion_term) /
                                 max(1., radius, abs(pressure_term), abs(cohesion_term)))
            principal = [.5*(sx+sy)+radius, .5*(sx+sy)-radius, sigma_z[element[0], gp]]
            largest, smallest = max(principal), min(principal)
            full_radius = .5*(largest-smallest)
            full_pressure = .5*(largest+smallest)*math.sin(reduced_phi)
            full_yield_margins.append((full_radius-full_pressure-cohesion_term) /
                                     max(1., full_radius, abs(full_pressure), abs(cohesion_term)))
            # Workbook B and its stored stresses are compression-positive.
            internal[indices, 0] -= (derivative[0]*sx + derivative[1]*tau) * measure
            internal[indices, 1] -= (derivative[1]*sy + derivative[0]*tau) * measure
            body[indices, 1] -= material[4] * shape * measure
    norm_residual = float(np.linalg.norm((body-internal)[~constrained]))
    norm_reference = float(np.linalg.norm(body[~constrained]))
    all_dof_reference = float(np.linalg.norm(body))
    if not math.isfinite(norm_reference) or norm_reference <= 0:
        raise ValueError('No finite nonzero gravity load on free DOFs: ' + result['case'])
    geometry = displacement_geometry(fixture, result['displacements_flat'])
    geometry.pop('gauss')
    return dict(case=result['case'], build=result['build'],
                source_sha256=result['source_sha256'], status=result['status'],
                fos_pass=metrics['fos_pass'], fos_fail=metrics['fos_fail'],
                native_relative_residual=metrics['relative_residual'],
                material_only_relative_residual=norm_residual/norm_reference,
                material_only_residual_norm=norm_residual,
                reference_force_norm=norm_reference,
                reference_force_basis='gravity-force norm on free DOFs only',
                all_dof_reference_force_norm=all_dof_reference,
                all_dof_normalized_material_residual=norm_residual/all_dof_reference,
                body_force_y=float(body[:, 1].sum()),
                expected_body_force_y=fixture['checks']['body_force_sum_y'],
                material_force_balance_pass=norm_residual/norm_reference <= 1e-5,
                max_normalized_in_plane_mc_yield_margin=float(max(yield_margins)),
                in_plane_mc_necessary_condition_pass=bool(max(yield_margins) <= 1e-6),
                max_normalized_full_mc_yield_margin=float(max(full_yield_margins)),
                full_mc_admissibility_pass=bool(max(full_yield_margins) <= 1e-6),
                full_mc_basis='principal stresses reconstructed from sigma_x, sigma_y, tau_xy and sigma_z; unchanged small-strain geometry',
                gauss_point_count=len(full_yield_margins),
                **geometry)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--fixtures', type=Path, required=True)
    parser.add_argument('--results-root', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--expected-build')
    parser.add_argument('--expected-count', type=int)
    parser.add_argument('--require-balanced', action='store_true')
    args = parser.parse_args()
    fixtures = {f['name']: f for f in json.loads(args.fixtures.read_text(encoding='utf-8-sig'))['fixtures']}
    records = []
    for path in sorted(args.results_root.glob('**/results/*.json')):
        result = json.loads(path.read_text(encoding='utf-8-sig'))
        if 'metrics' not in result or not result['status'].startswith('PASS|'):
            continue
        if result['flow_policy'] != 'INCONSISTENT' or result['fixture'] not in fixtures:
            continue
        if args.expected_build and result['build'] != args.expected_build:
            continue
        fixture = fixtures[result['fixture']]
        if fixture['loads']:
            raise ValueError('This audit requires a gravity-only fixture: '+fixture['name'])
        records.append(audit(fixture, result))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(dict(schema='srm_material_force_audit_v3', cases=records),
                                     ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    for r in records:
        print(r['case'], 'material_residual=',r['material_only_relative_residual'],
              'balance_pass=',r['material_force_balance_pass'],
              'full_yield_margin=',r['max_normalized_full_mc_yield_margin'])
    if args.expected_count is not None and len(records) != args.expected_count:
        raise SystemExit(f'Expected {args.expected_count} accepted states, found {len(records)}')
    if args.require_balanced and (not records or
            not all(r['material_force_balance_pass'] and
                    r['full_mc_admissibility_pass'] for r in records)):
        raise SystemExit('Independent material-force balance audit failed')


if __name__ == '__main__':
    main()
