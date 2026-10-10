"""Regression for independent gravity-force residual normalization.

Run with --fixtures fixtures.json --result <completed native result.json>.
The input files are read-only; synthetic constraint/stress changes stay in memory.
"""
import argparse
import copy
import json

from audit_material_force import audit


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--fixtures', required=True)
    parser.add_argument('--result', required=True)
    args = parser.parse_args()
    with open(args.result, encoding='utf-8-sig') as stream:
        result = json.load(stream)
    with open(args.fixtures, encoding='utf-8-sig') as stream:
        fixtures = json.load(stream)['fixtures']
    fixture = copy.deepcopy(next(f for f in fixtures if f['name'] == result['fixture']))
    for node in fixture['nodes']:
        node[3:5] = [1, 1]
    fixture['nodes'][0][4] = 0
    for offset in range(0, len(result['stresses_flat']), 5):
        result['stresses_flat'][offset+2:offset+5] = [0, 0, 0]
    columns = result['gauss_state_columns']
    for offset in range(0, len(result['gauss_state_flat']), len(columns)):
        result['gauss_state_flat'][offset+columns.index('sigma_z')] = 0
    value = audit(fixture, result)
    assert abs(value['material_only_relative_residual'] - 1) < 1e-12
    assert value['all_dof_normalized_material_residual'] < 1
    assert value['material_force_balance_pass'] is False
    fixture['nodes'][0][4] = 1
    try:
        audit(fixture, result)
    except ValueError as error:
        assert 'No finite nonzero gravity load on free DOFs' in str(error)
    else:
        raise AssertionError('Zero free load was accepted')
    print('PASS: support loads cannot dilute the free-DOF residual; zero free load is refused.')


if __name__ == '__main__':
    main()
