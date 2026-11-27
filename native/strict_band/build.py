"""Build the optional strict band kernel with Zig 0.14.1, without a CRT DLL."""
import argparse, hashlib, json, subprocess
from pathlib import Path

root = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument('--zig', default='zig')
parser.add_argument('--output', type=Path, default=root/'soilfem_band_fast_x64.dll')
args = parser.parse_args()
command = [args.zig, 'cc', '-target', 'x86_64-windows-gnu', '-O3',
           '-fno-fast-math', '-ffp-contract=off', '-fno-builtin', '-ffreestanding',
           '-march=x86_64', '-msse2', '-mno-avx', '-shared', '-nostdlib',
           '-Wl,--no-entry', str(root/'soilfem_band_fast.c'), '-o', str(args.output)]
subprocess.run(command, check=True)
record = {'command': command, 'compiler': 'Zig 0.14.1',
          'source_sha256': hashlib.sha256((root/'soilfem_band_fast.c').read_bytes()).hexdigest(),
          'dll_sha256': hashlib.sha256(args.output.read_bytes()).hexdigest()}
args.output.with_suffix('.build.json').write_text(json.dumps(record, indent=2), encoding='utf8')
