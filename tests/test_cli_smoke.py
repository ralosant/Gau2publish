
import sys
import subprocess

def test_cli_help():
    proc = subprocess.run([sys.executable, "-m", "gau2publish", "--help"], capture_output=True, text=True)
    assert proc.returncode == 0
    out = (proc.stdout + proc.stderr).lower()
    assert "gau2publish" in out
