import sys
from pathlib import Path


def pytest_sessionstart(session):
    """
    Ensure the repository root is importable regardless of how pytest is invoked.
    """
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
