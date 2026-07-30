import pathlib
import sys

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session")
def root() -> pathlib.Path:
    return ROOT


@pytest.fixture(scope="session")
def sample_dir(root) -> pathlib.Path:
    return root / "sample_data"
