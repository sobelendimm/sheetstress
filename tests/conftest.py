import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "examples"))
import make_examples  # noqa: E402


@pytest.fixture(scope="session")
def ex(tmp_path_factory):
    return make_examples.build(tmp_path_factory.mktemp("ex"))
