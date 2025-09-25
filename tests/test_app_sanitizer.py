import math
import numpy as np
from app import _sanitize_for_json


def test_sanitize_converts_nan_and_numpy_types():
    payload = {
        "float_nan": float("nan"),
        "np_nan": np.float64("nan"),
        "np_int": np.int64(5),
        "array": np.array([1, np.nan, np.int64(2)]),
        "nested": [1, np.float64(3.5), float("nan")],
    }

    sanitized = _sanitize_for_json(payload)

    assert sanitized["float_nan"] is None
    assert sanitized["np_nan"] is None
    assert sanitized["np_int"] == 5
    assert sanitized["array"] == [1, None, 2]
    assert sanitized["nested"] == [1, 3.5, None]
