import sys
import pytest

if __name__ == "__main__":
    sys.path.insert(0, 'src')
    sys.exit(pytest.main(["tests", "-v"]))
