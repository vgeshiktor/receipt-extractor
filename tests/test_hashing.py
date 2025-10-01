from utils import file_sha256

def test_file_sha256_stable():
    h1 = file_sha256(b"abc")
    h2 = file_sha256(b"abc")
    assert h1 == h2
    assert len(h1) == 64
