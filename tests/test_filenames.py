from utils import safe_filename

def test_safe_filename():
    assert safe_filename("a/b\\c.pdf") == "a_b_c.pdf"
    assert safe_filename("") == "attachment"
