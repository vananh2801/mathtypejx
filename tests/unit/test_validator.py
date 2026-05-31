"""Unit tests for conversion quality validation."""

from mathtypejx.validator import validate_conversion_quality

MATHML_NS = "http://www.w3.org/1998/Math/MathML"
OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _mathml(body: str) -> str:
    return f'<math xmlns="{MATHML_NS}">{body}</math>'


def _omml(body: str) -> str:
    return f'<m:oMath xmlns:m="{OMML_NS}">{body}</m:oMath>'


def test_rejects_empty_subscript_slot():
    mathml = _mathml("<msub><mi>A</mi><mn>6</mn></msub>")
    omml = _omml(
        "<m:sSub>"
        "<m:e><m:r><m:t>A</m:t></m:r></m:e>"
        "<m:sub/>"
        "</m:sSub>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("m:sub" in error for error in result["errors"])


def test_rejects_empty_denominator_slot():
    mathml = _mathml("<mfrac><mn>1</mn><mn>6</mn></mfrac>")
    omml = _omml(
        "<m:f>"
        "<m:num><m:r><m:t>1</m:t></m:r></m:num>"
        "<m:den/>"
        "</m:f>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("m:den" in error for error in result["errors"])


def test_rejects_missing_token():
    mathml = _mathml("<msub><mi>A</mi><mn>6</mn></msub>")
    omml = _omml(
        "<m:sSub>"
        "<m:e><m:r><m:t>A</m:t></m:r></m:e>"
        "<m:sub><m:r><m:t>7</m:t></m:r></m:sub>"
        "</m:sSub>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("'6'" in error for error in result["errors"])


def test_rejects_fraction_structure_loss():
    mathml = _mathml("<mfrac><mi>a</mi><mi>b</mi></mfrac>")
    omml = _omml("<m:r><m:t>a</m:t></m:r><m:r><m:t>b</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("mfrac" in error for error in result["errors"])


def test_accepts_valid_subscript():
    mathml = _mathml("<msub><mi>A</mi><mn>6</mn></msub>")
    omml = _omml(
        "<m:sSub>"
        "<m:e><m:r><m:t>A</m:t></m:r></m:e>"
        "<m:sub><m:r><m:t>6</m:t></m:r></m:sub>"
        "</m:sSub>"
    )

    result = validate_conversion_quality(mathml, omml)


def test_accepts_valid_mroot_degree():
    mathml = _mathml("<mroot><mi>x</mi><mn>3</mn></mroot>")
    omml = _omml(
        "<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg>"
        "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]


def test_rejects_missing_mroot_degree():
    mathml = _mathml("<mroot><mi>x</mi><mn>3</mn></mroot>")
    omml = _omml("<m:rad><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("mroot degree" in error for error in result["errors"])


def test_accepts_valid_matrix_shape():
    mathml = _mathml(
        "<mtable>"
        "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
        "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr>"
        "</mtable>"
    )
    omml = _omml(
        "<m:m>"
        "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
        "<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr>"
        "</m:m>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]


def test_rejects_flattened_matrix():
    mathml = _mathml("<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr></mtable>")
    omml = _omml("<m:r><m:t>ab</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("matrix" in error for error in result["errors"])


def test_rejects_matrix_cell_loss():
    mathml = _mathml(
        "<mtable>"
        "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
        "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr>"
        "</mtable>"
    )
    omml = _omml(
        "<m:m>"
        "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
        "<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e></m:mr>"
        "</m:m>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("matrix cell" in error for error in result["errors"])


def test_accepts_mfenced_delimiter():
    mathml = _mathml("<mfenced><mi>x</mi></mfenced>")
    omml = _omml("<m:d><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>")

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]


def test_rejects_missing_mfenced_delimiter():
    mathml = _mathml("<mfenced><mi>x</mi></mfenced>")
    omml = _omml("<m:r><m:t>x</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("delimiter" in error or "mfenced" in error for error in result["errors"])


def test_accepts_accent_structure():
    mathml = _mathml("<mover><mi>x</mi><mo>¯</mo></mover>")
    omml = _omml("<m:bar><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar><m:r><m:t>¯</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]


def test_rejects_missing_accent_structure():
    mathml = _mathml("<mover><mi>x</mi><mo>¯</mo></mover>")
    omml = _omml("<m:r><m:t>x¯</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("accent" in error for error in result["errors"])


def test_accepts_nary_limits():
    mathml = _mathml("<munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover><mi>x</mi>")
    omml = _omml(
        "<m:nary><m:sub><m:r><m:t>i</m:t></m:r></m:sub>"
        "<m:sup><m:r><m:t>n</m:t></m:r></m:sup>"
        "<m:e><m:r><m:t>∑x</m:t></m:r></m:e></m:nary>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]


def test_rejects_flattened_nary_limits():
    mathml = _mathml("<munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover><mi>x</mi>")
    omml = _omml("<m:r><m:t>∑inx</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("n-ary" in error for error in result["errors"])


def test_warns_for_function_without_func_structure():
    mathml = _mathml("<mrow><mi>sin</mi><mfenced><mi>x</mi></mfenced></mrow>")
    omml = _omml("<m:d><m:e><m:r><m:t>sinx</m:t></m:r></m:e></m:d>")

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]
    assert any("function style" in warning for warning in result["warnings"])


def test_accepts_function_structure_without_warning():
    mathml = _mathml("<mrow><mi>sin</mi><mi>x</mi></mrow>")
    omml = _omml(
        "<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName>"
        "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]
    assert not any("function style" in warning for warning in result["warnings"])


def test_warns_for_phantom_spacing_and_alignment():
    mathml = _mathml("<mphantom><mi>x</mi></mphantom><mspace/><mpadded><mi>y</mi></mpadded><maligngroup/><malignmark/>")
    omml = _omml("<m:r><m:t>xy</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]
    assert any("phantom" in warning for warning in result["warnings"])
    assert any("spacing" in warning for warning in result["warnings"])
    assert any("alignment" in warning for warning in result["warnings"])


def test_rejects_prescript_loss():
    mathml = _mathml(
        "<mmultiscripts><mi>X</mi><mi>i</mi><none/><mprescripts/><mi>j</mi><none/></mmultiscripts>"
    )
    omml = _omml("<m:sSub><m:e><m:r><m:t>X</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub><m:r><m:t>j</m:t></m:r>")

    result = validate_conversion_quality(mathml, omml)

    assert not result["valid"]
    assert any("prescript" in error for error in result["errors"])


def test_accepts_prescript_structure():
    mathml = _mathml(
        "<mmultiscripts><mi>X</mi><mi>i</mi><none/><mprescripts/><mi>j</mi><none/></mmultiscripts>"
    )
    omml = _omml(
        "<m:sPre><m:sub><m:r><m:t>j</m:t></m:r></m:sub>"
        "<m:sup><m:r><m:t>none</m:t></m:r></m:sup>"
        "<m:e><m:r><m:t>Xi</m:t></m:r></m:e></m:sPre>"
    )

    result = validate_conversion_quality(mathml, omml)

    assert result["valid"], result["errors"]
