"""Regression tests for MTEF v5 record parsing."""

from lxml import etree

from mathtypejx.mtef import mtef_to_mathml
from mathtypejx.mtef.chars import replace as replace_chars
from mathtypejx.mtef.records5 import _parse_color_def
from mathtypejx.mtef.stream import ByteStream


def test_color_def_with_name_consumes_null_terminated_name():
    stream = ByteStream(
        bytes(
            [
                0x04,  # mtefCOLOR_NAME, RGB model
                0xE8,
                0x03,  # r = 1000
                0x00,
                0x00,  # g = 0
                0x00,
                0x00,  # b = 0
            ]
        )
        + b"Red\x00"
        + bytes([0x0F])
    )

    color_def = _parse_color_def(stream)

    assert color_def["color_name"] == "Red"
    assert stream.peek_uint8() == 0x0F


def test_textmode_char_without_specific_template_becomes_mtext():
    root = etree.fromstring(
        b"""
        <root>
          <char>
            <mt_code_value>0x002D</mt_code_value>
            <variation>textmode</variation>
          </char>
        </root>
        """
    )

    replace_chars(root)

    assert root.find("mtext").text == "-"
    assert root.find("mo") is None


def test_greek_char_is_left_for_xslt_font_mapping():
    root = etree.fromstring(
        b"""
        <root>
          <char>
            <mt_code_value>0x03C0</mt_code_value>
            <variation>textmode</variation>
          </char>
        </root>
        """
    )

    replace_chars(root)

    assert root.find("char") is not None
    assert root.find("mi") is None


def test_textmode_nbsp_is_preserved_as_text():
    root = etree.fromstring(
        b"""
        <root>
          <char>
            <mt_code_value>0x00A0</mt_code_value>
            <variation>textmode</variation>
          </char>
        </root>
        """
    )

    replace_chars(root)

    assert root.find("mtext").text == "\u00a0"


def test_mtef_to_mathml_smoke(mock_ole_binary):
    mathml = mtef_to_mathml(mock_ole_binary)

    assert mathml is not None
    root = etree.fromstring(mathml.encode("utf-8"))
    assert etree.QName(root).localname == "math"
