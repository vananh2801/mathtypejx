# mathtypejx

`mathtypejx` is a Python package for converting MathType and legacy Equation Editor OLE formulas in Word `.docx` files into native OMML equations.

The core MathType Equation Format (MTEF) parser is implemented in Python. The final MathML-to-OMML step uses Microsoft Office's `MML2OMML.XSL`, which is normally installed with Office.

## Features

- Scan Word `.docx` files for embedded MathType OLE formula objects.
- Extract MTEF bytes from OLE compound files.
- Parse MTEF v5 and legacy v3 equation streams in Python.
- Convert parsed equations to MathML, then to OMML.
- Replace OLE objects in the `.docx` XML with native Word math.
- Provide per-formula validation and conversion reports.

## Development and Validation Background

`mathtypejx` was developed primarily against Chinese Gaokao physics exam documents, where old Word files often contain MathType OLE formulas rather than native OMML. The private validation corpus is not redistributed in this repository, but the development run that drove the parser covered a 344-document Chinese physics exam corpus with 7,888 MathType OLE formulas. A normalized comparison against the Ruby-based pipeline recorded 7,886/7,888 MathML and OMML matches; the two differences were overbar formulas where the Python path preserved the base character that the Ruby path dropped.

The formulas exercised in that corpus include common high-school physics notation such as fractions, roots, superscripts/subscripts, isotope-style prescripts, vectors, overbars/underbars, hats/tildes/arcs, large operators, limits, bracketed expressions, matrices, stacked equations, boxed or crossed-out terms, long division, color/font records, and text-mode physical units such as `kg*m/s`.

The code also includes defensive handling for failures that appeared while processing those documents:

- Missing or unreadable OLE binaries are marked failed and left in place.
- Multiple MathType stream names are tried: `Equation Native`, `EquationNative`, and `Equation`.
- MTEF v3 and v5 records are parsed separately, with future/comment records skipped safely when possible.
- `EQN_PREFS` nibble-packed data is bounded and includes recovery logic for over-consumption that can otherwise hide the equation body.
- Top-level `PILE` and `MATRIX` records are accepted, and their internal `LINE` records are converted into the slot shape expected by the bundled XSLT.
- Subscript/superscript movement handles parenthesized bases, isotope-style preceding scripts, invisible MathType spacing, and per-formula mover state isolation.
- Text-mode characters are wrapped as MathML tokens so overbars and other embellishments keep their base character.
- MathML normalization fixes missing namespaces, bare text inside containers, empty root-degree cases, and `mtext` baseline issues before OMML conversion.
- OMML quality checks block replacements with empty critical slots, token loss, structure loss, matrix row/cell loss, delimiter loss, accent/bar loss, n-ary limit loss, or malformed XML.
- Unsupported-character noise from `MML2OMML.XSL` is stripped from generated OMML where possible.
- Failed conversions keep the original OLE object in the output document, and the report preserves per-formula status, risk level, warnings, and errors.

## Install

```powershell
python -m pip install .
```

For development:

```powershell
python -m pip install -e ".[dev]"
python -m pytest -q
```

## Requirements

- Python 3.10 or newer
- `lxml`
- `olefile`
- `python-docx`
- Microsoft Office `MML2OMML.XSL` for OMML output

Typical Windows Office paths:

```text
C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL
```

You can pass a custom XSL path with `--xsl`.

## CLI

```powershell
mathtypejx health
mathtypejx convert input.docx -o output.docx
mathtypejx convert input.docx -o output.docx --xsl "C:\path\to\MML2OMML.XSL"
```

## Python API

```python
from mathtypejx import convert_mathtype_to_omml

report = convert_mathtype_to_omml(
    "input.docx",
    "output.docx",
    remove_edit_info=True,
    parallel=True,
    max_workers=8,
)

print(report.succeeded, report.failed)
```

## Scope and Limitations

- Supports `.docx` files, not legacy binary `.doc` files.
- Supports MathType OLE (`Equation.DSMT4`) and older Equation Editor OLE streams.
- Does not convert formulas embedded only as WMF images.
- OMML output requires an available `MML2OMML.XSL`.
- Public tests use tiny binary OLE fixtures and synthetic documents; large private exam corpora are intentionally not included.

## Acknowledgements

This project builds on the public MathType-to-MathML work that came before it:

- [`jure/mathtype`](https://github.com/jure/mathtype) and [`sbulka/mathtype`](https://github.com/sbulka/mathtype), Ruby implementations for reading MathType binaries and representing MTEF as XML.
- [`jure/mathtype_to_mathml`](https://github.com/jure/mathtype_to_mathml), which provides the original XSLT-based MTEF XML to MathML conversion approach.
- [`transpect/mathtype-extension`](https://github.com/transpect/mathtype-extension), whose public documentation and bundled fontmaps/XSLT ecosystem helped clarify the conversion pipeline.
- The `mathtype_to_mathml_plus` Ruby gem, which combines the `mathtype` gem and XSLTs into a MathType binary to MathML conversion flow.

`mathtypejx` is a Python implementation and packaging of this conversion path for `.docx` to OMML workflows, not an original discovery of the MTEF conversion model.

## License

Project code is released under the MIT License. Bundled font map assets under `src/mathtypejx/mtef/xslt/xsl/fontmaps` retain their upstream BSD-style notice; see `NOTICE` and the upstream `LICENSE` file in that directory.
