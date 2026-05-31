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

## License

Project code is released under the MIT License. Bundled font map assets under `src/mathtypejx/mtef/xslt/xsl/fontmaps` retain their upstream BSD-style notice; see `NOTICE` and the upstream `LICENSE` file in that directory.
