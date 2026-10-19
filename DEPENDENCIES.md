# Dependencies

All packages are pinned in [requirements.txt](requirements.txt). Supported Python versions: 3.10, 3.11, 3.12.

## Direct dependencies

The code imports these packages.

| Package | Version | Used in | Use |
| ------- | ------- | ------- | --- |
| `pandas` | 2.2.3 | both scripts | Reads the Excel sheets and groups rows (`read_excel`, `groupby`). |
| `openpyxl` | 3.1.5 | indirect, through `pandas` | Engine that `pandas.read_excel` uses for `.xlsx` files. The code does not import it. |
| `ifctester` | 0.8.1 | both scripts | Builds the IDS objects (`ids.Ids`, `ids.Specification`, facets, `ids.Restriction`) and writes the XML. |
| `deepdiff` | 8.6.2 | `custom_functions.py` | Compares specification, general and applicability data (`DeepDiff`). |
| `deepmerge` | 2.0 | `custom_functions.py` | Merges requirement and general data dicts (`Merger`). |

The code also uses `xml.etree.ElementTree`, `argparse`, `itertools`, `copy`, `os` and `datetime` from the standard library.

## Transitive dependencies

These packages are pinned only to make the install reproducible. The code does not use them directly.

| Package | Version | Pulled in by |
| ------- | ------- | ------------ |
| `ifcopenshell` | 0.8.1.post1 | `ifctester` |
| `xmlschema` | 3.4.3 | `ifctester` |
| `elementpath` | 4.8.0 | `xmlschema` |
| `numpy` | 2.2.3 | `pandas`, `ifcopenshell` |
| `shapely` | 2.0.7 | `ifcopenshell` |
| `lark` | 1.2.2 | `ifcopenshell` |
| `isodate` | 0.7.2 | `ifcopenshell` |
| `python-dateutil` | 2.9.0.post0 | `pandas` |
| `six` | 1.17.0 | `python-dateutil` |
| `pytz` | 2025.1 | `pandas` |
| `tzdata` | 2025.1 | `pandas` |
| `et_xmlfile` | 2.0.0 | `openpyxl` |
| `orderly-set` | 5.5.0 | `deepdiff` |
| `typing_extensions` | 4.12.2 | several packages |

## Rules for a change

- Before you add a package, make sure that it has recent commits and more than one maintainer.
- Update `requirements.txt` and this file in the same commit.
- `ifctester` and `ifcopenshell` must have matching versions. Upgrade them together.
