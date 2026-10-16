# TODO

Findings from a code review of `main` at `513e521` (2026-10-16). Line numbers refer to that commit.
Remove an item in the commit that fixes it, and add a `Fixed` entry to [CHANGELOG.md](CHANGELOG.md).

## Bugs

### 11. A sheet without value columns crashes the converter

- Where: [custom_functions.py:150](custom_functions.py#L150)
- Cause: A sheet with neither `R.PropertyValue` nor `R.AttributeValue` gives `removed_items = []`.
  `groupby(...).agg({})` then raises `ValueError: No objects to concatenate`.
- Confirmed: A full run on `8c59647` with the columns `A.Entity`, `R.PropertySet` and `R.Property`
  stops with this error. The same workbook with `R.PropertyValue` converts.
- Fix: Skip step 1 when `removed_items` is empty. Use `df_filled` as `df_step1`.

### 12. A line break in a cell text gives two line breaks on Windows

- Where: `ids.Ids.to_xml` of `ifctester`, then [custom_functions.py:1092](custom_functions.py#L1092)
- Cause: On Windows, `to_xml` writes the cell text `a\r\nb` as `a\r\r\nb`. `ET.parse` in
  `add_comment_to_xml` reads this as two line breaks and writes both to the final file.
- Fix: Replace `\r\n` and `\r` in all text values with `\n` before the facets are created.
  Check the result on Windows and on Linux.

### 13. The order of purpose, milestone and IFC versions changes between runs

- Where: [custom_functions.py:1086](custom_functions.py#L1086), [custom_functions.py:752](custom_functions.py#L752)
- Cause: `Phase`, `Role` and `Usecase` are stored in a `set`. `list(set(...))` also builds the
  IFC version list. Python randomizes the order of strings per process.
- Fix: Keep the order of the first occurrence. Use a list with a membership check or `dict.fromkeys`.

## Cleanup

### 9. The workbook is read three times

- Where: [IDS4ALL-main.py:50](IDS4ALL-main.py#L50), [custom_functions.py:58](custom_functions.py#L58), [custom_functions.py:116](custom_functions.py#L116)
- Cause: Separate `pd.read_excel` calls for the metadata, the headers and the data.
- Fix: Open the workbook once with `pd.ExcelFile`. Read the input sheet once with `dtype=str` and
  get the columns from `df.columns`.
  Also remove the unused function `load_columns` ([custom_functions.py:384](custom_functions.py#L384)).

### 10. Dead checks in the separation code

- Where: [custom_functions.py:314](custom_functions.py#L314), [custom_functions.py:341](custom_functions.py#L341), [custom_functions.py:1062](custom_functions.py#L1062)
- Cause: No code creates a `MissingKey` key. `new_spec` is never `None` at line 341. General data
  values are lists, so `!= ''` is always true.
- Fix: Remove the checks. Merge the two branches in `separate_specs_by_generaldata` with `setdefault`.
