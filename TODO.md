# TODO

Findings from a code review of `main` at `513e521` (2026-10-16). Line numbers refer to that commit.
Remove an item in the commit that fixes it, and add a `Fixed` entry to [CHANGELOG.md](CHANGELOG.md).

## Bugs

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
