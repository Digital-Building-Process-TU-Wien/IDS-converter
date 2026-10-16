# TODO

Findings from a code review of `main` at `513e521` (2026-10-16). Line numbers refer to that commit.
Remove an item in the commit that fixes it, and add a `Fixed` entry to [CHANGELOG.md](CHANGELOG.md).

## Bugs

### 8. Spec title raises `IndexError` for mixed entity values

- Where: [custom_functions.py:904](custom_functions.py#L904)
- Cause: The title loop reads `PredefinedType[j]` for each entity `j`. `separate_dict_value` adds a
  predefined type only for the entities that have one.
- Effect: `SpecificationCardinality` = `required` (the OR values are not split) and
  `A.Entity` = `IFCWALL|IFCWALL.STANDARD` without a `SpecificationName` give
  `Entity=['IFCWALL','IFCWALL']`, `PredefinedType=['STANDARD']` and an `IndexError` at `j=1`.
- Fix: Check the list length, or reject the mixed value in `separate_dict_value`.

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
