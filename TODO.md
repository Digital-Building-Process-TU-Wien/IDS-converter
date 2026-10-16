# TODO

Findings from a code review of `main` at `513e521` (2026-10-16). Line numbers refer to that commit.
Remove an item in the commit that fixes it, and add a `Fixed` entry to [CHANGELOG.md](CHANGELOG.md).

## Bugs

### 5. `NameError` when the sheet has no known columns

- Where: [custom_functions.py:201](custom_functions.py#L201), [custom_functions.py:356](custom_functions.py#L356)
- Cause: `df_final` and `spec_ifc_version_col_used` are only set inside `if relevant_columns:`.
- Effect: A wrong sheet name, a typo in the headers, or a wrong `Skipped rows` value gives
  `NameError: name 'df_final' is not defined`.
- Fix: Raise a clear error when `relevant_columns` is empty.

### 6. Property descriptions are keyed by property name only

- Where: [custom_functions.py:279](custom_functions.py#L279), [custom_functions.py:1011](custom_functions.py#L1011)
- Cause: `property_description_dict` is a module-level dict with the property name as key.
- Effect: `Pset_WallCommon.FireRating` and `Pset_DoorCommon.FireRating` get the same description,
  the last one read. The dict also keeps its values if the function runs twice in one process.
- Fix: Key the dict by `(PropertySet, Property)` and create it inside `excel_to_spec_list`.

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
