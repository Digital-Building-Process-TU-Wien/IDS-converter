# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

The project has no tagged release yet. Changes before this file existed are in the git history of `main`.

## [Unreleased]

### Fixed

- An enumeration that mixes values with and without `PredefinedType`, such as `IFCWALL|IFCWALL.STANDARD`, now gives a clear error message instead of `IndexError`.
- Property descriptions are now kept per property set. Two properties with the same name in different property sets no longer get the same description. A property without its own description still uses the last description of the same property name.
- Property descriptions of one run no longer carry over to a second run in the same process.
- A sheet without any known column, for example after a wrong sheet name or a wrong `Skipped rows` value, now gives a clear error message instead of `NameError`.
- Requirement values of one separated IDS file no longer leak into another file when a specification belongs to several `Phase`, `Role` or `Usecase` values.
- Specifications with different `File separators` categories, such as `Phase=1` and `Role=1`, no longer end up in the same IDS file.
- Output file names have the separator again, for example `..._Specifications_PhaseLP3.ids` instead of `..._SpecificationsLP3.ids`.
- Merged rows with a `pattern=` value that contains `|` now raise the "complex restriction in an enumeration" error, in any row order. Before, the pattern was joined or split wrongly.
- Column `R.MaterialURI` is now read. The Material facet gets its `uri` attribute.
- A sheet without the columns `R.PropertyValue` and `R.AttributeValue` now converts. Before, it stopped with `No objects to concatenate`.
- The order of the purpose, the milestone and the IFC versions is now the order of the first occurrence in the sheet. Before, it changed from run to run.

### Security

- Raised `deepdiff` from 8.3.0 to 8.6.2 (and `orderly-set` from 5.3.0 to 5.5.0). The old version has two known flaws, CVE-2025-58367 and CVE-2026-33155, in its `Delta` class and `pickle_load`. The converter does not use them, so it was not affected.

[unreleased]: https://github.com/pkoska/IDS-converter/compare/main...dev
