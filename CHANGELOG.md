# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

The project has no tagged release yet. Changes before this file existed are in the git history of `main`.

## [Unreleased]

### Fixed

- A sheet without any known column, for example after a wrong sheet name or a wrong `Skipped rows` value, now gives a clear error message instead of `NameError`.
- Requirement values of one separated IDS file no longer leak into another file when a specification belongs to several `Phase`, `Role` or `Usecase` values.
- Specifications with different `File separators` categories, such as `Phase=1` and `Role=1`, no longer end up in the same IDS file.
- Output file names have the separator again, for example `..._Specifications_PhaseLP3.ids` instead of `..._SpecificationsLP3.ids`.
- Merged rows with a `pattern=` value that contains `|` now raise the "complex restriction in an enumeration" error, in any row order. Before, the pattern was joined or split wrongly.
- Column `R.MaterialURI` is now read. The Material facet gets its `uri` attribute.

[unreleased]: https://github.com/pkoska/IDS-converter/compare/main...dev
