# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [vNext] Unreleased

## [0.1.0] 2018-01-30

### Added

- Commit and Merged to master.
- Help file for Compare-VesterOutput.
- Composed Readme.

### Changed

- Renamed Function/Script name Compare-ScriptOutput to Compare-VesterOutput.
- Changed xlsxfile parameter to mandatory.

### Removed

- Removed Get-FileName private function due to set xlsxfile parameter to mandatory.

## [0.0.3] 2018-01-30

### Added

- Compare-ScriptOutput.ps1 implemented function process() part to fill up excel file.
- Private function Get-ExcelColumnInt convert-an-excel-column-letter-into-its-number.

## [0.0.2] 2018-01-30

### Added

- Compare-ScriptOutput.ps1 implemented function begin() part to Vester result to a variable.

### Changed

- VMwareBaselineCheck.psm1 changed public and private folder

### Fixed

- Tests file

## [0.0.1] 2018-01-20

- initial release with scaffold
