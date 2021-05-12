# Excel Export Template
Simple class help export data with exist template.

## Requirement
- PHPSpreadsheet library

## Options
Symbols
```
[[var_name]] => for two-dimensional arrays
[var_name] => for one-dimensional arrays
{var_name] => for string values
```
## Output
For download
```
$target = 'ExcelExportTemplate' . time() . '.xlsx';
And
'output' => 'download'
```
For save file
```
$target = './ExcelExportTemplate' . time() . '.xlsx';
And
'output' => 'file'
```
## Example
See export.php file
