php_writeexcel
==============

PHP port of John McNamara's Spreadsheet::WriteExcel by Johann Hanne,
with some tweaks by Thomas Roche (https://github.com/thoroc/php_writeexcel),
and forked by Craig Manley adding more tweaks and some added drop-in compatibility with PEAR's Spreadsheet_Excel_Writer.

Dependencies
============
To use this library, please install [php_ole](https://github.com/thoroc/php_ole) which allow to write big Excel files (larger than 7MB).


Example for your `composer.json` file:

```
{
    "minimum-stability": "dev",
    "repositories": [
      {
        "type": "vcs",
        "url": "https://github.com/thoroc/php_ole"
      },
      {
        "type": "vcs",
        "url": "https://github.com/thoroc/php_writeexcel"
      }
    ],
    "require": {
        "thoroc/php_ole": "master",
        "thoroc/php_writeexcel": "master",
    }
}
```

Perfomance comparison
=====================
Update by Craig Manley, 2015-12-06:
Writing a table of 8 fields and 950 rows to an Excel file resulted in these average elapsed times:
```
PHPExcel: 7.417
PEAR Spreadsheet_Excel_Writer: 1.078
php_writeexcel: 2.065
```
So clearly, PEAR's Spreadsheet_Excel_Writer is almost twice as fast.
PHPExcel is more modern, but by far the slowest and consumes huge amounts of memory,
which is the only reason I tried with this library, updating it slightly in the process.
