php_writeexcel
==============

Johann Hanne's php lib to write excel file


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

