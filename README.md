# xlsxsql

Execute SQL on xlsx file.

## Usage

### Basic usage

```console
$ xlsxsql query "SELECT * FROM test.xlsx"
```

If no sheet is specified, the first sheet will be targeted.

### Specify sheet

```console
$ xlsxsql query "SELECT * FROM test.xlsx::Sheet2"
```

### Shorthand designation

```console
$ xlsxsql table test.xlsx::Sheet2
```

### list sheets

```console
$ xlsxsql list test.xlsx
```