# xlsxsql

Execute SQL on xlsx file.

## Install

```console
go install github.com/noborus/xlsxsql/cmd/xlsxsql@latest
```

## Usage

### list sheets

```console
$ xlsxsql list test.xlsx
```

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

### Output format

```console
$ xlsxsql query --out JSONL "SELECT * FROM test.xlsx::Sheet2"
```

You can choose from csv, ltsv, json, jsonl, yaml, at, md, vf, and tbln formats.