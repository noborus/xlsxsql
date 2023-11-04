# xlsxsql

Execute SQL on xlsx file.

Execute SQL on xlsx files using [xcelize](https://github.com/qax-os/excelize) and [trdsql](https://github.com/noborus/trdsql).
Output to various formats.

## Install

```console
go install github.com/noborus/xlsxsql/cmd/xlsxsql@latest
```

## Usage

```console
$ xlsxsql --help
output to CSV and various formats.

Usage:
  xlsxsql [flags]
  xlsxsql [command]

Available Commands:
  completion  Generate the autocompletion script for the specified shell
  help        Help about any command
  list        List the sheets of the xlsx file
  query       Executes the specified SQL query against the xlsx file
  table       SQL(SELECT * FROM table) for xlsx

Flags:
  -H, --header       Output header
  -h, --help         help for xlsxsql
  -o, --out string   Output Format[CSV|AT|LTSV|JSON|JSONL|TBLN|RAW|MD|VF|YAML] (default "CSV")
  -s, --skip int     Skip the number of lines
  -v, --version      display version information

Use "xlsxsql [command] --help" for more information about a command.
```

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
