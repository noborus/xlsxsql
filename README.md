# xlsxsql

Execute SQL on xlsx file.

Execute SQL on xlsx files using [xcelize](https://github.com/qax-os/excelize) and [trdsql](https://github.com/noborus/trdsql).
Output to various formats.

## Install

```console
go install github.com/noborus/xlsxsql/cmd/xlsxsql@latest
```

### Binary Downloads

Precompiled binaries for xlsxsql are available for various platforms and architectures. You can download them from the [GitHub Releases](https://github.com/noborus/xlsxsql/releases) page.

Here are the available binaries:

- `xlsxsql_Darwin_arm64.tar.gz`
- `xlsxsql_Darwin_x86_64.tar.gz`
- `xlsxsql_Linux_arm64.tar.gz`
- `xlsxsql_Linux_i386.tar.gz`
- `xlsxsql_Linux_x86_64.tar.gz`
- `xlsxsql_Windows_arm64.zip`
- `xlsxsql_Windows_x86_64.zip`

To install a binary, download the appropriate file for your system, extract it, and place the `xlsxsql` executable in a directory included in your system's `PATH`.

For example, on a Unix-like system, you might do:

```console
tar xvf xlsxsql_Darwin_x86_64.tar.gz
mv xlsxsql /usr/local/bin/
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
Sheet1
Sheet2
```

### Basic usage

The basic usage of xlsxsql is to run a SQL query against an Excel file.
The `query` command is used followed by the SQL query in quotes.
The SQL query should include the name of the Excel file. If no sheet is specified, the first sheet will be targeted.

```console
xlsxsql query "SELECT * FROM test.xlsx"
```

For example, if test.xlsx contains the following data in its first sheet:

| Name  | Age |
| ----- | --- |
| Alice | 20  |
| Bob   | 25  |
| Carol | 30  |

The output will be:

```csv
Name,Age
Alice,20
Bob,25
Carol,30
```

### Specify sheet

```console
xlsxsql query "SELECT * FROM test.xlsx::Sheet2"
```

### Shorthand designation

The `table` command is a shorthand that allows you to quickly display the contents of a specified sheet in a table format.
The syntax is `xlsxsql table <filename>::<sheetname>`.
If no sheet name is specified, the first sheet of the Excel file will be targeted.

Here is an example:

```console
xlsxsql table test.xlsx::Sheet2
```

### Output format

```console
xlsxsql query --out JSONL "SELECT * FROM test.xlsx::Sheet2"
```

You can choose from CSV, LTSV, JSON, JSONL, TBLN, RAW, MD, VF, YAML.

### Skip and Header Options

You can use the `--skip` option to ignore a certain number of rows at the beginning of the sheet. 
For example, to skip the first two rows, you would use:

```console
xlsxsql query --skip 2 --out JSONL "SELECT * FROM test.xlsx::Sheet2"
```

The `--header` option treats the first row (excluding any rows ignored by `--skip``) as the header.
For example, you would use it like this:

```console
xlsxsql query --header --out JSONL "SELECT * FROM test.xlsx::Sheet2"
```
