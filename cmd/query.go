package cmd

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/noborus/trdsql"
	"github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

func exec(args []string) error {
	if Debug {
		trdsql.EnableDebug()
	}
	query := strings.Join(args, " ")

	writer, err := setWriter(OutFileName)
	if err != nil {
		return err
	}
	trd := trdsql.NewTRDSQL(
		trdsql.NewImporter(
			trdsql.InSkip(Skip),
			trdsql.InHeader(Header),
			trdsql.InPreRead(100),
		),
		trdsql.NewExporter(writer),
	)
	return trd.Exec(query)
}

func setWriter(fileName string) (trdsql.Writer, error) {
	dotExt := strings.TrimLeft(filepath.Ext(fileName), ".")
	if OutFormat == "GUESS" && dotExt != "" {
		OutFormat = strings.ToUpper(dotExt)
	}

	if strings.ToUpper(OutFormat) != "XLSX" {
		return stdWriter(fileName)
	}

	// XLSX Writer
	if fileName == "" {
		return nil, fmt.Errorf("file name is required to output with XLSX")
	}

	if OutSheetName == "" {
		OutSheetName = "Sheet1"
	}

	writer, err := xlsxsql.NewXLSXWriter(
		xlsxsql.FileName(fileName),
		xlsxsql.Sheet(OutSheetName),
		xlsxsql.Cell(OutCell),
		xlsxsql.ClearSheet(ClearSheet),
		xlsxsql.Header(OutHeader),
	)
	if err != nil {
		return nil, err
	}

	return writer, nil
}

func stdWriter(fileName string) (trdsql.Writer, error) {
	var file io.Writer
	if fileName == "" {
		file = os.Stdout
	} else {
		f, err := os.Create(fileName)
		if err != nil {
			return nil, err
		}
		file = f
	}
	format := trdsql.OutputFormat(strings.ToUpper(OutFormat))
	w := trdsql.NewWriter(
		trdsql.OutStream(file),
		trdsql.OutHeader(OutHeader),
		trdsql.OutFormat(format),
	)
	return w, nil
}

// queryCmd represents the query command
var queryCmd = &cobra.Command{
	Use:   "query",
	Short: "Executes the specified SQL query against the xlsx file",
	Long: `Executes the specified SQL query against the xlsx file.
Output to CSV and various formats.`,
	RunE: func(cmd *cobra.Command, args []string) error {
		return exec(args)
	},
}

func init() {
	rootCmd.AddCommand(queryCmd)
}
