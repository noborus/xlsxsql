package cmd

import (
	"strings"

	"github.com/noborus/trdsql"
	"github.com/spf13/cobra"
)

func exec(args []string) error {
	//trdsql.EnableDebug()
	query := strings.Join(args, " ")
	format := trdsql.OutputFormat(strings.ToUpper(OutFormat))

	trd := trdsql.NewTRDSQL(
		trdsql.NewImporter(
			trdsql.InSkip(Skip),
			trdsql.InHeader(Header),
			trdsql.InPreRead(100),
		),
		trdsql.NewExporter(
			trdsql.NewWriter(
				trdsql.OutHeader(OutHeader),
				trdsql.OutFormat(format),
			),
		),
	)
	return trd.Exec(query)
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
