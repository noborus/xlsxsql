package cmd

import (
	"strings"

	"github.com/noborus/trdsql"
	"github.com/spf13/cobra"
)

var OutFormat = "CSV"

func outFormat() trdsql.Format {
	switch strings.ToUpper(OutFormat) {
	case "CSV":
		return trdsql.CSV
	case "LTSV":
		return trdsql.LTSV
	case "JSON":
		return trdsql.JSON
	case "TBLN":
		return trdsql.TBLN
	case "RAW":
		return trdsql.RAW
	case "MD":
		return trdsql.MD
	case "AT":
		return trdsql.AT
	case "VF":
		return trdsql.VF
	case "JSONL":
		return trdsql.JSONL
	case "YAML":
		return trdsql.YAML
	default:
		return trdsql.CSV
	}
}

func exec(args []string) error {
	query := strings.Join(args, " ")
	format := outFormat()

	trd := trdsql.NewTRDSQL(
		trdsql.NewImporter(),
		trdsql.NewExporter(
			trdsql.NewWriter(
				trdsql.OutFormat(format),
			),
		),
	)
	return trd.Exec(query)
}

// queryCmd represents the query command
var queryCmd = &cobra.Command{
	Use:   "query",
	Short: "query for xlsx",
	Long:  `query for xlsx`,
	RunE: func(cmd *cobra.Command, args []string) error {
		return exec(args)
	},
}

func init() {
	rootCmd.AddCommand(queryCmd)
}
