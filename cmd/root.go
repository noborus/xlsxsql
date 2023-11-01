package cmd

import (
	"io"
	"os"
	"strings"

	"github.com/noborus/trdsql"
	_ "github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

var fileName string
var Sheet string
var OutFormat = "CSV"

func outFormat(outStream io.Writer, errStream io.Writer) trdsql.Writer {
	var format trdsql.Format
	switch strings.ToUpper(OutFormat) {
	case "CSV":
		format = trdsql.CSV
	case "LTSV":
		format = trdsql.LTSV
	case "JSON":
		format = trdsql.JSON
	case "TBLN":
		format = trdsql.TBLN
	case "RAW":
		format = trdsql.RAW
	case "MD":
		format = trdsql.MD
	case "AT":
		format = trdsql.AT
	case "VF":
		format = trdsql.VF
	case "JSONL":
		format = trdsql.JSONL
	case "YAML":
		format = trdsql.YAML
	default:
		format = trdsql.AT
	}
	w := trdsql.NewWriter(
		trdsql.OutFormat(format),
		trdsql.OutStream(outStream),
		trdsql.ErrStream(errStream),
	)
	return w
}

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "xlsxsql",
	Short: "A brief description of your application",
	Long: `A longer description that spans multiple lines and likely contains
examples and usage of using your application. For example:

Cobra is a CLI library for Go that empowers applications.
This application is a tool to generate the needed files
to quickly create a Cobra application.`,
	Run: func(cmd *cobra.Command, args []string) {
		w := outFormat(os.Stdout, os.Stderr)
		query := strings.Join(args, " ")

		trd := trdsql.NewTRDSQL(
			trdsql.NewImporter(),
			trdsql.NewExporter(
				w,
			),
		)
		trd.Exec(query)
	},
}

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {
	cobra.CheckErr(rootCmd.Execute())
}

func init() {
	rootCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
	rootCmd.Flags().StringVarP(&fileName, "file", "F", "", "XLSX File Name")
	rootCmd.Flags().StringVarP(&Sheet, "sheet", "s", "", "Sheet Name")
	rootCmd.Flags().StringVarP(&OutFormat, "format", "f", "CSV", "Output Format")
}
