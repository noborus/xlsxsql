package cmd

import (
	"fmt"
	"os"
	"strings"

	_ "github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

// rootCmd represents the base command when called without any subcommands.
var rootCmd = &cobra.Command{
	Use:   "xlsxsql",
	Short: "Execute SQL against xlsx file.",
	Long: `Execute SQL against xlsx file.
Output to CSV and various formats.`,
	SilenceUsage: true,
	Run: func(cmd *cobra.Command, args []string) {
		if Ver {
			fmt.Printf("xlsxsql version %s rev:%s\n", Version, Revision)
			return
		}
		if err := cmd.Help(); err != nil {
			cmd.SetOutput(os.Stderr)
			os.Exit(1)
		}
	},
}

var (
	// Version represents the version.
	Version string
	// Revision set "git rev-parse --short HEAD"
	Revision string
)

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute(version string, revision string) {
	Version = version
	Revision = revision
	if err := rootCmd.Execute(); err != nil {
		rootCmd.SetOutput(os.Stderr)
		os.Exit(1)
	}
}

// Ver is version information.
var Ver bool

// Debug is debug mode.
var Debug bool

// OutFormat is output format.
var OutFormat = "CSV"

// OutHeader is output header.
var OutHeader bool

// Skip is skip lines.
var Skip int

// Header is input header.
var Header bool

// OutFileName is output file name.
var OutFileName string

// OutSheetName is output sheet name.
var OutSheetName string

// OutCell is output cell name.
var OutCell string

// ClearSheet is clear sheet if true.
var ClearSheet bool

func init() {
	rootCmd.PersistentFlags().BoolVarP(&Ver, "version", "v", false, "display version information")
	rootCmd.PersistentFlags().BoolVarP(&Debug, "debug", "", false, "debug mode")

	// Input
	rootCmd.PersistentFlags().IntVarP(&Skip, "skip", "s", 0, "Skip the number of lines")
	rootCmd.PersistentFlags().BoolVarP(&Header, "header", "H", false, "Input header")
	// Output
	validOutFormats := []string{"CSV", "AT", "LTSV", "JSON", "JSONL", "TBLN", "RAW", "MD", "VF", "YAML", "XLSX"}
	outputFormats := fmt.Sprintf("Output Format[%s]", strings.Join(validOutFormats, "|"))
	rootCmd.PersistentFlags().StringVarP(&OutFormat, "out", "o", "GUESS", outputFormats)
	// Register the completion function for the --out flag
	_ = rootCmd.RegisterFlagCompletionFunc("out", func(cmd *cobra.Command, args []string, toComplete string) ([]string, cobra.ShellCompDirective) {
		return validOutFormats, cobra.ShellCompDirectiveDefault
	})
	rootCmd.PersistentFlags().StringVarP(&OutFileName, "out-file", "O", "", "File name to output to file")
	rootCmd.PersistentFlags().BoolVarP(&OutHeader, "out-header", "", false, "Output header")
	rootCmd.PersistentFlags().StringVarP(&OutSheetName, "out-sheet", "", "", "Sheet name to output to xlsx file")
	rootCmd.PersistentFlags().StringVarP(&OutCell, "out-cell", "", "", "Cell name to output to xlsx file")
	rootCmd.PersistentFlags().BoolVarP(&ClearSheet, "clear-sheet", "", false, "Clear sheet when outputting to xlsx file")
}
