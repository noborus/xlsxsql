package cmd

import (
	"fmt"
	"os"

	_ "github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

// rootCmd represents the base command when called without any subcommands
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
		cmd.Help()
	},
}

var (
	// Version represents the version
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
		rootCmd.Println(err)
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

var Skip int

var Header bool

var OutFileName string
var OutSheetName string

var ClearSheet bool

func init() {
	rootCmd.PersistentFlags().BoolVarP(&Ver, "version", "v", false, "display version information")
	rootCmd.PersistentFlags().BoolVarP(&Debug, "debug", "", false, "debug mode")
	rootCmd.PersistentFlags().IntVarP(&Skip, "skip", "s", 0, "Skip the number of lines")
	rootCmd.PersistentFlags().BoolVarP(&Header, "header", "H", false, "Input header")
	rootCmd.PersistentFlags().BoolVarP(&OutHeader, "out-header", "O", false, "Output header")
	rootCmd.PersistentFlags().StringVarP(&OutFormat, "out", "o", "GUESS", "Output Format[CSV|AT|LTSV|JSON|JSONL|TBLN|RAW|MD|VF|YAML]")
	rootCmd.PersistentFlags().StringVarP(&OutFileName, "out-xlsx", "x", "", "File name to output to xlsx file")
	rootCmd.PersistentFlags().StringVarP(&OutSheetName, "out-sheet", "S", "", "Sheet name to output to xlsx file")
	rootCmd.PersistentFlags().BoolVarP(&ClearSheet, "clear-sheet", "D", false, "Clear sheet when outputting to xlsx file")
}
