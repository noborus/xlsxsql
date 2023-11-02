package cmd

import (
	"fmt"
	"os"

	_ "github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:          "xlsxsql",
	Short:        "sql for xlsx",
	Long:         `sql for xlsx`,
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

func init() {
	rootCmd.PersistentFlags().BoolVarP(&Ver, "version", "v", false, "display version information")
	rootCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
	rootCmd.PersistentFlags().StringVarP(&OutFormat, "out", "o", "CSV", "Output Format")
}
