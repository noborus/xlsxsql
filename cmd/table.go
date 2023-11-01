package cmd

import (
	"github.com/spf13/cobra"
)

// tableCmd represents the table command
var tableCmd = &cobra.Command{
	Use:   "table",
	Short: "sql(SELECT * FROM table) for xlsx",
	Long:  ``,
	RunE: func(cmd *cobra.Command, args []string) error {
		query := "SELECT * FROM " + args[0]
		return exec([]string{query})
	},
}

func init() {
	rootCmd.AddCommand(tableCmd)
}
