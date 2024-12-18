package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

// tableCmd represents the table command.
var tableCmd = &cobra.Command{
	Use:   "table",
	Short: "SQL(SELECT * FROM table) for xlsx",
	Long:  `Execute SELECT * FROM table, assuming the xlsx file as a table.`,
	RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 1 {
			return fmt.Errorf("please specify the xlsx file")
		}
		query := "SELECT * FROM " + args[0]
		return exec([]string{query})
	},
}

func init() {
	rootCmd.AddCommand(tableCmd)
}
