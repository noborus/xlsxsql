package cmd

import (
	"fmt"

	"github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

// sheetsCmd represents the sheets command
var listSheetsCmd = &cobra.Command{
	Use:   "list",
	Short: "List the sheets of the xlsx file",
	Run: func(cmd *cobra.Command, args []string) {
		list, err := xlsxsql.XLSXSheet(args[0])
		if err != nil {
			fmt.Println(err)
			return
		}
		for _, v := range list {
			fmt.Println(v)
		}
	},
}

func init() {
	rootCmd.AddCommand(listSheetsCmd)
}
