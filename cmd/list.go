package cmd

import (
	"fmt"

	"github.com/noborus/xlsxsql"
	"github.com/spf13/cobra"
)

// listSheetsCmd represents the list command.
var listSheetsCmd = &cobra.Command{
	Use:   "list",
	Short: "List the sheets of the xlsx file",
	Long:  `List the sheets of the xlsx file.`,
	RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 1 {
			return fmt.Errorf("please specify the xlsx file")
		}
		list, err := xlsxsql.XLSXSheet(args[0])
		if err != nil {
			fmt.Println(err)
			return err
		}
		for _, v := range list {
			fmt.Println(v)
		}
		return nil
	},
}

func init() {
	rootCmd.AddCommand(listSheetsCmd)
}
