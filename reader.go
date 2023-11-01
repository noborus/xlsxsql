package xlsxsql

import (
	"fmt"
	"io"

	"github.com/noborus/trdsql"
	"github.com/xuri/excelize/v2"
)

type XLSXReader struct {
	reader    *excelize.File
	tableName string
	names     []string
	types     []string
	body      [][]interface{}
}

func NewXLSXReader(reader io.Reader, opts *trdsql.ReadOpts) (trdsql.Reader, error) {
	r := XLSXReader{}

	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}

	sheet := "Sheet1"
	if len(opts.InJQuery) > 0 {
		sheet = opts.InJQuery
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	r.tableName = sheet

	for j, row := range rows {
		if j == 0 {
			r.names = make([]string, len(row))
			r.types = make([]string, len(row))
			for i := 0; i < len(row); i++ {
				r.names[i] = fmt.Sprintf("C%d", i+1)
				r.types[i] = "text"
			}
		}

		data := make([]interface{}, len(row))
		for i, colCell := range row {
			data[i] = colCell
		}
		r.body = append(r.body, data)
	}
	r.reader = f
	return r, nil
}

func (t XLSXReader) Names() ([]string, error) {
	return t.names, nil
}

func (t XLSXReader) Types() ([]string, error) {
	return t.types, nil
}

func (t XLSXReader) PreReadRow() [][]interface{} {
	return t.body
}

// ReadRow only returns EOF.
func (t XLSXReader) ReadRow(row []interface{}) ([]interface{}, error) {
	return nil, io.EOF
}

func XLSXSheet(fileName string) ([]string, error) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	return f.GetSheetList(), nil
}

func init() {
	trdsql.RegisterReaderFunc("XLSX", NewXLSXReader)
}
