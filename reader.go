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
	trdsql.EnableDebug()

	r := XLSXReader{}
	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}

	list := f.GetSheetList()
	sheet := list[0]
	if len(opts.InJQuery) > 0 {
		sheet = opts.InJQuery
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	r.tableName = sheet

	columnNum := 0
	for i := 0; i < len(rows); i++ {
		row := rows[i]
		columnNum = max(columnNum, len(row))
		if i > opts.InPreRead {
			break
		}
	}

	r.names = make([]string, columnNum)
	r.types = make([]string, columnNum)
	for i := 0; i < columnNum; i++ {
		if opts.InHeader && len(rows[0]) > i && rows[0][i] != "" {
			r.names[i] = rows[0][i]
		} else {
			r.names[i] = fmt.Sprintf("C%d", i+1)
		}
		r.types[i] = "text"
	}

	for j, row := range rows {
		if j == 0 && opts.InHeader {
			continue
		}
		data := make([]interface{}, columnNum)
		for i, colCell := range row {
			data[i] = colCell
		}
		r.body = append(r.body, data)
	}

	r.reader = f
	return r, nil
}
func (r XLSXReader) Names() ([]string, error) {
	return r.names, nil
}

func (r XLSXReader) Types() ([]string, error) {
	return r.types, nil
}

func (r XLSXReader) PreReadRow() [][]interface{} {
	return r.body
}

// ReadRow only returns EOF.
func (r XLSXReader) ReadRow(row []interface{}) ([]interface{}, error) {
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
