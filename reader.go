package xlsxsql

import (
	"fmt"
	"io"

	"github.com/noborus/trdsql"
	"github.com/xuri/excelize/v2"
)

type XLSXReader struct {
	tableName string
	names     []string
	types     []string
	body      [][]interface{}
}

// NewXLSXReader returns a trdsql.Reader interface for XLSX.
func NewXLSXReader(reader io.Reader, opts *trdsql.ReadOpts) (trdsql.Reader, error) {
	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}

	sheet, err := getSheet(f, opts.InJQuery)
	if err != nil {
		return nil, err
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	r := XLSXReader{}
	r.tableName = sheet
	skip := opts.InSkip
	columnNum := 0
	header := 0
	for i := 0; i < len(rows); i++ {
		if i < skip {
			header = i + 1
			continue
		}
		row := rows[i]
		columnNum = max(columnNum, len(row))
		if i > opts.InPreRead {
			break
		}
	}
	if header > len(rows) {
		header = 0
	} else {
		if opts.InHeader {
			skip++
		}
	}
	r.names = make([]string, columnNum)
	r.types = make([]string, columnNum)
	for i := 0; i < columnNum; i++ {
		if opts.InHeader && len(rows[header]) > i && rows[header][i] != "" {
			r.names[i] = rows[header][i]
		} else {
			r.names[i] = fmt.Sprintf("C%d", i+1)
		}
		r.types[i] = "text"
	}

	r.body = make([][]interface{}, 0, len(rows)-skip)
	for j, row := range rows {
		if j < skip {
			continue
		}
		data := make([]interface{}, columnNum)
		for i, colCell := range row {
			data[i] = colCell
		}
		r.body = append(r.body, data)
	}

	return r, nil
}

func getSheet(f *excelize.File, sheet string) (string, error) {
	list := f.GetSheetList()
	if len(list) == 0 {
		return "", fmt.Errorf("sheet not found")
	}
	if sheet == "" {
		sheet = list[0]
	}
	for _, s := range list {
		if s == sheet {
			return s, nil
		}
	}
	return "", fmt.Errorf("sheet not found")
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

// XLSXSheet returns the sheet name of the XLSX file.
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
