// Package xlsxsql provides a reader for XLSX files.
// It uses the trdsql and excelize/v2 packages to read XLSX files and convert them into SQL tables.
// The main type is XLSXReader, which implements the trdsql.Reader interface.
package xlsxsql

import (
	"fmt"
	"io"
	"strings"

	"github.com/noborus/trdsql"
	"github.com/xuri/excelize/v2"
)

var (
	ErrSheetNotFound = fmt.Errorf("sheet not found")
	ErrNoData        = fmt.Errorf("no data")
)

// XLSXReader is a reader for XLSX files.
type XLSXReader struct {
	names []string
	types []string
	body  [][]interface{}
}

// NewXLSXReader function takes an io.Reader and trdsql.ReadOpts, and returns a new XLSXReader.
// It reads the XLSX file, retrieves the sheet specified by the InJQuery option, and reads the rows into the XLSXReader.
func NewXLSXReader(reader io.Reader, opts *trdsql.ReadOpts) (trdsql.Reader, error) {
	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}
	extSheet, extCell := parseExtend(opts.InJQuery)
	sheet, err := getSheet(f, extSheet)
	if err != nil {
		return nil, err
	}

	isFilter := false
	cellX, cellY := 0, 0
	if extCell != "" {
		cellX, cellY, err = excelize.CellNameToCoordinates(extCell)
		if err != nil {
			return nil, err
		}
		isFilter = true
		cellX--
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	r := XLSXReader{}
	skip := 0
	if cellY > 0 {
		skip = cellY - 1
	}
	if opts.InSkip > 0 {
		skip = opts.InSkip
	}

	columnNum := 0
	header := 0
	for i := 0; i < len(rows); i++ {
		if i < skip {
			header = i + 1
			continue
		}
		row := rows[i]
		columnNum = max(columnNum, len(row)-cellX)
		if i > opts.InPreRead {
			break
		}
	}
	if columnNum <= 0 {
		return nil, ErrNoData
	}

	if header > len(rows) {
		header = 0
	} else {
		if opts.InHeader {
			skip++
		}
	}

	r.names, r.types = nameType(rows[header], cellX, columnNum, opts.InHeader)
	rowNum := len(rows) - skip
	body := make([][]interface{}, 0, rowNum)
	validColumns := make([]bool, columnNum)
	for i := 0; i < len(r.names); i++ {
		if r.names[i] != "" {
			validColumns[i] = true
		} else {
			r.names[i] = fmt.Sprintf("C%d", i+1)
		}
	}
	for j, row := range rows {
		if j < skip {
			continue
		}
		data := make([]interface{}, columnNum)
		for c, i := 0, cellX; i < len(row); i++ {
			data[c] = row[i]
			if data[c] != "" {
				validColumns[c] = true
			}
			c++
		}
		body = append(body, data)
	}

	if !isFilter {
		r.body = body
		return r, nil
	}

	r.body = filterColumns(body, validColumns)
	r.names = r.names[:len(r.body[0])]
	r.types = r.types[:len(r.body[0])]
	return r, nil
}

func filterColumns(body [][]interface{}, dataFlag []bool) [][]interface{} {
	count := len(dataFlag)
	start := false
	for i, f := range dataFlag {
		if f {
			start = true
		}
		if start && !f {
			count = i
			break
		}
	}
	newBody := make([][]interface{}, 0, len(body))
	for _, row := range body {
		cols := make([]interface{}, count)
		for i := 0; i < count; i++ {
			cols[i] = row[i]
		}
		newBody = append(newBody, cols)
	}
	return newBody
}

func parseExtend(ext string) (string, string) {
	e := strings.Split(ext, ".")
	if len(e) == 1 {
		return e[0], ""
	} else if len(e) == 2 {
		return e[0], e[1]
	} else {
		return e[0], strings.Join(e[1:], ".")
	}
}

func nameType(row []string, cellX int, columnNum int, header bool) ([]string, []string) {
	nameMap := make(map[string]bool)
	names := make([]string, columnNum)
	types := make([]string, columnNum)
	c := 0
	for i := cellX; i < cellX+columnNum; i++ {
		if header && len(row) > i && row[i] != "" {
			if _, ok := nameMap[row[i]]; ok {
				names[c] = fmt.Sprintf("C%d", i+1)
			} else {
				nameMap[row[i]] = true
				names[c] = row[i]
			}
		} else {
			names[c] = ""
		}
		types[c] = "text"
		c++
	}
	return names, types
}

func getSheet(f *excelize.File, sheet string) (string, error) {
	list := f.GetSheetList()
	if len(list) == 0 {
		return "", ErrSheetNotFound
	}
	if sheet == "" {
		sheet = list[0]
	}
	for _, s := range list {
		if s == sheet {
			return s, nil
		}
	}
	return "", ErrSheetNotFound
}

// Names returns the column names of the XLSX file.
func (r XLSXReader) Names() ([]string, error) {
	return r.names, nil
}

// Types returns the column types of the XLSX file.
func (r XLSXReader) Types() ([]string, error) {
	return r.types, nil
}

// PreReadRow returns the rows of the XLSX file.
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
	// Use XLSXReader for extension xlsx.
	trdsql.RegisterReaderFunc("XLSX", NewXLSXReader)
}
