package xlsxsql

import (
	"io"
	"os"

	"github.com/noborus/trdsql"
	"github.com/xuri/excelize/v2"
)

// XLSXWriter is a writer for XLSX files.
type XLSXWriter struct {
	fileName string
	f        *excelize.File
	sheet    string
	rowID    int
}

func clear(f *excelize.File, sheet string) error {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return err
	}

	for i, row := range rows {
		for j := range row {
			axis, err := excelize.CoordinatesToCellName(j+1, i+1)
			if err != nil {
				return err
			}
			f.SetCellStr(sheet, axis, "")
		}
	}
	return nil
}

// NewXLSXWriter function takes an io.Writer and trdsql.WriteOpts, and returns a new XLSXWriter.
func NewXLSXWriter(writer io.Writer, fileName string, sheet string, clearSheet bool) (trdsql.Writer, error) {
	var f *excelize.File
	var err error

	if _, err = os.Stat(fileName); os.IsNotExist(err) {
		// File does not exist, create a new one
		f = excelize.NewFile()
	} else {
		// File exists, open it
		f, err = excelize.OpenFile(fileName)
		if err != nil {
			return nil, err
		}
	}

	n, err := f.GetSheetIndex(sheet)
	if err != nil {
		return nil, err
	}
	if clearSheet {
		if n >= 0 {
			// Sheet exists,clear it
			if err := clear(f, sheet); err != nil {
				return nil, err
			}
		}
	}
	if n < 0 {
		// Sheet does not exist, create a new one
		if _, err := f.NewSheet(sheet); err != nil {
			return nil, err
		}
	}

	return &XLSXWriter{
		fileName: fileName,
		f:        f,
		sheet:    sheet,
	}, nil
}

// PreWrite function opens the XLSXWriter.
func (w *XLSXWriter) PreWrite(columns []string, types []string) error {
	for i, v := range columns {
		cell, error := excelize.CoordinatesToCellName(i+1, 1)
		if error != nil {
			return error
		}
		w.f.SetCellValue(w.sheet, cell, v)
	}
	w.rowID++
	return nil
}

// WriteRow function writes a row to the XLSXWriter.
func (w *XLSXWriter) WriteRow(row []interface{}, columns []string) error {
	w.rowID++
	for i, v := range row {
		if v == nil {
			continue
		}
		cell, error := excelize.CoordinatesToCellName(i+1, w.rowID)
		if error != nil {
			return error
		}
		w.f.SetCellValue(w.sheet, cell, v)
	}
	return nil
}

// PostWrite function closes the XLSXWriter.
func (w *XLSXWriter) PostWrite() error {
	if err := w.f.SaveAs(w.fileName); err != nil {
		return err
	}
	return w.f.Close()
}
