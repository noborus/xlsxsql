package xlsxsql

import (
	"io"
	"os"

	"github.com/xuri/excelize/v2"
)

// XLSXWriter is a writer for XLSX files.
type XLSXWriter struct {
	fileName string
	f        *excelize.File
	sheet    string
	header   bool
	cellX    int
	cellY    int
	rowID    int
}

// WriteOpts represents options that determine the behavior of the writer.
type WriteOpts struct {
	// ErrStream is the error output destination.
	ErrStream io.Writer
	// FileName is the output file name.
	FileName string
	// Sheet is the sheet name.
	Sheet string
	// Cell is the cell name.
	Cell string
	// ClearSheet is the flag to clear the sheet.
	ClearSheet bool
	// WriteHeader is the flag to write the header.
	Header bool
}

// WriteOpt is a function to set WriteOpts.
type WriteOpt func(*WriteOpts)

// ErrStream sets the error output destination.
func ErrStream(f io.Writer) WriteOpt {
	return func(args *WriteOpts) {
		args.ErrStream = f
	}
}

// FileName sets the output file name.
func FileName(f string) WriteOpt {
	return func(args *WriteOpts) {
		args.FileName = f
	}
}

// Sheet sets the sheet name.
func Sheet(f string) WriteOpt {
	return func(args *WriteOpts) {
		args.Sheet = f
	}
}

// Cell sets the cell name.
func Cell(f string) WriteOpt {
	return func(args *WriteOpts) {
		args.Cell = f
	}
}

// ClearSheet sets the flag to clear the sheet.
func ClearSheet(f bool) WriteOpt {
	return func(args *WriteOpts) {
		args.ClearSheet = f
	}
}

func Header(f bool) WriteOpt {
	return func(args *WriteOpts) {
		args.Header = f
	}
}

// NewXLSXWriter function takes an io.Writer and trdsql.WriteOpts, and returns a new XLSXWriter.
func NewXLSXWriter(options ...WriteOpt) (*XLSXWriter, error) {
	writeOpts := &WriteOpts{
		ErrStream:  os.Stderr,
		FileName:   "",
		Sheet:      "Sheet1",
		Cell:       "",
		ClearSheet: false,
		Header:     true,
	}
	for _, option := range options {
		option(writeOpts)
	}

	f, err := openXLSXFile(writeOpts.FileName)
	if err != nil {
		return nil, err
	}
	cellX, cellY := getCell(writeOpts.Cell)

	n, err := f.GetSheetIndex(writeOpts.Sheet)
	if err != nil {
		return nil, err
	}
	// Only attempt to clear the sheet if it exists.
	if writeOpts.ClearSheet && n >= 0 {
		// Sheet exists,clear it
		if err := clearSheet(f, writeOpts.Sheet); err != nil {
			return nil, err
		}
	}

	if n < 0 {
		// Sheet does not exist, create a new one
		if _, err := f.NewSheet(writeOpts.Sheet); err != nil {
			return nil, err
		}
	}

	return &XLSXWriter{
		fileName: writeOpts.FileName,
		f:        f,
		cellX:    cellX,
		cellY:    cellY,
		sheet:    writeOpts.Sheet,
		header:   writeOpts.Header,
	}, nil
}

// openXLSXFile function opens the XLSX file.
func openXLSXFile(fileName string) (*excelize.File, error) {
	var f *excelize.File
	var err error
	if _, err = os.Stat(fileName); err != nil && !os.IsNotExist(err) {
		return nil, err
	}
	if os.IsNotExist(err) {
		// File does not exist, create a new one
		f = excelize.NewFile()
	} else {
		// File exists, open it
		f, err = excelize.OpenFile(fileName)
		if err != nil {
			return nil, err
		}
	}
	return f, nil
}

func getCell(cellName string) (int, int) {
	if cellName == "" {
		return 0, 0
	}
	x, y, err := excelize.CellNameToCoordinates(cellName)
	if err != nil {
		return 0, 0
	}
	return x - 1, y - 1
}

// clearSheet function clears the sheet.
func clearSheet(f *excelize.File, sheet string) error {
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

// PreWrite function opens the XLSXWriter.
func (w *XLSXWriter) PreWrite(columns []string, types []string) error {
	if !w.header {
		return nil
	}
	// Write header
	for i, v := range columns {
		cell, err := excelize.CoordinatesToCellName(w.cellX+i+1, w.cellY+1)
		if err != nil {
			return err
		}
		if err := w.f.SetCellValue(w.sheet, cell, v); err != nil {
			return err
		}
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
		cell, err := excelize.CoordinatesToCellName(w.cellX+i+1, w.cellY+w.rowID)
		if err != nil {
			return err
		}
		if err := w.f.SetCellValue(w.sheet, cell, v); err != nil {
			return err
		}
	}
	return nil
}

// PostWrite function closes the XLSXWriter.
func (w *XLSXWriter) PostWrite() error {
	return w.f.SaveAs(w.fileName)
}
