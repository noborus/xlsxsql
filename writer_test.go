package xlsxsql

import (
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestXLSXWriter_WriteRow(t *testing.T) {
	type opts struct {
		fileName string
		f        *excelize.File
		sheet    string
		cellName string
	}
	type args struct {
		row     []interface{}
		columns []string
		types   []string
	}
	tests := []struct {
		name     string
		fields   opts
		args     args
		wantErr  bool
		wantRows [][]string
	}{
		{
			name: "test1",
			fields: opts{
				fileName: "dummy.xlsx",
				f:        excelize.NewFile(),
				sheet:    "Sheet1",
				cellName: "A1",
			},
			args: args{
				row:     []interface{}{"a", "b", "c"},
				columns: []string{"A", "B", "C"},
				types:   []string{"string", "string", "string"},
			},
			wantErr: false,
			wantRows: [][]string{
				{"A", "B", "C"},
				{"a", "b", "c"},
			},
		},
		{
			name: "test2",
			fields: opts{
				fileName: "dummy.xlsx",
				f:        excelize.NewFile(),
				sheet:    "Sheet1",
				cellName: "A2",
			},
			args: args{
				row:     []interface{}{"a", "b", "c"},
				columns: []string{"A", "B", "C"},
				types:   []string{"string", "string", "string"},
			},
			wantErr: false,
			wantRows: [][]string{
				nil,
				{"A", "B", "C"},
				{"a", "b", "c"},
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			w, err := NewXLSXWriter(
				FileName(tt.fields.fileName),
				Sheet(tt.fields.sheet),
				Cell(tt.fields.cellName),
				ClearSheet(true),
			)
			if err != nil {
				t.Fatal(err)
			}
			if err := w.PreWrite(tt.args.columns, tt.args.types); err != nil {
				t.Fatal(err)
			}
			if err := w.WriteRow(tt.args.row, tt.args.columns); (err != nil) != tt.wantErr {
				t.Errorf("XLSXWriter.WriteRow() error = %v, wantErr %v", err, tt.wantErr)
			}
			rows, err := w.f.GetRows(tt.fields.sheet)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(rows, tt.wantRows) {
				t.Errorf("XLSXWriter.WriteRow() rows = %#v, wantRows %#v", rows, tt.wantRows)
			}
		})
	}
}
