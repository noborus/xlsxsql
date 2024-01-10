// Package xlsxsql provides a reader for XLSX files.
// It uses the trdsql and excelize/v2 packages to read XLSX files and convert them into SQL tables.
// The main type is XLSXReader, which implements the trdsql.Reader interface.
package xlsxsql

import (
	"os"
	"reflect"
	"testing"

	"github.com/noborus/trdsql"
)

func TestXLSXSheet(t *testing.T) {
	type args struct {
		fileName string
	}
	tests := []struct {
		name    string
		args    args
		want    []string
		wantErr bool
	}{
		{
			"test1",
			args{"testdata/test1.xlsx"},
			[]string{"Sheet1"},
			false,
		},
		{
			"test2",
			args{"testdata/test2.xlsx"},
			[]string{"Sheet1", "Sheet2"},
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := XLSXSheet(tt.args.fileName)
			if (err != nil) != tt.wantErr {
				t.Errorf("XLSXSheet() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("XLSXSheet() = %v, want %v", got, tt.want)
			}
		})
	}
}

type xlsxReader struct {
	fileName string
	opts     *trdsql.ReadOpts
}

func createXLSXReader(t *testing.T, xlsx xlsxReader) (trdsql.Reader, error) {
	t.Helper()
	reader, err := os.Open(xlsx.fileName)
	if err != nil {
		t.Errorf("os.Open() error = %v", err)
		return nil, err
	}
	r, err := NewXLSXReader(reader, xlsx.opts)
	if err != nil {
		t.Errorf("NewXLSXReader() error = %v", err)
		return nil, err
	}
	return r, nil
}

func TestXLSXReader_PreReadRow(t *testing.T) {
	tests := []struct {
		name string
		xlsx xlsxReader
		want [][]any
	}{
		{
			"test1",
			xlsxReader{
				fileName: "testdata/test1.xlsx",
				opts:     &trdsql.ReadOpts{InPreRead: 1},
			},
			[][]any{
				{"1", "a"},
				{"2", "b"},
				{"3", "c"},
				{"4", "d"},
				{"5", "e"},
				{"6", "f"},
			},
		},
		{
			"test2",
			xlsxReader{
				fileName: "testdata/test2.xlsx",
				opts: &trdsql.ReadOpts{
					InHeader:  true,
					InPreRead: 1,
				},
			},
			[][]any{
				{"1", "apple"},
				{"2", "orange"},
				{"3", "melon"},
			},
		},
		{
			"test3",
			xlsxReader{
				fileName: "testdata/test3.xlsx",
				opts: &trdsql.ReadOpts{
					InHeader: true,
					InJQuery: "Sheet1.C1",
				},
			},
			[][]any{
				{"1", "apple"},
				{"2", "orange"},
				{"3", "melon"},
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r, err := createXLSXReader(t, tt.xlsx)
			if err != nil {
				t.Errorf("createXLSXReader() error = %v", err)
				return
			}
			if got := r.PreReadRow(); !reflect.DeepEqual(got, tt.want) {
				t.Errorf("XLSXReader.PreReadRow() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestXLSXReader_Names(t *testing.T) {
	tests := []struct {
		name    string
		xlsx    xlsxReader
		want    []string
		wantErr bool
	}{
		{
			"test1",
			xlsxReader{
				fileName: "testdata/test1.xlsx",
				opts:     &trdsql.ReadOpts{InPreRead: 1},
			},
			[]string{"A1", "B1"},
			false,
		},
		{
			"test2",
			xlsxReader{
				fileName: "testdata/test2.xlsx",
				opts: &trdsql.ReadOpts{
					InHeader:  true,
					InPreRead: 1,
				},
			},
			[]string{"id", "name"},
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r, err := createXLSXReader(t, tt.xlsx)
			if err != nil {
				t.Errorf("createXLSXReader() error = %v", err)
				return
			}
			got, err := r.Names()
			if (err != nil) != tt.wantErr {
				t.Errorf("XLSXReader.Names() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("XLSXReader.Names() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestXLSXReader_Types(t *testing.T) {
	tests := []struct {
		name    string
		xlsx    xlsxReader
		want    []string
		wantErr bool
	}{
		{
			"test1",
			xlsxReader{
				fileName: "testdata/test1.xlsx",
				opts:     &trdsql.ReadOpts{InPreRead: 1},
			},
			[]string{"text", "text"},
			false,
		},
		{
			"test2",
			xlsxReader{
				fileName: "testdata/test2.xlsx",
				opts: &trdsql.ReadOpts{
					InHeader:  true,
					InPreRead: 1,
				},
			},
			[]string{"text", "text"},
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r, err := createXLSXReader(t, tt.xlsx)
			if err != nil {
				t.Errorf("createXLSXReader() error = %v", err)
				return
			}
			got, err := r.Types()
			if (err != nil) != tt.wantErr {
				t.Errorf("XLSXReader.Types() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("XLSXReader.Types() = %v, want %v", got, tt.want)
			}
		})
	}
}
