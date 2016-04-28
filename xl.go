package xl

import (
	"fmt"
	"log"
	"os"
	"github.com/tealeg/xlsx"
)

type Tab struct {
	Header []string
	rows   []*xlsx.Row
}

func rows2Names(row *xlsx.Row) []string {
	c := row.Cells
	name := make([]string, 0, len(c))
	for _, v := range c {
		name = append(name, v.Value)
	}
	return name
}

// only support .xlsx, sheet number start from 0
func Open(file string, sheet int) *Tab {
	fmt.Println("open file = " + file)
	f, e := xlsx.OpenFile(file)
	if e != nil {
		log.Fatalln("OpenFile : ", e)
	}
	if sheet < 0 || sheet >= len(f.Sheets) {
		log.Fatalln("bad sheet number: ", sheet)
	}
	rows := f.Sheets[sheet].Rows
	header := rows2Names(rows[0])
	fmt.Println("col size = ", len(header))
	t := &Tab{
		Header: header,
		rows:   rows[1:],
	}
	return t
}
func (t *Tab) RowSize() int {
	return len(t.rows)
}
func (t *Tab) ColSize() int {
	return len(t.Header)
}
func (t *Tab) At(row, col int) string {
	if row < 0 || col < 0 || row >= t.RowSize() || col >= t.ColSize() {
		log.Fatalf("invalid index = (%d, %d) - row=%d, col=%d\n", row, col, t.RowSize(), t.ColSize())
	}
	return t.rows[row].Cells[col].Value
}
func WriteString(outFileName, data string) {
	file, err := os.OpenFile(outFileName, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666)
	if err != nil {
		fmt.Println("dump file data :")
		fmt.Print(data)
		log.Fatalln("open file to write failed => ", outFileName)
	}
	defer file.Close()
	file.WriteString(data)
}

type Conv interface {
	// 转换第index列数据
	Field(index int, value string) string
	// 将一行数据转化为一个对象的字符串
	Line(fields []string) string
	// 将所有对象的字符串转化为一个文档字符串
	Merge(lines []string) string
}

func (t *Tab) Encode(fn Conv) string {
	R := t.RowSize()
	N := t.ColSize()
	rows := make([]string, 0, R)
	for i := 0; i < R; i++ {
		cells := t.rows[i].Cells
		if len(cells) > 0 {
			line := make([]string, N)
			line[0] = fn.Field(0, cells[0].Value)
			for j := 1; j < N; j++ {
				line[j] = fn.Field(j, cells[j].Value)
			}
			rows = append(rows, fn.Line(line))
		}
	}
	return fn.Merge(rows)
}

// example for encoder
type DummyEncoder struct {
}

func (s *DummyEncoder) Field(i int, value string) string {
	return value
}
func (se *DummyEncoder) Line(fields []string) string {
	s := ""
	for _, v := range fields {
		s += v + ","
	}
	return s
}
func (se *DummyEncoder) Merge(lines []string) string {
	s := ""
	for _, v := range lines {
		s += v + "\n"
	}
	return s
}
