package xl

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"log"
	"os"
	//"xlsx"
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
	t := &Tab{
		Header: header,
		rows:   rows[1:],
	}
	fmt.Printf("table size = %d x %d\n", len(t.rows), len(header))
	return t
}
func (t *Tab) RowSize() int {
	return len(t.rows)
}
func (t *Tab) ColSize() int {
	return len(t.Header)
}
func (t *Tab) At(row, col int) string {
	if row < 0 || col < 0 || row >= t.RowSize() {
		log.Fatalf("invalid index = (%d, %d) - row=%d, col=%d\n", row, col, t.RowSize(), t.ColSize())
	}
	cells := t.rows[row].Cells
	if len(cells) <= col {
		return ""
	}
	return cells[col].Value
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
	// 转换第(row, col)单元格数据
	Field(row, col int, value string) string
	// 将一行数据转化为一个对象的字符串
	Line(row int, fields []string) string
	// 将所有对象的字符串转化为一个文档字符串
	Merge(lines []string) string
}

func (t *Tab) Encode(fn Conv) string {
	R := t.RowSize()
	N := t.ColSize()
	rows := make([]string, 0, R)
	line := make([]string, 0, N)
	for row := 0; row < R; row++ {
		for col := 0; col < N; col++ {
			line = append(line, fn.Field(row, col, t.At(row, col)))
		}
		rows = append(rows, fn.Line(row, line))
		line = line[0:0]
	}
	return fn.Merge(rows)
}

type I bool
type Comma bool 
type N bool
// identity process, return the value itself
func (i I)Field(r,c int, value string)string{
    return value
}
func StringField(value string) string{
    value = strings.TrimSpace(value)
    if len(value) == 0 {
        return ""
    }
    if value[0]!='"' {
        value = `"` + value + `"`
    }
    return value
}
// catenate fields by ','
func (c Comma)Line(row int, fields []string)string {
    if len(fields)==0 {
        return ""
    }
    s := fields[0]
    for j:=1; j<len(fields); j++ {
        s += ","+fields[j]
    }
    return s
}
// append '\n' after each line and catenate them
func (m N)Merge(lines []string) string {
    s := ""
    for _, value := range lines {
        s += value+"\n"
    }
    return s
}
