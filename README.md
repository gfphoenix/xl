# xl

This library is used to convert xlsx data to other file with different format, such as csv, json, etc.
The provided interface is pretty simple, but give you full control.

# Table
The table sheet has an exported field, Header, and 3 functions to access data:
RowSize() , ColSize(), At(row, col) .

There are two format/encode ways: one is to access each field with two index, the other way is to implement
an interface `Conv'. 
For example:

package main

import (
	"fmt"
	"os"
	"strings"
	"xl"
)

type Csv struct{}

func (c Csv) Field(i int, value string) string {
	value = strings.Replace(value, "“", "\"", -1)
	value = strings.Replace(value, "”", "\"", -1)
	return value
}
func (c Csv) Line(fields []string) string {
	s := fields[0]
	for j := 1; j < len(fields); j++ {
		s += "," + fields[j]
	}
	return s
}
func (c Csv) Merge(lines []string) string {
	s := ""
	for _, v := range lines {
		s += v + "\n"
	}
	return s
}
func main() {
	var c Csv
	for _, in := range os.Args[1:] {
		out := strings.Replace(in, ".xlsx", ".csv", -1)
		tab := xl.Open(in, 0)
		fmt.Println(in, " => ", out)
		xl.WriteString(out, tab.Encode(c))
	}
}
