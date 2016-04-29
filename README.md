# xl

This library is used to convert xlsx data to other file with different format, such as csv, json, etc.
The provided interface is pretty simple, but give you full control.

How To Use It
=============
1. make use you have installed go-lang package, and set GOROOT correctly.
2. make a directory for your workspace, like 
	```
	mkdir workspace && cd workspace
	```
3. get this package and its dependency by 
 	```
	GOPATH=`pwd` go get github.com/gfphoenix/xl
	```
4. start coding your formatter.

Example
-------
For example:

```go
package main

import (
	"fmt"
	"os"
	"strings"
	"github.com/gfphoenix/xl"
)

type Csv struct{
    xl.Comma
    xl.N
}

func (c Csv) Field(row, col int, value string) string {
	value = strings.Replace(value, "“", "\"", -1)
	value = strings.Replace(value, "”", "\"", -1)
	return value
}
//func (c Csv) Line(row int, fields []string) string {
//	s := fields[0]
//	for j := 1; j < len(fields); j++ {
//		s += "," + fields[j]
//	}
//	return s
//}
//func (c Csv) Merge(lines []string) string {
//	s := ""
//	for _, v := range lines {
//		s += v + "\n"
//	}
//	return s
//}
func main() {
	var c Csv
	for _, in := range os.Args[1:] {
		out := strings.Replace(in, ".xlsx", ".csv", -1)
		tab := xl.Open(in, 0)
		fmt.Println(in, " => ", out)
		xl.WriteString(out, tab.Encode(c))
	}
}

```
