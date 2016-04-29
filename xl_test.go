package xl

import (
	"fmt"
	"testing"
)
type DummyEncoder struct {
    I
    Comma
    N
}
func TestXl(t *testing.T) {
	fmt.Println("hello")
	tab := Open("weapon.xlsx", 0)
	var s DummyEncoder
	str := tab.Encode(&s)
	WriteString("/tmp/out.txt", str)
}
