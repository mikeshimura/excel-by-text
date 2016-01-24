package main

import (
	"flag"
	"fmt"
	"github.com/mikeshimura/excel-by-text/util"
	"os"
)

func main() {
	var encoding string
	flag.StringVar(&encoding, "e", "", "encoding default UTF-8, accept ShiftJIS, EUCJP")
	flag.Parse()
	if len(flag.Args()) == 0 {
		fmt.Fprintf(os.Stderr, "Input file not specified\n")
		flag.PrintDefaults()
		os.Exit(2)
	}
	util.Execute(flag.Args()[0], encoding)
}
