package util

import (
	"fmt"
	ge "github.com/mikeshimura/goexcel"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
	"io"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
	"time"
)

func transformEncoding(rawReader io.Reader, trans transform.Transformer) (string, error) {
	ret, err := ioutil.ReadAll(transform.NewReader(rawReader, trans))
	if err == nil {
		return string(ret), nil
	} else {
		return "", err
	}
}
func FromShiftJIS(str string) (string, error) {
	return transformEncoding(strings.NewReader(str), japanese.ShiftJIS.NewDecoder())
}
func FromEUCJP(str string) (string, error) {
	return transformEncoding(strings.NewReader(str), japanese.EUCJP.NewDecoder())
}
func Execute(infile string, encoding string) {
	buf, err := ioutil.ReadFile(infile)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Input file:"+infile+" not found")
		os.Exit(2)
	}
	bufs := string(buf)
	switch encoding {
	case "":

	case "ShiftJIS":
		bufs, err = FromShiftJIS(bufs)
		if err != nil {
			fmt.Fprintf(os.Stderr, "ShiftJIS Convert Error")
			os.Exit(2)
		}
	case "EUCJP":
		bufs, err = FromEUCJP(bufs)
		if err != nil {
			fmt.Fprintf(os.Stderr, "EUCJP Convert Error")
			os.Exit(2)
		}
	default:
		fmt.Fprintf(os.Stderr, "encoding:"+encoding+
			" is illegal. ShiftJIS or EUCJP.")
		os.Exit(2)
	}

	//fmt.Printf("bufs %v\n",bufs)
	pos := strings.Index(bufs, "\r\n")
	//fmt.Printf("pos %v\n",pos)
	if pos > -1 {
		bufs = strings.Replace(bufs, "\r\n", "\n", -1)
	}

	excel := ge.CreateGoexcel()
	lines := strings.Split(bufs, "\n")
	for _, line := range lines {
		ExecuteSub(line, excel)
	}
	fmt.Println("Excel Generate Finished")
}

func ExecuteSub(line string, excel *ge.Goexcel) {
	cols := strings.Split(line, "\t")
	switch cols[0] {
	case "O":
		OpenFile(cols, line, excel)
	case "W":
		SaveFile(cols, line, excel)
	case "STA":
		AddSheet(cols, line, excel)
	case "STS":
		SetSheet(cols, line, excel)
	case "SN":
		CreateStyle(cols, line, excel)
	case "CS":
		CopyStyle(cols, line, excel)
	case "SFN":
		SetFontName(cols, line, excel)
	case "SFS":
		SetFontSize(cols, line, excel)
	case "SC":
		SetFontColor(cols, line, excel)
	case "SI":
		SetItalic(cols, line, excel)
	case "SBL":
		SetBold(cols, line, excel)
	case "SU":
		SetUnderline(cols, line, excel)
	case "SB":
		SetBorder(cols, line, excel)
	case "SBC":
		SetBorderColor(cols, line, excel)
	case "SF":
		SetFill(cols, line, excel)
	case "SH":
		SetHorizontalAlign(cols, line, excel)
	case "SV":
		SetVerticalAlign(cols, line, excel)
	case "CW":
		SetColWidth(cols, line, excel)
	case "SS":
		SetStyle(cols, line, excel)
	case "M":
		Merge(cols, line, excel)
	case "FS":
		SetFormat(cols, line, excel)
	case "S":
		SetString(cols, line, excel)
	case "N":
		SetNumber(cols, line, excel)
	case "NF":
		SetNumberFormat(cols, line, excel)
	case "D":
		SetDate(cols, line, excel)
	case "DF":
		SetDateFormat(cols, line, excel)
	case "DT":
		SetDateTime(cols, line, excel)
	case "DTF":
		SetDateTimeFormat(cols, line, excel)
	case "F":
		SetFormula(cols, line, excel)
	case "FF":
		SetFormulaFormat(cols, line, excel)

	default:
		fmt.Printf("Skip %v\n", line)
	}
}
func SetNumberFormat(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 5, 5, line)
	excel.SetFloatFormat(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseFloat(cols[3], line), cols[4])
}
func SetNumber(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetFloat(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseFloat(cols[3], line))
}
func SetDateFormat(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 5, 5, line)
	excel.SetDateFormat(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseDate(cols[3], line), cols[4])
}
func SetDate(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetDate(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseDate(cols[3], line))
}
func ParseDateTime(s string, line string) time.Time {
	res, err := time.Parse("2006/01/02 15:04:05", s)
	if err != nil {
		fmt.Fprintf(os.Stderr, s+" is not datetime. yyyy/mm/dd hh:mm:dd:"+line)
		os.Exit(2)
	}
	return res
}
func ParseDate(s string, line string) time.Time {
	res, err := time.Parse("2006/01/02", s)
	if err != nil {
		fmt.Fprintf(os.Stderr, s+" is not date. yyyy/mm/dd :"+line)
		os.Exit(2)
	}
	return res
}
func SetDateTimeFormat(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 5, 5, line)
	excel.SetDateFormat(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseDateTime(cols[3], line), cols[4])
}
func SetDateTime(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetDateTime(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseDateTime(cols[3], line))
}
func SetFormat(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetFormat(Atoi(cols[1], line), Atoi(cols[2], line), cols[3])
}
func SetFormulaFormat(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 5, 5, line)
	excel.SetFormulaFormat(Atoi(cols[1], line),
		Atoi(cols[2], line), cols[3], cols[4])
}
func SetFormula(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetFormula(Atoi(cols[1], line), Atoi(cols[2], line), cols[3])
}

func Merge(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 5, 5, line)
	excel.Merge(Atoi(cols[1], line), Atoi(cols[2], line),
		Atoi(cols[3], line), Atoi(cols[4], line))
}

func SetColWidth(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetColWidth(Atoi(cols[1], line), Atoi(cols[2], line),
		ParseFloat(cols[3], line))
}

func SetVerticalAlign(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetVerticalAlign(cols[1], CheckVAline(cols[2], line))
}
func CheckVAline(ptn string, line string) string {
	res := ge.VAlingnMap[strings.ToUpper(ptn[0:1])+ptn[1:]]
	if res == "" {
		fmt.Fprintf(os.Stderr, ptn+" is not Vertical Aline:"+line)
		os.Exit(2)
	}
	return res
}
func SetHorizontalAlign(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetHorizontalAlign(cols[1], CheckHAline(cols[2], line))
}
func CheckHAline(ptn string, line string) string {
	res := ge.HAlingnMap[strings.ToUpper(ptn[0:1])+ptn[1:]]
	if res == "" {
		fmt.Fprintf(os.Stderr, ptn+" is not Horizontal Aline:"+line)
		os.Exit(2)
	}
	return res
}
func SetFill(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 5, 5, line)
	defer PanicRecover(line)
	excel.SetFill(cols[1], CheckFillPattern(cols[2], line),
		SetColorSub(cols[3], line), SetColorSub(cols[4], line))
}
func CheckFillPattern(ptn string, line string) string {
	res := ge.PatternMap[strings.ToUpper(ptn[0:1])+ptn[1:]]
	if res == "" {
		fmt.Fprintf(os.Stderr, ptn+" is not Fill Pattern:"+line)
		os.Exit(2)
	}
	return res
}

func SetBorder(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	ptn := CheckBorderPattern(cols[3], line)
	excel.SetBorder(cols[1], cols[2], ptn)
}
func CheckBorderPattern(ptn string, line string) string {
	res := ge.BorderMap[strings.ToUpper(ptn[0:1])+ptn[1:]]
	if res == "" && ptn != "None" {
		fmt.Fprintf(os.Stderr, ptn+" is not Border Pattern:"+line)
		os.Exit(2)
	}
	return res
}
func SetUnderline(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetUnderline(cols[1], AtoBool(cols[2], line))
}
func SetBold(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetBold(cols[1], AtoBool(cols[2], line))
}
func SetItalic(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetItalic(cols[1], AtoBool(cols[2], line))
}
func SetBorderColor(cols []string, line string, excel *ge.Goexcel) {
	defer PanicRecover(line)
	CheckColno(cols, 4, 4, line)
	excel.SetBorderColor(cols[1], cols[2], SetColorSub(cols[3], line))
}
func SetFontColor(cols []string, line string, excel *ge.Goexcel) {
	defer PanicRecover(line)
	CheckColno(cols, 3, 3, line)
	excel.SetFontColor(cols[1], SetColorSub(cols[2], line))
}
func SetColorSub(color string, line string) string {
	pos := strings.Index(color, ":")
	if pos == -1 {
		return color
	}
	cs := strings.Split(color, ":")
	res := ge.ColorDencity(cs[0], Atoi(cs[1], line))
	return res
}
func PanicRecover(line string) {
	errx := recover()
	//fmt.Printf("errx %v \n",errx)
	if errx != nil {
		fmt.Fprintf(os.Stderr, "%v:"+line, errx)
		os.Exit(2)
	}
}
func SetFontSize(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetFontSize(cols[1], Atoi(cols[2], line))
}

func SetFontName(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.SetFontName(cols[1], cols[2])
}

func SetSheet(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 2, 2, line)
	err := excel.SetSheet(cols[1])
	if err != nil {
		fmt.Fprintf(os.Stderr, cols[1]+" can't set:"+line)
		os.Exit(2)
	}
}
func AddSheet(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 2, 2, line)
	err := excel.AddSheet(cols[1])
	if err != nil {
		fmt.Fprintf(os.Stderr, cols[1]+" can't add :"+line)
		os.Exit(2)
	}
}
func SetString(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetString(Atoi(cols[1], line), Atoi(cols[2], line), cols[3])
	//fmt.Printf("SetString %v\n", cols[3])
}

func SetStyle(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 4, line)
	excel.SetStyleByKey(Atoi(cols[1], line), Atoi(cols[2], line), cols[3])
}
func CopyStyle(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 3, 3, line)
	excel.CopyStyle(cols[1], cols[2])
}
func CreateStyle(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 4, 6, line)
	cols = AddBlank(cols, 6)
	ptn := ""
	if cols[5] != "" {
		ptn = CheckBorderPattern(cols[5], line)
	}
	excel.CreateStyleByKey(cols[1], cols[2], Atoi(cols[3], line), cols[4], ptn)
}
func AtoBool(s string, line string) bool {
	if s == "T" {
		return true
	}
	if s == "F" {
		return false
	}

	fmt.Fprintf(os.Stderr, s+" is not T or F for bool :"+line)
	os.Exit(2)
	return false
}

func Atoi(s string, line string) int {
	res, err := strconv.Atoi(s)
	if err != nil {
		fmt.Fprintf(os.Stderr, s+" is not integer :"+line)
		os.Exit(2)
	}
	return res
}
func ParseFloat(s string, line string) float64 {
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		fmt.Fprintf(os.Stderr, s+" is not numeric :"+line)
		os.Exit(2)
	}
	return f
}
func AddBlank(cols []string, max int) []string {
	for i := len(cols); i < max; i++ {
		cols = append(cols, "")
	}
	return cols
}
func OpenFile(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 2, 2, line)
	err := excel.OpenFile(cols[1])
	if err != nil {
		fmt.Fprintf(os.Stderr, "file not found :"+line)
		os.Exit(2)
	}
}

func SaveFile(cols []string, line string, excel *ge.Goexcel) {
	CheckColno(cols, 2, 2, line)
	err := excel.Save(cols[1])

	if err != nil {
		fmt.Fprintf(os.Stderr, "file write error :"+line)
		os.Exit(2)
	}
}

func CheckColno(cols []string, min int, max int, line string) {
	no := len(cols)
	if no < min || no > max {
		fmt.Fprintf(os.Stderr, "colno must be between "+
			strconv.Itoa(min)+" "+strconv.Itoa(max)+" :"+line)
		os.Exit(2)
	}
}
