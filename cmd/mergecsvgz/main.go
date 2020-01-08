package main

import (
	"bufio"
	"compress/gzip"
	"encoding/csv"
	"fmt"
	"io/ioutil"
	"os"
	"strings"

	"github.com/jinlingan/billcheck/utils"
)

const FILENAME = "all.csv"

var HEADER = []string{ // nolint: gochecknoglobals
	"记录类别",                  //A
	"数据记录",                  //B
	"客户号码",                  //C
	"订单号",                   //D
	"",                      //E
	"",                      //F
	"Payment Reference",     //G
	"",                      //H
	"Customer ID",           //I
	"Merchant Reference",    //J
	"",                      //K
	"Currency delivered",    //L
	"Amount delivered",      //M
	"",                      //N
	"",                      //O
	"",                      //P
	"",                      //Q
	"Payment country",       //R
	"",                      //S
	"Transaction data time", //T
	"",                      //U
	"",                      //V
	"",                      //W
	"",                      //X
	"",                      //Y
	"",                      //Z
	"Card number",           //AA
	"",                      //AB
	"",                      //AC
	"",                      //AD
	"Authorization code",    //AE
	"",                      //AF
	"",                      //AG
	"",                      //AH
	"",                      //AI
	"Issuer country",        //AJ
	"",                      //AK
	"MID",                   //AL
	"",                      //AM
	"",                      //AN
	"",                      //AO
	"Credit Card Company",   //AP
	"",                      //AQ
	"",                      //AR
	"Payment Amount",        //AS
	"",                      //AT
	"",                      //AU
	"",                      //AV
	"",                      //AW
	"",                      //AX
	"",                      //AY
	"",                      //AZ
}

func main() {
	var files []os.FileInfo

	var filesPath string

	for {
		filesPath = utils.GetInput("请输入要合并的文件所在目录：")

		var err error

		files, err = ioutil.ReadDir(filesPath)
		if err != nil {
			fmt.Println(err)
			continue
		}

		break
	}
	GetAllData(filesPath, files)
}

func GetAllData(dir string, files []os.FileInfo) {
	allLine := make([]string, 0, 1000)

	for _, fInfo := range files {
		if fInfo.IsDir() {
			continue
		}

		fullPath := dir + string(os.PathSeparator) + fInfo.Name()
		newLine, err := GetContentFromFile(fullPath)

		if err != nil {
			fmt.Println(err)
			continue
		}

		allLine = append(allLine, newLine...)
	}

	fmt.Printf("%d", len(allLine))
	SaveToCSV(allLine)
}
func ParseLine(line string) []string {
	cells := strings.Split(line, ";")
	for i := range cells {
		if i == 19 {
			continue
		}

		cells[i] = strings.TrimPrefix(cells[i], "\"")

		if strings.HasSuffix(cells[i], "\"") {
			cells[i] = cells[i][:len(cells[i])-1]
		}
	}

	return cells
}
func SaveToCSV(allLine []string) {
	for {
		var records [][]string

		filePath := utils.GetInput("文件保存目录：")

		fullFileName := filePath + string(os.PathSeparator) + FILENAME
		f, err := os.OpenFile(fullFileName, os.O_TRUNC|os.O_CREATE|os.O_WRONLY, 0666)

		if err != nil {
			fmt.Printf("保存文件失败：%s", err)
			continue
		}

		_, err = f.WriteString("\xEF\xBB\xBF")

		if err != nil {
			fmt.Printf("写入头部文件失败：%s", err)
			continue
		}

		w := csv.NewWriter(f)

		for ri := range allLine {
			cells := ParseLine(allLine[ri])
			if len(cells) > 0 {
				records = append(records, cells)
			}
		}

		err = w.Write(HEADER)

		if err != nil {
			fmt.Printf("保存文件失败：%s", err)
			continue
		}

		err = w.WriteAll(records)

		if err != nil {
			fmt.Printf("保存文件失败：%s", err)
			continue
		}

		return
	}
}

func GetContentFromFile(fullPath string) ([]string, error) {
	var lines []string

	file, err := os.Open(fullPath)

	if err != nil {
		return nil, fmt.Errorf("open file %s error: %s", fullPath, err)
	}

	gz, err := gzip.NewReader(file)

	if err != nil {
		return nil, fmt.Errorf("open file %s error: %s", fullPath, err)
	}

	defer file.Close()
	defer gz.Close()

	scanner := bufio.NewScanner(gz)
	for scanner.Scan() {
		l := scanner.Text()
		if strings.HasPrefix(l, "I\";") {
			continue
		}

		lines = append(lines, l)
	}

	return lines, nil
}
