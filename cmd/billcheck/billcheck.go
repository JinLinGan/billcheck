package main

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"os"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/jinlingan/billcheck/utils"

	"github.com/360EntSecGroup-Skylar/excelize"
	log "github.com/sirupsen/logrus"
)

var SkipPlatform = []string{
	"aliexpress",
	"批发商订单专用平台",
	"Amazon - 1",
	"Amazon - 2",
	"营销样品",
}

var PlatformOrderPrefix = map[string]string{
	"Getnamenecklace":                   "",
	"portraitnecklace.com":              "portrait-",
	"Obtenircollierprenom.fr":           "obten-",
	"039-家庭生辰石系列 - craftfamilytree.com": "craftfamilytree-",
	"034-家庭系列 - thefamilynecklace.com":  "thefamilynecklace-",
	"068-路标相框-signgifts.com":            "signgifts-",
	"073-照片logo球系列-myballgift.com":      "myballgift-",
	"046-城市概念 - mycityoutline.com":      "mycityoutline-",
	"077-硬币系列-mynamecoins.com":          "mynamecoins-",
	"Roseinside.com":                    "roseinside-",
	"074-定制人偶系列-myfacefigure.com":       "myfacefigure-",
	"028-无穷大品类 - namedinfinity.com":     "namedinfinity-",
	"Bekommenamenskette.com":            "bekomme-",
	"012-彩宝款品类 - gemadam.com":           "gemadam-",
	"029-无穷大品类 - myinfinitys.com":       "myinfinitys-",
	"016-月亮款品类 - hexmoon.com":           "hexmoon-",
	"052-骨灰盒系列-ashesnecklace.com":       "ashesnecklace-",
	"荷兰站 - krijgnaamketting":            "krijgnaamketting-",
	"080-自然花系列-floralnecklace.com":      "floralnecklace-",
	"011-家庭款品类 - familydesign.net":      "familydesign-",
	"Obtenercollarconnombre 西语":         "obtener-",
	"Cheapnamenecklace.com":             "cheapnamenecklace-",
	"意大利站 - nomecollana.com":            "nomecollana-",
	"035-家庭系列 - familyengraved.com":     "familyengraved-",
	"061-木质纪念品系列-thesephoto.com":        "thesephoto-",
	"Beaustar.com":                      "beaustar-",
	"022-照片品类 - photosfeel.com":         "photosfeel-",
	"Custom-necklace.com":               "cn-",
	"003-宠物系列法语站 - monchanceux.com":     "monchanceux-",
	"002-宠物系列英语站 - mypetbuzz.com":       "mypetbuzz-",
	"063-仿真头骨系列-runskull.com":           "runskull-",
	"013-彩宝款品类 - bestbirthstone.net":    "bestbirthstone-",
	"062-木质纪念品系列-myphotoideas.com":      "myphotoideas-",
	"076-宠物系列-petbey.com":               "petbey-",
	"021-MO品类 - monogramsign.com":       "monogramsign-",
	"024-名字品类 - mynamehut.com":          "mynamehut-",
	"045-嘻哈系列 - icedoutdesign.com":      "icedoutdesign-",
	"004-宠物系列德语站 - haustierkette.com":   "haustierkette-",
	"Sheown.com":                    "sheown-",
	"Obtercolarcomnome 葡语":          "obtercol-",
	"031-钱包系列 - walletree.com":      "walletree-",
	"036-Bar系列 - mybarnecklace.com": "mybarnecklace-",
	"010-家庭款品类 - belemom.com":       "belemom-",
	"067-路标相框-myheartgift.com":      "myheartgift-",
}

type BillInfo struct {
	// 免费重做
	IsFreeRedo bool
	// 是否跳过
	IsSkipByPlatform bool
	// 是否是订单
	IsBill bool
	// 平台名称
	PlatformName string
	// 原始订单编号
	OriBillNum string
	// 总价
	TotalPrice float64
	// 生成的订单ID
	BillID string
	// 旧数据
	OldData []string
	// 核对状态
	CheckStatus string
	// 核对金额
	CheckPrice float64
}

type PayInfo struct {
	BillID     string
	TotalPrice float64
}

func main() {
	bills := readBillInfo()
	paypalInfo := readPayPalInfoloop()
	worldpayInof := readWorldPayInfo()
	all, notOKBills := checkBill(bills, paypalInfo, worldpayInof)
	fmt.Println("保存异常订单")
	saveResult(notOKBills, "NotOKBills.csv")
	fmt.Println("保存所有订单")
	saveResult(all, "ALLBills.csv")

}

func saveResult(bills []BillInfo, fileName string) {
	records := make([][]string, 0, len(bills))
	head := []string{
		"订单ID",
		"订单状态",
		"应收金额",
		"实收金额",
	}
	head = append(head, bills[0].OldData...)
	records = append(records, head)
	for rowIndex := 1; rowIndex < len(bills); rowIndex++ {
		new := []string{"", "", "", ""}
		if bills[rowIndex].IsBill {

			new = []string{
				bills[rowIndex].BillID,
				bills[rowIndex].CheckStatus,
				fmt.Sprintf("%.2f", bills[rowIndex].TotalPrice),
				fmt.Sprintf("%.2f", bills[rowIndex].CheckPrice),
			}
		}

		all := append(new, bills[rowIndex].OldData...)
		records = append(records, all)
	}
	for {
		filePath := utils.GetInput("文件保存目录：")
		fullFileName := filePath + string(os.PathSeparator) + fileName
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
		for ri := range records {
			if records[ri][4] != "" && ri != 0 {
				records[ri][4] = "\"" + records[ri][4] + "\""
			}
		}
		err = w.WriteAll(records)
		if err != nil {
			fmt.Printf("保存文件失败：%s", err)
			continue
		}
		return
	}
}

func checkBill(bills []BillInfo, pInfo []PayInfo, wInfo []PayInfo) (all, notok []BillInfo) {
	payInfos := getAllPayInfo(pInfo, wInfo)
	billCount := 0
	okBillCount := 0
	priceUnequalCount := 0
	notFoundCount := 0
	skipByPlatformCount := 0
	freeRedoCount := 0

	billPlatformCount := map[string]int{}
	okBillPlatformCount := map[string]int{}
	priceUnequalPlatformCount := map[string]int{}
	notFoundPlatformCount := map[string]int{}
	skipByPlatformPlatformCount := map[string]int{}
	freeRedoPlatformCount := map[string]int{}

	notOKBills := make([]BillInfo, 0, 500000)
	notOKBills = append(notOKBills, bills[0])
	for i := range bills {
		if !bills[i].IsBill {
			continue
		}
		billPlatformCount[bills[i].PlatformName] = billPlatformCount[bills[i].PlatformName] + 1
		billCount++
		if bills[i].IsFreeRedo {
			freeRedoPlatformCount[bills[i].PlatformName] = freeRedoPlatformCount[bills[i].PlatformName] + 1
			freeRedoCount++
			continue
		}
		if bills[i].IsSkipByPlatform {
			bills[i].CheckStatus = "跳过此平台"
			bills[i].CheckPrice = 0
			notOKBills = append(notOKBills, bills[i])
			skipByPlatformPlatformCount[bills[i].PlatformName] = skipByPlatformPlatformCount[bills[i].PlatformName] + 1
			skipByPlatformCount++
			continue
		}
		if v, ok := payInfos[bills[i].BillID]; ok {
			if v.TotalPrice == bills[i].TotalPrice {
				bills[i].CheckStatus = "正常"
				bills[i].CheckPrice = v.TotalPrice
				okBillPlatformCount[bills[i].PlatformName] = okBillPlatformCount[bills[i].PlatformName] + 1
				okBillCount++
			} else {
				//log.Warn(fmt.Sprintf("订单号 '%s' 金额不符合，期望金额 %f ，实际金额 %f", bills[i].BillID, bills[i].TotalPrice, v.TotalPrice))
				bills[i].CheckStatus = "金额不符"
				bills[i].CheckPrice = v.TotalPrice
				notOKBills = append(notOKBills, bills[i])
				priceUnequalPlatformCount[bills[i].PlatformName] = priceUnequalPlatformCount[bills[i].PlatformName] + 1
				priceUnequalCount++
			}
		} else {
			//log.Warn(fmt.Sprintf("订单号 '%s' 未找到收款", bills[i].BillID))
			bills[i].CheckStatus = "未找到收款"
			bills[i].CheckPrice = 0
			notOKBills = append(notOKBills, bills[i])
			notFoundPlatformCount[bills[i].PlatformName] = notFoundPlatformCount[bills[i].PlatformName] + 1
			notFoundCount++

		}
	}
	fmt.Println("=======================================")
	fmt.Printf("一共加载了 %d 条有效订单，%d 条有效支付数据\n", billCount, len(payInfos))

	//billCount := 0
	//okBillCount := 0
	//priceUnequalCount := 0
	//notFoundCount := 0
	//skipByPlatformCount := 0
	//freeRedoCount := 0
	fmt.Printf("新做订单 %d 条，占比 %.2f%% \n", billCount-freeRedoCount, float64(billCount-freeRedoCount)/float64(billCount)*100)
	PrintPlatformInfo(billPlatformCount, billCount-freeRedoCount, billPlatformCount)
	fmt.Printf("重做订单 %d 条，占比 %.2f%% \n", freeRedoCount, float64(freeRedoCount)/float64(billCount)*100)
	PrintPlatformInfo(freeRedoPlatformCount, freeRedoCount, billPlatformCount)
	fmt.Printf("有效订单中因平台因素跳过的有 %d 条，占比 %.2f%% \n", skipByPlatformCount, float64(skipByPlatformCount)/float64(billCount-freeRedoCount)*100)
	PrintPlatformInfo(skipByPlatformPlatformCount, skipByPlatformCount, billPlatformCount)
	fmt.Printf("有效订单中正常收款的有 %d 条，占比 %.2f%% \n", okBillCount, float64(okBillCount)/float64(billCount-freeRedoCount)*100)
	PrintPlatformInfo(okBillPlatformCount, okBillCount, billPlatformCount)

	fmt.Printf("有效订单中金额不匹配的有 %d 条，占比 %.2f%% \n", priceUnequalCount, float64(priceUnequalCount)/float64(billCount-freeRedoCount)*100)
	PrintPlatformInfo(priceUnequalPlatformCount, priceUnequalCount, billPlatformCount)

	fmt.Printf("有效订单中未找到收款记录的有 %d 条，占比 %.2f%% \n", notFoundCount, float64(notFoundCount)/float64(billCount-freeRedoCount)*100)
	PrintPlatformInfo(notFoundPlatformCount, notFoundCount, billPlatformCount)

	return bills, notOKBills
}
func PrintPlatformInfo(info map[string]int, total int, platformsCount map[string]int) {
	type kv struct {
		Key   string
		Value int
	}

	var ss []kv
	for k, v := range info {
		ss = append(ss, kv{k, v})
	}

	sort.Slice(ss, func(i, j int) bool {
		return ss[i].Value > ss[j].Value
	})

	for _, kv := range ss {
		fmt.Printf("\t'%s' 站点 %d 条，占本站点订单 %.2f%%，占比 %.2f%% \n",
			kv.Key,
			kv.Value,
			float64(kv.Value)/float64(platformsCount[kv.Key])*100,
			float64(kv.Value)/float64(total)*100,
		)
	}

}
func getAllPayInfo(payInfosList ...[]PayInfo) map[string]PayInfo {
	pays := map[string]PayInfo{}
	for _, payInfos := range payInfosList {
		for _, pay := range payInfos {
			if _, ok := pays[pay.BillID]; ok {
				//log.Warn(fmt.Sprintf("发现重复订单号 '%s' ", pay.BillID))
				continue
			}
			pays[pay.BillID] = pay
		}
	}
	return pays
}

func readWorldPayInfo() []PayInfo {
	payInfos := make([]PayInfo, 0, 500000)
	for {
		fileName := utils.GetInput("请输入 WorldPay 收款平台数据文件（输入回车停止加载）：")
		if fileName == "" {
			return payInfos
		}
		f, err := excelize.OpenFile(fileName)
		if err != nil {
			log.Warnf("读取 Excel 文件异常： %s", err)
			continue
		}
		oriData := readExcelSheet(f)
		newPayInfo := parseWorldPayData(oriData)
		payInfos = append(payInfos, newPayInfo...)

		fmt.Printf("目前 WorldPay 付款信息一共有 %d 条\n", len(payInfos))
	}
}

func parseWorldPayData(oriData [][]string) []PayInfo {
	fmt.Printf("文件中共有 %d 条信息 \n", len(oriData)-8)
	p := make([]PayInfo, 0, 500000)
	for _, row := range oriData {
		if row[3] == "CAPTURED" || row[3] == "SETTLED" {

			billPrice, err := strconv.ParseFloat(strings.ReplaceAll(row[6], ",", ""), 32)
			if err != nil {
				log.Warn(fmt.Sprintf("金额格式错误 %s", row[5]))
				continue
			}
			p = append(p, PayInfo{
				BillID:     row[0],
				TotalPrice: billPrice,
			})
		}
	}
	fmt.Printf("文件中属于 'CAPTURED' 状态的信息有 %d 条\n", len(p))
	return p
}
func readPayPalInfoloop() []PayInfo {
	payInfos := make([]PayInfo, 0, 500000)

	for {
		fileName := utils.GetInput("请输入 PayPal 收款平台数据文件（输入回车停止加载）：")
		if fileName == "" {
			return payInfos
		}
		f, err := os.Open(fileName)
		if err != nil {
			log.Warnf("读取 CSV 文件异常： %s", err)
			continue
		}

		oriData, err := readCSVFile(f)
		f.Close()
		if err != nil {
			log.Warnf("读取 CSV 文件异常： %s", err)
			continue
		}
		newPayInfo := parsePayPalData(oriData)
		payInfos = append(payInfos, newPayInfo...)

		fmt.Printf("目前 PayPal 付款信息一共有 %d 条\n", len(payInfos))
	}

}
func readCSVFile(f *os.File) ([][]string, error) {
	r := csv.NewReader(f)
	return r.ReadAll()
}

func parsePayPalData(oriData [][]string) []PayInfo {
	fmt.Printf("文件中共有 %d 条信息 \n", len(oriData)-1)
	p := make([]PayInfo, 0, 500000)
	for _, row := range oriData {
		if row[3] == "快速結帳付款" || row[3] == "快速结账付款" {

			billPrice, err := strconv.ParseFloat(strings.ReplaceAll(row[5], ",", ""), 32)
			if err != nil {
				log.Warn(fmt.Sprintf("金额格式错误 %s", row[5]))
				continue
			}
			p = append(p, PayInfo{
				BillID:     row[16],
				TotalPrice: billPrice,
			})
		}
	}
	fmt.Printf("文件中属于 '快速结账付款' 状态的信息有 %d 条\n", len(p))
	return p
}

// /Users/jinlin/Downloads/11月平台数据导出（收入核对）1111.xlsx
func readBillInfo() []BillInfo {
	for {
		fileName := utils.GetInput("请输入部门内部平台订单数据文件：")
		f, err := excelize.OpenFile(fileName)
		if err != nil {
			log.Warnf("读取 Excel 文件异常： %s", err)
			continue
		}
		oriData := readExcelSheet(f)
		return parseCompanyData(oriData)
	}
}

func readExcelSheet(file *excelize.File) [][]string {
	for {
		var buffer bytes.Buffer
		buffer.WriteString("找到以下 Sheet 页：\n")
		sheets := file.GetSheetMap()

		ids := make([]int, 0, len(sheets))
		for k := range sheets {
			ids = append(ids, k)
		}
		sort.Ints(ids)
		for _, i := range ids {
			buffer.WriteString(fmt.Sprintf("%d - %s \n", i, sheets[i]))
		}
		buffer.WriteString("请选择一个进行加载（请输入左侧编号）：")
		fmt.Print(buffer.String())
		sNum := utils.GetInput("")
		num, err := strconv.Atoi(sNum)
		if err != nil {
			fmt.Println("解析输入异常，输入的内容好像不是编号")
			continue
		}

		if sheet, ok := sheets[num]; ok {
			oriData, err := readSheet(file, sheet)

			if err != nil {
				fmt.Println("读取 Sheet 页数据异常")
				continue
			}

			return oriData

		} else {
			fmt.Println("读取 Sheet 页数据异常，好像没有你输入的这个 Sheet 页")
			continue
		}
	}
}

func readSheet(file *excelize.File, sheet string) ([][]string, error) {
	begin := time.Now()
	rows := file.GetRows(sheet)
	if len(rows) < 2 {
		return nil, fmt.Errorf("%s Sheet 页中好像没什么数据", sheet)
	}
	fmt.Printf("读取数据 %d 条，耗时 %f 秒\n", len(rows), time.Since(begin).Seconds())
	return rows, nil
}

func parseCompanyData(data [][]string) []BillInfo {
	ps := map[string]struct{}{}
	bs := make([]BillInfo, 0, len(data))
	bs = append(bs, BillInfo{
		IsBill:  false,
		OldData: data[0],
	})
	billCount := 0
	for rowIndex := 1; rowIndex < len(data); rowIndex++ {
		if data[rowIndex][0] != "" {
			ps[data[rowIndex][2]] = struct{}{}
			billPrice, err := strconv.ParseFloat(data[rowIndex][7], 32)
			if err != nil {
				log.Warn(fmt.Sprintf("金额格式错误 %s", data[rowIndex][7]))
			}
			prefix := "UNKNOWN_PLATFORM-"
			skipByPlatForm := false
			freeRedo := false
			for _, v := range SkipPlatform {
				if v == data[rowIndex][2] {
					skipByPlatForm = true
					break
				}
			}

			if v, ok := PlatformOrderPrefix[data[rowIndex][2]]; ok {
				prefix = v
			} else {
				skipByPlatForm = true
			}
			if strings.HasSuffix(data[rowIndex][3], "免费重做") {
				freeRedo = true
			}

			bs = append(bs, BillInfo{
				IsBill:           true,
				IsSkipByPlatform: skipByPlatForm,
				IsFreeRedo:       freeRedo,
				PlatformName:     data[rowIndex][2],
				OriBillNum:       data[rowIndex][3],
				TotalPrice:       billPrice,
				OldData:          data[rowIndex],
				BillID:           prefix + data[rowIndex][3],
			})
			billCount++
		} else {
			bs = append(bs, BillInfo{
				IsBill:  false,
				OldData: data[rowIndex],
			})
		}
	}

	fmt.Printf("找到 %d 条有效订单\n", billCount)
	var unknown []string

	for k := range ps {
		skip := false
		for _, v := range SkipPlatform {
			if v == k {
				skip = true
				break
			}
		}

		if _, ok := PlatformOrderPrefix[k]; !ok && !skip {
			unknown = append(unknown, k)
		}
	}

	fmt.Printf("找到 %d 个售卖平台，其中有 %d 个未知平台:\n", len(ps), len(unknown))

	for k := range unknown {
		fmt.Println(unknown[k])
	}

	return bs
}
