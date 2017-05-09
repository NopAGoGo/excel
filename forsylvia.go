package main

import (
	"os"
	"strconv"
	_ "github.com/mattn/go-sqlite3"
	"github.com/Luxurioust/excelize"
	"github.com/sunday9th/forsylvia/models"
	"fmt"
	"github.com/go-xorm/xorm"
)

var F *xorm.Engine

func init() {
	var err error

	F, err = xorm.NewEngine("sqlite3", "./forsilvia.db")
	if err != nil {
		panic(fmt.Sprintf("Fail to connect to database: %v", err))
	}
}

type Txn struct {
	ID            int64         `xorm:"'id' pk autoincr"`
	DeptID        string        `xorm:"'dept_id'"`
	DeptName      string        `xorm:"'dept_name'"`
	StaffID       string        `xorm:"'staff_id'"`
	StaffName     string        `xorm:"'staff_name'"`
	TranDate      string        `xorm:"'tran_date'"`
	Amount        float64       `xorm:"'amount'"`
	Count         int           `xorm:"'count'"`
	TranType      string        `xorm:"'tran_type'"`
	MachineID     string        `xorm:"'machine_id'"`
	MachineName   string        `xorm:"'machine_name'"`
	Type          string        `xorm:"'type'"`
	TranDateShort string        `xorm:"'tran_date_short'"`
}

var header []string

func main() {
	if err := xlsx2db(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	os.RemoveAll("./out")
	os.MkdirAll("./out", 0777)
	if err := spiltData();err!=nil{
		fmt.Println(err)
		os.Exit(1)
	}
	if err := summaryData("E");err!=nil{
		fmt.Println(err)
		os.Exit(1)
	}
	if err := summaryData("W");err!=nil{
		fmt.Println(err)
		os.Exit(1)
	}
}

func xlsx2db() error {
	// 建表
	if err := models.F.Sync(models.Txn{}); err != nil {
		return err
	}
	// 读取 xlsx
	fileSource, err := excelize.OpenFile("./in.xlsx")
	if err != nil {
		return err
	}
	// 读取所有行
	rows := fileSource.GetRows("sheet1")
	// 筛选东西食堂午餐数据保存到数据库
	ml := make([]models.Txn, 0)
	m := models.Txn{}
	for index, row := range rows {
		if index == 0 {
			header = row
			continue
		}
		// 筛选午餐
		if row[10] != "午餐" {
			continue
		}
		// 筛选东西食堂的卡机
		if row[9][:1] != "d" && row[9][:1] != "x" {
			continue
		}
		m.DeptID = row[0]
		m.DeptName = row[1]
		m.StaffID = row[2]
		m.StaffName = row[3]
		m.TranDate = row[4]
		m.Amount, _ = strconv.ParseFloat(row[5], 64)
		m.Count, _ = strconv.Atoi(row[6])
		m.TranType = row[7]
		m.MachineID = row[8]
		m.MachineName = row[9]
		m.Type = row[10]
		m.TranDateShort = row[4][0:10]
		ml = append(ml, m)
		// 每80条执行一次保存
		if len(ml) == 80 {
			_, err := models.F.Insert(ml)
			if err != nil {
				return err
			}
			ml = make([]models.Txn, 0)
		}
	}
	// 保存最后不到80条的数据
	_, err = models.F.Insert(ml)
	if err != nil {
		return err
	}
	return nil
}

func spiltData() error {
	// 获取日期列表
	dateList := make([]models.Txn, 0)
	models.F.Asc("tran_date_short").Distinct("tran_date_short").Find(&dateList)
	// 按日循环
	for _, date := range dateList {
		// 创建当天 xlsx 文件
		fileSplit := excelize.CreateFile()
		// 获取当天活动机器列表
		machineList := make([]models.Txn, 0)
		models.F.Where("tran_date_short = ?", date.TranDateShort).Asc("machine_name").Distinct("machine_name").Find(&machineList)
		// 循环每个活动机器一个 sheet 页
		for sheetIndex, machine := range machineList {
			// 创建当天 xlsx 文件时自带 Sheet1
			if sheetIndex == 0 {
				fileSplit.SetSheetName("Sheet1", machine.MachineName)
			} else {
				fileSplit.NewSheet(sheetIndex+1, machine.MachineName)
			}
			// 写表头
			hStart := rune('A')
			for cellIndex, cell := range header {
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), string(hStart+rune(cellIndex))+"1", cell)
			}

			// 获取当天该机器的所有交易，按金额降序
			txnList := make([]models.Txn, 0)
			models.F.Where("tran_date_short = ? and machine_name = ?", date.TranDateShort, machine.MachineName).Desc("amount").Find(&txnList)

			// rowIndex 代替循环次数，在循环中需被修改
			rowIndex := 0
			// 写数据的起始行
			initRow := 4
			// 汇总
			sumNode := rowIndex + initRow - 1
			// 用于记录排序后的第一个刷卡金额，与之后的 4 比对
			var tmpAmount float64

			for _, txn := range txnList {
				// 记录排序后的第一个刷卡金额，之后出现4的金额，且与先前金额不等，汇总之前的所有金额，并 rowIndex 下移两行，合计节点
				if rowIndex == 0 {
					tmpAmount = txn.Amount
				} else if tmpAmount != txn.Amount && txn.Amount == 4 {
					// 写小计公式
					fileSplit.SetCellFormula("Sheet"+strconv.Itoa(sheetIndex+1), "F"+strconv.Itoa(sumNode), "=sum(F"+strconv.Itoa(sumNode+1)+":F"+strconv.Itoa(rowIndex+initRow-1)+")")
					fileSplit.SetCellFormula("Sheet"+strconv.Itoa(sheetIndex+1), "G"+strconv.Itoa(sumNode), "=sum(G"+strconv.Itoa(sumNode+1)+":G"+strconv.Itoa(rowIndex+initRow-1)+")")
					// 合并单元格
					fileSplit.MergeCell("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(sumNode), "E"+strconv.Itoa(sumNode))
					// 往合并单元格写字，居中
					fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(sumNode), "小计")
					fileSplit.SetCellStyle("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(sumNode), "A"+strconv.Itoa(sumNode), `{"alignment":{"horizontal":"center"}}`)
					// rowIndex 下移两行
					rowIndex += 2
					// 更新 sumNode
					sumNode = rowIndex + initRow - 1
					tmpAmount = txn.Amount
				}
				// 疯狂写表
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(rowIndex+initRow), txn.DeptID)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "B"+strconv.Itoa(rowIndex+initRow), txn.DeptName)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "C"+strconv.Itoa(rowIndex+initRow), txn.StaffID)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "D"+strconv.Itoa(rowIndex+initRow), txn.StaffName)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "E"+strconv.Itoa(rowIndex+initRow), txn.TranDate)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "F"+strconv.Itoa(rowIndex+initRow), txn.Amount)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "G"+strconv.Itoa(rowIndex+initRow), txn.Count)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "H"+strconv.Itoa(rowIndex+initRow), txn.TranType)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "I"+strconv.Itoa(rowIndex+initRow), txn.MachineID)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "J"+strconv.Itoa(rowIndex+initRow), txn.MachineName)
				fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "K"+strconv.Itoa(rowIndex+initRow), txn.Type)
				rowIndex++
			}
			// 汇总金额为 4 的所有记录
			fileSplit.SetCellFormula("Sheet"+strconv.Itoa(sheetIndex+1), "F"+strconv.Itoa(sumNode), "=sum(F"+strconv.Itoa(sumNode+1)+":F"+strconv.Itoa(rowIndex+initRow-1)+")")
			fileSplit.SetCellFormula("Sheet"+strconv.Itoa(sheetIndex+1), "G"+strconv.Itoa(sumNode), "=sum(G"+strconv.Itoa(sumNode+1)+":G"+strconv.Itoa(rowIndex+initRow-1)+")")
			fileSplit.MergeCell("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(sumNode), "E"+strconv.Itoa(sumNode))
			fileSplit.SetCellValue("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(sumNode), "小计")
			fileSplit.SetCellStyle("Sheet"+strconv.Itoa(sheetIndex+1), "A"+strconv.Itoa(sumNode), "A"+strconv.Itoa(sumNode), `{"alignment":{"horizontal":"center"}}`)
		}
		// 激活页签1
		fileSplit.SetActiveSheet(1)
		err := fileSplit.WriteTo("./out/" + date.TranDateShort + ".xlsx")
		if err != nil {
			return err
		}
	}
	return nil
}

func summaryData(news string) error {
	var news_zhCN string
	var news_sql string
	if news == "E" {
		news_zhCN = "东"
		news_sql = "d%"
	} else if news == "W" {
		news_zhCN = "西"
		news_sql = "x%"
	} else {
		return nil
	}
	// 获取日期列表
	dateList := make([]models.Txn, 0)
	models.F.Asc("tran_date_short").Distinct("tran_date_short").Find(&dateList)

	fileSummary := excelize.CreateFile()

	// 当月所有活动机器列表
	allMachineList := make([]models.Txn, 0)
	models.F.Asc("machine_id").Distinct("machine_id", "machine_name").Where("machine_name like ?", news_sql).Find(&allMachineList)


	// rowIndex
	rowIndex := 0
	// 写数据的起始行
	initRow := 4
	// 一天的占行数，所有活动机器数+表头+表头日期+合计+两行的间隔
	rowsOneDay := len(allMachineList) + 5
	// 整月总计
	var countSumTotal, amountSumTotal int
	var titleDate string

	for dateIndex, date := range dateList {
		titleDate = date.TranDateShort

		// 当天活动机器列表
		machineList := make([]models.Txn, 0)
		models.F.Where("tran_date_short = ? and machine_name like ?", date.TranDateShort,news_sql).Asc("machine_id").Distinct("machine_id", "machine_name").Find(&machineList)

		// 两天一行，偶数天换行
		xStart := rune('A')
		if dateIndex%2 == 1 {
			xStart = rune('F')
		}

		// 合并单元格，每天表头标题
		fileSummary.MergeCell("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow), string(xStart+3)+strconv.Itoa(rowIndex+initRow))
		fileSummary.SetCellValue("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow), date.TranDateShort)
		// 表头
		fileSummary.SetCellValue("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow+1), "机台编号")
		fileSummary.SetCellValue("Sheet1", string(xStart+1)+strconv.Itoa(rowIndex+initRow+1), "机台名称")
		fileSummary.SetCellValue("Sheet1", string(xStart+2)+strconv.Itoa(rowIndex+initRow+1), "份数")
		fileSummary.SetCellValue("Sheet1", string(xStart+3)+strconv.Itoa(rowIndex+initRow+1), "金额(元)")
		// 日总计
		var countSum, amountSum int
		// 写每日明细
		for mii, machineIterator := range allMachineList {
			fileSummary.SetCellValue("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow+2+mii), machineIterator.MachineID)
			fileSummary.SetCellValue("Sheet1", string(xStart+1)+strconv.Itoa(rowIndex+initRow+2+mii), machineIterator.MachineName)
			fileSummary.SetCellValue("Sheet1", string(xStart+2)+strconv.Itoa(rowIndex+initRow+2+mii), 0)
			fileSummary.SetCellValue("Sheet1", string(xStart+3)+strconv.Itoa(rowIndex+initRow+2+mii), 0)

			for _, machine := range machineList {
				if machineIterator == machine {
					txn := new(models.Txn)
					total, err := models.F.Where("tran_date_short = ? and machine_id = ? and machine_name = ? and amount = ?",
						date.TranDateShort, machineIterator.MachineID, machineIterator.MachineName, 4).Sums(txn, "count", "amount")
					if err != nil {
						return err
					}
					fileSummary.SetCellValue("Sheet1", string(xStart+2)+strconv.Itoa(rowIndex+initRow+2+mii), int(total[0]))
					fileSummary.SetCellValue("Sheet1", string(xStart+3)+strconv.Itoa(rowIndex+initRow+2+mii), int(total[1]))
					countSum += int(total[0])
					amountSum += int(total[1])
					break
				}
			}
		}

		// 调整每日样式
		fileSummary.MergeCell("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3), string(xStart+1)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3))
		fileSummary.SetCellValue("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3), "合计")
		fileSummary.SetCellValue("Sheet1", string(xStart+2)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3), countSum)
		fileSummary.SetCellValue("Sheet1", string(xStart+3)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3), amountSum)
		fileSummary.SetCellStyle("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow), string(xStart+3)+strconv.Itoa(rowIndex+initRow+1),
			`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"alignment":{"horizontal":"center"}}`)
		fileSummary.SetCellStyle("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow+2), string(xStart+3)+strconv.Itoa(rowIndex+initRow+rowsOneDay-4),
			`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}]}`)
		fileSummary.SetCellStyle("Sheet1", string(xStart)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3), string(xStart+1)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3),
			`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"alignment":{"horizontal":"center"}}`)
		fileSummary.SetCellStyle("Sheet1", string(xStart+2)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3), string(xStart+3)+strconv.Itoa(rowIndex+initRow+rowsOneDay-3),
			`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}]}`)

		// 两天一行，偶数天换行
		if dateIndex%2 == 1 {
			rowIndex += rowsOneDay
		}
		// 每日合计
		countSumTotal += countSum
		amountSumTotal += amountSum
	}

	// 全部合计
	fileSummary.MergeCell("Sheet1", "A1", "B1")
	fileSummary.MergeCell("Sheet1", "A2", "B2")
	fileSummary.SetCellValue("Sheet1", "A1", titleDate[:7])
	fileSummary.SetCellValue("Sheet1", "A2", "全部合计")
	fileSummary.SetCellValue("Sheet1", "C1", "份数")
	fileSummary.SetCellValue("Sheet1", "D1", "金额(元)")
	fileSummary.SetCellValue("Sheet1", "C2", countSumTotal)
	fileSummary.SetCellValue("Sheet1", "D2", amountSumTotal)
	fileSummary.SetCellStyle("Sheet1", "A1", "D1",
		`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"alignment":{"horizontal":"center"}}`)
	fileSummary.SetCellStyle("Sheet1", "A2", "B2",
		`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"alignment":{"horizontal":"center"}}`)
	fileSummary.SetCellStyle("Sheet1", "C2", "D2",
		`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}]}`)
	fileSummary.SetActiveSheet(1)
	err := fileSummary.WriteTo("./out/" + titleDate[:7] + news_zhCN + ".xlsx")
	if err != nil {
		return err
	}

	return nil
}
