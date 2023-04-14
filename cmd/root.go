package cmd

import (
	"database/sql"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/jmoiron/sqlx"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"

	"github.com/vvfock3r/mysqlexport/kernel/load"
	"github.com/vvfock3r/mysqlexport/kernel/module/logger"
	"github.com/vvfock3r/mysqlexport/kernel/module/mysql"
)

var (
	my    = &MySQL{}
	excel = &Excel{}
)

type MySQL struct {
	// flags
	execute   string // 执行的SQL命令
	batchSize int    // 数据库每遍历N次
	delayTime string // 延迟多久

	// 存储SQL查询结果
	rows        *sqlx.Rows
	columnNames []string          // 列名称
	columnTypes []*sql.ColumnType // 列类型

	// 数据库每遍历N次延迟多久
	delayDuration time.Duration // delayTime解析结果
	rowNextNumber int           // 记录数据库遍历次数
}

func (m *MySQL) Query() error {
	delayDuration, err := time.ParseDuration(m.delayTime)
	if err != nil {
		return err
	}

	rows, err := mysql.DB.Queryx(m.execute)
	if err != nil {
		return err
	}

	columnNames, err := rows.Columns()
	if err != nil {
		return err
	}

	// 获取列类型
	columnTypes, err := rows.ColumnTypes()
	if err != nil {
		return err
	}

	m.delayDuration = delayDuration
	m.rows = rows
	m.columnNames = columnNames
	m.columnTypes = columnTypes

	return nil
}

func (m *MySQL) ParseRow(row []any) ([]any, error) {
	var rowValue []any
	for i, v := range row {
		// 空值
		if v == nil {
			rowValue = append(rowValue, excelize.Cell{Value: nil})
			continue
		}

		// 获取数据库类型
		dTypeName := m.columnTypes[i].DatabaseTypeName()

		// 样式
		style, err := excel.getStyleID(i + 1)
		if err != nil {
			return nil, err
		}

		// 字符串类型
		if in(dTypeName, []string{"CHAR", "VARCHAR", "TEXT"}) {
			valueStr := string(v.([]byte))
			valueInt, err := strconv.Atoi(valueStr)
			if err == nil {
				rowValue = append(rowValue, excelize.Cell{Value: valueInt, StyleID: style})
			} else {
				rowValue = append(rowValue, excelize.Cell{Value: valueStr, StyleID: style})
			}
			continue
		}
		if in(dTypeName, []string{"BINARY", "VARBINARY", "BLOB"}) {
			// 转为Go十六进制表示
			value := fmt.Sprintf("0x%X\n", v)

			// 收集对象
			rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
			continue
		}

		// 数字类型
		if in(dTypeName, []string{"TINYINT", "SMALLINT", "MEDIUMINT", "INT", "BIGINT"}) {
			// 转为Go数字类型
			value, err := strconv.Atoi(string(v.([]byte)))
			if err != nil {
				return nil, err
			}
			// 收集对象
			rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
			continue
		}
		if in(dTypeName, []string{"DECIMAL", "FLOAT", "DOUBLE"}) {
			// 转为Go数字类型
			value, err := strconv.ParseFloat(string(v.([]byte)), 10)
			if err != nil {
				return nil, err
			}
			// 收集对象
			rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
			continue
		}
		if in(dTypeName, []string{"BIT"}) {
			value := fmt.Sprintf("0x%X\n", v)
			rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
			continue
		}

		// 时间类型
		if in(dTypeName, []string{"DATE"}) {
			value := v.(time.Time).Format(time.DateOnly)
			rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
			continue
		}
		if in(dTypeName, []string{"DATETIME", "TIMESTAMP"}) {
			value := v.(time.Time).Format(time.DateTime)
			rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
			continue
		}
		if in(dTypeName, []string{"TIME", "YEAR"}) {
			valueStr := string(v.([]byte))
			valueInt, err := strconv.Atoi(valueStr)
			if err == nil {
				rowValue = append(rowValue, excelize.Cell{Value: valueInt, StyleID: style})
			} else {
				rowValue = append(rowValue, excelize.Cell{Value: valueStr, StyleID: style})
			}
			continue
		}

		// 不支持的数据库类型
		logger.Error(fmt.Sprintf("Unsupported database type: %s", dTypeName))
		return nil, err
	}
	return rowValue, nil
}

func (m *MySQL) CheckSleep() {
	m.rowNextNumber++
	if m.rowNextNumber >= m.batchSize {
		time.Sleep(m.delayDuration)
		m.rowNextNumber = 0
	}
}

type Excel struct {
	// flags
	password       string // 设置密码
	output         string // 输出文件
	sheetName      string // 单个工作表直接使用此名称,多个工作表会自动添加数字后缀:-N
	styleRowHeight string // 行高
	styleColWidth  string // 列宽度
	styleColAlign  string // 列对齐

	// 存储样式解析结果
	rowHeightMap map[int]float64 // 存储行高的Map
	colAlignMap  map[int]string  // 存储列对齐的Map

	// StreamWriter
	f              *excelize.File
	sw             *excelize.StreamWriter // 每个Sheet拥有一个专属的StreamWriter
	header         []any                  // 表头
	curTotalLine   int                    // 当前累计写入了多少行
	curSheetLine   int                    // 当前Sheet写入了多少行
	maxSheetNumber int                    // 每个Sheet最多允许写入多少行
}

func (e *Excel) NewStreamWriter() error {
	f := excelize.NewFile()

	sw, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		return err
	}

	e.f = f
	e.sw = sw
	e.rowHeightMap = make(map[int]float64)
	e.colAlignMap = make(map[int]string)

	return nil
}

func (e *Excel) MustClose() {
	err := e.sw.Flush()
	if err != nil {
		logger.Fatal(err.Error())
	}

	err = e.f.SaveAs(e.output, excelize.Options{Password: e.password})
	if err != nil {
		logger.Fatal(err.Error())
	}

	err = e.f.Close()
	if err != nil {
		logger.Fatal(err.Error())
	}
}

func (e *Excel) SetHeader(header []any) {
	e.header = header
	if len(e.header) > 0 {
		e.maxSheetNumber += 1
	}
}

func (e *Excel) AddRow(values []any) error {
	// 超过最大行则新建Sheet
	if e.curSheetLine+1 > e.maxSheetNumber {
		err := e.sw.Flush()
		if err != nil {
			return err
		}

		name := "Sheet" + strconv.Itoa(e.curTotalLine/e.maxSheetNumber+1)

		_, err = e.f.NewSheet(name)
		if err != nil {
			return err
		}

		e.sw, err = e.f.NewStreamWriter(name)
		if err != nil {
			return err
		}

		e.curSheetLine = 0

		// 重新设置列宽
		err = e.SetColWidth()
		if err != nil {
			return err
		}
	}

	// 找到行高
	height, _ := excel.rowHeightMap[e.curSheetLine+1]

	// 第一行添加表头
	if e.curSheetLine == 0 && len(e.header) > 0 {
		err := e.sw.SetRow("A1", e.header, excelize.RowOpts{Height: height})
		if err != nil {
			return err
		}
		e.curSheetLine += 1
		e.curTotalLine += 1
	}

	// 重新找到行高
	height, _ = excel.rowHeightMap[e.curSheetLine+1]

	// 写入数据
	e.curSheetLine += 1
	e.curTotalLine += 1
	cell := "A" + strconv.Itoa(e.curSheetLine)
	return e.sw.SetRow(cell, values, excelize.RowOpts{Height: height})
}

func (e *Excel) SetColWidth() error {
	list, err := e.parseStyle(e.styleColWidth)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, widthStr := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		width, err := strconv.ParseFloat(widthStr, 10)
		if err != nil {
			return err
		}

		err = e.sw.SetColWidth(min, max, width)
		if err != nil {
			return err
		}
	}
	return nil
}

func (e *Excel) SetColAlign() error {
	list, err := e.parseStyle(e.styleColAlign)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, align := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		for i := min; i <= max; i++ {
			e.colAlignMap[i] = align
		}
	}
	return nil
}

func (e *Excel) SetRowHeight() error {
	list, err := e.parseStyle(e.styleRowHeight)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, heightStr := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		height, err := strconv.ParseFloat(heightStr, 10)
		if err != nil {
			return err
		}

		for i := min; i <= max; i++ {
			e.rowHeightMap[i] = height
		}
	}

	return nil
}

func (e *Excel) parseStyle(style string) (list [][]string, err error) {
	if style == "" {
		return
	}
	styleList := strings.Split(style, ",") // []string{"1:10", "2-7:40"}
	for _, element := range styleList {

		item := strings.Split(element, ":") // []string{"2-7", "40"}
		if len(item) < 2 {
			return nil, fmt.Errorf("parse style error: %s", element)
		}

		key, value := item[0], item[1]     // key="2-7", value="40"
		keyList := strings.Split(key, "-") // keyList=[]string{"2",7}

		var min, max string
		min = keyList[0]
		if len(keyList) >= 2 {
			max = keyList[1]
		} else {
			max = min
		}
		list = append(list, []string{min, max, value})
	}

	return
}

func (e *Excel) getStyleID(index int) (int, error) {
	// 列对齐方式
	align, ok := e.colAlignMap[index]
	if !ok {
		align = "left"
	}

	// 生成样式
	style, err := e.f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: align,
			Vertical:   "center",
		}})

	return style, err
}

func (e *Excel) SetSheetName() error {
	if excel.sheetName == "" {
		return nil
	}
	sheetList := excel.f.GetSheetList()
	if len(sheetList) <= 1 {
		return excel.f.SetSheetName("Sheet1", excel.sheetName)
	}
	for i, v := range sheetList {
		err := excel.f.SetSheetName(v, excel.sheetName+"-"+strconv.Itoa(i+1))
		if err != nil {
			return err
		}
	}
	return nil
}

var rootCmd = &cobra.Command{
	Use:           "mysqlexport",
	Short:         "Export mysql to excel\nFor details, please refer to https://github.com/vvfock3r/mysqlexport",
	SilenceUsage:  true,
	SilenceErrors: true,
	PersistentPreRunE: func(cmd *cobra.Command, args []string) error {
		for _, m := range load.ModuleList {
			m.MustCheck(cmd)
		}

		for _, m := range load.ModuleList {
			err := m.Initialize(cmd)
			if err != nil {
				return err
			}
		}

		return nil
	},
	Run: func(cmd *cobra.Command, args []string) {
		// 执行SQL
		err := my.Query()
		if err != nil {
			logger.Fatal(err.Error())
		}
		defer func() { _ = my.rows.Close() }()

		// 初始化 Excel
		err = excel.NewStreamWriter()
		if err != nil {
			logger.Fatal(err.Error())
		}
		defer excel.MustClose()

		// 设置列宽
		err = excel.SetColWidth()
		if err != nil {
			logger.Fatal(err.Error())
		}

		// 设置行高
		err = excel.SetRowHeight()
		if err != nil {
			logger.Fatal(err.Error())
		}

		// 设置列对齐方式
		err = excel.SetColAlign()
		if err != nil {
			logger.Fatal(err.Error())
		}

		// 设置表头
		var header []any
		for i, value := range my.columnNames {
			style, err := excel.getStyleID(i + 1)
			if err != nil {
				logger.Fatal(err.Error())
			}
			header = append(header, excelize.Cell{Value: value, StyleID: style})
		}
		excel.SetHeader(header)

		// 遍历每一条记录
		for my.rows.Next() {
			// 获取一行
			row, err := my.rows.SliceScan()
			if err != nil {
				logger.Fatal(err.Error())
			}

			// 遍历每个字段,收集值
			rowValue, err := my.ParseRow(row)
			if err != nil {
				logger.Fatal(err.Error())
			}

			// 添加一行到Excel
			err = excel.AddRow(rowValue)
			if err != nil {
				logger.Fatal(err.Error())
			}

			// 是否休眠一下以减轻MySQL的压力
			my.CheckSleep()
		}

		// 修改Sheet名称
		err = excel.SetSheetName()
		if err != nil {
			logger.Fatal(err.Error())
		}

		// 结束
		logger.Info("execution completed")
	},
}

func in(str string, list []string) bool {
	for _, s := range list {
		if str == s {
			return true
		}
	}
	return false
}

func init() {
	// register flags or others
	for _, m := range load.ModuleList {
		m.Register(rootCmd)
	}

	// mysql flags
	rootCmd.Flags().StringVarP(&my.execute, "execute", "e", "", "execute sql command")
	rootCmd.Flags().IntVarP(&my.batchSize, "batch-size", "", 10000, "batch size")
	rootCmd.Flags().StringVarP(&my.delayTime, "sleep-time", "", "1s", "sleep time")

	err := rootCmd.MarkFlagRequired("execute")
	if err != nil {
		panic(err)
	}

	// excel flags
	rootCmd.Flags().StringVarP(&excel.output, "output", "o", "", "output xlsx file")
	rootCmd.Flags().StringVarP(&excel.password, "excel-password", "", "", "excel-password")
	rootCmd.Flags().StringVarP(&excel.sheetName, "sheet-name", "", "", "sheet name")
	rootCmd.Flags().IntVarP(&excel.maxSheetNumber, "sheet-line", "", 1000000, "max line per sheet")
	rootCmd.Flags().StringVarP(&excel.styleColWidth, "col-width", "", "", "col-width")
	rootCmd.Flags().StringVarP(&excel.styleColAlign, "col-align", "", "", "col align")
	rootCmd.Flags().StringVarP(&excel.styleRowHeight, "row-height", "", "", "row height")

	err = rootCmd.MarkFlagRequired("output")
	if err != nil {
		panic(err)
	}
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		_, _ = fmt.Fprintf(os.Stderr, "%v\n", err)
	}
}
