package cmd

import (
	"database/sql"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/jmoiron/sqlx"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"

	"github.com/vvfock3r/mysqlexport/kernel/load"
	"github.com/vvfock3r/mysqlexport/kernel/module/logger"
	"github.com/vvfock3r/mysqlexport/kernel/module/mysql"
)

var (
	my    = NewMySQL()
	excel = NewExcel()
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

func NewMySQL() *MySQL {
	return &MySQL{}
}

func (m *MySQL) Ping() error {
	return mysql.DB.Ping()
}

func (m *MySQL) Close() error {
	return mysql.DB.Close()
}

func (m *MySQL) Query() error {
	delayDuration, err := time.ParseDuration(m.delayTime)
	if err != nil {
		return err
	}

	// SQL语句必须以SELECT开头,仅仅为了安全考虑,如果有特殊需求可以将下面的代码删除
	if !strings.HasPrefix(strings.ToUpper(m.execute), "SELECT") {
		return fmt.Errorf("the sql statement must start with the select keyword")
	}

	// 执行查询
	rows, err := mysql.DB.Queryx(m.execute)
	if err != nil {
		return err
	}

	// 获取列名称
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

func (m *MySQL) ParseRow(row []any) ([]excelize.Cell, error) {
	var rowValue []excelize.Cell
	for i, v := range row {
		// 空值
		if v == nil {
			rowValue = append(rowValue, excelize.Cell{Value: nil})
			continue
		}

		// 获取数据库类型
		dTypeName := m.columnTypes[i].DatabaseTypeName()

		// 字符串类型
		if in(dTypeName, []string{"CHAR", "VARCHAR", "TEXT"}) {
			valueStr := string(v.([]byte))
			valueInt, err := strconv.Atoi(valueStr)
			if err == nil {
				rowValue = append(rowValue, excelize.Cell{Value: valueInt})
			} else {
				rowValue = append(rowValue, excelize.Cell{Value: valueStr})
			}
			continue
		}
		if in(dTypeName, []string{"BINARY", "VARBINARY", "BLOB", "GEOMETRY"}) {
			// 转为Go十六进制表示
			value := fmt.Sprintf("0x%X\n", v)

			// 收集对象
			rowValue = append(rowValue, excelize.Cell{Value: value})
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
			rowValue = append(rowValue, excelize.Cell{Value: value})
			continue
		}
		if in(dTypeName, []string{"DECIMAL", "FLOAT", "DOUBLE"}) {
			// 转为Go数字类型
			value, err := strconv.ParseFloat(string(v.([]byte)), 10)
			if err != nil {
				return nil, err
			}
			// 收集对象
			rowValue = append(rowValue, excelize.Cell{Value: value})
			continue
		}
		if in(dTypeName, []string{"BIT"}) {
			value := fmt.Sprintf("0x%X\n", v)
			rowValue = append(rowValue, excelize.Cell{Value: value})
			continue
		}

		// 时间类型
		if in(dTypeName, []string{"DATE"}) {
			value := v.(time.Time).Format(time.DateOnly)
			rowValue = append(rowValue, excelize.Cell{Value: value})
			continue
		}
		if in(dTypeName, []string{"DATETIME", "TIMESTAMP"}) {
			value := v.(time.Time).Format(time.DateTime)
			rowValue = append(rowValue, excelize.Cell{Value: value})
			continue
		}
		if in(dTypeName, []string{"TIME", "YEAR"}) {
			valueStr := string(v.([]byte))
			valueInt, err := strconv.Atoi(valueStr)
			if err == nil {
				rowValue = append(rowValue, excelize.Cell{Value: valueInt})
			} else {
				rowValue = append(rowValue, excelize.Cell{Value: valueStr})
			}
			continue
		}

		// JSON类型
		if in(dTypeName, []string{"JSON"}) {
			value := string(v.([]byte))
			rowValue = append(rowValue, excelize.Cell{Value: value})
			continue
		}

		// 未测试过或不支持的类型
		value, ok := v.([]byte)
		if ok {
			logger.Warn(fmt.Sprintf("untested database type: %s, column name: %s", dTypeName, m.columnNames[i]))
			rowValue = append(rowValue, excelize.Cell{Value: string(value)})
		} else {
			logger.Fatal(fmt.Sprintf("unsupported database type: %s, column name: %s", dTypeName, m.columnNames[i]))
		}
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
	// 概念说明: 工作簿 workbook,工作表 sheet

	// flags
	password          string // 设置密码
	output            string // 输出文件
	sheetName         string // 单个工作表直接使用此名称,多个工作表会自动添加数字后缀:-N
	styleRowHeight    string // 行高
	styleColWidth     string // 列宽度
	styleColAlign     string // 列对齐
	styleRowBgColor   string // 背景色
	styleColBgColor   string // 背景色
	styleRowFontColor string // 字体颜色
	styleColFontColor string // 字体颜色

	// 存储样式解析结果和表头等一般不会变的数据
	rowHeightMap    map[int]float64 // 存储行高的Map
	colAlignMap     map[int]string  // 存储列对齐的Map
	rowBgColorMap   map[int]string  // 存储背景色的Map
	colBgColorMap   map[int]string  // 存储背景色的Map
	rowFontColorMap map[int]string  // 存储字体颜色的Map
	colFontColorMap map[int]string  // 存储字体颜色的Map

	// 表头
	header []excelize.Cell

	// StreamWriter
	f                  *excelize.File
	sw                 *excelize.StreamWriter // 每个Sheet拥有一个专属的StreamWriter
	maxWorkbookLine    int                    // 每个Workbook最多允许写入多少行，不包含表头
	maxSheetLine       int                    // 每个Sheet最多允许写入多少行，不包含表头
	curWorkbookLine    int                    // 当前Workbook累计写入了多少行，不包含表头
	curSheetLine       int                    // 当前Sheet写入了多少行，不包含表头
	curSheetHeaderLine int                    // 当前Sheet写入了多少行，包含表头
	curlTotalLine      int                    // 当前总共写入了多少行，不包含表头
}

func NewExcel() *Excel {
	return &Excel{
		f:               excelize.NewFile(),
		rowHeightMap:    make(map[int]float64),
		colAlignMap:     make(map[int]string),
		rowBgColorMap:   make(map[int]string),
		colBgColorMap:   make(map[int]string),
		rowFontColorMap: make(map[int]string),
		colFontColorMap: make(map[int]string),
	}
}

func (e *Excel) NewStreamWriter() (err error) {
	e.sw, err = e.f.NewStreamWriter("Sheet1")
	return err
}

func (e *Excel) getOutput() (string, error) {
	// 只有一个工作簿的情况下
	if e.maxWorkbookLine <= 0 || e.curlTotalLine < e.maxWorkbookLine {
		return e.output, nil
	}

	// 转为绝对路径
	absOutput, err := filepath.Abs(e.output)
	if err != nil {
		return "", err
	}

	// 绝对路径分割为 路径 和 文件名
	dir, fileName := filepath.Split(absOutput)

	// 文件名分割为 名称 和 扩展名, 名称中允许包含.
	outputList := strings.Split(fileName, ".")
	name := strings.Join(outputList[:len(outputList)-1], ".")
	ext := outputList[len(outputList)-1]

	index := math.Ceil(float64(e.curlTotalLine) / float64(e.maxWorkbookLine))
	indexStr := strconv.FormatFloat(index, 'f', 0, 64)

	// 组合出新路径
	newOutput := strings.Join([]string{dir, name, "-", indexStr, ".", ext}, "")
	return newOutput, nil
}

func (e *Excel) MustClose() {
	err := e.sw.Flush()
	if err != nil {
		logger.Fatal(err.Error())
	}

	output, err := e.getOutput()
	if err != nil {
		logger.Fatal(err.Error())
	}

	err = e.f.SaveAs(output, excelize.Options{Password: e.password})
	if err != nil {
		logger.Fatal(err.Error())
	}

	err = e.f.Close()
	if err != nil {
		logger.Fatal(err.Error())
	}
}

func (e *Excel) SetHeader(header []excelize.Cell) {
	e.header = header
}

func (e *Excel) AddRow(values []excelize.Cell) error {
	// 超过工作簿最大行数则重新建一个
	if e.maxWorkbookLine > 0 && e.curWorkbookLine+1 > e.maxWorkbookLine {
		// 修改Sheet名称
		err := e.SetSheetName()
		if err != nil {
			logger.Fatal(err.Error())
		}

		// 保存
		e.MustClose()

		e.f = excelize.NewFile()
		err = e.NewStreamWriter()
		if err != nil {
			return err
		}
		e.curSheetLine = 0
		e.curWorkbookLine = 0
		e.curSheetHeaderLine = 0

		err = e.SetStyle()
		if err != nil {
			logger.Fatal(err.Error())
		}
	}

	// 超过工作表最大行数则重新建一个
	if e.curSheetLine+1 > e.maxSheetLine {
		err := e.sw.Flush()
		if err != nil {
			return err
		}

		name := "Sheet" + strconv.Itoa(e.curWorkbookLine/e.maxSheetLine+1)
		_, err = e.f.NewSheet(name)
		if err != nil {
			return err
		}

		e.sw, err = e.f.NewStreamWriter(name)
		if err != nil {
			return err
		}

		e.curSheetLine = 0
		e.curSheetHeaderLine = 0

		// 重新设置列宽
		err = e.SetColWidth()
		if err != nil {
			return err
		}
	}

	// 第一行添加表头
	if e.curSheetLine == 0 && len(e.header) > 0 {
		// 设置样式
		for i := range e.header {
			style, err := e.getStyleID(e.curSheetHeaderLine+1, i+1)
			if err != nil {
				return err
			}
			e.header[i].StyleID = style
		}

		// 类型转换
		valueAny := e.ConvertAny(e.header)

		// 添加行
		err := e.sw.SetRow("A1", valueAny, excelize.RowOpts{Height: e.getNextRowHeight()})
		if err != nil {
			return err
		}
		e.curSheetHeaderLine++
	}

	// 设置颜色样式
	for i := range values {
		style, err := e.getStyleID(e.curSheetHeaderLine+1, i+1)
		if err != nil {
			return err
		}
		values[i].StyleID = style
	}

	// 类型转换
	valueAny := e.ConvertAny(values)

	// 写入数据
	cell := "A" + strconv.Itoa(e.curSheetHeaderLine+1)
	err := e.sw.SetRow(cell, valueAny, excelize.RowOpts{Height: e.getNextRowHeight()})
	if err != nil {
		return err
	}

	// 计数加1
	e.curSheetLine++
	e.curSheetHeaderLine++
	e.curWorkbookLine++
	e.curlTotalLine++

	return nil
}

func (e *Excel) ConvertAny(cells []excelize.Cell) []any {
	var values []any
	for _, cell := range cells {
		values = append(values, cell)
	}
	return values
}

func (e *Excel) SetStyle() error {
	err := e.SetColWidth()
	if err != nil {
		return err
	}

	// 设置行高
	err = e.SetRowHeight()
	if err != nil {
		return err
	}

	// 设置列对齐方式
	err = e.SetColAlign()
	if err != nil {
		return err
	}

	// 设置列背景色
	err = e.SetColBgColor()
	if err != nil {
		return err
	}

	// 设置行背景色
	err = e.SetRowBgColor()
	if err != nil {
		return err
	}

	// 设置列字体颜色
	err = e.SetColFontColor()
	if err != nil {
		return err
	}

	// 设置行字体颜色
	err = e.SetRowFontColor()
	if err != nil {
		return err
	}

	return nil
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

func (e *Excel) SetRowBgColor() error {
	list, err := e.parseStyle(e.styleRowBgColor)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, bgcolor := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		for i := min; i <= max; i++ {
			e.rowBgColorMap[i] = bgcolor
		}
	}
	return nil
}

func (e *Excel) SetColBgColor() error {
	list, err := e.parseStyle(e.styleColBgColor)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, bgcolor := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		for i := min; i <= max; i++ {
			e.colBgColorMap[i] = bgcolor
		}
	}
	return nil
}

func (e *Excel) SetRowFontColor() error {
	list, err := e.parseStyle(e.styleRowFontColor)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, fontColor := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		for i := min; i <= max; i++ {
			e.rowFontColorMap[i] = fontColor
		}
	}
	return nil
}

func (e *Excel) SetColFontColor() error {
	list, err := e.parseStyle(e.styleColFontColor)
	if err != nil {
		return err
	}

	for _, item := range list {
		minStr, maxStr, fontColor := item[0], item[1], item[2]

		min, err := strconv.Atoi(minStr)
		if err != nil {
			return err
		}

		max, err := strconv.Atoi(maxStr)
		if err != nil {
			return err
		}

		for i := min; i <= max; i++ {
			e.colFontColorMap[i] = fontColor
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

func (e *Excel) getStyleID(rowIndex, colIndex int) (int, error) {
	// 样式对象
	style := &excelize.Style{}

	// 对齐样式
	align, ok := e.colAlignMap[colIndex]
	if !ok {
		align = "left"
	}
	style.Alignment = &excelize.Alignment{
		Horizontal: align,
		Vertical:   "center",
	}

	// 列背景颜色
	colBgColor, ok := e.colBgColorMap[colIndex]
	if ok {
		style.Fill = excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{colBgColor},
		}
	}

	// 行背景颜色
	rowBgColor, ok := e.rowBgColorMap[rowIndex]
	if ok {
		style.Fill = excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{rowBgColor},
		}
	}

	// 列字体颜色
	colFontColor, ok := e.colFontColorMap[colIndex]
	if ok {
		style.Font = &excelize.Font{
			Color: colFontColor,
		}
	}

	// 行字体颜色
	rowFontColor, ok := e.rowFontColorMap[rowIndex]
	if ok {
		style.Font = &excelize.Font{
			Color: rowFontColor,
		}
	}

	// 生成样式
	styleID, err := e.f.NewStyle(style)

	return styleID, err
}

func (e *Excel) getNextRowHeight() float64 {
	height, _ := e.rowHeightMap[e.curSheetHeaderLine+1]
	return height
}

func (e *Excel) SetSheetName() error {
	if e.sheetName == "" {
		return nil
	}
	sheetList := e.f.GetSheetList()
	if len(sheetList) <= 1 {
		return e.f.SetSheetName("Sheet1", e.sheetName)
	}
	for i, v := range sheetList {
		err := e.f.SetSheetName(v, e.sheetName+"-"+strconv.Itoa(i+1))
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
		// 连接数据库
		err := my.Ping()
		if err != nil {
			logger.Fatal("connect database error", zap.Error(err))
		}
		logger.Info("connect database success")
		defer func() { _ = my.Close() }()

		// 执行SQL
		err = my.Query()
		if err != nil {
			logger.Fatal(err.Error())
		}
		defer func() { _ = my.rows.Close() }()

		// 初始化Excel流式写入器
		err = excel.NewStreamWriter()
		if err != nil {
			logger.Fatal(err.Error())
		}
		defer excel.MustClose()

		// 设置样式
		err = excel.SetStyle()
		if err != nil {
			logger.Fatal(err.Error())
		}

		// 设置表头
		var header []excelize.Cell
		for i, value := range my.columnNames {
			style, err := excel.getStyleID(1, i+1)
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
	rootCmd.Flags().StringVarP(&my.execute, "execute", "e", "", "specifies the SQL command to be executed.")
	rootCmd.Flags().IntVarP(&my.batchSize, "batch-size", "", 10000, "specifies the batch size to use when executing SQL commands")
	rootCmd.Flags().StringVarP(&my.delayTime, "delay-time", "", "1s", "specifies the time to delay between batches when executing SQL")

	err := rootCmd.MarkFlagRequired("execute")
	if err != nil {
		panic(err)
	}

	// excel flags
	rootCmd.Flags().StringVarP(&excel.output, "output", "o", "", "specifies the name of the output Excel file")
	rootCmd.Flags().StringVarP(&excel.password, "setup-password", "", "", "specifies the password for the Excel file")
	rootCmd.Flags().StringVarP(&excel.sheetName, "sheet-name", "", "", "specifies the name of the sheet in the Excel file")
	rootCmd.Flags().IntVarP(&excel.maxSheetLine, "sheet-line", "", 1000000, "specifies the maximum number of lines per sheet in the Excel file")
	rootCmd.Flags().IntVarP(&excel.maxWorkbookLine, "workbook-line", "", -1, "specifies the maximum number of lines all sheet in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleRowHeight, "row-height", "", "", "specifies the row height in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleRowBgColor, "row-bg-color", "", "", "specifies the row background color in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleRowFontColor, "row-font-color", "", "", "specifies the row font color in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleColWidth, "col-width", "", "", "specifies the column width in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleColAlign, "col-align", "", "", "specifies the column alignment in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleColBgColor, "col-bg-color", "", "", "specifies column background color in the Excel file")
	rootCmd.Flags().StringVarP(&excel.styleColFontColor, "col-font-color", "", "", "specifies column font color in the Excel file")

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
