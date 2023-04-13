package cmd

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"

	"github.com/vvfock3r/mysqlexport/kernel/load"
	"github.com/vvfock3r/mysqlexport/kernel/module/logger"
	"github.com/vvfock3r/mysqlexport/kernel/module/mysql"
)

var (
	execute     string // 执行的SQL查询命令
	batch       int    // 数据库每遍历N次
	sleep       string // 休眠N秒
	output      string // 输出文件名
	sheetName   string // 工作表名称，多个工作表会添加-1/-2后缀
	sheetLine   int    // 每个工作表最大的行树
	colWidth    string // 设置列宽，单独设置数字则代表设置所有列, A:20,B:30则代表单独设置列
	colAlign    string // 列对齐方式
	colAlignMap = make(map[int]string)

	rowHeight    string // 设置行高
	rowHeightMap = make(map[int]float64)
	password     string
)

// StreamWriter 对excelize.StreamWriter做了一层简单封装，用于自动添加并写入数据到新Sheet中
type StreamWriter struct {
	f              *excelize.File         // File对象,用于创建新的Sheet和StreamWriter
	w              *excelize.StreamWriter // 每个Sheet拥有一个专属的StreamWriter,当前的StreamWriter
	header         []any                  // 表头
	curTotalLine   int                    // 当前总共写入了多少行，包含表头(若有)
	curSheetLine   int                    // 当前Sheet写入了多少行，包含表头(若有)
	maxSheetNumber int                    // 每个Sheet最多允许写入多少行，自动添加表头行数(若有)
}

func NewStreamWriter(f *excelize.File, maxSheetNumber int) (*StreamWriter, error) {
	w, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		return nil, err
	}
	s := &StreamWriter{f: f, w: w, maxSheetNumber: maxSheetNumber}
	return s, nil
}

func (s *StreamWriter) SetHeader(header []any) {
	s.header = header
	if len(s.header) > 0 {
		s.maxSheetNumber += 1
	}
}

func (s *StreamWriter) AddRow(values []any) error {
	// 超过最大行则新建Sheet
	if s.curSheetLine+1 > s.maxSheetNumber {
		err := s.Flush()
		if err != nil {
			return err
		}

		name := "Sheet" + strconv.Itoa(s.curTotalLine/s.maxSheetNumber+1)

		_, err = s.f.NewSheet(name)
		if err != nil {
			return err
		}

		s.w, err = s.f.NewStreamWriter(name)
		if err != nil {
			return err
		}

		s.curSheetLine = 0

		// 重新设置列宽
		colWidthList := strings.Split(colWidth, ",")
		for _, v := range colWidthList {
			item := strings.Split(v, ":")
			if len(item) < 2 {
				continue
			}
			colStr, widthStr := item[0], item[1]

			colList := strings.Split(colStr, "-")

			minStr := colList[0]
			min, err := strconv.Atoi(minStr)
			if err != nil {
				panic(err)
			}
			var max int
			if len(colList) >= 2 {
				maxStr := colList[1]
				max, err = strconv.Atoi(maxStr)
				if err != nil {
					panic(err)
				}
			} else {
				max = min
			}

			width, err := strconv.ParseFloat(widthStr, 10)
			if err != nil {
				panic(err)
			}
			err = s.w.SetColWidth(min, max, width)
			if err != nil {
				panic(err)
			}
		}
	}

	// 找到行高
	height, _ := rowHeightMap[s.curSheetLine+1]

	// 第一行添加表头
	if s.curSheetLine == 0 && len(s.header) > 0 {
		err := s.w.SetRow("A1", s.header, excelize.RowOpts{Height: height})
		if err != nil {
			return err
		}
		s.curSheetLine += 1
		s.curTotalLine += 1
	}

	// 重新找到行高
	height, _ = rowHeightMap[s.curSheetLine+1]

	// 写入数据
	s.curSheetLine += 1
	s.curTotalLine += 1
	cell := "A" + strconv.Itoa(s.curSheetLine)
	return s.w.SetRow(cell, values, excelize.RowOpts{Height: height})
}

func (s *StreamWriter) Flush() error {
	return s.w.Flush()
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
		run()
	},
}

func run() {
	// 执行SQL
	rows, err := mysql.DB.Queryx(execute)
	if err != nil {
		panic(err)
	}

	// 获取列名称
	columns, err := rows.Columns()
	if err != nil {
		panic(err)
	}

	// 获取列类型
	columnsTypes, err := rows.ColumnTypes()
	if err != nil {
		panic(err)
	}

	// 初始化 Excel
	f := excelize.NewFile()
	defer func() {
		err = f.SaveAs(output, excelize.Options{Password: password})
		if err != nil {
			panic(err)
		}
		err = f.Close()
		if err != nil {
			panic(err)
		}
	}()

	// 初始化 Excel流式写入
	w, err := NewStreamWriter(f, sheetLine)
	if err != nil {
		panic(err)
	}
	defer func() {
		err = w.Flush()
		if err != nil {
			panic(err)
		}
	}()

	// 设置列宽,格式为: 1:10,2-3:20,代表第1列宽度为10,第2和3列宽度为20
	colWidthList := strings.Split(colWidth, ",")
	for _, v := range colWidthList {
		item := strings.Split(v, ":")
		if len(item) < 2 {
			continue
		}
		colStr, widthStr := item[0], item[1]

		colList := strings.Split(colStr, "-")

		minStr := colList[0]
		min, err := strconv.Atoi(minStr)
		if err != nil {
			panic(err)
		}
		var max int
		if len(colList) >= 2 {
			maxStr := colList[1]
			max, err = strconv.Atoi(maxStr)
			if err != nil {
				panic(err)
			}
		} else {
			max = min
		}

		width, err := strconv.ParseFloat(widthStr, 10)
		if err != nil {
			panic(err)
		}
		err = w.w.SetColWidth(min, max, width)
		if err != nil {
			panic(err)
		}
	}

	// 解析行高 1:20,2:30
	rowHeightList := strings.Split(rowHeight, ",")
	for _, v := range rowHeightList {
		item := strings.Split(v, ":")
		if len(item) < 2 {
			continue
		}
		rowStr, heightStr := item[0], item[1] // 1:20

		rowList := strings.Split(rowStr, "-") // 1-2:20

		minStr := rowList[0]
		min, err := strconv.Atoi(minStr)
		if err != nil {
			panic(err)
		}
		var max int
		if len(rowList) >= 2 {
			maxStr := rowList[1]
			max, err = strconv.Atoi(maxStr)
			if err != nil {
				panic(err)
			}
		} else {
			max = min
		}

		height, err := strconv.ParseFloat(heightStr, 10)
		if err != nil {
			panic(err)
		}

		for i := min; i <= max; i++ {
			rowHeightMap[i] = height
		}

	}

	// 解析列对齐方式 1:center,2-3:left
	colAlignList := strings.Split(colAlign, ",")
	for _, v := range colAlignList {
		item := strings.Split(v, ":")
		if len(item) < 2 {
			continue
		}
		colStr, align := item[0], item[1] // 1:left

		rowList := strings.Split(colStr, "-") // 1-2:left

		minStr := rowList[0]
		min, err := strconv.Atoi(minStr)
		if err != nil {
			panic(err)
		}
		var max int
		if len(rowList) >= 2 {
			maxStr := rowList[1]
			max, err = strconv.Atoi(maxStr)
			if err != nil {
				panic(err)
			}
		} else {
			max = min
		}

		for i := min; i <= max; i++ {
			colAlignMap[i] = align
		}
	}

	// 设置表头(对所有Sheet生效)
	var header []any
	for i, value := range columns {
		style, err := getStyleID(f, i+1)
		if err != nil {
			panic(err)
		}
		header = append(header, excelize.Cell{Value: value, StyleID: style})
	}
	w.SetHeader(header)

	// 遍历每一条记录
	n := 0
	t, err := time.ParseDuration(sleep)
	if err != nil {
		panic(err)
	}
	for rows.Next() {
		// 获取一行
		row, err := rows.SliceScan()
		if err != nil {
			panic(err)
		}

		// 遍历每个字段,收集值
		var rowValue []any
		for i, v := range row {
			// 空值
			if v == nil {
				rowValue = append(rowValue, excelize.Cell{Value: nil})
				continue
			}

			// 获取数据库类型
			dTypeName := columnsTypes[i].DatabaseTypeName()

			// 样式
			style, err := getStyleID(f, i+1)
			if err != nil {
				panic(err)
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
					panic(err)
				}
				// 收集对象
				rowValue = append(rowValue, excelize.Cell{Value: value, StyleID: style})
				continue
			}
			if in(dTypeName, []string{"DECIMAL", "FLOAT", "DOUBLE"}) {
				// 转为Go数字类型
				value, err := strconv.ParseFloat(string(v.([]byte)), 10)
				if err != nil {
					panic(err)
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

			// 不支持的类型
			logger.Error(fmt.Sprintf("Unsupported database type: %s", dTypeName))
			os.Exit(1)
		}

		// 添加一行到Excel
		err = w.AddRow(rowValue)
		if err != nil {
			panic(err)
		}

		// 是否休眠一下以减轻MySQL的压力
		n++
		if n >= batch {
			time.Sleep(t)
			n = 0
		}
	}

	// 修改Sheet名称
	sheetList := f.GetSheetList()
	if len(sheetList) <= 1 {
		err := f.SetSheetName("Sheet1", sheetName)
		if err != nil {
			panic(err)
		}
	} else {
		for i, v := range sheetList {
			err := f.SetSheetName(v, sheetName+"-"+strconv.Itoa(i+1))
			if err != nil {
				panic(err)
			}
		}
	}

	// 结束
	logger.Info("execution completed")
}

func in(str string, list []string) bool {
	for _, s := range list {
		if str == s {
			return true
		}
	}
	return false
}

func getStyleID(f *excelize.File, index int) (int, error) {
	align, ok := colAlignMap[index]
	if !ok {
		align = "left"
	}
	style, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: align,
			Vertical:   "center",
		}})
	if err != nil {
		return 0, err
	}
	return style, nil
}

func init() {
	// register flags or others
	for _, m := range load.ModuleList {
		m.Register(rootCmd)
	}

	rootCmd.Flags().StringVarP(&execute, "execute", "e", "", "execute sql command")
	err := rootCmd.MarkFlagRequired("execute")
	if err != nil {
		panic(err)
	}

	rootCmd.Flags().IntVarP(&batch, "batch-size", "", 10000, "batch size")
	rootCmd.Flags().StringVarP(&sleep, "sleep-time", "", "1s", "sleep time")

	rootCmd.Flags().StringVarP(&output, "output", "o", "", "output xlsx file")
	err = rootCmd.MarkFlagRequired("output")
	if err != nil {
		panic(err)
	}

	// 工作表
	rootCmd.Flags().StringVarP(&sheetName, "sheet-name", "", "Sheet", "sheet name")
	rootCmd.Flags().IntVarP(&sheetLine, "sheet-line", "", 1000000, "max line per sheet")

	// 样式
	rootCmd.Flags().StringVarP(&colWidth, "col-width", "", "", "col-width")
	rootCmd.Flags().StringVarP(&colAlign, "col-align", "", "left", "col align")
	rootCmd.Flags().StringVarP(&rowHeight, "row-height", "", "", "row height")

	// 密码
	rootCmd.Flags().StringVarP(&password, "excel-password", "", "", "excel-password")
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		_, _ = fmt.Fprintf(os.Stderr, "%v\n", err)
	}
}
