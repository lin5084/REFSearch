package main

import (
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/sys/windows"
)

func init() {
	stdin := windows.Handle(os.Stdin.Fd())
	var mode uint32
	_ = windows.GetConsoleMode(stdin, &mode)
	_ = windows.SetConsoleMode(stdin, mode|windows.ENABLE_VIRTUAL_TERMINAL_PROCESSING)
	_ = windows.SetConsoleOutputCP(65001)
}

type MergeInfo struct {
	startRow int
	endRow   int
	startCol int
	endCol   int
}

type MatchItem struct {
	fileName string
	rowNo    int
}

// splitRefDes 分割位号：英文逗号、中文逗号、和
func splitRefDes(s string) []string {
	s = strings.TrimSpace(s)
	if s == "" {
		return nil
	}
	re := regexp.MustCompile(`[,，和]`)
	parts := re.Split(s, -1)
	var res []string
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			res = append(res, p)
		}
	}
	return res
}

// toLowerAll 全部转小写用于不区分大小写比对
func toLowerAll(list []string) []string {
	out := make([]string, len(list))
	for i, v := range list {
		out[i] = strings.ToLower(v)
	}
	return out
}

// getMergeMap 获取合并单元格映射
func getMergeMap(f *excelize.File, sheet string) (map[string]*MergeInfo, error) {
	mergeCellMap := make(map[string]*MergeInfo)
	ranges, err := f.GetMergeCells(sheet)
	if err != nil {
		return nil, err
	}
	for _, rng := range ranges {
		sAxis, eAxis := rng.GetStartAxis(), rng.GetEndAxis()
		colS, rowS, _ := excelize.CellNameToCoordinates(sAxis)
		colE, rowE, _ := excelize.CellNameToCoordinates(eAxis)
		mi := &MergeInfo{
			startRow: rowS,
			endRow:   rowE,
			startCol: colS,
			endCol:   colE,
		}
		for r := rowS; r <= rowE; r++ {
			for c := colS; c <= colE; c++ {
				key := fmt.Sprintf("%d-%d", r, c)
				mergeCellMap[key] = mi
			}
		}
	}
	return mergeCellMap, nil
}

// getMergedCellValue 读取合并单元格，取首格值
func getMergedCellValue(f *excelize.File, sheet string, row int, col int, mergeCellMap map[string]*MergeInfo) (string, error) {
	key := fmt.Sprintf("%d-%d", row, col)
	mi, ok := mergeCellMap[key]
	cellAxis, _ := excelize.CoordinatesToCellName(col, row)
	if !ok {
		return f.GetCellValue(sheet, cellAxis)
	}
	firstAxis, _ := excelize.CoordinatesToCellName(mi.startCol, mi.startRow)
	return f.GetCellValue(sheet, firstAxis)
}

// parseArg 解析参数：料号=位号列表
func parseArg(arg string) (targetPart string, targetRefs []string, ok bool) {
	eq := strings.Index(arg, "=")
	if eq <= 0 {
		return "", nil, false
	}
	targetPart = strings.TrimSpace(arg[:eq])
	refStr := strings.TrimSpace(arg[eq+1:])
	targetRefs = splitRefDes(refStr)
	if targetPart == "" || len(targetRefs) == 0 {
		return "", nil, false
	}
	return targetPart, targetRefs, true
}

// scanFile 扫描单个BOM文件，返回匹配条目
func scanFile(path string, targetPart string, targetRefsLower []string) ([]MatchItem, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	sheet := f.GetSheetList()[0]
	a1, err := f.GetCellValue(sheet, "A1")
	if err != nil {
		return nil, err
	}
	if a1 != "料号" {
		return nil, nil
	}

	mergeCellMap, err := getMergeMap(f, sheet)
	if err != nil {
		return nil, err
	}
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	const colA = 1
	const colF = 6
	maxRow := len(rows)
	var hits []MatchItem

	for rowIdx := 3; rowIdx < maxRow; rowIdx++ {
		excelRow := rowIdx + 1
		row := rows[rowIdx]
		aVal := ""
		if len(row) > 0 {
			aVal = strings.TrimSpace(row[0])
		}
		// 判断F列是否是合并子行：子行跳过，避免重复统计
		keyF := fmt.Sprintf("%d-%d", excelRow, colF)
		if mi, ok := mergeCellMap[keyF]; ok && mi.startRow != excelRow {
			continue
		}
		fVal, errF := getMergedCellValue(f, sheet, excelRow, colF, mergeCellMap)
		if errF != nil {
			fVal = ""
		}
		fVal = strings.TrimSpace(fVal)
		refList := splitRefDes(fVal)
		refLower := toLowerAll(refList)
		// 料号匹配（区分大小写）
		if aVal != targetPart {
			continue
		}
		// 全部目标位号都必须存在（AND）
		allFound := true
		for _, want := range targetRefsLower {
			has := false
			for _, r := range refLower {
				if r == want {
					has = true
					break
				}
			}
			if !has {
				allFound = false
				break
			}
		}
		if allFound {
			hits = append(hits, MatchItem{
				fileName: filepath.Base(path),
				rowNo:    excelRow,
			})
		}
	}
	return hits, nil
}

func main() {
	if len(os.Args) != 2 {
		fmt.Println("用法：REFSearch.exe 料号=位号")
		fmt.Println("示例：")
		fmt.Println("REFSearch.exe R123=C1")
		fmt.Println("REFSearch.exe R123=C1,C2,C3和C4")
		return
	}
	arg := os.Args[1]
	targetPart, targetRefs, ok := parseArg(arg)
	if !ok {
		fmt.Println("参数格式错误！必须使用 料号=位号 格式")
		return
	}
	targetRefsLower := toLowerAll(targetRefs)

	workDir, err := os.Getwd()
	if err != nil {
		fmt.Println("获取当前目录失败：", err)
		return
	}
	var bomFiles []string
	_ = filepath.WalkDir(workDir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return nil
		}
		if d.IsDir() {
			return nil
		}
		base := filepath.Base(path)
		if strings.HasPrefix(base, "~$") {
			return nil
		}
		if strings.ToLower(filepath.Ext(path)) == ".xlsx" {
			bomFiles = append(bomFiles, path)
		}
		return nil
	})

	var allHits []MatchItem
	for _, fp := range bomFiles {
		hits, err := scanFile(fp, targetPart, targetRefsLower)
		if err != nil {
			fmt.Printf("读取文件 %s 失败：%v\n", filepath.Base(fp), err)
			continue
		}
		allHits = append(allHits, hits...)
	}

	// 输出文本
	title := fmt.Sprintf("要查找的“%s=%s”在以下BOM中：", targetPart, strings.Join(targetRefs, ","))
	var outLines []string
	outLines = append(outLines, title)
	for _, item := range allHits {
		line := fmt.Sprintf("%s 第%d行", item.fileName, item.rowNo)
		outLines = append(outLines, line)
	}

	outText := strings.Join(outLines, "\r\n")
	fmt.Println(outText)

	// 写入 result.txt
	err = os.WriteFile("result.txt", []byte(outText), 0644)
	if err != nil {
		fmt.Printf("\n写入result.txt失败：%v\n", err)
	} else {
		fmt.Println("\n结果已保存至 result.txt")
	}
}
