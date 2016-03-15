package main

import (
	"fmt"
	"github.com/shibukawa/xlsxrange"
	"github.com/tealeg/xlsx"
	"regexp"
	"strconv"
	"strings"
)

type Schema struct {
	Functions []*Function
}

type Function struct {
	Name               string
	Formula            string
	FormulaRange       *xlsxrange.Range
	ParameterFromRange map[string]*Parameter
	Parameters         []*Parameter
}

type Parameter struct {
	Name string
}

func ParseSchema(file *xlsx.File) (*Schema, error) {
	schema := &Schema{}

	schemaSheet, ok := file.Sheet["Schema"]
	if !ok {
		return nil, fmt.Errorf("Sheet 'Schema' is missing")
	}

	var targetSheet *xlsx.Sheet
	var currentFunction *Function
	var lastType string
	for i, row := range schemaSheet.Rows[1:] {
		rowNum := i + 2
		sheetName := get(row, 0).Value
		if sheetName != "" {
			newSheet, ok := file.Sheet[sheetName]
			if !ok {
				return nil, fmt.Errorf("A%d - Sheet '%s' at line %d is missing", rowNum, sheetName, rowNum)
			}
			targetSheet = newSheet
		}
		typeName := get(row, 1).Value
		if typeName == "" {
			typeName = lastType
		} else {
			lastType = typeName
		}
		rangeLabel := get(row, 2).Value
		name := get(row, 3).Value
		if rangeLabel == "" && name == "" {
			continue
		}
		if rangeLabel == "" {
			return nil, fmt.Errorf("C%d - Range label for '%s' is missing", rowNum, name)
		}
		if name == "" {
			return nil, fmt.Errorf("D%d - Name label for range '%s' is missing", rowNum, rangeLabel)
		}
		if typeName == "function" {
			currentFunction = nil
		}
		targetRange, err := searchRange(targetSheet, rowNum, rangeLabel, currentFunction)
		if err != nil {
			return nil, err
		}
		switch typeName {
		case "function":
			currentFunction = &Function{
				Name:               name,
				Formula:            targetRange.GetCell().Formula(),
				FormulaRange:       targetRange,
				ParameterFromRange: make(map[string]*Parameter),
			}
			schema.Functions = append(schema.Functions, currentFunction)
		case "param":
			if currentFunction == nil {
				return nil, fmt.Errorf("B%d - 'param' definition appears in front of 'function' definition", rowNum)
			}
			param := &Parameter{
				Name: name,
			}
			currentFunction.ParameterFromRange[targetRange.Format(false)] = param
			currentFunction.Parameters = append(currentFunction.Parameters, param)
		}
	}
	return schema, nil
}

var emptyCell *xlsx.Cell = &xlsx.Cell{}

func get(row *xlsx.Row, colNum int) *xlsx.Cell {
	if colNum < len(row.Cells) {
		return row.Cells[colNum]
	}
	return emptyCell
}

var namedRange *regexp.Regexp = regexp.MustCompile(`([1-9][0-9]*):([^:]+)(:([1-9][0-9]*))?`)

func searchRange(sheet *xlsx.Sheet, definedRow int, rangeLabel string, currentFunction *Function) (*xlsxrange.Range, error) {
	matched := namedRange.FindStringSubmatch(rangeLabel)
	if len(matched) > 0 {
		rowNum, _ := strconv.Atoi(matched[1])
		if (rowNum - 1) >= len(sheet.Rows) {
			return nil, fmt.Errorf("C%d - Row number '%d' is out of range (%s)", definedRow, rowNum, rangeLabel)
		}
		var targetRowNum int

		if matched[3] == "" {
			if currentFunction == nil {
				return nil, fmt.Errorf("C%d - Row number (e.g. trailing 2 in '1:label:2') is not eliminable if type is function", definedRow)
			}
			targetRowNum = currentFunction.FormulaRange.Row
		} else {
			targetRowNum, _ = strconv.Atoi(matched[4])
		}
		if (targetRowNum - 1) >= len(sheet.Rows) {
			return nil, fmt.Errorf("C%d - Target row number '%d' is out of range (%s)", definedRow, targetRowNum, rangeLabel)
		}
		searchLabel := strings.TrimSpace(matched[2])
		for i, cell := range sheet.Rows[rowNum-1].Cells {
			if strings.TrimSpace(cell.Value) == searchLabel {
				return xlsxrange.New(sheet, targetRowNum, i+1), nil
			}
		}
		return nil, fmt.Errorf("C%d - Target label '%s' is missing at row %d in the sheet '%s'", definedRow, searchLabel, targetRowNum, sheet.Name)
	}
	_, _, err := xlsxrange.ParseA1Notation(rangeLabel)
	if err != nil {
		return nil, err
	}
	return xlsxrange.New(sheet, rangeLabel), nil
}
