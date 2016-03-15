package main

import (
	"github.com/shibukawa/xlsxrange"
	"github.com/tealeg/xlsx"
)

func ParseConstant(file *xlsx.File) map[string]*xlsx.Cell {
	result := make(map[string]*xlsx.Cell)
	for _, name := range file.DefinedNames {
		if name.Function {
			continue
		}
		aRange := xlsxrange.NewWithFile(file, name.Data)
		result[name.Name] = aRange.GetCell()
	}
	return result
}
