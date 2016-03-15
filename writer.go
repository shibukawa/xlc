package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"github.com/shibukawa/xlsxformula"
	"github.com/shibukawa/xlsxrange"
	"github.com/tealeg/xlsx"
	"go/format"
	"io/ioutil"
	"log"
	"path"
	"strings"
)

func WriteGoFunction(outputDir, packageName string, function *Function, constants map[string]*xlsx.Cell) {
	filePath := path.Join(outputDir, fmt.Sprintf("%s.go", strings.ToLower(function.Name)))
	fmt.Printf("writing %s\n", filePath)
	var buffer bytes.Buffer
	fmt.Fprintf(&buffer, "func %s(", GoFuncName(function.Name))
	for i, param := range function.Parameters {
		if i != 0 {
			buffer.WriteString(", ")
		}
		fmt.Fprintf(&buffer, "%s float64", param.Name)
	}
	buffer.WriteString(") float64 {\nreturn ")
	node, err := xlsxformula.Parse(function.Formula)
	if err != nil {
		log.Fatalln(err)
	}
	imports := make(map[string]bool)
	WriteNode(&buffer, node, function, constants, imports, false)
	buffer.WriteString("\n}\n")

	var buffer2 bytes.Buffer

	fmt.Fprintf(&buffer2, "package %s\n", packageName)
	if len(imports) > 0 {
		buffer2.WriteString("import (")
		for _, importName := range imports {
			fmt.Fprintf(&buffer2, "\"%s\"\n", importName)
		}
		buffer2.WriteString(")\n")
	}
	buffer2.Write(buffer.Bytes())

	formattedCode, err := format.Source(buffer2.Bytes())
	if err != nil {
		log.Fatalln(err)
	}
	ioutil.WriteFile(filePath, formattedCode, 0777)
}

func GoFuncName(name string) string {
	runes := []rune(name)
	return strings.ToUpper(string(runes[0:1])) + string(runes[1:])
}

func WriteNode(buffer *bytes.Buffer, node *xlsxformula.Node, function *Function, constants map[string]*xlsx.Cell, imports map[string]bool, insideExpression bool) error {
	switch node.Type {
	case xlsxformula.Function:
		buffer.WriteString(node.Token.Text)
		buffer.WriteByte('(')
		for i, child := range node.Children {
			if i != 0 {
				buffer.WriteString(", ")
			}
			WriteNode(buffer, child, function, constants, imports, false)
		}
		buffer.WriteByte(')')
	case xlsxformula.Expression:
		if insideExpression {
			buffer.WriteByte('(')
		}
		for i, child := range node.Children {
			if i != 0 {
				buffer.WriteByte(' ')
			}
			WriteNode(buffer, child, function, constants, imports, true)
		}
		if insideExpression {
			buffer.WriteByte(')')
		}
	case xlsxformula.SingleToken:
		token := node.Token.Text
		switch node.Token.Type {
		case xlsxformula.Number:
			buffer.WriteString(token)
		case xlsxformula.String:
			byteExpression, _ := json.Marshal(token)
			buffer.Write(byteExpression)
		case xlsxformula.Bool:
			if token == "TRUE" {
				buffer.WriteString("true")
			} else {
				buffer.WriteString("false")
			}
		case xlsxformula.Operator:
			// todo: ^, %, & operator
			buffer.WriteString(token)
		case xlsxformula.Comparator:
			if token == "=" {
				buffer.WriteString("==")
			} else if token == "<>" {
				buffer.WriteString("!=")
			} else {
				buffer.WriteString(token)
			}
		case xlsxformula.Name:
			cell, ok := constants[token]
			if ok {
				buffer.WriteString(cell.Value)
			} else {
				return fmt.Errorf("Unknown name '%s' at %d:%d", token, node.Token.Line, node.Token.Col)
			}
		case xlsxformula.Range:
			variable, ok := function.ParameterFromRange[token]
			if ok {
				buffer.WriteString(variable.Name)
			} else {
				aRange := xlsxrange.New(function.FormulaRange.Sheet, token)
				buffer.WriteString(aRange.GetCell().Value)
			}
		}
	}
	return nil
}
