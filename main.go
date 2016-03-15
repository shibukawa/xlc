package main

import (
	"github.com/tealeg/xlsx"
	"log"
)

func main() {
	config := initConfig()

	file, err := xlsx.OpenFile(config.SourcePath)
	if err != nil {
		log.Fatalln(err)
	}
	schema, err := ParseSchema(file)
	if err != nil {
		log.Fatalln(err)
	}
	constants := ParseConstant(file)

	for _, function := range schema.Functions {
		WriteGoFunction(config.OutputPath, config.Package, function, constants)
	}
}
