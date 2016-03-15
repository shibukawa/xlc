package main

import (
	"gopkg.in/alecthomas/kingpin.v2"
)

type Config struct {
	SourcePath string
	OutputPath string
	Package    string
}

func initConfig() *Config {
	source := kingpin.Arg("source", "Input .xlsx file").Required().ExistingFile()
	output := kingpin.Arg("output", "Output path").Required().ExistingDir()
	packageName := kingpin.Flag("package", "Go package").Short('f').Default("main").String()

	kingpin.Parse()

	return &Config{
		SourcePath: *source,
		OutputPath: *output,
		Package:    *packageName,
	}
}
