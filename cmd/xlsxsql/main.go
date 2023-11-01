package main

import "github.com/noborus/xlsxsql/cmd"

// version represents the version
var version = "dev"

// revision set "git rev-parse --short HEAD"
var revision = "HEAD"

func main() {
	cmd.Execute(version, revision)
}
