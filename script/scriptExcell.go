package main

import (
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)
const (
	count = 2
)
func searchIfExist(sourceFile *xlsx.File, text string)(bool) {
	for _, sheet := range sourceFile.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				data := cell.String()
				if data == text {
					return true
				}
			}
		}
	}
	return false	
}
func main() {
	// Open the source Excel file
	sourceFilePath := "CNF-Report.xlsx"
	sourceFile, err := xlsx.OpenFile(sourceFilePath)
	if err != nil {
		log.Fatalf("Error opening source Excel file: %s", err)
	}
	// Open the dest Excel file
	destFilePath := "Telco-Operator-List.xlsx"
	destFile, err := xlsx.OpenFile(destFilePath)
	if err != nil {
		log.Fatalf("Error opening source Excel file: %s", err)
	}

	for _, test := range destFile.Sheets {
		for _, row := range test.Rows {
			counter := 0
			for _, cell := range row.Cells {
				text := cell.String()
				if text == "" {
					break
				}
				fmt.Println(text)
				counter++
				if searchIfExist(sourceFile,text) {
					//put YES in column C in dest file
					cell.SetString("Yes")
				} else {
					cell.SetString("No")
				}
				if counter == count {
					break
				}
			}
		}
	}
	// Save the modified Excel file
	err = destFile.Save(destFilePath)
	if err != nil {
		fmt.Println(err)
		return
	}
			
}