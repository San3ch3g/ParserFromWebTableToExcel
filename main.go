package main

import (
	"fmt"
	"log"
	"strings"

	"github.com/gocolly/colly"
	"github.com/xuri/excelize/v2"
)

func main() {
	url := "https://attack.mitre.org/matrices/enterprise/"

	c := colly.NewCollector()

	file := excelize.NewFile()

	sheetMap := make(map[string]bool)

	c.OnHTML(".matrix-container.p-3 a", func(e *colly.HTMLElement) {
		link := e.Attr("href")
		if strings.HasPrefix(link, "/tactics/TA") {
			fmt.Println(link)

			linkCollector := colly.NewCollector()

			linkCollector.OnHTML("table.table-techniques tbody tr.technique", func(innerE *colly.HTMLElement) {
				techniqueID := innerE.ChildText("td:nth-child(1) a")
				techniqueName := innerE.ChildText("td:nth-child(2) a")

				if techniqueID != "" && techniqueName != "" {
					fmt.Printf("%s  %s\n", techniqueID, techniqueName)

					sheetName := strings.TrimPrefix(link, "/tactics/")
					exists := sheetMap[sheetName]

					if !exists {
						file.NewSheet(sheetName)
						sheetMap[sheetName] = true
					}

					rows, err := file.GetRows(sheetName)
					if err != nil {
						log.Println("Error getting rows:", err)
					}
					rowIndex := len(rows) + 1

					file.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), techniqueID)
					file.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), techniqueName)
				}
			})

			if err := linkCollector.Visit(e.Request.AbsoluteURL(link)); err != nil {
				log.Println("Error visiting link:", err)
			}
		}
	})

	err := c.Visit(url)
	if err != nil {
		log.Fatal(err)
	}

	if err := file.SaveAs("output.xlsx"); err != nil {
		log.Fatal(err)
	}
}
