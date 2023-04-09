package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"time"

	"google.golang.org/api/sheets/v4"
)

type yearConfig struct {
	SpreadsheetId     string
	WriteRange        string
	LinkedSpreadsheet string
	SheetName         string
	Year              int
	RowsNumber        int
}

func addYear(srv *sheets.Service, yearCFG *yearConfig) (*sheets.UpdateValuesResponse, error) {
	var vr sheets.ValueRange

	// build heading with months from 1/yearCFG.Year to 12/yearCFG.Year
	heading := []interface{}{}
	date := time.Date(yearCFG.Year, time.January, 1, 0, 0, 0, 0, time.UTC)
	for i := 1; i <= 12; i++ {
		heading = append(heading, fmt.Sprintf("%d/%v", date.Month(), date.Year()))
		date = date.AddDate(0, 1, 0)
	}
	// add blank row
	vr.Values = append(vr.Values, heading, make([]interface{}, 12))

	// build rows
	for i := 1; i <= yearCFG.RowsNumber; i++ {
		row := []interface{}{}
		for letter := int('B'); letter <= int('M'); letter++ {
			// change letter for every column
			monthString := fmt.Sprintf("=IMPORTRANGE(\"%s\"; \"%v!%c%v\")", yearCFG.LinkedSpreadsheet, yearCFG.SheetName, letter, i)
			row = append(row, monthString)
		}
		vr.Values = append(vr.Values, row)
	}

	// send rows to spreadsheet
	return srv.Spreadsheets.Values.Update(yearCFG.SpreadsheetId, yearCFG.WriteRange, &vr).ValueInputOption("USER_ENTERED").Do()
}

func main() {
	srv := authenticate()

	var yearCFG yearConfig
	jsonFile, err := os.ReadFile("yearConfig.json")
	if err != nil {
		log.Fatalf("Error reading yearConfig file: %v", err)
	}

	json.Unmarshal(jsonFile, &yearCFG)

	_, err = addYear(srv, &yearCFG)
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}
}
