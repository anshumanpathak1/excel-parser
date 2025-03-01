package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	EmplID string
	Marks  float64
}

func main() {
	if len(os.Args) < 2 {
		log.Fatal("Usage: go run main.go <path_to_excel>")
	}

	excelPath := os.Args[1]
	file, err := excelize.OpenFile(excelPath)
	if err != nil {
		log.Fatalf("Error reading file: %v", err)
	}
	defer file.Close()

	sheetName := file.GetSheetName(0)
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Error reading sheet: %v", err)
	}

	branchTotals := make(map[string][]float64)
	componentSums := make(map[string]float64)
	componentCounts := make(map[string]int)
	var totalScores []float64
	componentRanks := map[string][]Student{}

	columns := map[string]int{
		"EMPLID": 1, "CAMPUS_ID": 0, "QUIZ": 2, "MIDSEM": 3, "LABTEST": 4, "WEEKLYLABS": 5, "COMPRE": 6, "TOTAL": 7,
	}

	validBranchCodes := []string{"A3", "A4", "A5", "A7", "A8", "AA", "AD"}

	for _, row := range rows[1:] {
		if len(row) < 8 {
			continue
		}

		emplID := row[columns["EMPLID"]]
		campusID := row[columns["CAMPUS_ID"]]
		if !strings.HasPrefix(campusID, "2024") {
			continue
		}

		branch := ""
		for _, code := range validBranchCodes {
			if strings.Contains(campusID, code) {
				branch = code
				break
			}
		}
		if branch == "" {
			continue
		}

		scores := make(map[string]float64)
		for key, index := range columns {
			if key == "EMPLID" || key == "CAMPUS_ID" {
				continue
			}
			score, err := strconv.ParseFloat(row[index], 64)
			if err != nil {
				continue
			}
			scores[key] = score
		}

		branchTotals[branch] = append(branchTotals[branch], scores["TOTAL"])
		totalScores = append(totalScores, scores["TOTAL"])

		for key, score := range scores {
			componentSums[key] += score
			componentCounts[key]++
			componentRanks[key] = append(componentRanks[key], Student{EmplID: emplID, Marks: score})
		}
	}

	fmt.Println("General Averages:")
	for comp, sum := range componentSums {
		avg := sum / float64(componentCounts[comp])
		fmt.Printf("%s: %.2f\n", comp, avg)
	}

	fmt.Println("\nBranch-wise Averages (2024 Batch):")
	for branch, scores := range branchTotals {
		sum := 0.0
		for _, score := range scores {
			sum += score
		}
		avg := sum / float64(len(scores))
		fmt.Printf("Branch %s: %.2f\n", branch, avg)
	}

	fmt.Println("\nTop 3 Students Per Component:")
	for comp, students := range componentRanks {
		sort.Slice(students, func(i, j int) bool {
			return students[i].Marks > students[j].Marks
		})

		fmt.Printf("\n%s Rankings:\n", comp)
		for i := 0; i < 3 && i < len(students); i++ {
			rankStr := ""
			switch i {
			case 0:
				rankStr = "1st"
			case 1:
				rankStr = "2nd"
			case 2:
				rankStr = "3rd"
			}
			fmt.Printf("%s: EmplID %s - Marks: %.2f\n", rankStr, students[i].EmplID, students[i].Marks)
		}
	}
}

