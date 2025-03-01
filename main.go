package main

import (
    "encoding/json"
    "fmt"
    "log"
    "os"
    "strconv"

    "github.com/xuri/excelize/v2"
)

type Student struct {
    Emplid        string  `json:"emplid"`
    ClassNo       string  `json:"class_no"`
    Quiz          float64 `json:"quiz"`
    MidSem        float64 `json:"mid_sem"`
    LabTest       float64 `json:"lab_test"`
    WeeklyLabs    float64 `json:"weekly_labs"`
    PreCompre     float64 `json:"pre_compre"`
    Compre        float64 `json:"compre"`
    Total         float64 `json:"total"`
    ComputedTotal float64 `json:"computed_total"`
    Error         string  `json:"error,omitempty"`
}

func parseExcel(filePath string) ([]Student, error) {
    f, err := excelize.OpenFile(filePath)
    if err != nil {
        return nil, err
    }
    defer f.Close()

    rows, err := f.GetRows("CSF111_202425_01_GradeBook")
    if err != nil {
        return nil, err
    }

    var students []Student
    for i, row := range rows {
        if i == 0 || len(row) < 11 {
            continue
        }

        quiz, _ := strconv.ParseFloat(row[4], 64)
        midSem, _ := strconv.ParseFloat(row[5], 64)
        labTest, _ := strconv.ParseFloat(row[6], 64)
        weeklyLabs, _ := strconv.ParseFloat(row[7], 64)
        preCompre, _ := strconv.ParseFloat(row[8], 64)
        compre, _ := strconv.ParseFloat(row[9], 64)
        total, _ := strconv.ParseFloat(row[10], 64)

        computedTotal := quiz + midSem + labTest + weeklyLabs + compre
        var errorMsg string
        if computedTotal != total {
            errorMsg = fmt.Sprintf("Discrepancy: Expected %.2f, Found %.2f", total, computedTotal)
        }

        students = append(students, Student{
            Emplid:        row[2],
            ClassNo:       row[1],
            Quiz:          quiz,
            MidSem:        midSem,
            LabTest:       labTest,
            WeeklyLabs:    weeklyLabs,
            PreCompre:     preCompre,
            Compre:        compre,
            Total:         total,
            ComputedTotal: computedTotal,
            Error:         errorMsg,
        })
    }
    return students, nil
}

func reportResults(students []Student) {
    for _, s := range students {
        if s.Error != "" {
            fmt.Printf("Error for %s: %s\n", s.Emplid, s.Error)
        }
    }

    data, _ := json.MarshalIndent(students, "", "  ")
    _ = os.WriteFile("report.json", data, 0644)
    fmt.Println("Report saved to report.json")
}

func main() {
    if len(os.Args) < 2 {
        log.Fatal("Usage: go run main.go <path_to_excel>")
    }

    filePath := os.Args[1]
    students, err := parseExcel(filePath)
    if err != nil {
        log.Fatal("Error reading file:", err)
    }

    reportResults(students)
}
