package main

import (
    "encoding/json"
    "fmt"
    "log"
    "os"
    "strconv"
    "strings"
    "github.com/xuri/excelize/v2"
    "math"
)

type Student struct {
    ClassNo       string  `json:"class_no"`
    Emplid        string  `json:"emplid"`
    CampusID      string  `json:"campus_id"`
    Quiz          float64 `json:"quiz"`
    MidSem        float64 `json:"mid_sem"`
    LabTest       float64 `json:"lab_test"`
    WeeklyLabs    float64 `json:"weekly_labs"`
    PreCompre     float64 `json:"pre_compre"`
    Compre        float64 `json:"compre"`
    Total         float64 `json:"total"`
    ComputedTotal float64 `json:"computed_total"`
    BranchCode    string  `json:"branch_code"`
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
        const epsilon = 0.01
        if math.Abs(computedTotal-total) > epsilon {
            errorMsg = fmt.Sprintf("Discrepancy: Expected %.2f, Found %.2f", total, computedTotal)
        }

        branchCode := ""
        if len(row) > 3 && strings.HasPrefix(row[3], "2024") && len(row[3]) > 6 {
            branchCode = row[3][4:6]
        }

        students = append(students, Student{
            ClassNo:       row[1],
            Emplid:        row[2],
            CampusID:      row[3],
            Quiz:          quiz,
            MidSem:        midSem,
            LabTest:       labTest,
            WeeklyLabs:    weeklyLabs,
            PreCompre:     preCompre,
            Compre:        compre,
            Total:         total,
            ComputedTotal: computedTotal,
            BranchCode:    branchCode,
            Error:         errorMsg,
        })
    }
    return students, nil
}

func computeAverages(students []Student) map[string]float64 {
    averages := make(map[string]float64)
    counts := make(map[string]int)
    fields := []string{"quiz", "mid_sem", "lab_test", "weekly_labs", "pre_compre", "compre", "total"}

    for _, student := range students {
        for _, field := range fields {
            var score float64
            switch field {
            case "quiz":
                score = student.Quiz
            case "mid_sem":
                score = student.MidSem
            case "lab_test":
                score = student.LabTest
            case "weekly_labs":
                score = student.WeeklyLabs
            case "pre_compre":
                score = student.PreCompre
            case "compre":
                score = student.Compre
            case "total":
                score = student.Total
            }
            averages[field] += score
            counts[field]++
        }
    }

    for field := range averages {
        if counts[field] > 0 {
            averages[field] /= float64(counts[field])
        }
    }
    return averages
}

func computeBranchAverages(students []Student) map[string]float64 {
    branchSums := make(map[string]float64)
    branchCounts := make(map[string]int)

    for _, student := range students {
        if student.BranchCode != "" {
            branchSums[student.BranchCode] += student.Total
            branchCounts[student.BranchCode]++
        }
    }

    branchAverages := make(map[string]float64)
    for branch, sum := range branchSums {
        count := branchCounts[branch]
        if count > 0 {
            branchAverages[branch] = sum / float64(count)
        }
    }

    return branchAverages
}

func findTopStudents(students []Student, field string) []Student {
    var sortedStudents []Student
    switch field {
    case "quiz":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.Quiz })
    case "mid_sem":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.MidSem })
    case "lab_test":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.LabTest })
    case "weekly_labs":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.WeeklyLabs })
    case "pre_compre":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.PreCompre })
    case "compre":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.Compre })
    case "total":
        sortedStudents = sortByScore(students, func(s Student) float64 { return s.Total })
    }

    if len(sortedStudents) > 3 {
        return sortedStudents[:3]
    }
    return sortedStudents
}

func sortByScore(students []Student, getScore func(Student) float64) []Student {
    sortedStudents := make([]Student, len(students))
    copy(sortedStudents, students)
    
    for i := 0; i < len(sortedStudents)-1; i++ {
        for j := i + 1; j < len(sortedStudents); j++ {
            if getScore(sortedStudents[i]) < getScore(sortedStudents[j]) {
                sortedStudents[i], sortedStudents[j] = sortedStudents[j], sortedStudents[i]
            }
        }
    }
    return sortedStudents
}

func reportResults(students []Student) {
    for _, s := range students {
        if s.Error != "" {
            fmt.Printf("Error for %s: %s\n", s.Emplid, s.Error)
        }
    }
    fmt.Println("\n")

    averages := computeAverages(students)
    fmt.Println("Averages:\n")
    for field, avg := range averages {
        fmt.Printf("%s: %.2f\n", field, avg)
    }
    fmt.Println("\n")

    topStudents := make(map[string][]Student)
    fields := []string{"quiz", "mid_sem", "lab_test", "weekly_labs", "pre_compre", "compre", "total"}
    for _, field := range fields {
        topStudents[field] = findTopStudents(students, field)
        fmt.Printf("Top 3 students for %s:\n", field)
        for rank, student := range topStudents[field] {
            var marks float64
            switch field {
            case "quiz":
                marks = student.Quiz
            case "mid_sem":
                marks = student.MidSem
            case "lab_test":
                marks = student.LabTest
            case "weekly_labs":
                marks = student.WeeklyLabs
            case "pre_compre":
                marks = student.PreCompre
            case "compre":
                marks = student.Compre
            case "total":
                marks = student.Total
            }
            fmt.Printf("%d: %s, Marks: %.2f\n", rank+1, student.Emplid, marks)
        }
    }
    fmt.Println("\n")
    branchAverages := computeBranchAverages(students)
    fmt.Println("Branch Averages (2024 entries):\n")
    for branch, avg := range branchAverages {
        fmt.Printf("Branch %s: %.2f\n", branch, avg)
    }
    fmt.Println("\n")

    results := map[string]interface{}{
        "top_students": topStudents,
        "averages":     averages,
        "branch_averages": branchAverages,
        "students": students,
    }

    data, _ := json.MarshalIndent(results, "", "  ")
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
