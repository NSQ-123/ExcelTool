package main

import (
    "encoding/csv"
    "fmt"
    "os"
    "path/filepath"
    "strings"

    "github.com/xuri/excelize/v2"
)

func main() {
    if len(os.Args) < 4 {
        fmt.Println("Usage: exceltool <inputExcelDir> <outputCsvDir> <outputCsharpDir>")
        return
    }
    inputDir := os.Args[1]
    outputCsvDir := os.Args[2]
    outputCsharpDir := os.Args[3]

    // 保证输出目录存在
    os.MkdirAll(outputCsvDir, 0755)
    os.MkdirAll(outputCsharpDir, 0755)

    // 遍历 excel 目录下所有 .xlsx 文件
    files, _ := os.ReadDir(inputDir)
    for _, file := range files {
        if file.IsDir() || !strings.HasSuffix(file.Name(), ".xlsx") {
            continue
        }
        excelPath := filepath.Join(inputDir, file.Name())
        baseName := strings.TrimSuffix(file.Name(), ".xlsx")
        csvPath := filepath.Join(outputCsvDir, baseName+".csv")
        csPath := filepath.Join(outputCsharpDir, "T_"+baseName+".cs")

        ExportCsv(excelPath, csvPath)
        ExportCSharp(excelPath, csPath, baseName)
        fmt.Println("Processed:", file.Name())
    }
}

// 导出CSV（取第一个sheet全部内容）
func ExportCsv(excelPath, csvPath string) {
    f, err := excelize.OpenFile(excelPath)
    if err != nil {
        fmt.Println("open excel error:", err)
        return
    }
    defer f.Close()
	sheet := f.GetSheetName(0)
    rows, err := f.GetRows(sheet)
    if err != nil {
        fmt.Println("read sheet error:", err)
        return
    }
    out, _ := os.Create(csvPath)
    defer out.Close()
    writer := csv.NewWriter(out)
    defer writer.Flush()
    for _, row := range rows {
        writer.Write(row)
    }
}

// 导出C#结构体（根据第二个sheet的前4行生成字段和注释）
func ExportCSharp(excelPath, csPath, baseName string) {
    f, err := excelize.OpenFile(excelPath)
    if err != nil {
        fmt.Println("open excel error:", err)
        return
    }
    defer f.Close()
    sheet := f.GetSheetName(1)
    rows, err := f.GetRows(sheet)
    if err != nil || len(rows) < 4 {
        fmt.Println("meta sheet error:", err)
        return
    }
    fieldNames := rows[0]
    fieldTypes := rows[1]
    fieldUsages := rows[2]
    fieldDescs := rows[3]

    sb := &strings.Builder{}
    sb.WriteString("using System;\nusing System.Collections.Generic;\n\n")
    sb.WriteString("namespace GameFramework.Table {\n")
    sb.WriteString(fmt.Sprintf("\tpublic partial class T_%s : ITable {\n", baseName))
    for i := range fieldNames {
        if i >= len(fieldUsages) || !strings.Contains(strings.ToLower(fieldUsages[i]), "c") {
            continue // 只导出客户端字段
        }
        if i < len(fieldDescs) && fieldDescs[i] != "" {
            sb.WriteString(fmt.Sprintf("\t\t/// <summary>\n\t\t/// %s\n\t\t/// </summary>\n", fieldDescs[i]))
        }
        goType := GoTypeToCSharp(fieldTypes[i])
        fieldName := fieldNames[i]
        if fieldName == "" {
            continue
        }
        fieldName = strings.Title(fieldName)
        sb.WriteString(fmt.Sprintf("\t\tpublic %s %s { get; set; }\n", goType, fieldName))
    }
    sb.WriteString("\t}\n}\n")
    os.WriteFile(csPath, []byte(sb.String()), 0644)
}

// Go类型到C#类型简单映射（可根据你的工程扩展）
func GoTypeToCSharp(t string) string {
    switch strings.ToLower(t) {
    case "int":
        return "int"
    case "string":
        return "string"
    case "float":
        return "float"
    case "double":
        return "double"
    case "bool":
        return "bool"
    case "intslice":
        return "List<int>"
    case "stringslice":
        return "List<string>"
    case "floatslice":
        return "List<float>"
    case "doubleslice":
        return "List<double>"
    case "boolslice":
        return "List<bool>"
    default:
        return "string"
    }
}