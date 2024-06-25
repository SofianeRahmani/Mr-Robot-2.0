using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
namespace RecordETL.Services;

public class ExcelService
{
    public static List<string> ReadSheetNames(ExcelWorkbook workbook)
    {
        List<string> sheetNames = new List<string>();
        foreach (var worksheet in workbook.Worksheets)
        {
            sheetNames.Add(worksheet.Name);
        }
        return sheetNames;
    }
    public static List<string> ReadColumnNames(ExcelWorkbook workbook, string sheetName)
    {
        List<string> columnNames = new List<string>();
        var worksheet = workbook.Worksheets[sheetName]; 

        if (worksheet == null) return columnNames; 
        if (worksheet.Dimension == null) return columnNames; 

            
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            if (worksheet.Cells[2, col].Value != null) 
                columnNames.Add(worksheet.Cells[2, col].Text); 
        }

        return columnNames;
    }
        
    public static string? GetColumnValue(int row, int column, ExcelWorksheet worksheet)
    {
        return column != -1 ? worksheet.Cells[row, column + 1].Text.Trim() : null;
    }


}