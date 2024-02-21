// See https://aka.ms/new-console-template for more information
using System.IO;
using CsvHelper;
using System.Collections.Generic;
using System.Globalization;
using Aspose.CAD;
using System.Text;
using System;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

public class ExportToExcel
{
    public void Run(FileNameReturnObject fileNameReturnObject, List<CalculationEntry> calculationEntries)
    {
        string csvFileName = fileNameReturnObject.FullFilePath.Replace(".dwg",".csv");
        
        using (var writer = new StreamWriter(csvFileName))

        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
        {
            csv.WriteRecords(calculationEntries);
        }
        convertToExcel(csvFileName, fileNameReturnObject.FileName, ",");

    }
    public void convertToExcel(string csvFilePath,string fileName, string splitter)
    {
        string excelDirectory = @"C:\\Users\\Michal\\Desktop\\Michis AutoCAD Kostenrechnungen\";
        //string excelDirectory = @"C:\\temp\\";

        var excelFilePath = excelDirectory + fileName;
        string newFileName = excelFilePath.Replace(".dwg", ".xls");

        string[] lines = File.ReadAllLines(csvFilePath, Encoding.UTF8);

        int columnCounter = 0;

        foreach (string s in lines)
        {
            string[] ss = s.Trim().Split(Convert.ToChar(splitter));

            if (ss.Length > columnCounter)
                columnCounter = ss.Length;
        }

        HSSFWorkbook workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var rowIndex = 0;
        var rowExcel = sheet.CreateRow(rowIndex);

        for (int lineNumber = 0; lineNumber < lines.Length;lineNumber++)
        {
            var s = lines[lineNumber];
            rowExcel = sheet.CreateRow(rowIndex);

            string[] ss = s.Trim().Split(Convert.ToChar(splitter));

            for (int i = 0; i < columnCounter; i++)
            {
                var headerName = ss[i];
                string data = !String.IsNullOrEmpty("s") && i < ss.Length ? ss[i] : "";

                if (lineNumber <= 0)
                {
                    rowExcel.CreateCell(i).SetCellType(CellType.String);
                    rowExcel.CreateCell(i).SetCellValue(data);
                }
                else {
                    if (headerName.StartsWith("=")) {
                        rowExcel.CreateCell(i).SetCellType(CellType.Formula);
                        rowExcel.CreateCell(i).SetCellFormula(data.Remove(0, 1)); //Remove the = from the formular
                    }
                    else if (headerName.Contains(".")){
                        rowExcel.CreateCell(i).SetCellType(CellType.Numeric);
                        DecimalFormat df = new DecimalFormat("#,##0.00");
                        var formattedValue = df.Format(data);
                        rowExcel.CreateCell(i).SetCellValue(formattedValue);
                    }
                    else
                    {
                        rowExcel.CreateCell(i).SetCellType(CellType.Numeric);
                        rowExcel.CreateCell(i).SetCellValue(data);
                    }
                }
              
            }

            rowIndex++;
        }
        for (var i = 0; i < sheet.GetRow(0).LastCellNum; i++)
            sheet.AutoSizeColumn(i);

        //Delete Excel File if existent

        using (FileStream file = new FileStream(newFileName, FileMode.OpenOrCreate, FileAccess.Write))
        {
            workbook.Write(file);
            file.Close();
        }

        Console.WriteLine($"Neues Excel wurde erfolgreich erstellt: {newFileName}");
    }

}


