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
using NPOI.POIFS.Crypt.Dsig;

public class ExportToExcel
{
    public void Run(FileNameReturnObject fileNameReturnObject, List<CalculationEntry> calculationEntries)
    {
        var csvFileName = convertToCsv(fileNameReturnObject, calculationEntries);
        convertToExcel(csvFileName, fileNameReturnObject.FileName, ",", calculationEntries);

    }
    public string convertToCsv(FileNameReturnObject fileNameReturnObject, List<CalculationEntry> calculationEntries)
    {
        string csvFileName = fileNameReturnObject.FullFilePath.Replace(".dwg", ".csv");

        using (var writer = new StreamWriter(csvFileName))

        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
        {
            csv.WriteRecords(calculationEntries);
        }
        return csvFileName;
    }
    public void convertToExcel(string csvFilePath, string fileName, string splitter, List<CalculationEntry> calculationEntries)
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
        CreateKostenSheet(splitter, lines, columnCounter, workbook);
        CreateAbmessungenSheet(workbook, calculationEntries);

        //Delete Excel File if existent

        using (FileStream file = new FileStream(newFileName, FileMode.OpenOrCreate, FileAccess.Write))
        {
            workbook.Write(file);
            file.Close();
        }

        Console.WriteLine($"Neues Excel wurde erfolgreich erstellt: {newFileName}");
    }

    private void CreateKostenSheet(string splitter, string[] lines, int columnCounter, HSSFWorkbook workbook)
    {
        var sheetKosten = workbook.CreateSheet("Kosten");
        var rowIndex = 0;
        var rowExcel = sheetKosten.CreateRow(rowIndex);

        for (int lineNumber = 0; lineNumber < lines.Length; lineNumber++)
        {
            var s = lines[lineNumber];
            rowExcel = sheetKosten.CreateRow(rowIndex);

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
                else
                {
                    if (headerName.StartsWith("="))
                    {
                        rowExcel.CreateCell(i).SetCellType(CellType.Formula);
                        rowExcel.CreateCell(i).SetCellFormula(data.Remove(0, 1)); //Remove the = from the formular
                    }
                    else if (headerName.Contains("."))
                    {
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
        for (var i = 0; i < sheetKosten.GetRow(0).LastCellNum; i++)
            sheetKosten.AutoSizeColumn(i);
    }

    private void CreateAbmessungenSheet(HSSFWorkbook workbook, List<CalculationEntry> calculationEntries)
    {
        var sheetAbmessungen = workbook.CreateSheet("Abmessungen");       

        for (int i = 0; i < calculationEntries.Count; i++)
        {
            var rowExcel = sheetAbmessungen.CreateRow(i);

            var calculationEntry = calculationEntries[i];

            rowExcel.CreateCell(0).SetCellType(CellType.Numeric);
            rowExcel.CreateCell(0).SetCellValue(calculationEntry.Thickness);

            for (int j = 0; j < calculationEntry.Abmessungen.Count; j++)
            {
                var abmessung = calculationEntry.Abmessungen[j];
                rowExcel.CreateCell(j+1).SetCellType(CellType.Numeric);
                DecimalFormat df = new DecimalFormat("#,##0.00");
                var formattedValue = df.Format(abmessung);
                rowExcel.CreateCell(j+1).SetCellValue(formattedValue);
            }


        }
        
        for (var i = 0; i < sheetAbmessungen.GetRow(0).LastCellNum; i++)
            sheetAbmessungen.AutoSizeColumn(i);
    }
}


