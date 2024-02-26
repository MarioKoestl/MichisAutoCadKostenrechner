// See https://aka.ms/new-console-template for more information
using System.IO;
using CsvHelper;
using System.Collections.Generic;
using System;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using CsvHelper.Configuration.Attributes;
using System.Reflection;
using System.Linq;
using NPOI.XWPF.UserModel;

public class ExportToExcel
{
    public void ConvertToExcel(FileNameReturnObject fileNameReturnObject, List<CalculationEntry> calculationEntries)
    {
        //string excelDirectory = @"C:\\Users\\Michal\\Desktop\\Michis AutoCAD Kostenrechnungen\";
        string excelDirectory = @"C:\\temp\\";

        var excelFilePath = excelDirectory + fileNameReturnObject.FileName;
        string newFileName = excelFilePath.Replace(".dwg", ".xls");

     
        HSSFWorkbook workbook = new HSSFWorkbook();
        CreateKostenSheet(workbook, calculationEntries);
        CreateAbmessungenSheet(workbook, calculationEntries);

        //Delete Excel File if existent

        using (FileStream file = new FileStream(newFileName, FileMode.OpenOrCreate, FileAccess.Write))
        {
            workbook.Write(file);
            file.Close();
        }

        Console.WriteLine($"Neues Excel wurde erfolgreich erstellt: {newFileName}");
    }
    private void CreateKostenSheet(HSSFWorkbook workbook, List<CalculationEntry> calculationEntries)
    {
        var sheetKosten = workbook.CreateSheet("Kosten");
        var rowExcel = sheetKosten.CreateRow(0);

        //Create Header row
        int columnCounter = 0;
        var calculationEntryType= typeof(CalculationEntry);
        foreach (PropertyInfo property in calculationEntryType.GetProperties())
        {
            
            var displayNameAttribute = property.GetCustomAttribute<NameAttribute>();
            string displayName = displayNameAttribute?.Names.FirstOrDefault() ?? property.Name;

            //Abmessungne will be set in new Sheet
            if (displayName != "Abmessungen")
            {
                rowExcel.CreateCell(columnCounter).SetCellType(CellType.String);
                rowExcel.CreateCell(columnCounter).SetCellValue(displayName);
            }
            columnCounter++;
        }

        for (int rowIndex = 0; rowIndex < calculationEntries.Count; rowIndex++)
        {
            rowExcel = sheetKosten.CreateRow(rowIndex+1);
            var calculationEntry = calculationEntries[rowIndex];
            columnCounter = 0;
            foreach (var property in calculationEntryType.GetProperties())
            {
                var propertyValue = property.GetValue(calculationEntry);
                if (propertyValue != null)
                {
                    if (property.PropertyType == typeof(string))
                    {
                        var data = propertyValue.ToString();
                        if (data != null && data.StartsWith("="))
                        {
                            rowExcel.CreateCell(columnCounter).SetCellType(CellType.Formula);
                            rowExcel.CreateCell(columnCounter).SetCellFormula(data.Remove(0, 1)); //Remove the = from the formular
                        }
                        else
                        {
                            rowExcel.CreateCell(columnCounter).SetCellType(CellType.String);
                            rowExcel.CreateCell(columnCounter).SetCellFormula(data);
                        }
                    }
                    if (property.PropertyType == typeof(double))
                    {
                        rowExcel.CreateCell(columnCounter).SetCellType(CellType.Numeric);
                        DecimalFormat df = new DecimalFormat("#,##0.00");
                        var formattedValue = df.Format(propertyValue);
                        rowExcel.CreateCell(columnCounter).SetCellValue(formattedValue);
                    }
                    if (property.PropertyType == typeof(int))
                    {
                        //propertyValue=Convert.ToDecimal((int)propertyValue);
                        rowExcel.CreateCell(columnCounter).SetCellType(CellType.Numeric);
                        DecimalFormat df = new DecimalFormat("#,##0.0");
                        var formattedValue = df.Format(propertyValue);
                        rowExcel.CreateCell(columnCounter).SetCellValue(formattedValue);

                    }
                }
                columnCounter++;
            }
        }
        //Create GesamtKosten Row
        rowExcel = sheetKosten.CreateRow(calculationEntries.Count+1);
        rowExcel.CreateCell(9).SetCellType(CellType.Formula);
        rowExcel.CreateCell(9).SetCellFormula("SUM(I:I)");
      
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


