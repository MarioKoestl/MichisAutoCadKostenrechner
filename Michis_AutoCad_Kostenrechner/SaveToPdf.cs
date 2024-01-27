// See https://aka.ms/new-console-template for more information
using Aspose.CAD;
using Aspose.CAD.Xmp.Types.Complex.Dimensions;
using System;
using System.ComponentModel.DataAnnotations;

//        //var drawing = Aspose.CAD.Project.LoadFrom(new File("input.dwg"));
//        //// Extract dimension and area data
//        //var dimensionDataList = new List<DimensionData>();
//        //var areaDataList = new List<AreaData>();

//        //foreach (var object in drawing.Objects)
//        //{
//        //    if (object.GetType() == typeof(Dimension))
//        //    {
//        //        var dimension = (Dimension)object;
//        //        dimensionDataList.Add(new DimensionData
//        //        {
//        //            ObjectType = "Dimension",
//        //            DimensionValue = dimension.Value,
//        //            DimensionType = dimension.Type,
//        //        });
//        //    }
//        //    else if (object.GetType() == typeof(Polyline))
//        //    {
//        //        var polyline = (Polyline)object;
//        //        areaDataList.Add(new AreaData
//        //        {
//        //            ObjectType = "Polygon",
//        //            AreaValue = polyline.Area,
//        //        });
//        //    }
//        //}

//        // Convert data to Excel-compatible format
//        //var tableData = GetExcelTableData(dimensionDataList, areaDataList);

//        // Save data to Excel file
//        //var workbook = Aspose.CAD.Project.CreateWorkbook();
//        //var sheet = workbook.Worksheets.Add("Data");
//        //var table = sheet.Tables.Add(tableData);
//        //workbook.SaveTo(new File("output.xlsx"));
//    }

//    //public TableData GetExcelTableData(List<DimensionData> dimensionData, List<AreaData> areaData)
//    //{
//    //    TableData tableData = new TableData();
//    //    tableData.Columns.Add(new TableDataColumn("Object Type", DataType.Text));
//    //    tableData.Columns.Add(new TableDataColumn("Dimension Value", DataType.Double));
//    //    tableData.Columns.Add(new TableDataColumn("Dimension Type", DataType.Text));
//    //    tableData.Columns.Add(new TableDataColumn("Area Value", DataType.Double));

//    //    foreach (var dimension in dimensionData)
//    //    {
//    //        Row row = tableData.Rows.Add();
//    //        row.Cells[0].Text = dimension.ObjectType;
//    //        row.Cells[1].DoubleValue = dimension.DimensionValue;
//    //        row.Cells[2].Text = dimension.DimensionType;
//    //    }

//    //    foreach (var area in areaData)
//    //    {
//    //        Row row = tableData.Rows.Add();
//    //        row.Cells[0].Text = area.ObjectType;
//    //        row.Cells[1].DoubleValue = area.AreaValue;
//    //    }

//    //    return tableData;
//    //}


//}


//  public class DimensionData
//{
//    public string ObjectType { get; set; }
//    public double DimensionValue { get; set; }
//    public string DimensionType { get; set; }
//}

//public class AreaData
//{
//    public string ObjectType { get; set; }
//    public double AreaValue { get; set; }
//}
public class SaveToPdf
{
    public void Run()
    {
        //ExStart:ExportToPDF
        // The path to the documents directory.
        string MyDir = @"C:\temp\";
        string fileName = "Nagl.dwg";
        string sourceFilePath = MyDir + fileName;
        using (Image image = Image.Load(sourceFilePath))
        {
            // Create an instance of CadRasterizationOptions and set its various properties
            Aspose.CAD.ImageOptions.CadRasterizationOptions rasterizationOptions = new Aspose.CAD.ImageOptions.CadRasterizationOptions();
            rasterizationOptions.BackgroundColor = Aspose.CAD.Color.White;
            rasterizationOptions.PageWidth = 1600;
            rasterizationOptions.PageHeight = 1600;

            // Create an instance of PdfOptions
            Aspose.CAD.ImageOptions.PdfOptions pdfOptions = new Aspose.CAD.ImageOptions.PdfOptions();
            // Set the VectorRasterizationOptions property
            pdfOptions.VectorRasterizationOptions = rasterizationOptions;

            MyDir = MyDir + $"{fileName}.pdf";
            //Export the DWG to PDF
            image.Save(MyDir, pdfOptions);
        }
        //ExEnd:ExportToPDF            
        Console.WriteLine("\nThe DWG file exported successfully to PDF.\nFile saved at " + MyDir);

    }
}