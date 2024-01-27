// See https://aka.ms/new-console-template for more information
using Aspose.CAD;
using Aspose.CAD.FileFormats.Cad.CadConsts;
using Aspose.CAD.FileFormats.Cad.CadObjects;
using System;
using System.Collections.Generic;
using System.Windows.Forms;


public class ExtractDataFromDWG
{
    public List<CalculationEntry> Value(string fileName)
    {
        List<CalculationEntry> calculationEntries = new List<CalculationEntry>();
        using (Aspose.CAD.FileFormats.Cad.CadImage image = (Aspose.CAD.FileFormats.Cad.CadImage)Image.Load(fileName))
        {
            List<CadLwPolyline> lwPolyline = new List<CadLwPolyline>();

            // Go through each entity inside the DWG file
            foreach (CadEntityBase baseEntity in image.Entities)
            {
                Console.WriteLine(baseEntity.TypeName);
                // selection or filtration of entities
                if (baseEntity.TypeName == CadEntityTypeName.LWPOLYLINE)
                {
                    var cadLwPolyLine = (CadLwPolyline)baseEntity;
                    //Thickness defines that this polyline is the arial view of the Object --> Therefore we need it for the Area and Laufmeter calculation
                    if (cadLwPolyLine.Thickness != 0)
                    {
                        lwPolyline.Add(cadLwPolyLine);
                    }
                    
                }
            }
            var unitType = image.UnitType;
            var excelRowNumber = 2;
            foreach (var polyLine in lwPolyline)
            {
                var laufmeter =(int) Math.Round(polyLine.Length, 0);
                var flaeche = (int)Math.Round(polyLine.Area, 0);
                var dicke = (int)Math.Round(polyLine.Thickness,0);
                var outputString = $"{polyLine.Id} --> Laufmeter: {laufmeter} {unitType}, Fläche: {flaeche} {unitType}, Dicke: {dicke}";
                Console.WriteLine(outputString);

                var calculationEntry = new CalculationEntry
                {
                    Id = polyLine.Id,
                    Laufmeter = laufmeter,
                    Flaeche = flaeche,
                    Thickness = dicke,
                    KostenGesamtLaufmeter = $"=(C{excelRowNumber}/100)*D{excelRowNumber}",
                    KostenGesamtFlaeche = $"=(F{excelRowNumber}/100)*G{excelRowNumber}",
                    KostenQuadratmeterUndLaufmeter = $"=E{excelRowNumber}+H{excelRowNumber}"
                };
                //Only the first Line schould contain the sum
                //if (excelRowNumber == 2)
                //{
                //    calculationEntry.SummeGesamt = $"=SUMME(I2:I1000)";
                //}
                calculationEntries.Add(calculationEntry) ;
                
                excelRowNumber++;
            }

        }
        return calculationEntries;

    }
 
}

