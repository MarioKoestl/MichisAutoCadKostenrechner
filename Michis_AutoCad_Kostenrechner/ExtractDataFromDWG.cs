// See https://aka.ms/new-console-template for more information
using Aspose.CAD;
using Aspose.CAD.FileFormats.Cad.CadConsts;
using Aspose.CAD.FileFormats.Cad.CadObjects;
using System;
using System.Collections.Generic;


public class ExtractDataFromDWG
{
 
    public List<CalculationEntry> Value(string fileName)
    {
        List<CalculationEntry> calculationEntries = new List<CalculationEntry>();
        List<CadLwPolyline> lwPolyline = new List<CadLwPolyline>();

        using (Aspose.CAD.FileFormats.Cad.CadImage image = (Aspose.CAD.FileFormats.Cad.CadImage)Image.Load(fileName))
        {
            //Go through each entity inside the DWG file
            foreach (var blockEntity in image.BlockEntities.Values)
            {
                var cadBlockEntity= (CadBlockEntity)blockEntity;

                foreach (CadEntityBase baseEntity in cadBlockEntity.Entities)
                {
                    Console.WriteLine(baseEntity.TypeName);
                    ;
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
            }
            var unitType = image.UnitType;
            var excelRowNumber = 2;
            foreach (var polyLine in lwPolyline)
                {
                var laufmeter =Math.Round(polyLine.Length/1000, 2);
                var flaeche = Math.Round(polyLine.Area/1000000, 2);
                var dicke = (int)Math.Round(polyLine.Thickness,0);
                var outputString = $"{polyLine.Id} --> Laufmeter: {laufmeter} {unitType}, Fläche: {flaeche} {unitType}, Dicke: {dicke}";
                Console.WriteLine(outputString);

                var calculationEntry = new CalculationEntry
                {
                    Id = polyLine.Id,
                    Laufmeter = laufmeter,
                    Flaeche = flaeche,
                    Thickness = dicke,
                    Stueck=1,
                    KostenGesamtLaufmeter = $"=D{excelRowNumber}*E{excelRowNumber}*C{excelRowNumber}",
                    KostenGesamtFlaeche = $"=G{excelRowNumber}*H{excelRowNumber}*C{excelRowNumber}",
                    KostenQuadratmeterUndLaufmeter = $"=F{excelRowNumber}+I{excelRowNumber}",
                    FertigmasLength = 0,
                    FertigmasBreite =0,
                    RohmasLength = $"=M{excelRowNumber}-16",
                    RohmasBreite = $"=L{excelRowNumber}-16"
                };
                //Only the first Line schould contain the sum
                //if (excelRowNumber == 2)
                //{
                //    calculationEntry.SummeGesamt = $"=SUMME(K2:K1000)";
                //}
                calculationEntries.Add(calculationEntry) ;
                
                excelRowNumber++;
            }

        }
        return calculationEntries;

    }
 
}

