// See https://aka.ms/new-console-template for more information
using CsvHelper.Configuration.Attributes;
using System;
using System.Text.Json.Serialization;

public class CalculationEntry
{
    [Name("Id")]
    public String Id { get; set; }

    [Name("Thickness (mm)")]
    public int Thickness { get; set; }

    [Name("Laufmeter (mm)")]
    public int Laufmeter { get; set; }

    [Name("Preis Pro Laufmeter")]
    public int PreisProLaufmeter { get; set; }

    [Name("Kosten gesamt Laufmeter")]
    public string KostenGesamtLaufmeter { get; set; }

    [Name("Quadratmeter (mm)")]
    public int Flaeche { get; set; }

    [Name("Preis pro Quadratmeter")]
    public int PreisProQuadratMeter { get; set; }

    [Name("Kosten gesamt Quadratmeter")]
    public string KostenGesamtFlaeche { get; set; }

    [Name("Kosten Laufmeter + Quadratmeter ")]
    public string KostenQuadratmeterUndLaufmeter { get; set; }

    [Name("Kosten Gesamt Projekt")]
    public string SummeGesamt { get; set; }


}

