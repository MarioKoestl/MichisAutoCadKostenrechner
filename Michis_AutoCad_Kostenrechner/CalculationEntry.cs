// See https://aka.ms/new-console-template for more information
using CsvHelper.Configuration.Attributes;
using System;
using System.Text.Json.Serialization;

public class CalculationEntry
{
    [Name("Id")]
    public String Id { get; set; }

    [Name("Position")]
    public int Thickness { get; set; }

    [Name("Stück")]
    public int Stueck { get; set; }

    [Name("Laufmeter (m)")]
    public double Laufmeter { get; set; }

    [Name("Preis Pro Laufmeter")]
    public double PreisProLaufmeter { get; set; }

    [Name("Kosten gesamt Laufmeter")]
    public string KostenGesamtLaufmeter { get; set; }

    [Name("Quadratmeter (m2)")]
    public double Flaeche { get; set; }

    [Name("Preis pro Quadratmeter")]
    public double PreisProQuadratMeter { get; set; }

    [Name("Kosten gesamt Quadratmeter")]
    public string KostenGesamtFlaeche { get; set; }

    [Name("Kosten Laufmeter + Quadratmeter")]
    public string KostenQuadratmeterUndLaufmeter { get; set; }

    [Name("Kosten Gesamt Projekt")]
    public string SummeGesamt { get; set; }

    [Name("Fertig Breite")]
    public int FertigmasBreite { get; set; }
    [Name("Fertig Länge")]
    public int FertigmasLength { get; set; }

    [Name("Roh Breite")]
    public string RohmasBreite { get; set; }
    [Name("Roh Länge")]
    public string RohmasLength { get; set; }

}

