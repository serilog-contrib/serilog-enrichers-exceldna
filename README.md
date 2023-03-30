| README.md |
|:---|

<div align="center">

<img src="assets/serilog-enrichers-exceldna-nuget.png" alt="Serilog.Enrichers.ExcelDna" width="100" />

</div>

<h1 align="center">Serilog.Enrichers.ExcelDna</h1>
<div align="center">

Enriches [Serilog](https://serilog.net) events with information from your [Excel-DNA](https://excel-dna.net) Add-in.

[![NuGet Version](http://img.shields.io/nuget/v/Serilog.Enrichers.ExcelDna.svg?style=flat)](https://www.nuget.org/packages/Serilog.Enrichers.ExcelDna/) [![Stack Overflow](https://img.shields.io/badge/stack%20overflow-serilog-orange.svg)](http://stackoverflow.com/questions/tagged/serilog) [![Stack Overflow](https://img.shields.io/badge/stack%20overflow-excel--dna-orange.svg)](http://stackoverflow.com/questions/tagged/excel-dna)

![Seq showing properties from Serilog.Enrichers.ExcelDna](assets/serilog-enrichers-exceldna-seq.png)

</div>

## Give a Star! :star:

If you like or are using this project please give it a star. Thanks!

## Getting started :rocket:

Install the [Serilog.Enrichers.ExcelDna](https://www.nuget.org/packages/Serilog.Enrichers.ExcelDna/) package from NuGet:

```powershell
Install-Package Serilog.Enrichers.ExcelDna
```

To apply the enricher to your logger configuration:

```csharp
var log = new LoggerConfiguration()
    .Enrich.WithXllPath()
    // ... other configuration ...
    .CreateLogger();
```

The `WithXllPath()` enricher will add an `XllPath` property to produced events.

### Included enrichers

The package includes:

* `WithXllPath()` - adds `XllPath` based with the value of `ExcelDnaUtil.XllPath`
* `WithExcelVersion()` - adds `ExcelVersion` with the value of `ExcelDnaUtil.ExcelVersion`
* `WithExcelVersionName()` - adds `ExcelVersionName` based on `ExcelDnaUtil.ExcelVersion`
* `WithExcelBitness()` - adds `ExcelBitness` based `Environment.Is64BitProcess`

### Example of an Excel-DNA add-in using Serilog with this sink

In the [sample](sample/) folder, there's an example of an Excel-DNA add-in that uses Serilog for logging to the `LogDisplay` of Excel-DNA using this sink, and apply the enrichers described above.

### Excel-DNA configuration for packing with `ExcelDnaPack`

In order for the Excel-DNA enricher to work from an add-in that was packaged using the `ExcelDnaPack` utility, you need to include references to `Serilog.dll` and `Serilog.Enrichers.ExcelDna.dll` in the `.dna` file of the add-in:

```xml
<DnaLibrary Name="My Add-In" RuntimeVersion="v4.0">
  <ExternalLibrary Path="MyAddIn.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" />

  <Reference Path="Serilog.dll" Pack="true" />
  <Reference Path="Serilog.Enrichers.ExcelDna.dll" Pack="true" />
```

## Release History

Click on the [Releases](https://github.com/serilog-contrib/serilog-enrichers-exceldna/releases) tab on GitHub.

---

_Copyright &copy; 2018-2023 C. Augusto Proiete & Contributors - Provided under the [Apache License, Version 2.0](LICENSE)._
