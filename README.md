# Search AFDatabases for PIPoints using .NET 10
Uses AF SDK for .NET 8 NuGet package.  Searches an AFDatabase for PIPoints and logs them to an Excel workbook.

You do not need Excel in order to run this application.  Granted, you will need Excel - or an equivalent viewer - to view the output from this application.

<h2>NuGet Packages</h2>

**Aveva.AFSDK** - obviously

**ClosedXml** - no need for sluggish Excel Automation and COM objects.  This package is very fast and super easy to use.  It will load and interact with OpenXml, which is not so easy to use.

**Microsoft.Extension.Hosting** - we want to do things the newer .NET way and not .NET Framework.  

**Serilog.Extensions.Hossting** - for our console and file logging

There are dozens of other transitive packages loaded by loading the above.

Sadly, one NuGet package I was unable to use was **Costura.Fody**, which is an assembly weaver.  It will bundle up all the libs and DLL's so the resulting build will have less than 10 files in the build folder.  I get AF SDK errors or no data when I use it, so I was forced to omit it.  And now I have over 130 files in the build folder.  Plus, I am not happy to think that every .NET 8+ application I build will need almost 20 AVEVA DLL's residing in the build folder.

<h2>Not a Trivial Application</h2>

Just one day after AVEVA released its first ever NuGet package **Aveva.AFSDK**, I was able to port over a not-so-trivial .NET Framework application.  The app uses a paged **AFAttributeSearch** filtering on PIPoint data reference plugin.  The Asset-related info is then linked to Tag-related info thanks to bulk calls for PIPoint definitions and PI data calls. 

The information is written to an Excel worksheet, and formatted with boders, color shadings, and a filter applied.  All this is thanks to **ClosedXml** and does not require the Excel application be installed to generate the worksheet.  There is a 1 million row limitation with Excel, so it may be best to choose a database that will find much less than that.  If you are sure you have a large database and unsure you can safely be less than 1 million rows, an **appsetting.json** property can request the output be a text file.

To write the output, I employ an **IReportWriter interface** along with a **ReportWriter abstract class** containing some virtual methods.  There are 2 concrete classes: **ExcelWriter** and **TextWriter**, and each has some overridden methods specific to its own implementation, but each also relies upon some common virtual methods defined in ReportWriter.


