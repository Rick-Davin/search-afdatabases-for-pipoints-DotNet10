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

<h2>Not a Trivial Port</h2>

I wrote the .NET Framework app a year ago in anticipation of migrating to .NET 8.  So I did not use the **AppConfig** file.  I used **appsettings.json**, though I had my own custom reader code, as well as a custom console logger.

For .NET 10, I start off using a **Generic Host**, where I can read appsettings.json and Environment variables easily.  I also switched so the **ILogger<T>** pattern.  **Serilog** was my chosen logging package, though you are free to change to your heart's desire.  All of this makes for a very different looking Program.Main that may be totally foreign to anyone who has only used .NET Framework.  While one could argue that AS SDK is rooted in the past, you can still be forward-looking as you migrate to .NET 8, 10, or beyond.

Though there are few **async** calls, I set up the app as if there would be lots of them.  Again, I did not want a trival application.

I give Claude AI some thanks for helping out big time with Program.Main and the Dependency Injection of a MainWorker instance.

After all that, you would think it would be ready to build but buckle in a LOT of nitpicky nullability issues.  Consider the .NET Framework snippet of **PIPoint tag = null;**  In .NET 8+, this produces a compile error.  You will have to use **PIPoint?** instead.  And you will find yourself going up-and-down in your code peppering it with **?** when it can be null or **!** when you know it absolutely is not null.

And I lost a lot of time not realizing that **Costura.Fody** was blocking the PI calls.  All in all, I was able to port it over in around 6-hours.  I expect to spend more hours editing the README or creating a Wiki.

<h2>Prepping for Aveva.AFSDK</h2>

It is not as simple as installing the NuGet package into your application.  I recommend you install all the prerequisite Microsoft Visual C++ Redistributables.  I don't care if you did it last month, do it again.

Next, what seems not be mentioned in the AVEVA docs yet is that you really should have AF Client 2024 installed.  The key reason for this is that the KST (Known Servers Table) is moved out of the Windows Registry and as separate file.  It may not run everything for you, so you may need to manually run C:\Program Files\PIPC\AF\KSTMigrationTool.exe to create the KST as a file.
