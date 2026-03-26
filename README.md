# Search AFDatabases for PIPoints using .NET 10
Uses AF SDK for .NET 8 NuGet package.  Searches an AFDatabase for PIPoints and logs them to an Excel workbook.

You do not need Excel in order to run this application.  Granted, you will need Excel - or an equivalent viewer - to view the output from this application.  In fact, I am using LibreOffice Calc to view the generated workbooks.

<h2>NuGet Packages</h2>

**Aveva.AFSDK** - obviously

**ClosedXml** - no need for sluggish Excel Automation and COM objects.  This package is very fast and super easy to use.  It will load and interact with OpenXml, which is not so easy to use to put it nicely.  If I were to put it not so nicely, it would be a 2-page rant.

**Microsoft.Extension.Hosting** - we want to do things the newer .NET way and not .NET Framework.  

**Serilog.Extensions.Hosting** - for our console and file logging.

There are dozens of other transitive packages loaded by loading the above.

Finally, one NuGet package to add but use with caution is **Costura.Fody**, which is an assembly weaver.  It will bundle up all the libs and DLL's so the resulting build will have less than 10 files in the build folder.  For this to work with AFSDK, I needed to exclude 1 DLL beginning with "Aveva" and 20 beginning with "OSIsoft".  This reduces my build folder from over 130 files down to about 30.  While that is big improvement, I would still like to eventually see less than 10 files in the build folder one day.  Meanwhile, I am not happy to think that every .NET 8+ application I build will need almost 21 AVEVA-related DLL's residing in the build folder.

<h2>Prepping for Aveva.AFSDK</h2>

It is not as simple as installing the NuGet package into your application.  I recommend you install all the prerequisite Microsoft Visual C++ Redistributables.  I don't care if you did it last month, do it again.

Next, what seems not be mentioned in the AVEVA docs yet is that you really should have **AF Client 2024** installed.  The key reason for this is that the KST (Known Servers Table) is moved out of the Windows Registry and as separate file.  It may not run everything for you, so you may need to manually run C:\Program Files\PIPC\AF\KSTMigrationTool.exe to create the KST as a file.  With AF Client 2018, the KST is still stored in the Registry, and that will not work with the AFSDK NuGet package.

<h2>Not a Trivial Application</h2>

Just one day after AVEVA released its first ever NuGet package **Aveva.AFSDK**, I was able to port over a not-so-trivial .NET Framework application.  The app uses a paged **AFAttributeSearch** filtering on PIPoint data reference plugin.  The Asset-related info is then linked to Tag-related info thanks to bulk calls for PIPoint definitions and PI data calls. 

The information is written to an Excel worksheet, and formatted with boders, color shadings, and a filter applied.  All this is thanks to **ClosedXml** and does not require the Excel application be installed to generate the worksheet.  There is a 1 million row limitation with Excel, so it may be best to choose a database that will find much less than that.  If you are sure you have a large database and unsure you can safely be less than 1 million rows, an **appsetting.json** property can request the output be a text file.

To write the output, I employ an **IReportWriter interface** along with a **ReportWriter abstract class** containing some virtual methods.  There are 2 concrete classes: **ExcelWriter** and **TextWriter**, and each has some overridden methods specific to its own implementation, but each also relies upon some common virtual methods defined in ReportWriter.

All in all, it is not a huge or major application but then again it is not small, trivial one either.  My eyes would glass over in sheer boredom if all you showed to me is that you could connect to a PIServer and query the snapshot of SINUSOID.

<h2>Not a Trivial Port</h2>

My overall goal was not to take a .NET Framework app and force fit it to run in .NET 10.  Rather, ***I wanted to create a .NET 10 app that just so happens to use AFSDK***.

I wrote the original .NET Framework app a year ago in anticipation of migrating to .NET 8.  At that time, I chose not to use the **AppConfig** file.  I used **appsettings.json** as it would also be used in .NET 8/10, though I had my own custom reader code, as well as a custom console logger.

For .NET 10, I start off using a **Generic Host**, where I can read appsettings.json and Environment variables easily.  I also switched so the **ILogger<T>** pattern.  **Serilog** was my chosen logging package, though you are free to change to your heart's desire.  All of this makes for a very different looking Program.Main that may be totally foreign to anyone who has only used .NET Framework.  While one could argue that AFSDK is rooted in the past, you can still be forward-looking as you migrate to .NET 8, 10, or beyond.  So get accustomed to a GlobalUsings class, and lots of nullability issues (more below).  You should absolutely bring your knowledge and experience about AFSDK forward but you may have to leave a lot of old .NET Framework habits behind as you embrace .NET.

Though there are few **async** calls, I set up the app as if there would be lots of them.  Again, I did not want a trival application.

I give Claude AI some thanks for helping out big time with Program.Main and the Dependency Injection of a MainWorker instance, as well as helping with ILogger calls and Serilog configuration.

After all that, you would think it would be ready to build but buckle in to fix a LOT of nitpicky nullability issues.  Consider the .NET Framework snippet of **PIPoint tag = null;**  In .NET 8+, this produces a compile error.  You will have to use **PIPoint?** instead.  And you will find yourself going up-and-down in your code peppering it with **?** when it can be null or **!** when you know it absolutely is not null.  I believe a developer should give specific thought as to whether a given object can ever be null versus that is should never be null.

And I lost a lot of time not realizing that **Costura.Fody** was blocking the PI calls.  All in all, I was able to port it over in around 6-hours and that including me fumbling around with migrating the KST.  I expect to spend more hours editing the README or creating a Wiki, mainly because I believe/hope that this is a good learning example.

I give a shout out to a special friend who shall remain nameless.  Thanks for letting me use your test server.
