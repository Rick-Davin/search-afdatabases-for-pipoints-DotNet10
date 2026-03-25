# search-afdatabases-for-pipoints-DotNet10
Uses AF SDK for .NET 8 NuGet package.  Searches an AFDatabase for PIPoints and logs them to an Excel workbook.

You do not need Excel in order to run this application.  Granted, you will need Excel - or an equivalent viewer - to view the output from this application.

<h2>NuGet Packages</h2>

Aveva.AFSDK - obviously

ClosedXml - no need for sluggish Excel Automation and COM objects.  This package is very fast and super easy to use.  It will load and interact with OpenXml, which is not so easy to use.

Microsoft.Extension.Hosting - we want to do things the newer .NET way and not .NET Framework.

Serilog.Extensions.Hossting - for our console and file logging

There are dozens of other transitive packages loaded by loading the above.
