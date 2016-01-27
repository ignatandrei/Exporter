$rootFolder=$MyInvocation.MyCommand.Path
$rootFolder=Split-Path $rootFolder

$dllPath =Join-Path $rootFolder "libs" 

$dllExport =Join-Path  $dllPath "ExportImplementation.dll"
#Add-Type -LiteralPath "D:\github\Exporter\ExporterWordExcelPDF\lib\ExportImplementation.dll"
Add-Type -LiteralPath $dllExport

[string[]] $myArray=@("Name,WebSite,CV", "Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc","you, http://yourwebsite.com/,cv.doc")
$b=[ExportImplementation.ExportFactory]::ExportDataCsv( $myArray , [ExportImplementation.ExportToFormat]::Excel2007)
$dest=Join-Path $rootFolder "a.xlsx"
[System.IO.File]::WriteAllBytes($dest,$b)
Start-Process -FilePath $dest

