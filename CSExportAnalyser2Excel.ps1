<#
PowerShell Script to run csexport.exe and CSExportAnalyzer.exe consecutively on
all MAs in MIM and Export them to an Excel document.

Module Dependencies:
- SqlServer from PSGallery
- ImportExcel from PSGallery

File Dependencies:
- Assumes that CSExportAnalyser.exe is in MIMs bin-folder.

Considerations:
- There is a character length limit on table and sheet names that the script gets around by
removing spaces and cropping the MA name to a maximum of 25 characters. This will create
a naming conflict in the excel file if multiple MAs share the first 25 charactes of their names.
#>

$serverInstance = "<DB Instance>"
$databaseName = "MIMSynchronizationService"
$schemaName = "dbo"
$tableName = "mms_management_agent"
$columnName = "ma_name"
$mimInstallFolder = "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service"

$csExportCommand = $mimInstallFolder+"\Bin\csexport.exe"
$csExportAnalyserCommand = $mimInstallFolder+"\Bin\CSExportAnalyzer.exe"
$csExportFolder = "D:\Temp\csExportFolder"

[System.IO.Directory]::CreateDirectory($csExportFolder)

$timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$outputXLSX = $csExportFolder+"\"+"csExport_"+$timestamp+".xlsx"

$dataRows = Read-SqlTableData -ServerInstance $serverInstance -DatabaseName $databaseName -SchemaName $schemaName -TableName $tableName -ColumnName $columnName -OutputAs DataRows

foreach ($dataRow in $dataRows)
{
    $maName = $dataRow.ItemArray[0]
    
    $xml = $csExportFolder+"\"+$maName+".xml"
    $csv = $csExportFolder+"\"+$maName+".csv"

    & $csExportCommand "$maName" $xml /f:x
    & $csExportAnalyserCommand $xml > $csv

    $maName = $maName -replace '\s',''
    $maName = $maName.Substring(0, [System.Math]::Min(25, $maName.Length))

    Import-Csv -Path $csv | Export-Excel -Path $outputXLSX -AutoSize -FreezeTopRow -TableName $maName -WorksheetName $maName

    Remove-Item -Path $xml
    Remove-Item -Path $csv
}
