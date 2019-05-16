<#  Read the Excel file WorkbookName an loop into its sheets skipping 1st sheet (named HLQ).
For each one of the sheets after sheet "HLQ", create a new workbook composed by 2 sheets:

    - 1st sheet (sheet HLQ)
	- current sheet 
#>

# set input Excel filename (also used as a base filename for the splited workbooks)
$WorkbookName = "GC_DATA_LINEAGE_v1D.XLSX"
$filepath ="C:\Users\0101675\Documents\BCP - CDO\Lineage - PS-GC\"+$WorkbookName
$WorkbookName = $filepath -replace ".xlsx", ""

#####################################################
$Excel = New-Object -ComObject "Excel.Application" 
$Excel.Visible = $false #Runs Excel in the background. 
$Excel.DisplayAlerts = $false #Supress alert messages. 
$Workbook = $Excel.Workbooks.open($filepath)
$numb_sheets = $Workbook.Worksheets.Count
$sheet_HLQ = $WorkBook.Worksheets(1)
#
$FileFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook 
$output_type = "xlsx"

# Loop on sheets
write-Output "Processing $WorkbookName with $numb_sheets sheets..." 
$i=0
foreach($Worksheet in $Workbook.Worksheets) {
    $i++
    if ( $i -gt 1 ){ # skip fist sheet as it is copied for all workbooks

		# create a new workbook with current sheet Content
		$Worksheet.copy()
        $newExcelFile = $WorkbookName + "_" + $Worksheet.Name + "." + $output_type 
        $Excel.ActiveWorkbook.SaveAs($newExcelFile, $FileFormat) 
        $Excel.ActiveWorkbook.Close

        write-Output "Created file: $newExcelFile with sheet $Worksheet.Name"
		
		# insert sheet HLQ at begining of the newly created workBook $new_wb
		$new_wb = $Excel.workbooks.open($newExcelFile) # open target
		$sheet_HLQ.Copy($new_wb.Worksheets($Worksheet.Name))
		$Excel.ActiveWorkbook.Save() 
        $Excel.ActiveWorkbook.Close
    }
}

# Closing all 
$Workbook.Close() 
$Excel.Quit() 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Stop-Process -Name EXCEL
Remove-Variable Excel