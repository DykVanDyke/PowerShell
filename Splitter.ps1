<#  Read the Excel file WorkbookName an loop into its sheets skipping 1st sheet (named HLQ).
For each one of the sheets after sheet "HLQ", create a new workbook composed by 2 sheets:

    - 1st sheet (sheet HLQ)
	- current sheet 
#>

# Get this Script Name 
$ThisScript = (Get-Item $PSCommandPath ).Name

# Check input parameter
if ( $args.count -eq 0  ){
   echo "================================================="
   echo "Missing input parameter!"
   echo " "
   echo "  Usage: .\$ThisScript <Input File>"
   echo "Example: .\$ThisScript GC_DATA_LINEAGE_v1D.XLSX"
   echo "================================================="
   echo " "
   exit
}

# Set input Excel filename from the given input parameter ( which will also be used as a base filename for the splited workbooks)
$WorkbookName=$args[0]
$filepath=(Get-Item $WorkbookName)
$filename=(Get-Item $WorkbookName).Name
$WorkbookName = $filepath -replace ".xlsx", ""

### Init Excel Object
$Excel = New-Object -ComObject "Excel.Application" 
$Excel.Visible = $false #Runs Excel in the background. 
$Excel.DisplayAlerts = $false #Supress alert messages. 
$Workbook = $Excel.Workbooks.open($filepath)

# set specific parameters
$numb_sheets = $Workbook.Worksheets.Count-1
$sheet_HLQ = $WorkBook.Worksheets(1)
$FileFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook 
$output_type = "xlsx"

# Loop on sheets
write-Output "============================================================================================================= "
write-Output "Processing Excel file $filename with $numb_sheets sheets containing Lineage Information Input ..." 
write-Output ""
$i=0
foreach($Worksheet in $Workbook.Worksheets) {
    $i++
    if ( $i -gt 1 ){ # skip fist sheet as it is copied for all workbooks

		# create a new workbook with current sheet Content
		$Worksheet.copy()
        $newExcelFile = $WorkbookName + "_" + $Worksheet.Name + "." + $output_type 
        $Excel.ActiveWorkbook.SaveAs($newExcelFile, $FileFormat) 
        $Excel.ActiveWorkbook.Close()

        write-Output "Created file: $newExcelFile"
		
		# insert sheet HLQ at begining of the newly created workBook $new_wb
		$new_wb = $Excel.workbooks.open($newExcelFile) # open target
		$sheet_HLQ.Copy($new_wb.Worksheets($Worksheet.Name))
		$Excel.ActiveWorkbook.Save() 
        $Excel.ActiveWorkbook.Close()
    }
}

# Closing all 
$Workbook.Close()
$Excel.Quit() 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Stop-Process -Name EXCEL
Remove-Variable Excel