$timestamp = get-date -format 'HH:mm:ss'
Write-Host 'Starting Weekly Script'$timestamp  -ForegroundColor Green

start-process '(Batch file* is called here to run an Access macro that exports files to a specific folder)' -Wait -NoNewWindow

<# *Here's the code for the batch script:

echo Generating Weekly Files

del "\\...\Weekly Report\output files\*.xlsx" /s /f /q

start "C:\Program Files (x86)\Microsoft Office\Office\MSACCESS.EXE" "\\...\Weekly Report\Database.mdb" /x WeeklyMacro

echo Files in \\...\Weekly Report\output files
#>

$timestamp1 = get-date -format 'HH:mm:ss'
Write-Host 'Batch Script Done. Joining Excel books'$timestamp1  -ForegroundColor Yellow

# This next part takes all 3 Excel files and combines them into one

$file1 = '...\output files\qselSumForecast01a.xlsx' # Source's fullpath
$file2 = '...\output files\qselSumForecast01a2.xlsx'  # Source's fullpath
$file3 = '...\output files\qselSumForecast01a3.xlsx' # Destination's fullpath 
$xl = new-object -c Excel.application
$xl.displayAlerts = $false                      # Don't prompt the user
$wb1 = $xl.workbooks.open($file1, $null, $true) # Open source, readonly
$wb2 = $xl.workbooks.open($file2, $null, $true) # Open source, readonly
$wb3 = $xl.workbooks.open($file3)                         # Open target

$sh1_wb3 = $wb3.sheets.item(1)    # First sheet in destination workbook
$sh2_wb3 = $wb3.sheets.item(1)    # First sheet in destination workbook

$sheetToCopy = $wb1.sheets.item('qselSumForecast01a')   # Source sheet to copy
$sheetToCopy.copy($sh1_wb3) # Copy source sheet to destination workbook

$sheetToCopy = $wb2.sheets.item('qselSumForecast01a2')   # Source sheet to copy
$sheetToCopy.copy($sh2_wb3) # Copy source sheet to destination workbook


$wb1.close($false)                   # Close source workbook w/o saving
$wb2.close($false)                   # Close source workbook w/o saving
$wb3.close($true)                 # Close and save destination workbook
$xl.quit()

# Deletes the old excel files: file1 & file2
Remove-Item $file1
Remove-Item $file2

# Renames the main Excel file
Rename-Item -Path "...\output files\qselSumForecast01a3.xlsx" -NewName "File_Weekly.xlsx" 


$timestamp2 = get-date -format 'HH:mm:ss'
Write-Host 'Finished joining Excel books' $timestamp2  -ForegroundColor White


# Transferring the information from one workbook to another. The first Sheet

$xl = new-object -c Excel.application
$xl.displayAlerts = $false                      # Don't prompt the user

$timestamp4 = get-date -format 'HH:mm:ss'
Write-Host 'Starting copy and paste page 1 '$timestamp4  -ForegroundColor Yellow

$CellDate = get-date -format 'MM/dd/yy'
$Finalwb = '...\TemplateFile\ForecastWeeklytemplate.xlsx'

$Finalbook = $xl.Workbooks.open($Finalwb)

$Datawb = '...\output files\File_Weekly.xlsx'

$Databook = $xl.Workbooks.open($Datawb)

$Works1 = $Databook.WorkSheets.item(1)

$Works1.activate()

$xlPasteValues = -4163          # Values only, not formulas
$xlCellTypeLastCell = 11       # To find last used cell
$used = $Works1.usedRange
$lastCell = $used.SpecialCells($xlCellTypeLastCell)

$row = $lastCell.row
$range = $Works1.UsedRange
$range = $Works1.Range("A2:FI2$row")
$range.Copy()
$Works1 = $Finalbook.WorkSheets.item(1)
$Works1.Range("A4:FI4").PasteSpecial(-4163)

$Databook.close($false)

$savepath = '...\Weekly Report\ForecastWeekly.xlsx'
$Finalbook.SaveAs($savepath)

$Finalbook.close($false)

$timestamp5 = get-date -format 'HH:mm:ss'
Write-Host 'Completed Page 1 copy and paste. Starting Foreach loop '$timestamp5  -ForegroundColor Yellow

$Finalwb = '...\Weekly Report\ForecastWeekly.xlsx'
$wbf = $xl.Workbooks.open($Finalwb) # Open the final Excel workbook
$firstsheet = $wbf.sheets.item(1)    # First sheet in the final workbook
$firstsheet.Cells.Item(1,36) = $CellDate # Insert Date Data in cell
$firstsheet.Cells.Item(1,2) = '=SUBTOTAL(2,B4:B12000)' # Insert subtotal equation


# Foreach loop to paste an equation starting at AJ4. 
$i = 4
foreach ($cell in $firstsheet.Range('AJ4:AJ12000').Cells) {

 $cell.Value = '=IF(MIN(AK' + $i + ':DS' + $i + ')<0,$AJ$1-1+MATCH(IF(COUNTIF(AK' + $i + ':DS' + $i + ',"<0"),INDEX(AK' + $i + ':DS' + $i + ',MATCH(TRUE,INDEX(AK' + $i + ':DS' + $i + '<0,0),0)),""),AK' + $i + ':DS' + $i + ',0),"")'
	$i++
}

$wbf.close($true)


$timestamp6 = get-date -format 'HH:mm:ss'
Write-Host 'Completed worksheet 1 out of 3  '$timestamp6  -ForegroundColor Yellow


# Transferring the data from 2nd page to 2nd final page
$Finalwb2 = '...\Weekly Report\ForecastWeekly.xlsx' # Locating the destination final workbook

$Finalbook2 = $xl.workbooks.open($Finalwb2)

$Datawb2 = '...\Weekly Report\output files\File_Weekly.xlsx'


# Copying data over
$Databook2 = $xl.workbooks.open($Datawb2)

$Works2 = $Databook2.WorkSheets.item(2)

$Works2.activate()


$xlPasteValues2 = -4163          # Values only, not formulas
$xlCellTypeLastCell2 = 11       # To find last used cell
$used2 = $Works2.usedRange
$lastCell2 = $used2.SpecialCells($xlCellTypeLastCell2)

$row2 = $lastCell2.row
$range2 = $Works2.UsedRange
$range2 = $Works2.Range("A2:FI2$row2")
$range2.Copy()
$Works2 = $Finalbook2.WorkSheets.item(2)
$Works2.Range("A4:FI4").PasteSpecial(-4163)  

$Databook2.close($false)
$Finalbook2.close($true)  

$timestamp7 = get-date -format 'HH:mm:ss'
Write-Host 'Completed Page 2 copy and paste. Starting foreach loop '$timestamp7  -ForegroundColor Yellow

$wbf2 = $xl.workbooks.open($Finalwb2) 
$secondsheet = $wbf2.sheets.item(2)    
$secondsheet.Cells.Item(1,2) = '=SUBTOTAL(2,B4:B3500)' 


$i = 4
foreach ($cell in $secondsheet.Range('AJ4:AJ3500').Cells) {

 $cell.Value = '=IF(MIN(AK' + $i + ':DS' + $i + ')<0,$AJ$1-1+MATCH(IF(COUNTIF(AK' + $i + ':DS' + $i + ',"<0"),INDEX(AK' + $i + ':DS' + $i + ',MATCH(TRUE,INDEX(AK' + $i + ':DS' + $i + '<0,0),0)),""),AK' + $i + ':DS' + $i + ',0),"")'
	$i++
}

$wbf2.close($true)

$timestamp8 = get-date -format 'HH:mm:ss'
Write-Host 'Completed worksheet 2 out of 3  '$timestamp8  -ForegroundColor Yellow
 

# Transferring the data from 3rd page to 3rd final page
$Finalwb3 = '...\Weekly Report\ForecastWeekly.xlsx' # Locating the destination final workbook

$Finalbook3 = $xl.workbooks.open($Finalwb3)

$Datawb3 = '...\Weekly Report\output files\File_Weekly.xlsx'


$Databook3 = $xl.workbooks.open($Datawb3)

$Works3 = $Databook3.WorkSheets.item(3)

$Works3.activate()

$xlPasteValues3 = -4163          
$xlCellTypeLastCell3 = 11       
$used3 = $Works3.usedRange
$lastCell3 = $used3.SpecialCells($xlCellTypeLastCell3)

$row3 = $lastCell3.row
$range3 = $Works3.UsedRange
$range3 = $Works3.Range("A1:R1$row3")
$range3.Copy()
$Works3 = $Finalbook3.WorkSheets.item(3)
$Works3.Range("A1:R1").PasteSpecial(-4163) 

$Databook3.close($false)
$Finalbook3.close($true)  
$xl.Quit()			
spps -n excel


$timestamp8 = get-date -format 'HH:mm:ss'
Write-Host 'Completed worksheet 3 out of 3  '$timestamp8  -ForegroundColor Yellow

# Renaming the file to give it a date
$filenameFormat = "File_Weekly_" + (Get-Date -Format "yyyy-MM-dd") + ".xlsx"
Rename-Item -path "...\Weekly Report\ForecastWeekly.xlsx" -NewName $filenameFormat

$timestampFinal = get-date -format 'HH:mm:ss'
Write-Host 'Done!'$timestampfinal  -ForegroundColor Green