﻿Clear-Host
Set-StrictMode -Version latest
$path = "C:\"
$files = Get-Childitem $path -Include *.xls,*.xlsx –Force –Recurse –ErrorAction SilentlyContinue –ErrorVariable AccessDenied | Where-Object { !($_.psiscontainer) }
$word = "word"
$Excel = New-Object -comobject excel.application
$Excel.visible = $true
$Excel.DisplayAlerts = $False
$count = 0
"Location:" | Add-Content -path "excelFindings.csv"

foreach($file In $files){
    try{
		$ExcelWorkBook = $Excel.Workbooks.Open($file,0,$true,5,"LETMEIN",$null,$true)
		$Worksheets = $ExcelWorkBook.worksheets
	}

    catch{
		Write-Output "CAUGHT - " $file.fullname
        $file.fullname | Add-Content -path "excelFindings.csv"
        continue
	}

    foreach($worksheet In $Worksheets){

		try{
			$Range = $Worksheet.Range("A1:Z1").EntireColumn
			$found = $false
			$found = $Range.find($word)

			if($found){
				$file.fullname | Add-Content -path "excelFindings.csv"
			}

        }

		catch{
            Write-Output "CAUGHT2 - " $file.fullname
		    $file.fullname | Add-Content -path "excelFindings.csv"
		    $ExcelWorkBook.close($false);
            continue
		}

    }

	Write-Output $file.fullname
    $ExcelWorkBook.close($false);

}

$Excel.Quit()
#Stop-Process -name "EXCEL.EXE"
