<#
	!!! Autorisée l'utilisation des script powerhsell: https://www.pcastuces.com/pratique/astuces/3908.htm !!!!
	Attention lors de l'utilisation de se script, ce n'est pas un script micosoft officiel. 
	Le script ci dessous permet de convrtir tout type de fichier exel ( xlsl, odt ect ) en csv, il permet aussi de convertir 
	dans d'autre format. Deplus il permet d'écrire le nom du fichier à la dernière colone. 
	Attention aussi à ne faire que un seul type de conversion à la fois et à changer les fichier convertit de dossier pour ne par récrire deux fois à l'intérieur. 
#>

$ErrorActionPreference = 'Continue' <# Une erreur sera toujours affiché, je force donc le script à continuer #> 

Function Convert-CsvInBatch
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][String]$Folder
	)
	$ExcelFiles = Get-ChildItem -Path $Folder -Filter *.xlsx -Recurse <# Ici changler le .xlsx en ."formart depuis lequel on veut convertir" #>

	$excelApp = New-Object -ComObject Excel.Application 
	$excelApp.DisplayAlerts = $false

	$ExcelFiles | ForEach-Object {
		$workbook = $excelApp.Workbooks.Open($_.FullName).WorkSheets.Item(1)
		$csvFilePath = $_.FullName -replace "\.xlsx$", ".csv"  <# de même ici #>
		$workbook.SaveAs($csvFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV) 
		$workbook.Close()		

	}

	

	# Fermeture du exel com object
	$excelApp.Workbooks.Close()
	$excelApp.Visible = $true
	Start-Sleep 1
	$excelApp.Quit()
	$null=[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
	$null=[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
	[System.GC]::Collect()#>

	$ExcelFiles1 = Get-ChildItem -Path $Folder -Filter *.csv -Recurse 


	$ExcelFiles1 | ForEach-Object {
	$csvFilePath = $_.FullName
	$csv = Import-Csv -Delimiter ',' -Path $csvFilePath 
	$csv | Select-Object -Property *,@{label='NomDuFichier';Expression={$csvFilePath.remove(0,50)}} | <# écriture de nom de fichier sur toute les cellule de la dernière colone #>
	 export-csv $csvFilePath -NoTypeInformation
	$csv = Import-Csv -Delimiter ',' -Path $csvFilePath
	$csv
	$compte =  $csv | Measure-Object 
	$compte 
	$csv | Export-Csv $csvFilePath -NoTypeInformation -delimiter ';' <# Changement du délimiteur ',' en ';' pour être lisible sur par exel en français. 
	
	}

	
}
	

#
# 0. Mette le folderpath ou se présente les fichier exels ( = chemin des fichiers ) 
$FolderPath = "./"

Convert-CsvInBatch -Folder $FolderPath

#Par Thomas Burette



