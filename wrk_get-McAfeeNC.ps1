<#
.SYNOPSIS
	Collects information about McAfee versions and returns the non compliant computers
	
	
.AUTHOR
	Philippe Martin - Huleu Consult s.p.r.l.

.DESCRIPTION
	This script reads the csv file McAfee53x_SIP.csv on server pvtwivir53epo, update the Excel file McAfee_DashBoard_KPI.xlsx and 
	copy the original csv file to '\\beprod01.eoc.net\offshare\Infrastructure Evolution\001 WSS Group\Windows Support\02. Services\AntiVirus\1. Control\McAfee53x_SIP_Backup_File 

.EXAMPLE
		.\Get-SqlServerInventoryToClixml.ps1 -DNSServer automatic -DNSDomain automatic -PrivateOnly
		
.OUTPUTS
		updated row in \\beprod01\OffShare\Infrastructure Evolution\001 WSS Group\Windows Support\02. Services\AntiVirus\1. Control\McAfee_DashBoard_KPI.xlsx

.NOTES
	This script is still under development
	
.LINK
	
#>

function no-lf ($str) {
	if ($str.IndexOf([char]10) -ne -1) { 
		$str = $str.replace([char]10, '¥') 
		}
	$str
	}

function get-XLHeaser {
	Param (
		[parameter(Mandatory=$true,
    	ValueFromPipeline=$true)]
		[string] $strWorkBook,
		
		[parameter(Mandatory=$true,
		ValueFromPipeline=$true)]
		[string] $strWorkSheet,
		
		[parameter(Mandatory=$false,
		ValueFromPipeline=$true)]
		[string[]] $arrCols
		)
		
	$objXL = New-Object -Com Excel.Application
	$objXL.DisplayAlerts = $false
	$objXL.visible = $false
	$wbk = $objXL.Workbooks.Open($strWorkBook, $null, $true)
	$wks = $wbk.Worksheets.Item($strWorkSheet)
	$rg = $wks.UsedRange	
	
	$arrHeader = $rg.Rows[1] | %{$_.Value2}
 	$objXL.Workbooks.Close()
	$objXL.Quit()
 	$arrHeader 
 	}

function set-XLData {
	Param (
		[parameter(Mandatory=$true,
    	ValueFromPipeline=$true)]
		[string] $strWorkBook,
		
		[parameter(Mandatory=$true,
		ValueFromPipeline=$true)]
		[string] $strWorkSheet,
		
		[parameter(Mandatory=$false,
		ValueFromPipeline=$true)]
		$hshRow
		)
		
	$objXL = New-Object -Com Excel.Application
	$objXL.DisplayAlerts = $false
	$objXL.visible = $false
	$wbk = $objXL.Workbooks.Open($strWorkBook)
	$wks = $wbk.Worksheets.Item($strWorkSheet)
	$rg = $wks.UsedRange	
	
	$arrHeader = $rg.Rows[1] | %{$_.Value2}
	$arrXL = @()
	
	# retrieve the row correspondign to the current date
	$objColDate = $rg.columns[$arrheader.indexof('DATE') + 1].cells
	$dtDate = [datetime]::ParseExact($hshRow['Date'],'dd/MM/yyyy',$null)
	$intRow = 1
	do {$intRow++} until ([datetime]$objColDate.Item($intRow,$objColDate.row).value() -eq $dtDate)

	# fill the cells 
	for ($i = 0; $i -le $hshRow.Count; $i++) {
		$strCol = no-lf $objColDate.Item(1,$i).Text
		if (!($strCol -in ('TOTAL¥NO Compliant', 'TOTAL¥Compliant', 'DAY', '%¥Compliant'))) {
			$objColDate.Item($intRow,$i) = switch ($strCol) {
				'DATE' {
					[datetime]::ParseExact($hshRow[$strCol],'dd/MM/yyyy',$null)
					}
				default {
					$hshRow[$strCol]
					}
				}
			}
		
		}
	$wbk.save()
 	$objXL.Workbooks.Close()
	$objXL.Quit()
 	}

cls
# Determine script location
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path	
$rootscriptDir = $ScriptDir.Substring(0,$ScriptDir.IndexOf('WindowsPowershell')  + ('WindowsPowershell').Length) 
#Import-Module ("{0}\Modules\EOC" -F $rootscriptDir)

# Making excessively long column names easier to manipulate
$hshHeader = @{DAT = 'DAT Version (VirusScan Enterprise)'; Product = 'Product Version (Endpoint Security Threat Prevention)'; AM = 'AMCore Content Version'}

$dtToday = get-date

# on Wednesday get the three last records, the two last on the other days
$intLastRec = if ($dtToday.DayOfWeek -eq 'Wednesday') {
				3
				}
			else {
				2
				}
				
# actuel location of destination Excel workbook
$strDestXL = '\\beprod01.eoc.net\offshare\Infrastructure Evolution\001 WSS Group\Windows Support\02. Services\AntiVirus\1. Control\McAfee_DashBoard_KPI.xlsx'
# location of a file for testing purpose
$strDestXL = "{0}\DATA\McAfee_DashBoard_KPI.xlsx" -f $ScriptDir

$hshResult = @{}
get-XLHeaser $strDestXL 'Data' | %{
	if ($_) {
		if ($_.IndexOf([char]10) -ne -1) {
			$hshResult.Add(($_ -replace [char]10, '¥'), '')
			}
		else {
			$hshResult.Add($_, '')
			}
		}
	}

$hshResult['Day'] = $dtToday.DayOfWeek
$hshResult['Date'] = "{0:dd}/{0:MM}/{0:yyyy}" -f $dtToday
if (Test-Connection pvtwivir53epo -quiet) {
	# get data from the original file
	$strWorkBook = '\\pvtwivir53epo\e$\reports\McAfee53x_SIP.csv'
#	$strWorkBook = "{0}\Data\2018_12_19_McAfee53x_SIP.csv" -f $rootscriptDir

	$arrCsv = Import-Csv $strWorkBook
	$arrNC = @()

	# get the number of records contained in the file
	$hshResult['Total'] = $arrCsv.Count

	# get the latest version of 'Product Version (Endpoint Security Threat Prevention)'
	$strLatest = ($arrCsv.($hshHeader.Product) | measure -Maximum).Maximum

	# get the number of computers running the latest version
	$intCurrRelease = ($arrCsv.($hshHeader.Product) | ?{$_ -eq $strLatest} | measure -Maximum).count
	$hshResult['McAfee End Point Security¥ Total'] = $intCurrRelease
	
	# add the number of computers where both 'Product Version (Endpoint Security Threat Prevention)' and 'DAT Version (VirusScan Enterprise)' are empty 
	$arrCsv | ?{$_.($hshHeader.Product) -ne $strLatest} | %{
		if ($_.($hshHeader.DAT) -eq '') {
			$arrNC += "{0};{1};{2};{3}" -f 'EndPoint and DAT empty', $_.'System Name', $_.'IP Address', $_.'Operating System'
			$hshResult['McAfee End Point Security¥ Total'] += 1
			}
		}

	# get the number of compliant computers (running the two (on Wednesday the three) last versions of 'AMCore Content Version'
	$arrValues = ($arrCsv.($hshHeader.AM)) | sort -Unique  | ?{$_ –ne ''}
	$arrCsv | sort ($hshHeader.AM) -Unique  | ?{$_ –ne ''} | %{
		$arrNC += "{0};{1};{2};{3}" -f 'DAT not empty', $_.'System Name', $_.'IP Address', $_.'Operating System'
		}
	$hshResult['EndPoint¥Compliant¥(3477 & 3480)'] = ($arrCsv | ?{$_.($hshHeader.AM) -in (($arrValues) | sort -Unique | select -Last $intLastRec | ?{$_ –ne ''})}).count
	($arrCsv | ?{$_.($hshHeader.AM) -in (($arrValues) | sort -Unique | select -Last $intLastRec | ?{$_ –ne ''})}).count

	# get the number of non compliant computers running other 'AMCore Content Version'
	$hshResult['EndPoint¥Non Compliant'] = ($arrCsv.($hshHeader.AM) | ?{$_ -in ($arrValues | select -First ($arrValues.Count - $intLastRec))}).count

	# get the number of computers unable to retrieve AMCore
	$hshResult['EndPoint ¥Unable To retrieve AMCore'] = 0
	$arrCsv | ?{$_.($hshHeader.AM) –eq ''}| %{
		if ($_.($hshHeader.DAT) -eq '') {
			$arrNC += "{0};{1};{2};{3}" -f 'Unable To retrieve AMCore', $_.'System Name', $_.'IP Address', $_.'Operating System'
			$hshResult['EndPoint ¥Unable To retrieve AMCore'] += 1
			}
		}
	
	# get the number of computers running McAfee VirusScan Enterprise
	$hshResult['McAfee VirusScan Enterprise¥Total '] = ($arrCsv | ?{$_.($hshHeader.DAT) -ne ''}).count

	# get the latest versions of 'DAT Version (VirusScan Enterprise)'
	$strLatest = ($arrCsv.($hshHeader.DAT) | measure -Maximum).Maximum

	# get the number of compliant computers (running the two (on Wednesday the three) last versions of 'DAT Version (VirusScan Enterprise)'
	$arrValues = $arrCsv | sort $hshHeader.DAT -Unique  | ?{$_ –ne ''}
	
	$arrCsv | ?{$_.($hshHeader.DAT) -in ($arrValues) | select -Last $intLastRec | ?{$_ –ne ''}}.count
	$hshResult['VirusScan¥Compliant ¥(8963 & 8968)'] = ($arrCsv | ?{$_.($hshHeader.DAT) -in ($arrValues)} ).count

	# get the number of non compliant VirusScan computers
	$hshResult['VirusScan¥Non Compliant'] = ($arrCsv | ?{!($_.($hshHeader.DAT) -in ($arrValues)) -and ($_ -ne '')}).count
	$arrCsv | ?{!($_.($hshHeader.DAT) -in ($arrValues)) -and ($_.($hshHeader.DAT) -ne '')} | %{
		$arrNC += "{0};{1};{2};{3}" -f 'VirusScan Non Compliant', $_.'System Name', $_.'IP Address', $_.'Operating System'
		}

	# get the number of computers unable to retrieve DAT
	$hshResult['VirusScan¥Unable To retrieve DAT'] = ($arrCsv.($hshHeader.DAT) | ?{!($_ -in ($arrValues)) -and ($_ -ne '')} ).count
#	$arrCsv | ?{$_.($hshHeader.DAT) –eq ''}| %{
#		if ($_.($hshHeader.DAT) -eq '') {
#			$hshResult['VirusScan¥Unable To retrieve DAT'] += 1
#			}
#		}

	set-XLData $strDestXL 'Data' $hshResult
	
	
# copy the file to the backup location
	$strFolder = "{0:yyyy}_{0:MM}" -f $dtToday
	$strDest = "\\beprod01.eoc.net\offshare\Infrastructure Evolution\001 WSS Group\Windows Support\02. Services\AntiVirus\1. Control\McAfee53x_SIP_Backup_File\PROD\{0}" -f $strFolder
	
	# Creating the folder if it does not yet exist
	if (!(Test-Path $strDest)) {
		New-Item -ItemType Directory -Force -Path $strDest
		}

	Set-Location $strDest
	$strNewFileName = "{0}\{1:dd}_{2}" -f $strFolder, $dtToday, $strWorkBook.Substring($strWorkBook.LastIndexOf('\') + 1)
	$strNewFileName
#		Copy-Item $strWorkBook $strNewFileName	
	}
else {
	'ePO server unreachable'
	}
	