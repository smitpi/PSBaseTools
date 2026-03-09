<#PSScriptInfo

.VERSION 0.1.0

.GUID 3c9dc69d-98ff-46a7-ae8e-3aea6b7fafca

.AUTHOR Pierre Smit

.COMPANYNAME Private

.COPYRIGHT 

.TAGS ps

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
Moved to dedicated module

.PRIVATEDATA


#>

#Requires -Module ImportExcel
#Requires -Module PSWriteHTML

<#

.DESCRIPTION
 Creates a excel or html report 

#>


<#
.SYNOPSIS
Creates an Excel, HTML, or XML report.

.DESCRIPTION
Creates a report in Excel, HTML, or XML format. When exporting to Excel (or using Export 'All'), additional parameters for conditional text formatting and text wrapping are available.

.PARAMETER Export
Export the result to a report file. (Excel, HTML, or XML). If 'Excel' or 'All' is specified, Excel-specific options become available.

.PARAMETER InputObject
Data for the report.

.PARAMETER ExcelConditionalText
Add conditional text color to Excel cells. Only available if Export includes 'Excel' or 'All'.

.PARAMETER TextWrap
Wrap the text in the Excel report. Only available if Export includes 'Excel' or 'All'.

.PARAMETER ReportTitle
Title of the report.

.PARAMETER ReportPath
Where to save the report.

.PARAMETER OpenReportsFolder
Open the directory containing the created reports.

.EXAMPLE
[System.Collections.generic.List[PSObject]]$Conditions = @()
$Conditions.Add((New-ConditionalText -Text 'Warning' -ConditionalTextColor black -BackgroundColor Yellow -Range 'E:E'))
$Conditions.Add((New-ConditionalText -Text 'Error' -ConditionalTextColor black -BackgroundColor orange -Range 'E:E'))
$Conditions.Add((New-ConditionalText -Text 'Critical' -ConditionalTextColor white -BackgroundColor Red -Range 'E:E'))

Write-PSReports -InputObject $report -ReportTitle 'Windows Events' -Export All -ReportPath C:\temp -ExcelConditionalText $Conditions

#>
function Write-PSReports {
	[CmdletBinding(HelpURI = 'https://smitpi.github.io/PSBaseTools/Write-PSReports')]
	[OutputType([System.Object[]])]
	param(
		[Parameter(Position = 0, Mandatory)]
		[PSCustomObject]$InputObject,

		[Parameter(Position = 1, Mandatory)]
		[string]$ReportTitle,

		[ValidateSet('Excel', 'HTML', 'XML')]
		[string[]]$Export,

		[ValidateScript( { if (Test-Path $_) { $true }
				else { New-Item -Path $_ -ItemType Directory -Force | Out-Null; $true }
			})]
		[System.IO.DirectoryInfo]$ReportPath = 'C:\Temp',
		[switch]$OpenReportsFolder
	)

	dynamicparam {
		$paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

		if ($PSBoundParameters.ContainsKey('Export') -and ($PSBoundParameters['Export'] -contains 'Excel' -or $PSBoundParameters['Export'] -contains 'All')) {
			$attrCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
			$attrCollection.Add((New-Object System.Management.Automation.ParameterAttribute))
			$paramDictionary.Add('ExcelConditionalText', (New-Object System.Management.Automation.RuntimeDefinedParameter('ExcelConditionalText', [PSCustomObject], $attrCollection)))

			$attrCollection2 = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
			$attrCollection2.Add((New-Object System.Management.Automation.ParameterAttribute))
			$paramDictionary.Add('TextWrap', (New-Object System.Management.Automation.RuntimeDefinedParameter('TextWrap', [switch], $attrCollection2)))
		}
		return $paramDictionary
	}

	Write-Verbose "[$(Get-Date -Format HH:mm:ss) BEGIN] Starting $($myinvocation.mycommand)"

	Write-Verbose "[$(Get-Date -Format HH:mm:ss) PROCESS] Checking Members"
	$MemberCheck = ($InputObject.psobject.members | Where-Object { $_.MemberType -like 'NoteProperty' }).Name
	if (-not($MemberCheck)) {
		Write-Verbose "[$(Get-Date -Format HH:mm:ss) PROCESS] Creating custom object"
		$ToReport = [PSCustomObject]@{
			$($ReportTitle) = $InputObject
		}# PSObject
	} else {
		Write-Verbose "[$(Get-Date -Format HH:mm:ss) PROCESS] Using InputObject as is"
		$ToReport = $InputObject
	}

	# Retrieve dynamic parameters if present
	$ExcelConditionalText = $null
	$TextWrap = $false
	if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('ExcelConditionalText')) {
		$ExcelConditionalText = $PSCmdlet.MyInvocation.BoundParameters['ExcelConditionalText']
	}
	if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('TextWrap')) {
		$TextWrap = $PSCmdlet.MyInvocation.BoundParameters['TextWrap']
	}

	Write-Verbose "[$(Get-Date -Format HH:mm:ss) PROCESS] Rechecking Members"
	$Members = ($ToReport.psobject.members | Where-Object { $_.MemberType -like 'NoteProperty' }).Name

	if (($PSBoundParameters['Export'] -contains 'Excel') -or ($PSBoundParameters['Export'] -contains 'All')) {
		Write-Verbose "[$(Get-Date -Format HH:mm:ss) END] Creating Excel Report "
		$ExcelOptions = @{
			Path              = $(Join-Path -Path $ReportPath -ChildPath "\$($ReportTitle.Replace(' ','_'))_$(Get-Date -Format yyyy.MM.dd-HH.mm).xlsx")
			AutoSize          = $True
			AutoFilter        = $True
			TitleBold         = $True
			TitleSize         = '28'
			TitleFillPattern  = 'LightTrellis'
			TableStyle        = 'Light20'
			FreezeTopRow      = $True
			FreezePane        = '3'
			FreezeFirstColumn = $True
			MaxAutoSizeRows   = 50
		}

		if ($ExcelConditionalText) {
			$ExcelOptions.Add('ConditionalText', $ExcelConditionalText)
		}

		foreach ($Member in $Members) {
			if ($ToReport.$member) { $ToReport.$Member | Export-Excel -Title $Member -WorksheetName $Member @ExcelOptions }
		}
		if ($TextWrap) {
			$excel = Open-ExcelPackage -Path $ExcelOptions.Path
			foreach ($Member in $Members) {
				if ($ToReport.$member) { 
					$WorkSheet = $excel.Workbook.Worksheets[$member]
					$range = $WorkSheet.Dimension.address.Replace('A1', 'A2')
					Set-ExcelRange -Address $WorkSheet.Cells[$($range)] -WrapText -VerticalAlignment Center
				}
			}
			Close-ExcelPackage $excel
		}
	}

	if (($PSBoundParameters['Export'] -contains 'HTML') -or ($PSBoundParameters['Export'] -contains 'All')) {
		Write-Verbose "[$(Get-Date -Format HH:mm:ss) END] Creating HTML Report "
		$TableSettings = @{
			Style           = 'cell-border'
			TextWhenNoData  = 'No Data to display here'
			Buttons         = 'searchBuilder', 'pdfHtml5', 'excelHtml5'
			FixedHeader     = $true
			HideFooter      = $true
			SearchHighlight = $true
			PagingStyle     = 'full'
			PagingLength    = 100
			AutoSize        = $true
			ScrollX         = $true
			ScrollCollapse  = $true
			ScrollY         = $true
			DisablePaging   = $true
		}
		$SectionSettings = @{
			BackgroundColor       = 'grey'
			CanCollapse           = $true
			HeaderBackGroundColor = '#2b1200'
			HeaderTextAlignment   = 'center'
			HeaderTextColor       = '#f37000'
			HeaderTextSize        = '15'
			BorderRadius          = '20px'
		}
		$TableSectionSettings = @{
			BackgroundColor       = 'white'
			CanCollapse           = $true
			HeaderBackGroundColor = '#f37000'
			HeaderTextAlignment   = 'center'
			HeaderTextColor       = '#2b1200'
			HeaderTextSize        = '15'
		}
		$TabSettings = @{
			TextTransform = 'uppercase'
			IconBrands    = 'mix'
			TextSize      = '16' 
			TextColor     = '#00203F'
			IconSize      = '16'
			IconColor     = '#00203F'
		}
		$HeadingText = "$($ReportTitle) [$(Get-Date -Format dd) $(Get-Date -Format MMMM) $(Get-Date -Format yyyy) $(Get-Date -Format HH:mm)]"
		New-HTML -TitleText $($ReportTitle) -FilePath $(Join-Path -Path $ReportPath -ChildPath "\$($ReportTitle.Replace(' ','_'))_$(Get-Date -Format yyyy.MM.dd-HH.mm).html") {
			New-HTMLHeader {
				New-HTMLText -FontSize 20 -FontStyle normal -Color '#00203F' -Alignment left -Text $HeadingText
			}
			foreach ($Member in $Members) {
				if ($ToReport.$member) { New-HTMLTab -Name $Member @TabSettings -HtmlData { New-HTMLSection @TableSectionSettings { New-HTMLTable -DataTable $ToReport.$Member @TableSettings } } }
			}
		}
	}
	if (($PSBoundParameters['Export'] -contains 'XML') -or ($PSBoundParameters['Export'] -contains 'All')) {
		$Path = $(Join-Path -Path $ReportPath -ChildPath "\$($ReportTitle.Replace(' ','_'))_$(Get-Date -Format yyyy.MM.dd-HH.mm).xml")
		$InputObject | Export-Clixml -Depth 20 -Path $Path -Force -NoClobber
	}

	if ($OpenReportsFolder) { 
		Write-Verbose "[$(Get-Date -Format HH:mm:ss) END] Opening Folder"
		Start-Process -FilePath explorer.exe -ArgumentList $($ReportPath) 
	}
	Write-Verbose "[$(Get-Date -Format HH:mm:ss) END] Done"
} #end Function

