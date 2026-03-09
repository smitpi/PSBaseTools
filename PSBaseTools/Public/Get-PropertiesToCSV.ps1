<#PSScriptInfo

.VERSION 0.1.0

.GUID 5ab117c4-f29f-4b50-8fd0-c783240ab40d

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
<#

.DESCRIPTION
Get member data of an object. Use it to create other PSObjects.

#>


<#
.SYNOPSIS
Get member data of an object. Use it to create other PSObjects.

.DESCRIPTION
Get member data of an object. Use it to create other PSObjects.

.PARAMETER Data
Parameter description

.EXAMPLE
Show-ObjectPropertiesAsCSV -data $data

#>
function Show-ObjectPropertiesAsCSV {
    [Cmdletbinding(HelpURI = 'https://smitpi.github.io/PSToolKit/Show-ObjectPropertiesAsCSV')]

    param (
        [parameter( ValueFromPipeline = $True )]
        [object[]]$Data)

    process {
        $data | Get-Member -MemberType NoteProperty, Property | Sort-Object | ForEach-Object { $_.name } | Join-String -Separator ','
    }
} #end Function

