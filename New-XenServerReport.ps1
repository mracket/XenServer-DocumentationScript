<#
.Synopsis
   Generate XenServer documentation report
.DESCRIPTION
   Generate XenServer documentation report. This script requires PSCribo module which can be downloaded here: https://github.com/iainbrighton/PScribo
   The report can be set for word format or HTML format and the standard placement will be on the desktop.
.EXAMPLE
   New-XenServerReport -PoolMasterIP "10.10.10.10","10.10.10.20" -UserName "" -Password ""
#>
Function New-XenServerReport {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $PoolMasterIP,
        [ValidateSet("Word","HTML","Text")] 
        [String]$ReportFormat, 

        # Param2 help description
        [string]$UserName,
        [string]$Password
    )
    Begin{
        Import-Module PScribo -Force
        Import-Module XenServerPSModule -Force
    }
    Process{
        $XenServerReport = Document -Name 'XenServer Report' {
            Foreach ($PoolMaster in $PoolMasterIP) {
                Connect-XenServer -Server $PoolMaster -UserName $UserName -Password $Password -SetDefaultSession -NoWarnCertificates
                
                $InstalledPatches = @()
                $XSManagementIPResult = @()
                $LUNResult = @()
                $LUNDtails = [ordered]@{}
                $LUNTest = @()
            
                $XenPoolName = (Get-XenPool).name_label
                Write-Host $PoolMaster
                if ((Get-XenHost).Count -gt 1) {
                    $XSVersion = (Get-XenHost).software_version[0].product_version
                } else {
                    $XSVersion = (Get-XenHost).software_version.product_version
                }                
                $XenPoolHAStatus = (Get-XenPool).ha_enabled
                $PoolPatches = Get-XenPoolPatch | Sort-Object name_label
                Foreach ($PoolPatch in $PoolPatches) {
                    $InstalledPatches += $PoolPatch
                }
                $url = "http://updates.xensource.com/XenServer/updates.xml"
                [xml]$xml = (new-object System.Net.WebClient).DownloadString($url)
                
                $AvailablePatches = $xml.patchdata.patches.patch | where {$_.'#comment' -like "*$XSVersion*"} | Sort-Object name-label
                
                $AvailablePatchesInfo = @()
                foreach ($AvailablePatch in $AvailablePatches) {
                    $AvailablePatchesNames = New-Object psobject
                    $AvailablePatchesNames | Add-Member -MemberType NoteProperty -Name "Hotfix" -Value $AvailablePatch.'name-label'
                    $AvailablePatchesInfo += $AvailablePatchesNames
                }
                $InstalledPatchesInfo = @()
                foreach ($InstalledPatch in $InstalledPatches) {
                    $InstalledPatchesNames = New-Object psobject
                    $InstalledPatchesNames | Add-Member -MemberType NoteProperty -Name "Hotfix" -Value $InstalledPatch.'name_label'
                    $InstalledPatchesInfo += $InstalledPatchesNames
                }
            
                $Compare = Compare-Object $AvailablePatchesInfo $InstalledPatchesInfo -Property hotfix 
                $Compare = $Compare | where {$_.sideindicator -like "<="}

                $Xenhosts = Get-XenHost | Sort-Object name_label
                foreach ($XenHost in $XenHosts) {
                    $XSManagementIPResult += Get-XenHost $XenHost
                }
                $LUNs = Get-XenSR | where {$_.type -ne "udev"} | Sort-Object name_label
                $StorageInfo = @()
                foreach ($LUN in $LUNs) {
                    $LUNResult += $LUN                               
                    if ($LUN.shared -eq "true") {
                        $LUNInfo = New-Object psobject
                        $LUNInfo | Add-Member -MemberType NoteProperty -Name "MultipathCapable" -Value ($LUN.sm_config).multipathable
                        $LUNInfo | Add-Member -MemberType NoteProperty -Name "DeviceSerial" -Value ($LUN.sm_config).devserial
                        $LUNInfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $LUN.name_label
                        $StorageInfo += $LUNInfo
                    }           
                }
                $VLANResult =  @() 
                $VLANs = Get-XenNetwork | where {$_.bridge -like "xapi*"} | Sort-Object name_label
                foreach ($VLAN in $VLANs) {
                        $VLANResult += $VLAN
                }
                $MultiPath = Get-XenHost | Select-Object name_label, other_config -ExpandProperty other_config | Sort-Object name_label
                $MultiPathInfo = @()
                foreach ($xenhost in $MultiPath) {
                    if ($xenhost.multipathing.length -eq 0) {
                        $XenHostMultipathStatus = "Disabled"
                    } else {
                        $XenHostMultipathStatus = "Enabled"
                    }
                    $MultipathStatus = New-Object psobject
                    $MultipathStatus | Add-Member -MemberType NoteProperty -Name "Name" -Value $xenhost.name_label
                    $MultipathStatus | Add-Member -MemberType NoteProperty -Name "Status" -Value $XenHostMultipathStatus
                
                $MultiPathInfo +=  $MultipathStatus 
                }
                Paragraph "Xenserver pool information" -Style Heading2
                Paragraph "Xenserver pool name: $XenPoolName"
                Paragraph "XenServer version: $XSVersion" 
                Paragraph "XenServer pool high availability enabled: $XenPoolHAStatus"
                Paragraph "Installed patches" -Style Heading2 
                $InstalledPatches | Table -Columns 'name_label','name_description' -Width 75
                Paragraph "Available patches" -Style Heading2
                $AvailablePatches | Table -Columns 'name-label','name-description' -Width 75
                Paragraph "Missing patches" -Style Heading2
                $Compare | Table -Columns Hotfix -Width 25
                Paragraph "Xenserver host overview" -Style Heading2    
                $XSManagementIPResult | Table -Columns 'name_label','address','edition' -Headers 'Hostname','Management IP','License edition' -width 75
                Paragraph "VLANs present in pool" -Style Heading2
                $VLANResult | Table -Columns 'name_label','name_description','bridge' -Headers 'Name','Description','Bridge' -width 75
                Paragraph "Xenserver shared storage" -Style Heading2 
                $LUNResult | Table -Columns 'name_label','physical_size','physical_utilisation','shared','virtual_allocation','type' -Width 75
                Paragraph "Xenserver shared storage extra information" -Style Heading3     
                $StorageInfo | Table -Width 75
                Paragraph "Multipath status" -Style Heading2
                $MultiPathInfo | Table -Columns 'Name','Status' -Width 25
                PageBreak
                Disconnect-XenServer
            }
        }
        $XenServerReport | Export-Document -Format $ReportFormat  -Path ~\Desktop     
    }
    End{
        
    }
}
New-XenServerReport -PoolMasterIP "10.10.10.10" -UserName "user" -Password "password" -ReportFormat html