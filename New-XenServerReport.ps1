<#
.Synopsis
   Generate XenServer documentation report
.DESCRIPTION
   Generate XenServer documentation report. This script requires PSCribo module which can be downloaded here: https://github.com/iainbrighton/PScribo
   The report can be set for word format or HTML format and the standard placement will be on the desktop.
.EXAMPLE
   New-XenServerReport -PoolMasterIP "10.10.10.10","10.10.10.20" -UserName "" -Password ""
#>
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
                If ($AvailablePatches.count -eq 0) {
                    $AvailablePatchesInfo = "No patches available"
                } else {
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
                $XenBondMasters = (Get-Xenbond).master | Sort-Object opaque_ref
                $NICStatusInfo = @()
                foreach ($XenBondMaster in $XenBondMasters) {
                    $XenPif = Get-XenPIF -opaque_ref $XenBondMaster
                    $BondName = $XenPif.device 
                    $XenHost = (Get-XenPIF -opaque_ref $XenBondMaster).host
                    $XenHostName = (Get-XenHost -opaque_ref $XenHost).hostname
                    $XenPIFMasterOf = (Get-XenPIF -opaque_ref $XenBondMaster).bond_master_of
                    $Slaves = (Get-XenBond -opaque_ref ($XenPIFMasterOf).opaque_ref).slaves
                    Foreach ($Slave in $Slaves)  {
                        $SlavePif = Get-XenPif -opaque_ref $Slave.opaque_ref
                        $Metrics = ($SlavePif).metrics
                        $LinkSpeed = (Get-XenPIFMetrics -opaque_ref $Metrics.opaque_ref).speed
                        If ($LinkSpeed -gt 0) {
                            $LinkState = "Up"
                        } else {
                            $LinkState = "Down"
                        }
                        $Device = $SlavePif.device
                        $MAC = $SlavePif.MAC
                        $NICStatus = New-Object psobject
                        $NICStatus | Add-Member -MemberType NoteProperty -Name "Host name" -Value $XenHostName
                        $NICStatus | Add-Member -MemberType NoteProperty -Name "Bond name" -Value $BondName
                        $NICStatus | Add-Member -MemberType NoteProperty -Name "Device name" -Value $Device
                        $NICStatus | Add-Member -MemberType NoteProperty -Name "MAC address" -Value $MAC
                        $NICStatus | Add-Member -MemberType NoteProperty -Name "Link state" -Value $LinkState
                        $NICStatusInfo +=  $NICStatus
                    }
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
                If (($Compare).length -eq 0) {
                    Paragraph "XenServer is up to date" -Style Normal
                } else {
                    $Compare | Table -Columns Hotfix -Width 25                
                }
                Paragraph "Xenserver host overview" -Style Heading2    
                $XSManagementIPResult | Table -Columns 'name_label','address','edition' -Headers 'Hostname','Management IP','License edition' -width 75
                Paragraph "VLANs present in pool" -Style Heading2
                $VLANResult | Table -Columns 'name_label','name_description','bridge' -Headers 'Name','Description','Bridge' -width 75
                Paragraph "Network information" -Style Heading2
                $NICStatusInfo | Sort-Object "Host name", "Device name" | Table -Width 75
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

# SIG # Begin signature block
# MIINFAYJKoZIhvcNAQcCoIINBTCCDQECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUsDc3Id5llKTS9ZonsoGTTing
# 2p6gggpWMIIFHjCCBAagAwIBAgIQDXS9akkSjo7r759FxeBS1zANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDIwNTAwMDAwMFoXDTE5MDIx
# MzEyMDAwMFowWzELMAkGA1UEBhMCREsxFDASBgNVBAcTC1NrYW5kZXJib3JnMRow
# GAYDVQQKExFNYXJ0aW4gVGhlcmtlbHNlbjEaMBgGA1UEAxMRTWFydGluIFRoZXJr
# ZWxzZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCitzrGHQMHiYmK
# nIJg1k6GY+RVIB88GbrFRXCHdZYZeBjuLVpMUNxAFAo4fivAjOIXp9gdqORaoWSO
# Rqs4GqttuqViM0cpGVI8hoVw5yfAxwWk9bX9/e8P5yzV3rUJF3wVHkO80oSWeOTW
# geIvABuyRsoA6lQUs+WTEw1BgR6X1d5dLN9mJjhTWUqB41lBliGE204IbhHrvfsS
# czfKaXovg+MnKK51cnkqav+mdUsS0IyDE/18WlOMroIi8NpmJcyZp1ejCv78SEcM
# PWJ7LJE9VX/iUpjMWZt1XY1XM3Hxcs6yF8A26gUzkKzMUG7kEOHuilMHh/Wh+c16
# jNvz0//hAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32
# ZXUOWDAdBgNVHQ4EFgQUfF54GE6E0FN6g8hk8fic5PoO9qowDgYDVR0PAQH/BAQD
# AgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Ax
# hi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUH
# MAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1
# cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEB
# CwUAA4IBAQDsfRHsWvsAZnRbk7FxBnkRWf67MyfyMrY00mAkg53qhasp75Fzs0Yb
# NxTDJVHiC+W9Inm/uuDI+UgjqLEOSL7R4co/gtRHznn93HKna5NuQ+AJdGSKWVpX
# owL4mKTMLu3DLS12gArnF/ozvbugzFb9EU5pspX1jxFU1UtX6rmIvRJCWi3q6FMJ
# jRx5QmkIerq2jlZs1ZAgmJheMah5OO+CMC7UFtD6vY7Rq/swdv48dHXrSOE4ZF2G
# cx8hccptp+v5+zPPBdEi1sVJ0GiE3ZHTcslfvC07VXsrR09wuTyF34xoazqubeWa
# kVhqS28oio2n/hBTT/lWCE25CbXhm1/1MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1
# U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQD
# ExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcN
# MjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2Vy
# dCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid
# 2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sj
# lOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjf
# DPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzL
# fnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR
# 93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckw
# EgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgw
# OqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIE
# MCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYI
# YIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1a
# JLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUP
# UbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1
# UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjF
# Emifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM
# 1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhs
# RDKyZqHnGKSaZFHvMYICKDCCAiQCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8G
# A1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQQIQ
# DXS9akkSjo7r759FxeBS1zAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAig
# AoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgEL
# MQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUBKqpaVfCK8Nw42R3PDko
# TS1/yq4wDQYJKoZIhvcNAQEBBQAEggEAedWxNxT9lTeemxEiZhIXLj2OhY6XZRhs
# TNfeWjPiYX9aUlsUEfy5/Qp8A3l3Z9j9rQhzSFaJAL7eZZTaARQjeMy5DKtVhB70
# V5We0rEEizUAg+lSZgcPRU/L8vPSgi1/MfAkgnPLVQDX6sug1E3dKv0WbkaPzq9N
# MRPESdhp+EgAnYK6TYRTkJL7eSCLWAIaMLdbG8T3TpnKquB1FbRbfSQBI47G78Qb
# KKbPa4TO0QqBkm7on8UGjh7XgdxznQsI3D45LgCNqg/v9khvCXIWXsj8F8/3NVO9
# 9jzaabBGfc1hFZETJkN/MBKJdO4oSoCmsUbJCFaQpVumc/cmf2Tdrg==
# SIG # End signature block
