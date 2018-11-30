####
# Description: Script to toggle monitor refresh rate between two hardcoded values.
# Author: Hannes Grothkopf
# TODOs: Better rate setting handling, make monitor ID configurable
# Info: 
# - https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/getting-wmi-objects--get-wmiobject-?view=powershell-6
# - https://superuser.com/questions/1184815/is-a-way-to-quickly-set-screen-refresh-rate
####

$RefreshRate1 = 60
$RefreshRate2 = 144
$win32_VideoController = Get-WmiObject "Win32_VideoController"

foreach ($objItem in $win32_VideoController) {
    $CurrentHorizontalResolution = $objItem.CurrentHorizontalResolution
    $CurrentVerticalResolution = $objItem.CurrentVerticalResolution
    $CurrentBitsPerPixel = $objItem.CurrentBitsPerPixel
    $CurrentRefreshRate = [int]$objItem.CurrentRefreshRate

    if($CurrentRefreshRate.Equals($RefreshRate1)){
        Write-Host "Refresh rate was $RefreshRate1 is now $RefreshRate2."
        nircmdc.exe setdisplay $CurrentHorizontalResolution $CurrentVerticalResolution $CurrentBitsPerPixel $RefreshRate2
    }elseif($CurrentRefreshRate.Equals($RefreshRate2)){
        Write-Host "Refresh rate was $RefreshRate2 is now $RefreshRate1."
        nircmdc.exe setdisplay $CurrentHorizontalResolution $CurrentVerticalResolution $CurrentBitsPerPixel $RefreshRate1
    }else{
        Write-Host "Unknown refresh rate. Doing nothing."
    }
}

Read-Host -Prompt "Press Enter to exit"