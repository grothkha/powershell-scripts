####
# Description: Script to toggle monitor refresh rate between two hardcoded values.
# Author: Hannes Grothkopf
# TODOs: Better rate setting handling, make monitor ID configurable
# Info: 
# - https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/getting-wmi-objects--get-wmiobject-?view=powershell-6
# - https://superuser.com/questions/1184815/is-a-way-to-quickly-set-screen-refresh-rate
####

$refreshRate1 = 60
$refreshRate2 = 144
$win32_VideoController = Get-WmiObject "Win32_VideoController"

foreach ($objItem in $win32_VideoController) {
    $currentHorizontalResolution = $objItem.CurrentHorizontalResolution
    $currentVerticalResolution = $objItem.CurrentVerticalResolution
    $currentBitsPerPixel = $objItem.CurrentBitsPerPixel
    $currentRefreshRate = [int]$objItem.CurrentRefreshRate

    if($currentRefreshRate.Equals($refreshRate1)){
        Write-Host "Current refresh rate is $refreshRate1."
        $userInput = Read-Host -Prompt "Set rate to $refreshRate2? [y/ANY]"
        if($userInput.Equals('y')){
            nircmdc.exe setdisplay $currentHorizontalResolution $currentVerticalResolution $currentBitsPerPixel $refreshRate2
        }
    }elseif($currentRefreshRate.Equals($refreshRate2)){
        Write-Host "Current refresh rate is $refreshRate2."
        $userInput = Read-Host -Prompt "Set rate to $refreshRate1? [y/ANY]"
        if($userInput.Equals('y')){
            nircmdc.exe setdisplay $currentHorizontalResolution $currentVerticalResolution $currentBitsPerPixel $refreshRate1
        }
    }else{
        Write-Host "Unknown refresh rate. Doing nothing."
    }
}