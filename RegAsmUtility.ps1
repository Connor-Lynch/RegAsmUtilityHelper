using namespace System.Management.Automation.Host

param(
    [string]$esignUtilityPath = $home + "\source\repos\Document E-Signing\ESign.Utility\bin\Debug",
    [switch]$verbose = $false
)

function New-Menu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Question
    )
    
    $esign = [ChoiceDescription]::new('&ESign', 'Update the ESign.Utility.dll and ESign.Utility.tlb: ESign')
    $newtonsoft = [ChoiceDescription]::new('&Newtonsoft', 'Load the Newtonsoft.Json.dll and Newtonsoft.Json.tlb: Newtonsoft')
    $exit = [ChoiceDescription]::new('&Quit', 'Exit RegAsm.exe Utility: Quit')

    $options = [ChoiceDescription[]]($esign, $newtonsoft, $exit)

    $Title = "================= " + $Title + " ================="
    return $host.ui.PromptForChoice($Title, $Question, $options, 0)
}

function Close-Excel {
    $excel = Get-Process "excel" -ErrorAction SilentlyContinue
    if ($excel) {
        Write-Host "Closing Excel" -foregroundcolor Cyan
        $excel.CloseMainWindow() | Out-Null
        Start-Sleep 5
        $excel | Stop-Process -Force
    }
}

function Open-Excel {
    Write-Host "Starting Excel" -foregroundcolor Green
    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $true
    $curDir = Get-Location
    $fileName = $curDir.toString() + "\TestCOMVBA.xlsm"
    $workbook = $excel.Workbooks.Open($fileName)
}

do
{
    if(!$verbose) {
        cls
    }
    $result = New-Menu -Title 'RegAsm.exe Utility' -Question 'Please assembly to load:'

    switch ($result) {
        0 
        {
            Close-Excel

            Write-Host 'Starting Udate' -foregroundcolor Cyan
            $stack = .\UpdateESignUtility.bat $esignUtilityPath
            Write-Host 'Update Complete' -foregroundcolor Green

            Open-Excel
        } 
        1
        {
            Close-Excel

            Write-Host 'Starting' -foregroundcolor Cyan
            $stack = .\UpdateNewtonsoftUtility.bat $esignUtilityPath
            Write-Host 'Done' -foregroundcolor Green
        }
    }
}
until ($result -eq 2)

# regasm ESign.Utility.dll /tlb ESign.Utility.dll/codebase
# C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\x86\Debug\ESign.Utility.dll" /tlb "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\x86\Debug\ESign.Utility.dll"/codebase

# regasm "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\Debug\Newtonsoft.Json.dll" /tlb "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\Debug\Newtonsoft.Json.dll"/codebase
# regasm "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\Debug\ESign.Utility.dll" /tlb "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\Debug\ESign.Utility.dll"/codebase


# regasm "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\x86\Debug\Newtonsoft.Json.dll" /tlb "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\x86\Debug\Newtonsoft.Json.dll"/codebase
# regasm "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\x86\Debug\ESign.Utility.dll" /tlb "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\x86\Debug\ESign.Utility.dll"/codebase

# gacutil -i "C:\Users\ConnorLynch\source\repos\Document E-Signing\ESign.Utility\bin\Debug\ESign.Utility.dll"