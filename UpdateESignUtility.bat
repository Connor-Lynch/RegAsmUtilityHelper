set esignUtilityPath=%1
set utilityPath=%esignUtilityPath%\ESign.Utility.dll
echo %utilityPath%

"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\gacutil.exe" -u ESign.Utility
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" -u %utilityPath% /tlb %utilityPath%/codebase

"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\gacutil.exe" -i %utilityPath%
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" %utilityPath% /tlb %utilityPath%/codebase