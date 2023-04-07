##
## Download NuGet.exe
##
$sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
$targetNugetExe = ".\nuget.exe"
Remove-Item .\Tools -Force -Recurse -ErrorAction Ignore
Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
Set-Alias nuget $targetNugetExe -Scope Global -Verbose

##
## Download Microsoft.Identity.Client
##
./nuget install Microsoft.Identity.Client -O .\Tools
md .\Tools\Microsoft.Identity.Client
md .\Tools\Microsoft.IdentityModel.Abstractions
$prtFolder = Get-ChildItem ./Tools | Where-Object {$_.Name -match 'Microsoft.Identity.Client.'}
$prtFolder2 = Get-ChildItem ./Tools | Where-Object {$_.Name -match 'Microsoft.IdentityModel.Abstractions.'}
move .\Tools\$prtFolder\lib\net45\*.* .\Tools\Microsoft.Identity.Client
move .\Tools\$prtFolder2\lib\net45\*.* .\Tools\Microsoft.IdentityModel.Abstractions
Remove-Item .\Tools\$prtFolder -Force -Recurse
Remove-Item .\Tools\$prtFolder2 -Force -Recurse

##
## Remove NuGet.exe
##
Remove-Item nuget.exe
