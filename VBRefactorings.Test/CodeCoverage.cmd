echo %time%
Rem Set MANUAL_TESTS=
Rem dotnet test --collect:"XPlat Code Coverage" --settings coverletArgs.runsettings
Rem JSON
dotnet test ./VBRefactorings.Test.vbproj /p:CollectCoverage=true /p:CoverletOutputFormat=json /p:CoverletOutput=./TestResults/LastRun/Coverage.json /p:Exclude=\"[coverlet.*]*,[*]Coverlet.Core*,[xunit*]*,[Microsoft.DotNet.XUnitExtensions]*,[VBMsgBox]*,[HashLibrary]*"
Rem Opencover
Rem dotnet test ./VBRefactorings.Test.vbproj /p:CollectCoverage=true /p:CoverletOutputFormat=Opencover /p:CoverletOutput=./TestResults/LastRun/Coverage.Opencover /p:Exclude=\"[coverlet.*]*,[*]Coverlet.Core*,[xunit*]*,[Microsoft.DotNet.XUnitExtensions]*,[VBMsgBox]*,[HashLibrary]*"
Rem Cobertura
Rem dotnet test ./VBRefactorings.Test.vbproj /p:CollectCoverage=true /p:CoverletOutputFormat=Cobertura /p:CoverletOutput=./TestResults/LastRun/Coverage.Cobertura  /p:Exclude=\"[coverlet.*]*,[*]Coverlet.Core*,[xunit*]*,[Microsoft.DotNet.XUnitExtensions]*,[VBMsgBox]*,[HashLibrary]*"
