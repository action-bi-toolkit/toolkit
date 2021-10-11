@ECHO OFF

PUSHD %~dp0

nuget install Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64 -ExcludeVersion -o ./NuGet
nuget install Microsoft.AnalysisServices.AdomdClient.retail.amd64 -ExcludeVersion -o ./NuGet
nuget install Microsoft.AnalysisServices.NetCore.retail.amd64 -ExcludeVersion -o ./NuGet
nuget install Microsoft.AnalysisServices.retail.amd64 -ExcludeVersion -o ./NuGet

POPD