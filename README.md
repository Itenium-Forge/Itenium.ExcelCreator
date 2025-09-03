Itenium.ExcelCreator
====================

## HttpClient Nuget

```sh
dotnet add package Itenium.ExcelCreator
```

## WebApi Deployment

```sh
cd backend
docker build -t itenium-excel-creator .
docker run -d -p 8080:5000 -e ASPNETCORE_URLS="http://*:5000" -e DOTNET_ENVIRONMENT=Development --name excel-creator itenium-excel-creator
```
