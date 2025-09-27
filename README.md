Itenium.ExcelCreator
====================

## WebApi Deployment

Configure `backend/nuget.config` or create `.env` or set the
`Nuget_CustomFeed*` environment variables.

```sh
docker-compose up -d --build
```

### Link

In the `docker-compose` of the project that is going to use it:

```
networks:
  forge-excel_excelnet:
    external: true
```

In the project, refer to the excel-creator:

```text
http://excel-creator:5000/api/Excel
```


## HttpClient Nuget

(this one doesn't exist yet)

```sh
dotnet add package Itenium.ExcelCreator
```
