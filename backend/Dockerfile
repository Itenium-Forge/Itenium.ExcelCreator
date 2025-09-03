FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src

# Copy files & build
COPY *.sln .
COPY nuget.config ./
COPY **/*.csproj ./
RUN for file in $(ls *.csproj); do mkdir -p ./${file%.*}/ && mv $file ./${file%.*}/; done

COPY . .
RUN dotnet restore
RUN dotnet publish ./Itenium.ExcelCreator.WebApi/Itenium.ExcelCreator.WebApi.csproj -c Release -o /app


FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS runtime
WORKDIR /app

# RUN apt-get update && apt-get install -y \
#     some-package \
#     && rm -rf /var/lib/apt/lists/*

COPY --from=build /app .
# USER $APP_UID
ENTRYPOINT ["dotnet", "Itenium.ExcelCreator.WebApi.dll"]
