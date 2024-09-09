FROM mcr.microsoft.com/azure-functions/dotnet-isolated:4-dotnet-isolated8.0 AS base
WORKDIR /home/site/wwwroot
EXPOSE 80

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["PDFFunctions.csproj", "."]
RUN dotnet restore "./PDFFunctions.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "./PDFFunctions.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./PDFFunctions.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /home/site/wwwroot
COPY --from=publish /app/publish /home/site/wwwroot
ENV AzureWebJobsScriptRoot=/home/site/wwwroot
RUN mkdir -p /usr/share/man/man1
RUN apt-get update && apt-get upgrade -y
RUN apt-get -y install default-jre-headless libreoffice

ENTRYPOINT ["dotnet", "PDFFunctions.dll"]