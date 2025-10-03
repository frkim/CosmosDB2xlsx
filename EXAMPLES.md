# CosmosDB2xlsx Usage Examples

This document provides detailed examples for using the CosmosDB2xlsx tool.

## Prerequisites

Before using this tool, you need:

1. An Azure Cosmos DB account
2. A connection string with read permissions
3. .NET 9.0 SDK or the published executable

## Getting Your Connection String

You can get your Cosmos DB connection string from:
- Azure Portal: Navigate to your Cosmos DB account → Keys → Primary Connection String
- Azure CLI: `az cosmosdb keys list --name <account-name> --resource-group <resource-group> --type connection-strings`

## Basic Usage Examples

### Example 1: Export All Containers

Export all containers from a database to the current directory:

```bash
dotnet run --project CosmosDB2xlsx/CosmosDB2xlsx.csproj -- \
  --connection-string "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" \
  --database "MyDatabase"
```

### Example 2: Export Specific Containers

Export only the "Users" and "Orders" containers:

```bash
dotnet run --project CosmosDB2xlsx/CosmosDB2xlsx.csproj -- \
  --connection-string "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" \
  --database "MyDatabase" \
  --containers "Users" "Orders"
```

### Example 3: Export to Specific Directory

Export all containers to a specific output directory:

```bash
dotnet run --project CosmosDB2xlsx/CosmosDB2xlsx.csproj -- \
  --connection-string "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" \
  --database "MyDatabase" \
  --output "./exports/$(date +%Y%m%d)"
```

### Example 4: Using Short Aliases

Use short aliases for cleaner command lines:

```bash
dotnet run --project CosmosDB2xlsx/CosmosDB2xlsx.csproj -- \
  -c "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" \
  -d "MyDatabase" \
  -t "Products" "Categories" \
  -o "./exports"
```

## Using Environment Variables

For better security, you can store the connection string in an environment variable:

### Linux/macOS

```bash
export COSMOS_CONNECTION_STRING="AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key=="

dotnet run --project CosmosDB2xlsx/CosmosDB2xlsx.csproj -- \
  -c "$COSMOS_CONNECTION_STRING" \
  -d "MyDatabase"
```

### Windows (PowerShell)

```powershell
$env:COSMOS_CONNECTION_STRING = "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key=="

dotnet run --project CosmosDB2xlsx/CosmosDB2xlsx.csproj -- `
  -c "$env:COSMOS_CONNECTION_STRING" `
  -d "MyDatabase"
```

## Using Published Executable

After publishing the application, you can use it without the .NET SDK:

### Windows

```cmd
CosmosDB2xlsx.exe ^
  -c "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" ^
  -d "MyDatabase" ^
  -o "C:\exports"
```

### Linux/macOS

```bash
./CosmosDB2xlsx \
  -c "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" \
  -d "MyDatabase" \
  -o "/home/user/exports"
```

## Batch Processing Multiple Databases

You can create a script to export multiple databases:

### Bash Script

```bash
#!/bin/bash

CONNECTION_STRING="AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key=="
DATABASES=("Database1" "Database2" "Database3")
OUTPUT_BASE="./exports"

for db in "${DATABASES[@]}"; do
  echo "Exporting database: $db"
  ./CosmosDB2xlsx \
    -c "$CONNECTION_STRING" \
    -d "$db" \
    -o "$OUTPUT_BASE/$db"
done
```

### PowerShell Script

```powershell
$ConnectionString = "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key=="
$Databases = @("Database1", "Database2", "Database3")
$OutputBase = ".\exports"

foreach ($db in $Databases) {
    Write-Host "Exporting database: $db"
    .\CosmosDB2xlsx.exe `
        -c $ConnectionString `
        -d $db `
        -o "$OutputBase\$db"
}
```

## Output Format

The tool creates one XLSX file per container with the following characteristics:

- **File name**: `<ContainerName>.xlsx`
- **Worksheet name**: Same as the container name
- **Headers**: First row contains all unique property names across all documents
- **Data**: Each subsequent row represents one document
- **Complex types**: Objects and arrays are serialized as JSON strings
- **Missing properties**: Empty cells for properties not present in a document

## Troubleshooting

### Connection Errors

If you see connection errors:
1. Verify your connection string is correct
2. Check that your IP address is allowed in Cosmos DB firewall rules
3. Ensure you have network connectivity to Azure

### Permission Errors

If you see permission errors:
1. Verify the account key has read permissions
2. Check that the database exists
3. Ensure container names are spelled correctly

### Large Datasets

For very large datasets:
1. Consider exporting specific containers instead of all
2. The tool loads all data into memory before writing to Excel
3. Excel has a limit of ~1,048,576 rows per worksheet

## Performance Tips

1. **Export during off-peak hours** to minimize impact on production workloads
2. **Use specific containers** instead of exporting all to reduce time and resources
3. **Consider partitioning** exports by date ranges if you have time-series data
4. **Increase RU/s temporarily** if you need faster exports and can afford the cost
