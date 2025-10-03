# CosmosDB2xlsx

A .NET command-line tool to export data from Azure Cosmos DB to Excel (XLSX) files.

## Features

- Export data from Cosmos DB containers to XLSX format
- Support for connection string authentication
- Export all containers or specify specific containers
- Automatic handling of dynamic schemas
- Each container exported to a separate XLSX file

## Installation

### Prerequisites

- [.NET 9.0 SDK](https://dotnet.microsoft.com/download) or later

### Build from Source

```bash
git clone https://github.com/frkim/CosmosDB2xlsx.git
cd CosmosDB2xlsx/CosmosDB2xlsx
dotnet build
```

## Usage

### Basic Syntax

```bash
dotnet run -- --connection-string "<your-connection-string>" --database "<database-name>" [options]
```

### Options

| Option | Alias | Description | Required |
|--------|-------|-------------|----------|
| `--connection-string` | `-c` | Cosmos DB connection string | Yes |
| `--database` | `-d` | Database name | Yes |
| `--containers` | `-t` | List of container names to export (space-separated) | No |
| `--output` | `-o` | Output directory for XLSX files | No (default: current directory) |
| `--help` | `-h` | Show help information | No |

### Examples

#### Export all containers from a database

```bash
dotnet run -- -c "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" -d "MyDatabase"
```

#### Export specific containers

```bash
dotnet run -- -c "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" -d "MyDatabase" -t "Users" "Orders" "Products"
```

#### Export to a specific directory

```bash
dotnet run -- -c "AccountEndpoint=https://your-account.documents.azure.com:443/;AccountKey=your-key==" -d "MyDatabase" -o "./exports"
```

## Output

- Each container is exported to a separate XLSX file named `<ContainerName>.xlsx`
- All documents in a container are exported to a single worksheet
- Column headers are automatically generated from document properties
- Dynamic schemas are supported - all unique properties across all documents are included as columns
- Complex objects and arrays are serialized as JSON strings

## Requirements

- Azure Cosmos DB account with valid credentials
- Network connectivity to the Cosmos DB endpoint
- Read permissions on the target database and containers

## NuGet Packages

This tool uses the following packages:

- [Microsoft.Azure.Cosmos](https://www.nuget.org/packages/Microsoft.Azure.Cosmos/) - Cosmos DB SDK
- [ClosedXML](https://www.nuget.org/packages/ClosedXML/) - Excel file generation
- [System.CommandLine](https://www.nuget.org/packages/System.CommandLine/) - Command-line parsing
- [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json/) - JSON serialization

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
