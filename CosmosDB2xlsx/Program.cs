using System.CommandLine;
using Microsoft.Azure.Cosmos;
using ClosedXML.Excel;
using System.Text.Json;

namespace CosmosDB2xlsx;

class Program
{
    static async Task<int> Main(string[] args)
    {
        var connectionStringOption = new Option<string>(
            name: "--connection-string",
            description: "Cosmos DB connection string")
        {
            IsRequired = true
        };
        connectionStringOption.AddAlias("-c");

        var containersOption = new Option<string[]>(
            name: "--containers",
            description: "List of container names to export (default: all containers)")
        {
            AllowMultipleArgumentsPerToken = true
        };
        containersOption.AddAlias("-t");

        var outputDirOption = new Option<string>(
            name: "--output",
            description: "Output directory for XLSX files (default: current directory)",
            getDefaultValue: () => Directory.GetCurrentDirectory());
        outputDirOption.AddAlias("-o");

        var databaseOption = new Option<string>(
            name: "--database",
            description: "Database name (required)")
        {
            IsRequired = true
        };
        databaseOption.AddAlias("-d");

        var rootCommand = new RootCommand("Export data from Cosmos DB to XLSX files")
        {
            connectionStringOption,
            databaseOption,
            containersOption,
            outputDirOption
        };

        rootCommand.SetHandler(async (connectionString, database, containers, outputDir) =>
        {
            await ExportCosmosDbToXlsx(connectionString, database, containers, outputDir);
        }, connectionStringOption, databaseOption, containersOption, outputDirOption);

        return await rootCommand.InvokeAsync(args);
    }

    static async Task ExportCosmosDbToXlsx(string connectionString, string databaseName, string[]? containerNames, string outputDir)
    {
        try
        {
            Console.WriteLine("Connecting to Cosmos DB...");
            using var cosmosClient = new CosmosClient(connectionString);
            
            var database = cosmosClient.GetDatabase(databaseName);
            
            // Get list of containers to export
            List<string> containersToExport;
            if (containerNames != null && containerNames.Length > 0)
            {
                containersToExport = containerNames.ToList();
                Console.WriteLine($"Exporting specified containers: {string.Join(", ", containersToExport)}");
            }
            else
            {
                // Get all containers
                Console.WriteLine("Getting all containers from database...");
                containersToExport = new List<string>();
                using var containerIterator = database.GetContainerQueryIterator<ContainerProperties>();
                while (containerIterator.HasMoreResults)
                {
                    var response = await containerIterator.ReadNextAsync();
                    foreach (var containerProps in response)
                    {
                        containersToExport.Add(containerProps.Id);
                    }
                }
                Console.WriteLine($"Found {containersToExport.Count} containers to export");
            }

            // Create output directory if it doesn't exist
            Directory.CreateDirectory(outputDir);

            // Export each container
            foreach (var containerName in containersToExport)
            {
                try
                {
                    await ExportContainer(database, containerName, outputDir);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error exporting container '{containerName}': {ex.Message}");
                }
            }

            Console.WriteLine("Export completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            throw;
        }
    }

    static async Task ExportContainer(Database database, string containerName, string outputDir)
    {
        Console.WriteLine($"\nExporting container: {containerName}");
        
        var container = database.GetContainer(containerName);
        
        // Query all items in the container
        var query = new QueryDefinition("SELECT * FROM c");
        using var iterator = container.GetItemQueryIterator<JsonDocument>(query);
        
        var items = new List<JsonDocument>();
        while (iterator.HasMoreResults)
        {
            var response = await iterator.ReadNextAsync();
            items.AddRange(response);
            Console.WriteLine($"  Retrieved {items.Count} items so far...");
        }

        if (items.Count == 0)
        {
            Console.WriteLine($"  Container '{containerName}' is empty. Skipping.");
            return;
        }

        Console.WriteLine($"  Total items retrieved: {items.Count}");
        Console.WriteLine($"  Creating XLSX file...");

        // Create Excel workbook
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(containerName);

        // Extract all unique property names from all items
        var allProperties = new HashSet<string>();
        foreach (var item in items)
        {
            var jsonElement = item.RootElement;
            foreach (var property in jsonElement.EnumerateObject())
            {
                allProperties.Add(property.Name);
            }
        }

        var propertyList = allProperties.OrderBy(p => p).ToList();
        
        // Write headers
        for (int i = 0; i < propertyList.Count; i++)
        {
            worksheet.Cell(1, i + 1).Value = propertyList[i];
            worksheet.Cell(1, i + 1).Style.Font.Bold = true;
        }

        // Write data rows
        int row = 2;
        foreach (var item in items)
        {
            var jsonElement = item.RootElement;
            
            for (int col = 0; col < propertyList.Count; col++)
            {
                var propertyName = propertyList[col];
                if (jsonElement.TryGetProperty(propertyName, out var propertyValue))
                {
                    var cellValue = GetCellValue(propertyValue);
                    worksheet.Cell(row, col + 1).Value = cellValue;
                }
            }
            row++;
        }

        // Auto-fit columns
        worksheet.Columns().AdjustToContents();

        // Save the workbook
        var fileName = Path.Combine(outputDir, $"{containerName}.xlsx");
        workbook.SaveAs(fileName);
        
        Console.WriteLine($"  Saved to: {fileName}");
    }

    static string GetCellValue(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.String:
                return element.GetString() ?? "";
            case JsonValueKind.Number:
                return element.ToString();
            case JsonValueKind.True:
                return "true";
            case JsonValueKind.False:
                return "false";
            case JsonValueKind.Null:
                return "";
            case JsonValueKind.Array:
            case JsonValueKind.Object:
                return element.ToString();
            default:
                return element.ToString();
        }
    }
}
