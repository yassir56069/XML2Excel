using System;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using OfficeOpenXml;
using System.Data.SqlClient;

/// <summary>
/// Provides a robust XML to Excel conversion utility with batch and file watching capabilities
/// </summary>
class XmlToExcelConverter
{
    /// <summary>
    /// Establish SQL Connection
    /// by retrieving connection string from app.config
    /// </summary>
    /// <remarks> TODO: use secrets.json for connection string </remarks>
    private static string connectionString;

    /// <summary>
    /// Folder path where XML input files are located
    /// </summary>
    private readonly string watchFolder;

    /// <summary>
    /// Folder for processed xmls in batch runs, to avoid duplicate batch runs where possible
    /// </summary>
    private readonly string processedBatchFolder;

    /// <summary>
    /// Destination folder where converted Excel files will be saved
    /// </summary>
    private readonly string destinationFolder;

    /// <summary>
    /// Subfolder for initially found XML files
    /// </summary>
    private readonly string initialFilesFolder;

    /// <summary>
    /// Initializes a new instance of the XmlToExcelConverter with specified folder paths
    /// </summary>
    /// <param name="watchFolder">Directory to monitor for XML files</param>
    /// <param name="destinationFolder">Directory where converted files will be saved</param>
    public XmlToExcelConverter(string watchFolder, string destinationFolder, string processedBatchFolder)
    {
        // Establish SQL Connection
        connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;


        // Ensure all required directories exist
        this.watchFolder = watchFolder;
        this.destinationFolder = destinationFolder;
        this.processedBatchFolder = processedBatchFolder;
        this.initialFilesFolder = Path.Combine(destinationFolder, "InitialFiles");

        // Create directories if they don't exist
        Directory.CreateDirectory(watchFolder);
        Directory.CreateDirectory(destinationFolder);
        Directory.CreateDirectory(processedBatchFolder);
        Directory.CreateDirectory(initialFilesFolder);

        // Set EPPlus license context (required for EPPlus 5.0+)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    /// <summary>
    /// Converts existing XML files in the input folder to Excel
    /// </summary>
    /// <returns>Number of files converted</returns>
    public int ConvertXmlToExcelBatchHandling()
    {
        // Check for existing XML files
        string[] existingXmlFiles = Directory.GetFiles(watchFolder, "*.xml");
        int convertedFiles = 0;
        Console.WriteLine($"Opering in folder: {watchFolder}");

        if (existingXmlFiles.Length > 0)
        {
            Console.WriteLine($"Found {existingXmlFiles.Length} existing XML file(s).");

            // Convert existing files
            foreach (var file in existingXmlFiles)
            {
                try
                {
                    ConvertXmlToExcel(file, true);
                    convertedFiles++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting {Path.GetFileName(file)}: {ex.Message}");
                }
            }

            Console.WriteLine($"Converted {convertedFiles} file(s) in batch mode.");
        }
        else
        {
            Console.WriteLine("No XML files found in the input folder.");
        }

        return convertedFiles;
    }

    /// <summary>
    /// Starts continuous file system monitoring for new XML files
    /// </summary>
    public void ConvertXmlToExcelWatchHandling()
    {
        FileSystemWatcher watcher = new FileSystemWatcher(watchFolder);

        // Configure watcher settings
        watcher.Created += OnXmlFileCreated;
        watcher.Filter = "*.xml";
        watcher.EnableRaisingEvents = true;

        Console.WriteLine($"Watching folder: {watchFolder} for new XML files");
        Console.WriteLine($"Converted files will be saved to: {destinationFolder}");
        Console.WriteLine("Press 'Q' to quit the application.");

        // Keep the application running
        while (Console.ReadKey().Key != ConsoleKey.Q)
        {
            Thread.Sleep(1000);
        }
    }

    /// <summary>
    /// Event handler for newly created XML files in the watched directory
    /// </summary>
    /// <param name="sender">The source of the event</param>
    /// <param name="e">File system event arguments</param>
    private void OnXmlFileCreated(object sender, FileSystemEventArgs e)
    {
        try
        {
            // Wait to ensure file is fully written
            Thread.Sleep(500);

            // Convert the XML to Excel (not from initial files)
            ConvertXmlToExcel(e.FullPath, false);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file {e.Name}: {ex.Message}");
        }
    }

    /// <summary>
    /// Converts an XML file to Excel, with optional handling for initial files
    /// </summary>
    /// <param name="xmlFilePath">Full path to the XML file</param>
    /// <param name="isInitialFile">Indicates if the file is from the initial batch</param>
    private void ConvertXmlToExcel(string xmlFilePath, bool isInitialFile)
    {
        try
        {
            // Load XML document
            XDocument xdoc = XDocument.Load(xmlFilePath);

            // Create a new Excel package
            using (var package = new ExcelPackage())
            {
                // Convert XML to flat data structure
                var flattenedData = FlattenXml(xdoc.Root);

                // Create worksheets for different sections
                foreach (var section in flattenedData)
                {
                    var worksheet = package.Workbook.Worksheets.Add(section.Key);

                    // Get all unique keys across all rows
                    var allKeys = section.Value
                        .SelectMany(dict => dict.Keys)
                        .Distinct()
                        .ToList();

                    // Write headers
                    for (int i = 0; i < allKeys.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = allKeys[i];
                    }

                    // Write data
                    for (int rowIndex = 0; rowIndex < section.Value.Count; rowIndex++)
                    {
                        var rowData = section.Value[rowIndex];
                        for (int colIndex = 0; colIndex < allKeys.Count; colIndex++)
                        {
                            var key = allKeys[colIndex];
                            if (rowData.TryGetValue(key, out var value))
                            {
                                worksheet.Cells[rowIndex + 2, colIndex + 1].Value = value;
                            }
                        }
                    }
                }

                // Determine output path based on whether it's an initial file
                string outputFolder = isInitialFile ? initialFilesFolder : destinationFolder;
                string outputFileName = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(xmlFilePath) + ".xlsx"
                );

                // Save the Excel file
                FileInfo fileInfo = new FileInfo(outputFileName);
                package.SaveAs(fileInfo);

                Console.WriteLine($"Converted {Path.GetFileName(xmlFilePath)} to Excel: {outputFileName}");
            }
            // Ensure the ProcessedXMLs directory exists
            Directory.CreateDirectory(processedBatchFolder);

            // Move the original XML file to the ProcessedXMLs folder
            string processedXmlPath = Path.Combine(processedBatchFolder, Path.GetFileName(xmlFilePath));
            File.Move(xmlFilePath, processedXmlPath);

            Console.WriteLine($"Moved {Path.GetFileName(xmlFilePath)} to ProcessedXMLs: {processedXmlPath}");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing {Path.GetFileName(xmlFilePath)}: {ex.Message}");
        }
    }

    /// <summary>
    /// Flattens a complex XML structure into a dictionary of lists for Excel conversion
    /// </summary>
    /// <param name="element">Root XML element to flatten</param>
    /// <returns>A dictionary where keys are element names and values are lists of flattened data</returns>
    private Dictionary<string, List<Dictionary<string, string>>> FlattenXml(XElement element)
    {
        var result = new Dictionary<string, List<Dictionary<string, string>>>();

        // Recursively process the XML
        ProcessElement(element, result);

        return result;
    }

    /// <summary>
    /// Recursively processes XML elements to extract data into a flattened structure
    /// </summary>
    /// <param name="element">Current XML element to process</param>
    /// <param name="result">Dictionary to store flattened data</param>
    private void ProcessElement(XElement element, Dictionary<string, List<Dictionary<string, string>>> result)
    {
        // Traverse through child elements
        foreach (var childGroup in element.Elements().GroupBy(e => e.Name.LocalName))
        {
            string groupName = childGroup.Key;
            var groupList = new List<Dictionary<string, string>>();

            foreach (var childElement in childGroup)
            {
                var rowData = new Dictionary<string, string>();

                // Process attributes
                foreach (var attr in childElement.Attributes())
                {
                    rowData[attr.Name.LocalName] = attr.Value;
                }

                // Process direct child elements
                foreach (var leaf in childElement.Elements())
                {
                    // Only add leaf nodes (elements without further children)
                    if (!leaf.Elements().Any())
                    {
                        rowData[leaf.Name.LocalName] = leaf.Value;
                    }
                }

                // Add row data to the group list
                groupList.Add(rowData);
            }

            // Add the group to the result
            result[groupName] = groupList;

            // Recursively process child elements that might have nested structures
            foreach (var childElement in childGroup)
            {
                ProcessElement(childElement, result);
            }
        }
    }

    /// <summary>
    /// Main entry point of the application
    /// </summary>
    /// <param name="args">Command-line arguments to specify mode</param>
    static void Main(string[] args)
    {
        // Specify the folders 
        string watchFolder = @"C:\XmlWatcherService\XML2EXCEL\XMLInput";
        string destinationFolder = @"C:\XmlWatcherService\XML2EXCEL\ConvertedXML";
        string processedXmlsFolder = @"C:\XmlWatcherService\XML2EXCEL\ProcessedXMLs";

        // Create converter
        var converter = new XmlToExcelConverter(watchFolder, destinationFolder, processedXmlsFolder);

        BasicSQLQueryTest();


        // Determine mode based on command-line argument
        if (args.Length > 0)
        {
            switch (args[0].ToLower())
            {
                case "batch":
                    Console.WriteLine("Running in Batch Mode");
                    converter.ConvertXmlToExcelBatchHandling();
                    break;

                case "watch":
                    Console.WriteLine("Running in Watch Mode");
                    converter.ConvertXmlToExcelWatchHandling();
                    break;

                default:
                    Console.WriteLine("Invalid mode. Use 'batch' or 'watch'.");
                    ShowUsage();
                    return;
            }
        }
        else
        {
            ShowUsage();
        }
    }

    /// <summary>
    /// Tests the SQL Connection with a basic query from the XmlRepo File. If this isn't working then there's a problem with the connection.
    /// </summary>
    private static void BasicSQLQueryTest()
    {
        // SQL query
        string query = "SELECT TOP 1 * FROM XmlRepo";

        // connection
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            try
            {
                connection.Open();
                Console.WriteLine("Connection successful!");

                // Execute query
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            // Replace 0 and 1 with your column indices
                            Console.WriteLine($"ID: {reader[0]}, Name: {reader[1]}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Displays usage instructions for the application
    /// </summary>
    static void ShowUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("XmlToExcelConverter.exe batch   - Convert existing XML files");
        Console.WriteLine("XmlToExcelConverter.exe watch   - Watch folder for new XML files");
    }
}