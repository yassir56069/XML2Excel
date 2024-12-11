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
    private readonly  string watchFolder;

    /// <summary>
    /// Folder for processed xmls in batch runs, to avoid duplicate batch runs where possible. ALso receives reverted xmls
    /// </summary>
    private readonly  string processedBatchFolder;

    /// <summary>
    /// Folder used for revert operation, 
    /// </summary>
    private readonly string revertFolder;

    /// <summary>
    /// Destination folder where converted Excel files will be saved
    /// </summary>
    private readonly  string destinationFolder;

    /// <summary>
    /// Subfolder for initially found XML files
    /// </summary>
    private readonly  string initialFilesFolder;

    /// <summary>
    /// Initializes a new instance of the XmlToExcelConverter with specified folder paths
    /// </summary>
    /// <param name="watchFolder">Directory to monitor for XML files</param>
    /// <param name="destinationFolder">Directory where converted files will be saved</param>
    public XmlToExcelConverter(string watchFolder, string destinationFolder, string processedBatchFolder, string revertFolder)
    {
        // Establish SQL Connection
        connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;


        // Ensure all required directories exist
        this.watchFolder = watchFolder;
        this.destinationFolder = destinationFolder;
        this.processedBatchFolder = processedBatchFolder;
        this.revertFolder = revertFolder;
        this.initialFilesFolder = Path.Combine(destinationFolder, "InitialFiles");

        // Create directories if they don't exist
        Directory.CreateDirectory(watchFolder);
        Directory.CreateDirectory(destinationFolder);
        Directory.CreateDirectory(processedBatchFolder);
        Directory.CreateDirectory(initialFilesFolder);
        Directory.CreateDirectory(revertFolder);

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
                    ConvertXmlToExcel(file);
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
    /// Queries the XmlRepo table to retrieve the earliest entry or an entry by specific ID
    /// </summary>
    /// <param name="id">Optional ID to query a specific record</param>
    private static void QueryXmlRepo(int? id = null)
    {
        string query;
        if (id == null)
        {
            // Query for the earliest entry by XmlDateOfEntry
            query = @"
                SELECT TOP 1 XmlID, XmlFile, XmlDateOfEntry 
                FROM XmlRepo 
                ORDER BY XmlDateOfEntry ASC";
        }
        else
        {
            // Query for a specific entry by XmlID
            query = @"
                SELECT XmlID, XmlFile, XmlDateOfEntry 
                FROM XmlRepo 
                WHERE XmlID = @XmlID";
        }

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            try
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Add parameter if querying by ID
                    if (id.HasValue)
                    {
                        command.Parameters.AddWithValue("@XmlID", id.Value);
                    }

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int xmlId = reader.GetInt32(0);
                            string xmlName = reader.GetString(1);
                            DateTime xmlDateOfEntry = reader.GetDateTime(2);

                            Console.WriteLine($"Query Result:");
                            Console.WriteLine($"XmlID: {xmlId}");   
                            Console.WriteLine($"XmlName: {xmlName}");
                            Console.WriteLine($"XmlDateOfEntry: {xmlDateOfEntry}");
                        }
                        else if (id.HasValue)
                        {
                            Console.WriteLine($"No entry found with XmlID: {id.Value}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during query: {ex.Message}");
            }
        }
    }

/// <summary>
/// Inserts XML data from the specified file into the XmlRepo table.
/// </summary>
/// <param name="xmlFilePath">The path to the XML file to be inserted.</param>
private static void InsertXmlToDB(string xmlFilePath)
{
    // Define the query with parameters
    string query = @"
        INSERT INTO XmlRepo (XmlFile, XmlDateOfEntry)
        VALUES (@XmlFile, GETDATE());
    ";

    // Load XML document as a string
    XDocument xmlContent = XDocument.Load(xmlFilePath);


        using (SqlConnection connection = new SqlConnection(connectionString))
    {
        try
        {
            connection.Open();

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                // Add the XML content as a parameter
                command.Parameters.AddWithValue("@XmlFile", xmlContent.ToString());

                // Execute the query
                int rowsAffected = command.ExecuteNonQuery();

                Console.WriteLine($"{rowsAffected} row(s) inserted into the XmlRepo table.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while inserting the XML data: {ex.Message}");
        }
    }
}


    /// <summary>
    /// Converts existing Excel files in the revert folder back to XML
    /// </summary>
    /// <returns>Number of files converted</returns>
    public int ConvertExcelToXmlBatchHandling()
    {
        // Check for existing Excel files
        string[] existingExcelFiles = Directory.GetFiles(revertFolder, "*.xlsx");
        int convertedFiles = 0;
        Console.WriteLine($"Operating in revert folder: {revertFolder}");

        if (existingExcelFiles.Length > 0)
        {
            Console.WriteLine($"Found {existingExcelFiles.Length} existing Excel file(s).");

            // Convert existing files
            foreach (var file in existingExcelFiles)
            {
                try
                {
                    ConvertExcelToXml(file);
                    convertedFiles++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting {Path.GetFileName(file)}: {ex.Message}");
                }
            }

            Console.WriteLine($"Converted {convertedFiles} file(s) back to XML in batch mode.");
        }
        else
        {
            Console.WriteLine("No Excel files found in the revert folder.");
        }

        return convertedFiles;
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
            ConvertXmlToExcel(e.FullPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file {e.Name}: {ex.Message}");
        }
    }

    private void ConvertXmlToExcel(string xmlFilePath)
    {
        // Load XML document
        XDocument xdoc = XDocument.Load(xmlFilePath);

        // Create a new Excel package
        using (var package = new ExcelPackage())
        {
            // Add a new worksheet
            var worksheet = package.Workbook.Worksheets.Add("XMLData");

            // Determine the structure of the XML
            var elements = xdoc.Descendants().Where(e => e.HasElements == false).ToList();

            // Get unique element names for headers
            var headers = elements.Select(e => e.Name.LocalName).Distinct().ToList();

            // Write headers
            for (int i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
            }

            // Group elements by their parent
            var groupedElements = elements
                .GroupBy(e => e.Parent)
                .Where(g => g.Key != null);

            // Write data
            int rowIndex = 2;
            foreach (var group in groupedElements)
            {
                int colIndex = 1;
                foreach (var header in headers)
                {
                    var value = group.FirstOrDefault(e => e.Name.LocalName == header);
                    worksheet.Cells[rowIndex, colIndex].Value = value?.Value;
                    colIndex++;
                }
                rowIndex++;
            }

            // Generate output filename
            string outputFileName = Path.Combine(
                destinationFolder,
                Path.GetFileNameWithoutExtension(xmlFilePath) + ".xlsx"
            );

            // Save the Excel file
            FileInfo fileInfo = new FileInfo(outputFileName);
            package.SaveAs(fileInfo);

            Console.WriteLine($"Converted {Path.GetFileName(xmlFilePath)} to Excel: {outputFileName}");
        }
    }


    /// <summary>
    /// Converts an Excel file back to XML
    /// </summary>
    /// <param name="excelFilePath">Full path to the Excel file</param>
    private void ConvertExcelToXml(string excelFilePath)
    {
        try
        {
            FileInfo fileInfo = new FileInfo(excelFilePath);

            // Check if file exists
            if (!fileInfo.Exists)
            {
                Console.WriteLine($"File not found: {excelFilePath}");
                return;
            }

            // Load the Excel package
            using (var package = new ExcelPackage(fileInfo))
            {
                // Check if workbook is null
                if (package.Workbook == null)
                {
                    Console.WriteLine($"Error: Workbook is null for file {excelFilePath}");
                    return;
                }

                // Create a new XML document
                XDocument xmlDoc = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement("root")
                );

                // Process each worksheet
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    // Skip worksheets with no dimension (empty worksheets)
                    if (worksheet.Dimension == null)
                    {
                        Console.WriteLine($"Skipping empty worksheet: {worksheet.Name}");
                        continue;
                    }

                    // Validate worksheet dimensions
                    int startRow = worksheet.Dimension.Start.Row;
                    int endRow = worksheet.Dimension.End.Row;
                    int startCol = worksheet.Dimension.Start.Column;
                    int endCol = worksheet.Dimension.End.Column;

                    // Ensure there are headers
                    if (endRow < startRow)
                    {
                        Console.WriteLine($"Worksheet {worksheet.Name} has no rows");
                        continue;
                    }

                    // Get headers (first row)
                    var headers = Enumerable.Range(startCol, endCol - startCol + 1)
                        .Select(col =>
                        {
                            string headerText = worksheet.Cells[startRow, col].Text;
                            return string.IsNullOrWhiteSpace(headerText) ? $"Column{col}" : headerText;
                        })
                        .ToList();

                    // Create a parent element for this worksheet
                    var worksheetElement = new XElement(worksheet.Name);
                    xmlDoc.Root.Add(worksheetElement);

                    // Process data rows (start from second row, assuming first row is headers)
                    for (int row = startRow + 1; row <= endRow; row++)
                    {
                        var rowElement = new XElement(GetSingularForm(worksheet.Name));

                        // Add each cell as an element
                        for (int col = startCol; col <= endCol; col++)
                        {
                            // Ensure headers list is not out of bounds
                            if (col - startCol >= headers.Count)
                            {
                                Console.WriteLine($"Warning: Column index out of bounds for worksheet {worksheet.Name}");
                                break;
                            }

                            string cellValue = worksheet.Cells[row, col].Text;
                            string headerName = headers[col - startCol];

                            // Only add non-empty values
                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                rowElement.Add(new XElement(headerName, cellValue));
                            }
                        }

                        // Only add row if it has elements
                        if (rowElement.Elements().Any())
                        {
                            worksheetElement.Add(rowElement);
                        }
                    }
                }

                // Determine output path in ProcessedXMLs folder
                string outputFileName = Path.Combine(
                    processedBatchFolder,
                    Path.GetFileNameWithoutExtension(excelFilePath) + ".xml"
                );

                // Save the XML file
                xmlDoc.Save(outputFileName);
                InsertXmlToDB(outputFileName);

                Console.WriteLine($"Converted {Path.GetFileName(excelFilePath)} back to XML: {outputFileName}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Detailed error converting {Path.GetFileName(excelFilePath)} to XML:");
            Console.WriteLine($"Error Type: {ex.GetType().Name}");
            Console.WriteLine($"Error Message: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }
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
    /// Converts a plural worksheet name to its singular form
    /// </summary>
    /// <param name="pluralName">Plural name of the worksheet</param>
    /// <returns>Singular form of the worksheet name</returns>
    private string GetSingularForm(string pluralName)
    {
        // Simple pluralization rules (can be expanded)
        if (pluralName.EndsWith("s"))
        {
            return pluralName.Substring(0, pluralName.Length - 1);
        }
        return pluralName;
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
        string revertFolder =  @"C:\XmlWatcherService\XML2EXCEL\RevertToXML";

        // Create converter
        var converter = new XmlToExcelConverter(watchFolder, destinationFolder, processedXmlsFolder, revertFolder);


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

                case "revert":
                    Console.WriteLine("Running in Revert Mode");
                    converter.ConvertExcelToXmlBatchHandling();
                    break;

                case "query":
                    Console.WriteLine("Running Query Mode");
                    if (args.Length == 1)
                    {
                        // No ID specified, query earliest entry
                        QueryXmlRepo();
                    }
                    else if (args.Length == 2 && int.TryParse(args[1], out int id))
                    {
                        // ID specified, query specific entry
                        QueryXmlRepo(id);
                    }
                    else
                    {
                        Console.WriteLine("Invalid query syntax. Use 'query' or 'query <id>, eg: \"XML2Excel.exe query \" or \"XML2Excel.exe query 1\"");
                        ShowUsage();
                    }
                    break;

                default:
                    Console.WriteLine("Invalid mode. Use 'batch', 'watch', 'query', or 'revert'.");
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
    /// Displays usage instructions for the application
    /// </summary>
    static void ShowUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("XmlToExcelConverter.exe batch        - Convert existing XML files");
        Console.WriteLine("XmlToExcelConverter.exe watch        - Watch folder for new XML files");
        Console.WriteLine("XmlToExcelConverter.exe revert       - Convert Excel files back to XML");
        Console.WriteLine("XmlToExcelConverter.exe query        - Query earliest XmlRepo entry");
        Console.WriteLine("XmlToExcelConverter.exe query <id>   - Query XmlRepo entry by ID");
    }
}