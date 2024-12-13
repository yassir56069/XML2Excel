﻿using System;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using OfficeOpenXml;

/// <summary>
/// Provides a robust XML to Excel conversion utility with file watching and batch conversion capabilities.
/// This application monitors a specified folder for XML files, allows batch conversion of existing files,
/// and supports continuous file monitoring.
/// </summary>
/// <remarks>
/// Key Features:
/// - Scans input folder on launch for existing XML files
/// - Provides user option to convert existing files
/// - Continuous file system watching for new XML files
/// - Flexible XML to Excel conversion supporting complex XML structures
/// </remarks>
class XmlToExcelConverter
{
    /// <summary>
    /// Folder path where XML input files are located
    /// </summary>
    private string watchFolder;

    /// <summary>
    /// Destination folder where converted Excel files will be saved
    /// </summary>
    private string destinationFolder;

    /// <summary>
    /// Subfolder for initially found XML files
    /// </summary>
    private string initialFilesFolder;

    /// <summary>
    /// Initializes a new instance of the XmlToExcelConverter with specified folder paths
    /// </summary>
    /// <param name="watchFolder">Directory to monitor for XML files</param>
    /// <param name="destinationFolder">Directory where converted files will be saved</param>
    public XmlToExcelConverter(string watchFolder, string destinationFolder)
    {
        // Ensure all required directories exist
        this.watchFolder = watchFolder;
        this.destinationFolder = destinationFolder;
        this.initialFilesFolder = Path.Combine(destinationFolder, "InitialFiles");

        // Create directories if they don't exist
        Directory.CreateDirectory(watchFolder);
        Directory.CreateDirectory(destinationFolder);
        Directory.CreateDirectory(initialFilesFolder);

        // Set EPPlus license context (required for EPPlus 5.0+)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    /// <summary>
    /// Starts the XML conversion process by first checking for existing files
    /// and then setting up continuous file system monitoring
    /// </summary>
    public void Start()
    {
        // Check for existing XML files
        string[] existingXmlFiles = Directory.GetFiles(watchFolder, "*.xml");

        if (existingXmlFiles.Length > 0)
        {
            Console.WriteLine($"Found {existingXmlFiles.Length} existing XML file(s).");
            Console.Write("Would you like to convert these files now? (Y/N): ");

            var response = Console.ReadKey();
            Console.WriteLine(); // Move to next line

            if (char.ToUpper(response.KeyChar) == 'Y')
            {
                // Convert existing files
                foreach (var file in existingXmlFiles)
                {
                    try
                    {
                        ConvertXmlToExcel(file, true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
            }
        }

        // Start file system watcher for new files
        StartWatching();
    }

    /// <summary>
    /// Sets up a FileSystemWatcher to continuously monitor the input folder for new XML files
    /// </summary>
    private void StartWatching()
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
    /// <param name="args">Command-line arguments (not used)</param>
    static void Main(string[] args)
    {
        // Specify the folders 
        string watchFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "XMLInput"
        );
        string destinationFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ConvertedXML"
        );

        // Create converter and start processing
        var converter = new XmlToExcelConverter(watchFolder, destinationFolder);
        converter.Start();
    }
}