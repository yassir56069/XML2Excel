using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Xml.Linq;
using System.Linq;
using System.Threading;
using OfficeOpenXml;

class XmlToExcelConverter
{
    private string watchFolder;
    private string destinationFolder;

    public XmlToExcelConverter(string watchFolder, string destinationFolder)
    {
        // Ensure the destination folder exists
        Directory.CreateDirectory(destinationFolder);

        this.watchFolder = watchFolder;
        this.destinationFolder = destinationFolder;

        // Set EPPlus license context (required for EPPlus 5.0+)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public void StartWatching()
    {
        // Create a FileSystemWatcher to monitor the folder
        FileSystemWatcher watcher = new FileSystemWatcher(watchFolder);

        // Set up event handlers
        watcher.Created += OnXmlFileCreated;

        // Only watch XML files
        watcher.Filter = "*.xml";

        // Enable the watcher
        watcher.EnableRaisingEvents = true;

        Console.WriteLine($"Watching folder: {watchFolder} for XML files");
        Console.WriteLine($"Converted files will be saved to: {destinationFolder}");
        Console.WriteLine("Press 'Q' to quit the application.");

        // Keep the application running
        while (Console.ReadKey().Key != ConsoleKey.Q)
        {
            Thread.Sleep(1000);
        }
    }

    private void OnXmlFileCreated(object sender, FileSystemEventArgs e)
    {
        try
        {
            // Wait a moment to ensure file is fully written
            Thread.Sleep(500);

            // Convert the XML to Excel
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

    static void Main(string[] args)
    {
        // Specify the folders (you can modify these paths as needed)
        string watchFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "XMLInput"
        );
        string destinationFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ConvertedXML"
        );

        // Create converter and start watching
        var converter = new XmlToExcelConverter(watchFolder, destinationFolder);
        converter.StartWatching();
    }
}