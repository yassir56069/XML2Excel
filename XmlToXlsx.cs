using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XML2Excel
{

    using System;
    using System.Data;
    using System.IO;
    using System.Xml;
    using OfficeOpenXml;

    public class XmlToXlsx
    {
        public static void ConvertXmlToExcel(string xmlFilePath, string excelFilePath)
        {
            // Load the XML document
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(xmlFilePath);

            // Create a DataTable to hold the data
            DataTable dataTable = new DataTable("Data");

            // Add the "Invoices" column at the start
            dataTable.Columns.Add("Invoices");

            // Parse XML and populate the DataTable
            ParseXmlNode(xmlDocument.DocumentElement, dataTable);

            // Create the Excel file
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("XML Data");

                // Add column headers
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true; // Make header bold
                }

                // Add data rows
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j]?.ToString();
                    }
                }

                // Generate output filename
                string outputFileName = Path.Combine(
                    excelFilePath,
                    Path.GetFileNameWithoutExtension(xmlFilePath) + ".xlsx"
                );

                // Save the Excel file
                FileInfo fileInfo = new FileInfo(outputFileName);
                package.SaveAs(fileInfo);

                

                Console.WriteLine($"Converted {Path.GetFileName(xmlFilePath)} to Excel: {outputFileName}");
            }
        }
        public static void RevertToXml(string excelFilePath, string xmlFilePath)
        {
            // Load the Excel file
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Create a new XmlDocument
                XmlDocument xmlDocument = new XmlDocument();

                // Create the root element based on the first column header (e.g., "Invoices")
                string rootName = SanitizeXmlName(worksheet.Cells[1, 1].Value?.ToString() ?? "Root");
                XmlElement rootElement = xmlDocument.CreateElement(rootName);
                xmlDocument.AppendChild(rootElement);

                // Process rows in the worksheet
                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    XmlElement rowElement = xmlDocument.CreateElement("row");

                    for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                    {
                        string columnName = worksheet.Cells[1, j].Value?.ToString();
                        string cellValue = worksheet.Cells[i, j].Value?.ToString();

                        if (!string.IsNullOrEmpty(columnName))
                        {
                            // Sanitize the column name
                            string sanitizedColumnName = SanitizeXmlName(columnName);

                            XmlElement columnElement = xmlDocument.CreateElement(sanitizedColumnName);
                            columnElement.InnerText = cellValue ?? string.Empty;
                            rowElement.AppendChild(columnElement);
                        }
                    }

                    rootElement.AppendChild(rowElement);
                }

                // Save the XML file
                xmlDocument.Save(xmlFilePath);
            }
        }

        // Helper method to sanitize XML names
        private static string SanitizeXmlName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "Unnamed";

            // Use XmlConvert.EncodeName to handle invalid characters
            string encodedName = XmlConvert.EncodeName(name);

            // Ensure the name does not start with a number or invalid character
            if (!XmlConvert.IsStartNCNameChar(encodedName[0]))
            {
                encodedName = "_" + encodedName;
            }

            return encodedName;
        }
        private static void ParseXmlNode(XmlNode node, DataTable dataTable)
        {
            if (node == null) return;

            // Ensure all attributes and child elements become columns
            if (node.Attributes != null)
            {
                foreach (XmlAttribute attribute in node.Attributes)
                {
                    if (!dataTable.Columns.Contains(attribute.Name))
                        dataTable.Columns.Add(attribute.Name);
                }
            }

            if (node.ChildNodes != null)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (!dataTable.Columns.Contains(childNode.Name))
                        dataTable.Columns.Add(childNode.Name);
                }
            }

            // Create a new DataRow for each <row> tag
            if (node.Name.Equals("row", StringComparison.OrdinalIgnoreCase))
            {
                DataRow row = dataTable.NewRow();

                // Populate attributes
                if (node.Attributes != null)
                {
                    foreach (XmlAttribute attribute in node.Attributes)
                    {
                        row[attribute.Name] = attribute.Value;
                    }
                }

                // Populate child node values
                if (node.ChildNodes != null)
                {
                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        if (!string.IsNullOrEmpty(childNode.Name))
                        {
                            row[childNode.Name] = childNode.InnerText ?? string.Empty;
                        }
                    }
                }

                dataTable.Rows.Add(row);
            }

            // Recursively process child nodes
            if (node.ChildNodes != null)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    ParseXmlNode(childNode, dataTable);
                }
            }
        }
    }
}

