using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Xml.Linq;

namespace XML2Excel
{
    public class ExcelToXmlConverter
    {
        public static void ConvertExcelToXml(string inputExcelPath, string outputXmlPath)
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read the Excel file
            using (var package = new ExcelPackage(new FileInfo(inputExcelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Create XML document
                XDocument xmlDoc = new XDocument(
                    new XElement("Invoices",
                        new XElement("BOM",
                            new XElement("BO",
                                new XElement("AdmInfo",
                                    new XElement("Object", "oInvoices")
                                ),
                                new XElement("QueryParams",
                                    new XElement("DocEntry")
                                ),
                                new XElement("Documents",
                                    CreateDocumentsElement(worksheet)
                                ),
                                new XElement("Document_Lines",
                                    CreateDocumentLinesElement(worksheet)
                                )
                            )
                        )
                    )
                );

                // Save XML file
                xmlDoc.Save(outputXmlPath);
            }
        }

        private static XElement CreateDocumentsElement(ExcelWorksheet worksheet)
        {
            var documentsElement = new XElement("row");

            // Map document-level columns
            var documentColumns = new[]
            {
            "DocType", "HandWritten", "MRAQRCode", "MRAIRN", "PrevHashMra",
            "MraInvoiceCounter", "Inv_Time", "NumAtCard", "LocalSalesType",
            "CardCode", "MRAID", "InvNum", "Company", "CustCode", "CustId",
            "DocDate", "DeliveryDate", "Address1", "Address2", "Address3",
            "Vat_No", "CustDescription", "OwnerCode", "Inv_No", "SalesmanCode",
            "InvWhse", "UserId", "AgentID", "Statusid"
        };

            foreach (var column in documentColumns)
            {
                int columnIndex = FindColumnIndex(worksheet, column);
                if (columnIndex != -1)
                {
                    documentsElement.Add(new XElement(column,
                        worksheet.Cells[2, columnIndex].Text.Trim()));
                }
                else
                {
                    documentsElement.Add(new XElement(column));
                }
            }

            return documentsElement;
        }

        private static XElement CreateDocumentLinesElement(ExcelWorksheet worksheet)
        {
            var documentLinesElement = new XElement("row");

            // Find start of document lines
            int startRow = FindStartOfDocumentLines(worksheet);

            // Map document lines columns
            var documentLinesColumns = new[]
            {
            "LineNum", "WarehouseCode", "ProdDescription", "Qty", "UnitPrice",
            "OriginalUnitPrice", "ProdGroup", "LotNo", "UnitPriceVat",
            "VatPercent", "StockId", "TaxCode", "_id", "TotalExclVat",
            "TotalInclVat", "NumBox", "UOM", "RSP", "CutsSpec", "DiscountPer",
            "TotalDiscount", "TotalVatAmt"
        };

            // Create multiple rows for each set of detail lines
            var rows = new XElement("row");
            int lastPopulatedRow = worksheet.Dimension.End.Row;

            for (int row = startRow; row <= lastPopulatedRow; row++)
            {
                var currentRow = new XElement("row");

                foreach (var column in documentLinesColumns)
                {
                    int columnIndex = FindColumnIndex(worksheet, column);
                    if (columnIndex != -1)
                    {
                        currentRow.Add(new XElement(column,
                            worksheet.Cells[row, columnIndex].Text.Trim()));
                    }
                    else
                    {
                        currentRow.Add(new XElement(column));
                    }
                }

                rows.Add(currentRow);
            }

            return rows;
        }

        private static int FindColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (worksheet.Cells[1, col].Text.Trim() == columnName)
                {
                    return col;
                }
            }
            return -1;
        }

        private static int FindStartOfDocumentLines(ExcelWorksheet worksheet)
        {
            // Start from row 2 (assuming first row is headers)
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                // Check if the row contains data in a column like WarehouseCode
                int warehouseCodeIndex = FindColumnIndex(worksheet, "WarehouseCode");
                if (warehouseCodeIndex != -1 &&
                    !string.IsNullOrWhiteSpace(worksheet.Cells[row, warehouseCodeIndex].Text))
                {
                    return row;
                }
            }
            return 2; // Default to second row if no specific start found
        }
    }
}
