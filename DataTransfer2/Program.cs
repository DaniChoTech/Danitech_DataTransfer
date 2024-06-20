using System;
using System.IO;
using System.Xml;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Console.Write("FileName : ");
        string tableName = Console.ReadLine();
        string excelFilePath = $"./{tableName}.xlsx";

        string xmlFilePath = $"./{tableName}.xml";

        try
        {
            ConvertExcelToXml(tableName, excelFilePath, xmlFilePath);
            Console.WriteLine($"Successfully converted '{excelFilePath}' to '{xmlFilePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static void ConvertExcelToXml(string tableName, string excelFilePath, string xmlFilePath)
    {
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // 첫 번째 워크시트를 선택

            if (worksheet.Dimension == null)
            {
                throw new Exception("Worksheet is empty");
            }

            var headers = new string[worksheet.Dimension.End.Column];
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                headers[col - 1] = worksheet.Cells[1, col].Text;
            }

            using (XmlWriter writer = XmlWriter.Create(xmlFilePath, new XmlWriterSettings { Indent = true }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement(tableName);

                writer.WriteStartElement("dataCategory");
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // 첫 번째 행을 건너뛰고 시작
                {
                    writer.WriteStartElement("data");
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        writer.WriteAttributeString(headers[col - 1], worksheet.Cells[row, col].Text);
                    }
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
    }
}
