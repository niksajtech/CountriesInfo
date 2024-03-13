using System;
using System.Data;
using System.Data.SqlClient;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CountriesInfo
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Data Source=.;Initial Catalog=niksaj;Integrated Security=True";
            string query = "SELECT CountryName,CountryFlagImage, Flag, State_emblem, Capital, Largest_cities, " +
                           "Official_Languages, Religion, Independence_Date, Republic_Date, Total_Area, " +
                           "Currency, Date_format, Driving_side, Calling_code, ISO_3166_code, Internet_TLD, " +
                           "National_Game, National_Bird, Important_Rivers, Highest_Citizen_award, Parliament, " +
                           "Judiciary FROM dbo.Country ORDER BY CountryName ASC";

            string filePath = @"C:\Users\hp\Documents\ebook\CountryInformation.docx";

            // Create Word document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Add a new main document part
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Set default font and size
                RunProperties defaultRunProperties = new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman" },
                    new FontSize() { Val = "24" } // 12pt
                );

                // Set font size for CountryName to 36pt
                RunProperties countryNameRunProperties = new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman" },
                    new FontSize() { Val = "72" } // 36pt
                );

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(query, connection);
                    SqlDataReader reader = command.ExecuteReader();

                    DataTable schemaTable = reader.GetSchemaTable();

                    while (reader.Read())
                    {
                        // Iterate through schema to get column names
                        foreach (DataRow row in schemaTable.Rows)
                        {
                            string columnName = row["ColumnName"].ToString();
                            string columnValue = reader[columnName].ToString();

                            // Add country information as paragraphs
                            if (columnName == "CountryName")
                            {
                                AddParagraph(body, "", columnValue, true, countryNameRunProperties, true); // Centered and bold
                            }
                            else
                            {
                                AddParagraph(body, $"{columnName}: ", columnValue, true, defaultRunProperties, false);
                            }
                        }

                        // Insert a page break
                        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    }
                }
            }

            Console.WriteLine("Word document created successfully.");
        }

        static void AddParagraph(Body body, string columnName, string columnValue, bool isBold, RunProperties runProperties, bool isCentered)
        {
            // Create a new paragraph
            Paragraph paragraph = body.AppendChild(new Paragraph());

            // Set alignment if required
            if (isCentered)
            {
                paragraph.AppendChild(new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ));
            }

            // Append the column name
            Run columnNameRun = paragraph.AppendChild(new Run(new Text(columnName)));

            // Apply bold formatting to the column name if necessary
            if (isBold)
            {
                Bold bold = new Bold();
                columnNameRun.RunProperties = new RunProperties(bold, runProperties.CloneNode(true));
            }
            else
            {
                columnNameRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
            }

            // Append the column value
            Run columnValueRun = paragraph.AppendChild(new Run(new Text(columnValue)));

            // Apply bold formatting to the column value if it's the country name
            if (columnName.Trim().Equals("CountryName", StringComparison.OrdinalIgnoreCase))
            {
                Bold bold = new Bold();
                columnValueRun.RunProperties = new RunProperties(bold, runProperties.CloneNode(true));
            }
            else
            {
                columnValueRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
            }
        }
    }
}
