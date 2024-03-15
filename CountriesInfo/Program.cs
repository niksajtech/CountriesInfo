using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SkiaSharp;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Linq;

namespace CountriesInfo
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Data Source=.;Initial Catalog=niksaj;Integrated Security=True";
            string query = "SELECT CountryName, CountryFlagImage, Capital, Largest_cities, " +
                "Official_Languages, Religion, " +
                "CASE WHEN Independence_Date IS NULL THEN 'Not Available' ELSE CONVERT(varchar, Independence_Date, 103) END AS Independence_Date, " +
                "CASE WHEN Republic_Date IS NULL THEN 'Not Available' ELSE CONVERT(varchar, Republic_Date, 103) END AS Republic_Date, " +
                "Total_Area, Currency, Date_format, Driving_side, Calling_code, ISO_3166_code, Internet_TLD, " +
                "National_Game, National_Bird, Important_Rivers, Highest_Citizen_award, Parliament, " +
                "Judiciary " +
                "FROM dbo.Country ORDER BY CountryName ASC";

            string filePath = @"C:\Users\hp\Documents\ebook\CountryInformation.docx";

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                AddCoverPage(body); // Add cover page

                RunProperties defaultRunProperties = new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman" },
                    new FontSize() { Val = "28" } // 14pt
                );

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

                    int rowCount = 0; // Counter to track the number of rows processed

                    while (reader.Read())
                    {
                        // Check if there's enough space for the content and the image
                        if (rowCount >= 15)
                        {
                            // Check if the previous page is empty
                            if (IsPageEmpty(body))
                            {
                                RemoveLastPage(body);
                            }

                            // Insert a page break
                            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                            rowCount = 0; // Reset the row count after inserting a page break
                        }

                        // Iterate through schema to get column names and values
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
                                if (columnName != "CountryFlagImage")
                                {
                                    AddParagraph(body, $"{columnName}: ", columnValue, true, defaultRunProperties, false);
                                }
                            }
                            if (columnName == "CountryFlagImage")
                            {
                                byte[] imageData = (byte[])reader["CountryFlagImage"];
                                AddImageToBody(wordDocument, body, imageData);
                            }

                            // Check if the current column is "Judiciary" and insert a page break after it
                            if (columnName == "Judiciary")
                            {
                                // Check if the previous page is empty
                                if (IsPageEmpty(body))
                                {
                                    RemoveLastPage(body);
                                }

                                // Insert a page break
                                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                                rowCount = 0; // Reset the row count after inserting a page break
                            }
                        }

                        rowCount++; // Increment the row count for each row processed
                    }

                    // Check if the last page is empty and remove it
                    if (IsPageEmpty(body))
                    {
                        RemoveLastPage(body);
                    }
                }
            }

            Console.WriteLine("Word document created successfully.");
        }

        static void AddParagraph(Body body, string columnName, string columnValue, bool isBold, RunProperties runProperties, bool isCentered)
        {
            // Remove underscores from column name and replace with spaces
            columnName = columnName.Replace("_", " ");

            Paragraph paragraph = body.AppendChild(new Paragraph());

            if (isCentered)
            {
                paragraph.AppendChild(new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ));
            }

            Run columnNameRun = paragraph.AppendChild(new Run(new Text(columnName)));

            if (columnName.Equals("CountryName", StringComparison.OrdinalIgnoreCase))
            {
                // Apply header style for "CountryName"
                Bold bold = new Bold();
                RunProperties headerRunProperties = new RunProperties(bold, runProperties.CloneNode(true));
                columnNameRun.RunProperties = headerRunProperties;
            }
            else
            {
                // Apply the specified run properties for other columns
                if (isBold)
                {
                    Bold bold = new Bold();
                    columnNameRun.RunProperties = new RunProperties(bold, runProperties.CloneNode(true));
                }
                else
                {
                    columnNameRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                }
            }

            // Check if the column value is a date
            if (columnName.Trim().Equals("Independence_Date", StringComparison.OrdinalIgnoreCase) ||
                columnName.Trim().Equals("Republic_Date", StringComparison.OrdinalIgnoreCase))
            {
                // Parse the date
                if (DateTime.TryParse(columnValue, out DateTime date))
                {
                    // Format the date as dd/mm/yyyy without time part
                    columnValue = date.ToString("dd/MM/yyyy");
                }
                else
                {
                    columnValue = "Not Available"; // Display "N/A" if the date is not available
                }
            }

            Run columnValueRun = paragraph.AppendChild(new Run(new Text(columnValue)));

            if (columnName.Trim().Equals("CountryName", StringComparison.OrdinalIgnoreCase))
            {
                Bold bold = new Bold();
                RunProperties headerRunProperties = new RunProperties(bold, runProperties.CloneNode(true));
                columnNameRun.RunProperties = headerRunProperties;
            }
            else
            {
                columnValueRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
            }
        }

        static void AddImageToBody(WordprocessingDocument wordDoc, Body body, byte[] imageData)
        {
            using (MemoryStream imageStream = new MemoryStream(imageData))
            {
                using (SKBitmap bitmap = SKBitmap.Decode(imageStream))
                {
                    var imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

                    using (MemoryStream stream = new MemoryStream())
                    {
                        bitmap.Encode(stream, SKEncodedImageFormat.Jpeg, 100);
                        stream.Seek(0, SeekOrigin.Begin);
                        imagePart.FeedData(stream);
                    }

                    long originalWidth = bitmap.Width;
                    long originalHeight = bitmap.Height;

                    var element =
                        new Drawing(
                            new DW.Inline(
                                new DW.Extent() { Cx = originalWidth * 9525, Cy = originalHeight * 9525 },
                                new DW.EffectExtent()
                                {
                                    LeftEdge = 0L,
                                    TopEdge = 0L,
                                    RightEdge = 0L,
                                    BottomEdge = 0L
                                },
                                new DW.DocProperties()
                                {
                                    Id = (UInt32Value)1U,
                                    Name = "Picture 1"
                                },
                                new DW.NonVisualGraphicFrameDrawingProperties(
                                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                                new A.Graphic(
                                    new A.GraphicData(
                                        new PIC.Picture(
                                            new PIC.NonVisualPictureProperties(
                                                new PIC.NonVisualDrawingProperties()
                                                {
                                                    Id = (UInt32Value)0U,
                                                    Name = "New Bitmap Image.jpg"
                                                },
                                                new PIC.NonVisualPictureDrawingProperties()),
                                            new PIC.BlipFill(
                                                new A.Blip(
                                                    new A.BlipExtensionList(
                                                        new A.BlipExtension()
                                                        {
                                                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                        })
                                                )
                                                {
                                                    Embed = wordDoc.MainDocumentPart.GetIdOfPart(imagePart),
                                                    CompressionState = A.BlipCompressionValues.Print
                                                },
                                                new A.Stretch(
                                                    new A.FillRectangle())),
                                            new PIC.ShapeProperties(
                                                new A.Transform2D(
                                                    new A.Offset() { X = 0L, Y = 0L },
                                                    new A.Extents() { Cx = originalWidth * 9525, Cy = originalHeight * 9525 }),
                                                new A.PresetGeometry(
                                                    new A.AdjustValueList()
                                                )
                                                { Preset = A.ShapeTypeValues.Rectangle }))
                                    )
                                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                            )
                            {
                                DistanceFromTop = (UInt32Value)0U,
                                DistanceFromBottom = (UInt32Value)0U,
                                DistanceFromLeft = (UInt32Value)0U,
                                DistanceFromRight = (UInt32Value)0U,
                                EditId = "50D07946"
                            });

                    // Create border
                    var border = new A.Outline(
                        new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = "000000" } })
                    {
                        Width = 1 // Width of the border in EMUs (English Metric Units)
                    };

                    // Apply border to the shape properties of the image
                    element.Descendants<PIC.ShapeProperties>().First().Append(border);

                    Paragraph paragraph = new Paragraph();
                    paragraph.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
                    paragraph.AppendChild(new Run(element));
                    body.AppendChild(paragraph);
                }
            }
        }

        static bool IsPageEmpty(Body body)
        {
            return body.Elements<Paragraph>().All(p => string.IsNullOrWhiteSpace(p.InnerText));
        }

        static void RemoveLastPage(Body body)
        {
            // Remove the last paragraph, which represents the empty page
            var lastParagraph = body.Elements<Paragraph>().LastOrDefault();
            if (lastParagraph != null)
            {
                lastParagraph.Remove();
            }
        }

        static void AddCoverPage(Body body)
        {
            Paragraph title = new Paragraph();
            title.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
            title.AppendChild(new Run(new Text() { Text = "WORLD WANDERER" })
            {
                RunProperties = new RunProperties(new RunFonts() { Ascii = "Times New Roman" }, new FontSize() { Val = "144" })// 72pt
            });
            body.AppendChild(title);

            Paragraph subtitle = new Paragraph();
            subtitle.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
            subtitle.AppendChild(new Run(new Text() { Text = "A TOUR OF 195 NATIONS" })
            {
                RunProperties = new RunProperties(new RunFonts() { Ascii = "Times New Roman" }, new FontSize() { Val = "44" })// 22pt
            });
            body.AppendChild(subtitle);

            for (int i = 0; i < 18; i++) // Add 18 line breaks
            {
                body.AppendChild(new Paragraph());
            }

            Paragraph author = new Paragraph();
            author.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
            author.AppendChild(new Run(new Text() { Text = "SAJID AHMED SHAHSROHA" })
            {
                RunProperties = new RunProperties(new RunFonts() { Ascii = "Times New Roman" }, new FontSize() { Val = "44" })// 22pt
            });
            body.AppendChild(author);

            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page }))); // Insert page break after cover page
        }



    }
}
