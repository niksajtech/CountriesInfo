using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MO=Microsoft.Office.Interop.Word;
using SkiaSharp;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


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
                "National_Game, National_Bird, Important_Rivers, Highest_Citizen_award, Parliament,Surrounded_By, " +
                "Judiciary " +
                "FROM dbo.Country ORDER BY CountryName ASC";

            string filePath = @"C:\Users\hp\Documents\ebook\CountryInformation.docx";

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                RunProperties countryNameRunProperties = new RunProperties(
                    new RunStyle() { Val = "Heading 1" },
                    new RunFonts() { Ascii = "Times New Roman" },
                    new FontSize() { Val = "72" } // 36pt
                );

                // Define RunProperties for header content
                RunProperties headerRunProperties = new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman" },
                    new FontSize() { Val = "20" } // Adjust font size as needed
                );

                // Define RunProperties before calling AddCoverPage method
                RunProperties defaultRunProperties = new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman" },
                    new FontSize() { Val = "28" } // 14pt
                );


                AddCoverPage(body); // Add cover page

                AddCopyRightPage(body); // Add copy right page

                // Create header part and generate header content
                HeaderPart headerPartSubsequentPages = wordDocument.MainDocumentPart.AddNewPart<HeaderPart>();
                GenerateHeader(headerRunProperties, "TITLE : WORLD WANDERER", "AUTHOR : NIKHAT SHAHIN KHAN").Save(headerPartSubsequentPages);

                // Create footer part and generate footer content
                FooterPart footerPart = wordDocument.MainDocumentPart.AddNewPart<FooterPart>();
                GenerateFooterWithPageNumber().Save(footerPart);

                wordDocument.MainDocumentPart.Document.AppendChild(
       new SectionProperties(
           new HeaderReference { Id = wordDocument.MainDocumentPart.GetIdOfPart(headerPartSubsequentPages) },
           new FooterReference { Id = wordDocument.MainDocumentPart.GetIdOfPart(footerPart) }
       )
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

                        // Iterate through schema to get column names and values
                        int paragraphsPerPage = 19; // Define the maximum number of paragraphs per page
                        int paragraphCount = 0; // Counter to track the number of paragraphs added

                        foreach (DataRow row in schemaTable.Rows)
                        {
                            string columnName = row["ColumnName"].ToString();
                            string columnValue = reader[columnName].ToString();

                            // Add country information as paragraphs
                            if (columnName == "CountryName")
                            {
                                AddParagraph(body, "", columnValue.ToUpper(), true, countryNameRunProperties, true); // Centered and bold
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

                            // Increment paragraph count
                            paragraphCount++;

                            // Check if the current column is "Judiciary" or if the maximum paragraphs per page is reached
                            if (columnName == "Judiciary" || paragraphCount >= paragraphsPerPage)
                            {
                                // Check if the previous page is empty
                                if (IsPageEmpty(body))
                                {
                                    RemoveLastPage(body);
                                }

                                // Insert a page break
                                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                                // Reset the paragraph count after inserting a page break
                                paragraphCount = 0;
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

		private static void AddCopyRightPage(Body body)
        {
            // Add title
            Paragraph title = new Paragraph();
            string copyrightHeading = "Copyright/Disclaimers".ToUpper();
            title.AppendChild(new Run(new Text(copyrightHeading))
            {
                RunProperties = new RunProperties(new RunFonts() { Ascii = "Times New Roman" }, new Bold(), new FontSize() { Val = "56" })
            });
            body.AppendChild(title);

            // Add copyright text
            string[] copyrightText = new string[]
            {
        "All rights reserved. No part of this publication may be reproduced, distributed, or transmitted in any form or by any means, including photocopying, recording, or other electronic or mechanical methods, without the prior written permission of the publisher, except in the case of brief quotations embodied in critical reviews and certain other noncommercial uses permitted by copyright law.",
         "",
        "Disclaimer:",
        "",
        "The information contained in this ebook is for general information purposes only. While we endeavor to keep the information up to date and correct, we make no representations or warranties of any kind, express or implied, about the completeness, accuracy, reliability, suitability, or availability with respect to the ebook or the information, products, services, or related graphics contained in the ebook for any purpose. Any reliance you place on such information is therefore strictly at your own risk.",
        "",
        "In no event will we be liable for any loss or damage including without limitation, indirect or consequential loss or damage, or any loss or damage whatsoever arising from loss of data or profits arising out of, or in connection with, the use of this ebook.",
        "",
        "Through this ebook, you are able to link to other websites that are not under the control of the publisher. We have no control over the nature, content, and availability of those sites. The inclusion of any links does not necessarily imply a recommendation or endorse the views expressed within them.",
        "",
        "Every effort is made to keep the ebook up and running smoothly. However, the publisher takes no responsibility for, and will not be liable for, the ebook being temporarily unavailable due to technical issues beyond our control."
            };

            foreach (string paragraphText in copyrightText)
            {
                Paragraph paragraph = new Paragraph();
                paragraph.AppendChild(new Run(new Text(paragraphText))
                {
                    RunProperties = new RunProperties(new FontSize() { Val = "28" }, new RunFonts() { Ascii = "Times New Roman" })
                });
                body.AppendChild(paragraph);
            }

            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page }))); // Insert page break after cover page
        }






        // Method to generate header content for subsequent pages
        static Header GenerateHeader(RunProperties runProperties, string title, string author)
        {
            Header header = new Header();

            // Create a table with a single row and two columns
            Table table = new Table();
            TableRow row = new TableRow();

            // Create the first cell (left column) in the row
            TableCell cell1 = new TableCell();
            cell1.AppendChild(new TableCellProperties(new TableCellWidth { Width = "50%" }
            ));
            Paragraph titleParagraph = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
            titleParagraph.AppendChild(new Run(new Text(title))
            {
                RunProperties = runProperties.CloneNode(true) as RunProperties
            });
            cell1.AppendChild(titleParagraph);
            row.AppendChild(cell1);

            // Create the second cell (right column) in the row
            TableCell cell2 = new TableCell();
            cell2.AppendChild(new TableCellProperties(new TableCellWidth { Width = "50%" }
            ));
            Paragraph authorParagraph = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
            authorParagraph.AppendChild(new Run(new Text(author))
            {
                RunProperties = runProperties.CloneNode(true) as RunProperties
            });
            cell2.AppendChild(authorParagraph);
            row.AppendChild(cell2);

            // Add the row to the table
            table.AppendChild(row);

            // Set table width to match the width of the page
            table.AppendChild(new TableProperties(new TableWidth { Type = TableWidthUnitValues.Pct, Width = "100%" }));

            // Add the table to the header
            header.AppendChild(table);

            return header;
        }




        static Footer GenerateFooterWithPageNumber()
        {
            Footer footer = new Footer();

            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });

            Run pageNumberRun = new Run();
            pageNumberRun.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Begin });
            pageNumberRun.AppendChild(new RunProperties(new NoProof()));
            pageNumberRun.AppendChild(new FieldCode() { Text = @"PAGE" });
            pageNumberRun.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.End });

            paragraph.AppendChild(paragraphProperties);
            paragraph.Append(pageNumberRun);

            footer.Append(paragraph);

            return footer;
        }





        static void AddParagraph(Body body, string columnName, string columnValue, bool isBold, RunProperties runProperties, bool isCentered)
        {
            // Remove underscores from column name and replace with spaces
            columnName = columnName.Replace("_", " ");
            columnName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(columnName);

            Paragraph paragraph = body.AppendChild(new Paragraph());

            if (isCentered)
            {
                paragraph.AppendChild(new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ));
            }

            Run columnNameRun = new Run(new Text(columnName));

            if (columnName.Equals("CountryName", StringComparison.OrdinalIgnoreCase))
            {
                
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

            paragraph.AppendChild(columnNameRun);

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
             
            }
            else
            {
                if (columnName.Trim().Equals(""))
                {
                    columnValueRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                    columnValueRun.RunProperties.Color = new Color() { Val = "000EEE" };
                }
                else
                {
                    columnValueRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
                }
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
            author.AppendChild(new Run(new Text() { Text = "NIKHAT SHAHIN KHAN" })
            {
                RunProperties = new RunProperties(new RunFonts() { Ascii = "Times New Roman" }, new FontSize() { Val = "44" })// 22pt
            });
            body.AppendChild(author);

            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page }))); // Insert page break after cover page
        }
    }
}
