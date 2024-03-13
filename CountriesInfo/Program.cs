using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.IO;
using System;
using DocumentFormat.OpenXml;
using System.Data.SqlClient;
using System.Data;

//By Niksaj Technologies Pvt Ltd
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
                                if (columnName != "CountryFlagImage")
                                {
                                    AddParagraph(body, $"{columnName}: ", columnValue, true, defaultRunProperties, false);
                                }
                            }
                            if (columnName == "CountryFlagImage")
                            {
                                string imagePath1 = @"C:\Users\hp\\Downloads\163237683_10159249574270050_8620494328625146217_n.jpg";

                                AddImageToBody(wordDocument, body, imagePath1);
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

        static void AddImageToWordDocument(string imagePath, string docPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                AddImageToBody(wordDoc, body, imagePath);
            }
        }

        static void AddImageToBody(WordprocessingDocument wordDoc, Body body, string imagePath)
        {
            var imagePartId = "myImageId";

            // Add ImagePart to the document
            ImagePart imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg, imagePartId);

            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Define the desired width and height for the image (1000px)
            long desiredWidth = 200L;
            long desiredHeight = 200L;

            // Define the reference of the image with increased size
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = desiredWidth * 9525, Cy = desiredHeight * 9525 }, // 1 inch = 9525 EMUs
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
                                            Embed = imagePartId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = desiredWidth * 9525, Cy = desiredHeight * 9525 }), // Convert to EMUs
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


            // Create a paragraph
            Paragraph paragraph = new Paragraph();

            // Set paragraph alignment to center
            paragraph.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));

            // Append the image to the paragraph
            paragraph.AppendChild(new Run(element));

            // Append the paragraph to the body
            body.AppendChild(paragraph);
            // Append the reference of the image to the document body.

            //body.AppendChild(new Paragraph(new Run(element)).AppendChild(new ParagraphProperties(new Justification() { Val=JustificationValues.Center})));
        }
    }
}
