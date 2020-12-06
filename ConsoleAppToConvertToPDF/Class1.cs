using MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppToConvertToPDF
{

    class Program
    {
        /// <summary>
        /// ref taken from http://www.pdfsharp.net/
        /// 
        /// </summary>
        static void Main()
        {
            // Create a MigraDoc document
            Document document = CreateDocument();

            //string ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document, "aMigraDoc.mdddl");

            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always);
            renderer.Document = document;

            renderer.RenderDocument();

            // Save the document...
            string filename = "HelloMigraDocss.pdf";
            renderer.PdfDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }

        public static Document CreateDocument()
        {
            // Create a new MigraDoc document
            Document document = new Document();
            document.Info.Title = "PROMIS CASE NOTE";
            document.Info.Subject = "SerPro generated file of all case notes";
            document.Info.Author = "Data Migration Lange";

            DefineStyles(document);
            DefineCover(document);

            Section section = document.LastSection;
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph("Table of Contents");
            paragraph.Format.Font.Size = 14;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 24;
            paragraph.Format.OutlineLevel = OutlineLevel.Level1;

            var list = new List<LogEntry>()
            {
                new LogEntry()
                {
                    LogId=1,
                    Desc="sdfsdfsd sdf",
                    Summary="lorem sdfds sdfs dfds",
                    CreatedOn = DateTime.Now
                },
                new LogEntry()
                {
                    LogId=2,
                    Desc="XXXX XXXX",
                    Summary="SSSSS SSSSdfds sdSS SSS fs dfds"                    
                }
            };


            foreach(var i in list)
            //loop for each item 
            {
                paragraph = section.AddParagraph();
                paragraph.Style = "TOC";
                Hyperlink hyperlink = paragraph.AddHyperlink(i.LogId.ToString());
                hyperlink.AddText(i.Desc + "\t");
                hyperlink.AddPageRefField(i.LogId.ToString());
            }

            DefineContentSection(document);

            foreach (var i in list)
            //loop for each item 
            {
                //for each item do this
                Paragraph paragraph_S = document.LastSection.AddParagraph(i.Desc, "Heading1");
                paragraph_S.AddBookmark(i.LogId.ToString());
                Table table = new Table();
                table.Borders.Width = 0.75;

                Column column = table.AddColumn("16cm");
                column.Format.Alignment = ParagraphAlignment.Left;

                Row row = table.AddRow();
               
                paragraph = row.Cells[0].AddParagraph();
                paragraph.AddFormattedText("hello", TextFormat.Bold);
                paragraph.AddFormattedText(" by ", TextFormat.Italic);
                paragraph.AddLineBreak();
                paragraph.AddText("XXXXX ");

                row = table.AddRow();
                paragraph = row.Cells[0].AddParagraph();
                paragraph.AddFormattedText("hellosdf", TextFormat.Bold);
                paragraph.AddFormattedText(" by ssdfds", TextFormat.Italic);
                paragraph.AddText(i.CreatedOn?.ToString("MM/dd/yyyy h:mm tt") ?? "");


                table.SetEdge(0, 0, 1, table.Rows.Count, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

                document.LastSection.Add(table);
                
            }
            return document;
        }

        /// <summary>
        /// Defines the styles used in the document.
        /// </summary>
        public static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Times New Roman";

            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            // in PDF.

            style = document.Styles["Heading1"];
            style.Font.Name = "Tahoma";
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.Font.Color = Colors.DarkBlue;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading3"];
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 3;

            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called TextBox based on style Normal
            style = document.Styles.AddStyle("TextBox", "Normal");
            style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            style.ParagraphFormat.Borders.Width = 2.5;
            style.ParagraphFormat.Borders.Distance = "3pt";
            style.ParagraphFormat.Shading.Color = Colors.SkyBlue;

            // Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Colors.Blue;

            // Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("BOLD", "Normal");            
            style.ParagraphFormat.Font.Bold = true;

            // Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("NOR", "Normal");
            style.ParagraphFormat.Font.Bold = false;

            
        }

        /// <summary>
        /// Defines the cover page.
        /// </summary>
        public static void DefineCover(Document document)
        {
            Section section = document.AddSection();

            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.SpaceAfter = "3cm";

            //Image image = section.AddImage("../../images/Logo landscape.png");
            //image.Width = "10cm";

            paragraph = section.AddParagraph("SerPro Data Migration \n Case Notes Summary");
            paragraph.Format.Font.Size = 16;
            paragraph.Format.Font.Color = Colors.DarkRed;
            paragraph.Format.SpaceBefore = "8cm";
            paragraph.Format.SpaceAfter = "3cm";

            paragraph = section.AddParagraph("Rendering date: ");
            paragraph.AddDateField();
        }
              
        /// <summary>
        /// Defines page setup, headers, and footers.
        /// </summary>
        static void DefineContentSection(Document document)
        {
            Section section = document.AddSection();            
            section.PageSetup.StartingNumber = 1;

            HeaderFooter header = section.Headers.Primary;
            header.Format.Alignment = ParagraphAlignment.Right;
            header.AddParagraph("\t SerPro Data Migration ");


            // Create a paragraph with centered page number. See definition of style "Footer".
            Paragraph paragraph = new Paragraph();
            paragraph.AddTab();
            paragraph.AddPageField();

            // Add paragraph to footer for odd pages.
            section.Footers.Primary.Add(paragraph);            
        }
        

        public class LogEntry
        {
            public int LogId { get; set; }
            public string Desc { get; set; }
            public string Summary { get; set; }
            public DateTime? CreatedOn { get; set; }
        }


    }
}
