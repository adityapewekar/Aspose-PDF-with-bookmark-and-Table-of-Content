using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace Aspose_GeneratePDF_ConsoleApp
{
    class Program
    {
        private const string DocumentFolderPath = "F:\\Project\\Aspose_GeneratePDF_ConsoleApp\\Aspose-PDF-with-bookmark-and-Table-of-Content\\Documents\\";
        private const string PDFFolderPath = "F:\\Project\\Aspose_GeneratePDF_ConsoleApp\\Aspose-PDF-with-bookmark-and-Table-of-Content\\GeneratedPDF\\";
        private const string PDFFileName = "MergedPDF.pdf";
        private readonly string PDFFilePath;

        Program()
        {
            PDFFilePath = PDFFolderPath + PDFFileName;
        }
        
        static void Main(string[] args)
        {
            Program obj = new Program();
            string pdfFilePath = obj.GeneratePDF();

            // Opens the folder of PDF generated file
            Process.Start(new ProcessStartInfo()
            {
                FileName = PDFFolderPath,
                UseShellExecute = true,
                Verb = "open"
            });
            Console.WriteLine(pdfFilePath);
            Console.ReadLine();
        }

        /// <summary>
        /// Responsible to merge different pdf documents and generate into one single PDF
        /// along with the bookmark and table of content 
        /// </summary>
        /// <returns> Path of the PDF document </returns>
        public string GeneratePDF()
        {
            // Used to keep a track of current generated PDF page number
            int pdfPageNo = 2;

            // Used to create bookmark and Table of content
            List<TOC> toc = new List<TOC>();

            if (!Directory.Exists(PDFFolderPath))
            {
                Directory.CreateDirectory(PDFFolderPath);
            }

            //aspose words license
            //var AsposeLicenseFilePath = "C://Users//Aditya//.nuget//packages//aspose.pdf//19.9.0//lib//net4.0 - client//Aspose.Pdf.lic";
            //if (System.IO.File.Exists(AsposeLicenseFilePath))
            //{
            //    Aspose.Pdf.License WordsLic = new Aspose.Pdf.License();
            //    WordsLic.SetLicense(AsposeLicenseFilePath);
            //}

            if (File.Exists(PDFFilePath))
            {
                File.Delete(PDFFilePath);
            }

            // Used to maintaint the list of all the documents to be merged.
            List<String> inputDocs = new List<String>();

            // Fetches all the documents from the specified DocumentFolderPath
            foreach (string file in Directory.EnumerateFiles(DocumentFolderPath, "*.pdf"))
            {
                inputDocs.Add(file);
            }

            Document targetDoc = new Document();
            Document doc;

            for (int i = 0; i < inputDocs.Count; i++)
            {
                doc = new Document(inputDocs[i]);

                toc.Add(new TOC
                {
                    BookmarkName=Path.GetFileName(doc.FileName),
                    PageNo=pdfPageNo
                });

                // Add the pages of the source documents to the target document
                targetDoc.Pages.Add(doc.Pages);

                #region Code to generate Blank page between two documents
                Page pageI = targetDoc.Pages.Insert(pdfPageNo);
                TextFragment textI = new TextFragment("Blank Page");
                pageI.Paragraphs.Add(textI);

                // Create a bookmark object
                OutlineItemCollection pdfOutlineI = new OutlineItemCollection(targetDoc.Outlines);
                pdfOutlineI.Title = "Blank Page";
                pdfOutlineI.Italic = true;
                pdfOutlineI.Bold = true;

                // Add bookmark in the document's outline collection.
                targetDoc.Outlines.Add(pdfOutlineI);
                pdfPageNo = pdfPageNo + doc.Pages.Count;
                pdfPageNo++;
                #endregion
            }


            #region Code to generate Table of Content
            string html = "";// above html
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            Page tocPage = targetDoc.Pages.Insert(1);
            tocPage.Background = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White);

            // Create object to represent TOC information
            TocInfo tocInfo = new TocInfo();
            TextFragment title = new TextFragment("Table Of Contents");
            title.TextState.FontSize = 20;
            title.TextState.FontStyle = FontStyles.Bold;

            // Set the title for TOC
            tocInfo.Title = title;
            tocPage.TocInfo = tocInfo;


            foreach(var item in toc)
            {
                // Create Heading object
                Heading heading2 = new Heading(1);

                TextSegment segment2 = new TextSegment();
                segment2 = new TextSegment();
                heading2.TocPage = tocPage;
                heading2.Segments.Add(segment2);

                // Specify the destination page for heading object
                heading2.DestinationPage = targetDoc.Pages[item.PageNo];

                // Destination page
                heading2.Top = targetDoc.Pages[item.PageNo].Rect.Height;

                // Destination coordinate
                segment2.Text = item.BookmarkName;

                // Add heading to page containing TOC
                tocPage.Paragraphs.Add(heading2);
            }
            #endregion

            targetDoc.Save(PDFFilePath);

            #region Code to bookmark the pages
            // For complete examples and data files, please go to https://github.com/aspose-pdf/Aspose.Pdf-for-.NET

            // Open document
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(PDFFilePath);

            // Create bookmark of a range of pages
            foreach(var item in toc)
                bookmarkEditor.CreateBookmarkOfPage(item.BookmarkName, item.PageNo);

            // Save updated PDF file
            bookmarkEditor.Save(PDFFilePath);
            #endregion

            return PDFFilePath;

        }
    }

    public class TOC
    {
        public string BookmarkName { get; set; }
        public int PageNo { get; set; }
    }
}
