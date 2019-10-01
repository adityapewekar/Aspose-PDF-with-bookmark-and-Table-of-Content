using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Aspose_GeneratePDF_ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Program obj = new Program();
            string pdfFilePath = obj.PdfAppendTest();
            Console.WriteLine(pdfFilePath);
            Console.ReadLine();
        }

        public string PdfAppendTest()
        {
            int pdfPageNo = 2;
            List<TOC> toc = new List<TOC>();


            var AsposeLicensePath =Directory.GetCurrentDirectory()+ "//PdfFile//";
            if (!Directory.Exists(AsposeLicensePath))
            {
                Directory.CreateDirectory(AsposeLicensePath);
            }

            //aspose words license
            //var AsposeLicenseFilePath = "C://Users//Aditya//.nuget//packages//aspose.pdf//19.9.0//lib//net4.0 - client//Aspose.Pdf.lic";

            //if (System.IO.File.Exists(AsposeLicenseFilePath))
            //{
            //    Aspose.Pdf.License WordsLic = new Aspose.Pdf.License();
            //    WordsLic.SetLicense(AsposeLicenseFilePath);
            //}

            if (File.Exists(Directory.GetCurrentDirectory() + "//PdfFile///merged.pdf"))
            {
                File.Delete(Directory.GetCurrentDirectory() + "//PdfFile///merged.pdf");
            }

            List<String> inputDocs = new List<String>();
            foreach (string file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.pdf"))
            {
                inputDocs.Add(file);
            }

            Aspose.Pdf.Document targetDoc = new Aspose.Pdf.Document();
            Aspose.Pdf.Document doc;

            for (int i = 0; i < inputDocs.Count; i++)
            {
                doc = new Aspose.Pdf.Document(inputDocs[i]);
                toc.Add(new TOC
                {
                    BookmarkName=Path.GetFileName(doc.FileName),
                    PageNo=pdfPageNo
                });

                // Add the pages of the source documents to the target document
                targetDoc.Pages.Add(doc.Pages);

                #region Code to generate Blank page between two documents
                Aspose.Pdf.Page pageI = targetDoc.Pages.Insert(pdfPageNo);
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

            Aspose.Pdf.Page tocPage = targetDoc.Pages.Insert(1);
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
                Aspose.Pdf.Heading heading2 = new Aspose.Pdf.Heading(1);

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

            targetDoc.Save(Directory.GetCurrentDirectory() + "//PdfFile//merged.pdf");

            #region Code to bookmark the pages
            // For complete examples and data files, please go to https://github.com/aspose-pdf/Aspose.Pdf-for-.NET
            // The path to the documents directory.
            string dataDir = Directory.GetCurrentDirectory() + "//PdfFile//";

            // Open document
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(dataDir + "merged.pdf");
            // Bookmark name list
            string[] bookmarkList = { "First" };
            // Page list
            int[] pageList = { 1 };
            // Create bookmark of a range of pages

            foreach(var item in toc)
                bookmarkEditor.CreateBookmarkOfPage(item.BookmarkName, item.PageNo);

            // Save updated PDF file
            bookmarkEditor.Save(dataDir + "merged.pdf");
            #endregion

            return Directory.GetCurrentDirectory() + "//PdfFile//merged.pdf";

        }
    }

    public class TOC
    {
        public string BookmarkName { get; set; }
        public int PageNo { get; set; }
    }
}
