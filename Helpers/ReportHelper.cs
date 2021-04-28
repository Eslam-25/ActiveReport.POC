using GrapeCity.ActiveReports;
using GrapeCity.ActiveReports.Document;
using GrapeCity.ActiveReports.Export.Pdf.Page;
using GrapeCity.ActiveReports.Export.Word.Page;
using GrapeCity.ActiveReports.Rendering.IO;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Text;
using System;
using System.Collections.Generic;
using System.IO;

namespace ActiveReport.POC.Helpers
{
    public class ReportHelper
    {
        private PageReport report { get; set; }
        private readonly IEnumerable<object> _models;
        public ReportHelper(IEnumerable<object> models)
        {
            _models = models;
        }

        public PageReport GenerateReport(string reportPath)
        {
            FileInfo filePath = new FileInfo(reportPath);
            report = new PageReport(filePath);
            PageDocument runt = new PageDocument(report);
            runt.LocateDataSource += new LocateDataSourceEventHandler(runt_LocateDataSource);

            PrepareFontCollection();
            return report;
        }
        private void runt_LocateDataSource(object sender, LocateDataSourceEventArgs args)
        {
            if (_models != null)
                args.Data = _models;
        }

        private void PrepareFontCollection()
        {
            FontCollection.SystemFonts.DefaultFont = StandardFonts.Times;
            FontCollection fontCollection = new FontCollection();
            fontCollection.RegisterDirectory("./Fonts");
            var font = fontCollection.FindFamilyName("CorpoS");
            if (font != null)
            {
                FontCollection.SystemFonts.DefaultFont = font;
            }
        }

        public string SavePDFFile(string pdfFileName)
        {
            //Output pdf file.
            string pdfFolderPath = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Downloads\\PR Report";
            DirectoryInfo outputDirectory = new DirectoryInfo(pdfFolderPath);
            outputDirectory.Create();

            GrapeCity.ActiveReports.Export.Pdf.Page.Settings pdfSetting = new GrapeCity.ActiveReports.Export.Pdf.Page.Settings();
            PdfRenderingExtension pdfRenderingExtension = new PdfRenderingExtension();
            FileStreamProvider outputProvider = new FileStreamProvider(outputDirectory, pdfFileName);

            outputProvider.OverwriteOutputFile = true;
            report.Document.Render(pdfRenderingExtension, outputProvider, pdfSetting);

            return pdfFolderPath;
        }

        public string SaveWordFile(string wordFileName)
        {
            //Output word file.
            string wordFolderPath = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Downloads\\PR Report";
            DirectoryInfo outputDirectory = new DirectoryInfo(wordFolderPath);

            outputDirectory.Create();

            GrapeCity.ActiveReports.Export.Word.Page.Settings wordSetting = new GrapeCity.ActiveReports.Export.Word.Page.Settings();
            WordRenderingExtension wordRenderingExtension = new WordRenderingExtension();
            FileStreamProvider outputProvider = new FileStreamProvider(outputDirectory, wordFileName);

            outputProvider.OverwriteOutputFile = true;
            report.Document.Render(wordRenderingExtension, outputProvider, wordSetting);

            return wordFolderPath;
        }
    }
}
