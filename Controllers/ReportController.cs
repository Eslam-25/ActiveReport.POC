using ActiveReport.POC.Helpers;
using Microsoft.AspNetCore.Mvc;

namespace ActiveReport.POC.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportController : ControllerBase
    {

        [HttpGet("pdf")]
        public IActionResult GetPDFFile()
        {
            var models = ReportCollectionHelper.GetReportModels();
            ReportHelper reportHelper = new ReportHelper(models);

            var report = reportHelper.GenerateReport("Reports/report.rdlx");
            foreach (var reportParameter in report.Report.ReportParameters)
            {
                if (reportParameter.Name == "htmlCode")
                    reportParameter.DefaultValue.Values.Add(HtmlHelper.GetHtml());
            }

            var pdfFolderPath = reportHelper.SavePDFFile("PDF Report");

            return Ok("The PDF file downloaded on : " + pdfFolderPath);
        }

        [HttpGet("word")]
        public IActionResult GetWordFile()
        {
            var models = ReportCollectionHelper.GetReportModels();
            ReportHelper reportHelper = new ReportHelper(models);

            var report = reportHelper.GenerateReport("Reports/report.rdlx");
            foreach (var reportParameter in report.Report.ReportParameters)
            {
                if (reportParameter.Name == "htmlCode")
                    reportParameter.DefaultValue.Values.Add(HtmlHelper.GetHtml());
            }

            var wordFolderPath = reportHelper.SaveWordFile("Word Report");
            return Ok("The word file downloaded on : " + wordFolderPath);
        }
    }
}
