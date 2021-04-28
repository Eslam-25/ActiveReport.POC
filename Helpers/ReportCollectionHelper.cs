using ActiveReport.POC.Models;
using System.Collections.Generic;

namespace ActiveReport.POC.Helpers
{
    public class ReportCollectionHelper
    {
        public static IEnumerable<ReportModel> GetReportModels()
        {
            var reportModels = new List<ReportModel>();

            reportModels.Add(new ReportModel { Id = 1, Title = "Report 1", SubTitle = "SubTitle 1", IsActive = true });
            reportModels.Add(new ReportModel { Id = 2, Title = "Report 2",  IsActive = true });
            reportModels.Add(new ReportModel { Id = 3, Title = "Report 3", SubTitle = "SubTitle 3", IsActive = false });

            return reportModels;
        }
    }
}
