using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebUIAutomation.POM
{
    internal class excelreadwrite
    {
        Code

using ClosedXML.Excel;
public bool CreateDailyReportTemplate(string type, string runEnv = null)
    {
        string reportPath = GetFilePath(type, runEnv);

        DirectoryInfo directory = new DirectoryInfo(ReportsBasePath);

        if (!directory.Exists)
        {
            directory.Create();
        }

        var archivePath = $@"{ReportsBasePath}\Archive";
        var archiveDirectory = new DirectoryInfo(archivePath);
        if (!archiveDirectory.Exists)
        {
            archiveDirectory.Create();
        }

        FileInfo[] files = directory.GetFiles();
        foreach (var file in files)
        {
            if (file.FullName.Contains(reportPath)) // skip existing path
                continue;

            if (file.Name.Contains(GetFileName("Team"))) // To make sure we don't move the team automation file of current date
                continue;

            if (file.Name.Contains(GetFileName("Test"))) // To make sure we don't move the team automation file of current date
                continue;
            else
            {
                if (/*file.Name.Contains("Test") || (file.Name.Contains("Team") &&*/ file.Name.Contains(DateTime.Now.ToString("MM-dd-yyyy")))
                    continue;
                file.MoveTo($@"{archivePath}\{file.Name}");
            }

        }

        if (File.Exists(reportPath))
        {
            Console.WriteLine("File already exists" + reportPath);
            return true;
        }

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Sheet 1");

            worksheet.Cell(1, SubmissionIdColumn).Value = "SubmissionID";
            worksheet.Cell(1, UpdatedSubmissionIDColumn).Value = "UpdatedSubmissionID";
            worksheet.Cell(1, SubmissionTypeColumn).Value = "SubmissionType";
            worksheet.Cell(1, EnvironmentColumn).Value = "Environment";
            worksheet.Cell(1, UpdatedTimeColumn).Value = "UpdatedTime";
            worksheet.Cell(1, G3PStatusColumn).Value = "G3PStatus";
            worksheet.Cell(1, AALColumn).Value = "AAL";
            worksheet.Cell(1, SevCatColumn).Value = "SevCat";
            worksheet.Cell(1, NonSevCatColumn).Value = "NonSevCat";
            worksheet.Cell(1, NonCatColumn).Value = "NonCat";
            worksheet.Cell(1, BenchmarkColumn).Value = "Benchmark";
            worksheet.Cell(1, BenchmarkPremiumColumn).Value = "BenchmarkPremium";
            worksheet.Cell(1, PriceComparisonColumn).Value = "PriceComparison";
            worksheet.Cell(1, DescriptionColumn).Value = "Description";
            worksheet.Cell(1, CommentsColumn).Value = "Comments";
            worksheet.Cell(1, LossReportsColumn).Value = "LossReports";
            worksheet.Cell(1, LossReportCommentsColumn).Value = "LossReportComments";
            worksheet.Cell(1, LossCostModellingColumn).Value = "LossCostModelling";
            worksheet.Cell(1, LossCostModellingCommentsColumn).Value = "LossCostModellingComments";
            worksheet.Cell(1, PricingReportsColumn).Value = "PricingReports";
            worksheet.Cell(1, PricingReportCommentsColumn).Value = "PricingReportComments";
            worksheet.Cell(1, AdminValidationColumn).Value = "AdminValidation";
            worksheet.Cell(1, G3PorEWHColumn).Value = "G3P/EWH";
            worksheet.Cell(1, AddLayerColumn).Value = "AddLayer";
            worksheet.Cell(1, DeleteLayerColumn).Value = "DeleteLayer";

            workbook.SaveAs(reportPath);
        }

        Console.WriteLine("File created successfully" + reportPath);
        return true;

        Code 2

 public void InsertExcel(Dictionary<string, string> ActualSubmissionIds, string type, string env, string TiV, string BusinessType, string runEnv, string basePath = null)
        {
            try
            {
                string filePath = GetFilePath(type, runEnv);

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault();

                    int currentRow = worksheet.LastRowUsed().RowNumber() + 1;
                    foreach (var kvp in ActualSubmissionIds)
                    {
                        string actualsubmission = kvp.Key;
                        string UWSubmission = kvp.Value;

                        if (TiV.ToLowerInvariant().Contains("both"))
                        {
                            InsertRow(worksheet, currentRow, type, actualsubmission, UWSubmission, env, TiV, "Calc", BusinessType);
                            currentRow++;
                            InsertRow(worksheet, currentRow, type, actualsubmission, UWSubmission, env, TiV, "Org", BusinessType);
                            currentRow++;
                        }
                        else
                        {
                            InsertRow(worksheet, currentRow, type, actualsubmission, UWSubmission, env, TiV, TiV, BusinessType);
                            currentRow++;
                        }
                    }

                    [workbook.Save] (https://workbook.Save)();
         }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void InsertRow(IXLWorksheet worksheet, int rowNum, string type, string actualSubmission, string UWSubmission, string env, string tiv, string submissionType, string BusinessType)
        {
            worksheet.Cell(rowNum, SubmissionIdColumn).Value = actualSubmission;
            worksheet.Cell(rowNum, UpdatedSubmissionIDColumn).Value = UWSubmission;
            worksheet.Cell(rowNum, G3PorEWHColumn).Value = BusinessType;
            worksheet.Cell(rowNum, EnvironmentColumn).Value = env;
            worksheet.Cell(rowNum, UpdatedTimeColumn).Value = DateTime.Now.ToString("HH:mm");
            worksheet.Cell(rowNum, G3PStatusColumn).Value = "NA";
            worksheet.Cell(rowNum, AALColumn).Value = "NA";
            worksheet.Cell(rowNum, SevCatColumn).Value = "NA";
            worksheet.Cell(rowNum, NonSevCatColumn).Value = "NA";
            worksheet.Cell(rowNum, NonCatColumn).Value = "NA";
            worksheet.Cell(rowNum, BenchmarkColumn).Value = "NA";
            worksheet.Cell(rowNum, BenchmarkPremiumColumn).Value = "NA";
            worksheet.Cell(rowNum, SubmissionTypeColumn).Value = submissionType;
            worksheet.Cell(rowNum, DescriptionColumn).Value = "NA";
            worksheet.Cell(rowNum, PriceComparisonColumn).Value = "NA";
            worksheet.Cell(rowNum, CommentsColumn).Value = "NA";
            worksheet.Cell(rowNum, LossReportsColumn).Value = " ";
            worksheet.Cell(rowNum, LossReportCommentsColumn).Value = " ";
            worksheet.Cell(rowNum, LossCostModellingColumn).Value = " ";
            worksheet.Cell(rowNum, LossCostModellingCommentsColumn).Value = " ";
            worksheet.Cell(rowNum, PricingReportsColumn).Value = " ";
            worksheet.Cell(rowNum, PricingReportCommentsColumn).Value = " ";
            worksheet.Cell(rowNum, AdminValidationColumn).Value = " ";
            worksheet.Cell(rowNum, AddLayerColumn).Value = " ";
            worksheet.Cell(rowNum, DeleteLayerColumn).Value = " ";
        }

        Code 3

 public void UpdateExcel(string type, Dictionary<string, string> ValuesfromSOV, string UVsubmissionName, string submissionType, string runEnv)
        {
            string filePath = GetFilePath(type, runEnv);
            submissionType = submissionType.Contains("OrgT") ? "Org" : submissionType;
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var row = worksheet.Rows().FirstOrDefault(x => x.Cell(UpdatedSubmissionIDColumn).Value.Equals(UVsubmissionName) && x.Cell(SubmissionTypeColumn).Value.Equals(submissionType));

                foreach (var kvp in ValuesfromSOV)
                {
                    var column = kvp.Key;
                    var value = kvp.Value;
                    if (!value.Contains("NA"))
                    {
                        row.Cell(Columns[column]).Value = value;
                    }
                }

                [workbook.Save] (https://workbook.Save)();
     }
        }

        Code 5

  public List<Dictionary<string, string>> SelectFromExcelReport(string type, string G3PStatus, string runEnv = null)
        {
            string filePath = GetFilePath(type, runEnv);

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();

                var rows = worksheet
                    .Rows()
                    .Where(x => x.Cell(G3PStatusColumn).Value.ToString() == G3PStatus);

                return rows.Select(x => SelectFromRow(x)).ToList();
            }
        }
    }
}
}
