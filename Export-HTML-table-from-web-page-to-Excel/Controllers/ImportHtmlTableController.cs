using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;

namespace ImportHtmlTable.Controllers
{
    public class ImportHtmlTableController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public ImportHtmlTableController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public ActionResult ImportHtmlTable(string button, string tableHTML)
        {
            if (button == null)
                return View();

            MemoryStream ms = new MemoryStream();

            // The instantiation process consists of two steps.
            // Step 1 : Instantiate the spreadsheet creation engine.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                // Step 2 : Instantiate the excel application object.
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                // A workbook is created.
                IWorkbook workbook = application.Workbooks.Create(1);

                // The first worksheet object in the worksheets collection is accessed.
                IWorksheet worksheet = workbook.Worksheets[0];

                byte[] byteArray = Encoding.UTF8.GetBytes(tableHTML);

                MemoryStream file = new MemoryStream(byteArray);

                // Imports HTML table into the worksheet from first row and first column
                worksheet.ImportHtmlTable(file, 1, 1);

                worksheet.UsedRange.AutofitColumns();
                worksheet.UsedRange.AutofitRows();

                workbook.SaveAs(ms);
                ms.Position = 0;
            }

            return File(ms, "Application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Export-HTML-Table.xlsx");
        }
    }
    
}