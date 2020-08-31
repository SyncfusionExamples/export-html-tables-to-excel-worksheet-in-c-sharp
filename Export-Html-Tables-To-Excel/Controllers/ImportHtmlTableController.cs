using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public ActionResult ImportHtmlTable(string button)
        {
            string basePath = _hostingEnvironment.WebRootPath;

            if (button == null)
                return View();
            else if (button == "Input Template")
            {
                Stream ms = new FileStream(basePath + @"/Import-HTML-Table.html", FileMode.Open, FileAccess.Read);
                return File(ms, "text/html", "Import-HTML-Table.html");
            }
            else
            {
                // The instantiation process consists of two steps.
                // Step 1 : Instantiate the spreadsheet creation engine.
                ExcelEngine excelEngine = new ExcelEngine();

                // Step 2 : Instantiate the excel application object.
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                // A workbook is created.
                IWorkbook workbook = application.Workbooks.Create(1);

                // The first worksheet object in the worksheets collection is accessed.
                IWorksheet worksheet = workbook.Worksheets[0];

                Stream file = new FileStream(basePath + @"/Import-HTML-Table.html", FileMode.Open, FileAccess.Read);

                //Imports HTML table into the worksheet from first row and first column
                worksheet.ImportHtmlTable(file, 1, 1);

                worksheet.UsedRange.AutofitColumns();				
				worksheet.UsedRange.AutofitRows();

                MemoryStream ms = new MemoryStream();
                workbook.SaveAs(ms);
                ms.Position = 0;

                excelEngine.Dispose();
                    
                return File(ms, "Application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Import-HTML-Table.xlsx");
            }
        }
    }
    
}