using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace XlsxToXmlConvertor.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConvertorController : ControllerBase
    {
        // GET: api/<ConvertorController>
        [HttpPost]
        public async Task<IActionResult> Get(IFormFile formFile)
        {
            DataTable dt = new DataTable();
            var ms = new MemoryStream();
            formFile.CopyTo(ms);          
            using (Stream inputStream = new MemoryStream(ms.ToArray()))
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    IWorkbook workbook = application.Workbooks.Open(inputStream);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    dt = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);

                }
            }

            string xml = @"<?xml version=""1.0"" encoding=""utf-8"" ?><resources>";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                xml += @"<string name=""" + dt.Rows[i][0];
                xml += @""">" + dt.Rows[i][1] + "</string>";

            }
            xml += @"</resources>";
            return Ok(xml);
        }


    }
}
