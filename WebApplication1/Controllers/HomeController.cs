using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using ExcelDataReader;

[ApiController]
[Route("api/GetData")]
public class HomeController : ControllerBase
{
    [HttpGet]
    public IActionResult ConvertExcelToJson()
    {
        string excelPath = Path.Combine(Directory.GetCurrentDirectory(), "dados.xlsx");
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using (var stream = System.IO.File.Open(excelPath, FileMode.Open, FileAccess.Read))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var jsonData = new List<Dictionary<string, object>>();
            string[] columnNames = { "Name", "Email", "Telephone", "UpdateDate" };
            reader.Read();
            while (reader.Read())
            {
                var rowData = new Dictionary<string, object>();
                for (int i = 0; i < columnNames.Length; i++)
                {
                    var cellValue = reader.GetValue(i);
                    rowData[columnNames[i]] = cellValue;
                }
                jsonData.Add(rowData);
            }
            var json = JsonConvert.SerializeObject(jsonData);
            return Ok(json);
        }
    }
}
