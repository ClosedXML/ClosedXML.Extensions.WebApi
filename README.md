# ClosedXML.Extensions.WebApi
WebApi Extensions for [ClosedXML](https://github.com/ClosedXML/ClosedXML)

## Install via NuGet

To install ClosedXML.Extensions.WebApi, run the following command in the Package Manager Console

```
PM> Install-Package ClosedXML.Extensions.WebApi
```
or
```
dotnet add package ClosedXML.Extensions.WebApi
```


## Usage
In your WebApi controller define an action that will generate and download your file:

```c#
public class ExcelController : ApiController
{
    [HttpGet]
    [Route("api/file/{id}")]
    public async Task<HttpResponseMessage> DownloadFile(int id)
    {
        var wb = await BuildExcelFile(id);
        return wb.Deliver("excelfile.xlsx");
    }

    private async Task<XLWorkbook> BuildExcelFile(int id)
    {
        //Creating the workbook
        var t = Task.Run(() =>
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(id);

            return wb;
        });

        return await t;
    }
}
```
