using ClosedXML.Excel;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace ClosedXML.Extensions
{
    public static class WebApiExtensions
    {
        public static HttpResponseMessage Deliver(this IXLWorkbook workbook, string fileName, string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            var memoryStream = new MemoryStream();
            workbook.SaveAs(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);

            var message = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };
            message.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            message.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };

            return message;
        }
    }
}