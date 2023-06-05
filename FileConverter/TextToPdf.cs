using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.OpenApi.Models;
using System.Net;
using Syncfusion.Pdf;
using System.IO;
using Newtonsoft.Json;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Drawing;

namespace FileConverter
{
    public static class TextToPdf
    {
        [FunctionName("TextToPdf")]
        [OpenApiOperation()]
        [OpenApiParameter(name: "Text", Required = true, In = ParameterLocation.Query, Type = typeof(string))]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/pdf", bodyType: typeof(string))]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string text = req.Query["text"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            text = text ?? data?.text;

            using PdfDocument document = new();
            var page = document.Pages.Add();
            PdfGraphics graphics = page.Graphics;
            graphics.DrawString(text,
                new PdfStandardFont(PdfFontFamily.Helvetica, 20),
                PdfBrushes.Black,
                new PointF(0, 0));

            using MemoryStream outputPdfStream = new MemoryStream();
            document.Save(outputPdfStream);
            outputPdfStream.Position = 0;
            document.Dispose();

            string contentType = "application/pdf";
            string fileName = "document.pdf";

            req.HttpContext.Response.Headers.Add("Content-Disposition", $"attachment;{fileName}");

            return new FileContentResult(outputPdfStream.ToArray(), contentType);
        }
    }
}
