using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using ClosedXML.Excel;

namespace demoExcel.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        IConfiguration _iConfiguration;
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger, IConfiguration Configuration)
        {
            _logger = logger;
            _iConfiguration = Configuration;
        }

        [HttpGet]
        public IEnumerable<WeatherForecast> Get()
        {
            var rng = new Random();
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = rng.Next(-20, 55),
                Summary = Summaries[rng.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpPost]
        public void Post()
        {
            string pathFile = AppDomain.CurrentDomain.BaseDirectory + "demo1.xlsx";
            MemoryStream ms = new MemoryStream();
            SLDocument document = new SLDocument();
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("Nombre", typeof(string));
            dt.Columns.Add("Edad", typeof(int));
            dt.Columns.Add("Sexo", typeof(string));

            dt.Rows.Add("Pepe", 25, "Hombre");
            dt.Rows.Add("Laura",28, "Mujer");
            dt.Rows.Add("Juan", 20, "Hombre");
         

            document.ImportDataTable(1, 1, dt, true);
            document.SaveAs(ms);
            //se configura email a enviar con el pdf generado
            string EmailOrigen = _iConfiguration["EmailConfiguration:User"];
            // string EmailDestino = _iConfiguration["EmailConfiguration:User"];
            string EmailDestino = _iConfiguration["EmailConfiguration:User"]; ;
            ///  string Contraseña = Cryptography.Decrypt(ListConfiguraciones.Where(p => p.Codigo == 10).First().Valor);
            string Contraseña = _iConfiguration["EmailConfiguration:Password"];
            MailMessage oMailMessage = new MailMessage(EmailOrigen, EmailDestino, "Demo Excel", null);
            oMailMessage.Body = "Adjunto encontrara orden de compra para su despacho";
            oMailMessage.Subject = "Demo Excel";
            ms.Position = 0;
            oMailMessage.Attachments.Add(new Attachment(ms, "DemoFinal.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            ///  oMailMessage.IsBodyHtml = true;
            SmtpClient oSmtpCliente = new SmtpClient(_iConfiguration["EmailConfiguration:Host"]);
            oSmtpCliente.EnableSsl = Convert.ToBoolean(_iConfiguration["EmailConfiguration:IsEnabledSsl"]);
            oSmtpCliente.Port = Convert.ToInt32(_iConfiguration["EmailConfiguration:Port"]);
            oSmtpCliente.Credentials = new System.Net.NetworkCredential(EmailOrigen, Contraseña);
            oSmtpCliente.Send(oMailMessage);
            oSmtpCliente.Dispose();
            //string fileName = "Subscribers.xlsx";
            //string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Task<IActionResult>
            //return File(ms, fileType, fileName);

        }
        [HttpPost]
        [Route("GeneratePdf")]
        public void GeneratePdf()
        {
            PdfWriter writer = new PdfWriter(AppDomain.CurrentDomain.BaseDirectory+"demo.pdf");
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            Paragraph header = new Paragraph("HEADER")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(20);

            document.Add(header);
            document.Close();
        }
        [HttpPost]
        [Route("GenerateExcel")]
        public void GenerateExcel()
        {
            MemoryStream ms = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Users");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Username";
              
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = 1;
                    worksheet.Cell(currentRow, 2).Value = "David";    
                    workbook.SaveAs(ms);

                string EmailOrigen = _iConfiguration["EmailConfiguration:User"];
                string EmailDestino = _iConfiguration["EmailConfiguration:User"]; ;
                string Contraseña = _iConfiguration["EmailConfiguration:Password"];
                MailMessage oMailMessage = new MailMessage(EmailOrigen, EmailDestino, "Demo Excel", null);
                oMailMessage.Body = "Adjunto encontrara orden de compra para su despacho";
                oMailMessage.Subject = "Demo Excel";
                ms.Position = 0;
                oMailMessage.Attachments.Add(new Attachment(ms, "Demoopenxml.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                SmtpClient oSmtpCliente = new SmtpClient(_iConfiguration["EmailConfiguration:Host"]);
                oSmtpCliente.EnableSsl = Convert.ToBoolean(_iConfiguration["EmailConfiguration:IsEnabledSsl"]);
                oSmtpCliente.Port = Convert.ToInt32(_iConfiguration["EmailConfiguration:Port"]);
                oSmtpCliente.Credentials = new System.Net.NetworkCredential(EmailOrigen, Contraseña);
                oSmtpCliente.Send(oMailMessage);
                oSmtpCliente.Dispose();


            }
        }

    }
}
