using DinkToPdf;
using DinkToPdf.Contracts;
using System.IO;

namespace Plant_Sale_Tool
{
    //[Route("api/pdfcreator")]
    //[ApiController]
    public class PdfCreatorController
    {
        private IConverter _converter;

        public PdfCreatorController()
        {

        }

        public PdfCreatorController(IConverter converter)
        {
            _converter = converter;
        }

        //[HttpGet]
        public byte[] CreatePDF(string HtmlContent, string plantName)
        {
            var globalSettings = new GlobalSettings
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4,
                Margins = new MarginSettings { Top = 10 },
                DocumentTitle = string.Format("{0} -  2020 Plant Sale", plantName),
                //Out = @"D:\PDFCreator\Employee_Report.pdf"
                //Out = string.Format("{0}.pdf", plantName)
            };

            var objectSettings = new ObjectSettings
            {
                PagesCount = true,
                //HtmlContent = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. In consectetur mauris eget ultrices  iaculis. Ut
                //odio viverra, molestie lectus nec, venenatis turpis.",
                HtmlContent = HtmlContent,
                //WebSettings = { DefaultEncoding = "utf-8" },
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "ReportStyles.css") },
                //HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true, Spacing = 2.812 }
                HeaderSettings = { FontName = "Arial", FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                FooterSettings = { FontName = "Arial", FontSize = 9, Line = true, Center = "Report Footer" }
            };

            var converter = new BasicConverter(new PdfTools());

            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = globalSettings,
                Objects = { objectSettings }
                
            };
            byte[] pdf = converter.Convert(doc);

            return pdf;
        }
    }
}