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
                HtmlContent = HtmlContent,
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "ReportStyles.css") },
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

        public static byte[] ConvertDocument(HtmlToPdfDocument doc)
        {
            var converter = new BasicConverter(new PdfTools());
            return converter.Convert(doc);
        }

        public HtmlToPdfDocument AddPageToPDF(string HtmlContent, string plantName, HtmlToPdfDocument doc = null, bool useCss = true)
        {
            if(doc == null)
            {
                doc = CreateHtmlToPdfDocument();
            }

            var page = new ObjectSettings
            {
                PagesCount = true,
                HtmlContent = HtmlContent,
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "ReportStyles.css") },
                HeaderSettings = { FontName = "Arial", FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                FooterSettings = { FontName = "Arial", FontSize = 9, Line = true, Center = "Plant Sale - 2020" }
            };

            if (!useCss) { page.WebSettings = null; }

            doc.Objects.Add(page);
            return doc;

        }

        private GlobalSettings CreateGlobals()
        {
            var globalSettings = new GlobalSettings
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4,
                Margins = new MarginSettings { Top = 10 },
                DocumentTitle = string.Format("Plant Sale - 2020"),
            };

            return globalSettings;
        }

        private HtmlToPdfDocument CreateHtmlToPdfDocument(GlobalSettings setting = null)
        {
            if(setting == null)
            {
                setting = CreateGlobals();
            }

            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = setting,
            };

            return doc;
        }
    }
}