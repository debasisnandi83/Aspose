using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Facades;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AsposeTest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AsposeController : ControllerBase
    {

        [HttpGet("CreatePDF")]
        public IActionResult CreatePDF()
        {
            string dataDir = @"D:\Study Material\Aspose\App\Aspose\AsposePDF\";
            List<String> fileList = new List<String>();

            // Initialize document object
            Document document = new Document();
            // Add page
            Page page = document.Pages.Add();

            CreateTable(page);

            // Save updated PDF
            document.Save(dataDir + "file1.pdf");
            fileList.Add(dataDir + "file1.pdf");

            Document document1 = new Document();
            // Add page
            Page page1 = document1.Pages.Add();

            CreateTable(page1);

            // Save updated PDF
            document1.Save(dataDir + "file2.pdf");
            fileList.Add(dataDir + "file2.pdf");

            PdfFileEditor pdfEditor = new PdfFileEditor();
            // merge files
            pdfEditor.Concatenate(fileList.ToArray(), dataDir+"merged.pdf");

            return Ok();
        }

        private void CreateTable(Page page)
        {
            // Instantiate a table object
            Aspose.Pdf.Table tab1 = new Aspose.Pdf.Table();
            // Add the table in paragraphs collection of the desired section
            page.Paragraphs.Add(tab1);

            // Set with column widths of the table
            tab1.ColumnWidths = "50 50 50";
            tab1.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;

            // Set default cell border using BorderInfo object
            tab1.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1F);

            // Set table border using another customized BorderInfo object
            tab1.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 1F);
            // Create MarginInfo object and set its left, bottom, right and top margins
            Aspose.Pdf.MarginInfo margin = new Aspose.Pdf.MarginInfo();
            margin.Top = 5f;
            margin.Left = 5f;
            margin.Right = 5f;
            margin.Bottom = 5f;

            // Set the default cell padding to the MarginInfo object
            tab1.DefaultCellPadding = margin;

            // Create rows in the table and then cells in the rows
            Aspose.Pdf.Row row1 = tab1.Rows.Add();
            row1.Cells.Add("col1");
            row1.Cells.Add("col2");
            row1.Cells.Add("col3");
            Aspose.Pdf.Row row2 = tab1.Rows.Add();
            row2.Cells.Add("item1");
            row2.Cells.Add("item2");
            row2.Cells.Add("item3");

            CreateHyperLink(page);
        }

        private void CreateHyperLink(Page page)
        {
            LinkAnnotation link = new LinkAnnotation(page, new Aspose.Pdf.Rectangle(100, 600, 300, 300));
            // Create border object for LinkAnnotation
            Border border = new Border(link);
            // Set the border width value as 0
            border.Width = 0;
            // Set the border for LinkAnnotation
            link.Border = border;
            // Specify the link type as remote URI
            link.Action = new GoToURIAction("www.aspose.com");
            // Add link annotation to annotations collection of first page of PDF file
            page.Annotations.Add(link);

            // Create Free Text annotation
            FreeTextAnnotation textAnnotation = new FreeTextAnnotation(page, new Aspose.Pdf.Rectangle(100, 600, 300, 300), new DefaultAppearance(Aspose.Pdf.Text.FontRepository.FindFont("TimesNewRoman"), 14, System.Drawing.Color.Blue));
            // String to be added as Free text
            textAnnotation.Contents = "Link to Aspose website";
            // Set the border for Free Text Annotation
            textAnnotation.Border = border;
            // Add FreeText annotation to annotations collection of first page of Document
            page.Annotations.Add(textAnnotation);
        }
    }
}
