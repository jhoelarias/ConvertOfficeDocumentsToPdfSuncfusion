using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.OfficeChartToImageConverter;
using System.Collections.Generic;
using System.IO;
using System.Web.Http;

namespace ConvertOfficeDocumentsToPdfSuncfusion.Controllers
{
    public class ValuesController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get()
        {
            var mappedPath = System.Web.Hosting.HostingEnvironment.MapPath("~/bin/Docs/");

            var directory = new DirectoryInfo(mappedPath);

            var files = directory.GetFiles("*.docx");

            foreach (var file in files)
            {
                //Loads an existing Word document
                var wordDocument = new WordDocument($"{mappedPath}{file.Name}", FormatType.Docx)
                {
                    //Initializes the ChartToImageConverter for converting charts during Word to pdf conversion
                    ChartToImageConverter = new ChartToImageConverter()
                };

                //Creates an instance of the DocToPDFConverter
                var converter = new DocToPDFConverter();
                //Converts Word document into PDF document
                var pdfDocument = converter.ConvertToPDF(wordDocument);
                //Saves the PDF file
                pdfDocument.Save($"{mappedPath}{file.Name.Replace(file.Extension, "")}.pdf");

                //Closes the instance of document objects
                pdfDocument.Close(true);
                wordDocument.Close();
            }
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}