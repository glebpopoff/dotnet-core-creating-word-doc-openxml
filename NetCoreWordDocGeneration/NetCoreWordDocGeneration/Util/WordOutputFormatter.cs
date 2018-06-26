using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Formatters;
using Microsoft.Net.Http.Headers;
using NetCoreWordDocGeneration.Dtos;

namespace NetCoreWordDocGeneration.Util
{
    public class WordOutputFormatter : OutputFormatter
    {
        public string ContentType { get; }

        public WordOutputFormatter()
        {
            ContentType = "application/ms-word";
            SupportedMediaTypes.Add(MediaTypeHeaderValue.Parse(ContentType));
        }

        public override bool CanWriteResult(OutputFormatterCanWriteContext context)
        {
            return context.Object is DemoDto;
        }

        // this needs to be overwritten
        public override async Task WriteResponseBodyAsync(OutputFormatterWriteContext context)
        {
            var response = context.HttpContext.Response;
            var filePath = string.Format("./DataExport/myfile-{0}.docx", DateTime.Now.Ticks);
            var templatePath = string.Format("./DataExport/my-template.docx");

            var viewModel = context.Object as DemoDto;

            //open the template then save it as another file (while also stream it to the user)

            byte[] byteArray = File.ReadAllBytes(templatePath);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);

                //to create a new document
                // using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {

                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var paras = body.Elements<Paragraph>();

                    //search & replace string
                    foreach (var para in paras)
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                if (text.Text.Trim().Contains("Welcome to Word"))
                                {
                                    text.Text = text.Text.Trim().Replace("Welcome to Word", viewModel.Welcome);
                                }
                            }
                        }
                    }

                    //another approach to search & replace string
                    foreach (var text in wordDoc.MainDocumentPart.Document.Descendants<Text>())
                    {
                        if (text.Text.Trim().Contains("Welcome to Word"))
                        {
                            text.Text = text.Text.Trim().Replace("Welcome to Word", viewModel.Welcome);
                        }
                    }

                    //append some stuff to the document
                    Paragraph p = new Paragraph();
                    Run r = new Run();
                    Text t = new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent quam augue, tempus id metus in, laoreet viverra quam. Sed vulputate risus lacus, et dapibus orci porttitor non.");
                    r.Append(t);
                    p.Append(r);
                    body.Append(p);

                    p = new Paragraph();
                    r = new Run();
                    t = new Text(viewModel.HelloWorld);
                    r.Append(t);
                    p.Append(r);
                    body.Append(p);


                    wordDoc.Close();


                }

                using (FileStream fileStream = new FileStream(filePath, System.IO.FileMode.CreateNew))
                {
                    mem.WriteTo(fileStream);
                    mem.Close();
                    fileStream.Close();
                }

                response.Headers.Add("Content-Disposition", "inline;filename=MyFile.docx");
                response.ContentType = "application/ms-word";

                await response.SendFileAsync(filePath);
            }

        }
    }
}
