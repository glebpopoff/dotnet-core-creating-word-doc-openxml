# dotnet-core-creating-word-doc-openxml
.NET Core 2.X Creating Word Doc Documents using OpenXML

The solution uses .NET Core 2.0 and OpenXml 2.8.1. The solution demonstrates how to create Word Documents (Docx) using Microsoft OpenXml library. The solution relies on .NET OutputFormatter to create a word document using a template (DataExport folder).

Here's a quick setup:

# Create a new project/solution
- Add DocumentFormat.OpenXml 2.8.1 NuGet package as dependency

# Create a DTO / Model class
```c#
public class DemoDto
    {
        public string Welcome { get; set; }
        public string HelloWorld { get; set; }
    }
```

# Create a WordOutputFormatter class
```c#
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
                
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {

                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var paras = body.Elements<Paragraph>();

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
```


# Register WordOutputFormatter in MVC startup options
In the Startup.cs
```c#
services.AddMvc(config =>
            {
                config.OutputFormatters.Add(new WordOutputFormatter());
                config.FormatterMappings.SetMediaTypeMappingForFormat(
                  "docx", MediaTypeHeaderValue.Parse("application/ms-word"));

            });
```

# Define your controller. Notice how the controller responds to 'application/ms-word' content type
```c#
[Route("api/[controller]")]
    [Produces("application/ms-word")]
    public class DemoController : Controller
    {
        [HttpGet("Export")]
        [Produces("application/ms-word")]
        public async Task<IActionResult> Export()
        {
            try
            {
                var demoDto = new DemoDto() { Welcome = "Lorem Ipsum", HelloWorld = "Hello World!!!" };
                return Ok(demoDto);
            }
            catch (Exception ex)
            {
                //log the exception
                return BadRequest();
            }
        }
    }
```
# Optional: add Swagger support...because it's awesome!
Add Swashbuckle.AspNetCore (2.3+) as a dependency
In your startup file (Startup.cs) ConfigureServices method, after the MVC configuration:
```c#
services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Swashbuckle.AspNetCore.Swagger.Info { Title = "My App!", Version = "v1" });

            });
```
In the Startup.cs Configure method, add the following
```c#
app.UseSwagger(c => { c.RouteTemplate = "api/swagger/{documentName}/swagger.json"; });
            app.UseSwaggerUI(c => { c.SwaggerEndpoint("v1/swagger.json", "My App Api"); c.RoutePrefix = "api/swagger"; });
```
# Test! Either use Swagger or directly navigate to your controller
Swagger: http://localhost:5000/api/swagger/index.html
Directly naviage: http://localhost:5000/api/export/demo
