using HtmlToOpenXml;
using Markdig;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using ppt_lib;

namespace Ppt_lib
{
    public class DgPpt
    {

        public async static Task md_to_ppt(String md, Stream outputStream)
        {
            MarkdownPipeline pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

            var html = Markdown.ToHtml(md, pipeline);
            using (PresentationDocument presentationDocument = PresentationDocument.Create(outputStream, PresentationDocumentType.Presentation, true))
            {
                PresentationPart presentationPart = presentationDocument.AddPresentationPart();
                presentationPart.Presentation = new Presentation();
                
                CreatePresentationDocument.CreatePresentationParts(presentationPart);


                //HtmlConverter converter = new HtmlConverter(mainPart);
                //converter.ParseHtml(html);
                presentationDocument .Save();
            }
        }

        public async static Task ppt_to_md(Stream infile, Stream outfile, String name = "")
        { }
    }
}
