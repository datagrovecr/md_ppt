﻿
using Markdig;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using ppt_lib;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using Markdig.Syntax;


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


                //HtmlConverter converter = new HtmlConverter(presentationDocument);
                //converter.ParseHtml(html);
                presentationDocument .Save();
            }
        }

        public async static Task ppt_to_md(Stream infile, Stream outfile, String name = "")
        {
            PresentationDocument presDoc = PresentationDocument.Open(infile, false);
            PresentationPart presPart= presDoc.PresentationPart;
            IEnumerable <SlideMasterPart> slideMasterPart = presPart.SlideMasterParts;
            IEnumerable<SlidePart> slidePart = presPart.SlideParts;
            StringBuilder textBuilder = new StringBuilder();

           

            foreach (var slides in slidePart)
            {
                //DocumentFormat.OpenXml.Packaging.HyperlinkRelationship




                foreach (var treeBranch in slides.Slide.Descendants<ShapeTree>().FirstOrDefault())
                {


                  

                    //DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties
                    if (treeBranch is NonVisualGroupShapeProperties) { 
                    
                    }
                    //DocumentFormat.OpenXml.Presentation.GroupShapeProperties
                    if (treeBranch is GroupShapeProperties) {
                    
                    }
                    //DocumentFormat.OpenXml.Presentation.Shape
                    if (treeBranch is Shape)
                    {

                        openXmlProcessing. ProcessParagraph((Shape)treeBranch, textBuilder,slides);
                    }

                }
            }
            //var parts = wordDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
            //StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
            //if (parts != null)
            //{
                //var asd = parts.Descendants<HyperlinkList>();


                /*foreach (var block in parts.ChildElements)
                {


                    if (block is Paragraph)
                    {
                        //This method is for manipulating the style of Paragraphs and text inside
                        ProcessParagraph((Paragraph)block, textBuilder);
                    }

                    if (block is Table) ProcessTable((Table)block, textBuilder);

                }*/

            //}

            //This code is replacing the below one because I need to check the .md file faster
            //writing the .md file in test_result folder
            if (name != "")
            {
                using (var streamWriter = new StreamWriter(name + ".md"))
                {
                    String s = textBuilder.ToString();
                    streamWriter.Write(s);
                }
            }
            else
            {

                var writer = new StreamWriter(outfile);
                String s = textBuilder.ToString();
                writer.Write(s);
                writer.Flush();
            }

        }

    }
}
