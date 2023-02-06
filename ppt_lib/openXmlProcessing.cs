using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Linq.Expressions;
using Markdig.Syntax;

namespace ppt_lib
{
    internal class openXmlProcessing
    {
        public static int isInlineStile = 0;
        public static IEnumerable<HyperlinkRelationship> links = null;
        public static void ProcessParagraph(Shape treeBranch, StringBuilder textBuilder,DocumentFormat.OpenXml.Packaging.SlidePart slides)
        {
            //http://schemas.openxmlformats.org/officeDocument/2006/relationships
            links = slides.HyperlinkRelationships;
         
            string text = "";
             bool isFullCodeBlock = false;
            foreach (var element in treeBranch)
            {
              
                if (element is DocumentFormat.OpenXml.Presentation.ShapeProperties)
                {
                    DocumentFormat.OpenXml.Drawing.SolidFill fill = element.Descendants<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
                    if (fill!=null)
                    {
                        isFullCodeBlock = true;
                        textBuilder.Append("``` \n");
                    }
                }
                if (element is TextBody)
                 {
                    int orderedLits=1;
                    foreach (var item in element)
                    {
                        
                        //DocumentFormat.OpenXml.Drawing.BodyProperties
                        if (item is DocumentFormat.OpenXml.Drawing.BodyProperties){}

                        //DocumentFormat.OpenXml.Drawing.ListStyle
                        if (item is DocumentFormat.OpenXml.Drawing.ListStyle){}
                        

                       
                        //DocumentFormat.OpenXml.Drawing.Paragraph -this has the size 
                        //pull the run
                        if (item is DocumentFormat.OpenXml.Drawing.Paragraph)
                        {

                           
                            //DocumentFormat.OpenXml.Drawing.ParagraphProperties
                            DocumentFormat.OpenXml.Drawing.ParagraphProperties paragraphProperties = item.Descendants<DocumentFormat.OpenXml.Drawing.ParagraphProperties>().FirstOrDefault();
                            DocumentFormat.OpenXml.Drawing.RunProperties runProp = item.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>().FirstOrDefault();
                            
                            
                            foreach (var paragraphChild in item.ChildElements)
                            {
                                if (paragraphChild is DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties)
                                {
                                    textBuilder.Append("\n");
                                }

                                if (paragraphChild is DocumentFormat.OpenXml.Drawing.Run)
                                {
                                    DocumentFormat.OpenXml.Drawing.Run run = (DocumentFormat.OpenXml.Drawing.Run)paragraphChild;
                                    if (run?.InnerText == null) continue;
                                    DocumentFormat.OpenXml.Drawing.RunProperties props = run.RunProperties;
                                    //apply bold & italic styles
                                    text += inlineStyle(run, props);

                                    //IS A HEADER?
                                    if (props.FontSize != null && props.FontSize > 2500) text = processHeader(text, props.FontSize);


                                    //IS AUTONUMBERED LIST
                                    if (isAutoNum((DocumentFormat.OpenXml.Drawing.Paragraph)item))
                                    {

                                        //if last one was  auto num AND this is auto num
                                        textBuilder.Append(orderedLits + ". " + text + "\n");
                                        text = "";
                                        orderedLits++;
                                        continue;
                                    }
                                    orderedLits = 1;

                                    //IS A BULLET LIST?
                                    if (isBullet((DocumentFormat.OpenXml.Drawing.Paragraph)item)) text = "* " + text+"\n";

                                    
                                    
                                        //textBuilder.Append(text + "\n");
                                        textBuilder.Append(text + "");

                                    
                                    text = "";
                                }


                            }
                            if (paragraphProperties != null)
                            {
                                if (paragraphProperties.Level > 0)
                                {
                                    text = textBuilder.ToString();
                                    textBuilder.Clear();
                                    textBuilder.Append("> " + text + "");
                                }
                            }
                            

                            //why is this variable declaration here?
                            //Answer: I don't know but if you erase it probably colapse 
                            var fontSize = runProp?.FontSize ?? 0;

                        }
                        

                    }
                }
               
            }

            if (isFullCodeBlock)
            {
                textBuilder.Append("``` \n");
            }
            textBuilder.Append("\n");


        }

        private static bool isBullet(DocumentFormat.OpenXml.Drawing.Paragraph paragraph)
        {
            //THE EASIEST WAY TO CHECK IF A LINE IS A BULLET-LIST IS TO CHECK THE EXISTENCE OF THIS 2 PROPERTIES 
            //INSIDE DocumentFormat.OpenXml.Drawing.ParagraphProperties SHOULD BE 
            //DocumentFormat.OpenXml.Drawing.BulletSizePercentage
            //DocumentFormat.OpenXml.Drawing.CharacterBullet
            return paragraph.Descendants<DocumentFormat.OpenXml.Drawing.CharacterBullet>().Count() > 0;
        }

        private static bool isAutoNum(DocumentFormat.OpenXml.Drawing.Paragraph paragraph)
        {
            //THE EASIEST WAY TO CHECK IF A LINE IS A BULLET-LIST IS TO CHECK THE EXISTENCE OF THIS 2 PROPERTIES 
            //INSIDE DocumentFormat.OpenXml.Drawing.ParagraphProperties SHOULD BE 
            //DocumentFormat.OpenXml.Drawing.BulletSizePercentage
            //DocumentFormat.OpenXml.Drawing.CharacterBullet
            return paragraph.Descendants<DocumentFormat.OpenXml.Drawing.AutoNumberedBullet>().Count() > 0;
        }

        public static string inlineStyle(DocumentFormat.OpenXml.Drawing.Run run, DocumentFormat.OpenXml.Drawing.RunProperties props)
        {
            if (run.InnerText==" "|| run.InnerText == "")
            {
                return run.InnerText;
            }
            
            
            if(props.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>().Count() > 0) {


                foreach (var taglinkref in props.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>())
                {
                    //"rId2
                    foreach (var link in links)
                    {
                        if (taglinkref.Id==link.Id)
                        {
                            return "[" + run.InnerText.Trim() + "](" +link.Uri.AbsoluteUri+ ")" ?? "";

                        }
                    }
                }


            }

            if (props.Italic != null && props.Bold != null)
            {
                isInlineStile++;
                string txt = run.InnerText;
                //return "***" + txt+ "***" ?? "";
                return "***" + run.InnerText.Trim() + "*** " ?? "";
            }
            else if (props.Italic != null)
            {
                isInlineStile++;
                string txt = run.InnerText;
                //return "*" +txt+ "*" ?? "";
                return "*" + run.InnerText.Trim() + "* " ?? "";
            }
            else if (props.Bold != null)
            {
                isInlineStile++;
                string txt = run.InnerText;
                //return "**" + txt + "**" ?? "";
                return "**" + run.InnerText.Trim() + "** " ?? "";
            }
            else if (props.Descendants<DocumentFormat.OpenXml.Drawing.Highlight>().Count()>0)
            {
                isInlineStile++;
                return "`" + run.InnerText.Trim() + "` " ?? "";
            }
          
            else
            {
                if (isInlineStile>0)
                {
                    //here ends a line
                    isInlineStile = 0;
                    //stringBuilder.Append("\n");
                }
                return run.InnerText ?? "";
            }

         
        }

        public static string processHeader(string text, int fontSize=0)
        {
            
            switch (fontSize)
            {
                case >= 5500:
                    return "# " + text + "";
                case >= 5000:
                    return "## " + text + "";
                case >= 4500:
                    return "### " + text + "";
                case >= 4000:
                    return "#### " + text + "";
                case >= 3500:
                    return "##### " + text + "";
                case >= 3000:
                    return "###### " + text + "";
                default:
                    return text;
            }
        }

    }
}
