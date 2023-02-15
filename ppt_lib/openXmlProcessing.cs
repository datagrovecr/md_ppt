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
using DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ppt_lib
{
    internal class openXmlProcessing
    {
        public static int isInlineStile = 0;
        public static IEnumerable<HyperlinkRelationship> links = null;
        public static void ProcessParagraph(Shape treeBranch, StringBuilder textBuilder,DocumentFormat.OpenXml.Packaging.SlidePart slides)
        {
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
            if (props==null)
            {
                return run.InnerText;

            }

            if (run.InnerText==" "|| run.InnerText == "")
            {
                return run.InnerText;
            }
            
            
            if(props.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>()?.Count() > 0) {


                foreach (var taglinkref in props.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>())
                {
                    foreach (var link in links)
                    {
                        if (taglinkref.Id==link.Id)
                        {
                            return "[" + ProcessEscapeCharacters(run.InnerText.Trim()) + "](" +link.Uri.AbsoluteUri+ ")" ?? "";

                        }
                    }
                }


            }

            string txt= run.InnerText.Trim();

            /* if (txt == "☑")
             {
                 return "[x] " ?? "";
             }
             if (txt == "☐")
             {
                 return "[ ] " ?? "";
             }
 */
            //un-checkbox
            if (txt == "☐")
            {
                return "[ ] " ?? "";
            }

            if (props.Italic != null && props.Bold != null)
            {
                isInlineStile++;
                return "***" + ProcessEscapeCharacters(txt) + "*** " ?? "";
            }
            else if (props.Italic != null)
            {
                isInlineStile++;
                return "*" + ProcessEscapeCharacters(txt) + "* " ?? "";
            }
            else if (props.Bold != null)
            {
                isInlineStile++;
                //return "**" + txt + "**" ?? "";
                return "**" + ProcessEscapeCharacters(txt) + "** " ?? "";
            }
            else if (props.Descendants<DocumentFormat.OpenXml.Drawing.Highlight>().Count()>0)
            {
                isInlineStile++;
                return "`" + ProcessEscapeCharacters(txt) + "` " ?? "";
            }
            else
            {
               
                return ProcessEscapeCharacters(run.InnerText) ?? "";
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


        public static string ProcessEscapeCharacters(string input)
        {
          
            string result = "";
            for (int i = 0; i < input.Length; i++)
            {
                switch (input[i])
                {
                   
                    case '☑':
                        result += "[x]";
                        break;
                    case '\\':
                        result += '\\';
                        break;
                    case '`':
                        result += "\\`";
                        break;
                    case '-':
                        result += "\\-";
                        break;
                    case '_':
                        result += "\\_";
                        break;
                    case '*':
                        result += "\\*";
                        break;
                    case '{':
                        result += "\\{";
                        break;
                    case '}':
                        result += "\\}";
                        break;
                    case '[':
                        result += "\\[";
                        break;
                    case ']':
                        result += "\\]";
                        break;
                    case '<':
                        result += "\\<";
                        break;
                    case '>':
                        result += "\\>";
                        break;
                    case '(':
                        result += "\\(";
                        break;
                    case ')':
                        result += "\\)";
                        break;
                    case '#':
                        result += "\\#";
                        break;
                    case '+':
                        result += "\\+";
                        break;
                    case '!':
                        result += "\\!";
                        break;
                    default:

                        result += input[i];
                        break;
                }
            }
            return result;
        }

        public static void ProcessPicture(DocumentFormat.OpenXml.Presentation.Picture picture, StringBuilder textBuilder, DocumentFormat.OpenXml.Packaging.SlidePart slides)
        {
            Dictionary<string, string> imagesDict = new();
          
            var mediaId = slides.Parts.ToArray();
            var Urlparts = slides.ImageParts.ToArray();
            for (int i = 0; i < mediaId.Length; i++)
            {
                string rls = mediaId[i].OpenXmlPart.RelationshipType;
                if (rls == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                {
                   
                    imagesDict.Add(mediaId[i].RelationshipId, Urlparts[i].Uri.OriginalString);

                }

            }
            DocumentFormat.OpenXml.Presentation.BlipFill blip = picture.Descendants<DocumentFormat.OpenXml.Presentation.BlipFill>().FirstOrDefault();
          
            textBuilder.Append("![" + picture.InnerText + "](" + imagesDict[blip.Blip.Embed.ToString()] + ")");

        }

        public static string ProcessTable(DocumentFormat.OpenXml.Drawing.Table table)
        {
            string markdown = "";
            // Add table header
            markdown += "|";
            foreach (DocumentFormat.OpenXml.Drawing.TableRow tableRow in table.Descendants<DocumentFormat.OpenXml.Drawing.TableRow>().Take(1))
            {
                foreach (DocumentFormat.OpenXml.Drawing.TableCell tableCell in tableRow.Descendants< DocumentFormat.OpenXml.Drawing.TableCell >())
                {
                    var text = tableCell.InnerText;
                    markdown += text + "|";
                }
                markdown += "\n";
                markdown += "|";
                foreach (DocumentFormat.OpenXml.Drawing.TableCell tableCell in tableRow.Descendants<DocumentFormat.OpenXml.Drawing.TableCell>())
                {
                    markdown += "---|";
                }
                markdown += "\n";
            }
            // Add table rows
            foreach (DocumentFormat.OpenXml.Drawing.TableRow tableRow in table.Descendants<DocumentFormat.OpenXml.Drawing.TableRow>().Skip(1))
            {
                markdown += "|";
                foreach (DocumentFormat.OpenXml.Drawing.TableCell tableCell in tableRow.Descendants<DocumentFormat.OpenXml.Drawing.TableCell>())
                {
                    var text = tableCell.InnerText;
                    markdown += text + "|";
                }
                markdown += "\n";
            }
            return markdown;
        }

    }
}
