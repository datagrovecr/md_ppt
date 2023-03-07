using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using HtmlAgilityPack;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using HtmlAgilityPack;
using System.Xml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Runtime.Serialization;

namespace ppt_lib
{
    internal class processSlidesAdd
    {
        public static uint drawingObjectId2 = 2;
        public static int y = 1000000;
        public static int HyperlinkCount = 1;
        public static PresentationDocument presentationDocument=null;
        public static List<HyperLInkElement> hyperlinkList = new();
        public static List<List<Shape>> listOfShapes = new();
        
        // Insert a slide into the specified presentation.
        public static void ProcessHtmlToPresentation(PresentationDocument presentationDocumentF, HtmlDocument htmlDoc)
        {
            presentationDocument = presentationDocumentF;
            List<Shape> shapes = new();
            processFragment(htmlDoc.DocumentNode,shapes);
            int i = 1;

            //HERE LOOP YOUR LIST OF LIST'S
            foreach (var shapeElements in listOfShapes)
            {
                InsertNewSlideBaby(presentationDocument, i, shapeElements);
                i++;
            }//HERE ENDS LOOP
        }
        
        public static void processFragment(HtmlNode htmlNodeDoc, List<Shape> shapes)
        {
          
            

            //InsertNewSlide(presentationDocument, 1, htmlDoc);
            //CREATE ALL SHAPES AND THEN DECIDE TO ADD IT TO INSERT SLIDE
            //HANDLE THE POSITION HERE
            //SET THEIR POSITION BY DEFAULT AND LATER ACCESS  ShapeProperties
            //OR USE A FOR LOOP TO ADJUST     Y=VALUE
            // find all the url insert them on the presentation document
            //get the in order later
            //htmlDoc.DocumentNode.ChildNodes.Descendants<>


            if (htmlNodeDoc.HasChildNodes==false)
            {
                ProcessLine(htmlNodeDoc, shapes);
            }
            else
            {
                //go for A
                ProcessLine(htmlNodeDoc, shapes);
                foreach (var htmlNode in htmlNodeDoc.ChildNodes)
                {

                    ProcessLine(htmlNode,shapes);


                }//HERE ENDS LOOP
            }
           

            if (shapes.Count > 0)  listOfShapes.Add(shapes);


            
        }

        public static OpenXmlElement ProcessLine(HtmlNode htmlNode, List<Shape> shapes)
        {

            if (y >= 5000000)
            {
                listOfShapes.Add(shapes);
                shapes = new();
            }

            // var result = processFragment(htmlNode);
            if (htmlNode.Name == "#text")
            {

                TextBody S = new TextBody(new Drawing.BodyProperties(),
                        new Drawing.ListStyle()
                        );

                Drawing.Paragraph para = new Drawing.Paragraph(new Drawing.ParagraphProperties() { Alignment = Drawing.TextAlignmentTypeValues.Center });
                int FontSize = 18000;
                //split n 
                //create a Run for each one of the elements
                if (htmlNode.InnerText.Contains("\n"))
                {
                    string[] text = htmlNode.InnerText.Split('\n');
                    int i = 0;
                    foreach (string lines in text)
                    {

                        // \n happens

                        if (i == text.Length - 1)
                        {
                            para.AppendChild(
                            new Drawing.Run(
                            new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = FontSize },
                            new Drawing.Text() { Text = lines })
                            );

                        }
                        else
                        {
                            para.AppendChild(
                            new Drawing.Run(
                            new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = FontSize },
                            new Drawing.Text() { Text = lines })
                            );
                            S.AppendChild(para);

                            para = new Drawing.Paragraph(new Drawing.ParagraphProperties() { Alignment = Drawing.TextAlignmentTypeValues.Center });

                        }
                        i++;
                    }//here ends loop

                }
                else
                {
                    para.AppendChild(
                           new Drawing.Run(
                           new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = FontSize },
                           new Drawing.Text() { Text = htmlNode.InnerText })
                           );
                    S.AppendChild(para);
                }
                return S;

            }
            else if (htmlNode.Name == "h1" || htmlNode.Name == "h2" || htmlNode.Name == "h3" || htmlNode.Name == "h4")
            {
                // Declare and instantiate the title shape of the new slide.
                //shapes.Add(shapeList.TitleShape(y, htmlNode));

                listOfShapes.Add(shapes);
                shapes = new();


                drawingObjectId2++;
                y = 1000000;

            }

            //
            else if (htmlNode.Name == "p")
            {

                //CALL #TEXT
                //set the font size 
                // Declare and instantiate the body shape of the new slide.
                Shape bodyShape = new Shape();

                // Specify the required shape properties for the body shape.
                bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = processSlidesAdd.drawingObjectId2, Name = "Content Placeholder" },
                        new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
                bodyShape.ShapeProperties = new ShapeProperties()
                {
                    Transform2D = new Drawing.Transform2D(
                                             new Drawing.Offset() { X = 0, Y = y },
                                             new Drawing.Extents() { Cx = 9144000, Cy = 457200 }
                                             )
                };

                Drawing.Paragraph Slide = new Drawing.Paragraph();



                // Specify the text of the title shape.
                bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                        new Drawing.ListStyle()

                        );

                foreach (var htmlNodeBaby in htmlNode.ChildNodes)
                {
                    bodyShape.TextBody.AppendChild(ProcessLine(htmlNodeBaby,shapes));
                }

                //listOfShapes.AddRange(result);
                //shapes.Add(shapeList.TextShape(y, htmlNode));
                y += 1000000;
                shapes.Add(bodyShape);
                return null;
            }
            else if (htmlNode.Name == "ul")
            {
                //shapes.Add(shapeList.BulletListShape(y, htmlNode));

                //y += 1000000;
            }
            else if (htmlNode.Name == "ol")
            {
                //shapes.Add(shapeList.OrderedListShape(y, htmlNode));

                //y += 1000000;
            }
            else if (htmlNode.Name == "a")
            {
                Drawing.Paragraph para = new Drawing.Paragraph(new Drawing.ParagraphProperties() { Alignment = Drawing.TextAlignmentTypeValues.Center });
                int FontSize = 18000;
                string href = htmlNode.Attributes["href"].Value;
                string text = htmlNode.InnerText;
                HyperLInkElement hyperlink = processSlidesAdd.UrlProcess(href);
                //here add the link 
                //Name: "title", Value: "The best search engine for privacy"
                string tooltip = htmlNode.Attributes["title"]?.Value != null ? htmlNode.Attributes["title"]?.Value : "";

                /*
                  <a:rPr lang="en-US" b="0" i="0" u="sng" dirty="0">
                            <a:solidFill>
                                <a:srgbClr val="A0AABF"/>
                            </a:solidFill>
                            <a:effectLst/>
                            <a:latin typeface="Georgia" panose="02040502050405020303" pitchFamily="18" charset="0"/>
                            <a:hlinkClick r:id="rId2" tooltip="The best search engine for privacy"/>
                        </a:rPr>
                 */
                para.AppendChild(
                            new Drawing.Run(
                            new Drawing.RunProperties(
                                new Drawing.SolidFill(
                                    new Drawing.RgbColorModelHex() { Val = "A0AABF" }
                                    ),
                                new Drawing.EffectList()
                                , new Drawing.HyperlinkOnClick() { Id = hyperlink.id, Tooltip = tooltip }
                                )
                            { Language = "en-US", Dirty = false, Bold = true, FontSize = FontSize },
                            new Drawing.Text() { Text = htmlNode.InnerText })
                            );
                return para;
            }
            else if (htmlNode.Name == "pre")
            {
                foreach (var htmlNodeSon in htmlNode.ChildNodes)
                {
                    if (htmlNodeSon.Name == "code")
                    {
                        // shapes.Add(shapeList.codeblockShape(y, htmlNode));
                        y += 1000000;

                    }
                }

            }
            else if (htmlNode.Name == "blockquote")
            {
                //shapes.Add(shapeList.BlockQuoteShape(y, htmlNode));
                y += 1000000;
            }
            else
            {
                //everithing else

            }

            return new Drawing.Break();
        }

        public static HyperLInkElement UrlProcess(string linkUrl)
        {
            HyperlinkCount++;
            HyperLInkElement CURRENT = new HyperLInkElement(new Uri(linkUrl, UriKind.Absolute),
            "rId" +new ObjectIDGenerator().GetHashCode());
            
            //save it on the list
            hyperlinkList.Add(CURRENT);
            return CURRENT;
           
        }


        public static void InsertNewSlideBaby(PresentationDocument presentationDocument, int position, List<Shape> shapes3)
        {
            // HtmlNode htmlNode
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            /*if (htmlNode.InnerText == null)
            {
                throw new ArgumentNullException("slideTitle");
            }*/

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            foreach (var shapeNode in shapes3)
            {

                shapeNode.NonVisualShapeProperties.NonVisualDrawingProperties.Id = drawingObjectId;

                // Declare and instantiate the title shape of the new slide.
                slide.CommonSlideData.ShapeTree.AppendChild(shapeNode);

                
                drawingObjectId++;
              


            }



            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            //hyperlinkList
            foreach (HyperLInkElement helement in hyperlinkList)
            {
                slidePart.AddHyperlinkRelationship(helement.url,true, helement.id);
            }
            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

     
        public static void InsertNewHeader(PresentationDocument presentationDocument, int position,HtmlNode htmlNode)
        {

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (htmlNode.InnerText == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties()
            {
                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 0, Y = 2900000 },
                                         new Drawing.Extents() { Cx = 9144000, Cy = 557200 }
                                         )
            };

            //set the font size 
            int fontSize(){
                switch (htmlNode.Name)
                {
                    case "h1":
                        return 6500;
                    case "h2":
                        return 5500;
                    case "h3":
                        return 5000;
                    case "h4":
                        return 4000;
                    default:
                        return 6500;
                }
            };

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties() { Alignment=Drawing.TextAlignmentTypeValues.Center},
                        new Drawing.Run(
                         new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = fontSize()},
                        new Drawing.Text() { Text = htmlNode.InnerText })
                                        )
                    );


            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        // Insert the specified slide into the presentation at the specified position.
        public static void InsertNewSlide2(PresentationDocument presentationDocument, int position, string slideTitle)
        {

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties()
            {
                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 0, Y = 3200000 },
                                         new Drawing.Extents() { Cx = 9144000, Cy = 457200 }
                                         )
            };

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties()
            {
                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 0, Y = 4200000 },
                                         new Drawing.Extents() { Cx = 9144000, Cy = 457200 }
                                         )
            };

            // Specify the text of the body shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph());

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }//HERE ENDS LOOP

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }
    }
}