﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using HtmlAgilityPack;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using HtmlAgilityPack;
namespace ppt_lib
{
    internal class processSlidesAdd
    {
        public static uint drawingObjectId2 = 2;

        // Insert a slide into the specified presentation.
        public static void ProcessHtmlToPresentation(PresentationDocument presentationDocument, HtmlDocument htmlDoc)
        {
            
            List<List<Shape>> listOfShapes=new();
            List<Shape> shapes = new();
            int i = 1;
            int y = 1000000;
            //InsertNewSlide(presentationDocument, 1, htmlDoc);
            //CREATE ALL SHAPES AND THEN DECIDE TO ADD IT TO INSERT SLIDE
            //HANDLE THE POSITION HERE
            //SET THEIR POSITION BY DEFAULT AND LATER ACCESS  ShapeProperties
            //OR USE A FOR LOOP TO ADJUST     Y=VALUE
            foreach (var htmlNode in htmlDoc.DocumentNode.ChildNodes)
            {

                if (y >= 3000000)
                {
                    listOfShapes.Add(shapes);
                    shapes = new();
                }
                if (htmlNode.Name == "h1" || htmlNode.Name == "h2" || htmlNode.Name == "h3" || htmlNode.Name == "h3")
                {
                    // Declare and instantiate the title shape of the new slide.
                    shapes.Add(shapeList.TitleShape(y, htmlNode));
                    
                    listOfShapes.Add(shapes);
                    shapes = new();


                    drawingObjectId2++;
                    y = 1000000;

                }
                else if (htmlNode.Name == "p")
                {
                    shapes.Add(shapeList.TextShape(y, htmlNode));
                    y += 1000000;
                }
                else
                {
                    //everithing else

                }
                
               
            }
            if (shapes.Count > 0)  listOfShapes.Add(shapes);

            //HERE LOOP YOUR LIST OF LIST'S
            foreach (var shapes1 in listOfShapes)
            {
                InsertNewSlideBaby(presentationDocument, i, shapes1);
                i++;
            }

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
                //Shape shapeIncoming = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
                slide.CommonSlideData.ShapeTree.AppendChild(shapeNode);

                //deal with this later
                drawingObjectId++;
              


            }



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

        public static void InsertNewSlide(PresentationDocument presentationDocument, int position, HtmlDocument htmlDoc)
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

            int i = 0;
            int y = 1900000;
            foreach (var htmlNode in htmlDoc.DocumentNode.ChildNodes)
            {
                if (i==2)
                {
                    //create slide

                    break;
                }
                
                if (htmlNode.Name == "h1" || htmlNode.Name == "h2" || htmlNode.Name == "h3" || htmlNode.Name == "h3")
                {
                    // Declare and instantiate the title shape of the new slide.
                    Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
                    //deal with this later
                    drawingObjectId++;

                    // Specify the required shape properties for the title shape. 
                    titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                        (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                        new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
                    titleShape.ShapeProperties = new ShapeProperties()
                    {
                        Transform2D = new Drawing.Transform2D(
                                                // new Drawing.Offset() { X = 0, Y = 2900000 },
                                                 new Drawing.Offset() { X = 0, Y = y },
                                                 new Drawing.Extents() { Cx = 9144000, Cy = 557200 }
                                                 )
                    };
                    //asdasdasdasdasdasd

                    ///add header alone
                    ///
                    int fontSize()
                    {
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
                    //set the font size 


                    // Specify the text of the title shape.
                    titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(
                                new Drawing.ParagraphProperties() { Alignment = Drawing.TextAlignmentTypeValues.Center },
                                new Drawing.Run(
                                 new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = fontSize() },
                                new Drawing.Text() { Text = htmlNode.InnerText })
                                                )
                            );
                    i++;
                    y += 1000000;
                }
                else if (htmlNode.Name == "photo")
                {

                }
                else
                {
                    //everithing else

                }

                
            }
         


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
    }
}