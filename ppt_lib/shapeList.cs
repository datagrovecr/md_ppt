using DocumentFormat.OpenXml.Presentation;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Drawing = DocumentFormat.OpenXml.Drawing;


namespace ppt_lib
{
    internal class shapeList
    {

        public static Shape TitleShape(int y,HtmlNode htmlNode)
        {

            //drawingObjectId++;
            Shape titleShape=new Shape();
            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = processSlidesAdd.drawingObjectId2, Name = "Title" },
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
            return titleShape;
        }
        public static Shape TextShape(int y, HtmlNode htmlNode)
        {

           
            //set the font size 
            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape =new Shape();

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


            // Specify the text of the title shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties() { Alignment = Drawing.TextAlignmentTypeValues.Center },
                        new Drawing.Run(
                         new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 1800 },
                        new Drawing.Text() { Text = htmlNode.InnerText })
                                        )
                    );
            return bodyShape;
        }
    }
}
