using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
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

        public static Shape BulletListShape(int y, HtmlNode htmlNode)
        {


            //set the font size 
            // Declare and instantiate the body shape of the new slide.
            Shape listShape = new Shape();

            // Specify the required shape properties for the body shape.
            listShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = processSlidesAdd.drawingObjectId2, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            listShape.ShapeProperties = new ShapeProperties()
            {
                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 0, Y = y },
                                         new Drawing.Extents() { Cx = 9144000, Cy = 457200 }
                                         )
            };

            /*
             <a:pPr marL="285750" indent="-285750" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0" />
            <a:buChar char="•" />
            </a:pPr><a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:rPr lang="en-US" dirty="0" />
            <a:t>Element1</a:t>
            </a:r>
             */
            string textList = "";
            int lastNode =htmlNode.ChildNodes.Count-2;
            int index = 0;
            foreach (var list in htmlNode.ChildNodes)
            {
                if (index== lastNode && list.Name== "li")
                {
                    textList += list.InnerText;
                    processSlidesAdd.y += 300000;
                }
                else if (list.Name == "li")
                {
                    textList += list.InnerText + "\n";
                    processSlidesAdd.y +=  300000;
                }
                index++;
            }

            // Specify the text of the title shape.
            listShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle() ,
                    
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties(
                            new Drawing.BulletFont() { Typeface= "Arial", Panose= "020B0604020202020204", PitchFamily = 34 , CharacterSet = 0 },
                            new Drawing.CharacterBullet() { Char= "•" }
                        ) { Alignment = Drawing.TextAlignmentTypeValues.Center, LeftMargin= 285750,Indent= -285750 },
                        new Drawing.Run(
                         new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 1800 },
                        new Drawing.Text() { Text = textList })
                                        )
                    );
            return listShape;
        }

        public static Shape OrderedListShape(int y, HtmlNode htmlNode)
        {


            //set the font size 
            // Declare and instantiate the body shape of the new slide.
            Shape listShape = new Shape();

            // Specify the required shape properties for the body shape.
            listShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = processSlidesAdd.drawingObjectId2, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            listShape.ShapeProperties = new ShapeProperties()
            {
                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 0, Y = y },
                                         new Drawing.Extents() { Cx = 9144000, Cy = 457200 }
                                         )
            };

            /*
            <a:pPr marL="342900" indent="-342900" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:buFont typeface="+mj-lt" />
            <a:buAutoNum type="arabicPeriod" />
            </a:pPr>
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:rPr lang="en-US" dirty="0" />
            <a:t>One
            </a:t>
            </a:r>
             */
            string textList = "";
            int lastNode = htmlNode.ChildNodes.Count - 2;
            int index = 0;
            foreach (var list in htmlNode.ChildNodes)
            {
                if (index == lastNode && list.Name == "li")
                {
                    textList += list.InnerText;
                    processSlidesAdd.y += 300000;
                }
                else if (list.Name == "li")
                {
                    textList +=  list.InnerText + "\n";
                    processSlidesAdd.y += 300000;
                }
                index++;
            }

            // Specify the text of the title shape.
            listShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),

                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties(
                            new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 },
                            new Drawing.AutoNumberedBullet() {Type=Drawing.TextAutoNumberSchemeValues.ArabicPeriod  }
                        )
                        { Alignment = Drawing.TextAlignmentTypeValues.Center, LeftMargin = 285750, Indent = -285750 },
                        new Drawing.Run(
                         new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 1800 },
                        new Drawing.Text() { Text = textList })
                                        )
                    );
            return listShape;
        }


        public static Shape codeblockShape(int y, HtmlNode htmlNode)
        {

            //set the font size 
            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = new Shape();

            ///
            /// 
            /// <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            /// <a:off x="500000" y="1064008" />
            /// <a:ext cx="8144000" cy="923330" />
            /// </a:xfrm>
            /// <a:prstGeom prst="rect" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            /// <a:avLst />
            /// </a:prstGeom>
            /// <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            /// <a:schemeClr val="bg1">
            /// <a:lumMod val="65000" />
            /// </a:schemeClr>
            /// </a:solidFill>
            /// 

            
            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = processSlidesAdd.drawingObjectId2, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties(
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset=Drawing.ShapeTypeValues.Rectangle}
                ,
                new Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor(new Drawing.LuminanceModulation(){Val= 65000 }) { Val=Drawing.SchemeColorValues.Background1} )
               

                )
            {

                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 500000, Y = y },
                                         new Drawing.Extents() { Cx = 8144000, Cy = 923330 }
                                         )

            };


            // Specify the text of the title shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties() { Wrap=Drawing.TextWrappingValues.Square},
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties() { Alignment = Drawing.TextAlignmentTypeValues.Left },
                        new Drawing.Run(
                         new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 1800 },
                        new Drawing.Text() { Text = htmlNode.InnerText })
                                        )
                    );

            return bodyShape;

        }

        public static Shape BlockQuoteShape(int y, HtmlNode htmlNode)
        {

            //set the font size 
            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = new Shape();

            ///
            /// 
            /// <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            /// <a:off x="500000" y="1064008" />
            /// <a:ext cx="8144000" cy="923330" />
            /// </a:xfrm>
            /// <a:prstGeom prst="rect" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            /// <a:avLst />
            /// </a:prstGeom>
            /// <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            /// <a:schemeClr val="bg1">
            /// <a:lumMod val="65000" />
            /// </a:schemeClr>
            /// </a:solidFill>
            /// 


            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = processSlidesAdd.drawingObjectId2, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 })
                    )
            ;

            bodyShape.ShapeProperties = new ShapeProperties(
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
               /* ,
                new Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor(new Drawing.LuminanceModulation() { Val = 65000 }) { Val = Drawing.SchemeColorValues.Background1 })*/
                )
            {

                Transform2D = new Drawing.Transform2D(
                                         new Drawing.Offset() { X = 500000, Y = y },
                                         new Drawing.Extents() { Cx = 8144000, Cy = 923330 }
                                         )

            };


            // Specify the text of the title shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(new Drawing.ShapeAutoFit()) { Wrap = Drawing.TextWrappingValues.Square, RightToLeftColumns = false },
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties() {Level=1 },
                        new Drawing.Run(
                         new Drawing.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 1800 },
                        new Drawing.Text() { Text = htmlNode.InnerText })
                        ,
                        new Drawing.EndParagraphRunProperties()
                                        )
                    );

            return bodyShape;

        }
    }
}
