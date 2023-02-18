using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using HtmlAgilityPack;

namespace ppt_lib
{
    internal class CreatePresentationDocumentCopy
    {
        
        public static void CreatePresentationParts(PresentationDocument presentationDoc, HtmlDocument htmlDoc)
        {

            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
      
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            processSlidesAdd.InsertNewSlide2(presentationDoc, 1, "AAAAAAAAAAAAAAAAAAAAAAAAAAA 1");
            processSlidesAdd.InsertNewSlide2(presentationDoc, 2, "AAAAAAAAAAAAAAAAAAAAAAAAAAA 2");

            //Close the presentation handle
            presentationDoc.Close();

            for (int i = 2; i < 4; i++)
            {
               InsertNewSlide(presentationPart, i, "filmina" + i);

            }
        }


        public static void InsertNewSlide(PresentationPart presentationPart, int position, string slideTitle)
        {
           
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            P.NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            P.Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new P.NonVisualShapeProperties
                (new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new P.ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new P.TextBody(new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new Run(new P.Text() { Text = "AAAAAAAAAAAA" })));

            // Declare and instantiate the body shape of the new slide.
            P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new P.ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new P.TextBody(new BodyProperties(),
                    new ListStyle(),
                    new Paragraph());

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





        public static SlidePart CreateSlidePart(PresentationPart presentationPart,HtmlDocument htmlDoc)
        {
            // offset max Y 6400000 
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
            /*slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties()
                                {

                                    Transform2D = new Transform2D(
                                         new Offset() { X = 0, Y = 6400000 },
                                         new Extents() { Cx = 9144000, Cy = 457200 }
                                         //new Extents() { Cx = 9144000, Cy = 457200 }
                                         )
                                },
                                new P.TextBody(  //THIS PART SHOULD BE LOOPED
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new Run(
                                        new D.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false,FontSize=4400 },
                                        new D.Text() { Text = "WORKING" }),
                                        new EndParagraphRunProperties() { Language = "en-US" }
                                                 )
                                    )
                                )
                            )
                        ),
                    new ColorMapOverride(new MasterColorMapping())
                    );
*/
            Paragraph paragraph = new Paragraph();
            foreach (var itemSack in htmlDoc.DocumentNode.ChildNodes)
            {


                paragraph.AppendChild(
                      new Run( //multiple runs
                                      new D.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 2200 },
                                      new D.Text() { Text = itemSack.InnerText })


                    );

            }

            paragraph.AppendChild(new EndParagraphRunProperties() { Language = "en-US" });




            P.Shape shape = new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties()
                                {
                                    //asdasdasd
                                    Transform2D = new Transform2D(
                                         new Offset() { X = 0, Y = 0 },
                                         new Extents() { Cx = 9144000, Cy = 457200 }
                                         //new Extents() { Cx = 9144000, Cy = 457200 }
                                         )
                                },
                                new P.TextBody( 
                                    new BodyProperties(),
                                    new ListStyle(),
                                    paragraph
                                    )
                                );



            slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            shape
                            )
                        ),
                    new ColorMapOverride(new MasterColorMapping())
                    );

            return slidePart;
        }
        public static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            //edit this
            //create 1 slidelayout
            //add it to 
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
          //  slideLayout.InnerXml = " <p:txStyles>\r\n    <p:titleStyle>\r\n      <a:defPPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n        <a:defRPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" sz=\"3600\" b=\"true\"/>\r\n      </a:defPPr>\r\n    </p:titleStyle>\r\n    <p:bodyStyle>\r\n      <a:defPPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n        <a:defRPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" sz=\"2400\"/>\r\n      </a:defPPr>\r\n    </p:bodyStyle>\r\n  </p:txStyles>\r\n  <p:cs>\r\n    <p:spTree>\r\n      <p:nvGrpSpPr>\r\n        <p:cNvPr id=\"1\" name=\"Title Placeholder\"/>\r\n        <p:cNvGrpSpPr/>\r\n        <p:nvPr/>\r\n      </p:nvGrpSpPr>\r\n      <p:grpSpPr>\r\n        <a:xfrm  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n          <a:off  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" x=\"0\" y=\"0\"/>\r\n          <a:ext  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" cx=\"9144000\" cy=\"6858000\"/>\r\n        </a:xfrm>\r\n      </p:grpSpPr>\r\n      <p:sp>\r\n        <p:nvSpPr>\r\n          <p:cNvPr id=\"2\" name=\"Subtitle Placeholder\"/>\r\n          <p:cNvSpPr>\r\n            <a:spLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></a:spLocks> noGrp=\"1\"/>\r\n          </p:cNvSpPr>\r\n          <p:nvPr>\r\n            <p:ph type=\"subTitle\"/>\r\n          </p:nvPr>\r\n        </p:nvSpPr>\r\n        <p:spPr>\r\n          <a:xfrm  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n            <a:off xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" x=\"0\" y=\"1875446\"/>\r\n            <a:ext xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" cx=\"9144000\" cy=\"1598269\"/>\r\n          </a:xfrm>\r\n          <a:prstGeom  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" prst=\"rect\">\r\n            <a:avLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></a:avLst>/>\r\n          </a:prstGeom>\r\n        </p:spPr>\r\n        <p:txBody>\r\n          <a:bodyPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" anchor=\"ctr\" bIns=\"457200\" tIns=\"457200\" lIns=\"914400\" rIns=\"914400\" rtlCol=\"0\">\r\n            <a:schemeClr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" val=\"tx1\">\r\n              <a:lumMod  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" val=\"65000\"/>\r\n              <a:lumOff  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" val=\"35000\"/>\r\n            </a:schemeClr>\r\n          </a:bodyPr>\r\n          <a:lstStyle  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n            <a:lvl1pPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" marL=\"0\" indent=\"0\">\r\n              <a:defRPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" sz=\"4400\"/>\r\n            </a:lvl1pPr>\r\n            <a:lvl2pPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" marL=\"457200\" indent=\"0\">\r\n              <a:defRPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" sz=\"3200\"/>\r\n            </a:lvl2pPr>\r\n            <a:lvl3pPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" marL=\"914400\" indent=\"0\">\r\n              <a:defRPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" sz=\"2400\"/>\r\n            </a:lvl3pPr>\r\n          </a:lstStyle>\r\n          <a:p  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n            <a:pPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" marL=\"0\" lvl=\"0\">\r\n              <a:defRPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" sz=\"4400\" b=\"true\"/>\r\n            </a:pPr>\r\n            <a:r  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n              <a:rPr  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" lang=\"en-US\" sz=\"4400\"></a:rPr>  \r\n            </a:r>\r\n           </a:p>  \r\n        </p:txBody>\r\n      </p:sp>\r\n    </p:spTree>\r\n  </p:cs>";
            slideLayoutPart1.SlideLayout= slideLayout;
            return slideLayoutPart1;
        }

        public static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster slideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        public static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }))
              { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
              { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }
    
    }
}
