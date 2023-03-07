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

namespace ppt_lib
{
    internal class CreatePresentationDocument
    {

        public static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;


            slidePart1 = CreateSlidePart(presentationPart);
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
        }
        private static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            // offset max Y 6400000 
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = new Slide(
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
                                         new Offset() { X = 0, Y = 3200000 },
                                         new Extents() { Cx = 9144000, Cy = 457200 }
                                         //new Extents() { Cx = 9144000, Cy = 457200 }
                                         )
                                },
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(
                                        new ParagraphProperties() { Alignment = new EnumValue<TextAlignmentTypeValues> { Value = TextAlignmentTypeValues.Center } },
                                    new Run(
                                         new D.RunProperties() { Language = "en-US", Dirty = false, SpellingError = false, FontSize = 2200 },
                                      new D.Text() { Text = "Presentation" }),
                                        new EndParagraphRunProperties() { Language = "en-US" })
                                    )))),
                    new ColorMapOverride(new MasterColorMapping()));
            return slidePart1;
        }

        private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
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
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
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
               new P.ShapeProperties()
               {

                   Transform2D = new Transform2D(
                                          new Offset() { X = 0, Y = 3200000 },
                                          new Extents() { Cx = 9144000, Cy = 457200 }
                                          )
               },
               new P.TextBody(
                 new BodyProperties(),
                 new ListStyle(),
                 new Paragraph())))),
             new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
             new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
             //new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle())
             new TextStyles() { InnerXml= "<p:titleStyle xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:lvl1pPr algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPct val=\"0\" /></a:spcBef><a:buNone /><a:defRPr sz=\"4400\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mj-lt\" /><a:ea typeface=\"+mj-ea\" /><a:cs typeface=\"+mj-cs\" /></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:lvl1pPr marL=\"228600\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"1000\" /></a:spcBef><a:defRPr sz=\"2800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl1pPr><a:lvl2pPr marL=\"685800\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"2400\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl2pPr><a:lvl3pPr marL=\"1143000\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"2000\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl3pPr><a:lvl4pPr marL=\"1600200\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl4pPr><a:lvl5pPr marL=\"2057400\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl5pPr><a:lvl6pPr marL=\"2514600\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl6pPr><a:lvl7pPr marL=\"2971800\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl7pPr><a:lvl8pPr marL=\"3429000\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl8pPr><a:lvl9pPr marL=\"3886200\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:lnSpc><a:spcPct val=\"90000\" /></a:lnSpc><a:spcBef><a:spcPts val=\"500\" /></a:spcBef><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl9pPr></p:bodyStyle><p:otherStyle xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:defPPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr lang=\"en-US\" /></a:defPPr><a:lvl1pPr marL=\"0\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl1pPr><a:lvl2pPr marL=\"457200\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl2pPr><a:lvl3pPr marL=\"914400\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl3pPr><a:lvl4pPr marL=\"1371600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl4pPr><a:lvl5pPr marL=\"1828800\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl5pPr><a:lvl6pPr marL=\"2286000\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl6pPr><a:lvl7pPr marL=\"2743200\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl7pPr><a:lvl8pPr marL=\"3200400\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl8pPr><a:lvl9pPr marL=\"3657600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\" /></a:solidFill><a:latin typeface=\"+mn-lt\" /><a:ea typeface=\"+mn-ea\" /><a:cs typeface=\"+mn-cs\" /></a:defRPr></a:lvl9pPr></p:otherStyle>" }
             );
             slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
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
