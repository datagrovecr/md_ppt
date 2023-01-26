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

        public static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
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
                                new P.ShapeProperties(),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph( new Run( new D.Text() { Text="WORKING"}),
                                        new EndParagraphRunProperties() { Language = "en-US" }
                                                 )
                                    )
                                )
                            )
                        ),
                    new ColorMapOverride(new MasterColorMapping())
                    );
            return slidePart1;
        }
        public static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
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
            #region ol'ways
            /*
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
                        { Name = "Office" });*/
            #endregion


            //theme1.Append(themeElements1);
            //add themeElements with innerXml
            
            theme1.InnerXml = "<a:themeElements>  <a:clrScheme name=\"Office\"> <a:dk1>  <a:sysClr val=\"windowText\" lastClr=\"000000\"/>             </a:dk1>             <a:lt1>  <a:sysClr val=\"window\" lastClr=\"FFFFFF\"/>             </a:lt1>             <a:dk2>                 <a:srgbClr val=\"44546A\"/>             </a:dk2>             <a:lt2>                 <a:srgbClr val=\"E7E6E6\"/>             </a:lt2>             <a:accent1>                 <a:srgbClr val=\"4472C4\"/>             </a:accent1>             <a:accent2>                 <a:srgbClr val=\"ED7D31\"/>             </a:accent2>             <a:accent3>                 <a:srgbClr val=\"A5A5A5\"/>             </a:accent3>             <a:accent4>                 <a:srgbClr val=\"FFC000\"/>             </a:accent4>             <a:accent5>                 <a:srgbClr val=\"5B9BD5\"/>             </a:accent5>             <a:accent6>                 <a:srgbClr val=\"70AD47\"/>             </a:accent6>             <a:hlink>                 <a:srgbClr val=\"0563C1\"/>             </a:hlink>             <a:folHlink>                 <a:srgbClr val=\"954F72\"/>             </a:folHlink>         </a:clrScheme>         <a:fontScheme name=\"Office\">             <a:majorFont>                 <a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/>                 <a:ea typeface=\"\"/>                 <a:cs typeface=\"\"/>                 <a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/>                 <a:font script=\"Hang\" typeface=\"맑은 고딕\"/>                 <a:font script=\"Hans\" typeface=\"等线 Light\"/>                 <a:font script=\"Hant\" typeface=\"新細明體\"/>                 <a:font script=\"Arab\" typeface=\"Times New Roman\"/>                 <a:font script=\"Hebr\" typeface=\"Times New Roman\"/>                 <a:font script=\"Thai\" typeface=\"Angsana New\"/>                 <a:font script=\"Ethi\" typeface=\"Nyala\"/>                 <a:font script=\"Beng\" typeface=\"Vrinda\"/>                 <a:font script=\"Gujr\" typeface=\"Shruti\"/>                 <a:font script=\"Khmr\" typeface=\"MoolBoran\"/>                 <a:font script=\"Knda\" typeface=\"Tunga\"/>                 <a:font script=\"Guru\" typeface=\"Raavi\"/>                 <a:font script=\"Cans\" typeface=\"Euphemia\"/>                 <a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>                 <a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>                 <a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>                 <a:font script=\"Thaa\" typeface=\"MV Boli\"/>                 <a:font script=\"Deva\" typeface=\"Mangal\"/>                 <a:font script=\"Telu\" typeface=\"Gautami\"/>                 <a:font script=\"Taml\" typeface=\"Latha\"/>                 <a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Orya\" typeface=\"Kalinga\"/>                 <a:font script=\"Mlym\" typeface=\"Kartika\"/>                 <a:font script=\"Laoo\" typeface=\"DokChampa\"/>                 <a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>                 <a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>                 <a:font script=\"Viet\" typeface=\"Times New Roman\"/>                 <a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>                 <a:font script=\"Geor\" typeface=\"Sylfaen\"/>                 <a:font script=\"Armn\" typeface=\"Arial\"/>                 <a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/>                 <a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/>                 <a:font script=\"Java\" typeface=\"Javanese Text\"/>                 <a:font script=\"Lisu\" typeface=\"Segoe UI\"/>                 <a:font script=\"Mymr\" typeface=\"Myanmar Text\"/>                 <a:font script=\"Nkoo\" typeface=\"Ebrima\"/>                 <a:font script=\"Olck\" typeface=\"Nirmala UI\"/>                 <a:font script=\"Osma\" typeface=\"Ebrima\"/>                 <a:font script=\"Phag\" typeface=\"Phagspa\"/>                 <a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Sora\" typeface=\"Nirmala UI\"/>                 <a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/>                 <a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/>                 <a:font script=\"Tfng\" typeface=\"Ebrima\"/>             </a:majorFont>             <a:minorFont>                 <a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/>                 <a:ea typeface=\"\"/>                 <a:cs typeface=\"\"/>                 <a:font script=\"Jpan\" typeface=\"游ゴシック\"/>                 <a:font script=\"Hang\" typeface=\"맑은 고딕\"/>                 <a:font script=\"Hans\" typeface=\"等线\"/>                 <a:font script=\"Hant\" typeface=\"新細明體\"/>                 <a:font script=\"Arab\" typeface=\"Arial\"/>                 <a:font script=\"Hebr\" typeface=\"Arial\"/>                 <a:font script=\"Thai\" typeface=\"Cordia New\"/>                 <a:font script=\"Ethi\" typeface=\"Nyala\"/>                 <a:font script=\"Beng\" typeface=\"Vrinda\"/>                 <a:font script=\"Gujr\" typeface=\"Shruti\"/>                 <a:font script=\"Khmr\" typeface=\"DaunPenh\"/>                 <a:font script=\"Knda\" typeface=\"Tunga\"/>                 <a:font script=\"Guru\" typeface=\"Raavi\"/>                 <a:font script=\"Cans\" typeface=\"Euphemia\"/>                 <a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>                 <a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>                 <a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>                 <a:font script=\"Thaa\" typeface=\"MV Boli\"/>                 <a:font script=\"Deva\" typeface=\"Mangal\"/>                 <a:font script=\"Telu\" typeface=\"Gautami\"/>                 <a:font script=\"Taml\" typeface=\"Latha\"/>                 <a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Orya\" typeface=\"Kalinga\"/>                 <a:font script=\"Mlym\" typeface=\"Kartika\"/>                 <a:font script=\"Laoo\" typeface=\"DokChampa\"/>                 <a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>                 <a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>                 <a:font script=\"Viet\" typeface=\"Arial\"/>                 <a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>                 <a:font script=\"Geor\" typeface=\"Sylfaen\"/>                 <a:font script=\"Armn\" typeface=\"Arial\"/>                 <a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/>                 <a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/>                 <a:font script=\"Java\" typeface=\"Javanese Text\"/>                 <a:font script=\"Lisu\" typeface=\"Segoe UI\"/>                 <a:font script=\"Mymr\" typeface=\"Myanmar Text\"/>                 <a:font script=\"Nkoo\" typeface=\"Ebrima\"/>                 <a:font script=\"Olck\" typeface=\"Nirmala UI\"/>                 <a:font script=\"Osma\" typeface=\"Ebrima\"/>                 <a:font script=\"Phag\" typeface=\"Phagspa\"/>                 <a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/>                 <a:font script=\"Sora\" typeface=\"Nirmala UI\"/>                 <a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/>                 <a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/>                 <a:font script=\"Tfng\" typeface=\"Ebrima\"/>             </a:minorFont>         </a:fontScheme>         <a:fmtScheme name=\"Office\">             <a:fillStyleLst>                 <a:solidFill>                     <a:schemeClr val=\"phClr\"/>                 </a:solidFill>                 <a:gradFill rotWithShape=\"1\">                     <a:gsLst>                         <a:gs pos=\"0\">                             <a:schemeClr val=\"phClr\">                                 <a:lumMod val=\"110000\"/>                                 <a:satMod val=\"105000\"/>                                 <a:tint val=\"67000\"/>                             </a:schemeClr>                         </a:gs>                         <a:gs pos=\"50000\">                             <a:schemeClr val=\"phClr\">                                 <a:lumMod val=\"105000\"/>                                 <a:satMod val=\"103000\"/>                                 <a:tint val=\"73000\"/>                             </a:schemeClr>                         </a:gs>                         <a:gs pos=\"100000\">                             <a:schemeClr val=\"phClr\">                                 <a:lumMod val=\"105000\"/>                                 <a:satMod val=\"109000\"/>                                 <a:tint val=\"81000\"/>                             </a:schemeClr>                         </a:gs>                     </a:gsLst>                     <a:lin ang=\"5400000\" scaled=\"0\"/>                 </a:gradFill>                 <a:gradFill rotWithShape=\"1\">                     <a:gsLst>                         <a:gs pos=\"0\">                             <a:schemeClr val=\"phClr\">                                 <a:satMod val=\"103000\"/>                                 <a:lumMod val=\"102000\"/>                                 <a:tint val=\"94000\"/>                             </a:schemeClr>                         </a:gs>                         <a:gs pos=\"50000\">                             <a:schemeClr val=\"phClr\">                                 <a:satMod val=\"110000\"/>                                 <a:lumMod val=\"100000\"/>                                 <a:shade val=\"100000\"/>                             </a:schemeClr>                         </a:gs>                         <a:gs pos=\"100000\">                             <a:schemeClr val=\"phClr\">                                 <a:lumMod val=\"99000\"/>                                 <a:satMod val=\"120000\"/>                                 <a:shade val=\"78000\"/>                             </a:schemeClr>                         </a:gs>                     </a:gsLst>                     <a:lin ang=\"5400000\" scaled=\"0\"/>                 </a:gradFill>             </a:fillStyleLst>             <a:lnStyleLst>                 <a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">                     <a:solidFill>                         <a:schemeClr val=\"phClr\"/>                     </a:solidFill>                     <a:prstDash val=\"solid\"/>                     <a:miter lim=\"800000\"/>                 </a:ln>                 <a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">                     <a:solidFill>                         <a:schemeClr val=\"phClr\"/>                     </a:solidFill>                     <a:prstDash val=\"solid\"/>                     <a:miter lim=\"800000\"/>                 </a:ln>                 <a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">                     <a:solidFill>                         <a:schemeClr val=\"phClr\"/>                     </a:solidFill>                     <a:prstDash val=\"solid\"/>                     <a:miter lim=\"800000\"/>                 </a:ln>             </a:lnStyleLst>             <a:effectStyleLst>                 <a:effectStyle>                     <a:effectLst/>                 </a:effectStyle>                 <a:effectStyle>                     <a:effectLst/>                 </a:effectStyle>                 <a:effectStyle>                     <a:effectLst>                         <a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\">                             <a:srgbClr val=\"000000\">                                 <a:alpha val=\"63000\"/>                             </a:srgbClr>                         </a:outerShdw>                     </a:effectLst>                 </a:effectStyle>             </a:effectStyleLst>             <a:bgFillStyleLst>                 <a:solidFill>                     <a:schemeClr val=\"phClr\"/>                 </a:solidFill>                 <a:solidFill>                     <a:schemeClr val=\"phClr\">                         <a:tint val=\"95000\"/>                         <a:satMod val=\"170000\"/>                     </a:schemeClr>                 </a:solidFill>                 <a:gradFill rotWithShape=\"1\">                     <a:gsLst>                         <a:gs pos=\"0\">                             <a:schemeClr val=\"phClr\">                                 <a:tint val=\"93000\"/>                                 <a:satMod val=\"150000\"/>                                 <a:shade val=\"98000\"/>                                 <a:lumMod val=\"102000\"/>                             </a:schemeClr>                         </a:gs>                         <a:gs pos=\"50000\">                             <a:schemeClr val=\"phClr\">                                 <a:tint val=\"98000\"/>                                 <a:satMod val=\"130000\"/>                                 <a:shade val=\"90000\"/>                                 <a:lumMod val=\"103000\"/>                             </a:schemeClr>                         </a:gs>                         <a:gs pos=\"100000\">                             <a:schemeClr val=\"phClr\">                                 <a:shade val=\"63000\"/>                                 <a:satMod val=\"120000\"/>                             </a:schemeClr>                         </a:gs>                     </a:gsLst>                     <a:lin ang=\"5400000\" scaled=\"0\"/>                 </a:gradFill>             </a:bgFillStyleLst>         </a:fmtScheme> </a:themeElements>";
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());
            theme1.Append(new D.ExtensionList() { InnerXml = "<a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\">             <thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme 2013 - 2022\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/>         </a:ext>" });
            themePart1.Theme = theme1;
            return themePart1;

        }
    
    }
}
