using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using HtmlAgilityPack;

namespace ppt_lib
{
    internal class ConvertPresentationDocument
    {
        public static void ProcessHtml(PresentationPart presentationPart, HtmlDocument htmlDoc)
        {
            //  HtmlDocument htmlDocument
            


            foreach (var itemSack in htmlDoc.DocumentNode.ChildNodes)
            {





            }

            // PresentationPart presentationPart = presentationDocument.PresentationPart;
            // Get the presentation object
            Presentation presentation = presentationPart.Presentation;

            // Create a new slide
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));

            // Get the slide ID of the last slide in the presentation
            int slideId = 3;
            if (presentation.SlideIdList!=null)
            {
                slideId = presentation.SlideIdList.ChildElements.Count() + 256 ;

            }
            else
            {
                presentation.SlideIdList = new SlideIdList();
            }

            // Create a new slide ID for the new slide
            SlideId slideIdObj = new SlideId();
            slideIdObj.Id =(UInt32) slideId;
            slideIdObj.RelationshipId = presentationPart.GetIdOfPart(presentationPart.AddNewPart<SlidePart>());
            presentation.SlideIdList.Append(slideIdObj);

            // Save the new slide to the slide part
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>(slideIdObj.RelationshipId);
            slide.Save(slidePart);
        }

        public static void addSlides(PresentationPart presentationPart)
        {
            
        }
    }
}