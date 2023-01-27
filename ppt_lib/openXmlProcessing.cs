using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;


namespace ppt_lib
{
    internal class openXmlProcessing
    {

        public static void ProcessParagraph(Shape treeBranch, StringBuilder textBuilder)
        {
            foreach (var element in treeBranch)
            {
                if (element is TextBody)
                {
                    foreach (var item in element)
                    {
                        // DocumentFormat.OpenXml.Drawing.ListStyle
                    }
                }
            }
            textBuilder.Append(treeBranch.InnerText + "\n");

        }
    }
}
