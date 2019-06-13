using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  public static class Clean
  {
    public static void run(string pathFolder)
    {
      using (PresentationDocument presentationDocument = PresentationDocument.Open(
      pathFolder + 
      "\\pptx_template\\" +
      "NEW_TEMPLATE.pptx", true))
      {
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        int slidesCount = presentationPart.SlideParts.Count();

        Presentation presentation = presentationPart.Presentation;

        presentation.SlideSize.Cy = 5150000;

        SlideIdList slideIdList = presentation.SlideIdList;

        for(int i = 0; i <= slidesCount - 1; i++)
        {
          SlideId slideId = slideIdList.ChildElements[i] as SlideId;

          string slidePartRelationshipId = slideId.RelationshipId;

          SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

          Slide sld = slidePart.Slide;

          if (sld.InnerText.IndexOf("Evaluation only.") != -1)
          {
            int index = sld.InnerText.IndexOf("Evaluation only.");

            int textNumber = sld.Descendants<TextBody>().Count();

            for(int y = 0; y <= textNumber - 1; y++)
            {
              TextBody textBody = sld.Descendants<TextBody>().ElementAt(y);

              if(textBody.InnerText == "Evaluation only.Created with Aspose.Slides for .NET 4.0 Client Profile 19.5.Copyright 2004-2019Aspose Pty Ltd.")
              {
                A.Paragraph p1 = textBody.Elements<A.Paragraph>().ElementAt(0);
                A.Paragraph p2 = textBody.Elements<A.Paragraph>().ElementAt(1);
                A.Paragraph p3 = textBody.Elements<A.Paragraph>().ElementAt(2);

                textBody.RemoveChild<A.Paragraph>(p1);
                textBody.RemoveChild<A.Paragraph>(p2);
                textBody.RemoveChild<A.Paragraph>(p3);

                textBody.AppendChild<A.Paragraph>(new A.Paragraph());
              }
            }
          }
        }
        presentationDocument.PresentationPart.Presentation.Save();
      }
    }
  }
}
