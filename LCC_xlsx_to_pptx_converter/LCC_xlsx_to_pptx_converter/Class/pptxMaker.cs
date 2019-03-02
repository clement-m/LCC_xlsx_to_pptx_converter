using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class
{
  class pptxMaker
  {
    const string DIR_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\";

    public pptxData pptxData;
    public PresentationDocument presentationDoc;

    public pptxMaker(pptxData pptxData)
    {
      this.pptxData = pptxData;
    }

    public void switchSlide(string relId, string text, SlidePart slide)
    {
      switch (relId)
      {
        case "rId2":
          //OpenXmlUtils.InsertNewSlide(DIR_PATH, 1, "TAMERLAPUT");
          //title
          break;
        case "rId4":
          // Cimentery..
          Console.WriteLine("slide id: '{0}' with textcontent: {1}", relId, text);
          break;
        case "rId25":
          // R+2
          Console.WriteLine("slide id: '{0}' with textcontent: {1}", relId, text);
          break;
      }
    }

    public void run()
    {
      string path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), @"..\\..\\pptx_template\\template.pptx");
      Stream stream = File.Open(path, FileMode.Open);
      using (PresentationDocument presentationDocument =
          PresentationDocument.Open(stream, false))
      {
        Console.WriteLine("Ouverture du template");

        // Get the relationship ID of the first slide.
        PresentationPart part = presentationDocument.PresentationPart;
        OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

        for (int i = 0; i <= slideIds.Count(); i++)
        {
          string relId = (slideIds[i] as SlideId).RelationshipId;
          Console.WriteLine("slide id: {0}", relId);

          // Get the slide part from the relationship ID.
          SlidePart slide = (SlidePart)part.GetPartById(relId);

          // Build a StringBuilder object.
          StringBuilder paragraphText = new StringBuilder();

          // Get the inner text of the slide.
          IEnumerable<D.Text> texts = slide.Slide.Descendants<D.Text>();
          foreach (D.Text text in texts)
          {
            paragraphText.Append(text.Text);
            Console.WriteLine("text content {0}", text.Text);
            this.switchSlide(relId, text.Text, slide);

          }
          string sldText = paragraphText.ToString();


        }


      }
    }
  }
}
