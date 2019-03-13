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
  public static class PPTXReader
  {
    public static string getNumberImage(ImagePart img)
    {
      string imageFileNumber = img.Uri.OriginalString.Substring(16, 2);
      if (imageFileNumber.IndexOf('.') != -1)
      {
        imageFileNumber = imageFileNumber.Substring(0, 1);
      }

      return imageFileNumber;
    }
    
    public static void switchSlide(string relId, string text, SlidePart slide)
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
  }
}
