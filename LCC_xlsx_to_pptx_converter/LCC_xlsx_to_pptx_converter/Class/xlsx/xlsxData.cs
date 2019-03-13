using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace LCC_xlsx_to_pptx_converter.Class
{
  class XLSXData
  {
    public Dictionary<string, byte[]> images;

    public XLSXData()
    {
      this.images = new Dictionary<string, byte[]>();
    }

    public void addImage(WorksheetPart worksheetPart)
    {
      foreach (ImagePart i in worksheetPart.DrawingsPart.ImageParts)
      {
        var rId = worksheetPart.DrawingsPart.GetIdOfPart(i);
        string imageFileNumber = i.Uri.OriginalString.Substring(15, 2);
        if (imageFileNumber.IndexOf('.') != -1)
        {
          imageFileNumber = imageFileNumber.Substring(0, 1);
        }
        Stream stream = i.GetStream();
        long length = stream.Length;
        byte[] byteStream = new byte[length];
        stream.Read(byteStream, 0, (int)length);

        string imageAsString = Convert.ToBase64String(byteStream);
        Console.WriteLine("The rId of this Image is '{0}' and he image file number is image{1}.png", rId, imageFileNumber);

        try
        {
          this.images.Add(rId.ToString(), byteStream);
        }
        catch (Exception e)
        {
          Console.WriteLine(e.Message);
        }
        Console.WriteLine("The rId of this Image is '{0}' and he image file number is image{1}.png", rId, imageFileNumber);
      }

    }
  }
}
