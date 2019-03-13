using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  class PPTXData
  {
    public List<string> images;
    public Dictionary<string, byte[]> imagesData;

    public PPTXData()
    {
      this.images = new List<string>();
      this.imagesData = new Dictionary<string, byte[]>();
    }

    public void addImagesData(string name, byte[] data)
    {
      try
      {
        this.imagesData.Add(name, data);
      } catch(Exception e) {
        
      }
      if (this.imagesData[name] == null) {
        
      }
    }

    public List<string> getImages(SlidePart slidePartTemplate, string relId)
    {
      // Copy the image parts
      foreach (ImagePart image in slidePartTemplate.ImageParts)
      {
        string imageFileNumber = image.Uri.OriginalString.Substring(16, 2);
        if (imageFileNumber.IndexOf('.') != -1)
        {
          imageFileNumber = imageFileNumber.Substring(0, 1);
        }

        this.images.Add("image" + imageFileNumber + ".png");

        Console.WriteLine("Slide id: {0}: image 1: image{1}.png", relId, imageFileNumber);
        ImagePart imageClone = image;
        //ImagePart imageClone = slidePartClone.AddImagePart(image.ContentType, slideTemplate.GetIdOfPart(image));
        //using (var imageStream = image.GetStream())
        //{
        //  imageClone.FeedData(imageStream);
        //}
      }

      return this.images;
    }
  }
}
