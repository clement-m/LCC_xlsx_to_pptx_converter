using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  class changeImageSlide
  {
    public static void run(Presentation pres, Slide slide, string imageName)
    {
      IShape shape = FindShape(slide, "Picture 11");
      int toto = 1;

      if (shape != null)
      {
        IPictureFrame picFrame = (IPictureFrame)shape;

        Image newImage = Image.FromFile(imageName);
        IPPImage oldImage = pres.Images[2];
        foreach(IPPImage img in pres.Images)
        {
          img.ReplaceImage(newImage);
        }

      }
    }
    /*
    public static void TestImageReplace(string dataDir)
    {
      Presentation TemplatePres = new Presentation(dataDir + "BCNE_19S07_VDI_T-CSFB-D-S_OSM-CIMENTERIE-EQIOM.pptx");

      Presentation TargetPres = new Presentation(dataDir + "SORTIE_PPTX.pptx");

      ISlide slide2Copy = TemplatePres.Slides[0];

      ISlideCollection TargetPresSlides = TargetPres.Slides;

      //Now Copying Slide from Template To Target
      //We will copy 5 slides
      ISlide slide = null;
      for (int i = 0; i < 5; i++)
      {
        slide = TargetPresSlides.AddClone(slide2Copy);


        // alternative text of the shape to be found
        IShape shape = FindShape(slide, “ImageTestname”);

        if (shape != null)
        {
          IPictureFrame picFrame = (IPictureFrame)shape;

          Image newImage = Image.FromFile(“img.jpg”);
          IPPImage oldImage = TargetPres.Images[2];
          oldImage.ReplaceImage(newImage);

        }
      }


      TargetPres.Save(“SavedPres.pptx”, SaveFormat.Pptx);
    }
    */


    // Method implementation to find a shape in a slide using its alternative text
    public static IShape FindShape(ISlide slide, string alttext)
    {
      // Iterating through all shapes inside the slide
      for (int i = 0; i < slide.Shapes.Count; i++)
      {
        // If the alternative text of the slide matches with the required one then
        // return the shape 
        if (slide.Shapes[i].Name.CompareTo("Picture 11") == 0
        || slide.Shapes[i].Name.CompareTo("Picture 12") == 0)
          return slide.Shapes[i];
      }
      return null;
    }

    
  }
}
