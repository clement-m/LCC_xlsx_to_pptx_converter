using Aspose.Slides;
using LCC_xlsx_to_pptx_converter.Class.main;
using System.Drawing;
using System.IO;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  public static class createPPTX
  {
    public static Presentation run(string dataDir, Data D, string customerName)
    {
      using (Presentation pres = new Presentation(dataDir + "BCNE_19S07_VDI_T-CSFB-D-S_OSM-CIMENTERIE-EQIOM.pptx"))
      {
        int slideId = 0;
        foreach (Slide slide in pres.Slides)
        {
          slideId++;

          switch (slideId)
          {
            case 11:
              changeTextInSlide.run(slide, customerName);
              break;
            case 12:
              changeImageSlide.run(pres, slide, getImage(dataDir, 2, 1, 11));
              break;
            default:
              changeImageSlide.run(pres, slide, getImage(dataDir, 2, 1, 11));
              break;
          break;
          }
        }
        //pres.Save(dataDir + "RectShpPic_out.pptx", SaveFormat.Pptx);
        return pres;
      }
    }

    public static string getImage(string dataDir, int WorkBook, int WorkSheet, int imageNumber)
    {
      return dataDir +
        "WorkBook" + WorkBook + "\\" +
        "WorkSheet" + WorkSheet + "\\" +
        imageNumber + ".png"
      ;
    }
  }
}
