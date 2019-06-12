using System.Drawing;
using Aspose.Slides;
using LCC_xlsx_to_pptx_converter.Class.datas;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class changeImageSlide
  {
    public static void run(Presentation pres, int imageWanted, int WB, int slideTarget)
    {
      IPPImage oldImage;
      using (Image newImage = Image.FromFile(getImage.run(getProgramDirectory.run(), WB, imageWanted)))
      {
        for (int i = 1; i <= pres.Images.Count - 1; i++)
        {
          oldImage = pres.Images[i];
          oldImage.ReplaceImage(newImage);
        }
      }
    }

    public static void run2(Presentation pres, int imageWanted, int WB, int slideTarget)
    {
      using (Image newImage = Image.FromFile(getImage.run(getProgramDirectory.run(), WB, imageWanted)))
      {
        IPPImage oldImage;
        oldImage = pres.Images[slideTarget];
        oldImage.ReplaceImage(newImage);
      }
    }
  }
}
