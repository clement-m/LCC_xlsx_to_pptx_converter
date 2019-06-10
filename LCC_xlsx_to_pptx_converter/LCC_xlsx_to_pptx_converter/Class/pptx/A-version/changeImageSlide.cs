using System.Drawing;
using Aspose.Slides;
using LCC_xlsx_to_pptx_converter.Class.datas;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  class changeImageSlide
  {
    public static void run(Presentation pres,int WB, int imageNumber, int slideTargetImage)
    {
      IPPImage oldImage;

      Image newImage = Image.FromFile(getImage.run(getProgramDirectory.run(), WB, imageNumber));

      for(int i = 11; i <= 250; i++)
      {
        oldImage = pres.Images[i];
        oldImage.ReplaceImage(newImage);
      }

      oldImage = null;
      newImage.Dispose();
    }
  }
}
