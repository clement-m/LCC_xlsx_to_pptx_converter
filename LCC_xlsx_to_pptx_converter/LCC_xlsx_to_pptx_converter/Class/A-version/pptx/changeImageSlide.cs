using System.Drawing;
using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class changeImageSlide
  {
    public static void run2(A.Presentation pres, int imageWanted, int WB, int slideTarget)
    {
      using (Image newImage = Image.FromFile(getImage.run(getProgramDirectory.run(), WB, imageWanted)))
      {
        A.IPPImage oldImage;
        oldImage = pres.Images[slideTarget];
        oldImage.ReplaceImage(newImage);
      }
    }
  }
}
