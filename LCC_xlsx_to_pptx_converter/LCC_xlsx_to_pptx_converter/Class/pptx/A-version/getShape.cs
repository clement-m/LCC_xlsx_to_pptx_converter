using System.Linq;
using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class getShape
  {
    public static IShape run(Slide slide, string alttext)
    {
      for (int i = 0; i < slide.Shapes.Count; i++)
      {
        if (slide.Shapes[i].Name.CompareTo("Picture 11") == 0
        || slide.Shapes[i].Name.CompareTo("Picture 12") == 0)
          return slide.Shapes[i];
      }
      return null;
    }
  }
}
