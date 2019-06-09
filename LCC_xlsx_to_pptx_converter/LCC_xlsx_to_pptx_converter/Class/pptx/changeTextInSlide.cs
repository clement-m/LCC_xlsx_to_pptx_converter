using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  class changeTextInSlide
  {
    public static void run(Slide slide, string text)
    {
      foreach (IShape shp in slide.Shapes)
      if (shp.Placeholder != null)
      {
        ((IAutoShape)shp).TextFrame.Text = text;
      }
    }
  }
}
