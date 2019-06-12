using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class changeTextInSlide
  {
    public static void run(Slide slide, string text)
    {
      foreach (IShape shp in slide.Shapes)
      if (shp.Placeholder != null)
      {
        try
        {
            ((IAutoShape)shp).TextFrame.Text = text;
        }
        catch
        {

        }  
      }
    }
  }
}
