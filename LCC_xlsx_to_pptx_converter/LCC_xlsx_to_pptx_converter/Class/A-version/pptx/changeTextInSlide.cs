using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class changeTextInSlide
  {
    public static void run(A.Slide slide, string text)
    {
      foreach (A.IShape shp in slide.Shapes)
      if (shp.Placeholder != null)
      {
        try
        {
            ((A.IAutoShape)shp).TextFrame.Text = text;
        }
        catch
        {

        }  
      }
    }
  }
}
