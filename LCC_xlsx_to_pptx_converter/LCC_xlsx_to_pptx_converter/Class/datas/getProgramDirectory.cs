using System.IO;

namespace LCC_xlsx_to_pptx_converter.Class.datas
{
  class getProgramDirectory
  {
    public static string run()
    {
      var parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
      string startDirectory = null;
      if (parent != null)
      {
        var directoryInfo = parent.Parent;
        if (directoryInfo != null)
        {
          startDirectory = directoryInfo.FullName;
        }
      }
      else
      {
        startDirectory = parent.FullName;
      }
      return startDirectory != null ? Path.Combine(startDirectory, "Data\\") : null;
    }
  }
}
