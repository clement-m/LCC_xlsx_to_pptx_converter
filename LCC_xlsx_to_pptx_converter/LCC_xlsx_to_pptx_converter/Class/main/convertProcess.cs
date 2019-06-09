using System;
using System.IO;
using LCC_xlsx_to_pptx_converter.Class.main;
using LCC_xlsx_to_pptx_converter.Class.xlsx;
using LCC_xlsx_to_pptx_converter.Class.pptx;

namespace LCC_xlsx_to_pptx_converter.Class
{
  public static class convertProcess
  {
    public static void run(string customerName)
    {
      Console.WriteLine("LCC xlsx_to_pptx_convertor Launched...");
      Console.WriteLine("=====================================================");
      string pathFolder = GetDataDir_Data();
      Console.WriteLine("Step I:\n-Copying file");
      CopyXlsxFile.run(pathFolder);
      Console.WriteLine("-Copying file ok");
      Console.WriteLine("-Opening and extract data");
      Data D = OpenXlsx.run(pathFolder);
      Console.WriteLine("-data extracted ok");
      Console.WriteLine("Step II:\n-Create A-Version Presentation"); 
      Aspose.Slides.Presentation pres = createPPTX.run(pathFolder, D, customerName);
      Console.WriteLine("-Dispose data");
      D.dispose();
      Console.WriteLine("-Data disposed");
      Console.WriteLine("-Delete temporaries file");
      DeleteXlsx.run(pathFolder);
      Console.WriteLine("-Files deleted");
      Console.WriteLine("Step III:\n-Cleaning Images");
      //DeleteImages.run(pathFolder);
      Console.WriteLine("-Cleaning success");
      Console.WriteLine("Step IV:\n-Converting A-Version to PPTX");
      Clean.run(pres, pathFolder);
      Console.WriteLine("-Convertion succeeded");
      Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
      Console.ReadLine();
    }

    private static string GetDataDir_Data()
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
