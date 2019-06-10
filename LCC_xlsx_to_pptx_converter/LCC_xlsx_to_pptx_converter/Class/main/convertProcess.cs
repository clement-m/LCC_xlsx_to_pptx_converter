using System;
using System.Collections.Generic;
using LCC_xlsx_to_pptx_converter.Class.main;
using LCC_xlsx_to_pptx_converter.Class.xlsx;
using LCC_xlsx_to_pptx_converter.Class.pptx;
using LCC_xlsx_to_pptx_converter.Class.datas;

namespace LCC_xlsx_to_pptx_converter.Class
{
  public static class convertProcess
  {
    public static void run(List<string> listFile, string title)
    {
      Console.WriteLine("LCC xlsx_to_pptx_convertor Launched...");
      Console.WriteLine("=====================================================");

      //Console.WriteLine("Step I:\n-Copying file");
      //CopyXlsxFile.run(listFile);
      //Console.WriteLine("done.");

      Console.WriteLine("-Opening and extract data");
      Data D = OpenXlsx.run(listFile);
      Console.WriteLine("done.");

      //Console.WriteLine("-Delete temporaries file");
      //DeleteXlsx.run(getProgramDirectory.run());
      //Console.WriteLine("done.");

      Console.WriteLine("Step II:\n-Create A-Version Presentation"); 
      Aspose.Slides.Presentation pres = createPPTX.run(D, title);
      Console.WriteLine("done.");

      Console.WriteLine("Step IV:\n-Converting A-Version to PPTX");
      Clean.run(pres, getProgramDirectory.run());
      Console.WriteLine("-Convertion succeeded");

      Console.WriteLine("-Dispose data");
      D.dispose();
      Console.WriteLine("done.");

      Console.WriteLine("Step III:\n-Cleaning Images");
      DeleteImages.run(getProgramDirectory.run(), listFile);
      Console.WriteLine("done.");

      Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
      Console.ReadLine();

      listFile.Clear();
      pres.Dispose();
    }
  }
}
