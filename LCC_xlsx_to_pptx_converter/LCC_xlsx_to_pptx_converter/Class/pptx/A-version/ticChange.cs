using System.Collections.Generic;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class ticChange
  {
    public static void run(ref int slideId, ref int switchTic, ref int WB, List<string> listFile)
    {
      if (listFile.Count != 1)
      {
        WB++;
        if (WB == listFile.Count + 1)
        {
          WB = 1;
          slideId = slideId + (3 * (listFile.Count - 1)) + 1;
        }
        else
        {
          slideId = 11;
          switchTic = 11;
        }
      }
    }
  }
}
