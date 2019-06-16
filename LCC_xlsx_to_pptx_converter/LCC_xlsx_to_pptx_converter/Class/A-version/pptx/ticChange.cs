using System.Collections.Generic;

namespace LCC_xlsx_to_pptx_converter.Class.openXML
{
  class ticChange
  {
    public static void run(ref int slideId, ref int switchTic, ref int WB, List<string> listFile,
      int slideNumberInGroup, int slideIdReplacing)
    {
      if (listFile.Count != 1)
      {
        WB++;
        if (WB == listFile.Count + 1)
        {
          WB = 1;
          slideId = slideId + (slideNumberInGroup * (listFile.Count - 1)) + 1;
        }
        else
        {
          switchTic = slideIdReplacing;
          slideId = slideIdReplacing;
        }
      }
    }
  }
}
