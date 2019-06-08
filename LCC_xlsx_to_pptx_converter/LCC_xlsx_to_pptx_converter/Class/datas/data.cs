using System.Collections.Generic;
using System.Linq;

namespace LCC_xlsx_to_pptx_converter.Class.main
{
  public class Data
  {
    private Dictionary<int, DataSet> D = new Dictionary<int, DataSet>();

    public void addDataSet(DataSet dataset)
    {
      D.Add(D.Count() + 1, dataset);
    }

    public void dispose()
    {
      D = null;
    }
  }
}
