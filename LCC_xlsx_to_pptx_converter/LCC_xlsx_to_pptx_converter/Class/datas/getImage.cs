namespace LCC_xlsx_to_pptx_converter.Class.datas
{
  class getImage
  {
    public static string run(string dataDir, int WorkBook, int imageNumber)
    {
      return dataDir +
        "WorkBook" + WorkBook + "\\" +
        imageNumber + ".png"
      ;
    }
  }
}
