using Aspose.Cells;

namespace LCC_xlsx_to_pptx_converter.Class.xlsx
{
  public static class CopyXlsxFile
  {
    public static void run(string dataDir)
    {
      Workbook workbook1 = new Workbook(dataDir + "Outdoor.xlsx");
      Workbook workbook2 = new Workbook(dataDir + "RDC Eqiom heming.xlsx");
      Workbook workbook3 = new Workbook(dataDir + "r+1 bureaux administratifs.xlsx");
      Workbook workbook4 = new Workbook(dataDir + "r+2 bureaux administratifs.xlsx");

      workbook1.Save(dataDir + "TEMP1.xlsx", Aspose.Cells.SaveFormat.Xlsx);
      workbook2.Save(dataDir + "TEMP2.xlsx", Aspose.Cells.SaveFormat.Xlsx);
      workbook3.Save(dataDir + "TEMP3.xlsx", Aspose.Cells.SaveFormat.Xlsx);
      workbook4.Save(dataDir + "TEMP4.xlsx", Aspose.Cells.SaveFormat.Xlsx);
    }
  }
}
