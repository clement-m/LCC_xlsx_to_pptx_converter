namespace LCC_xlsx_to_pptx_converter.Class.main
{
  public class DataSet
  {
    private int    workbookId;
    private int    worksheetId;
    private int    row;
    private int    col;
    private string textContent;
    private byte[] image;

    public DataSet(int workbookId, int worksheetId, int row, int col, string textContent)
    {
      this.workbookId  = workbookId;
      this.worksheetId = worksheetId;
      this.row         = row;
      this.col         = col;
      this.textContent = textContent;
    }

    public void addImage(byte[] image)
    {
      this.image = image;
    }
  }
}
