using System;
using System.Windows.Forms;
using LCC_xlsx_to_pptx_converter.Class;

namespace LCC_xlsx_to_pptx_converter
{
  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      convertProcess.run(textBoxClient.Text);
    }
  }
}