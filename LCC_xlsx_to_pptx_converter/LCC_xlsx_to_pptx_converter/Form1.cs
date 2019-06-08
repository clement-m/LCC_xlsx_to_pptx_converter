using System;
using System.Windows.Forms;
using LCC_xlsx_to_pptx_converter.Class;

namespace LCC_xlsx_to_pptx_converter
{
  public partial class Form1 : Form
  {
    const string DIR_PATH = @"C:\Users\daggo\Desktop\pptx_xlsx\extract analyze et rapport\";

    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      convertProcess.run();
    }
  }
}