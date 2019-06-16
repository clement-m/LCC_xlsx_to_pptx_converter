using System;
using System.Collections.Generic;
using System.Windows.Forms;
using LCC_xlsx_to_pptx_converter.Class.main;

namespace LCC_xlsx_to_pptx_converter
{
  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
    }

    private void Button2_Click(object sender, EventArgs e)
    {
      List<string> listFile = new List<string>();

      try
      {
        var o = new OpenFileDialog();

        o.Multiselect = true;

        if (o.ShowDialog() == DialogResult.OK)
        {
          string[] ddd = o.FileNames;
          
          int size = 0;

          foreach(string moncul in ddd)
          {
            size++;
          }

          for (int i = 0; i <= size - 1; i++)
          {
            listFile.Add(o.FileNames[i]);
          }
        }
        else
        {
          MessageBox.Show("File Not Uploaded", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }

      convertProcess.run(listFile, textBoxTitle.Text);
    }
  }
}