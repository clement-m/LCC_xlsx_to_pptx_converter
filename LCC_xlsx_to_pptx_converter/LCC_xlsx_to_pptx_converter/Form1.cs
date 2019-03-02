using System;
using System.Windows.Forms;
using LCC_xlsx_to_pptx_converter.Class;

namespace LCC_xlsx_to_pptx_converter
{
  
  public partial class Form1 : Form
  {
    const string DIR_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\";

    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      string path = DIR_PATH + "make.pptx";

      Console.WriteLine("Creating pptx:");
      PPTXClass.CreatePresentation(path);

      int slideId = 1;

      Console.WriteLine("-add slide " + slideId + " to make.pptx");
      PPTXClass.addSlide(path, slideId, "test slide 1");

      string imagePath = "test.png";

      Console.WriteLine("-add image " + imagePath + " to slide " + slideId + " in make.pptx");
      //PPTXClass.InsertImageInLastSlide(imagePath);

      slideId++;

      Console.WriteLine("-add slide " + slideId + " to make.pptx");
      PPTXClass.addSlide(path, slideId, "test slide 3");
      
      //Console.WriteLine("-add image " + imagePath + " to slide " + slideId + " in make.pptx");
      //PPTXClass.InsertImageInLastSlide(imagePath);
      
      // SUITE DU PROGRAMME

      // Ajouter une image venant du fichier excel

      // verifier la correspondance

      // ajouter les deux images aux bonnes positions pour les gros titres

      // ajouter les images 

      
      Console.WriteLine("END PROGRAM");
      Console.WriteLine("CONVERTION XLSX to PPTX SUCCESS");
    }
  }
}