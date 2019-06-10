namespace LCC_xlsx_to_pptx_converter
{
  partial class Form1
  {
    /// <summary>
    /// Variable nécessaire au concepteur.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Nettoyage des ressources utilisées.
    /// </summary>
    /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Code généré par le Concepteur Windows Form

    /// <summary>
    /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
    /// le contenu de cette méthode avec l'éditeur de code.
    /// </summary>
    private void InitializeComponent()
    {
      this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
      this.textBoxTitle = new System.Windows.Forms.TextBox();
      this.label1 = new System.Windows.Forms.Label();
      this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
      this.button2 = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // openFileDialog1
      // 
      this.openFileDialog1.FileName = "openFileDialog1";
      this.openFileDialog1.Multiselect = true;
      // 
      // textBoxTitle
      // 
      this.textBoxTitle.Location = new System.Drawing.Point(161, 12);
      this.textBoxTitle.Name = "textBoxTitle";
      this.textBoxTitle.Size = new System.Drawing.Size(100, 20);
      this.textBoxTitle.TabIndex = 1;
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(50, 15);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(105, 13);
      this.label1.TabIndex = 2;
      this.label1.Text = "Nom lieu / entreprise";
      // 
      // button2
      // 
      this.button2.Location = new System.Drawing.Point(304, 120);
      this.button2.Name = "button2";
      this.button2.Size = new System.Drawing.Size(75, 23);
      this.button2.TabIndex = 3;
      this.button2.Text = "start";
      this.button2.UseVisualStyleBackColor = true;
      this.button2.Click += new System.EventHandler(this.Button2_Click);
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(391, 155);
      this.Controls.Add(this.button2);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.textBoxTitle);
      this.Name = "Form1";
      this.Text = "Form1";
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion
    private System.Windows.Forms.OpenFileDialog openFileDialog1;
    private System.Windows.Forms.TextBox textBoxTitle;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    private System.Windows.Forms.Button button2;
  }
}

