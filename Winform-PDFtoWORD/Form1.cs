using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using org.apache.pdfbox.pdfviewer;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using Xceed.Words.NET;

namespace Winform_PDFtoWORD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "PDF files |*.pdf" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtName.Text = ofd.FileName;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            PDDocument doc = PDDocument.load(txtName.Text);
            PDFTextStripper stripper = new PDFTextStripper();
            richTextBox1.Text = (stripper.getText(doc));
            var docName = Path.GetFileNameWithoutExtension(txtName.Text) + ".docx";
            var worddoc = DocX.Create(docName);
            worddoc.InsertParagraph(richTextBox1.Text);
            worddoc.Save();
            Process.Start(docName);
        }
    }
}
