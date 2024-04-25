using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Shifa_Lab.Properties
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchcase = true;
            object matchWholeword = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchkashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;


            wordApp.Selection.Find.Execute(ref ToFindText,
            ref matchcase, ref matchWholeword, ref matchWildCards,
            ref matchSoundLike, ref nmatchAllforms, ref forward, ref wrap,
            ref format, ref replaceWithText, ref replace, ref matchkashida,
            ref matchDiactitics, ref matchAlefHamza, ref matchControl);


        }

        private void CreateWordDocument(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();



                //find and replace
                this.FindAndReplace(wordApp, "<name>", name.Text);
                this.FindAndReplace(wordApp, "<age>", age.Text);
                this.FindAndReplace(wordApp, "<sex>", sex.Text);
                this.FindAndReplace(wordApp, "<lab>", lab.Text);
                this.FindAndReplace(wordApp, "<ref>", prefer.Text);
                this.FindAndReplace(wordApp, "<date>", date.Text);
                this.FindAndReplace(wordApp, "<pot>", pot.Text);
                this.FindAndReplace(wordApp, "<sod>", sod.Text);
                this.FindAndReplace(wordApp, "<chl>", chlo.Text);
                this.FindAndReplace(wordApp, "<cre>", crea.Text);
                this.FindAndReplace(wordApp, "<ure>", urea.Text);


            }
            else
            {
                MessageBox.Show("File not found!");
            }

            //Save as
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created!");


        }


        private void button5_Click(object sender, EventArgs e)
        {
            //Pick up Template document                      //Address of new Document file            
            CreateWordDocument(@"F:\Shifa_Lab\SE,RPM.docx", @"F:\Lab Reports\" + file.Text + " (SE,RPM).docx");

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            file.Text = String.Empty;
            name.Text = String.Empty;
            sex.Text = String.Empty;
            age.Text = String.Empty;
            prefer.Text = String.Empty;
            lab.Text = String.Empty;
            pot.Text = String.Empty;
            sod.Text = String.Empty;
            chlo.Text = String.Empty;
            crea.Text = String.Empty;
            urea.Text = String.Empty;

            MessageBox.Show("Input fields Cleared.");
        }
    }
}
