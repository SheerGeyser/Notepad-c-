using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace ENDE_coder
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.Filter = "Text File (.txt)|*.txt|Word File (.docx ,.doc)|*.docx;*.doc"; //|PDF (.pdf)|*.pdf|Spreadsheet (.xls ,.xlsx)|  *.xls ;*.xlsx|Presentation (.pptx ,.ppt)|*.pptx;*.ppt"
            saveFileDialog1.Filter = "Text File (.txt)|*.txt|Word File (.docx ,.doc)|*.docx;*.doc";
        }

        private void oPENToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                FileInfo file = new FileInfo(openFileDialog1.FileName);
                string extenshion = file.Extension;
                string path1 = openFileDialog1.FileName;

                if (extenshion == ".docx" || extenshion == ".doc")
                {
                    Word._Application wordApp = new Word.Application();
                    object objFile = path1;
                    object objNull = System.Reflection.Missing.Value;
                    object objReadOnly = true;


                    Word._Document Doc = wordApp.Documents.Open(ref objFile, ref objNull, ref objReadOnly,
                    ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull,
                    ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull);

                    int i = 1;
                    foreach (Word.Paragraph objParagraph in Doc.Paragraphs)
                    {
                        try
                        {
                            richTextBox1.Text = null;
                            richTextBox1.Text += Doc.Paragraphs[i].Range.Text;
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        i++;
                    }

                    Doc.Close(ref objNull, ref objNull, ref objNull);
                    wordApp.Quit(ref objNull, ref objNull, ref objNull);
                }
            }
        }

        private void sAVEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            // получаем выбранный файл
            string filename = saveFileDialog1.FileName;
            // сохраняем текст в файл
            System.IO.File.WriteAllText(filename, richTextBox1.Text);
        }
    }
}
