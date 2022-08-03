using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace InsertHeaderPDF
{
    public partial class Form1 : Form
    {
        string program = @"c:\vjrun.exe";
        string project = @"C:\ETIQUETA.vjr";
        string dataFile = @"C:\process.txt";
        string printer = "Adobe PDF";
        string options_1 = "INICIAL=1";
        string options_2 = "TOTAL=100";
        string fileTmp = string.Empty;


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process process = new Process();
            process.StartInfo.FileName = program;
            process.StartInfo.Arguments = $"print -p\"{project}\" -d\"{dataFile}\" -w\"{printer}\"" +
                                          $" -yS -a\"{options_1}\" -a\"{options_2}\"";
            process.Start();
            process.WaitForExit();

            //Thread.Sleep(80000);
            Thread.Sleep(5000);

            var fileInfo = new FileInfo(@"D:\Debug\DADOS\GRAFICA_SENADOR\QR_CODE_SEQ.pdf");

            fileTmp = $"{fileInfo.Directory}\\Tmp_{fileInfo.Name}";

            FontFactory.RegisterDirectory("C:\\WINDOWS\\Fonts");

            float pos_x = Utilities.MillimetersToPoints(100);
            float pos_y = Utilities.MillimetersToPoints(10);

            using (PdfReader reader = new PdfReader(fileInfo.FullName))
            {

                var pageRotation = reader.GetPageRotation(1);
                var pageWidth = reader.GetPageSizeWithRotation(1).Width;
                var pageHeight = reader.GetPageSizeWithRotation(1).Height;

                //Size Page 
                iTextSharp.text.Rectangle sizePage = reader.GetPageSizeWithRotation(1);

                //Document
                using (Document doc = new Document(new iTextSharp.text.Rectangle(sizePage.Width, sizePage.Height)))
                {

                    PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(fileTmp, FileMode.Create));
                    doc.Open();

                    var cb = writer.DirectContent;

                    doc.NewPage();

                    int ini = 1, end = 15, reg = 15;

                    for (int indexPage = 1; indexPage <= reader.NumberOfPages; indexPage++)
                    {
                        var page = writer.GetImportedPage(reader, indexPage);

                        cb.AddTemplate(page, 0, 0);

                        cb.BeginText();

                        var font = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.BOLD, new BaseColor(0, 0, 0));

                        cb.SetColorFill(font.Color);

                        cb.SetFontAndSize(font.BaseFont, font.Size);

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, $"{ini:00000} até {end:00000}", pos_x, page.Height - pos_y, 0);

                        cb.EndText();

                        doc.NewPage();

                        ini = end + 1;
                        end += reg;
                    }

                }

            }
            //doc.Close();

            MessageBox.Show("done!");
        }
    }
}
