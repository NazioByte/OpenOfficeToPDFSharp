using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZioByte.OpenOffice;

namespace Sample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

       

        private void Form1_Load(object sender, EventArgs e)
        {

           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenOfficeToPDF officeToPDF = new OpenOfficeToPDF();

            officeToPDF.ConvertComplete += OfficeToPDF_ConvertComplete; 
          
            officeToPDF.Convert(@"C:\Users\ACER\Desktop\test\a\sample4.docx", @"C:\Users\ACER\Desktop\test\o");
        }

        private void OfficeToPDF_ConvertComplete(object sender, ConvertCompleteEventArgs args)
        {
            MessageBox.Show(args.FileName);
        }
    }
}
