using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZioByte.OpenOffice
{
    public class OpenOfficeToPDF
    {
        private ProcessInterface process = new ProcessInterface();

        private string app_3rd_path = System.IO.Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            @"dist\program", "soffice.exe");

        protected string SourcePath { get; set; }
        protected string OutputDirPath { get; set; }

        public OpenOfficeToPDF(string sourcePath, string outputDirectory)
        {
            this.SourcePath = sourcePath;
            this.OutputDirPath = outputDirectory;

            process.OnProcessOutput += Process_OnProcessOutput;
        }

        public void Convert()
        {
            var app_args = $"--headless -convert-to pdf {this.SourcePath} --outdir {this.OutputDirPath}";

            process.StartProcess(app_3rd_path, app_args);
        }

        private void Process_OnProcessOutput(object sender, ProcessEventArgs args)
        {
            if (args.Content.IndexOf("writer_pdf_Export") != -1)
            {
                OnMessageReceived(new ConvertCompleteEventArgs("Successs"));
            }
        }

        public delegate void ConvertCompleteEventHandler(object sender, ConvertCompleteEventArgs args);
        public event ConvertCompleteEventHandler ConvertComplete;
        protected virtual void OnMessageReceived(ConvertCompleteEventArgs args)
        {
            if (ConvertComplete != null)
                ConvertComplete(this, args);
        }

    }
}
