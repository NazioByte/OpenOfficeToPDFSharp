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

        private string SourcePath { get; set; }
        private string OutPutDir { get; set; }

        private  string app_3rd_path = System.IO.Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            @"dist\program", "soffice");       

        public OpenOfficeToPDF()
        {
            process.OnProcessOutput += Process_OnProcessOutput;
        }

        public void Convert(string sourcePath, string outputDirectory)
        {
            var app_args = $"--headless -convert-to pdf {sourcePath} --outdir {outputDirectory}";

            this.SourcePath = sourcePath;
            this.OutPutDir = outputDirectory;

            process.StartProcess(app_3rd_path, app_args);
        }

        private void Process_OnProcessOutput(object sender, ProcessEventArgs args)
        {
            var new_pdf_file = System.IO.Path.Combine(this.OutPutDir,  
                System.IO.Path.GetFileNameWithoutExtension
                (this.SourcePath) + ".pdf");

            if (args.Content.IndexOf("writer_pdf_Export") != -1)
            {
                OnMessageReceived(new ConvertCompleteEventArgs(new_pdf_file, "Successs"));
                process.StopProcess();
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
