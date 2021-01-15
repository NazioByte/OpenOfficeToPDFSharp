using System;
using System.Collections.Generic;
using System.Text;

namespace ZioByte.OpenOffice
{
    public class OpenOfficeToPDF
    {
        private ProcessInterface process = new ProcessInterface();

        private const string STR_PDF_EXPORT = "writer_pdf_Export";
        private const string STR_PDF_OVERWRITE = "Overwriting";

        private StringBuilder bf;
        private IList<string> cmd = new List<string>();

        private string app_3rd_path = System.IO.Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            @"dist\program", "soffice");

        public OpenOfficeToPDF()
        {
            process.OnProcessOutput += Process_OnProcessOutput;
        }

        public void Set(string sourcePath, string outputDirectory)
        {
            var app_args = $"--headless -convert-to pdf {sourcePath} --outdir {outputDirectory}";

            cmd.Add($"{app_3rd_path} {app_args}");
        }

        public void Convert()
        {
            bf = new StringBuilder();

            foreach (string content in cmd)
            {
                bf.Append(content + " && ");
            }

            var str_result_conv = "/C " + bf.ToString().Remove(bf.ToString().Length - 3);

            Console.WriteLine(str_result_conv);

            process.StartProcess("cmd.exe", str_result_conv);
        }

        private void Process_OnProcessOutput(object sender, ProcessEventArgs args)
        {
            if (args.Content.IndexOf(STR_PDF_EXPORT) != -1)
            {
                var temp_filename = args.Content.Split(new char[] { '-', '>' }, StringSplitOptions.RemoveEmptyEntries);

                var filename_pdf = temp_filename[1].Replace("using filter : writer_pdf_Export", "").Replace("Overwriting: ", "").Trim();

                OnMessageReceived(new ConvertStatusChangeEventArgs(filename_pdf, "Writing"));
            }

            if (args.Content.IndexOf(STR_PDF_OVERWRITE) != -1)
            {
                var temp_filename = args.Content.Replace("Overwriting: ", "").Trim();

                OnMessageReceived(new ConvertStatusChangeEventArgs(temp_filename, STR_PDF_OVERWRITE));
            }

            if (args.Content == string.Empty)
            {
                OnMessageReceived(new ConvertCompleteEventArgs(true));
            }
        }

        public delegate void ConvertCompleteEventHandler(object sender, ConvertCompleteEventArgs args);
        public delegate void ConvertStatusChangeEventHandler(object sender, ConvertStatusChangeEventArgs args);

        public event ConvertCompleteEventHandler ConvertComplete;
        public event ConvertStatusChangeEventHandler ConvertStatusChange;

        protected virtual void OnMessageReceived(ConvertCompleteEventArgs args)
        {
            if (ConvertComplete != null)
                ConvertComplete(this, args);
        }

        protected virtual void OnMessageReceived(ConvertStatusChangeEventArgs args)
        {
            if (ConvertStatusChange != null)
                ConvertStatusChange(this, args);
        }

    }
}
