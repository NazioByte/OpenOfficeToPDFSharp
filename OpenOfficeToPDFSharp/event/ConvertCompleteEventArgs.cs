using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZioByte.OpenOffice
{
    public class ConvertCompleteEventArgs
    {
        public string Status { get; set; }
        public string FileName { get; set; }
        public ConvertCompleteEventArgs(string filename, string message)
        {
            this.Status = message;
            this.FileName = filename;
        }
    }
}
