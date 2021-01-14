using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ConvertCompleteEventArgs
{
    public string Status { get; set; } 
    public ConvertCompleteEventArgs(string message)
    {
        this.Status = message;
    }
}