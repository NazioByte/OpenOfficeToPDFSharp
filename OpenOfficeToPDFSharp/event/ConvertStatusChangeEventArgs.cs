namespace ZioByte.OpenOffice
{
    public class ConvertStatusChangeEventArgs
    {
        public string Status { get; set; }
        public string FileName { get; set; }

        public ConvertStatusChangeEventArgs(string filename, string status)
        {
            this.Status = filename;
            this.FileName = status;
        }
    }
}