namespace ZioByte.OpenOffice
{
    public class ConvertCompleteEventArgs
    {
        public bool Success { get; set; } = false;

        public ConvertCompleteEventArgs(bool status)
        {
            this.Success = status;
        }
    }
}
