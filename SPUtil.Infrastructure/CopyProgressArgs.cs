namespace SPUtil.Infrastructure
{
    public class CopyProgressArgs
    {
        public int Processed { get; set; }
        public int Total { get; set; }
        public required string Message { get; set; }
    }
}
