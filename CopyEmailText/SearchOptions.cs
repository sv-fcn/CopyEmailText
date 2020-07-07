namespace CopyEmailText
{
    public class SearchOptions
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string SearchFrom { get; set; }
        public string SearchSubject { get; set; }
        public int SearchNumberOfMessages { get; set; }
        public int ShowConsoleSeconds { get; set; }

    }
}
