namespace CopyEmailText
{
    public class SearchOptions
    {
        public TestMode TestMode { get; set; }
        public string Host { get; set; }
        public int Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string SearchFrom { get; set; }
        public string SearchSubject { get; set; }
        public int SearchNumberOfMessages { get; set; }
        public int EmailMaxValidAgeSeconds { get; set; }
        public bool DeleteMessages { get; set; }
        public int ShowConsoleSeconds { get; set; }
    }

    public class TestMode
    {
        private bool _enabled = false;
        private bool _imapConnect = false;
        private bool _deleteMessages = false;

        public bool Enabled
        {
            get => _enabled;
            set => _enabled = value;
        }

        public bool ImapConnect
        {
            get => _enabled && _imapConnect;
            set => _imapConnect = value;
        }

        public bool DeleteMessages
        {
            get => _enabled && _deleteMessages;
            set => _deleteMessages = value;
        }

        public bool TestModeSetting( bool setting, out bool testModeEnabled )
        {
            testModeEnabled = _enabled;

            if( _enabled )
            {
                return setting;
            }

            return false;
        }

        public bool TestModeSetting( bool setting )
        {
            return TestModeSetting( setting, out _ );
        }
    }
}
