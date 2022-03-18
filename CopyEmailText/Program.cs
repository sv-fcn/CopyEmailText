using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Security;
using System.Threading;
using GemBox.Email;
using GemBox.Email.Imap;
using GemBox.Email.Mime;
using GemBox.Email.Security;
using TextCopy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.Security.Cryptography.X509Certificates;

namespace CopyEmailText
{
    class Program
    {
        private const int HEIGHT = 25;
        private const int WIDTH = 75;
        private const ConsoleColor BACKGROUNDCOLOR = ConsoleColor.Black;
        private static IConfigurationRoot _configuration;

        static void Main( string[] args )
        {
            SearchOptions options = null;
            int defaultBufferHeight = Console.BufferHeight;

            try
            {
                Console.Title = nameof( CopyEmailText );
                Console.WindowHeight = HEIGHT;
                Console.WindowWidth = WIDTH;
                Console.BufferHeight = HEIGHT;
                Console.BufferWidth = Console.WindowWidth;
                Console.BackgroundColor = BACKGROUNDCOLOR;
                Console.Clear( );

                WriteColor( $"Reading appsettings...", false );
                var serviceCollection = new ServiceCollection( );
                ConfigureServices( serviceCollection );
                options = _configuration.GetSection( "Options" ).Get<SearchOptions>( );
                WriteColor( " Success", ConsoleColor.DarkGreen );

                if( options.TestMode.Enabled )
                {
                    WriteColor( "[ Test Mode Enabled ]", ConsoleColor.Yellow );
                }

                using( var imapClient = ImapConnect( options ) )
                {
                    var emailList = GetEmailIdList( options, imapClient );

                    if( !emailList.Any( ) )
                    {
                        throw new Exception( "No email found" );
                    }

                    var email = GetTextFromEmail( emailList, options.SearchSubject );

                    if( string.IsNullOrWhiteSpace( email.text ) )
                    {
                        throw new Exception( "Email contained no text" );
                    }

                    WriteColor( $"Copying password to clipboard...", false );
                    var c = new Clipboard( );
                    c.SetText( email.text );
                    WriteColor( " Success", ConsoleColor.DarkGreen );

                    Console.WriteLine( );
                    var textAge = GetTextAge( email.date, options );
                    WriteColor( $"{textAge.text} ago\t\t{email.text}", textAge.color );

                    Console.WriteLine( );
                    DeleteMessages( options, imapClient, emailList );


                    Thread.Sleep( options.ShowConsoleSeconds * 1000 );
                }
            }
            catch( Exception e )
            {
                Console.WriteLine( );
                if( options is null || options.OutputFullException )
                {
                    Console.BufferHeight = defaultBufferHeight;

                    WriteColor( e.ToString(), ConsoleColor.DarkRed );
                }
                else
                {
                    WriteColor( e.Message, ConsoleColor.DarkRed );
                }

                WriteColor( "Press enter to quit" );
                Console.ReadLine( );
            }
        }

        private static ImapClient ImapConnect( SearchOptions options )
        {
            if( !options.TestMode.TestModeSetting( options.TestMode.ImapConnect, out var enabled ) && enabled )
            {
                WriteColor( "[ Test Mode ImapConnect Disabled ]", ConsoleColor.Yellow );
                return null;
            }

            ComponentInfo.SetLicense( "FREE-LIMITED-KEY" );

            var sec = Enum.TryParse( typeof( ConnectionSecurity ), options.ConnectionSecurity, true, out var cs ) ? (ConnectionSecurity)cs : ConnectionSecurity.Auto;

            var imapClient = new ImapClient( options.Host, options.Port, sec, new RemoteCertificateValidationCallback(ValidateServerCertificate) );
            WriteColor( $"Connecting to {options.Host}:{options.Port} using {sec}...", false );
            

            imapClient.Connect( );
            WriteColor( " Success", ConsoleColor.DarkGreen );

            WriteColor( $"Authenticating {options.Username}...", false );
            imapClient.Authenticate( options.Username, options.Password );
            WriteColor( " Success", ConsoleColor.DarkGreen );

            return imapClient;
        }

        private static bool ValidateServerCertificate(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            const SslPolicyErrors ignoredErrors =
                SslPolicyErrors.RemoteCertificateChainErrors |  // self-signed
                SslPolicyErrors.RemoteCertificateNameMismatch;  // name mismatch
            
            if ((sslPolicyErrors & ~ignoredErrors) == SslPolicyErrors.None)
            {
                return true;
            }

            Console.WriteLine("Certificate error: {0}", sslPolicyErrors);

            return false;
        }

        private static List<(DateTime date, string subject, int messageId)> GetEmailIdList( SearchOptions options, ImapClient imapClient )
        {
            var emails = new List<(DateTime date, string subject, int messageId)>( );

            if( !options.TestMode.TestModeSetting( options.TestMode.ImapConnect, out var enabled ) && enabled )
            {
                WriteColor( "[ Test Mode creating fake emails ]", ConsoleColor.Yellow );
                emails.AddRange(
                    new List<(DateTime date, string subject, int messageId)>
                    {
                        (DateTime.Now.AddSeconds(-120), $"{options.SearchSubject} ABC123", 1000),
                        (DateTime.Now.AddSeconds(-320), $"{options.SearchSubject} MMM456", 2000),
                        (DateTime.Now.AddSeconds(-720), $"{options.SearchSubject} ZZZ999", 3000),
                    } );
                return emails;
            }

            imapClient.SelectInbox( );

            WriteColor( $"Search for \"{options.SearchSubject}\"...", false );

            var subjectCmd = $"SUBJECT \"{options.SearchSubject}\"";
            var fromCmd = $"FROM {options.SearchFrom}";

            var command = $"{fromCmd} {subjectCmd}";
            var list = imapClient.SearchMessageNumbers( command );

            if( !list.Any( ) )
                return emails;

            WriteColor( " Success", ConsoleColor.DarkGreen );

            WriteColor( $"Downloading headers...", false );
            foreach( var msgId in list.OrderByDescending( i => i ).Take( options.SearchNumberOfMessages ) )
            {
                var headers = imapClient.GetHeaders( msgId );

                if( DateTime.TryParse( headers[HeaderId.Date].Body, out var date ) )
                    emails.Add( (date, headers[Headers.Subject].Body, msgId) );
            }
            WriteColor( " Success", ConsoleColor.DarkGreen );

            return emails;
        }

        private static (string text, DateTime date) GetTextFromEmail( List<(DateTime date, string subject, int messageId)> emails, string subject )
        {
            WriteColor( $"Finding newest email...", false );


            var filteredEmail = emails
                .OrderByDescending( e => e.date )
                .FirstOrDefault( );

            var text = filteredEmail.subject?
                .Replace( subject, "", StringComparison.InvariantCultureIgnoreCase )
                .Trim( );

            WriteColor( " Success", ConsoleColor.DarkGreen );

            return (text, filteredEmail.date);
        }

        private static void DeleteMessages( SearchOptions options, ImapClient imapClient, List<(DateTime date, string subject, int messageId)> emails )
        {
            var testModeEnabled = options.TestMode.Enabled;

            if( options.DeleteMessages )
            {
                if( !options.TestMode.TestModeSetting( options.TestMode.DeleteMessages ) && testModeEnabled )
                {
                    WriteColor( "[ Test Mode not deleting messages ]", ConsoleColor.Yellow );
                    return;
                }

                if( !options.TestMode.TestModeSetting( options.TestMode.ImapConnect ) && testModeEnabled )
                {
                    WriteColor( "[ Test Mode pretend to delete fake messages ]", ConsoleColor.Yellow );
                }

                WriteColor( $"Deleting {emails.Count} emails..." );
                emails.ForEach( e =>
                {
                    WriteColor( $"   Deleting message from {e.date:G}...", false );
                    if( !testModeEnabled || options.TestMode.TestModeSetting( options.TestMode.ImapConnect ) && options.TestMode.TestModeSetting( options.TestMode.DeleteMessages ) )
                    {
                        imapClient.DeleteMessage( e.messageId, false );
                    }
                    WriteColor( " Success", ConsoleColor.DarkGreen );
                } );
            }
        }

        private static void WriteColor( string text, bool newLine = true )
        {
            WriteColor( text, Console.ForegroundColor, newLine );
        }

        private static void WriteColor( string text, ConsoleColor color, bool newLine = true )
        {
            var formattedText = newLine ? $"{text}{Environment.NewLine}" : $"{text,-( WIDTH - 20 )}";
            Console.ForegroundColor = color;
            Console.Write( $"{formattedText}" );
            Console.ResetColor( );
            Console.BackgroundColor = BACKGROUNDCOLOR;
        }

        private static (string text, ConsoleColor color) GetTextAge( DateTime date, SearchOptions options )
        {
            var now = DateTime.Now;
            var age = now - date;

            var color = age.TotalSeconds > options.EmailMaxValidAgeSeconds ? ConsoleColor.Yellow : ConsoleColor.White;

            if( age > new TimeSpan( 1, 0, 0 ) )
            {
                return ($"{( int ) age.TotalMinutes}m", ConsoleColor.DarkRed);
            }

            if( age > TimeSpan.FromSeconds( options.EmailMaxValidAgeSeconds ) )
            {
                return ($"{( int ) age.TotalMinutes}m {( int ) age.Seconds}s", color);
            }

            return ($"{( int ) age.TotalSeconds}s", color);
        }

        private static void ConfigureServices( IServiceCollection serviceCollection )
        {

            _configuration = new ConfigurationBuilder( )
                .SetBasePath( Directory.GetParent( AppContext.BaseDirectory ).FullName )
                .AddJsonFile( "appsettings.json", false )
                .AddEnvironmentVariables( )
                .Build( );

            serviceCollection.AddSingleton( _configuration );
        }
    }
}
