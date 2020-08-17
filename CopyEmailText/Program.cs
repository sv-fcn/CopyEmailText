using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using GemBox.Email;
using GemBox.Email.Imap;
using GemBox.Email.Mime;
using GemBox.Email.Security;
using TextCopy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace CopyEmailText
{
    class Program
    {
        private static IConfigurationRoot _configuration;
        static void Main( string[] args )
        {
            Console.Write( $"Reading appsettings..." );
            var serviceCollection = new ServiceCollection( );
            ConfigureServices( serviceCollection );
            var options = _configuration.GetSection( "Options" ).Get<SearchOptions>( );
            WriteColor( " Success", ConsoleColor.DarkGreen );

            using( var imapClient = ImapConnect( options.Host, options.Port, options.Username, options.Password ) )
            {
                var emailList = GetEmailIdList( imapClient, options.SearchFrom, options.SearchSubject, options.SearchNumberOfMessages );
                var email = GetTextFromEmail( emailList, options.SearchSubject );

                Console.Write( $"Copying password to clipboard..." );
                var c = new Clipboard( );
                c.SetText( email.text );
                WriteColor( " Success", ConsoleColor.DarkGreen );

                DeleteMessages( options, imapClient, emailList );

                Console.WriteLine( );
                var textAge = GetTextAge( email.date );
                WriteColor( $"{textAge.text} ago\t\t{email.text}", textAge.color );
                Thread.Sleep( options.ShowConsoleSeconds * 1000 );
            }
        }

        private static ImapClient ImapConnect( string host, int port, string username, string password )
        {
            ComponentInfo.SetLicense( "FREE-LIMITED-KEY" );

            var imapClient = new ImapClient( host, port, ConnectionSecurity.Auto );
            Console.Write( $"Connecting to {host}:{port}..." );
            imapClient.Connect( );
            WriteColor( " Success.", ConsoleColor.DarkGreen );

            Console.Write( $"Authenticating {username}..." );
            imapClient.Authenticate( username, password );
            WriteColor( " Success.", ConsoleColor.DarkGreen );

            return imapClient;
        }

        private static List<(DateTime date, string subject, int messageId)> GetEmailIdList( ImapClient imapClient, string from, string subject, int searchNumber )
        {
            var emails = new List<(DateTime date, string subject, int messageId)>( );
            imapClient.SelectInbox( );

            Console.Write( $"Search for \"{subject}\"..." );

            var subjectCmd = $"SUBJECT \"{subject}\"";
            var fromCmd = $"FROM {from}";

            var command = $"{fromCmd} {subjectCmd}";
            var list = imapClient.SearchMessageNumbers( command );
            WriteColor( " Success", ConsoleColor.DarkGreen );

            Console.Write( $"Downloading headers..." );
            foreach( var msgId in list.OrderByDescending( i => i ).Take( searchNumber ) )
            {
                var headers = imapClient.GetHeaders( msgId );

                if( DateTime.TryParse( headers[HeaderId.Date].Body, out var date ) )
                    emails.Add( (date, headers[Headers.Subject].Body, msgId) );
            }
            WriteColor( " Success.", ConsoleColor.DarkGreen );

            return emails;
        }

        private static (string text, DateTime date) GetTextFromEmail( List<(DateTime date, string subject, int messageId)> emails, string subject )
        {
            Console.Write( $"Finding newest email..." );


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
            if( options.DeleteMessages )
            {
                Console.Write( $"Deleting {emails.Count} emails..." );
                emails.ForEach( e => imapClient.DeleteMessage( e.messageId, false ) );
                WriteColor( " Success", ConsoleColor.DarkGreen );
            }
        }

        private static void WriteColor( string text, ConsoleColor color )
        {
            Console.ForegroundColor = color;
            Console.WriteLine( text );
            Console.ResetColor( );
        }

        private static (string text, ConsoleColor color) GetTextAge( DateTime date )
        {
            var now = DateTime.Now;
            var age = now - date;

            var color = age.TotalSeconds > 30 ? ConsoleColor.Yellow : ConsoleColor.White;

            if( age > new TimeSpan( 1, 10, 0 ) )
            {
                return ($"{( int ) age.TotalMinutes}m", ConsoleColor.DarkRed);
            }
            
            if( age > new TimeSpan( 0, 2, 0 ) )
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
