using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using GemBox.Email;
using GemBox.Email.Imap;
using GemBox.Email.Mime;
using GemBox.Email.Security;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Clipboard = TextCopy.Clipboard;

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
                var email = GetTextFromEmail( emailList, options.SearchSubject ) ?? (text: null, date: DateTime.MinValue);

                if( email.text != null )
                {
                    Console.Write( $"Copying password to clipboard..." );
                    var c = new Clipboard( );
                    c.SetText( email.text );
                    WriteColor( " Success", ConsoleColor.DarkGreen );

                    try
                    {
                        FindWindow( "vpnui" );
                    }
                    catch( Exception e )
                    {
                        Console.WriteLine( e );
                    }

                    Console.WriteLine( );
                    var textAge = GetTextAge( email.date );
                    WriteColor( $"{textAge.text} ago\t\t{email.text}", textAge.color );

                }

                Console.WriteLine( );
                DeleteMessages( options, imapClient, emailList );

                Thread.Sleep( options.ShowConsoleSeconds * 1000 );
            }
        }

        private static ImapClient ImapConnect( string host, int port, string username, string password )
        {
            try
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
            catch( Exception e )
            {
                WriteColor( " Failed.", ConsoleColor.DarkRed );
                WriteColor( $"{e.Message}", ConsoleColor.White );
                throw;
            }
        }

        private static List<(DateTime date, string subject, int messageId)> GetEmailIdList( ImapClient imapClient, string from, string subject, int searchNumber )
        {
            try
            {
                var emails = new List<(DateTime date, string subject, int messageId)>( );
                imapClient.SelectInbox( );

                Console.Write( $"Search for \"{subject}\"..." );

                var subjectCmd = $"SUBJECT \"{subject}\"";
                var fromCmd = $"FROM {from}";

                var command = $"{fromCmd} {subjectCmd}";
                var list = imapClient.SearchMessageNumbers( command );
                WriteColor( " Success", ConsoleColor.DarkGreen );

                if( !list.Any( ) )
                {
                    WriteColor( " No emails found.", ConsoleColor.DarkYellow );
                }
                else
                {
                    Console.Write( $"Downloading headers..." );
                    foreach( var msgId in list.OrderByDescending( i => i ).Take( searchNumber ) )
                    {
                        var headers = imapClient.GetHeaders( msgId );

                        if( DateTime.TryParse( headers[HeaderId.Date].Body, out var date ) )
                            emails.Add( (date, headers[Headers.Subject].Body, msgId) );
                    }
                    WriteColor( " Success.", ConsoleColor.DarkGreen );
                }

                return emails;
            }
            catch( Exception e )
            {
                WriteColor( " Failed.", ConsoleColor.DarkRed );
                WriteColor( $"{e.Message}", ConsoleColor.White );
                throw;
            }
        }

        private static (string text, DateTime date)? GetTextFromEmail( List<(DateTime date, string subject, int messageId)> emails, string subject )
        {
            try
            {
                if( emails.Any( ) )
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

                return null;
            }
            catch( Exception e )
            {
                WriteColor( " Failed.", ConsoleColor.DarkRed );
                WriteColor( $"{e.Message}", ConsoleColor.White );
                throw;
            }
        }

        private static void DeleteMessages( SearchOptions options, ImapClient imapClient, List<(DateTime date, string subject, int messageId)> emails )
        {
            if( options.DeleteMessages )
            {
                if( emails.Any( ) )
                {
                    try
                    {
                        Console.WriteLine( $"Deleting {emails.Count} emails..." );
                        emails.ForEach( e =>
                        {
                            Console.Write( $"\tDeleting message from {e.date:G}..." );
                            imapClient.DeleteMessage( e.messageId, false );
                            WriteColor( " Success", ConsoleColor.DarkGreen );
                        } );
                        WriteColor( " Success", ConsoleColor.DarkGreen );
                    }
                    catch( Exception e )
                    {
                        WriteColor( " Failed.", ConsoleColor.DarkRed );
                        WriteColor( $"{e.Message}", ConsoleColor.White );
                        throw;
                    }
                }
            }
        }

        private static void FindWindow( string windowName )
        {
            /*
                vpnui
                Cisco AnyConnect Secure Mobility Client
                6308
                C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe
                vpnui.exe
            */
            var process = Process.GetProcesses( ).Where( p => p.ProcessName == windowName ).FirstOrDefault( );

            try
            {
                WriteProcess(process);

                if( process != null )
                {
                    var subHandles = new WindowHandleInfo( process.MainWindowHandle ).GetAllChildHandles( );
                    
                    Console.WriteLine( subHandles.Count( ) );

                    foreach( var subHandle in subHandles )
                    {
                        Console.WriteLine("...");
                        SetForegroundWindow( subHandle );
                        SendKeys.SendWait("^v");
                        SendKeys.SendWait( "{Enter}" );
                        Console.WriteLine("---");
                    }
                }
            }
            catch( Exception e )
            {
                Console.WriteLine( e.Message );
            }

            
            

            //var prc = Process.GetProcessesByName("notepad");
            //if( prc.Length > 0 )
            //{
            //    SetForegroundWindow( prc[0].MainWindowHandle );
            //}
        }

        [DllImport( "user32.dll" )]
        private static extern bool SetForegroundWindow( IntPtr hWnd );

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        //public static IEnumerable<Process> GetChildProcesses(Process process)
        //{
        //    List<Process> children = new List<Process>();
        //    var mos = new ManagementObjectSearcher(String.Format("Select * From Win32_Process Where ParentProcessID={0}", process.Id));

        //    foreach (ManagementObject mo in mos.Get())
        //    {
        //        children.Add(Process.GetProcessById(Convert.ToInt32(mo["ProcessID"])));
        //    }

        //    return children;
        //}

        public static void WriteProcess( Process process )
        {
            WriteColor( process.ProcessName, ConsoleColor.White );
            Console.WriteLine( process.MainWindowTitle );
            Console.WriteLine( process.Id );
            Console.WriteLine( process.MainModule.FileName );
            Console.WriteLine( process.MainModule.ModuleName );
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
