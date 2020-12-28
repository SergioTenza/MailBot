﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using MailKit.Net.Pop3;
using static MailBot.Program;
using System.Text.RegularExpressions;

namespace MailBot
{
    class Program
    {
        public static class Connection
        {
            public static string userName = Environment.GetEnvironmentVariable("USER_TEST_EMAIL", EnvironmentVariableTarget.User);
            public static string userPass = Environment.GetEnvironmentVariable("USER_TEST_PASSWORD", EnvironmentVariableTarget.User);
            public static string mailImapServer = "imap.gmail.com";
            public static string mailPopServer = "pop.gmail.com";
            public static Int32 mailImapPort = 993;
            public static Int32 mailPopPort = 995;            
        }         
        static void Main(string[] args)
        {
            
            Console.WriteLine("Bienvenidos a MailBot");
            Options();
            //            Get

            //string getEnv = Environment.GetEnvironmentVariable("envVar");
            //            Set

            //string setEnv = Environment.SetEnvironmentVariable("envvar", varEnv);
        }
        private static void Options()
        {
            Console.WriteLine("\nSelecciona tu opcion:\n1 POP\n2 IMAP\n3 Cambiar Usuario");
            if (Console.ReadKey().Key == ConsoleKey.D1)
            {
                try
                {
                    using (var client = new Pop3Client())
                    {
                        client.Connect(Connection.mailPopServer, Connection.mailPopPort, true);

                        client.Authenticate(Connection.userName, Connection.userPass);

                        for (int i = 0; i < client.Count; i++)
                        {
                            var message = client.GetMessage(i);
                            Console.WriteLine("Subject: {0}", message.Subject);
                        }

                        client.Disconnect(true);
                    }
                }
                catch (Exception Error)
                {

                    Console.WriteLine("Ha ocurrido un error.\nPulsa 'E' si quieres ver el error.\nPulsa cualquier tecla para continuar...");
                    if (Console.ReadKey().Key == ConsoleKey.E)
                    {
                        Console.WriteLine(Error);
                        Options();
                    }
                    else
                    {
                        Options();
                    }
                }

            }
            else if (Console.ReadKey().Key == ConsoleKey.D2)
            {
                try
                {
                    using (var client = new ImapClient())
                    {
                        client.Connect(Connection.mailImapServer, Connection.mailImapPort, true);

                        client.Authenticate(Connection.userName, Connection.userPass);

                        // The Inbox folder is always available on all IMAP servers...
                        var inbox = client.Inbox;
                        inbox.Open(FolderAccess.ReadOnly);

                        Console.WriteLine("Total messages: {0}", inbox.Count);
                        Console.WriteLine("Recent messages: {0}", inbox.Recent);

                        for (int i = 0; i < inbox.Count; i++)
                        {
                            var message = inbox.GetMessage(i);
                            Console.WriteLine("Subject: {0}", message.Subject);
                        }

                        client.Disconnect(true);
                    }
                }
                catch (Exception Error)
                {
                    Console.WriteLine("Ha ocurrido un error.\nPulsa 'E' si quieres ver el error.\nPulsa cualquier tecla para continuar...");
                    if (Console.ReadKey().Key == ConsoleKey.E)
                    {
                        Console.WriteLine(Error);
                        Options();
                    }
                    else
                    {
                        Options();
                    }
                }

            }
            else if (Console.ReadKey().Key == ConsoleKey.D3)
            {
                string emailString = Console.ReadLine();
                bool isEmail = Regex.IsMatch(emailString, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase);
                if (isEmail)
                {
                    Connection.userName = emailString;
                    Console.WriteLine("Introduce tu password: ");
                    Connection.userPass = Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("El email no es valido. Pruebe de neuvo.");
                }
            }
            else
            {
                Console.WriteLine("Comando no reconocido\nPulsa 'R' para reintentarlo o cualquier tecla para salir de la aplicacion.");
                if (Console.ReadKey().Key == ConsoleKey.R)
                {
                    Options();
                }
                else
                {
                    Environment.Exit(0);
                }
            }
        }
    }
}
