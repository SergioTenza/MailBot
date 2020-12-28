using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using MailKit.Net.Pop3;

namespace MailBot
{
    class Program
    {
        private static string userName = Environment.GetEnvironmentVariable("USER_TEST_EMAIL", EnvironmentVariableTarget.User);

        private static string userPass = Environment.GetEnvironmentVariable("USER_TEST_PASSWORD", EnvironmentVariableTarget.User);

        private static string mailImapServer = "imap.gmail.com";

        private static Int32 mailImapPort = 993;

        private static string mailPopServer = "pop.gmail.com";

        private static Int32 mailPopPort = 995;
        static void Main(string[] args)
        {
            Console.WriteLine("Bienvenidos a MailBot");
            Options();
            Console.WriteLine("Iniciando conexion\nUsuario: {0}\nServidor: {1} Puerto: {2}", userName, mailImapServer, mailImapPort);
            try
            {
                using (var client = new ImapClient())
                {
                    client.Connect(mailImapServer, mailImapPort, true);

                    client.Authenticate(userName, userPass);

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
                Console.WriteLine("Ha ocurrido un error.\nPulsa 'E' si quieres ver el error\nPulsa cualquier otra tecla para cerrar la aplicacion ");
                if (Console.ReadKey().Key == ConsoleKey.E)
                {
                    Console.WriteLine(Error);
                    Console.WriteLine("Pulsa cualquier tecla para salir de la aplicacion.");
                    Console.ReadKey();
                }
            }

            //            Get

            //string getEnv = Environment.GetEnvironmentVariable("envVar");
            //            Set

            //string setEnv = Environment.SetEnvironmentVariable("envvar", varEnv);
        }
        private static void Options()
        {
            Console.WriteLine("\nSelecciona tu opcion:\n1 POP\n2 IMAP");
            if (Console.ReadKey().Key == ConsoleKey.D1)
            {
                try
                {
                    using (var client = new Pop3Client())
                    {
                        client.Connect(mailPopServer, mailPopPort, true);

                        client.Authenticate(userName, userPass);

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
                        client.Connect(mailImapServer, mailImapPort, true);

                        client.Authenticate(userName, userPass);

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
