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
using System.Text.RegularExpressions;

namespace MailBot
{
    class Program
    {
        public static class Connection
        {
            public static string userName = Environment.GetEnvironmentVariable("USER_TEST_EMAIL", EnvironmentVariableTarget.User);
            public static string userPass = Environment.GetEnvironmentVariable("USER_TEST_PASSWORD", EnvironmentVariableTarget.User);
            public static string mailImapGmail = "imap.gmail.com";
            public static string mailPopGmail = "pop.gmail.com";
            public static string mailImapIonos = "imap.ionos.es";
            public static string mailPopIonos = "pop.ionos.es";
            public static Int32 mailImapPort = 993;
            public static Int32 mailPopPort = 995;
             //Get string getEnv = Environment.GetEnvironmentVariable("envVar");
            //Set string setEnv = Environment.SetEnvironmentVariable("envvar", varEnv);
        }         
        static void Main(string[] args)
        {
            
            Console.WriteLine("Bienvenidos a MailBot");
            Domenu();
           
        }
        static void DoMenu()
        {
	        DisplaySelection();

	        SelectionIndex = 0;

	        bool resultSelection = int.TryParse(CommandIndex, out SelectionIndex);

	        if(resultSelection)
	        {
		        while(SelectionIndex != 99)
		        {
			        switch(SelectionIndex)
				        {
                            case 10:
                                break;
                            case 11:
                                break;
                            case 12:
                                break;
                            case 13:
                                break;
					        case 98:
					        displayOptions();
                                break;
					        default:
					            break;
				        }
		        }
	        }

        }
        static private void displaySelection()
        {           
            Console.WriteLine("\n");
            Console.WriteLine("\nType your selection Number. Type '98 for OPTIONS'");
            Console.Write("> ");
        }
        static private void displayOptions()
        {
            Console.WriteLine("\n");
            Console.WriteLine("OSC Camera Wireless Controller .NET Core Application");
            Console.WriteLine("Developed by Sergio Tenza\n");
            Console.WriteLine("Copyright (c) TNZ Servicios Informaticos. All Rights Reserved.");
            Console.WriteLine("==============================================================\n");            
            Console.ForegroundColor = ConsoleColor.Green;     
            Console.WriteLine("10.  POP Gmail Server");
            Console.WriteLine("11.  IMAP Gmail Server");
            Console.WriteLine("12.  POP Ionos Server");
            Console.WriteLine("13.  IMAP Ionos Server");
            Console.WriteLine("98.  Show Application Options");
            Console.WriteLine("99.  Exit Application");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\n");
            Console.Write("> ");
        }
        static private void Server(string type,string service)
        {            
            var _command; 

            switch (type)
	        {
                    case POP:
                        switch (service)
	                    {
                                case gmail:
                                    _command = Tuple.Create(mailPopGmail, mailPopPort);
                                    break;
                                case ionos:
                                    _command = Tuple.Create(mailPopIonos, mailPopPort);
                                    break;
		                        default:
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("==============================================================\n"); 
                                    Console.WriteLine("Ha ocurrido un error\n");
                                    Console.WriteLine("==============================================================\n");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                    displaySelection();
                                    break;
	                    }
                        break;
                    case IMAP: 
                        switch (service)
	                    {
                                case gmail:
                                    _command = Tuple.Create(mailImapGmail, mailImapPort);
                                    break;
                                case ionos:
                                    _command = Tuple.Create(mailImapIonos, mailImapPort);
                                    break;                                
		                        default:
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("==============================================================\n"); 
                                    Console.WriteLine("Ha ocurrido un error\n");
                                    Console.WriteLine("==============================================================\n");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                    displaySelection();
                                    break;
	                    }
                        break;
		            default:
                        break;
	        }
            //bool checkCommand = _command.Item1.Contains("");
        }
       
        #region Console Handler
        static bool ConsoleEventCallback(int eventType)
        {
            switch (eventType)
            {
                //Ctrl+C
                case 0:                    
                //Close
                case 2:
                    // Stop the server 
                    Console.Write("Server stopping...");
                    server.Stop();
                    Console.WriteLine("Done!");
                    break;
            }
            return false;
        }

        static ConsoleEventDelegate handler;

        private delegate bool ConsoleEventDelegate(int eventType);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);
        #endregion
    }
}
