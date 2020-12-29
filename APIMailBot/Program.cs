using ExcelDataReader;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Json;

namespace APIMailBot
{
    class Program
    {
        private static readonly HttpClient clientHttp = new HttpClient();
        private static string url = "https://httpbin.org/post";

        public static string userName = Environment.GetEnvironmentVariable("USER_TEST_EMAIL", EnvironmentVariableTarget.User);
        public static string userPass = Environment.GetEnvironmentVariable("USER_TEST_PASSWORD", EnvironmentVariableTarget.User);
        public static string mailImapGmail = "imap.gmail.com";
        public static string mailPopGmail = "pop.gmail.com";
        public static string mailImapIonos = "imap.ionos.es";
        public static string mailPopIonos = "pop.ionos.es";
        public static Int32 mailImapPort = 993;
        public static Int32 mailPopPort = 995;
        public static Int32 SelectionIndex;
        public static string basePath = Directory.GetCurrentDirectory();
        //Get string getEnv = Environment.GetEnvironmentVariable("envVar");
        //Set string setEnv = Environment.SetEnvironmentVariable("envvar", varEnv);
        public class Mesa
        {
            public string name { get; set; }
            public List<string> participants { get; set; }
        }

        public class TareaMaquina
        {
            public string idMaquina { get; set; }
            public List<Mesa> mesas { get; set; }
        }

        private static async Task Main(string[] args)
        {
            // Initial Setup
            handler = new ConsoleEventDelegate(ConsoleEventCallback);
            SetConsoleCtrlHandler(handler, true);
            // Main Loop           
            await DoMenuAsync();
        }
        static async Task DoMenuAsync()
        {            
            uint selectionIndex = 0;

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Bienvenidos a APIMailBot\n");
            Console.WriteLine("Developed by Sergio Tenza");
            Console.WriteLine("Copyright (c) TNZ Servicios Informaticos. All Rights Reserved.\n");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("==============================================================\n");

            while (selectionIndex != 99)
            {
                displaySelection();
                
                bool resultSelection = int.TryParse(Console.ReadLine(), out SelectionIndex);

                if (resultSelection)
                {
                    switch (SelectionIndex)
                    {
                        case 1:
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Usuarios en el sistema\n");
                            Console.WriteLine("Usuario: {0}", userName);                           
                            Console.ForegroundColor = ConsoleColor.Gray;
                            break;
                        case 2:
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Usuarios en el sistema\n");
                            Console.WriteLine("Usuario: {0}\nPassword: {1}", userName, userPass);
                            Console.ForegroundColor = ConsoleColor.Gray;
                            break;
                        case 3:
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Introduce tu direccion de Email de usuario:> ");
                            Console.ForegroundColor = ConsoleColor.Gray;
                            string emailString = Console.ReadLine();
                            bool isEmail = Regex.IsMatch(emailString, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase);
                            if (isEmail)
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("Introduce tu password:> ");
                                Console.ForegroundColor = ConsoleColor.Gray;
                                string tmpPass = Console.ReadLine();
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                Console.WriteLine(
                                    "\nLos datos introducidos son:\n\n" +
                                    "Usuario: {0}\n" +
                                    "Password: {1}\n\n" +
                                    "PULSE 'S' para guardar los credenciales o 'N' para descartar cambios\n", emailString, tmpPass);
                                if (Console.ReadKey().Key == ConsoleKey.S)
                                {
                                    userName = emailString;
                                    userPass = tmpPass;
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("\nLOS CAMBIOS HAN SIDO GUARDADOS EN EL SISTEMA.");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("\nTODOS LOS CAMBIOS DESCARTADOS.");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }                                
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("El email no es valido. Pruebe de nuevo.");
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }
                            break;
                        case 4:
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Introduce el nuevo endpoint:> ");
                            Console.ForegroundColor = ConsoleColor.Gray;
                            string endPointString = Console.ReadLine();
                            bool isHttp = Regex.IsMatch(endPointString, @"^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$", RegexOptions.IgnoreCase);
                            if (isHttp)
                            {                                
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                Console.WriteLine(
                                    "\nLos datos introducidos son:\n\n" +
                                    "EndPoint: {0}\n" +
                                    "PULSE 'S' para guardar los credenciales o 'N' para descartar cambios\n", endPointString);
                                if (Console.ReadKey().Key == ConsoleKey.S)
                                {
                                    url = endPointString;                                    
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("\nLOS CAMBIOS HAN SIDO GUARDADOS EN EL SISTEMA.");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("\nTODOS LOS CAMBIOS DESCARTADOS.");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }

                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("El EndPoint no es valido. Pruebe de nuevo.");
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }
                            break;
                        case 10:
                            break;
                        case 11:
                            break;
                        case 12:
                            break;
                        case 13:
                            try
                            {
                                using (var client = new ImapClient())
                                {
                                    client.Connect(mailImapIonos, mailImapPort, true);

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

                                        foreach (MimeEntity attachment in message.Attachments)
                                        {                                            
                                            var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                                            string name = Path.GetFileNameWithoutExtension(fileName);
                                            string ext = Path.GetExtension(fileName);

                                            if (ext == ".xls" || ext == ".xlsx")
                                            {
                                                
                                                Console.ForegroundColor = ConsoleColor.Green;
                                                Console.WriteLine("Adjunto excel: {0}\nName: {1}\nExtension: {2}", fileName, name, ext);
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                using (var stream = File.Create(fileName))
                                                {
                                                    if (attachment is MessagePart)
                                                    {
                                                        var rfc822 = (MessagePart)attachment;

                                                        rfc822.Message.WriteTo(stream);
                                                    }
                                                    else
                                                    {
                                                        var part = (MimePart)attachment;

                                                        part.Content.DecodeTo(stream);
                                                    }
                                                }

                                                var dirName = $@"{basePath}\{fileName}";
                                                
                                                DataTableCollection dataTableCollection;

                                                TareaMaquina _tarea = new TareaMaquina();

                                                List<string> _tables = new List<string>();
                                                List<string> _columns = new List<string>();
                                                List<string> _rows= new List<string>();

                                                using (var stream = File.Open(dirName, FileMode.Open, FileAccess.Read))
                                                {
                                                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                                                    {
                                                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                                        {
                                                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                                                        }
                                                        );
                                                        dataTableCollection = result.Tables;

                                                        _tables.Clear();
                                                        _columns.Clear();
                                                        _rows.Clear();

                                                        _tables.AddRange(dataTableCollection.Cast<DataTable>().Select(t => t.TableName).ToArray<string>());

                                                        DataTable dt;

                                                        foreach (var table in _tables)
                                                        {
                                                            dt = dataTableCollection[table];
                                                            var columNames = (from c in dt.Columns.Cast<DataColumn>() select c.ColumnName).ToArray();
                                                            _columns = new List<string>(columNames);
                                                            
                                                            Console.ForegroundColor = ConsoleColor.Green;
                                                            Console.WriteLine("Nombre Maquina: {0}\n", table);
                                                            _tarea.idMaquina = table;
                                                            Console.ForegroundColor = ConsoleColor.Gray;

                                                            List<Mesa> tmpMesas = new List<Mesa>();

                                                            foreach (var column in _columns)
                                                            {
                                                                Mesa tmpMesa = new Mesa();
                                                                Console.ForegroundColor = ConsoleColor.DarkGreen;
                                                                Console.WriteLine("Nombre Mesa: {0}\n", column);
                                                                tmpMesa.name = column;
                                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                                if (dt != null)
                                                                {
                                                                    List<string> tmpParticipants = new List<string>();
                                                                    foreach (DataRow row in dt.Rows)
                                                                    {
                                                                        string tmp = row.Field<string>(column);
                                                                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                                                                        Console.WriteLine("Asistente: {0}\n", tmp);
                                                                        tmpParticipants.Add(tmp);
                                                                        Console.ForegroundColor = ConsoleColor.Gray;
                                                                    }
                                                                    
                                                                    tmpMesa.participants = tmpParticipants;
                                                                    
                                                                    tmpMesas.Add(tmpMesa);                                                                    
                                                                }
                                                            }
                                                            _tarea.mesas = tmpMesas;

                                                        }
                                                    }
                                                }
                                                var options = new JsonSerializerOptions
                                                {
                                                    WriteIndented = true,
                                                };
                                                var jsonString = JsonSerializer.Serialize(_tarea, options);
                                                
                                                Console.ForegroundColor = ConsoleColor.Green;
                                                Console.WriteLine(jsonString);
                                                Console.ForegroundColor = ConsoleColor.Gray;


                                                var petition = await PostJsonContenturl(url, clientHttp, jsonString);

                                                var jsonStringResponse = JsonSerializer.Serialize(petition, options);

                                                Console.ForegroundColor = ConsoleColor.Green;
                                                Console.WriteLine(jsonStringResponse);
                                                Console.ForegroundColor = ConsoleColor.Gray;


                                            }
                                            else 
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine("Adjunto NO VALIDO");
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                            }                                            
                                        }
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
                                }                                
                            }
                            break;
                        case 98:
                            displayOptions();                            
                            break;
                        case 99:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.Write("Aplicacion Deteniendose en 5 Segundos..\n");
                            Thread.Sleep(1000);
                            Console.Write("Aplicacion Deteniendose en 4 Segundos...\n");
                            Thread.Sleep(1000);
                            Console.Write("Aplicacion Deteniendose en 3 Segundos...\n");
                            Thread.Sleep(1000);
                            Console.Write("Aplicacion Deteniendose en 2 Segundos...\n");
                            Thread.Sleep(1000);
                            Console.Write("Aplicacion Deteniendose en 1 Segundos...\n");
                            Thread.Sleep(1000);
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Hecho!");
                            Thread.Sleep(2000);
                            Environment.Exit(0);
                            break;
                        default:
                            break;
                    }
                }
            }            
        }

        private static async Task<TestResponse> PostJsonContenturl(string url, HttpClient clientHttp, string jsonString)
        {
            var postRequest = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = JsonContent.Create(jsonString)
            };

            var postResponse = await clientHttp.SendAsync(postRequest);

            postResponse.EnsureSuccessStatusCode();

            Console.WriteLine($"{(postResponse.IsSuccessStatusCode ? "Success" : "Error")} - {postResponse.StatusCode}");
            
            if (postResponse.IsSuccessStatusCode)
            {
                // perhaps check some headers before deserialising

                try
                {
                    return await postResponse.Content.ReadFromJsonAsync<TestResponse>();
                }
                catch (NotSupportedException) // When content type is not valid
                {
                    Console.WriteLine("The content type is not supported.");
                }
                catch (JsonException) // Invalid JSON
                {
                    Console.WriteLine("Invalid JSON.");
                }
            }

            return null;
        }

        static private void displaySelection()
        {
            Console.WriteLine("\n");
            Console.WriteLine("\nType your selection Number. Type '98 for OPTIONS'");
            Console.Write("> ");
        }
        static private void displayOptions()
        {
           
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("1.  Mostrar usuarios en el sistema");
            Console.WriteLine("2.  Mostrar usuarios + password en el sistema");
            Console.WriteLine("3.  Agregar usuario + password en el sistema");
            Console.WriteLine("4.  Agregar nuevo ENDPOINT en el sistema");
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
        static private void Server(string type, string service)
        {
            Tuple<string, int> _command;  

            switch (type)
            {
                case "POP":
                    switch (service)
                    {
                        case "gmail":
                             _command = Tuple.Create(mailPopGmail, mailPopPort);
                            break;
                        case "ionos":
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
                case "IMAP":
                    switch (service)
                    {
                        case "gmail":
                             _command = Tuple.Create(mailImapGmail, mailImapPort);
                            break;
                        case "ionos":
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
        }
        
        #region Console Handler
        static bool ConsoleEventCallback(int eventType)
        {
            switch (eventType)
            {
                //Ctrl+C
                case 0:
                    Environment.Exit(0);
                    break;
                //Close
                case 2:
                    // Stop the server 
                    Console.Write("Aplicacion Deteniendose...");                    
                    Console.WriteLine("Hecho!");
                    Environment.Exit(0);
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
