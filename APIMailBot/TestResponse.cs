using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APIMailBot
{
    class TestResponse
    {     
            public Args args { get; set; }
            public string data { get; set; }
            public Files files { get; set; }
            public Form form { get; set; }
            public Headers headers { get; set; }
            public object json { get; set; }
            public string origin { get; set; }
            public string url { get; set; }
       

        public class Args
        {
        }

        public class Files
        {
        }

        public class Form
        {
        }

        public class Headers
        {
            public string Accept { get; set; }
            public string AcceptEncoding { get; set; }
            public string AcceptLanguage { get; set; }
            public string ContentLength { get; set; }
            public string ContentType { get; set; }
            public string Host { get; set; }
            public string Origin { get; set; }
            public string Referer { get; set; }
            public string SecFetchDest { get; set; }
            public string SecFetchMode { get; set; }
            public string SecFetchSite { get; set; }
            public string UserAgent { get; set; }
            public string XAmznTraceId { get; set; }
        }

    }
}
