using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace DataImporter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static List<string> DataFiles { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            if (e.Args.Length < 2)
            {
                Console.WriteLine(@"The data files are not passed to the command.");
                Console.WriteLine(@"Example:");
                Console.WriteLine(@"DataImporter.exe -datafiles ""c:\data1.zip"" ""c:\data2.zip""");
            }

            DataFiles = new List<string>();
            for (var i = 1; i < e.Args.Length; i++)
            {
                if (File.Exists(e.Args[i]))
                {
                    DataFiles.Add(e.Args[i]);
                }
                else
                {
                    Console.WriteLine($@"The file {e.Args[i]} does not exist.");
                }
            }

            base.OnStartup(e);
        }
    }
}
