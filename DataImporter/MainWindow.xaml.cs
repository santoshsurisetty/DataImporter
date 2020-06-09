using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Tooling.CrmConnectControl;
using Microsoft.Xrm.Tooling.Dmt.DataMigCommon.Utility;
using Microsoft.Xrm.Tooling.Dmt.ImportProcessor.DataInteraction;

namespace DataImporter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private CrmConnectionManager _connectionManager = new CrmConnectionManager
        {
            ClientId = "2ad88395-b77d-4561-9441-d0e40824f9bc",
            RedirectUri = new Uri("app://5d3e90d6-aa8e-48a8-8f2c-58b45cc67315"),
            HostApplicatioNameOveride = "DataImporter", // To allow passing command-line parameters to the application
            UseUserLocalDirectoryForConfigStore = true
        };

        private BackgroundWorker _worker = new BackgroundWorker();

        private ImportCrmDataHandler _importCrmDataHandler;

        public MainWindow()
        {
            InitializeComponent();
            CrmLoginCtrl.Visibility = Visibility.Visible;
            CrmLoginCtrl.SetGlobalStoreAccess(_connectionManager);

            CrmLoginCtrl.ConnectionStatusEvent += CrmLoginCtrlOnConnectionStatusEvent;
            CrmLoginCtrl.UserCancelClicked += CrmLoginCtrlOnUserCancelClicked;
            _connectionManager.ConnectionCheckComplete += ConnectionManagerOnConnectionCheckComplete;
        }

        private void CrmLoginCtrlOnUserCancelClicked(object sender, EventArgs e)
        {
            Console.WriteLine(@"User cancelled login");
        }

        private void CrmLoginCtrlOnConnectionStatusEvent(object sender, ConnectStatusEventArgs e)
        {
            if (e.ConnectSucceeded)
            {
                Console.WriteLine(@"Connected to CRM instance successfully.");
            }
            else
            {
                Console.WriteLine(@"Failed to connect to CRM instance.");
            }
        }

        private void ConnectionManagerOnConnectionCheckComplete(object sender, ServerConnectStatusEventArgs e)
        {
            if (_connectionManager.CrmSvc != null && _connectionManager.CrmSvc.IsReady)
            {
                _importCrmDataHandler = new ImportCrmDataHandler();
                _importCrmDataHandler.AddNewProgressItem += ImportCrmDataHandlerOnAddNewProgressItem;
                _importCrmDataHandler.UpdateProgressItem += ImportCrmDataHandlerOnUpdateProgressItem;
                _importCrmDataHandler.UserMappingRequired += ImportCrmDataHandlerOnUserMappingRequired;

                _importCrmDataHandler.CrmConnection = _connectionManager.CrmSvc;
                _importCrmDataHandler.ImportConnections = new Dictionary<int, CrmServiceClient> { { 1, _connectionManager.CrmSvc } };
                _worker = new BackgroundWorker();
                _worker.DoWork += Worker_DoWork;
                _worker.ProgressChanged += Worker_ProgressChanged;
                _worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
                _worker.RunWorkerAsync();
            }
        }

        private void ImportCrmDataHandlerOnAddNewProgressItem(object sender, ProgressItemEventArgs e)
        {
            Console.WriteLine(@"LOG :" + e.progressItem.ItemText);
        }

        private void ImportCrmDataHandlerOnUserMappingRequired(object sender, UserMapRequiredEventArgs e)
        {
        }

        private void ImportCrmDataHandlerOnUpdateProgressItem(object sender, ProgressItemEventArgs e)
        {
            Console.WriteLine(@"LOG :" + e.progressItem.ItemText);
        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var dataFiles = App.DataFiles;
                var workingImportFolders = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
                foreach (var dataFile in dataFiles)
                {
                    if (ImportCrmDataHandler.CrackZipFileAndCheckContents(dataFile, null, out var workingImportFolder))
                    {
                        if (_importCrmDataHandler.ValidateSchemaFile(workingImportFolder))
                        {
                            workingImportFolders[dataFile] = workingImportFolder;
                        }
                        else
                        {
                            Console.WriteLine($@"Schema file validation failed for {dataFile}");
                            Environment.Exit(-1);
                        }
                    }
                    else
                    {
                        Console.WriteLine($@"Invalid zip for importing data {dataFile}");
                        Environment.Exit(-1);
                    }
                }

                foreach (var importFolder in workingImportFolders)
                {
                    var importCrmResult = _importCrmDataHandler.ImportDataToCrm(importFolder.Value, false);

                    if (!importCrmResult)
                    {
                        Console.WriteLine($@"Failed to import {importFolder.Key}");
                        Environment.Exit(-1);
                    }
                }

                e.Result = (object)true;
                Console.WriteLine(@"Import datafile(s) successfully.");
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine($@"Import datafile(s) failed {ex}");
                Environment.Exit(-1);
            }
        }
    }
}
