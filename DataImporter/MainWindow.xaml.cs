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
            RedirectUri = new Uri("app://5d3e90d6-aa8e-48a8-8f2c-58b45cc67315")
        };

        private BackgroundWorker _worker = new BackgroundWorker();

        private ImportCrmDataHandler _importCrmDataHandler;

        public MainWindow()
        {
            InitializeComponent();
            CrmLoginCtrl.Visibility = Visibility.Visible;
            CrmLoginCtrl.SetGlobalStoreAccess(_connectionManager);

            _connectionManager.ConnectionCheckComplete += ConnectionManagerOnConnectionCheckComplete;
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
        }

        private void ImportCrmDataHandlerOnUserMappingRequired(object sender, UserMapRequiredEventArgs e)
        {
        }

        private void ImportCrmDataHandlerOnUpdateProgressItem(object sender, ProgressItemEventArgs e)
        {
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
                var dataPath = @"C:\work\CRMDataFolder\data.zip";
                ImportCrmDataHandler.CrackZipFileAndCheckContents(dataPath, null, out var workingImportFolder);

                _importCrmDataHandler.ValidateSchemaFile(workingImportFolder);
                bool crm = _importCrmDataHandler.ImportDataToCrm(workingImportFolder, false);
                e.Result = (object)crm;
            }
            catch (Exception ex)
            {

            }
        }
    }
}
