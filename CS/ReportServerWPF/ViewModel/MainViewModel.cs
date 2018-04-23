using System;
using System.IO;
using System.Threading.Tasks;
using DevExpress.DocumentServices.ServiceModel;
using DevExpress.DocumentServices.ServiceModel.DataContracts;
using DevExpress.Mvvm;
using DevExpress.Mvvm.POCO;
using DevExpress.ReportServer.ServiceModel.Client;
using DevExpress.ReportServer.ServiceModel.ConnectionProviders;
using DevExpress.ReportServer.ServiceModel.DataContracts;
using DevExpress.XtraPrinting;
using System.Printing;
using System.Windows.Controls;

namespace ReportServerWPF.ViewModel {
    public class MainViewModel {
        const string ServerAddress = "https://reportserver.devexpress.com";
        readonly ConnectionProvider serverConnection = new GuestConnectionProvider(ServerAddress);
        readonly int ReportId = 1113; //Customer Order History report
        readonly ReportParameter[] reportParameters = new ReportParameter[] {
            new ReportParameter {
                Path = "@CustomerID", 
                Value = "BERGS",
                MultiValue = false
            }};


        public virtual bool IsBusy { get; protected set; }
        protected void OnIsBusyChanged() {
            this.RaiseCanExecuteChanged(x => x.Print());
            this.RaiseCanExecuteChanged(x => x.Export());
        }
        
        protected virtual IMessageBoxService MessageBoxService { get { return null; } }
        protected virtual ISaveFileDialogService SaveFileDialogService { get { return null; } }

        public bool CanExport() {
            return !IsBusy;
        }

        public void Export() {
            SaveFileDialogService.Filter = "PDF files (*.pdf)|*.pdf";
            if(SaveFileDialogService.ShowDialog()) {
                IsBusy = true;
                ExportTo(serverConnection, ReportId, new PdfExportOptions(), reportParameters, x => File.WriteAllBytes(SaveFileDialogService.GetFullFileName(), x));
            }            
        }

        public bool CanPrint() {
            return !IsBusy;
        }

        public void Print() {
            ExportTo(serverConnection, ReportId, new XpsExportOptions(), reportParameters, x => PrintReportFromXps(x));           
        }

        void PrintReportFromXps(byte[] xps) {
            using(PrintSystemJobInfo jobInfo = new PrintDialog().PrintQueue.AddJob("Print Job Name")) {
                jobInfo.JobStream.Write(xps, 0, xps.Length);
            }
        }

        void ExportTo(ConnectionProvider serverConnection, int reportId, ExportOptionsBase exportOptions, ReportParameter[] parameters, Action<byte[]> action) {
            IsBusy = true;
            serverConnection
                .ConnectAsync()
                .ContinueWith(t => {
                    IReportServerClient client = t.Result;
                    return Task.Factory.ExportReportAsync(client, new ReportIdentity(reportId), exportOptions, parameters, null);
                }).Unwrap()
                .ContinueWith(t => {
                    IsBusy = false;
                    try {
                        if(t.IsFaulted) {
                            throw new Exception(t.Exception.Flatten().InnerException.Message);
                        }
                        action(t.Result);
                    } catch(Exception e) {
                        MessageBoxService.Show(e.Message, "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}
