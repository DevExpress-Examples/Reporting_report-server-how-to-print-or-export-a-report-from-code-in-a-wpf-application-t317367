Imports System
Imports System.IO
Imports System.Threading.Tasks
Imports DevExpress.DocumentServices.ServiceModel
Imports DevExpress.DocumentServices.ServiceModel.DataContracts
Imports DevExpress.Mvvm
Imports DevExpress.Mvvm.POCO
Imports DevExpress.ReportServer.ServiceModel.Client
Imports DevExpress.ReportServer.ServiceModel.ConnectionProviders
Imports DevExpress.ReportServer.ServiceModel.DataContracts
Imports DevExpress.XtraPrinting
Imports System.Printing
Imports System.Windows.Controls

Namespace ReportServerWPF.ViewModel
    Public Class MainViewModel
        Private Const ServerAddress As String = "https://reportserver.devexpress.com"
        Private ReadOnly serverConnection As ConnectionProvider = New GuestConnectionProvider(ServerAddress)
        Private ReadOnly ReportId As Integer = 1113 'Customer Order History report
        Private ReadOnly reportParameters() As ReportParameter = { _
            New ReportParameter With {.Path = "@CustomerID", .Value = "BERGS", .MultiValue = False} _
        }


        Private privateIsBusy As Boolean
        Public Overridable Property IsBusy() As Boolean
            Get
                Return privateIsBusy
            End Get
            Protected Set(ByVal value As Boolean)
                privateIsBusy = value
            End Set
        End Property
        Protected Sub OnIsBusyChanged()
            Me.RaiseCanExecuteChanged(Sub(x) x.Print())
            Me.RaiseCanExecuteChanged(Sub(x) x.Export())
        End Sub

        Protected Overridable ReadOnly Property MessageBoxService() As IMessageBoxService
            Get
                Return Nothing
            End Get
        End Property
        Protected Overridable ReadOnly Property SaveFileDialogService() As ISaveFileDialogService
            Get
                Return Nothing
            End Get
        End Property

        Public Function CanExport() As Boolean
            Return Not IsBusy
        End Function

        Public Sub Export()
            SaveFileDialogService.Filter = "PDF files (*.pdf)|*.pdf"
            If SaveFileDialogService.ShowDialog() Then
                IsBusy = True
                ExportTo(serverConnection, ReportId, New PdfExportOptions(), reportParameters, Sub(x) File.WriteAllBytes(SaveFileDialogService.GetFullFileName(), x))
            End If
        End Sub

        Public Function CanPrint() As Boolean
            Return Not IsBusy
        End Function

        Public Sub Print()
            ExportTo(serverConnection, ReportId, New XpsExportOptions(), reportParameters, Sub(x) PrintReportFromXps(x))
        End Sub

        Private Sub PrintReportFromXps(ByVal xps() As Byte)
            Using jobInfo As PrintSystemJobInfo = (New PrintDialog()).PrintQueue.AddJob("Print Job Name")
                jobInfo.JobStream.Write(xps, 0, xps.Length)
            End Using
        End Sub

        Private Sub ExportTo(ByVal serverConnection As ConnectionProvider, ByVal reportId As Integer, ByVal exportOptions As ExportOptionsBase, ByVal parameters() As ReportParameter, ByVal action As Action(Of Byte()))
            IsBusy = True
            serverConnection.ConnectAsync().ContinueWith(Function(t)
                Dim client As IReportServerClient = t.Result
                Return Task.Factory.ExportReportAsync(client, New ReportIdentity(reportId), exportOptions, parameters, Nothing)
            End Function).Unwrap().ContinueWith(Sub(t)
                IsBusy = False
                Try
                    If t.IsFaulted Then
                        Throw New Exception(t.Exception.Flatten().InnerException.Message)
                    End If
                    action(t.Result)
                Catch e As Exception
                    MessageBoxService.Show(e.Message, "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
                End Try
End Sub, TaskScheduler.FromCurrentSynchronizationContext())
        End Sub
    End Class
End Namespace
