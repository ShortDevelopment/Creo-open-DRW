Imports System.IO
Imports System.Runtime.InteropServices
Imports pfcls

Public Class EntryPoint

    Protected Shared CreoConnection As IpfcAsyncConnection = Nothing
    Protected Shared WithEvents Timer As New Timer()

    <STAThread>
    Public Shared Sub Main()
        Dim commandline = My.Application.CommandLineArgs
        If Not commandline.Count = 0 Then
            For i As Integer = 0 To commandline.Count - 1
                Dim cmd = commandline(i).Replace("/", "").Replace("-", "")
                If cmd = "gui" Then
                    Application.Run(New Form1())
                    Exit Sub
                End If
            Next
        End If

        _handler = Function(ByVal sig As CtrlType) As Boolean
                       Try
                           CreoConnection.InterruptEventProcessing()
                       Catch : End Try
                       Try
                           CreoConnection.Disconnect(1)
                       Catch : End Try
                       Environment.[Exit](-1)
                       Return True
                   End Function
        SetConsoleCtrlHandler(_handler, True)

        Try

            Console.Title = Application.ProductName
            Console.BackgroundColor = ConsoleColor.Black

            Try
                For Each p As Process In Process.GetProcessesByName("pfclscom")
                    p.Kill()
                Next
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Terminated last connections")
                Console.ForegroundColor = ConsoleColor.White
            Catch ex As Exception

            End Try

            Dim LangDirectory As String = "C:\Temp\creo_ex"

            If Directory.Exists(LangDirectory) Then
                Directory.Delete(LangDirectory, True)
            Else
                Directory.CreateDirectory(LangDirectory)
            End If
            My.Computer.FileSystem.CopyDirectory(Application.StartupPath + "\lang\", LangDirectory)

            If Process.GetProcessesByName("genlwsc").Count = 0 Then
                Console.WriteLine()
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Waiting for Creo Parametric...")
                Console.ForegroundColor = ConsoleColor.White
                While Process.GetProcessesByName("genlwsc").Count = 0
                    Threading.Thread.Sleep(500)
                End While
                Console.WriteLine()
            End If

            Console.WriteLine("Connecting to Creo Parametric...")
            CreoConnection = New CCpfcAsyncConnection().Connect(Nothing, Nothing, Nothing, Nothing)
            Console.WriteLine("Connection established.")

            Console.WriteLine()

            Console.WriteLine("Creating Button...")
            Dim CommandListener As New CreoCommandListener()
            AddHandler CommandListener.Command, Sub()
                                                    Try
                                                        Dim model As IpfcModel = CreoConnection.Session.CurrentModel

                                                        Console.WriteLine(model.Descr.GetFileName())

                                                        If IO.Path.GetExtension(model.Descr.GetFileName()) = ".asm" Then
                                                            If CreoConnection.Session.CurrentSelectionBuffer.Contents Is Nothing Then
                                                                'Throw New Exception("No Model Selected!")
                                                            Else
                                                                model = CreoConnection.Session.CurrentSelectionBuffer.Contents(0).SelModel
                                                            End If

                                                        End If
                                                        If IO.Path.GetExtension(model.Descr.GetFileName()) = ".drw" Then
                                                            If CreoConnection.Session.CurrentSelectionBuffer.Contents Is Nothing Then
                                                                Throw New Exception("Do you think there should be any *.drw files belonging to a *.drw file ?!")
                                                            Else
                                                                model = CreoConnection.Session.CurrentSelectionBuffer.Contents(0).SelModel
                                                            End If
                                                        End If

                                                        Dim path = $"{model.Descr.Device}:{IO.Path.Combine(model.Descr.Path, model.Descr.GetFileName())}"

                                                        'Console.WriteLine($"path: {path}")

                                                        Dim modelname = Split(model.Descr.GetFileName(), ".")(0)
                                                        'Console.WriteLine(model.Descr.InstanceName)
                                                        Dim pathnew As String
                                                        Try
                                                            pathnew = My.Computer.FileSystem.GetFiles(IO.Path.GetDirectoryName(path)).Where(Function(x) Split(IO.Path.GetFileName(x), ".")(0) = modelname AndAlso IO.Path.GetFileName(x).Contains(".drw") AndAlso Integer.TryParse(IO.Path.GetExtension(x).Replace(".", ""), Nothing)).OrderBy(Function(x)
                                                                                                                                                                                                                                                                                                                                                   Dim extension = IO.Path.GetExtension(x).Replace(".", "")
                                                                                                                                                                                                                                                                                                                                                   Return Int(extension)
                                                                                                                                                                                                                                                                                                                                               End Function).Reverse()(0)

                                                        Catch ex As Exception

                                                        End Try
                                                        path = path.Replace(".prt", ".drw").Replace(".asm", ".drw")
                                                        Console.WriteLine(path)

                                                        'If Not File.Exists(path) Then
                                                        '    Throw New FileNotFoundException("No *.drw file for this model!")
                                                        'End If

                                                        Dim drwModelDescriptor = New CCpfcModelDescriptor().CreateFromFileName(path)

                                                        If String.IsNullOrEmpty(pathnew) Then
                                                            'drwModelDescriptor.FileVersion = CType(1, Int32)
                                                        Else
                                                            Dim version = CType(Int(IO.Path.GetExtension(pathnew).Replace(".", "")), Int32)
                                                            Console.WriteLine($"Version: {version}")
                                                            drwModelDescriptor.FileVersion = version
                                                        End If
                                                        Dim info = New CCpfcRetrieveModelOptions().Create()
                                                        Dim loaded_model = CreoConnection.Session.RetrieveModelWithOpts(drwModelDescriptor, info)

                                                        CreoConnection.Session.OpenFile(drwModelDescriptor)

                                                        Dim window = CreoConnection.Session.GetModelWindow(loaded_model)
                                                        window.Activate()

                                                        Console.WriteLine("Method executed sucessfully.")
                                                    Catch ex As NullReferenceException
                                                        HandleException(New NullReferenceException("Kein Model / Keine Baugruppe geöffnet!"))
                                                    Catch ex As COMException
                                                        If ex.Message = "pfcExceptions::XToolkitNotFound" Then
                                                            HandleException(New FileNotFoundException("Datei konnte nicht gefunden werden!" + vbNewLine + "Sind illegale Zeichen im Pfad enthalten?"))
                                                        Else
                                                            HandleException(ex)
                                                        End If
                                                    Catch ex As Exception
                                                        HandleException(ex)
                                                    End Try
                                                End Sub

            Dim ButtonCommand As IpfcUICommand
            Try
                ButtonCommand = CreoConnection.Session.UICreateCommand("DRWOpen", CommandListener)
            Catch ex As Exception
                ButtonCommand = CreoConnection.Session.UIGetCommand("DRWOpen")
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Command Already Exists!")
                Console.ForegroundColor = ConsoleColor.White
            End Try

            ButtonCommand.Designate($"{LangDirectory}\de.lang", "DRWOpen Label", "DRWOpen Help", "DRWOpen Description")
            Console.WriteLine("Button Created.")


            Timer.Interval = 10
            AddHandler Timer.Tick, Sub()
                                       Try
                                           CreoConnection.EventProcess()
                                       Catch ex As Exception
                                           HandleException(ex, False)
                                       End Try
                                   End Sub
            Timer.Enabled = True

            'Console.WriteLine("Adding ActionListener...")
            'CreoConnection.AddActionListener(New CreoConnectionHandler(CreoConnection))
            'Console.WriteLine("ActionListener added.")

            Console.WriteLine("Entering Message Loop...")

            Console.WriteLine()
            Console.Beep()

            Application.Run()

        Catch ex As Exception
            HandleException(ex)
        Finally

            If Not CreoConnection Is Nothing AndAlso CreoConnection.IsRunning Then
                Try
                    Console.WriteLine("Disconnecting from Creo Parametric...")
                    CreoConnection.Disconnect(1)
                    Console.WriteLine("Sucessfully disconnected.")
                Catch ex As Exception
                    Console.WriteLine("Could not disconnected!")
                End Try
            End If

            Console.WriteLine("Press any Key to terminate...")
            Console.ReadKey()

            Console.WriteLine()
            Console.WriteLine("Terminating self...")
        End Try
    End Sub

    <DllImport("Kernel32")>
    Private Shared Function SetConsoleCtrlHandler(ByVal handler As ConsoleEventHandler, ByVal add As Boolean) As Boolean : End Function
    Private Delegate Function ConsoleEventHandler(ByVal sig As CtrlType) As Boolean
    Shared _handler As ConsoleEventHandler

    Enum CtrlType
        CTRL_C_EVENT = 0
        CTRL_BREAK_EVENT = 1
        CTRL_CLOSE_EVENT = 2
        CTRL_LOGOFF_EVENT = 5
        CTRL_SHUTDOWN_EVENT = 6
    End Enum

    Private Shared Function Handler(ByVal sig As CtrlType) As Boolean
        Try
            CreoConnection.InterruptEventProcessing()
        Catch : End Try
        Try
            CreoConnection.Disconnect(1)
        Catch : End Try

        Environment.[Exit](-1)
        Return True
    End Function

    Protected Shared Sub HandleException(ex As Exception, Optional ShouldNotifyUser As Boolean = True)
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine($"{ex.GetType().Name} thrown:")
        Console.WriteLine(ex.Message)
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine(ex.StackTrace)
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine()

        Try
            'Timer.Enabled = False
            'Dim dialogoptions = New CCpfcMessageDialogOptions().Create()
            'dialogoptions.DialogLabel = $"Error: {ex.GetType().Name}"
            'dialogoptions.MessageDialogType = EpfcMessageDialogType.EpfcMESSAGE_WARNING
            'Dim result = CreoConnection.Session.UIShowMessageDialog(ex.Message, dialogoptions)
            'Timer.Enabled = True
            If ShouldNotifyUser Then MessageBox.Show(ex.Message, $"Error: {ex.GetType().Name}", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, CType(&H40000, MessageBoxOptions) Or MessageBoxOptions.DefaultDesktopOnly)
        Catch : End Try
    End Sub

    Public Class CreoCommandListener
        Implements IpfcUICommandActionListener, ICIPClientObject

        Public Event Command()
        Public Sub OnCommand() Implements IpfcUICommandActionListener.OnCommand
            Console.WriteLine("Button was pressed!")
            RaiseEvent Command()
        End Sub

        Public Function GetClientInterfaceName() As String Implements ICIPClientObject.GetClientInterfaceName
            Return GetType(IpfcUICommandActionListener).Name
        End Function
    End Class
    Public Class CreoConnectionHandler
        Implements IpfcAsyncActionListener, IpfcActionListener, ICIPClientObject

        Public ReadOnly Property Connection As IpfcAsyncConnection
        Public Sub New(Connection As IpfcAsyncConnection)
            Me.Connection = Connection
        End Sub

        Public Sub OnTerminate(_Status As Integer) Implements IpfcAsyncActionListener.OnTerminate
            Connection.InterruptEventProcessing()
            Console.WriteLine()
            Console.WriteLine()
            Console.WriteLine("Creo Parametrics was terminated!")
            Application.Exit()
        End Sub

        Public Function GetClientInterfaceName() As String Implements ICIPClientObject.GetClientInterfaceName
            Return GetType(IpfcAsyncActionListener).Name
        End Function
    End Class

End Class
