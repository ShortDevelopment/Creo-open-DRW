Imports System.ComponentModel
Imports System.Configuration
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Threading
Imports pfcls

Public Class EntryPoint

    Protected Shared CreoConnection As IpfcAsyncConnection = Nothing
    Protected Shared WithEvents Timer As New Timer()

    Protected Shared ButtonCommand, SearchButtonCommand As IpfcUICommand
    Protected Shared CommandListener, CommandListener2 As CreoCommandListener

    Public Shared ReadOnly Property ThreadControl As Dispatcher = Dispatcher.CurrentDispatcher
    Public Shared ReadOnly Property UIDispatcher As Dispatcher
    Public Shared ReadOnly Property UIThread As New Threading.Thread(Sub()
                                                                         _UIDispatcher = Dispatcher.CurrentDispatcher
                                                                         Application.Run()
                                                                         Console.WriteLine($"[{Threading.Thread.CurrentThread.Name}]: Has Message Loop: {Application.MessageLoop}")
                                                                     End Sub)
    Public Shared ReadOnly Property LogDir As String = Path.Combine(Application.StartupPath, "Logs")
    <STAThread>
    Public Shared Sub Main()

        Application.EnableVisualStyles()
        Threading.Thread.CurrentThread.Name = "MainThread"

        Dim CreateNewConsole As Boolean = False

        If Not CreateNewConsole Then AttachConsole(-1)

        Dim IsConsoleAttached As Boolean = If(GetConsoleWindow() = IntPtr.Zero, False, True)

        Dim LogWriter As StreamWriter
        If Not IsConsoleAttached Then
            Try
                If Not Directory.Exists(LogDir) Then Directory.CreateDirectory(LogDir)
                LogWriter = New StreamWriter(File.Create(Path.Combine(LogDir, DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss") + ".log")))
                LogWriter.AutoFlush = True
                LogWriter.WriteLine("Started Log")
                'LogWriter.Flush()
                Console.SetOut(LogWriter)
                Console.SetError(LogWriter)
            Catch : End Try
        End If

        _handler = Function(ByVal sig As CtrlType) As Boolean
                       Try
                           Timer.Enabled = False
                           Timer.Stop()
                       Catch ex As Exception

                       End Try
                       If Not CreoConnection Is Nothing AndAlso CreoConnection.IsRunning Then
                           Try
                               CreoConnection.InterruptEventProcessing()
                           Catch : End Try
                           Try
                               Console.WriteLine("Disconnecting from Creo Parametric...")
                               CreoConnection.Disconnect(0)
                               Console.WriteLine("Sucessfully disconnected.")
                           Catch ex As Exception
                               Console.WriteLine("Could not disconnected!")
                           End Try
                       End If
                       Environment.[Exit](-1)
                       Return True
                   End Function
        If IsConsoleAttached Then SetConsoleCtrlHandler(_handler, True)

        Try

            If IsConsoleAttached Then Console.Title = Application.ProductName
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
            CommandListener = New CreoCommandListener()
            AddHandler CommandListener.Command, Sub()
                                                    Try
                                                        Dim model As CreoModel = New CreoModel(CType(CreoConnection.Session, IpfcBaseSession).CurrentModel)

                                                        Console.WriteLine(model.FileName)

                                                        If IO.Path.GetExtension(model.FileName) = ".asm" Then
                                                            If CreoConnection.Session.CurrentSelectionBuffer.Contents Is Nothing Then
                                                                'Throw New Exception("No Model Selected!")
                                                            Else
                                                                model = New CreoModel(CreoConnection.Session.CurrentSelectionBuffer.Contents(0).SelModel)
                                                            End If

                                                        End If
                                                        If IO.Path.GetExtension(model.FileName) = ".drw" Then
                                                            If CreoConnection.Session.CurrentSelectionBuffer.Contents Is Nothing Then
                                                                Throw New Exception("Do you think there should be any *.drw files belonging to a *.drw file ?!")
                                                            Else
                                                                model = New CreoModel(CreoConnection.Session.CurrentSelectionBuffer.Contents(0).SelModel)
                                                            End If
                                                        End If

                                                        Dim NewestFile As CreoFile = My.Computer.FileSystem.GetFiles(model.Directory).AsParallel().Select(Function(x) New CreoFile(x)).Where(Function(x) x.ModelName.ToLower() = model.ModelName.ToLower() AndAlso x.Extension = ".drw").OrderBy(Function(x) x.Version).Reverse()(0)

                                                        If NewestFile Is Nothing Then
                                                            HandleException(New NullReferenceException("Keine Zeichnung gefunden!"))
                                                        Else
                                                            CreoConnection.Session.OpenFile(NewestFile)
                                                        End If

                                                        Console.WriteLine("Method executed sucessfully.")
                                                    Catch ex As NullReferenceException
                                                        HandleException(New NullReferenceException("Kein Model / Keine Baugruppe geöffnet!", ex))
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

            'Dim ButtonCommand As IpfcUICommand
            Try
                ButtonCommand = CreoConnection.Session.UICreateCommand("DRWOpen", CommandListener)
            Catch ex As Exception
                ButtonCommand = CreoConnection.Session.UIGetCommand("DRWOpen")
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Command Already Exists! Using FallBack!")
                Console.ForegroundColor = ConsoleColor.White
                CommandListener.IsFallBack = True
                CType(ButtonCommand, IpfcActionSource).AddActionListener(CommandListener)
                CType(ButtonCommand, IpfcActionSource).AddActionListener(New CreoButtonActivateListener())
            End Try

            ButtonCommand.Designate($"{LangDirectory}\de.lang", "DRWOpen Label", "DRWOpen Help", "DRWOpen Description")

            Dim CommandListener3 = New CreoCommandListener()
            AddHandler CommandListener3.Command, Sub()
                                                     Try
                                                         Dim dir = IO.Path.GetDirectoryName(CType(CreoConnection.Session, IpfcBaseSession).CurrentModel.Origin)
                                                         CType(CreoConnection.Session, IpfcBaseSession).ChangeDirectory(dir)
                                                     Catch ex As NullReferenceException
                                                         HandleException(New NullReferenceException("Kein Model / Keine Baugruppe geöffnet!", ex))
                                                     Catch ex As Exception
                                                         HandleException(ex)
                                                     End Try
                                                 End Sub

            Dim WorkingDirectoryCommand As IpfcUICommand
            Try
                WorkingDirectoryCommand = CreoConnection.Session.UICreateCommand("WorkingDirectoryCommand", CommandListener3)
            Catch ex As Exception
                WorkingDirectoryCommand = CreoConnection.Session.UIGetCommand("WorkingDirectoryCommand")
            End Try
            WorkingDirectoryCommand.Designate($"{LangDirectory}\de.lang", "WorkingDirectoryCommand Label", "WorkingDirectoryCommand Help", "WorkingDirectoryCommand Description")

            Console.WriteLine("Buttons Created.")

            Timer.Interval = 10
            AddHandler Timer.Tick, Sub()
                                       If Not CreoConnection.IsRunning Then Exit Sub
                                       Try
                                           CreoConnection.EventProcess()
                                       Catch ex As Exception
                                           HandleException(ex, False)
                                       End Try
                                   End Sub
            Timer.Enabled = True

            Console.WriteLine("Adding ActionListener...")
            Dim ConnectionHandler = New CreoConnectionHandler()
            AddHandler ConnectionHandler.CreoTerminated, Sub()

                                                             Timer.Enabled = False
                                                             Timer.Stop()

                                                             Console.WriteLine()
                                                             Console.WriteLine()
                                                             Console.WriteLine("Preparing for Shutdown...")

                                                             Try
                                                                 UIDispatcher.Invoke(Sub()
                                                                                         'If Not DialogInstance Is Nothing Then DialogInstance.Close()
                                                                                         Application.ExitThread()
                                                                                         'ThreadControl.Invoke(Sub() UIThread.Abort())
                                                                                     End Sub)
                                                             Catch ex As Exception
                                                                 'HandleException(ex, False)
                                                             End Try

                                                             Try
                                                                 CreoConnection.Disconnect(2)
                                                             Catch ex As Exception

                                                             End Try

                                                             'Application.Exit()

                                                             Threading.Thread.Sleep(100)

                                                             Process.GetCurrentProcess().Kill()
                                                         End Sub
            CType(CreoConnection, IpfcActionSource).AddActionListener(ConnectionHandler)
            Console.WriteLine("ActionListener added.")

            Console.WriteLine("Starting UI Thread...")
            UIThread.IsBackground = True
            UIThread.Priority = Threading.ThreadPriority.AboveNormal
            UIThread.Name = "UIThread"
            UIThread.Start()
            Console.WriteLine("UI Thread started.")

            Console.Beep()
            Console.WriteLine("Entering Message Loop...")
            Console.WriteLine()
            Application.Run()

        Catch ex As Exception
            HandleException(ex)
        Finally
            Console.WriteLine("Goodbye!")
        End Try
    End Sub

    <DllImport("user32", SetLastError:=True)>
    Public Shared Sub SetForegroundWindow(handle As IntPtr) : End Sub

    <DllImport("user32", SetLastError:=True)>
    Public Shared Sub LockSetForegroundWindow(uLockCode As UInteger) : End Sub

    <DllImport("kernel32", SetLastError:=True)>
    Public Shared Sub FreeConsole() : End Sub
    <DllImport("kernel32", SetLastError:=True)>
    Public Shared Sub AttachConsole(pid As Integer) : End Sub
    <DllImport("kernel32", SetLastError:=True)>
    Public Shared Sub AllocConsole() : End Sub

    <DllImport("kernel32.dll")>
    Private Shared Function GetConsoleWindow() As IntPtr : End Function
    Private Const StdOutputHandle As Integer = &HFFFFFFF5
    <DllImport("kernel32.dll")>
    Private Shared Function GetStdHandle(ByVal nStdHandle As Integer) As IntPtr : End Function
    <DllImport("kernel32.dll")>
    Private Shared Sub SetStdHandle(ByVal nStdHandle As Integer, ByVal handle As IntPtr) : End Sub

    Protected Const LSFW_LOCK As UInteger = 1
    Protected Const LSFW_UNLOCK As UInteger = 2

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
            CType(ButtonCommand, IpfcActionSource).RemoveActionListener(CommandListener)
            CType(SearchButtonCommand, IpfcActionSource).RemoveActionListener(CommandListener2)
        Catch ex As Exception

        End Try
        Try
            CreoConnection.InterruptEventProcessing()
        Catch : End Try
        Try
            CreoConnection.Disconnect(1)
        Catch : End Try

        Environment.[Exit](-1)
        Return True
    End Function

    Public Shared Sub HandleException(ex As Exception, Optional ShouldNotifyUser As Boolean = True)
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
        Implements IpfcUICommandActionListener, IpfcUICommandBracketListener, IpfcActionListener, ICIPClientObject
        Public Property IsFallBack As Boolean
        Public Event Command()
        Public Sub OnCommand() Implements IpfcUICommandActionListener.OnCommand
            Console.WriteLine("Button was pressed!")
            RaiseEvent Command()
        End Sub

        Public Function GetClientInterfaceName() As String Implements ICIPClientObject.GetClientInterfaceName
            If IsFallBack Then
                Return GetType(IpfcUICommandBracketListener).Name
            Else
                Return GetType(IpfcUICommandActionListener).Name
            End If
        End Function

        Public Sub OnBeforeCommand() Implements IpfcUICommandBracketListener.OnBeforeCommand
            Console.WriteLine("Button was pressed! (FallBack)")
            RaiseEvent Command()
        End Sub

        Public Sub OnAfterCommand() Implements IpfcUICommandBracketListener.OnAfterCommand
        End Sub
    End Class
    Public Class CreoButtonActivateListener
        Implements IpfcUICommandAccessListener, IpfcActionListener, ICIPClientObject

        Public Function GetClientInterfaceName() As String Implements ICIPClientObject.GetClientInterfaceName
            Return GetType(IpfcUICommandAccessListener).Name
        End Function

        Public Function OnCommandAccess(_AllowErrorMessages As Boolean) As Integer Implements IpfcUICommandAccessListener.OnCommandAccess
            Return EpfcCommandAccess.EpfcACCESS_AVAILABLE
        End Function
    End Class

    Public Class CreoConnectionHandler
        Implements IpfcAsyncActionListener, IpfcActionListener, ICIPClientObject

        Public Event CreoTerminated()

        Public Sub OnTerminate(_Status As Integer) Implements IpfcAsyncActionListener.OnTerminate
            RaiseEvent CreoTerminated()
        End Sub

        Public Function GetClientInterfaceName() As String Implements ICIPClientObject.GetClientInterfaceName
            Return GetType(IpfcAsyncActionListener).Name
        End Function
    End Class

    Public Class CreoWindowHandle
        Implements IWin32Window

        <DllImport("user32")>
        Private Shared Function GetForegroundWindow() As IntPtr : End Function
        Public Sub New(window As IpfcWindow)
            CreoWindow = window
            window.Activate()
            _Handle = GetForegroundWindow()
        End Sub
        Public ReadOnly Property CreoWindow As IpfcWindow
        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
    End Class
End Class
