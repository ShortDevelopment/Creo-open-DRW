Imports System.IO
Imports System.Runtime.CompilerServices
Imports pfcls

Public Class CreoFile
    Public ReadOnly Property FileName As String
    Public ReadOnly Property Extension As String
    Public ReadOnly Property Version As Int32
    Public ReadOnly Property Origin As String
    Public ReadOnly Property ModelName As String
    Public ReadOnly Property Directory As String
        Get
            Return Path.GetDirectoryName(Origin)
        End Get
    End Property
    Public Sub New(path As String)
        Origin = path
        Extension = IO.Path.GetExtension(Origin)
        Dim i As Integer
        If Integer.TryParse(Extension.Replace(".", ""), i) Then
            Extension = IO.Path.GetExtension(IO.Path.GetFileNameWithoutExtension(Origin)).ToLower()
            FileName = IO.Path.GetFileNameWithoutExtension(Origin)
            Version = i
        Else
            FileName = IO.Path.GetFileName(path)
            Version = 0
        End If
        ModelName = IO.Path.GetFileNameWithoutExtension(FileName)
    End Sub
End Class

Public Class CreoModel
    Public ReadOnly Property CurrentModel As IpfcModel
    Public ReadOnly Property Extension As String
    Public ReadOnly Property Directory As String
        Get
            Return Path.GetDirectoryName(Origin)
        End Get
    End Property
    Public ReadOnly Property FileName As String
        Get
            Return $"{ModelName}{Extension}"
        End Get
    End Property
    Public ReadOnly Property ModelName As String
    Public ReadOnly Property Version As Int32
    Public ReadOnly Property Origin As String
        Get
            Return CurrentModel.Origin
        End Get
    End Property
    Public Sub New(model As IpfcModel)
        CurrentModel = model
        _ModelName = CurrentModel.Descr.InstanceName
        Dim File As New CreoFile(model.Origin)
        Extension = File.Extension
        Version = File.Version
    End Sub
End Class

Module Extensions
    <Extension>
    Public Sub OpenFile(ByRef session As IpfcSession, model As CreoModel)
        OpenInternal(session, Path.Combine(model.Directory, model.FileName), model.Version)
    End Sub
    <Extension>
    Public Sub OpenFile(ByRef session As IpfcSession, file As CreoFile)
        OpenInternal(session, Path.Combine(file.Directory, file.FileName), file.Version)
    End Sub
    Private Sub OpenInternal(ByRef session As IpfcSession, path As String, version As Int32)
        Dim drwModelDescriptor As IpfcModelDescriptor = New CCpfcModelDescriptor().CreateFromFileName(path)

        drwModelDescriptor.Path = IO.Path.GetDirectoryName(path) + IO.Path.DirectorySeparatorChar
        Console.WriteLine(drwModelDescriptor.GenericName)
        Console.WriteLine(drwModelDescriptor.InstanceName)

        If version > 0 Then
            drwModelDescriptor.FileVersion = version
        End If
        Dim info = New CCpfcRetrieveModelOptions().Create()
        Dim loaded_model = CType(session, IpfcBaseSession).RetrieveModelWithOpts(drwModelDescriptor, info)

        CType(session, IpfcBaseSession).OpenFile(drwModelDescriptor)

        Dim window = CType(session, IpfcBaseSession).GetModelWindow(loaded_model)
        window.Activate()
    End Sub
    <Extension>
    Public Sub OpenFile(ByRef session As IpfcSession, file As String)
        OpenFile(session, New CreoFile(file))
    End Sub
End Module