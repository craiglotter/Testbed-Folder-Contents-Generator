Imports System.IO
Imports Microsoft.Win32

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread


   
    Public term1, term2, term3, term4, term5, foldername As String
    Public extension1, extension2, extension3, extension4, extension5 As String
    Public depth, amount As Integer

    
    Public result As String = ""

    Public Event WorkerErrorEncountered(ByVal ex As Exception)
    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress()
    Public Event WorkerFileCreated()



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        term1 = ""
        term2 = ""
        term3 = ""
        term4 = ""
        term5 = ""
        foldername = ""
        depth = 0
        extension1 = ""
        extension2 = ""
        extension3 = ""
        extension4 = ""
        extension5 = ""
        amount = 0
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal exc As Exception)
        Try
            RaiseEvent WorkerErrorEncountered(exc)
        Catch ex As Exception
            MsgBox("An error occurred in Testbed Folder Contents Generator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

   

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()
                Case 2
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerRemoveFolders)
                    ' Starts the thread.
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Public Sub WorkerExecute()
        Try
            'foldername = "C:\Documents and Settings\Administrator\Desktop\TestBed"
            'term1 = "hello"
            'term2 = "bye"
            'term3 = "cia"
            'term4 = "google"
            'term5 = "pookie"
            'depth = 5
            'amount = 9
            'extension1 = "txt"
            'extension2 = "asp"
            'extension3 = "inc"
            'extension4 = "bat"
            'extension5 = "exe"


            Dim terms(5) As String
            terms(0) = term1
            terms(1) = term2
            terms(2) = term3
            terms(3) = term4
            terms(4) = term5

            Dim extensions(5) As String
            extensions(0) = extension1
            extensions(1) = extension2
            extensions(2) = extension2
            extensions(3) = extension4
            extensions(4) = extension5
            Dim dinfo As DirectoryInfo = New DirectoryInfo(foldername)
            Dim r As DirectoryInfo
            Dim rt As DirectoryInfo
            Dim rtt As DirectoryInfo

            If dinfo.Exists = False Then
                dinfo.Create()
                RaiseEvent WorkerProgress()
            End If
            Dim dwinfo As DirectoryInfo = New DirectoryInfo(foldername)
            If dwinfo.Exists = True Then
                For kj As Integer = 1 To amount
                    dwinfo = New DirectoryInfo(foldername)
                    Randomize() ' Initialize random-number generator.
                    Dim MyValue As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                    Dim myvalue4 As Integer
                    Dim testinfo As DirectoryInfo = New DirectoryInfo((dwinfo.FullName & "\" & terms(MyValue)).Replace("\\", "\"))
                    If testinfo.Exists = False Then
                        dwinfo.CreateSubdirectory(terms(MyValue))
                        RaiseEvent WorkerProgress()
                        MyValue4 = CInt(Int((amount * Rnd()) + 1)) - 1  ' Generate random value between 0 and 4.
                        For t As Integer = 1 To MyValue4
                            create_file((dwinfo.FullName & "\" & terms(MyValue)).Replace("\\", "\"), terms, extensions)
                        Next
                    Else
                        Dim MyValue2 As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                        testinfo = New DirectoryInfo((dwinfo.FullName & "\" & terms(MyValue) & "_" & terms(MyValue2)).Replace("\\", "\"))
                        If testinfo.Exists = False Then
                            dwinfo.CreateSubdirectory(terms(MyValue) & "_" & terms(MyValue2))
                            RaiseEvent WorkerProgress()
                            MyValue4 = CInt(Int((amount * Rnd()) + 1)) - 1  ' Generate random value between 0 and 4.
                            For t As Integer = 1 To MyValue4
                                create_file((dwinfo.FullName & "\" & terms(MyValue)).Replace("\\", "\"), terms, extensions)
                            Next
                        End If
                    End If
                    testinfo = Nothing
                Next
                Dim MyValue3 As Integer = CInt(Int((amount * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                For t As Integer = 1 To MyValue3
                    create_file(dwinfo.FullName, terms, extensions)
                Next
                create_file(dwinfo.FullName, terms, extensions)
                If depth >= 2 Then
                    generate_sub_folders(foldername, terms, extensions, amount)
                End If
                If depth >= 3 Then
                    For Each r In dwinfo.GetDirectories
                        generate_sub_folders(r.FullName, terms, extensions, amount)
                        If depth >= 4 Then
                            For Each rt In r.GetDirectories
                                generate_sub_folders(rt.FullName, terms, extensions, amount)
                                If depth >= 5 Then
                                    For Each rtt In rt.GetDirectories
                                        generate_sub_folders(rtt.FullName, terms, extensions, amount)
                                    Next
                                End If

                            Next
                        End If
                    Next
                End If
            End If
            rtt = Nothing
            rt = Nothing
            r = Nothing
            dwinfo = Nothing
            dinfo = Nothing
            result = "success"
        Catch ex As Exception
            Error_Handler(ex)
            result = "fail"
        Finally
            RaiseEvent WorkerComplete(result)
        End Try
    End Sub

    Private Sub generate_sub_folders(ByVal foldername As String, ByVal terms As String(), ByVal extensions As String(), ByVal amount As Integer)
        Try
            Dim dwinfo As DirectoryInfo = New DirectoryInfo(foldername)
            If dwinfo.Exists = True Then
                For Each dwinfo In dwinfo.GetDirectories
                    Randomize() ' Initialize random-number generator.
                    Dim MyValue3 As Integer = CInt(Int((amount * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                    For kj As Integer = 1 To MyValue3
                        Randomize() ' Initialize random-number generator.
                        Dim MyValue As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                        Dim testinfo As DirectoryInfo = New DirectoryInfo((dwinfo.FullName & "\" & terms(MyValue)).Replace("\\", "\"))
                        If testinfo.Exists = False Then
                            dwinfo.CreateSubdirectory(terms(MyValue))
                            RaiseEvent WorkerProgress()
                            For t As Integer = 1 To MyValue3
                                create_file((dwinfo.FullName & "\" & terms(MyValue)).Replace("\\", "\"), terms, extensions)
                            Next
                        Else
                            Dim MyValue2 As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                            testinfo = New DirectoryInfo((dwinfo.FullName & "\" & terms(MyValue) & "_" & terms(MyValue2)).Replace("\\", "\"))
                            If testinfo.Exists = False Then
                                dwinfo.CreateSubdirectory(terms(MyValue) & "_" & terms(MyValue2))
                                RaiseEvent WorkerProgress()
                                For t As Integer = 1 To MyValue3
                                    create_file((dwinfo.FullName & "\" & terms(MyValue) & "_" & terms(MyValue2)).Replace("\\", "\"), terms, extensions)
                                Next
                            End If
                        End If
                        testinfo = Nothing
                    Next
              

                Next
                If dwinfo.GetFiles.Length = 0 Then
                    Randomize()
                    Dim MyValue3 As Integer = CInt(Int((amount * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
                    For t As Integer = 1 To MyValue3
                        create_file(dwinfo.FullName, terms, extensions)
                    Next
                End If
                dwinfo = Nothing
            End If
            dwinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub removefolders(ByVal foldername As String)
        Try
            Dim dinfo As DirectoryInfo = New DirectoryInfo(foldername)
            Dim finfo As DirectoryInfo
            For Each finfo In dinfo.GetDirectories
                removefolders(finfo.FullName)
            Next
            Dim f As FileInfo
            For Each f In dinfo.GetFiles
                f.Delete()
                RaiseEvent WorkerFileCreated()
            Next
            f = Nothing
            dinfo.Delete()
            RaiseEvent WorkerProgress()
            finfo = Nothing
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub workerremovefolders()
        Dim result As String
        Try
            '    foldername = "C:\Documents and Settings\Administrator\Desktop\TestBed"
            removefolders(foldername)
            result = "success"
        Catch ex As Exception
            Error_Handler(ex)
            result = "failure"
        Finally
            RaiseEvent WorkerComplete(result)
        End Try
    End Sub

    Private Sub create_file(ByVal foldername As String, ByVal terms As String(), ByVal extensions As String())
        Try
            Dim filename As String
            Randomize() ' Initialize random-number generator.
            Dim MyValue As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
            Dim MyValue2 As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
            Dim MyValue3 As Integer = CInt(Int((5 * Rnd()) + 1)) - 1 ' Generate random value between 0 and 4.
            Dim testinfo As FileInfo = New FileInfo((foldername & "\" & terms(MyValue)).Replace("\\", "\") & "." & extensions(MyValue2))
            If testinfo.Exists = False Then
                filename = testinfo.FullName
                Dim filewriter As IO.StreamWriter = New StreamWriter(filename, True)
                For i As Integer = 0 To MyValue
                    filewriter.Write(terms(MyValue2) & " " & terms(MyValue3) & " ")
                    filewriter.WriteLine(terms(MyValue))
                    MyValue2 = CInt(Int((5 * Rnd()) + 1)) - 1  ' Generate random value between 0 and 4.
                    MyValue3 = CInt(Int((5 * Rnd()) + 1)) - 1  ' Generate random value between 0 and 4.
                Next
                filewriter.Close()
                RaiseEvent WorkerFileCreated()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub
End Class
