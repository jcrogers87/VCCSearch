'\\Plataine 2015 for AUT
'\\searches a merged CSV for a column with VCC inside it and creates two csv files with the separated output. 
Imports System.IO
Module Module1
    'readconfig globals
    Dim inputJobSheet() As String, vccOutput As String, manualOutput As String, vccInt As Integer, headers As Boolean, errorDir As String
    'output files globals
    Dim vccJobsFilename As String, manualCutFilename As String, errorOutputFilename As String
    Sub Main()
        Call ReadConfig()
        Call ReadInputFile()
    End Sub
    Public Sub ReadInputFile()
        Dim i As Integer, j As Integer
        j = 1
        For i = 0 To inputJobSheet.Length - 1
            Dim inputJob() As String = File.ReadAllLines(inputJobSheet(i))
            For Each line As String In inputJob
                If j = 1 And headers = True Then
                    CreateVCCOutput(line)
                    'CreateManualOutput(line)
                    j = 0
                End If
                Dim column() As String = Split(line, ",")
                If Not column.Length < vccInt And j = 0 Then
                    If UCase(column(vccInt - 1)) = "VCC" Then
                        CreateVCCOutput(line)
                    ElseIf Not IsNothing(UCase(column(vccInt - 1))) And Not column(vccInt - 1) = "" And j = 0 Then
                        CreateManualOutput(line)
                    End If
                Else
                    Dim commas As String, k As Integer
                    For k = 1 To vccInt - Split(line, ",").Length
                        commas = commas & ","
                    Next
                    CreateManualOutput(line & commas & "Missing Data")
                    CreateErrorOutput(line & commas & "Missing from BOM")
                    commas = ""
                End If
            Next
        Next
    End Sub
    Public Sub CreateErrorOutput(line As String)
        If errorOutputFilename = "" Then
            If (Not Directory.Exists(errorDir)) Then
                Directory.CreateDirectory(errorDir)
            End If
            errorOutputFilename = errorDir & DateTime.Now.ToString("MM-dd.HH-mm") & "-Error.csv"
            If (File.Exists(errorOutputFilename)) Then errorOutputFilename = errorDir & DateTime.Now.ToString("MM-dd.HH-mm-ss") & "-Error.csv"
        End If
        Dim sw As New StreamWriter(errorOutputFilename, True)
        sw.WriteLine(line)
        sw.Close()
    End Sub
    Public Sub CreateVCCOutput(line As String)
        If vccJobsFilename = "" Then
            If (Not Directory.Exists(vccOutput)) Then
                Directory.CreateDirectory(vccOutput)
            End If
            vccJobsFilename = vccOutput & DateTime.Now.ToString("MM-dd.HH-mm") & "-VCCJobs.csv"
            If (File.Exists(vccJobsFilename)) Then vccJobsFilename = vccOutput & DateTime.Now.ToString("MM-dd.HH-mm-ss") & "-VCCJobs.csv"
        End If
        Dim sw As New StreamWriter(vccJobsFilename, True)
        sw.WriteLine(line)
        sw.Close()
    End Sub
    Public Sub CreateManualOutput(line As String)
        If manualCutFilename = "" Then
            If (Not Directory.Exists(manualOutput)) Then
                Directory.CreateDirectory(manualOutput)
            End If
            manualCutFilename = manualOutput & DateTime.Now.ToString("MM-dd.HH-mm") & "-ManualCut.csv"
            If (File.Exists(manualCutFilename)) Then manualCutFilename = vccOutput & DateTime.Now.ToString("MM-dd.HH-mm-ss") & "-ManualCut.csv"
        End If
        Dim sw As New StreamWriter(manualCutFilename, True)
        sw.WriteLine(line)
        sw.Close()
    End Sub
    Public Sub ReadConfig()
        Try
            headers = False
            Dim configFile() As String = File.ReadAllLines("C:\ProgramData\Plataine\VCCSearch.config")
            For Each line As String In configFile
                Dim setting() As String = Split(line, "=")
                If UCase(setting(0)) = "INPUTFILE" Then
                    inputJobSheet = Directory.GetFiles(setting(1))
                ElseIf UCase(setting(0)) = "VCCOUTPUT" Then
                    vccOutput = setting(1).ToString
                    If Not Right(vccOutput, 1) = "\" Then vccOutput = vccOutput & "\"
                ElseIf UCase(setting(0)) = "MANUALOUTPUT" Then
                    manualOutput = setting(1).ToString
                    If Not Right(manualOutput, 1) = "\" Then manualOutput = manualOutput & "\"
                    'column mappings:
                ElseIf UCase(setting(0)) = "HEADERS" Then
                    If UCase(setting(1).ToString) = "TRUE" Then headers = True
                ElseIf UCase(setting(0)) = "VCC" Then
                    vccInt = CInt(setting(1).ToString)
                ElseIf UCase(setting(0)) = "ERRORDIRECTORY" Then
                    errorDir = setting(1).ToString
                    If Not Right(errorDir, 1) = "\" Then errorDir = errorDir & "\"
                End If
            Next
            If IsNothing(inputJobSheet) Or IsNothing(vccOutput) Or IsNothing(manualOutput) Or IsNothing(vccInt) Then
                Call MsgBox("Your config file is invalid. Must be of the form:" _
                           & Chr(13) & "inputfile=pathtojobs\" _
                           & Chr(13) & "VCCOutput=path" _
                           & Chr(13) & "ManualOutput=path\" _
                           & Chr(13) & "Config location must be: C:\ProgramData\Plataine\VCCSearch.config")
                End
            End If
        Catch ex As Exception
            Call MsgBox("Your config file is missing, or missing required column mappings." _
                               & Chr(13) & "Config location must be: C:\ProgramData\Plataine\VCCSearch.config")
            End
        End Try
    End Sub
End Module
