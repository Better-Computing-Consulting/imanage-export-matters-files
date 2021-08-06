Imports System.Text.RegularExpressions
Imports System.Data.SqlClient
Imports System.IO
Module Module1
    Sub Main()
        Dim Exproot As String = "C:\TEMP\outfolder"
        Using sr As New StreamReader(Exproot & "\clientmatters.txt")
            Do While sr.Peek >= 0
                Dim line As String() = sr.ReadLine.Split(",")
                Console.WriteLine(".." & line(0) & ".." & line(1) & "..")
                GetFilesFromSQL(line(0), line(1), Exproot)
            Loop
        End Using
        Console.WriteLine(vbCrLf & "done")
        Console.ReadLine()
    End Sub
    Sub GetFilesFromSQL(ClientNo As String, CaseNo As String, RootFolder As String)
        Dim FilesToExport As New List(Of FileToExport)
        Dim connString1 As String = "Data Source=sql;Initial Catalog=iManage_Active;Integrated Security=SSPI"
        Dim queryString As String = "select DOCNAME,DOCLOC,T_ALIAS,C_ALIAS,C2ALIAS from MHGROUP.DOCMASTER where C_ALIAS != 'WEBDOC' and C1ALIAS = '" & ClientNo & "' and  C2ALIAS = '" & CaseNo & "' order by DOCNUM"
        Using sw As New StreamWriter(RootFolder & "\ExportedFiles.txt", True)
            sw.AutoFlush = True
            sw.WriteLine("# " & ClientNo & "." & CaseNo)
            Using conn As New SqlConnection(connString1)
                Dim cmd As New SqlCommand(queryString, conn)
                conn.Open()
                Dim r As SqlDataReader = cmd.ExecuteReader()
                If r.HasRows Then
                    Try
                        While r.Read
                            Dim docName As String = Trim(r("DOCNAME"))
                            Dim dpath As String = Trim(r("DOCLOC"))
                            Dim docpath As String = ""
                            If dpath.Contains("DEFSERVER2") Then
                                docpath = dpath.Replace("DEFSERVER2:", "\\svrwebctrl\imandocs")
                            Else
                                docpath = dpath.Replace("DEFSERVER:", "\\svrfps1\imandocs")
                            End If
                            sw.WriteLine(docpath)
                            Dim doctype As String = Trim(r("T_ALIAS"))
                            Dim docdir As String = Trim(r("C_ALIAS"))
                            Dim newdocname As String = GetDMSFileName(docName, doctype)
                            Dim newdocpath As String = RootFolder & "\" & ClientNo & "." & CaseNo & "\" & GetDMSDirectory(docdir)
                            FilesToExport.Add(New FileToExport(docpath, newdocpath & "\" & newdocname))
                        End While
                    Catch ex As Exception
                    End Try
                End If
            End Using
        End Using
        For Each doc As FileToExport In FilesToExport
            Console.WriteLine("{0,-60}{1}", doc.MoveFrom, GetNextFreeFile(doc.MoveTo))
            Try
                Dim destDirectory As String = New FileInfo(doc.MoveTo).DirectoryName
                If Not Directory.Exists(destDirectory) Then
                    Directory.CreateDirectory(destDirectory)
                End If
                File.Copy(doc.MoveFrom, GetNextFreeFile(doc.MoveTo))
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        Next
    End Sub
    Function GetDMSFileName(dbname As String, filetype As String)
        If String.IsNullOrEmpty(dbname.Trim) Then
            Return "(no title)" & GetDMSSuffix(filetype)
        End If
        Dim tmpfname As String = Regex.Replace(dbname, "[<>:""/\\|?*]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5))
        If String.IsNullOrEmpty(tmpfname.Trim) Then
            Return "(no title)" & GetDMSSuffix(filetype)
        Else
            Return tmpfname & GetDMSSuffix(filetype)
        End If
    End Function
    Function GetDMSDirectory(DirType As String) As String
        Select Case DirType.ToUpper
            Case "BANKRUPTCY"
                Return "Bankruptcy" 'top
            Case "CORR"
                Return "Correspondence" 'top
            Case "DEP"
                Return "Deposition"
            Case "DISC"
                Return "Discovery" 'top
            Case "DOC"
                Return "Documents" 'top
            Case "DOCKET"
                Return "Docket"
            Case "E-MAIL"
                Return "E-Mail" ' top
            Case "EXECUTED"
                Return "Executed" 'top
            Case "EXPERT"
                Return "Expert"
            Case "HOA_CLOSED"
                Return "HOA Closed Files"
            Case "HOA_COLLECTIONS"
                Return "HOA Collections"
            Case "HOA_COMMERCIAL"
                Return "HOA Commercial"
            Case "HOA_INJUNCTION"
                Return "HOA Injunction"
            Case "MNG"
                Return "Manage"
            Case "NTM"
                Return "Notes-Memos"
            Case "PENDING"
                Return "Pending"
            Case "PLD"
                Return "Pleadings" 'top
            Case "PRIVATE"
                Return "Private"
            Case "PROD"
                Return "Production"
            Case "PROFILE"
                Return "Profile"
            Case "TBF"
                Return "To be Filed" 'top
            Case "TRIAL"
                Return "Trial" 'top
            Case "WASTE"
                Return "Waste" 'top
            Case "WEBDOC"
                Return "infoLink Web Page"
            Case Else
                Return ""
        End Select
    End Function
    Function GetNextFreeFile(aFileName As String) As String
        Dim fileone As String = aFileName
        Dim newfile As String = ""
        If File.Exists(fileone) Then
            Dim filesufix As String = New FileInfo(fileone).Extension
            Dim filebase As String = fileone.Substring(0, fileone.Length - filesufix.Length)
            Dim i As Int16 = 0
            Do
                i += 1
                newfile = filebase & "(" & i & ")" & filesufix
            Loop While File.Exists(newfile)
        Else
            newfile = fileone
        End If
        Return newfile
    End Function
    Function GetDMSSuffix(filetype As String) As String
        Select Case filetype.ToUpper
            Case "ANSI"
                Return ""
            Case "HTML"
                Return ".html"
            Case "ACROBAT", "ECOPY"
                Return ".pdf"
            Case "BMP"
                Return ".bmp"
            Case "ETRANSCRIPT"
                Return ".ptx"
            Case "EXCEL"
                Return ".xls"
            Case "EXCELX"
                Return ".xlsx"
            Case "GIF"
                Return ".gif"
            Case "JPEG"
                Return ".jpeg"
            Case "MIME"
                Return ".msg"
            Case "NOTES"
                Return ".dxl"
            Case "PCX"
                Return ".pcx"
            Case "PPT"
                Return ".ppt"
            Case "PPTX"
                Return ".pptx"
            Case "TIFF"
                Return ".tiff"
            Case "URL"
                Return ".url"
            Case "WAV"
                Return ".wav"
            Case "WMV"
                Return ".wmv"
            Case "WORD"
                Return ".doc"
            Case "WORDX"
                Return ".docx"
            Case "WORDXT"
                Return ".doct"
            Case "WPF"
                Return ".wpd"
            Case "XML"
                Return ".xml"
            Case Else
                Return ""
        End Select
    End Function
End Module
Class FileToExport
    Sub New(InFrom As String, InTo As String)
        MoveFrom = InFrom
        MoveTo = InTo
    End Sub
    Public MoveFrom As String
    Public MoveTo As String
End Class