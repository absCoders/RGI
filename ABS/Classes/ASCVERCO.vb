Imports System.Xml

Public Class ASCVERCO

    Private xmlPath As String = String.Empty
    Private basePath As String = String.Empty
    Private wkPath As String = String.Empty
    Private Const xmlFileName As String = "Assemblies.xml"

    Private Const projectCodes As String = "ABS,ABSCS,ABSX,AP,AR,AS,CC,DE,ED,GL,IC,PO,PP,SA,SH,SO,TA,TAC,WB"
    Private Const testingMode As Boolean = False


    Public Sub New(ByVal CreateMissingXmlDocument As Boolean)

        If ASCMAIN1.Running_in_VS Then
            ' This line is here for testing purposes
            xmlPath = My.Settings.Item("Assemblies").ToString

            xmlPath = Application.StartupPath
            If xmlPath.Length > 0 AndAlso Not xmlPath.EndsWith("\") Then
                xmlPath &= "\"
            End If

            basePath = xmlPath
            basePath = basePath.ToUpper.Replace("\ABS\", "\@@@@\")
            basePath = basePath.ToUpper.Replace("\DEBUG\", "\RELEASE\")
        Else
            xmlPath = My.Settings.Item("Assemblies").ToString
            If xmlPath.Length > 0 AndAlso Not xmlPath.EndsWith("\") Then
                xmlPath &= "\"
            End If

            basePath = Application.StartupPath
            If basePath.Length > 0 AndAlso Not basePath.EndsWith("\") Then
                basePath &= "\"
            End If
        End If

        wkPath = String.Empty
        xmlPath &= xmlFileName

        If Not My.Computer.FileSystem.FileExists(xmlPath) AndAlso CreateMissingXmlDocument Then
            CreateXmlVersionFile()
        End If

    End Sub

    Public ReadOnly Property ABSProjectCodes() As String
        Get
            Return projectCodes
        End Get
    End Property

    Private Sub CreateXmlVersionFile()

        Using writer As New XmlTextWriter(xmlPath, System.Text.Encoding.UTF8)

            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 5
            writer.WriteStartElement("Assemblies")

            For Each AssemblyName As String In projectCodes.Split(",")
                AssemblyName = AssemblyName.Trim
                If AssemblyName.Length = 0 Then
                    Continue For
                End If

                writer.WriteStartElement("Assembly")

                writer.WriteStartElement("Name")
                writer.WriteString(AssemblyName)
                writer.WriteEndElement()

                wkPath = basePath.Replace("@@@@", AssemblyName)
                If AssemblyName <> "ABS" Then
                    wkPath &= AssemblyName & ".DLL"
                Else
                    wkPath &= AssemblyName & ".EXE"
                End If

                writer.WriteStartElement("Version")

                If My.Computer.FileSystem.FileExists(wkPath) Then
                    writer.WriteString(FileVersionInfo.GetVersionInfo(wkPath).FileVersion)
                Else
                    writer.WriteString("Unknown")
                End If
                writer.WriteEndElement()

                writer.WriteEndElement()
            Next

            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()
        End Using

        If ASCMAIN1.Running_in_VS Then
            My.Computer.FileSystem.CopyFile(xmlPath, "C:\VS\" & ASCMAIN1.CLIENT_CODE & "\BIN\" & xmlFileName, True)
        End If

    End Sub

    Public Sub UpdateVersionNumber(ByVal AssemblyName As String)

        '<?xml version="1.0" encoding="utf-8" standalone="yes"?>
        '<Assemblies>
        '     <Assembly>
        '          <Name>ABS</Name>
        '          <Version>1.1.1229.1</Version>
        '     </Assembly>

        Dim VersionNo As String = String.Empty

        wkPath = basePath.Replace("@@@@", AssemblyName)
        If AssemblyName <> "ABS" Then
            wkPath &= AssemblyName & ".DLL"
        Else
            wkPath &= AssemblyName & ".EXE"
        End If

        If My.Computer.FileSystem.FileExists(wkPath) Then
            VersionNo = (FileVersionInfo.GetVersionInfo(wkPath).FileVersion)
        Else
            VersionNo = "Unknown"
        End If


        Dim myXmlDocument As XmlDocument = New XmlDocument()
        myXmlDocument.Load(xmlPath)

        Dim node As XmlNode
        node = myXmlDocument.DocumentElement

        For Each node In node.ChildNodes
            If node.FirstChild.InnerText = AssemblyName Then
                node.ChildNodes(1).InnerText = VersionNo
                Exit For
            End If
        Next

        myXmlDocument.Save(xmlPath)

        If ASCMAIN1.Running_in_VS Then
            My.Computer.FileSystem.CopyFile(xmlPath, "C:\VS\" & ASCMAIN1.CLIENT_CODE & "\BIN\" & xmlFileName, True)
        End If

    End Sub

    Public Function CompareAssemblyVersions(ByRef ErrorMessage As String) As Boolean

        ErrorMessage = String.Empty
        Dim VersionNo As String = String.Empty
        Dim xmlVersionNo As String = String.Empty

        Dim xmlFileToRead As String = xmlPath
        If ASCMAIN1.Running_in_VS Then
            xmlFileToRead = "C:\VS\" & ASCMAIN1.CLIENT_CODE & "\BIN\" & xmlFileName
        End If

        If testingMode Then
            MessageBox.Show("xmlPath: " & xmlPath & Environment.NewLine & "basePath: " & basePath & Environment.NewLine & "xmlFileToRead: " & xmlFileToRead)
        End If

        Try
            If Not My.Computer.FileSystem.FileExists(xmlFileToRead) Then
                ErrorMessage = "Cannot locate Assembly Version XML Document"
                Return False
            End If

            Dim myXmlDocument As XmlDocument = New XmlDocument()
            myXmlDocument.Load(xmlFileToRead)

            For Each AssemblyName As String In projectCodes.Split(",")
                Dim node As XmlNode
                node = myXmlDocument.DocumentElement

                For Each node In node.ChildNodes
                    If node.FirstChild.InnerText = AssemblyName Then
                        xmlVersionNo = node.ChildNodes(1).InnerText

                        wkPath = basePath.Replace("@@@@", AssemblyName)
                        If AssemblyName <> "ABS" Then
                            wkPath &= AssemblyName & ".DLL"
                        Else
                            wkPath &= AssemblyName & ".EXE"
                        End If

                        If testingMode Then
                            MessageBox.Show("wkPath: " & wkPath)
                        End If

                        If My.Computer.FileSystem.FileExists(wkPath) Then
                            VersionNo = (FileVersionInfo.GetVersionInfo(wkPath).FileVersion)
                        Else
                            VersionNo = ""
                        End If

                        If VersionNo <> xmlVersionNo Then
                            ErrorMessage &= "Assembly " & AssemblyName & " has the wrong version: " & VersionNo & "/" & xmlVersionNo & Environment.NewLine
                        End If

                    End If
                Next
            Next

            Return ErrorMessage.Length = 0

        Catch ex As Exception
            ErrorMessage = ex.Message
            Return False
        End Try
    End Function

End Class

