Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Reflection
Imports Microsoft.PowerShell
Imports System.Text.RegularExpressions

Public Class ASFDEPL1
    Private F As New List(Of String)
    Private tblProjects As New DataTable

    Private client As String = String.Empty
    Private clientEnvironment As String = String.Empty
    Private WorkingDirectory As String = String.Empty

    Private deployScript As String = String.Empty
    Private deployScriptfileName As String = String.Empty

    Private Sub ASFDEPL1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        tblProjects.Columns.Add("SELECTED", GetType(System.Int16))
        tblProjects.Columns.Add("DLL_NAME", GetType(System.String))
        tblProjects.Columns.Add("DLL_DESC", GetType(System.String))

        Dim tblASTMENU1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ASTMENU1 WHERE MENU_ITEM_TYPE = 'M' AND MENU_ID = 'MAIN'", "ASTMENU1")

        For Each dllName As String In New String() {"ABS", "ABSCS", "ABSX", "AP", "AR", "AS", "AT", "CC", "EC", "ED", "GL", "IC", "PO", "SA", "SO", "TA", "TAC", "WB", "WH", "WHC", "WO"}
            Dim DLL_DESC As String = String.Empty
            If tblASTMENU1.Select($"MENU_ITEM_OBJECT = '{dllName}'").Length > 0 Then
                DLL_DESC = tblASTMENU1.Select($"MENU_ITEM_OBJECT = '{dllName}'")(0).Item("MENU_ITEM_DESC") & String.Empty
            End If

            If dllName = "ABS" Then
                DLL_DESC = "Main Application"
            End If
            tblProjects.Rows.Add(New Object() {"0", dllName.Trim, DLL_DESC})
        Next

        grdDLLS.DataSource = tblProjects
        grdDLLS.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdDLLS.DisplayLayout.Bands(0).SortedColumns.Add(grdDLLS.DisplayLayout.Bands(0).Columns("DLL_NAME"), False)


        WorkingDirectory = Application.StartupPath
        If ASCMAIN1.Running_in_VS Then
            Dim USERNAME As String = System.Environment.GetEnvironmentVariable("USERNAME") & String.Empty
            WorkingDirectory = $"C:\Users\{USERNAME}\VS\VDI"
        End If

        deployScriptfileName = $"{WorkingDirectory}\deploy.ps1"
        deployScript = File.ReadAllText(deployScriptfileName)

        Dim pattern As String = "clientSettings\s*=\s*@{\s*(""(.+?)""\s*=\s*@{(.*?)\};\s*)+\s*}"
        Dim mtch As Match = Regex.Match(deployScript, pattern, RegexOptions.Singleline)

        Dim tbl As New DataTable
        tbl.Columns.Add("Client", GetType(System.String))
        tbl.Columns.Add("IPADDRESS_PROD", GetType(System.String))
        tbl.Columns.Add("IPADDRESS_TEST", GetType(System.String))

        For icapt As Integer = 0 To mtch.Groups(2).Captures.Count - 1
            Dim data As String = mtch.Groups(3).Captures(icapt).Value
            data = data.Replace(Chr(34), "").Replace(vbCrLf, "")
            Dim splData() As String = data.Split(";")
            Dim dict As Dictionary(Of String, String) = splData.Select(Function(x) x.Split("=")).ToDictionary(Function(x) x(0).Trim, Function(x) x(1).Trim)

            Dim IPADDRESS_PROD As String = dict("PROD")
            Dim IPADDRESS_TEST As String = dict("QA")
            tbl.Rows.Add(New Object() {mtch.Groups(2).Captures(icapt).Value, IPADDRESS_PROD, IPADDRESS_TEST})
        Next

        cmbClient.DataSource = tbl
        cmbClient.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

    End Sub

    Private Sub cmdDeploy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeploy.Click

        Dim releaseDirectory As String = String.Empty
        If Not ValidateSelections(True, True, releaseDirectory) Then
            Exit Sub
        End If

        Dim lstAssemblies As New List(Of String)
        For Each row As DataRow In tblProjects.Select("SELECTED = '1'")
            lstAssemblies.Add(row.Item("DLL_NAME"))
        Next
        lstAssemblies.Sort()

        client = cmbClient.Text
        Dim region As String = optRegion.CheckedItem.DisplayText

        Dim msg As String = String.Empty
        msg &= "Selected Assemblies: " & String.Join(", ", lstAssemblies.ToArray)
        msg &= Environment.NewLine & Environment.NewLine
        msg &= $"Do you want to deploy the above assemblies to the {region} region for client: {client} ?"
        If MessageBox.Show(msg, "Deploy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Exit Sub
        End If

        msg = "Selected Assemblies: " & String.Join(", ", lstAssemblies.ToArray)
        msg &= Environment.NewLine & Environment.NewLine
        msg &= $"Are you sure you want to deploy the above assemblies to the {region} region for client: {client}?"
        If MessageBox.Show(msg, "Deploy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Exit Sub
        End If

        Deploy_Assemblies()

        MessageBox.Show("Deployment Complete", "Deploy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Sub Deploy_Assemblies()

        Try

            Dim releaseDirectory As String = String.Empty
            If Not ValidateSelections(True, True, releaseDirectory) Then
                Exit Sub
            End If

            Dim selectedAssemblies As New List(Of String)

            For Each row As DataRow In tblProjects.Select("SELECTED = '1'")
                selectedAssemblies.Add(row.Item("DLL_NAME"))
            Next

            Dim client As String = cmbClient.Text
            Dim region As String = String.Empty
            Select Case optRegion.Value
                Case "P"
                    region = "PROD"
                Case "T"
                    region = "QA"
            End Select

            ASCMAIN1.Progress($"Deploying to {client}...")

            Using runspace As Runspace = RunspaceFactory.CreateRunspace()
                runspace.Open()

                Dim sessionState As InitialSessionState = InitialSessionState.Create()

                'sessionState.ExecutionPolicy = ExecutionPolicy.RemoteSigned                
                Dim execPolProp As PropertyInfo = sessionState.GetType().GetProperty("ExecutionPolicy")
                If (execPolProp IsNot Nothing AndAlso execPolProp.CanWrite) Then
                    execPolProp.SetValue(sessionState, ExecutionPolicy.Bypass, Nothing)
                End If

                sessionState.LanguageMode = PSLanguageMode.FullLanguage
                Dim ps As PowerShell = PowerShell.Create(sessionState)
                ps.Runspace = runspace

                ps.AddScript(File.ReadAllText($"{WorkingDirectory}\deploy.ps1"))

                ps.AddScript($"Deploy-Assemblies -deployToEnvironments (""{region}"") -assembliesToDeploy (""{String.Join(""",""", selectedAssemblies)}"") -client ""{client}""")

                Dim results As Collection(Of PSObject) = ps.Invoke()

                If ps.HadErrors Then

                    For Each errorStream As ErrorRecord In ps.Streams.Error
                        MsgBox(errorStream.Exception.Message)
                    Next
                Else
                    'Need to add code to verify what was deployed and display to screen
                    'Probably need to add output to the deploy.ps1 script
                End If

                ps.Dispose()

            End Using

            SelectDlls("0")
            ASCMAIN1.Progress("", "")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Private Function ValidateSelections(ByVal validateDlls As Boolean, ByVal validateDirectories As Boolean, ByRef releaseDirectory As String) As Boolean

        Try
            ValidateSelections = False
            tblProjects.AcceptChanges()

            If validateDlls Then
                If tblProjects.Select("SELECTED = '1'").Length = 0 Then
                    MessageBox.Show("You must select at least one DLL.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Function
                End If
            End If

            If cmbClient.SelectedRow Is Nothing Then
                MessageBox.Show("You must select a client.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Function
            End If

            If validateDirectories Then

                If optRegion.Value = "P" Then
                    releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_PROD").Text
                    If releaseDirectory.Length = 0 Then
                        MessageBox.Show("The selected client does not have a Production Application Directory.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If

                    If Not My.Computer.FileSystem.DirectoryExists(releaseDirectory) Then
                        MessageBox.Show("The selected client's Production Application Directory cannot be located.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
                End If

                If optRegion.Value = "T" Then
                    releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_TEST").Text
                    If releaseDirectory.Length = 0 Then
                        MessageBox.Show("The selected client does not have a Test Application Directory.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If

                    If Not My.Computer.FileSystem.DirectoryExists(releaseDirectory) Then
                        MessageBox.Show("The selected client's Test Application Directory cannot be located.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
                End If

                Select Case optRegion.Value
                    Case "P"
                        releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_PROD").Text
                    Case "T"
                        releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_TEST").Text
                End Select

                If Not releaseDirectory.EndsWith("\") Then releaseDirectory &= "\"
                releaseDirectory &= "Releases"

                If Not My.Computer.FileSystem.DirectoryExists(releaseDirectory) Then
                    My.Computer.FileSystem.CreateDirectory(releaseDirectory)
                End If
            End If

            ValidateSelections = True
        Catch ex As Exception
            MessageBox.Show("Validate Selections Error: " & ex.Message, "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        SelectDlls("1")
    End Sub

    Private Sub btnDeSelect_Click(sender As Object, e As EventArgs) Handles btnDeSelect.Click
        SelectDlls("0")
    End Sub

    Private Sub SelectDlls(ByVal value As String)
        ' tblProjects.Columns.Add("SELECTED", GetType(System.Int16))
        For Each row As DataRow In tblProjects.Select("")
            row.Item("SELECTED") = value
        Next
    End Sub

End Class
