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
    Private WorkingDirectory As String = "C:\Users\ABS\Projects\VDI"

    Private deployScript As String = String.Empty
    Private deployScriptfileName As String = String.Empty
    Private dicClients As New Dictionary(Of String, String)

    Private Sub ASFDEPL1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        tblProjects.Columns.Add("SELECTED", GetType(System.Int16))
        tblProjects.Columns.Add("DLL_NAME", GetType(System.String))

        For Each dllName As String In New String() {"ABS", "ABSCS", "ABSX", "AP", "AR", "AS", "AT", "CC", "ED", "GL", "IC", "PO", "SA", "SO", "TA", "TAC", "WB", "WH", "WHC", "WO"}
            tblProjects.Rows.Add(New Object() {"0", dllName.Trim})
        Next

        grdDLLS.DataSource = tblProjects
        grdDLLS.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdDLLS.DisplayLayout.Bands(0).SortedColumns.Add(grdDLLS.DisplayLayout.Bands(0).Columns("DLL_NAME"), False)

        deployScriptfileName = $"{WorkingDirectory}\deploy.ps1"
        deployScript = File.ReadAllText(deployScriptfileName)

        Dim pattern As String = "environmentSettings\s*=\s*@{\s*(""(.+?)""\s*=\s*\(""(.*?)""\);\s*)+\s*}"
        Dim mtch As Match = Regex.Match(deployScript, pattern, RegexOptions.Singleline)

        Dim tbl As New DataTable
        tbl.Columns.Add("Client", GetType(System.String))
        tbl.Columns.Add("IPADDRESS_PROD", GetType(System.String))
        tbl.Columns.Add("IPADDRESS_TEST", GetType(System.String))

        For icapt As Integer = 0 To mtch.Groups(2).Captures.Count - 1
            dicClients.Add(mtch.Groups(2).Captures(icapt).Value, mtch.Groups(3).Captures(icapt).Value)

            Dim IPADDRESS_PROD As String = mtch.Groups(3).Captures(icapt).Value.Split(";")(0)
            Dim IPADDRESS_TEST As String = mtch.Groups(3).Captures(icapt).Value.Split(";")(1)

            tbl.Rows.Add(New Object() {mtch.Groups(2).Captures(icapt).Value, IPADDRESS_PROD, IPADDRESS_TEST})
        Next

        cmbClient.DataSource = tbl
        cmbClient.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
    End Sub

    Private Sub SetReleaseFolders(ByVal validateDirectories As Boolean)

        Try
            Me.Cursor = Cursors.WaitCursor
            Dim releaseDirectory As String = String.Empty

            cmbReleases.Items.Clear()
            If Not ValidateSelections(False, validateDirectories, releaseDirectory) Then
                Exit Sub
            End If

            If validateDirectories Then
                Dim latestReleases = Directory.GetDirectories(releaseDirectory, "*", SearchOption.TopDirectoryOnly).AsEnumerable().OrderByDescending(Function(x) x).Take(5).Select(Function(x) x.Substring(x.LastIndexOf("\") + 1)).ToArray()
                cmbReleases.Items.Clear()
                cmbReleases.Items.AddRange(latestReleases)

                If cmbReleases.Items.Count > 0 Then
                    cmbReleases.SelectedIndex = 0
                End If
            End If


        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("SetReleaseFolders Error: " & ex.Message, "Set Release Folders", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
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

        Dim msg As String = String.Empty
        msg &= "Selected Assemblies: " & String.Join(", ", lstAssemblies.ToArray)
        msg &= Environment.NewLine & Environment.NewLine
        msg &= $"Do you want to deploy the above assemblies to the following client: {client}?"
        If MessageBox.Show(msg, "Deploy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Exit Sub
        End If

        msg = "Selected Assemblies: " & String.Join(", ", lstAssemblies.ToArray)
        msg &= Environment.NewLine & Environment.NewLine
        msg &= $"Are you sure you want to deploy the above assemblies to the following client: {client}?"
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
                row.Item("SELECTED") = "0"
            Next

            Dim client As String = cmbClient.Text

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
                ps.AddScript($"Deploy-Assemblies -deployToEnvironments (""{client}"") -assembliesToDeploy (""{String.Join(""",""", selectedAssemblies)}"")")

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

            ASCMAIN1.Progress("", "")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Private Sub btnCreateRelease_Click(sender As Object, e As EventArgs) Handles btnCreateRelease.Click


        Dim releaseDirectory As String = String.Empty
        If Not ValidateSelections(True, True, releaseDirectory) Then
            Exit Sub
        End If

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
            ps.AddScript($"Create-Release-Folder")

            Dim results As Collection(Of PSObject) = ps.Invoke()

            For Each result As PSObject In results
                Dim success As Boolean = Convert.ToBoolean(result.Properties("Success").Value)
                Dim message As String = Convert.ToString(result.Properties("Message").Value)

                If success Then
                    Dim release As String = Convert.ToString(result.Properties("Release").Value)
                    SetReleaseFolders(True)
                End If

                MessageBox.Show(message, "Create Release", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Next

            If ps.HadErrors Then
                For Each errorStream As ErrorRecord In ps.Streams.Error
                    If errorStream.ErrorDetails IsNot Nothing Then
                        MessageBox.Show(errorStream.ErrorDetails.Message, "Create Release", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Next
            Else
                'Need to add code to verify what was deployed and display to screen
                'Probably need to add output to the deploy.ps1 script
            End If

            ps.Dispose()

        End Using
    End Sub

    Private Sub btnProdDeploy_Click(sender As Object, e As EventArgs) Handles btnProdDeploy.Click

        Dim releaseDirectory As String = String.Empty
        If Not ValidateSelections(False, True, releaseDirectory) Then
            Exit Sub
        End If

        If cmbReleases.Items.Count = 0 Then
            SetReleaseFolders(True)
        End If

        ' See if the user selected a release folder
        If cmbReleases.SelectedItem Is Nothing OrElse cmbReleases.Items.Count = 0 Then
            MessageBox.Show("You are required to select a release folder.", "Deploy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim rootDirectory As String = cmbClient.SelectedRow.Cells("IPADRESS_PROD").Text

        Dim emsg As String = "Do you want to use folder " & cmbReleases.SelectedItem & " to Update the Production Region?"
        If MessageBox.Show(emsg, "Deploy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Exit Sub
        End If

        Dim dirProd As String = rootDirectory
        If Not dirProd.EndsWith("\") Then
            dirProd &= "\"
        End If

        dirProd &= cmbReleases.SelectedItem
        Dim lstFiles As New List(Of String)
        For Each fileName As String In My.Computer.FileSystem.GetFiles(dirProd)
            fileName = My.Computer.FileSystem.GetName(fileName)
            lstFiles.Add(fileName)
        Next

        lstFiles.Sort()

        If lstFiles.Count = 0 Then
            MessageBox.Show("There are no files in the selected directory: " & cmbReleases.SelectedItem, "Deploy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        emsg = "Do you want to copy the following files to the Production Region?" & Environment.NewLine & Environment.NewLine
        emsg &= String.Join(Environment.NewLine, lstFiles.ToArray)
        If MessageBox.Show(emsg, "Deploy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Exit Sub
        End If

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
            ps.AddScript($"Deploy-Release -releaseFolder {cmbReleases.SelectedItem}")

            Dim results As Collection(Of PSObject) = ps.Invoke()

            Try
                For Each result As PSObject In results
                    Dim success As Boolean = Convert.ToBoolean(result.Properties("Success").Value)
                    Dim message As String = Convert.ToString(result.Properties("Message").Value)

                    MsgBox(message)
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            If ps.HadErrors Then
                For Each errorStream As ErrorRecord In ps.Streams.Error
                    MsgBox(errorStream.ErrorDetails.Message)
                Next
            Else
                'Need to add code to verify what was deployed and display to screen
                'Probably need to add output to the deploy.ps1 script
            End If

            ps.Dispose()

        End Using
    End Sub

    Private Sub cmbClient_ValueChanged(sender As Object, e As EventArgs) Handles cmbClient.ValueChanged
        SetReleaseFolders(False)
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
                releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_PROD").Text
                If releaseDirectory.Length = 0 Then
                    MessageBox.Show("The selected client does not have a Production Application Directory.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Function
                End If

                If Not My.Computer.FileSystem.DirectoryExists(releaseDirectory) Then
                    MessageBox.Show("The selected client's Production Application Directory cannot be located.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Function
                End If

                releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_TEST").Text
                If releaseDirectory.Length = 0 Then
                    MessageBox.Show("The selected client does not have a Test Application Directory.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Function
                End If

                If Not My.Computer.FileSystem.DirectoryExists(releaseDirectory) Then
                    MessageBox.Show("The selected client's Test Application Directory cannot be located.", "Validate Selections", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Function
                End If

                releaseDirectory = cmbClient.SelectedRow.Cells("IPADDRESS_PROD").Text

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

End Class
