Imports System
Imports System.IO
Imports Microsoft.Build.Utilities
Imports Microsoft.Build.Framework

Namespace ABSTasks

    Public Class IncrementBuildNumber
        Inherits Task

        Dim m_fileName As String    ' AssemblyInfo.vb file
        Dim m_buildNumber As Integer    ' Build number based on current month/day
        Dim m_revisionNumber As Integer ' Revision number, increments from 0 each day


        ''' <summary>
        ''' MSBuild entry point into this task.
        ''' </summary>
        Public Overrides Function Execute() As Boolean

            Try
                IncrementNumbers()
                Log.LogMessage(MessageImportance.Normal, "Build {0}.{1}", m_buildNumber, m_revisionNumber)
            Catch ex As Exception
                ' Log Failure
                Log.LogErrorFromException(ex)
                Log.LogMessage(MessageImportance.High, "Failed to increment")
                Return False
            End Try
            Return True

        End Function

        <Required()> _
        Public Property File() As String
            Get
                Return m_fileName
            End Get
            Set(ByVal value As String)
                m_fileName = value
            End Set
        End Property


        <Output()> _
        Public Property BuildNumber() As Integer
            Get
                Return m_buildNumber
            End Get
            Set(ByVal value As Integer)
                m_buildNumber = value
            End Set
        End Property

        <Output()> _
        Public Property RevisionNumber() As Integer
            Get
                Return m_revisionNumber
            End Get
            Set(ByVal value As Integer)
                m_revisionNumber = value
            End Set
        End Property

        Private Sub IncrementNumbers()
            Dim dDate As DateTime = DateTime.Now

            'm_buildNumber = (dDate.Year Mod 2000) * 10000 'The build number can't be larger than 65535?
            ' Set build number to current month and day, e.g. July 13 = 713, October 5 = 1005
            m_buildNumber += dDate.Month * 100
            m_buildNumber += dDate.Day

            ' Default build revision to 0
            m_revisionNumber = 0

            ' Update based on values in AssemblyInfo.vb
            If (System.IO.File.Exists(m_fileName)) Then

                Dim AssemblyInfo() As String = System.IO.File.ReadAllLines(m_fileName)
                Dim previousNumbers As String()
                Dim vtext As String = "1.0.0.0"

                For i As Integer = 0 To AssemblyInfo.Length - 1
                    'When we find the line containing the Assembly Version...
                    If (AssemblyInfo(i).StartsWith("<Assembly: AssemblyVersion")) Then
                        previousNumbers = AssemblyInfo(i).Split("""")(1).Split(".")
                        If m_buildNumber = Integer.Parse(previousNumbers(2)) Then
                            m_revisionNumber = Integer.Parse(previousNumbers(3)) + 1
                        End If
                        vtext = previousNumbers(0) & "." & previousNumbers(1) & "." _
                        & m_buildNumber.ToString() & "." & m_revisionNumber.ToString
                        'Replace old version string with new one
                        AssemblyInfo(i) = "<Assembly: AssemblyVersion(""" _
                        & vtext & """)> "
                    End If
                Next

                System.IO.File.WriteAllLines(m_fileName, AssemblyInfo)
                System.IO.File.WriteAllText("version.txt", "Version " & vtext)

            End If

        End Sub

    End Class

End Namespace