Public Class WHCFACT1
    Public Shared Function CreateWhcClass(ByVal className As String, ByVal gunEnvironment As GunEnvironment) As WHCRF000
        Select Case className
            Case "WHCRF001"
                Return New WHCRF001(gunEnvironment)
            Case "WHCRF002"
                Return New WHCRF002(gunEnvironment)
            Case "WHCRF003"
                Return New WHCRF003(gunEnvironment)
            Case "WHCRF004"
                Return New WHCRF004(gunEnvironment)
            Case "WHCRF005"
                Return New WHCRF005(gunEnvironment)
            Case "WHCRF006"
                Return New WHCRF006(gunEnvironment)
            Case "WHCRF007"
                Return New WHCRF007(gunEnvironment)
            Case "WHCRF008"
                Return New WHCRF008(gunEnvironment)
            Case "WHCRF009"
                Return New WHCRF009(gunEnvironment)
            Case "WHCRF010"
                Return New WHCRF010(gunEnvironment)
            Case "WHCRF011"
                Return New WHCRF011(gunEnvironment)
            Case "WHCRF012"
                Return New WHCRF012(gunEnvironment)
            Case "WHCRF013"
                Return New WHCRF013(gunEnvironment)
            Case "WHCRF014"
                Return New WHCRF014(gunEnvironment)
            Case "WHCRF015"
                Return New WHCRF015(gunEnvironment)
            Case "WHCRF016"
                Return New WHCRF016(gunEnvironment)
            Case "WHCRF017"
                Return New WHCRF017(gunEnvironment)
            Case "WHCRF018"
                Return New WHCRF018(gunEnvironment)
            Case "WHCRF019"
                Return New WHCRF019(gunEnvironment)
            Case "WHCRF020"
                Return New WHCRF020(gunEnvironment)
            Case "WHCRF021"
                Return New WHCRF021(gunEnvironment)
            Case "WHCRF022"
                Return New WHCRF022(gunEnvironment)
            Case "WHCRF023"
                Return New WHCRF023(gunEnvironment)
            Case "WHCRF024"
                Return New WHCRF024(gunEnvironment)
        End Select
        Return Nothing
    End Function

End Class
