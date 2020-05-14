Public Class ASTPARMP

    Public Overrides Sub Proceed_PreReq_Special(eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Absx1.chkFor("AS_PARM_USE_ENCRYPTION").Enabled = False
    End Sub

    Public Overrides Sub Mode_Settings(tf As Boolean, Optional MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        Absx1.chkFor("AS_PARM_USE_ENCRYPTION").Enabled = False
    End Sub
End Class