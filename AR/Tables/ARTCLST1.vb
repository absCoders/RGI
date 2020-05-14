Public Class ARTCLST1

    Overrides Sub Show_Record_Special()

    End Sub

    Public Overrides Sub Proceed_Update_Special_Pre()
        MyBase.Proceed_Update_Special_Pre()
        INIT_LAST("ARTCLST1")
    End Sub

End Class