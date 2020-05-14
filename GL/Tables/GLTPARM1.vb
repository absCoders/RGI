Public Class GLTPARM1

    Overrides Sub Proceed_Update_Special_Pre()

        ASCMAIN1.sql = "Delete from ASTCODE1 where TABLE_NAME = :PARM1 and COLUMN_NAME = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"GLTSEGM1", "ACCT_SEG_ID"})

        ASCMAIN1.sql = "Insert into ASTCODE1 Values (:PARM1,:PARM2,:PARM3,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"GLTSEGM1", "ACCT_SEG_ID", "2", Absx1.txtFor("GL_PARM_SEG2_DESC").Text})
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"GLTSEGM1", "ACCT_SEG_ID", "3", Absx1.txtFor("GL_PARM_SEG3_DESC").Text})
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"GLTSEGM1", "ACCT_SEG_ID", "4", Absx1.txtFor("GL_PARM_SEG4_DESC").Text})

    End Sub
End Class