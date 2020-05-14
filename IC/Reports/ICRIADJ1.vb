Imports System.Math

Public Class ICRIADJ1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Shadows SUBT As String = ""

    Dim ICTIADJ1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
        Absx1.optFor("RANGE").CheckedIndex = 2

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        SUBT = ""

        Dim sqlw As String = IIf(MENU_ITEM_OBJECT = "ICRIADJ1", "NVL(ICTIADJ1.JOURNAL_IND,'0') = '0'", "NVL(ICTIADJ1.REGISTER_IND,'0') = '0'")

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Adjustments Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Adjustments Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "ICTIADJ1.ADJ_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Adjustments Posted in " & xRYP0_legend
            Else
                SUBT = "Adjustments Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = "ICTIADJ1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If

        If RWU = "R" Then
            If MENU_ITEM_PP = "CJ" Then
                ' sqlw = "ICTIADJ1.OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "'"
                RWU = "N" ' WHEN RUN FROM THE COSTING MENU THIS REPORT SHOWS COSTS, BUT NO UPDATE REQUIRED
            Else
                'If MENU_ITEM_OBJECT = "ICRIADJ1" And ASCMAIN1.EOM <> "1" Then
                '    RWU = "N"
                'End If
            End If
        End If

        Prepare_dst(True, sqlw)

        Check_if_Empty("ICTIADJ1")
    End Sub

    Public Overrides Sub Print_Report()
        RPT = IIf(MENU_ITEM_FORM = "", MENU_ITEM_OBJECT, MENU_ITEM_FORM)
        Generate_Report(RPT, , SUBT)

        'If MENU_ITEM_PP = "CJ" Then
        '    Print_GL()
        'End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()

        If MENU_ITEM_PP = "CJ" Then
        Else
            Dim sql As String = "Update ICTIADJ1 " _
           & IIf(MENU_ITEM_OBJECT = "ICRIADJ1", _
                 " Set JOURNAL_IND = :PARM1, JOURNAL_XNO = :PARM2", _
                 " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2") _
           & " where ADJ_NO in (Select ADJ_NO from " & ICTIADJ1 & " )"
            ASCDATA1.ExecuteSQL(sql, "VV", New Object() {"1", MyBase.XNO})
        End If
       
        'If MENU_ITEM_PP = "CJ" Then
        '    GL_Update()
        'End If
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"
        ASCMAIN1.sql = "Select * from ICTIADJ1 where " & sqlw
        ICTIADJ1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTIADJ1 & " Add Primary Key (ADJ_NO)")

        ASCMAIN1.sql = "Select ICTIADJ1.*,ICTREAS1.REASON_DESC " _
            & " from " & ICTIADJ1 & " ICTIADJ1,ICTREAS1 " _
            & " where ICTREAS1.REASON_CODE (+) = ICTIADJ1.REASON_CODE " _
            & "   and " & sqlw
        Create_TDA(dst.Tables.Add, "ICTIADJ1", "**", 0)

        ASCMAIN1.sql = "Select ICTIADJ2.*, ICTSTYL1.STYLE_DESC " _
            & " from ICTIADJ2," & ICTIADJ1 & " ICTIADJ1, ICTSTYL1 " _
            & " where ICTIADJ2.ADJ_NO = ICTIADJ1.ADJ_NO" _
            & " and ICTSTYL1.STYLE_CODE = ICTIADJ2.STYLE_CODE"
        Create_TDA(dst.Tables.Add, "ICTIADJ2", "**", 0)

        ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC " _
            & " from ICTIADJ3," & ICTIADJ1 & " ICTIADJ1, GLTACCT1 " _
            & " where ICTIADJ3.ADJ_NO = ICTIADJ1.ADJ_NO" _
            & " and GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE"
        Create_TDA(dst.Tables.Add, "ICTIADJ3", "**", 0)

        'If MENU_ITEM_PP = "CJ" Then
        '    Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        'End If

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)
        Fill_Records("ICTIADJ1")
        Fill_Records("ICTIADJ2")
        Fill_Records("ICTIADJ3")
        'If MENU_ITEM_PP = "CJ" Then
        '    TAC.ICCMAIN1.Prepare_GL_Interface("ICIA", ICTIADJ1)
        'End If
        EnforceConstraints(True)
    End Sub
End Class