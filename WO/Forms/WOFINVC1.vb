Imports System.Drawing

Public Class WOFINVC1

#Region "Declarations"
    Dim Virtual_Location() As String
    Dim WHTSTYLX As String
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        WHTSTYLX = TAC.WHCMAIN1.Prepare_WHTSTYLX("", "", True)
        ASCMAIN1.sql = String.Format("Truncate table {0}", WHTSTYLX)
        ASCDATA1.ExecuteSQL()

        Dim new_Index As Integer = 0
        For Each VLs As String In New String() _
            {"00006A", "00000A", "00003A", "00004A", "00005A", "00007A", "00008A"}
            ReDim Preserve Virtual_Location(new_Index)
            Virtual_Location(new_Index) = VLs
            new_Index += 1
        Next

        Get_PARM("WHTLOCB1")
        With dst
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(" SELECT")
            SQLS.AppendLine(" B1.LOCATION_CODE,")
            SQLS.AppendLine(" B1.BAR_CODE,")
            SQLS.AppendLine(" (B1.STYLE_CODE || B1.COLOR_CODE) AS ITEM_CODE,")
            SQLS.AppendLine(" B1.STYLE_CODE,")
            SQLS.AppendLine(" B1.COLOR_CODE,")
            SQLS.AppendLine(" NVL(B2.STYLE_COUNT,SUM(B1.LOCATION_QTY)) STYLE_COUNT,")
            SQLS.AppendLine(" NVL(B2.CARTON_PACK_QTY,SUM(B1.LOCATION_QTY)) CARTON_PACK_QTY,")
            SQLS.AppendLine(" SUM(B1.LOCATION_QTY) AS LOCATION_QTY,")
            SQLS.AppendLine(" NVL((TRUNC(((SUM(B1.LOCATION_QTY) / B2.CARTON_PACK_QTY) * 1000),3) /1000),1) AS CASES")
            SQLS.AppendLine(" FROM WHTLOCB1 B1")
            'SQLS.AppendLine(" INNER Join")
            SQLS.AppendLine(" LEFT OUTER JOIN")
            SQLS.AppendLine(" (SELECT BAR_CODE,")
            SQLS.AppendLine(" COUNT((STYLE_CODE||COLOR_CODE)) AS STYLE_COUNT,")
            SQLS.AppendLine(" SUM(CARTON_PACK_QTY) AS CARTON_PACK_QTY")
            SQLS.AppendLine(" FROM WHTBARC2")
            SQLS.AppendLine(" GROUP BY BAR_CODE")
            SQLS.AppendLine(" ) B2")
            SQLS.AppendLine(" ON B1.BAR_CODE = B2.BAR_CODE")
            'SQLS.AppendLine(" WHERE B1.LOCATION_CODE NOT IN ('00006A', '00000A', '00003A', '00004A', '00005A', '00007A', '00008A')")
            SQLS.AppendLine(" GROUP BY")
            SQLS.AppendLine(" B1.LOCATION_CODE,")
            SQLS.AppendLine(" B1.BAR_CODE,")
            SQLS.AppendLine(" (B1.STYLE_CODE || B1.COLOR_CODE),")
            SQLS.AppendLine(" B1.STYLE_CODE,")
            SQLS.AppendLine(" B1.COLOR_CODE,")
            SQLS.AppendLine(" B2.STYLE_COUNT,")
            SQLS.AppendLine(" B2.CARTON_PACK_QTY")
            SQLS.AppendLine(" HAVING(SUM(B1.LOCATION_QTY) <> 0)")
            SQLS.AppendLine(" ORDER BY B1.LOCATION_CODE,")
            SQLS.AppendLine(" B1.STYLE_CODE,")
            SQLS.AppendLine(" B1.COLOR_CODE")
            ASCMAIN1.sql = SQLS.ToString()
            Create_TDA(.Tables.Add, "WHTINVCT", "**", 0, False, "", 6)
            Create_TDA(.Tables.Add, "WHTINVC1", "**", 0, True, "", 6)

            SQLS.Length = 0
            SQLS.AppendLine(" SELECT")
            SQLS.AppendLine(" LOCATION_CODE,")
            SQLS.AppendLine(" BAR_CODE,")
            SQLS.AppendLine(String.Format(" '{0}' AS ITEM_CODE,", New String(" ", 20)))
            SQLS.AppendLine(" 0 AS CART_NO,")
            SQLS.AppendLine(" 0 AS CART_QTY")
            SQLS.AppendLine(" FROM WHTLOCB1")
            SQLS.AppendLine(" WHERE ROWNUM = 0")
            ASCMAIN1.sql = SQLS.ToString()
            Create_TDA(.Tables.Add, "WHTINVC2", "**", 0, True, "", 4)

            ASCMAIN1.sql = "SELECT * FROM WHTPPKM1"
            Create_TDA(.Tables.Add, "WHTPPKX1", "**", 0, True, "", 1)

            ASCMAIN1.sql = "SELECT * FROM WHTPPKM2"
            Create_TDA(.Tables.Add, "WHTPPKX2", "**", 0, True, "", 3)

        End With

        grdWHTINVC1.DataSource = dst.Tables("WHTINVC1")
        grdWHTINVC2.DataSource = dst.Tables("WHTINVC2")
        grdWHTPPKX2.DataSource = dst.Tables("WHTPPKX2")

        Create_Summary(grdWHTINVC2, "CART_NO", "Count")
        Create_Summary(grdWHTINVC2, "CART_QTY", "Sum")

        With grdWHTINVC1
            .DisplayLayout.UseFixedHeaders = False
            With .DisplayLayout.Bands(0)
                .Columns("STYLE_COUNT").Format = "#,##0"
                .Columns("CARTON_PACK_QTY").Format = "#,##0"
                .Columns("LOCATION_QTY").Format = "#,##0.00"
                .Columns("CASES").Format = "#,##0.00"
            End With
        End With

        With grdWHTINVC2
            .DisplayLayout.UseFixedHeaders = False
            With .DisplayLayout.Bands(0)
                .Columns("CART_NO").Format = "#,##0"
                .Columns("CART_QTY").Format = "#,##0"
            End With
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Load"
            Case Else
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Load"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)
            Case "Done"
                Mode_Settings(False)
            Case "Transmit"
                Update_Record(True)
                Mode_Settings(False)
            Case Else
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If Not ScreenMode Then
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTINVC1", "WHTINVC2", "WHTPPKX1", "WHTPPKX2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Fill_Records("WHTINVCT")
        Fill_Records("WHTPPKX1")
        Fill_Records("WHTPPKX2")

        EnforceConstraints(True)

        For Each row As DataRow In TBLs("WHTINVCT").Rows
            Dim nRow As DataRow = dst.Tables("WHTINVC1").NewRow
            nRow.Item("LOCATION_CODE") = row.Item("LOCATION_CODE")
            nRow.Item("BAR_CODE") = row.Item("BAR_CODE")
            nRow.Item("ITEM_CODE") = row.Item("ITEM_CODE")
            nRow.Item("STYLE_CODE") = row.Item("STYLE_CODE")
            nRow.Item("COLOR_CODE") = row.Item("COLOR_CODE")
            nRow.Item("STYLE_COUNT") = row.Item("STYLE_COUNT")
            nRow.Item("CARTON_PACK_QTY") = row.Item("CARTON_PACK_QTY")
            nRow.Item("LOCATION_QTY") = row.Item("LOCATION_QTY")
            nRow.Item("CASES") = row.Item("CASES")
            dst.Tables("WHTINVC1").Rows.Add(nRow)
        Next

        BuildCartonRecords()

        Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Summary()

        Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Doing Things...")

        Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record(ShowMsg As Boolean)
        BeginTrans()

        Dim sqlx As String = "" 'String.Format("WKORDER_NO = '{0}'", WKORDER_NO)

        DeleteRecords()

        Update_Record_TDA("WHTINVC1", sqlx)
        Update_Record_TDA("WHTINVC2", sqlx)
        Update_Record_TDA("WHTPPKX1", sqlx)
        Update_Record_TDA("WHTPPKX2", sqlx)

        ASCMAIN1.sql = "INSERT INTO WHTPPKM1 SELECT * FROM WHTPPKX1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "INSERT INTO WHTPPKM2 SELECT * FROM WHTPPKX2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        SendDataToADS()

        If ShowMsg Then
            CommitTrans("Update Complete")
        Else
            CommitTrans("")
        End If
    End Sub

    Sub SendDataToADS()
        ASCMAIN1.sql = "DELETE FROM ADS.WHTINVC1@ADSIIS"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "DELETE FROM ADS.WHTINVC2@ADSIIS"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "DELETE FROM ADS.WHTPPKM1@ADSIIS"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "DELETE FROM ADS.WHTPPKM2@ADSIIS"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "INSERT INTO ADS.WHTINVC1@ADSIIS SELECT * FROM WHTINVC1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "INSERT INTO ADS.WHTINVC2@ADSIIS SELECT * FROM WHTINVC2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "INSERT INTO ADS.WHTPPKM1@ADSIIS SELECT * FROM WHTPPKM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "INSERT INTO ADS.WHTPPKM2@ADSIIS SELECT * FROM WHTPPKM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        '--- Special Area that WZ work with to update misc styles.
        ASCMAIN1.sql = "select distinct STYLE_CODE, COLOR_CODE from WHTLOCB1"
        Dim ICTSTAT2 As String = ASCMAIN1.Temp_Table()

        TAC.WHCMAIN1.Prepare_WHTSTYLX(WHTSTYLX, "NJ", , ICTSTAT2)

        ASCMAIN1.sql = "Update WHTSTYLX Set STATUS = '1',LP_XNO = 'NJ'" _
            & " where LP_CODE = 'NJ' and STATUS = '0'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ADS.WHTSTYLX@ADSIIS where LP_CODE = 'NJ'" _
            & " and (ITEM_CODE) in (Select ITEM_CODE from WHTSTYLX where LP_XNO = 'NJ')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ADS.WHTSTYLX@ADSIIS" _
            & " Select * from WHTSTYLX where LP_CODE = 'NJ' and LP_XNO = 'NJ'"
        ASCDATA1.ExecuteSQL()
        '--- End special area
    End Sub

    Sub DeleteRecords()
        ASCMAIN1.sql = "DELETE FROM WHTINVC1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "DELETE FROM WHTINVC2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "DELETE FROM WHTPPKX1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        ASCMAIN1.sql = "DELETE FROM WHTPPKX2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Load"
                Click_Command("Load")
            Case Else
        End Select

        Return Nothing
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ARTCUST1"
            E.COLUMN_NAME = "CUST_CODE"
            E.CODE_VALUE = Absx1.txtFor("CUST_CODE").Text
            E.DESC_VALUE = "Customer"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTINVC1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

#End Region

#Region "ABSColumn Controls"
#End Region

#Region "grdWHTINVC1"

    Private Sub grdWOTINVC1_AfterRowActivate(sender As System.Object, e As System.EventArgs) Handles grdWHTINVC1.AfterRowActivate
        If ScreenMode Then
            Dim dvw As DataView = DirectCast(grdWHTINVC2.DataSource, DataTable).DefaultView
            dvw.RowFilter = String.Format("LOCATION_CODE = '{0}' and BAR_CODE = '{1}'", grdWHTINVC1.ActiveRow.Cells("LOCATION_CODE").Value, grdWHTINVC1.ActiveRow.Cells("BAR_CODE").Value)
            grdWHTINVC2.Text = "Details for BarCode " & grdWHTINVC1.ActiveRow.Cells("BAR_CODE").Value
        End If
    End Sub
#End Region

#Region "grdWHTINVC2"
    Private Sub grdWHTINVC2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTINVC2.AfterRowActivate
        If ScreenMode Then
            Dim dvw As DataView = DirectCast(grdWHTPPKX2.DataSource, DataTable).DefaultView
            dvw.RowFilter = String.Format("PPK_CODE = '{0}'", grdWHTINVC2.ActiveRow.Cells("ITEM_CODE").Value)
            'grdWHTINVC2.Text = "Details for BarCode " & grdWHTINVC1.ActiveRow.Cells("BAR_CODE").Value
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub BuildCartonRecords()
        Dim NEXT_CASE_NO As Integer = 1
        dst.Tables("WHTINVC2").Clear()
        Dim LOCATION_CODE_LAST As String = ""
        Dim BAR_CODE_LAST As String = ""
        For Each rowWHTINVC1 As DataRow In dst.Tables("WHTINVC1").Select("", "LOCATION_CODE, BAR_CODE, ITEM_CODE")
            If Virtual_Location.Contains(rowWHTINVC1.Item("LOCATION_CODE")) Then
                If rowWHTINVC1.Item("LOCATION_CODE") <> LOCATION_CODE_LAST Then
                    Dim LOCATION_QTY As Integer = rowWHTINVC1.Item("LOCATION_QTY")
                    For Cases As Integer = 1 To RoundUp(rowWHTINVC1.Item("CASES"))
                        Dim rowWHTINVC2 As DataRow = dst.Tables("WHTINVC2").NewRow
                        rowWHTINVC2.Item("LOCATION_CODE") = rowWHTINVC1.Item("LOCATION_CODE")
                        rowWHTINVC2.Item("BAR_CODE") = rowWHTINVC1.Item("BAR_CODE")
                        rowWHTINVC2.Item("CART_NO") = NEXT_CASE_NO
                        rowWHTINVC2.Item("ITEM_CODE") = rowWHTINVC1.Item("ITEM_CODE")
                        If rowWHTINVC1.Item("CARTON_PACK_QTY") <= LOCATION_QTY Then
                            rowWHTINVC2.Item("CART_QTY") = rowWHTINVC1.Item("CARTON_PACK_QTY")
                        Else
                            rowWHTINVC2.Item("CART_QTY") = LOCATION_QTY
                        End If
                        dst.Tables("WHTINVC2").Rows.Add(rowWHTINVC2)
                        LOCATION_QTY -= rowWHTINVC2.Item("CART_QTY")
                        NEXT_CASE_NO += 1
                    Next
                End If
            Else
                If rowWHTINVC1.Item("STYLE_COUNT") = 1 Then
                    Dim LOCATION_QTY As Integer = rowWHTINVC1.Item("LOCATION_QTY")
                    For Cases As Integer = 1 To RoundUp(rowWHTINVC1.Item("CASES"))
                        Dim rowWHTINVC2 As DataRow = dst.Tables("WHTINVC2").NewRow
                        rowWHTINVC2.Item("LOCATION_CODE") = rowWHTINVC1.Item("LOCATION_CODE")
                        rowWHTINVC2.Item("BAR_CODE") = rowWHTINVC1.Item("BAR_CODE")
                        rowWHTINVC2.Item("CART_NO") = NEXT_CASE_NO
                        rowWHTINVC2.Item("ITEM_CODE") = rowWHTINVC1.Item("ITEM_CODE")
                        If rowWHTINVC1.Item("CARTON_PACK_QTY") <= LOCATION_QTY Then
                            rowWHTINVC2.Item("CART_QTY") = rowWHTINVC1.Item("CARTON_PACK_QTY")
                        Else
                            rowWHTINVC2.Item("CART_QTY") = LOCATION_QTY
                        End If
                        dst.Tables("WHTINVC2").Rows.Add(rowWHTINVC2)
                        LOCATION_QTY -= rowWHTINVC2.Item("CART_QTY")
                        NEXT_CASE_NO += 1
                    Next
                Else
                    If rowWHTINVC1.Item("LOCATION_CODE") <> LOCATION_CODE_LAST Or rowWHTINVC1.Item("BAR_CODE") <> BAR_CODE_LAST Then
                        LOCATION_CODE_LAST = rowWHTINVC1.Item("LOCATION_CODE")
                        BAR_CODE_LAST = rowWHTINVC1.Item("BAR_CODE")
                        CreateMultiCarton(LOCATION_CODE_LAST, BAR_CODE_LAST, NEXT_CASE_NO)
                    End If
                End If
            End If
        Next
    End Sub

    Private Function RoundUp(ByVal NumIn As Double) As Integer
        If Math.Truncate(NumIn) < NumIn Then
            Return Math.Truncate(NumIn) + 1
        Else
            Return CDbl(NumIn)
        End If
    End Function

    Private Sub CreateVirtualCarton(ByVal LOCATION_CODE As String, BAR_CODE As String, ByRef NextCase As Integer)
        'Dim FILTER As String = String.Format("LOCATION_CODE = '{0}' AND BAR_CODE = '{1}'", LOCATION_CODE, BAR_CODE)
        'Dim STYLE_COUNT As Integer = dst.Tables("WHTINVC1").Select(FILTER, "").Count()
        'Dim LOCATION_QTY_PER_LINE As Integer()
        'ReDim LOCATION_QTY_PER_LINE(STYLE_COUNT - 1)
        'Dim ITEMCODE As String = ""
        'Dim TOTALPACKEDCASE As Integer = 0
        'Dim CASES As Double = 0
        'Dim Counter As Integer = 0

        'For Each rowWHTINVC1 As DataRow In dst.Tables("WHTINVC1").Select(FILTER, "")
        '    If rowWHTINVC1.Item("CASES") > CASES Then
        '        CASES = rowWHTINVC1.Item("CASES")
        '    End If
        '    LOCATION_QTY_PER_LINE(Counter) = rowWHTINVC1.Item("LOCATION_QTY")
        '    Counter += 1
        'Next
        'CASES = RoundUp(CASES)
        'For carts As Integer = 1 To CASES
        '    TOTALPACKEDCASE = 0
        '    Dim PrePackList As New List(Of PrePackDetail)
        '    Dim cntWHTINVC1 As Integer = 0
        '    For Each rowWHTINVC1 As DataRow In dst.Tables("WHTINVC1").Select(FILTER, "")
        '        If LOCATION_QTY_PER_LINE(cntWHTINVC1) > 0 Then
        '            Dim PPKD As New PrePackDetail() With {.STYLE_CODE = rowWHTINVC1.Item("STYLE_CODE"), .COLOR_CODE = rowWHTINVC1.Item("COLOR_CODE")}
        '            If rowWHTINVC1.Item("CARTON_PACK_QTY") <= LOCATION_QTY_PER_LINE(cntWHTINVC1) Then
        '                PPKD.PPK_QTY = rowWHTINVC1.Item("CARTON_PACK_QTY")
        '                LOCATION_QTY_PER_LINE(cntWHTINVC1) -= rowWHTINVC1.Item("CARTON_PACK_QTY")
        '                TOTALPACKEDCASE += rowWHTINVC1.Item("CARTON_PACK_QTY")
        '            Else
        '                PPKD.PPK_QTY = LOCATION_QTY_PER_LINE(cntWHTINVC1)
        '                TOTALPACKEDCASE += LOCATION_QTY_PER_LINE(cntWHTINVC1)
        '                LOCATION_QTY_PER_LINE(cntWHTINVC1) = 0
        '            End If
        '            PrePackList.Add(PPKD)
        '        End If
        '        cntWHTINVC1 += 1
        '    Next

        '    ITEMCODE = MakeUsePrePack(PrePackList)

        '    Dim rowWHTINVC2 As DataRow = dst.Tables("WHTINVC2").NewRow
        '    rowWHTINVC2.Item("LOCATION_CODE") = LOCATION_CODE
        '    rowWHTINVC2.Item("BAR_CODE") = BAR_CODE
        '    rowWHTINVC2.Item("CART_NO") = NextCase
        '    rowWHTINVC2.Item("ITEM_CODE") = ITEMCODE
        '    rowWHTINVC2.Item("CART_QTY") = TOTALPACKEDCASE
        '    dst.Tables("WHTINVC2").Rows.Add(rowWHTINVC2)
        '    NextCase += 1
        'Next
    End Sub

    Private Sub CreateMultiCarton(ByVal LOCATION_CODE As String, BAR_CODE As String, ByRef NextCase As Integer)
        Dim FILTER As String = String.Format("LOCATION_CODE = '{0}' AND BAR_CODE = '{1}'", LOCATION_CODE, BAR_CODE)
        Dim STYLE_COUNT As Integer = dst.Tables("WHTINVC1").Select(FILTER, "").Count()
        Dim LOCATION_QTY_PER_LINE As Integer()
        ReDim LOCATION_QTY_PER_LINE(STYLE_COUNT - 1)
        Dim ITEMCODE As String = ""
        Dim TOTALPACKEDCASE As Integer = 0
        Dim CASES As Double = 0
        Dim Counter As Integer = 0

        For Each rowWHTINVC1 As DataRow In dst.Tables("WHTINVC1").Select(FILTER, "")
            If rowWHTINVC1.Item("CASES") > CASES Then
                CASES = rowWHTINVC1.Item("CASES")
            End If
            LOCATION_QTY_PER_LINE(Counter) = rowWHTINVC1.Item("LOCATION_QTY")
            Counter += 1
        Next
        CASES = RoundUp(CASES)
        For carts As Integer = 1 To CASES
            TOTALPACKEDCASE = 0
            Dim PrePackList As New List(Of PrePackDetail)
            Dim cntWHTINVC1 As Integer = 0
            For Each rowWHTINVC1 As DataRow In dst.Tables("WHTINVC1").Select(FILTER, "")
                If LOCATION_QTY_PER_LINE(cntWHTINVC1) > 0 Then
                    Dim PPKD As New PrePackDetail() With {.STYLE_CODE = rowWHTINVC1.Item("STYLE_CODE"), .COLOR_CODE = rowWHTINVC1.Item("COLOR_CODE")}
                    If rowWHTINVC1.Item("CARTON_PACK_QTY") <= LOCATION_QTY_PER_LINE(cntWHTINVC1) Then
                        PPKD.PPK_QTY = rowWHTINVC1.Item("CARTON_PACK_QTY")
                        LOCATION_QTY_PER_LINE(cntWHTINVC1) -= rowWHTINVC1.Item("CARTON_PACK_QTY")
                        TOTALPACKEDCASE += rowWHTINVC1.Item("CARTON_PACK_QTY")
                    Else
                        PPKD.PPK_QTY = LOCATION_QTY_PER_LINE(cntWHTINVC1)
                        TOTALPACKEDCASE += LOCATION_QTY_PER_LINE(cntWHTINVC1)
                        LOCATION_QTY_PER_LINE(cntWHTINVC1) = 0
                    End If
                    PrePackList.Add(PPKD)
                End If
                cntWHTINVC1 += 1
            Next

            ITEMCODE = MakeUsePrePack(PrePackList)

            Dim rowWHTINVC2 As DataRow = dst.Tables("WHTINVC2").NewRow
            rowWHTINVC2.Item("LOCATION_CODE") = LOCATION_CODE
            rowWHTINVC2.Item("BAR_CODE") = BAR_CODE
            rowWHTINVC2.Item("CART_NO") = NextCase
            rowWHTINVC2.Item("ITEM_CODE") = ITEMCODE
            rowWHTINVC2.Item("CART_QTY") = TOTALPACKEDCASE
            dst.Tables("WHTINVC2").Rows.Add(rowWHTINVC2)
            NextCase += 1
        Next
    End Sub

    Private Function MakeUsePrePack(ByVal PREPACKIN As List(Of PrePackDetail)) As String
        Dim PREPACKFOUND As Boolean = False
        Dim PREPACKUSED As String = ""
        For Each rowWHTPPKX1 As DataRow In dst.Tables("WHTPPKX1").Select()
            Dim ISLINEINLIST As Boolean = True
            For Each rowWHTPPKX2 As DataRow In dst.Tables("WHTPPKX2").Select(String.Format("PPK_CODE = '{0}'", rowWHTPPKX1.Item("PPK_CODE")))
                Dim PREPACKSEARCH As New PrePackDetail() With {.STYLE_CODE = rowWHTPPKX2.Item("STYLE_CODE"), .COLOR_CODE = rowWHTPPKX2.Item("COLOR_CODE"), .PPK_QTY = rowWHTPPKX2.Item("PPK_QTY")}
                If Not PrePackContains(PREPACKIN, PREPACKSEARCH) Then
                    ISLINEINLIST = False
                    Exit For
                End If
            Next
            If ISLINEINLIST Then
                PREPACKFOUND = True
                PREPACKUSED = rowWHTPPKX1.Item("PPK_CODE")
            End If
        Next
        If PREPACKFOUND Then
            Return PREPACKUSED
        Else
            Dim PPK_CODE As String = ASCMAIN1.Next_Control_No("PPK_CODE") & "PPK"
            PPK_CODE = Mid(PPK_CODE, 2)

            Dim rowWHTPPKX1_New As DataRow = dst.Tables("WHTPPKX1").NewRow
            rowWHTPPKX1_New.Item("PPK_CODE") = PPK_CODE
            rowWHTPPKX1_New.Item("PPK_DESC") = "Created In Conversion"
            rowWHTPPKX1_New.Item("INIT_DATE") = Now()
            rowWHTPPKX1_New.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTPPKX1_New.Item("LAST_DATE") = Now()
            rowWHTPPKX1_New.Item("LAST_OPER") = ASCMAIN1.USER_ID
            dst.Tables("WHTPPKX1").Rows.Add(rowWHTPPKX1_New)

            For Each P As PrePackDetail In PREPACKIN
                Dim rowWHTPPKX2_New As DataRow = dst.Tables("WHTPPKX2").NewRow
                rowWHTPPKX2_New.Item("PPK_CODE") = PPK_CODE
                rowWHTPPKX2_New.Item("STYLE_CODE") = P.STYLE_CODE
                rowWHTPPKX2_New.Item("COLOR_CODE") = P.COLOR_CODE
                rowWHTPPKX2_New.Item("PPK_QTY") = P.PPK_QTY
                dst.Tables("WHTPPKX2").Rows.Add(rowWHTPPKX2_New)
            Next

            Return PPK_CODE
        End If
    End Function

    Private Function PrePackContains(ByVal PrePack As List(Of PrePackDetail), ByVal Search As PrePackDetail) As Boolean
        Dim RETVAL As Boolean = False
        For Each DETAIL As PrePackDetail In PrePack
            If DETAIL.STYLE_CODE = Search.STYLE_CODE And DETAIL.COLOR_CODE = Search.COLOR_CODE And DETAIL.PPK_QTY = Search.PPK_QTY Then
                RETVAL = True
            End If
        Next
        Return RETVAL
    End Function

#End Region

End Class
