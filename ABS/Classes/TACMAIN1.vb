Public Class TACMAIN1

    ' nSoftware License Keys
    '    Public nSoftwareZipkey As String = "315A4E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004A424A545848315458354B300000"
    Public nSoftwareZipkey As String = "315A4E46414431535542323032333033313352415331544531414D483134323600000000000000004E484D59345531420000415853545441344A504653500000"
    'Public nSoftwareIPWorksV9Key As String = "31504E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004B4857525953375A4A5A375A0000"
    Public nSoftwareIPWorksV9Key As String = "31504E46414431535542323032333033313352415331544531414D483134323600000000000000004E484D593455314200005959563339583230314535340000"
    Public nSoftwareftpkey As String = nSoftwareIPWorksV9Key
    Public nSoftwareipportkey As String = nSoftwareIPWorksV9Key
    Public nSoftwarepopkey As String = nSoftwareIPWorksV9Key
    Public nSoftwarehttpkey As String = nSoftwareIPWorksV9Key
    'Public nSoftwareInship As String = "42584E354141315355425241533154453345383933333331580000000000000000000000000000004A52344B5057583900003059573859305A4A545958520000"
    'Public nSoftwareInship As String = "42584E3541413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600003030584254504E57374345330000"
    Public nSoftwareInship As String = "42584E424144315355423230313930393238524153315445334538393333333158000000000000004D554637445A525A00005947563236374346554E50540000"
    'Public nSoftwareEncryptkey As String = "31454E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004E4B54534E383157353733320000"
    Public nSoftwareEncryptkey As String = "31454E46414431535542323032333033313352415331544531414D483134323600000000000000004E484D593455314200005045465654425334425956450000"
    'Public nSoftwaresftpkey As String = "31484E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D5446000044483650384E5454444B4D4B0000"
    Public nSoftwaresftpkey As String = "31484E46414431535542323032333033313352415331544531414D483134323600000000000000004E484D593455314200005A4E3737415244503535394E0000"
    'Public nSoftwareEncryptionkey As String = "31454E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004E4B54534E383157353733320000"
    Public nSoftwareEncryptionkey As String = "31454E46414431535542323032333033313352415331544531414D483134323600000000000000004E484D593455314200005045465654425334425956450000"
    'Public nSoftwareInPay As String = "42504E364141315355425241533154453345383933333331580000000000000000000000000000004532594A58424252000032455A374B3543414D5841330000"
    'Public nSoftwareInPay As String = "42504E46414431535542323032313033313352415331544531414D483134323600000000000000003034484459543145000050484B4E42484A464B4A53530000"
    Public e4DPayments As String = "44504E4641414E5852464145504130303536000000000000000000000000000000000000000000004E484D5934553142000058314B365943314434424B530000"

    Public Overridable Sub Site_Specific_Settings()

    End Sub

    Public Overridable Sub Get_Column_Expression_Exceptions(ByVal FORM_NAME As String, ByVal DATA_SOURCE As String, ByVal COLUMN_NAME As String, ByRef sql_SELECT_col As String) ' , ByRef sql_GROUP_BY_col As String)

    End Sub

    Public Overridable Function Get_Code_SQL_X(ByVal FORM_NAME As String, ByVal COLUMN_NAME As String, ByRef GROUP_KEY As String) As String
        Return Nothing
    End Function

    Public Overridable Sub Write_Group_Record_X(ByVal GROUP_KEY As String, ByVal COLUMN_NAME As String, ByVal GROUP_CODEs As ArrayList, ByVal GROUP_DESCs As ArrayList)

    End Sub

    Public Overridable Function CodeValues(ByVal TABLE_COLUMN As String) As Dictionary(Of String, String)
        Return Nothing
    End Function

    Public Overridable Function Send_email(ByVal frmASFBASE0 As ASFBASE0, _
                                 ByVal EMAIL_ADDRESSs As Dictionary(Of String, String), _
                                 ByVal ATTACHMENTs As Dictionary(Of String, String), _
                                 ByVal SUBJECT As String, _
                                 ByVal EMAIL_KEY As String, _
                                 Optional ByVal auto_send As Boolean = False, _
                                 Optional SEND_CC_to_USER_ID As Boolean = False, _
                                 Optional ENTITY_KEY As String = "", _
                                 Optional ENTITY_NAME As String = "", _
                                 Optional ENTITY_CAPTION As String = "", _
                                 Optional EMAIL_BODY As String = "") As String
        Return Nothing
    End Function

    Public Overridable Sub Application_Initialization()

    End Sub

    Public Overridable Sub Maintain_Contacts(ByVal frmASFBASE1 As ASFBASE1, _
                                           ByVal CONTACT_ENTITY_TABLE As String, _
                                           ByVal CONTACT_ENTITY_KEY As String, _
                                           ByVal CONTACT_ENTITY_NAME As String)

    End Sub

    Public Overridable Function Custom_sqlwhere( _
    ByVal sqlwhere As String, _
    ByVal grd As UltraWinGrid.UltraGrid, _
    ByVal COLUMN_NAME As String) As String
        Return sqlwhere
    End Function

    Public Sub Record_Event( _
        ByVal TABLE_NAME As String, _
        ByVal TABLE_KEY As String, _
        ByVal INIT_DATE As Date, _
        ByVal INIT_OPER As String, _
        ByVal EVENT_TYPE As String, _
        ByVal EVENT_DESC As String, _
        Optional ByVal EVENT_KEY As String = "", _
        Optional FORM_NAME As String = "")

        If FORM_NAME = "" Then
            FORM_NAME = ASCMAIN1.ActiveForm.Name
        End If

        Dim SELECTION_NO As String = ""
        Dim XNO As String = ""

        If ASCMAIN1.ActiveForm IsNot Nothing Then
            SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
            XNO = ASCMAIN1.ActiveForm.XNO
        End If

        ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME, SESSION_NO, SELECTION_NO, XNO) " _
                             & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8,:PARM9,:PARM10,:PARM11)", _
                             "VVDVVVVVVVV", _
                             New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME, ASCMAIN1.SESSION_NO, SELECTION_NO, XNO})

    End Sub

#Region "Grid Layout Saving"
    Public Overridable Sub SaveGridLayout(ByRef frm As ASFBASE0, ByRef grd As UltraWinGrid.UltraGrid)
    End Sub

    Public Overridable Sub loadGridLayout(ByRef frm As ASFBASE0, ByRef grd As UltraWinGrid.UltraGrid)
    End Sub
#End Region

End Class
