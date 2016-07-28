Attribute VB_Name = "WildeTool_Menards_850"
'Author:      Dharmendra,Pranayesh
'Date:        03 March 2016
'Customer:    Wilde Tool
'TP:          Menards
'Note:

Option Explicit
Dim oCon_Edi_Foundation_2008 As ADODB.Connection, oCon_Edi_2008 As ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim oRsComposite As ADODB.Recordset
Dim oXlApp As New Excel.Application

Dim ls_XlFileName As String
Dim li_XlLineCounter As Integer
Dim lb_Local_Server As Boolean

Dim ls_ISA15 As String, ls_BEG3 As String
Dim ld_RecIdLog As Double
Dim ls_CustomErrorMsg As String
Dim ls_ExeName As String
Dim ls_Developer As String
Dim ls_TransetNo As String
Dim ls_FileNameDownloaded As String

'Program execution starts from this point
Private Sub Main()
    'MANDATORY - DO NOT LEAVE IT
    ls_ExeName = "EDI_850_R_WildeTool_Menards.exe"
    ls_Developer = "Dharmendra"
    ls_TransetNo = "850"

    'Used for Local Testing or Server Executions
    'Plese Comments as per ur Choice for Debugging.
    'lb_Local_Server = False 'Local Testing
    lb_Local_Server = True  'Server Executions
    
    'Openning Database Connections
    Call s_OpenConnection

    'Entry into Error Log on Project Startup
    Call s_ErrorLogEntryOnProjectStart
    
    'Reading EDI file into Recordset
    Call s_ReadEdiSegmentsData
    
    'Generate Custom XML
    Call s_GenerateCustomXML

    'Generating Excel Sheet
    Call s_ExcelReport

    'Saving Records into Database
    Call s_AddToDataBase
    
    'Entry into Error Log on Project Completion
    Call s_ErrorLogEntryOnProjectEnd
    
    'Program execution ends at this point
    End
End Sub

'procedure for generating xml for Shuqualak
 Private Sub s_GenerateCustomXML()
 
    On Error GoTo ErrEntry
    ls_CustomErrorMsg = "s_GenerateCustomXML"

    Dim ldTotal As Double
    Dim liFreeFile As Integer
    Dim lsFilePath As String, lsTemp As String
    Dim lsPONumber As String, ls_ShiptoLocationId, ls_OrderDate, ls_Price As String, ls_PriceUOM As String
    Dim ls_PriceQTY As String, ls_Notes As String, ls_MSGNotes As String
    Dim ls_ReasonCode As String, ls_QTY As String
    Dim ls_PONO As String, lsN101 As String, ls_ICN As String
    
    'ISA
    s_RsFilter "[Segment]='ISA'"
    If Not oRs.EOF Then
        ls_ICN = Trim(oRs.Fields(13).Value)
    End If
    
    lsTemp = "<?xml version='1.0' encoding='utf-8'?>"
        lsTemp = lsTemp & "<orders>"
        lsTemp = lsTemp & "<user_fields></user_fields>"
        lsTemp = lsTemp & "<orders_envelope></orders_envelope>"
        
        lsTemp = lsTemp & "<order>"
            lsTemp = lsTemp & "<header>"
    
                'BEG
                s_RsFilter "[Segment]='BEG'"
                If Not oRs.EOF Then
                    ls_PONO = Trim(oRs.Fields(3).Value)
                    lsTemp = lsTemp & "<purchase_order>" & Trim(oRs.Fields(3).Value) & "</purchase_order>"
                    lsTemp = lsTemp & "<purpose_order_purpose_code>" & Trim(oRs.Fields(1).Value) & "</purpose_order_purpose_code>"
                    lsTemp = lsTemp & "<customer_number>850ALLWILDETOOL</customer_number>"
                    lsTemp = lsTemp & "<order_date>" & f_DateFormat(Trim(oRs.Fields(5).Value)) & "</order_date>"
                    ls_BEG3 = ls_PONO
                End If
                'DTM
                s_RsFilter "[Segment]='DTM' and [1]='038'"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<order_date>" & f_DateFormat(Trim(oRs.Fields(2).Value)) & "</order_date>"
                End If
                'DTM
                s_RsFilter "[Segment]='DTM' and [1]='106'"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<ship_date>" & f_DateFormat(Trim(oRs.Fields(2).Value)) & "</ship_date>"
                End If
                'N1
                s_RsFilter "([Segment]='N1_N1' and [30]='BT') OR ([Segment]='N1_N1' and [30]='BS')"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<bill_to_no>" & Trim(oRs.Fields(2).Value) & "</bill_to_no>"
                End If
                'N1
                s_RsFilter "([Segment]='N1_N1' and [30]='ST') OR ([Segment]='N1_N1' and [30]='BS')"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<ship_to_no>" & Trim(oRs.Fields(2).Value) & "</ship_to_no>"
                End If
                'N1
                s_RsFilter "([Segment]='N1_N1' and [30]='OB') or ([Segment]='N1_N1' and [30]='BY')"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<buyer_no>" & Trim(oRs.Fields(2).Value) & "</buyer_no>"
                End If
                'FOB
                s_RsFilter "[Segment]='FOB'"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<ship_via>" & f_GetDescription(Trim(oRs.Fields(2).Value), "LocationQualifier_309") & "</ship_via>"
                End If
                'ITD
                s_RsFilter "[Segment]='ITD'"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<terms>" & Trim(oRs.Fields(12).Value) & "</terms>"
                End If
                'FOB
                s_RsFilter "[Segment]='FOB'"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<ship_via_description>" & Trim(oRs.Fields(3).Value) & "</ship_via_description>"
                End If
                'BEG
                s_RsFilter "[Segment]='BEG'"
                If Not oRs.EOF Then
                    lsTemp = lsTemp & "<order_type>" & Trim(oRs.Fields(2).Value) & "</order_type>"
                End If
                'N1
                lsTemp = lsTemp & "<billing_address>"
                    s_RsFilter "([Segment]='N1_N1' and [30]='BT') Or ([Segment]='N3_N1' and [30]='BT') Or ([Segment]='N4_N1' and [30]='BT') Or ([Segment]='PER_N1' and [30]='BT') " _
                                & " or ([Segment]='N1_N1' and [30]='BS') Or ([Segment]='N3_N1' and [30]='BS') Or ([Segment]='N4_N1' and [30]='BS') Or ([Segment]='PER_N1' and [30]='BS') "
                    While Not oRs.EOF
                        If Trim(oRs.Fields(0).Value) = "N1_N1" Then
                            lsTemp = lsTemp & "<name>" & Trim(oRs.Fields(2).Value) & "</name>"
                        End If
                        If Trim(oRs.Fields(0).Value) = "N3_N1" Then
                            lsTemp = lsTemp & "<address_1>" & Trim(oRs.Fields(1).Value) & "</address_1>"
                        End If
                        If Trim(oRs.Fields(0).Value) = "N4_N1" Then
                            lsTemp = lsTemp & "<city>" & Trim(oRs.Fields(1).Value) & "</city>"
                            lsTemp = lsTemp & "<state>" & Trim(oRs.Fields(2).Value) & "</state>"
                            lsTemp = lsTemp & "<zip>" & Trim(oRs.Fields(3).Value) & "</zip>"
                        End If
                        If Trim(oRs.Fields(0).Value) = "PER_N1" Then
                            If Trim(oRs.Fields(3).Value) = "TE" Then
                                lsTemp = lsTemp & "<phone_number>" & Trim(oRs.Fields(4).Value) & "</phone_number>"
                            End If
                            If Trim(oRs.Fields(5).Value) = "FX" Then
                                lsTemp = lsTemp & "<fax_number>" & Trim(oRs.Fields(6).Value) & "</fax_number>"
                            End If
                        End If
                        oRs.MoveNext
                    Wend
                lsTemp = lsTemp & "</billing_address>"
        
                'N1
                lsTemp = lsTemp & "<ship_to_address>"
                    s_RsFilter "([Segment]='N1_N1' and [30]='ST') Or ([Segment]='N3_N1' and [30]='ST') Or ([Segment]='N4_N1' and [30]='ST') Or ([Segment]='PER_N1' and [30]='ST')" _
                                & " or ([Segment]='N1_N1' and [30]='BS') Or ([Segment]='N3_N1' and [30]='BS') Or ([Segment]='N4_N1' and [30]='BS') Or ([Segment]='PER_N1' and [30]='BS') "
                    While Not oRs.EOF
                        If Trim(oRs.Fields(0).Value) = "N1_N1" Then
                            lsTemp = lsTemp & "<name>" & Trim(oRs.Fields(2).Value) & "</name>"
                        End If
                        If Trim(oRs.Fields(0).Value) = "N3_N1" Then
                            lsTemp = lsTemp & "<address_1>" & Trim(oRs.Fields(1).Value) & "</address_1>"
                        End If
                        If Trim(oRs.Fields(0).Value) = "N4_N1" Then
                            lsTemp = lsTemp & "<city>" & Trim(oRs.Fields(1).Value) & "</city>"
                            lsTemp = lsTemp & "<state>" & Trim(oRs.Fields(2).Value) & "</state>"
                            lsTemp = lsTemp & "<zip>" & Trim(oRs.Fields(3).Value) & "</zip>"
                        End If
                        If Trim(oRs.Fields(0).Value) = "PER_N1" Then
                            If Trim(oRs.Fields(3).Value) = "TE" Then
                                lsTemp = lsTemp & "<phone_number>" & Trim(oRs.Fields(4).Value) & "</phone_number>"
                            End If
                            If Trim(oRs.Fields(5).Value) = "FX" Then
                                lsTemp = lsTemp & "<fax_number>" & Trim(oRs.Fields(6).Value) & "</fax_number>"
                            End If
                        End If
                        oRs.MoveNext
                    Wend
                lsTemp = lsTemp & "</ship_to_address>"
        
                lsTemp = lsTemp & "<markfor_address>"
                lsTemp = lsTemp & "</markfor_address>"
            
                lsTemp = lsTemp & "<tax>"
                lsTemp = lsTemp & "<tax_codes>"
                lsTemp = lsTemp & "<tax_code>"
                lsTemp = lsTemp & "</tax_code>"
                lsTemp = lsTemp & "</tax_codes>"
                lsTemp = lsTemp & "</tax>"
    
                lsTemp = lsTemp & "<prepayments>"
                lsTemp = lsTemp & "<prepayment>"
                lsTemp = lsTemp & "</prepayment>"
                lsTemp = lsTemp & "</prepayments>"
            
                lsTemp = lsTemp & "<user_fields>"
                lsTemp = lsTemp & "</user_fields>"
                
                lsTemp = lsTemp & "<applied_coupon_codes>"
                lsTemp = lsTemp & "</applied_coupon_codes>"
                
                lsTemp = lsTemp & "<creditcardinfo>"
                lsTemp = lsTemp & "</creditcardinfo>"
                
                lsTemp = lsTemp & "<ref_details>"
                lsTemp = lsTemp & "</ref_details>"
                
                lsTemp = lsTemp & "<date_details>"
                lsTemp = lsTemp & "</date_details>"
                
                lsTemp = lsTemp & "<sac_details>"
                lsTemp = lsTemp & "</sac_details>"
                
            lsTemp = lsTemp & "</header>"
        
            'PO1 Loop
            lsTemp = lsTemp & "<lines>"
                s_RsFilter "[Segment]='PO1_PO1' Or [Segment]='PID_PO1;PID'"
                While Not oRs.EOF
                    If Trim(oRs.Fields("Segment").Value) = "PO1_PO1" Then
                        lsTemp = lsTemp & "<line>"
                        lsTemp = lsTemp & "<item_no>" & Trim(oRs.Fields(7).Value) & "</item_no>"
                        lsTemp = lsTemp & "<qty_ordered>" & Trim(oRs.Fields(2).Value) & "</qty_ordered>"
                        lsTemp = lsTemp & "<unit_price>" & Trim(oRs.Fields(4).Value) & "</unit_price>"
                        lsTemp = lsTemp & "<unit_of_measure>" & Trim(oRs.Fields(3).Value) & "</unit_of_measure>"
                        lsTemp = lsTemp & "<price_uom>" & Trim(oRs.Fields(5).Value) & "</price_uom>"
                        lsTemp = lsTemp & "<line_sequence>" & Trim(oRs.Fields(1).Value) & "</line_sequence>"
                    End If
                    
                    'PID
                    If Trim(oRs.Fields("Segment").Value) = "PID_PO1;PID" Then
                        lsTemp = lsTemp & "<item_description_1>" & f_ReplaceXMLKeyWords(Trim(oRs.Fields(5).Value)) & "</item_description_1>"
                        lsTemp = lsTemp & "<line_comments>"
                        lsTemp = lsTemp & "<line_comment></line_comment>"
                        lsTemp = lsTemp & "</line_comments>"
                        lsTemp = lsTemp & "<user_fields></user_fields>"
                        lsTemp = lsTemp & "<uom_details></uom_details>"
                        lsTemp = lsTemp & "<ship_to_address></ship_to_address>"
                        lsTemp = lsTemp & "<sac_details></sac_details>"
                    End If
                    
                    oRs.MoveNext
            
                    If oRs.EOF Then
                        lsTemp = lsTemp & "</line>"
                    ElseIf Trim(oRs.Fields("Segment").Value) = "PO1_PO1" Then
                        lsTemp = lsTemp & "</line>"
                    End If
                Wend
            lsTemp = lsTemp & "</lines>"
              
        lsTemp = lsTemp & "</order>"
    lsTemp = lsTemp & "</orders>"
    
    'Change the file Name as per customer\ tp
    Dim ls_FileName As String
    ls_FileName = "\WildTool_Menards_" & ls_PONO & "_" & Trim(ls_ICN) & ".XML"
    
    If lb_Local_Server Then
        lsFilePath = Replace(App.Path, "EDIEXE", "EDIOUT\FlatFile") & ls_FileName
    Else
        lsFilePath = App.Path & "\WildTool_Menards_" & Trim(ls_PONO) & ".xml"
    End If

    liFreeFile = FreeFile
    Open lsFilePath For Output As #liFreeFile
        Print #liFreeFile, Trim(lsTemp)
    Close #liFreeFile
    
    'Special case for stop the duplicate file placing on FTP
    Dim fso As New FileSystemObject
    If lb_Local_Server Then
        If Not fso.FileExists("S:\FTPSite\WILDE TOOL\ARCHIVE" & Trim(ls_FileName)) Then
            FileCopy lsFilePath, "S:\FTPSite\WILDE TOOL\INBOUND" & Trim(ls_FileName)
        End If
        FileCopy lsFilePath, "S:\FTPSite\WILDE TOOL\ARCHIVE" & Trim(ls_FileName)
    Else
        If Not fso.FileExists(App.Path & "\FTPFolder\Archive" & Trim(ls_FileName)) Then
            FileCopy lsFilePath, App.Path & "\FTPFolder\Inbound" & Trim(ls_FileName)
        End If
        FileCopy lsFilePath, App.Path & "\FTPFolder\Archive" & Trim(ls_FileName)
    End If
    
   
    Exit Sub
ErrEntry:
    'If any Error raised in this procedure, it's log will be maintained
    s_ErrorLogEntryOnError Err
    End
End Sub

'Procedure for stablished the connection with database
Private Sub s_OpenConnection()

    On Error GoTo ErrEntry
    ls_CustomErrorMsg = "s_OpenConnection"
    
    Dim ls_LineData As String
    Dim ls_Sql_Server As String, ls_Sql_Username As String, ls_Sql_Password As String, ls_Sql_Database_Foundation As String, ls_Sql_Database_Raw As String, ls_Sql_Database_2008 As String
    Dim ls_ConnectionString As String
    
    Open App.Path & "\MXSAC.inf" For Input As #1
        Do While Not EOF(1)
            Line Input #1, ls_LineData
            If Mid(ls_LineData, 1, 4) = "IPST" Then ls_Sql_Server = Mid(ls_LineData, 5, Len(Trim(ls_LineData)) - 8)
            If Mid(ls_LineData, 1, 6) = "USERST" Then ls_Sql_Username = Mid(ls_LineData, 7, Len(Trim(ls_LineData)) - 12)
            If Mid(ls_LineData, 1, 6) = "PASSST" Then ls_Sql_Password = Mid(ls_LineData, 7, Len(Trim(ls_LineData)) - 12)
            If Mid(ls_LineData, 1, 10) = "DATABASEST" Then ls_Sql_Database_Foundation = Mid(ls_LineData, 11, Len(Trim(ls_LineData)) - 20)
            If Mid(ls_LineData, 1, 11) = "DATABASE2ST" Then ls_Sql_Database_Raw = Mid(ls_LineData, 12, Len(Trim(ls_LineData)) - 22)
            If Mid(ls_LineData, 1, 11) = "DATABASE3ST" Then ls_Sql_Database_2008 = Mid(ls_LineData, 12, Len(Trim(ls_LineData)) - 22)
        Loop
    Close #1

    ls_ConnectionString = "Provider=SQLOLEDB.1;User ID=" & ls_Sql_Username & ";pwd=" & ls_Sql_Password & ";Data Source=" & ls_Sql_Server
    
    Set oCon_Edi_Foundation_2008 = New ADODB.Connection
    oCon_Edi_Foundation_2008.CursorLocation = adUseClient
    oCon_Edi_Foundation_2008.Open ls_ConnectionString
    oCon_Edi_Foundation_2008.Execute "use " & ls_Sql_Database_Foundation
    
    Set oCon_Edi_2008 = New ADODB.Connection
    oCon_Edi_2008.CursorLocation = adUseClient
    oCon_Edi_2008.Open ls_ConnectionString
    oCon_Edi_2008.Execute "use " & ls_Sql_Database_2008

    Exit Sub
ErrEntry:
    'If there is any problem in Database Connection, Error log will be maintained in "BackEndExeConnectionErrorLog.txt"
    Dim intFile As Integer
    
    intFile = FreeFile
    Open App.Path & "\BackEndExeConnectionErrorLog.txt" For Append As #intFile
    Print #intFile, ls_ExeName & "#" & ls_Developer & "#" & Now & "#" & ls_ConnectionString
    
    Close #intFile
    End
End Sub

'Procedure for reading the EDI file
Private Sub s_ReadEdiSegmentsData()

    On Error GoTo ErrEntry
    ls_CustomErrorMsg = "s_ReadEdiSegmentsData"
    
    Dim oEngine As EDIdEV.ediEngine
    Dim oInterChange As EDIdEV.ediInterchange
    Dim oSegment As ediDataSegment
    Dim oDataElement As New ediDataElement
    Dim li_SerailNo As Integer
    Dim ls_SegmentId As String, ls_LoopSection As String, lsN101 As String
    Dim li_i As Integer, li_j As Integer
    Dim lb_Composite As Boolean, lb_CompAddNew As Boolean
    
    oRs.Fields.Append "Segment", adChar, 50, adFldIsNullable
    For li_i = 1 To 32
        oRs.Fields.Append (li_i), adChar, 500, adFldIsNullable
    Next
    oRs.CursorType = adOpenDynamic
    oRs.CursorLocation = adUseClient
    oRs.Open
    
    Set oRsComposite = New ADODB.Recordset
    oRsComposite.Fields.Append "Segment", adChar, 50, adFldIsNullable
    For li_i = 1 To 15
        oRsComposite.Fields.Append (li_i), adChar, 500, adFldIsNullable
    Next
    oRsComposite.CursorType = adOpenDynamic
    oRsComposite.CursorLocation = adUseClient
    oRsComposite.Open
        
    'To fetch the Path & Name of EDI file
    If lb_Local_Server Then
        With oCon_Edi_Foundation_2008.Execute("Select FileNameDownloaded from tbl_ExcelSheetInboundFiles where RecIDExcel = " & Trim(Command))
            If Not .BOF And Not .EOF Then
                ls_FileNameDownloaded = .Fields(0).Value
            Else
                ls_CustomErrorMsg = "FileNameDownloaded not found for RecIDExcel = " & Trim(Command)
                GoTo ErrEntry
            End If
        End With
    Else
        ls_FileNameDownloaded = Command
    End If
    
    Set oEngine = New EDIdEV.ediEngine
    Set oInterChange = oEngine.CreateInterchange(ls_FileNameDownloaded, 0)
    oInterChange.Include App.Path & "\SefFiles\850_X12-4010.SEF", 0
    oInterChange.Import
    oInterChange.SegmentTerminator = "~" & vbCrLf
    oInterChange.ElementTerminator = "*"
    oInterChange.CompositeTerminator = "|"
        
    While Not oInterChange.EOF
        Set oSegment = oInterChange.Segment
        ls_SegmentId = oSegment.SegmentID
        ls_LoopSection = Trim(oSegment.LoopSection)
        If Len(ls_LoopSection) = 0 Then
            ls_SegmentId = Trim(oSegment.SegmentID)
        Else
            ls_SegmentId = Trim(oSegment.SegmentID) & "_" & Trim(ls_LoopSection)
        End If
        
        li_SerailNo = li_SerailNo + 1
        lb_Composite = False
        
        If Trim(ls_SegmentId) = "N1_N1" Then lsN101 = Trim(oSegment.DataElementValue(1))
        
        oRs.AddNew
        oRs.Fields("Segment").Value = Trim(ls_SegmentId)
        For li_i = 1 To 30
            Set oDataElement = oSegment.DataElement(li_i)
            If oSegment.IsElementComposite = False Then
                oRs.Fields(li_i).Value = oDataElement.Value
            Else
                
                lb_Composite = True
                lb_CompAddNew = False
                For li_j = 1 To 15
                    If Len(Trim(oSegment.DataElementValue(li_i, li_j))) > 0 Then
                        If lb_CompAddNew = False Then
                            lb_CompAddNew = True
                            
                            oRsComposite.AddNew
                            oRsComposite.Fields("Segment").Value = Trim(CStr(li_SerailNo))
                            oRsComposite.Fields("15").Value = Trim(CStr(li_i))
                        End If
                        oRsComposite.Fields(li_j).Value = oSegment.DataElementValue(li_i, li_j)
                    End If
                Next
            End If
        Next
        If Trim(ls_LoopSection) = "N1" Then oRs.Fields("30").Value = Trim(lsN101)
        oRs.Fields("31").Value = Trim(CStr(li_SerailNo))
        If lb_Composite = False Then
            oRs.Fields("32").Value = "F"
        Else
            oRs.Fields("32").Value = "T"
        End If
        
        oRs.Update
        
        oInterChange.MoveNext
    Wend
    
    If Not lb_Local_Server Then
        Call s_WriteRecordsetToTextFile(oRs, "Recordset.txt")
        Call s_WriteRecordsetToTextFile(oRsComposite, "RecordsetComposite.txt")
    End If
    
    Exit Sub
ErrEntry:
    'If any Error raised in this procedure, it's log will be maintained
    s_ErrorLogEntryOnError Err
    End
End Sub

'Function for getting Description of a Code from an specified Table
Private Function f_GetDescription(ps_Value As String, ps_Table As String, Optional ps_SelectField As String = "Description", Optional ps_WhereField As String = "Code") As String
    If Trim(ps_Value) <> "" Then
        With oCon_Edi_2008.Execute("Select " & Trim(ps_SelectField) & " From " & Trim(ps_Table) & " Where " & Trim(ps_WhereField) & " = '" & Trim(ps_Value) & "'")
            If Not .EOF Then
                f_GetDescription = Trim(.Fields(ps_SelectField).Value)
            Else
                f_GetDescription = Trim(ps_Value)
            End If
        End With
    Else
        f_GetDescription = ""
    End If
End Function

'Procedure for filtering main recordset according to given Criteria
Private Sub s_RsFilter(ps_Criteria As String)
    oRs.Filter = ""
    oRs.MoveFirst
    oRs.Filter = ps_Criteria
End Sub

'Procedure for filtering composite recordset according to given Criteria
Private Sub s_CompositeRsFilter(ps_CompositeCriteria As String)
    oRsComposite.Filter = ""
    oRsComposite.MoveFirst
    oRsComposite.Filter = ps_CompositeCriteria
End Sub

''''''Procedure for generating the Excel Sheet
'''''Private Sub s_ExcelReport()
'''''
'''''    On Error GoTo ErrEntry
'''''    ls_CustomErrorMsg = "s_ExcelReport"
'''''
'''''
'''''    If lb_Local_Server Then
'''''        ls_XlFileName = Replace(App.Path, "EDIEXE", "EDIExcel") & "\WildTool_Menards_850_" & ls_BEG3 & Format(Now, "HHMMSS") & ".csv"
'''''    Else
'''''        ls_XlFileName = App.Path & "\WildTool_Menards_850_" & ls_BEG3 & ".csv"
'''''    End If
'''''
'''''
'''''    Dim ls_Temp As String, li_i As Integer
'''''    Dim li_FreeFile As Integer
'''''
'''''    oRs.Filter = ""
'''''    oRs.MoveFirst
'''''
'''''    li_FreeFile = 51
'''''    Open ls_XlFileName For Output As #li_FreeFile
'''''        ls_Temp = ""
'''''        For li_i = 0 To oRs.Fields.Count - 1
'''''            ls_Temp = ls_Temp & oRs.Fields(li_i).Name & ","
'''''        Next
'''''        Print #li_FreeFile, ls_Temp
'''''
'''''        If Not oRs.EOF And Not oRs.BOF Then oRs.MoveFirst
'''''        Do While Not oRs.EOF
'''''            ls_Temp = ""
'''''            For li_i = 0 To oRs.Fields.Count - 1
'''''                ls_Temp = ls_Temp & Trim(oRs.Fields(li_i).Value) & ","
'''''            Next
'''''            Print #li_FreeFile, ls_Temp
'''''            oRs.MoveNext
'''''        Loop
'''''    Close #li_FreeFile
'''''
'''''
'''''    Exit Sub
'''''ErrEntry:
'''''    'If any Error raised in this procedure, it's log will be maintained
'''''    s_ErrorLogEntryOnError Err
'''''    End
'''''End Sub

'Procedure for generating the Excel Sheet
Private Sub s_ExcelReport()
    On Error GoTo ErrEntry

    ls_CustomErrorMsg = "s_ExcelReport"

    Dim li_ITDLine As Integer, li_SACLine As Integer, li_N1Line As Integer
    Dim li_HeaderlLine As Integer, li_POLine As Integer, li_PIDLine As Integer
    Dim li_PO1Count As Integer

    Dim ld_Total As Double
    Dim ld_GTotal As Double
    Dim lb_HeaderPrinted As Boolean
    Dim li_CTTLine As Integer
    Dim ls_TempPrint As String
    Dim li_N9MSGLine As Integer

    Set oXlApp = New Excel.Application
    Set oXlApp = CreateObject("Excel.Application")
    oXlApp.Workbooks.Add

    If Not lb_Local_Server Then oXlApp.Visible = True

    With oXlApp
        'ISA
        s_RsFilter "Segment='ISA'"
        If Not oRs.EOF Then
            If Trim(oRs.Fields(15).Value) = "T" Then
                ls_ISA15 = "Test"
            ElseIf Trim(oRs.Fields(15).Value) = "P" Then
                ls_ISA15 = "Production"
            Else
                ls_ISA15 = Trim(oRs.Fields(15).Value)
            End If
        End If

        s_HeaderShow ""
        s_DisplayHeader

        li_XlLineCounter = 3
        s_HeaderShow ""
        li_HeaderlLine = 5
        s_HeaderShow "Purchage Order Details"
        s_SolidCellsWhiteText
        s_CreateBorder "B", "G", 5, 5
        li_XlLineCounter = 5

        s_HeaderShowLeft "Document Indicator"
        s_HeaderShowRight IIf(ls_ISA15 <> "", ls_ISA15, "---N/A---")

        'ST
        s_RsFilter "Segment='ST'"
        If Not oRs.EOF Then
            s_HeaderShowLeft "Transaction Set Control No"
            s_HeaderShowRight IIf(Trim(oRs.Fields(2).Value) <> "", Trim(oRs.Fields(2).Value), "---N/A---")
            .Selection.NumberFormat = "#0000"
        End If

        'BEG
        s_RsFilter "Segment='BEG'"
        If Not oRs.EOF Then
            s_HeaderShowLeft "Transaction Set Purpose Code"
            s_HeaderShowRight IIf(Trim(oRs.Fields(1).Value) <> "", f_GetDescription(Trim(oRs.Fields(1).Value), "TransactionPurposeMaster_353"), "---N/A---")

            s_HeaderShowLeft "Purchase Order Type Code"
            s_HeaderShowRight IIf(Trim(oRs.Fields(2).Value) <> "", f_GetDescription(Trim(oRs.Fields(2).Value), "POTypemaster_92"), "---N/A---")

            ls_BEG3 = Trim(oRs.Fields(3).Value)
            s_HeaderShowLeft "Purchase Order Number"
            s_HeaderShowRight IIf(ls_BEG3 <> "", "'" & ls_BEG3, "---N/A---")

            If Trim(oRs.Fields(4).Value) <> "" Then
                s_HeaderShowLeft "Release Number"
                s_HeaderShowRight Trim(oRs.Fields(4).Value)
            End If

            s_HeaderShowLeft "Purchase Order Date"
            s_HeaderShowRight IIf(Trim(oRs.Fields(5).Value) <> "", f_DateFormat(Trim(oRs.Fields(5).Value)), "---N/A---")

        End If

        'REF
        s_RsFilter "Segment='REF'"
        While Not oRs.EOF
            If Trim(oRs.Fields(1).Value) <> "" And (Trim(oRs.Fields(2).Value) <> "" Or Trim(oRs.Fields(3).Value) <> "") Then
                s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(1).Value), "ReferenceNumberQualifier_128")
                s_HeaderShowRight "'" & Trim(oRs.Fields(2).Value) & "  " & Trim(oRs.Fields(3).Value)
            End If
            oRs.MoveNext
        Wend

         'PER
         s_RsFilter "Segment='PER'"
         While Not oRs.EOF
            If Trim(oRs.Fields("1").Value) <> "" And Trim(oRs.Fields("2").Value) <> "" Then
                s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields("1").Value), "ContactFunctionCode_366")
                s_HeaderShowRight "'" & Trim(oRs.Fields("2").Value)
            ElseIf Trim(oRs.Fields("1").Value) <> "" Then
                s_HeaderShowLeft "Contact Function Code"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields("1").Value), "ContactFunctionCode_366")
            End If
            If Trim(oRs.Fields("3").Value) <> "" And Trim(oRs.Fields("4").Value) <> "" Then
               s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields("3").Value), "CommunicationNumberQualifier_365")
               s_HeaderShowRight Trim(oRs.Fields("4").Value)
            End If
            oRs.MoveNext
         Wend

       'FOB
        s_RsFilter "Segment='FOB'"
        While Not oRs.EOF
            If Trim(oRs.Fields(1).Value) <> "" Then
                s_HeaderShowLeft "Shipment Method of Payment"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "ShipmentMethodofPayment_146")
            End If

            If Trim(oRs.Fields(2).Value) <> "" Then
                s_HeaderShowLeft "Location Qualifier"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(2).Value), "LocationQualifier_309")
            End If
            If Trim(oRs.Fields(3).Value) <> "" Then
                s_HeaderShowLeft "Free form :" & Trim(oRs.Fields(3).Value), "B", "G"
                .Selection.Font.Bold = False
            End If
            oRs.MoveNext
        Wend

          'CSH
        s_RsFilter "Segment='CSH'"
        While Not oRs.EOF
            If Trim(oRs.Fields(1).Value) <> "" Then
                s_HeaderShowLeft "Sales Requirement Code"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "SalesRequirementCode_563")
            End If
            oRs.MoveNext
        Wend

        'DTM
        s_RsFilter "Segment='DTM'"
        While Not oRs.EOF
            If Trim(oRs.Fields(2).Value) <> "" Then
                s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(1).Value), "DateTimeQualifier_374")
                s_HeaderShowRight f_DateFormat(Trim(oRs.Fields(2).Value))
            End If
            oRs.MoveNext
        Wend

        'TD5
        s_RsFilter "Segment='TD5'"
        While Not oRs.EOF
            If Trim(oRs.Fields(1).Value) <> "" Then
                s_HeaderShowLeft "Routing Sequence Code"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "RoutingSequenceCode_133")
            End If

            If Trim(oRs.Fields(2).Value) <> "" And Trim(oRs.Fields(3).Value) <> "" Then
                s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(2).Value), "IdentificationCodeQualifier_66")
                s_HeaderShowRight Trim(oRs.Fields(3).Value)
            End If

            If Trim(oRs.Fields(4).Value) <> "" Then
                s_HeaderShowLeft "Transportation Method"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(4).Value), "TransportationMethodCode_91")
            End If

            If Trim(oRs.Fields(5).Value) <> "" Then
                s_HeaderShowLeft "Routing"
                s_HeaderShowRight Trim(oRs.Fields(5).Value)
            End If
            oRs.MoveNext
        Wend
        s_CreateBorder "B", "G", li_HeaderlLine, li_XlLineCounter

        'SAC
        s_RsFilter "Segment='SAC_SAC'"
        lb_HeaderPrinted = False
        While Not oRs.EOF
            If Not lb_HeaderPrinted Then
                s_HeaderShow ""
                s_HeaderShow "Service, Promotion, Allowance, or Charge Information"
                li_SACLine = li_XlLineCounter
                s_SolidCellsWhiteText
                s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
                lb_HeaderPrinted = True
            End If
            If Trim(oRs.Fields(1).Value) <> "" Then
                s_HeaderShowLeft "Allowance or Charge Indicator"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "AllowanceOrChargeIndicator_248")
            End If
            If Trim(oRs.Fields(1).Value) <> "" Then
                s_HeaderShowLeft "Allowance or Charge Indicator"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "AllowanceOrChargeIndicator_248")
            End If
            If Trim(oRs.Fields(3).Value) <> "" Then
                s_HeaderShowLeft "Agency Qualifier Code"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(3).Value), "AgencyQualifierCode_559")
            End If
            If Trim(oRs.Fields(4).Value) <> "" Then
                s_HeaderShowLeft "Charge Code"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(4).Value), "ServicePromotionAllowanceCode_1300")
            End If
            If Trim(oRs.Fields(5).Value) <> "" Then
                s_HeaderShowLeft "Amount"
                s_HeaderShowRight "'" & Trim(oRs.Fields(5).Value)
            End If

            oRs.MoveNext
        Wend
        If li_SACLine > 0 Then s_CreateBorder "B", "G", li_SACLine, li_XlLineCounter

        'ITD
        s_RsFilter "Segment='ITD'"
        lb_HeaderPrinted = False
        While Not oRs.EOF
            If Not lb_HeaderPrinted Then
                s_HeaderShow ""
                s_HeaderShow "Terms of Sale/Deferred Terms of Sale"
                li_ITDLine = li_XlLineCounter
                s_SolidCellsWhiteText
                s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
                lb_HeaderPrinted = True
            End If
            If Trim(oRs.Fields(1).Value) <> "" Then
                s_HeaderShowLeft "Terms Type Code"
                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "BasicTermCode_336")
            End If

            If Trim(oRs.Fields(3).Value) <> "" Then
                s_HeaderShowLeft "Terms Discount Percent"
                s_HeaderShowRight Trim(oRs.Fields(3).Value)
            End If

            If Trim(oRs.Fields(5).Value) <> "" Then
                s_HeaderShowLeft "Terms Discount Days Due"
                s_HeaderShowRight Trim(oRs.Fields(5).Value)
            End If

            If Trim(oRs.Fields(7).Value) <> "" Then
                s_HeaderShowLeft "Terms Net Days"
                s_HeaderShowRight Trim(oRs.Fields(7).Value)
            End If

            If Trim(oRs.Fields(12).Value) <> "" Then
                s_HeaderShowLeft "Free-form :" & Trim(oRs.Fields(12).Value), "B", "G"
                .Selection.Font.Bold = False
            End If

            oRs.MoveNext
        Wend
        If li_ITDLine > 0 Then s_CreateBorder "B", "G", li_ITDLine, li_XlLineCounter

        'N9---------------------------------------------------------------------------------------------------------------
        s_RsFilter "Segment='N9_N9' Or Segment='MSG_N9'"
        While Not oRs.EOF
            If Trim(oRs.Fields("Segment").Value) = "N9_N9" Then

                If Trim(oRs.Fields(1)) <> "" Then
                    s_HeaderShow ""
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(1).Value), "ReferenceNumberQualifier_128")
                    li_N9MSGLine = li_XlLineCounter
                    s_SolidCellsWhiteText
                    .Selection.HorizontalAlignment = xlLeft
                    s_HeaderShowRight Trim(oRs.Fields(3).Value)
                    s_SolidCellsWhiteText
                    .Selection.HorizontalAlignment = xlLeft
                    s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
                End If

                oRs.MoveNext

                'MSG
                Do While Not oRs.EOF
                    If Trim(oRs.Fields("Segment").Value) = "MSG_N9" Then
                        If Trim(oRs.Fields(1).Value) <> "" Then
                            s_HeaderShow "Free Form : " & Trim(oRs.Fields(1).Value)
                              oXlApp.Rows(li_XlLineCounter & ":" & li_XlLineCounter).RowHeight = (CInt(Len(Trim(oRs.Fields(1).Value)) / 90) + 1) * 12.75
                              Range("B" & li_XlLineCounter & ":G" & li_XlLineCounter).Select
                              Selection.WrapText = True
                            .Selection.Font.Bold = False
                            .Selection.HorizontalAlignment = xlLeft
                        End If
                        oRs.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                If li_N9MSGLine > 0 Then s_CreateBorder "B", "G", li_N9MSGLine, li_XlLineCounter
            Else
                oRs.MoveNext
            End If
        Wend

        'N1-LOOP==================================================================================================
        lb_HeaderPrinted = False
        s_RsFilter "Segment='N1_N1' Or Segment='N2_N1' Or Segment='N3_N1' Or Segment='N4_N1' Or Segment='PER_N1' "
        While Not oRs.EOF
            If Trim(oRs.Fields("Segment").Value) = "N1_N1" Then
                 If Trim(oRs.Fields(1).Value) <> "" Then
                     If Not lb_HeaderPrinted Then
                         s_HeaderShow ""
                         s_HeaderShow "Name And Address Details"
                         li_N1Line = li_XlLineCounter
                         s_SolidCellsWhiteText
                         s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
                         lb_HeaderPrinted = True
                     End If

                     s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(1).Value), "EntityIdentifierCode_98")
                     .Selection.Interior.ColorIndex = 15
                 End If
                 If Trim(oRs.Fields(2).Value) <> "" Then
                     s_HeaderShowRight Trim(oRs.Fields(2).Value)
                 End If
                 If Trim(oRs.Fields(3).Value) <> "" Then
                     ls_TempPrint = f_GetDescription(Trim(oRs.Fields(3).Value), "EntityIDQualifier_66")
                 End If
                 If Trim(oRs.Fields(4).Value) <> "" Then
                     If Len(ls_TempPrint) <> 0 Then
                         ls_TempPrint = ls_TempPrint & ": " & Trim(oRs.Fields(4).Value)
                     Else
                         ls_TempPrint = Trim(oRs.Fields(4).Value)
                     End If
                 End If

                 If Len(ls_TempPrint) <> 0 Then
                     If Trim(oRs.Fields(2).Value) <> "" Then
                         s_HeaderShowLeft ""
                         s_HeaderShowRight ls_TempPrint
                     Else
                         s_HeaderShowRight ls_TempPrint
                     End If
                 End If
                 ls_TempPrint = ""

                 oRs.MoveNext

                 'N2
                 Do While Not oRs.EOF
                     If Trim(oRs.Fields("Segment").Value) = "N2_N1" Then
                         If Trim(oRs.Fields(1).Value) <> "" Then
                             ls_TempPrint = Trim(oRs.Fields(1).Value)
                             If Trim(oRs.Fields(2).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(2).Value)
                             End If
                             s_HeaderShowLeft ""
                             s_HeaderShowRight ls_TempPrint
                         ElseIf Trim(oRs.Fields(2).Value) <> "" Then
                             s_HeaderShowLeft ""
                             s_HeaderShowRight "," & Trim(oRs.Fields(2).Value)
                         End If
                         oRs.MoveNext
                     Else
                         Exit Do
                     End If
                 Loop
                 ls_TempPrint = ""

                 'N3
                 Do While Not oRs.EOF
                     If Trim(oRs.Fields("Segment").Value) = "N3_N1" Then
                         If Trim(oRs.Fields(1).Value) <> "" Then
                             ls_TempPrint = Trim(oRs.Fields(1).Value)
                             If Trim(oRs.Fields(2).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(2).Value)
                             End If
                             s_HeaderShowLeft ""
                             s_HeaderShowRight ls_TempPrint
                         ElseIf Trim(oRs.Fields(2).Value) <> "" Then
                             s_HeaderShowLeft ""
                             s_HeaderShowRight "," & Trim(oRs.Fields(2).Value)
                         End If
                         oRs.MoveNext
                     Else
                         Exit Do
                     End If
                 Loop
                 ls_TempPrint = ""

                 'N4
                 Do While Not oRs.EOF
                     If Trim(oRs.Fields("Segment").Value) = "N4_N1" Then
                         If Trim(oRs.Fields(1).Value) <> "" Then
                             ls_TempPrint = Trim(oRs.Fields(1).Value)
                             If Trim(oRs.Fields(2).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(2).Value)
                             End If
                             If Trim(oRs.Fields(3).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(3).Value)
                             End If
                             If Trim(oRs.Fields(4).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(4).Value)
                             End If
                             s_HeaderShowLeft ""
                             s_HeaderShowRight ls_TempPrint
                         ElseIf Trim(oRs.Fields(2).Value) <> "" Then
                             ls_TempPrint = Trim(oRs.Fields(2).Value)
                             If Trim(oRs.Fields(3).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(3).Value)
                             End If
                             If Trim(oRs.Fields(4).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(4).Value)
                             End If
                             s_HeaderShowLeft ""
                             s_HeaderShowRight ls_TempPrint
                         ElseIf Trim(oRs.Fields(3).Value) <> "" Then
                             ls_TempPrint = Trim(oRs.Fields(3).Value)
                             If Trim(oRs.Fields(4).Value) <> "" Then
                                 ls_TempPrint = ls_TempPrint & "," & Trim(oRs.Fields(4).Value)
                             End If
                             s_HeaderShowLeft ""
                             s_HeaderShowRight ls_TempPrint
                         ElseIf Trim(oRs.Fields(4).Value) <> "" Then
                             s_HeaderShowLeft ""
                             s_HeaderShowRight Trim(oRs.Fields(4).Value)
                         End If
                         oRs.MoveNext
                     Else
                         Exit Do
                     End If
                 Loop
                 ls_TempPrint = ""

                 'PER
                 Do While Not oRs.EOF
                     If Trim(oRs.Fields("Segment").Value) = "PER_N1" Then
                         If Trim(oRs.Fields("1").Value) <> "" And Trim(oRs.Fields("2").Value) <> "" Then
                               s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields("1").Value), "ContactFunctionCode_366")
                               s_HeaderShowRight "'" & Trim(oRs.Fields("2").Value)
                           ElseIf Trim(oRs.Fields("1").Value) <> "" Then
                               s_HeaderShowLeft "Contact Function Code"
                               s_HeaderShowRight f_GetDescription(Trim(oRs.Fields("1").Value), "ContactFunctionCode_366")
                           End If
                         If Trim(oRs.Fields(3).Value) <> "" And Trim(oRs.Fields(4).Value) <> "" Then
                             s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(3).Value), "CommunicationNumberQualifier_365")
                             s_HeaderShowRight Trim(oRs.Fields(4).Value)
                         End If
                          If Trim(oRs.Fields(5).Value) <> "" And Trim(oRs.Fields(6).Value) <> "" Then
                             s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(5).Value), "CommunicationNumberQualifier_365")
                             s_HeaderShowRight Trim(oRs.Fields(6).Value)
                         End If
                         oRs.MoveNext
                     Else
                         Exit Do
                     End If
                Loop
            Else
                oRs.MoveNext
            End If
        Wend
        If li_N1Line > 0 Then s_CreateBorder "B", "G", li_N1Line, li_XlLineCounter


        'PO1=========================================================================================================
        s_HeaderShow ""
        s_HeaderShow "ITEM DETAILS"
        li_POLine = li_XlLineCounter
        s_SolidCellsWhiteText
        s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
        s_HeaderShow ""
        li_XlLineCounter = li_XlLineCounter + 1
        s_HeaderShowRight "SNo", "B", "B"
        s_SolidCellsWhiteText
        s_HeaderShowRight "Quantity", "C", "C"
        s_SolidCellsWhiteText
        s_HeaderShowRight "Measurement", "D", "D"
        s_SolidCellsWhiteText
        s_HeaderShowRight "UnitPrice", "E", "E"
        s_SolidCellsWhiteText
        s_HeaderShowRight "UnitPriceCode", "F", "F"
        s_SolidCellsWhiteText
        s_HeaderShowRight "Total", "G", "G"
        s_SolidCellsWhiteText

        li_PO1Count = 1
        s_RsFilter "Segment='PO1_PO1' Or Segment='PID_PO1;PID'"
        While Not oRs.EOF
            'PO1_PO1
            If Trim(oRs.Fields("Segment").Value) = "PO1_PO1" Then
                li_XlLineCounter = li_XlLineCounter + 1
                If li_POLine < 1 Then li_POLine = li_XlLineCounter
                If oRs.Fields("1").Value <> "" Then
                    s_HeaderShowRight oRs.Fields("1").Value, "B", "B"
                Else
                    s_HeaderShowRight Format(li_PO1Count, "0000"), "B", "B"
                    li_PO1Count = li_PO1Count + 1
                End If
                .Selection.NumberFormat = "0000"
                s_SolidCells
                s_HeaderShowRight IIf(Trim(oRs.Fields(2).Value) <> "", Trim(oRs.Fields(2).Value), "---N/A---"), "C", "C"
                s_SolidCells
                s_HeaderShowRight IIf(Trim(oRs.Fields(3).Value) <> "", f_GetDescription(Trim(oRs.Fields(3).Value), "MeasurementCode_355"), "---N/A---"), "D", "D"
                s_SolidCells
                s_HeaderShowRight IIf(Trim(oRs.Fields(4).Value) <> "", "$" & Trim(oRs.Fields(4).Value), "---N/A---"), "E", "E"
                s_SolidCells
                s_HeaderShowRight IIf(Trim(oRs.Fields(5).Value) <> "", f_GetDescription(Trim(oRs.Fields(5).Value), "UnitPriceCode_639"), "---N/A---"), "F", "F"
                s_SolidCells
                ld_Total = Val(Trim(oRs.Fields(2).Value)) * Val(Trim(oRs.Fields(4).Value))
                s_HeaderShowRight Format(ld_Total, "#.00"), "G", "G"
                .Selection.NumberFormat = "$0.00"
                s_SolidCells
                ld_GTotal = ld_GTotal + ld_Total

                If Trim(oRs.Fields(6).Value) <> "" And Trim(oRs.Fields(7).Value) <> "" Then
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(6).Value), "ProductServiceIDQualifier_235")
                    s_HeaderShowRight "'" & Trim(oRs.Fields(7).Value)
                End If
                If Trim(oRs.Fields(8).Value) <> "" And Trim(oRs.Fields(9).Value) <> "" Then
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(8).Value), "ProductServiceIDQualifier_235")
                    s_HeaderShowRight "'" & Trim(oRs.Fields(9).Value)
                End If
                If Trim(oRs.Fields(10).Value) <> "" And Trim(oRs.Fields(11).Value) <> "" Then
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(10).Value), "ProductServiceIDQualifier_235")
                    s_HeaderShowRight "'" & Trim(oRs.Fields(11).Value)
                End If
                If Trim(oRs.Fields(12).Value) <> "" And Trim(oRs.Fields(13).Value) <> "" Then
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(12).Value), "ProductServiceIDQualifier_235")
                    s_HeaderShowRight "'" & Trim(oRs.Fields(13).Value)
                End If
                If Trim(oRs.Fields(14).Value) <> "" And Trim(oRs.Fields(15).Value) <> "" Then
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(14).Value), "ProductServiceIDQualifier_235")
                    s_HeaderShowRight "'" & Trim(oRs.Fields(15).Value)
                End If
                If Trim(oRs.Fields(16).Value) <> "" And Trim(oRs.Fields(17).Value) <> "" Then
                    s_HeaderShowLeft f_GetDescription(Trim(oRs.Fields(16).Value), "ProductServiceIDQualifier_235")
                    s_HeaderShowRight "'" & Trim(oRs.Fields(17).Value)
                End If

                oRs.MoveNext

               'PID_PO1;PID
               lb_HeaderPrinted = False
                li_PIDLine = 0
                Do While Not oRs.EOF
                    If Trim(oRs.Fields("Segment").Value) = "PID_PO1;PID" Then
                            If Not lb_HeaderPrinted Then
                                s_HeaderShow "Product/Item Description"
                                li_PIDLine = li_XlLineCounter
                                .Selection.Interior.ColorIndex = 48
                                s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
                                lb_HeaderPrinted = True
                            End If

                            If Trim(oRs.Fields(1).Value) <> "" And Trim(oRs.Fields(5).Value) <> "" Then
                                li_XlLineCounter = li_XlLineCounter + 1
                                s_HeaderShowRight f_GetDescription(Trim(oRs.Fields(1).Value), "ItemDescriptionType_349") & " : " & Trim(oRs.Fields(5).Value), "B", "G"
                                .Selection.Font.Bold = False
                            End If

                        oRs.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                If li_PIDLine > 0 Then s_CreateBorder "B", "G", li_PIDLine, li_XlLineCounter

            Else
                oRs.MoveNext
            End If
            s_HeaderShow ""
        Wend
        If li_POLine > 0 Then li_XlLineCounter = li_XlLineCounter - 1
        If li_POLine > 0 Then s_CreateBorder "B", "G", li_POLine, li_XlLineCounter

        'CTT
        s_RsFilter "Segment='CTT_CTT' "
        li_CTTLine = 0
        lb_HeaderPrinted = False
        While Not oRs.EOF
            If Trim(oRs.Fields("Segment").Value) = "CTT_CTT" Then
                If Not lb_HeaderPrinted Then
                    s_HeaderShow ""
                    s_HeaderShow "Transaction Totals"
                    li_CTTLine = li_XlLineCounter
                    s_SolidCellsWhiteText
                    s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
                    lb_HeaderPrinted = True
                End If

                If Trim(oRs.Fields(1).Value) <> "" Then
                    s_HeaderShowLeft "No Of Line Items"
                    s_HeaderShowRight Trim(oRs.Fields(1).Value)
                End If
                oRs.MoveNext
             Else
                oRs.MoveNext
            End If
        Wend
        If li_CTTLine > 0 Then s_CreateBorder "B", "G", li_CTTLine, li_XlLineCounter

        '********************  Summary ******************************
        s_HeaderShow ""
        s_HeaderShowLeft "Grand Total"
        s_HeaderShowRight ld_GTotal
        .Selection.HorizontalAlignment = xlRight
        .Selection.NumberFormat = "$#,##0.00"

        s_CreateBorder "B", "G", li_XlLineCounter, li_XlLineCounter
    '-------
        s_HeaderShow ""
        s_HeaderShow "Infocon Systems recommends all customers to check their Infocon EDI Mailbox on a routine basis."
        .Selection.Font.Bold = False
        .Selection.Font.Italic = True
        .Selection.Font.Size = 8.6
        s_HeaderShow "Infocon Systems is not responsible for lost EDI documents because of email transmission network failures."
        .Selection.Font.Bold = False
        .Selection.Font.Italic = True
        .Selection.Font.Size = 8.6

        .Columns("a:a").ColumnWidth = 1
        .Columns("b:b").ColumnWidth = 14
        .Columns("c:c").ColumnWidth = 14
        .Columns("d:d").ColumnWidth = 14
        .Columns("e:e").ColumnWidth = 14
        .Columns("f:f").ColumnWidth = 14
        .Columns("g:g").ColumnWidth = 14
    End With

    If lb_Local_Server Then
        ls_XlFileName = Replace(App.Path, "EDIEXE", "EDIExcel") & "\WildTool_Menards_850_" & ls_BEG3 & Format(Now, "HHMMSS") & ".xls"
    Else
        ls_XlFileName = App.Path & "\WildTool_Menards_850_" & ls_BEG3 & ".xls"
    End If

    oXlApp.ActiveWorkbook.SaveAs FileName:=ls_XlFileName, FileFormat:=xlNormal, Password:="", _
    WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False

    oXlApp.Quit
    Set oXlApp = Nothing

    Exit Sub
ErrEntry:
    'If any Error raised in this procedure, it's log will be maintained
    s_ErrorLogEntryOnError Err
    End
End Sub

'Procedure for updating "ExcelFileName" in "tbl_ExcelSheetInboundFiles" Table
Private Sub s_AddToDataBase()
    On Error GoTo ErrEntry
    
    ls_CustomErrorMsg = "s_AddToDataBase"
    
    If lb_Local_Server Then
        oCon_Edi_Foundation_2008.Execute "Update tbl_ExcelSheetInboundFiles Set ExcelFileName = '" & ls_XlFileName & "', ExcelSentStatus = 'N' where RecIDExcel = " & Trim(Command)
    End If
    
    Exit Sub
ErrEntry:
    'If any Error raised in this procedure, it's log will be maintained
    s_ErrorLogEntryOnError Err
    End
End Sub

'Procedure for writing Recordset onto a text file while Testing
Sub s_WriteRecordsetToTextFile(ps_Rs As ADODB.Recordset, ps_FileName As String)
    Dim ls_Temp As String, li_i As Integer
    Dim li_FreeFile As Integer
    
    li_FreeFile = 51
    Open App.Path & "\" & ps_FileName For Output As #li_FreeFile
        ls_Temp = ""
        For li_i = 0 To ps_Rs.Fields.Count - 1
            ls_Temp = ls_Temp & ps_Rs.Fields(li_i).Name & "#"
        Next
        Print #li_FreeFile, ls_Temp
        
        If Not ps_Rs.EOF And Not ps_Rs.BOF Then ps_Rs.MoveFirst
        Do While Not ps_Rs.EOF
            ls_Temp = ""
            For li_i = 0 To ps_Rs.Fields.Count - 1
                ls_Temp = ls_Temp & Trim(ps_Rs.Fields(li_i).Value) & "#"
            Next
            Print #li_FreeFile, ls_Temp
            ps_Rs.MoveNext
        Loop
    Close #li_FreeFile
End Sub

'Procedure for displaying left part (column B to D) of Excel Sheet
Private Sub s_HeaderShowLeft(Optional ByVal header As String, Optional ByVal cell1 As String, Optional ByVal cell2 As String)
    li_XlLineCounter = li_XlLineCounter + 1
    If cell1 = "" Then cell1 = "B"
    If cell2 = "" Then cell2 = "D"
    With oXlApp
        .Range(cell1 & li_XlLineCounter & ":" & cell2 & li_XlLineCounter).Select
        s_MergeCells
        .ActiveCell.FormulaR1C1 = header
        .Selection.HorizontalAlignment = xlLeft
        .Selection.Font.Bold = True
    End With
End Sub

'Procedure for displaying right part (column E to G) of Excel Sheet
Private Sub s_HeaderShowRight(Optional ByVal header As String, Optional ByVal cell1 As String, Optional ByVal cell2 As String)
    If cell1 = "" Then cell1 = "E"
    If cell2 = "" Then cell2 = "G"
    With oXlApp
        .Range(cell1 & li_XlLineCounter & ":" & cell2 & li_XlLineCounter).Select
        s_MergeCells
        .ActiveCell.FormulaR1C1 = header
        .Selection.HorizontalAlignment = xlLeft
        .Selection.Font.Bold = False
    End With
End Sub
    
'Procedure for displaying header parts (column B to G) of Excel Sheet
Private Sub s_HeaderShow(Optional ByVal header As String, Optional ByVal cell1 As String, Optional ByVal cell2 As String)
    li_XlLineCounter = li_XlLineCounter + 1
    If cell1 = "" Then cell1 = "B"
    If cell2 = "" Then cell2 = "G"
    With oXlApp
        .Range(cell1 & li_XlLineCounter & ":" & cell2 & li_XlLineCounter).Select
        s_MergeCells
        .ActiveCell.FormulaR1C1 = header
        .Selection.HorizontalAlignment = xlCenter
        .Selection.Font.Bold = True
    End With
End Sub

'Procedure for merging the selected cells in Excel Sheet
Private Sub s_MergeCells()
    With oXlApp.Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .ShrinkToFit = False
        .MergeCells = True
    End With
End Sub

'Procedure for formatting the elements of PO1 on Excel Sheet
Private Sub s_SolidCells()
    With oXlApp.Selection
        With .Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

'Procedure for formatting the Headers of Excel Sheet
Private Sub s_SolidCellsWhiteText()
    With oXlApp.Selection
        With .Interior
            .ColorIndex = 5
            .Pattern = xlSolid
        End With
        .Font.Bold = True
        .Font.ColorIndex = 2
        .HorizontalAlignment = xlCenter
    End With
End Sub

'Procedure for displaying the border in Excel Sheet
Private Sub s_CreateBorder(j1 As String, j2 As String, j3 As Integer, j4 As Integer)
    With oXlApp
        If j4 < 1 Then j4 = j3 + 1
        If j3 < 1 Then Exit Sub
        .Range(j1 & j3 & ":" & j2 & j4).Select
        .Range(j1 & j3 & ":" & j2 & j4).Activate
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("C3:E3").Select
    End With
End Sub

'Procedure for displaying the top most header on Excel Sheet
Private Sub s_DisplayHeader()
    With oXlApp
        .Range("B2").Select
        .ActiveCell.FormulaR1C1 = "EDI 850"
        .Range("B3").Select
        .ActiveCell.FormulaR1C1 = "PURCHASE ORDER DOCUMENT"
        .Range("B2:G3").Select
        .Selection.Font.Bold = True
        .Selection.HorizontalAlignment = xlCenter
        .Range("B2:G2").Select
        .Selection.HorizontalAlignment = xlCenter
        .Selection.MergeCells = True
        .Range("B3:G3").Select

        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = True
        End With

        .Range("B2:G3").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

        .Range("B2:G2").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("B5").Select
    End With
End Sub

'Function for displaying the date in Standard Format
Private Function f_DateFormat(p_date As String)
    f_DateFormat = Mid(p_date, 5, 2) & "/" & Right(p_date, 2) & "/" & Left(p_date, 4)
End Function

'Procedure for inserting a record in "tbl_BackEndExeErrorLog" Table on Project Startup
Private Sub s_ErrorLogEntryOnProjectStart()
    Dim ls_Sql As String
    
    ls_Sql = "INSERT INTO tbl_BackEndExeErrorLog " & _
            "(ExeName, Developer, StartDateTimeLog, EndDateTimeLog, " & _
            "ErrorStatus, ReceiptParameter, TransetNo) " & _
            "VALUES('" & ls_ExeName & "', '" & ls_Developer & "', '" & Now & "', '" & Now & _
            "', 'Process', '" & Command & "', '" & ls_TransetNo & "')"
    
    oCon_Edi_Foundation_2008.Execute ls_Sql
    
    With oCon_Edi_Foundation_2008.Execute("select @@Identity from tbl_BackEndExeErrorLog")
        ld_RecIdLog = .Fields(0).Value
    End With
End Sub

'Procedure for updating "tbl_BackEndExeErrorLog" Table on Project End
Private Sub s_ErrorLogEntryOnProjectEnd()
    Dim ls_Sql As String
    
    ls_Sql = "Update tbl_BackEndExeErrorLog " & _
            "SET EndDateTimeLog ='" & Now & _
            "', ErrorStatus = 'Successful', " & _
            "ExeOutputName = '" & ls_XlFileName & "'" & _
            "where RecIDLog = '" & ld_RecIdLog & "'"
    
    oCon_Edi_Foundation_2008.Execute ls_Sql
End Sub

'Procedure for updating "tbl_BackEndExeErrorLog" Table in case of any Error
Private Sub s_ErrorLogEntryOnError(Optional E As ErrObject)
    Dim ls_Sql As String
    
    ls_Sql = "Update tbl_BackEndExeErrorLog SET "
    
    If Not E Is Nothing Then
        ls_Sql = ls_Sql + "ErrorNo = '" & E.Number & "', " & _
                "ErrorSource = '" & E.Source & "', " & _
                "ErrorDescription = '" & Replace(E.Description, "'", "") & "', "
    End If
    
    ls_Sql = ls_Sql + "EndDateTimeLog = '" & Now & "', " & _
            "ErrorStatus = 'Partially Executed" & "', " & _
            "CustomErrorMsg = '" & ls_CustomErrorMsg & "' " & _
            "where RecIDLog = '" & ld_RecIdLog & "'"
    
    oCon_Edi_Foundation_2008.Execute ls_Sql
End Sub

Private Function f_ReplaceXMLKeyWords(ByVal ps_Value As String)
    ps_Value = Trim(ps_Value)
    ps_Value = Replace(ps_Value, "&", "&amp;")
    ps_Value = Replace(ps_Value, "<", "&lt;")
    ps_Value = Replace(ps_Value, ">", "&gt;")
    ps_Value = Replace(ps_Value, "'", "&apos;")
    ps_Value = Replace(ps_Value, """", "&quot;")
    f_ReplaceXMLKeyWords = ps_Value
End Function
