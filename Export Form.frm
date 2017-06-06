VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "Export Form"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3075
   LinkTopic       =   "Form3"
   ScaleHeight     =   2715
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cstatus 
      Height          =   315
      ItemData        =   "Export Form.frx":0000
      Left            =   480
      List            =   "Export Form.frx":000A
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cexport 
      Caption         =   "Export"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cdepartment 
      Height          =   315
      ItemData        =   "Export Form.frx":0016
      Left            =   480
      List            =   "Export Form.frx":0018
      TabIndex        =   1
      Text            =   "cdepartment"
      Top             =   840
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Pilih Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2085
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ApExcel As Object

Private Function setTitle(lngbaris As Long, lngkolom As Long, strValue As String, Optional intFontSize As Integer = 15)
    ApExcel.Cells(lngbaris, lngkolom).Font.Size = intFontSize
    ApExcel.Cells(lngbaris, lngkolom).Font.Bold = True
    ApExcel.Cells(lngbaris, lngkolom).Value = strValue
    ApExcel.Cells(lngbaris, lngkolom).WrapText = False
End Function

Private Function setColTitle(lngbaris As Long, lngkolom As Long, strValue As String)
    ApExcel.Cells(lngbaris, lngkolom).Font.Size = 10
    ApExcel.Cells(lngbaris, lngkolom).Font.Bold = True
    ApExcel.Cells(lngbaris, lngkolom).Value = strValue
    ApExcel.Cells(lngbaris, lngkolom).Interior.ColorIndex = 6
    ApExcel.Cells(lngbaris, lngkolom).WrapText = False
    ApExcel.Cells(lngbaris, lngkolom).HorizontalAlignment = xlCenter
    ApExcel.Cells(lngbaris, lngkolom).VerticalAlignment = xlCenter
    ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeLeft).LineStyle = xlContinuous
    ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeRight).LineStyle = xlContinuous
    ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeTop).LineStyle = xlContinuous
    ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Function

Private Function setColVal(lngbaris As Long, lngkolom As Long, strValue As String, Optional strFormat As String = "", Optional intRemark As Integer = 0, Optional bolRound As Boolean = True, Optional ByRef hAlign As Excel.Constants = xlLeft)
    ApExcel.Cells(lngbaris, lngkolom).Font.Size = 10
    ApExcel.Cells(lngbaris, lngkolom).Value = strValue
    ApExcel.Cells(lngbaris, lngkolom).HorizontalAlignment = hAlign
    If bolRound = True Then
        ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeRight).LineStyle = xlContinuous
        ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeTop).LineStyle = xlContinuous
        ApExcel.Cells(lngbaris, lngkolom).Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If

    If strFormat <> "" Then
        ApExcel.Cells(lngbaris, lngkolom).NumberFormat = strFormat
    End If
    If intRemark <> 0 Then
        ApExcel.Cells(lngbaris, lngkolom).Interior.ColorIndex = 37
    End If
End Function

Public Function Progress_Me(intValue As Integer, Optional intMax As Integer = 0)
    If intMax <> 0 Then pgb.Max = intMax
    
    pgb.Value = intValue
    
End Function

Public Function ExportExcel(strkodedept As String)
    Dim MyFieldCount, I As Integer
    Dim tanggal As String
    Dim MyIndex As Long
    Dim MyRecordCount As Long
    Set rs_export = New ADODB.Recordset
    
    If rs_export.State = 1 Then rs_export.Close
    sql = " SELECT * FROM tag_stock_opname " & strkodedept & " "
    
    rs_export.Open sql, conn, adOpenDynamic, adLockOptimistic
    
    If rs_export.EOF Then
        MsgBox "Data tidak ada !", vbOKOnly + vbInformation, "Information"
        rs_export.Close
        Set rs_export = Nothing
        Exit Function
    End If
    
    pgb.Visible = True

    intValue = 0
    pgb.Min = 0
    pgb.Max = IIf(rs_export.RecordCount < 1, 2, rs_export.RecordCount)
    pgb.Visible = True
    tanggal = Now
    
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Visible = False
    ApExcel.Workbooks.Add
    
    setTitle 1, 1, "Departemen : " & UCase(cdepartment.Text)
    setTitle 2, 1, "Dicetak : " & tanggal
    setTitle 3, 1, "LAPORAN STOCK OPNAME", 20
    
    setColTitle 5, 1, "NO"
    setColTitle 5, 2, "TAG NO"
    setColTitle 5, 3, "PART CODE"
    setColTitle 5, 4, "PART NAME"
    setColTitle 5, 5, "INCH"
    setColTitle 5, 6, "SCH"
    setColTitle 5, 7, "THICKNESS"
    setColTitle 5, 8, "LENGTH"
    setColTitle 5, 9, "WIDTH"
    setColTitle 5, 10, "CATEGORY"
    setColTitle 5, 11, "LOCATION"
    setColTitle 5, 12, "U/M"
    setColTitle 5, 13, "QTY ADMIN"
    setColTitle 5, 14, "QTY ACTUAL"
    setColTitle 5, 15, "QTY VARIANCE"
    setColTitle 5, 16, "TAHUN KEDATANGAN"
    setColTitle 5, 17, "REMARKS"
            
    I = 1
    lngctrlbrs = 6
    rs_export.MoveFirst
    Do While Not rs_export.EOF
        intValue = intValue + 1
        pgb.Value = intValue
        
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 1).Value = I
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 2).Value = "'" & rs_export!tag_no
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 3).Value = "'" & rs_export!part_no
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 4).Value = "'" & rs_export!part_name
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 5).Value = "'" & rs_export!INCH
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 6).Value = "'" & rs_export!sch
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 7).Value = "'" & rs_export!Thickness
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 8).Value = "'" & rs_export!Length
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 9).Value = "'" & rs_export!Width
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 10).Value = "'" & rs_export!Category
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 11).Value = "'" & rs_export!Location
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 12).Value = "'" & rs_export!satuan
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 13).Value = IIf(IsNull(rs_export!qty_admin), 0, rs_export!qty_admin)
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 14).Value = IIf(IsNull(rs_export!qty_actual), 0, rs_export!qty_actual)
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 15).Value = IIf(IsNull(rs_export!variance), 0, rs_export!variance)
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 16).Value = rs_export!tahun
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 17).Value = rs_export!remarks
        If rs_export!variance < 0 Then _
        ApExcel.Range(ApExcel.ActiveSheet.Cells(lngctrlbrs, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Interior.ColorIndex = 3
        If rs_export!variance > 0 Then _
        ApExcel.Range(ApExcel.ActiveSheet.Cells(lngctrlbrs, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Interior.ColorIndex = 4
        
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 1).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 2).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 3).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 4).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 5).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 6).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 7).Font.Size = 10
       ' ApExcel.ActiveSheet.Cells(lngctrlbrs, 8).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 9).Font.Size = 10
        'ApExcel.ActiveSheet.Cells(lngctrlbrs, 10).Font.Size = 10
       ' ApExcel.ActiveSheet.Cells(lngctrlbrs, 11).Font.Size = 10
        
        rs_export.MoveNext
        I = I + 1
        lngctrlbrs = lngctrlbrs + 1
    Loop
    
    lngctrlbrs = lngctrlbrs - 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(5, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Borders(1).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(5, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Borders(2).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(5, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Borders(3).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(5, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Borders(4).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(5, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 17)).Font.Size = 10

     
    'mod.24-12-2013
    pgb.Visible = False
    
    rs_export.Close
    Set rs_export = Nothing
    ApExcel.Columns.AutoFit
    ApExcel.Columns(1).ColumnWidth = 20
    ReportSODetail = MyRecordCount
    ApExcel.Visible = True
    
    Set ApExcel = Nothing
End Function

Private Sub cexport_Click()
    Select Case cdepartment.ListIndex
        Case 0
            ExportExcel "where status='" & cstatus.Text & "' "
        Case 1
            ExportExcel "WHERE left(tag_no,3)='TRA' and status='" & cstatus.Text & "'"
        Case 2
            ExportExcel "WHERE left(tag_no,3)='TRB' and status='" & cstatus.Text & "'"
        Case 3
            ExportExcel "WHERE left(tag_no,3)='TRC' and status='" & cstatus.Text & "'"
        Case 4
            ExportExcel "WHERE left(tag_no,3)='TRD' and status='" & cstatus.Text & "'"
        Case 5
            ExportExcel "WHERE left(tag_no,3)='TRE' and status='" & cstatus.Text & "'"
        Case 6
            ExportExcel "WHERE left(tag_no,3)='TRF' and status='" & cstatus.Text & "'"
        Case 7
            ExportExcel "WHERE left(tag_no,3)='TRH' and status='" & cstatus.Text & "'"
        Case 8
            ExportExcel "WHERE left(tag_no,3)='TRI' and status='" & cstatus.Text & "'"
        Case 9
            ExportExcel "WHERE left(tag_no,3)='TRJ' and status='" & cstatus.Text & "'"
        Case 10
            ExportExcel "WHERE left(tag_no,3)='TRK' and status='" & cstatus.Text & "'"
        Case 11
            ExportExcel "WHERE left(tag_no,3)='TRT' and status='" & cstatus.Text & "'"
        Case 12
            ExportExcel "WHERE left(tag_no,3)='TRU' and status='" & cstatus.Text & "'"
        Case 13
            ExportExcel "WHERE left(tag_no,3)='TRS' and status='" & cstatus.Text & "'"
        Case 14
            ExportExcel "WHERE left(tag_no,3)='TR1' and status='" & cstatus.Text & "'"
        Case 15
            ExportExcel "WHERE left(tag_no,3)='TR2' and status='" & cstatus.Text & "'"
        Case 16
            ExportExcel "WHERE left(tag_no,3)='TR3' and status='" & cstatus.Text & "'"
        Case 17
            ExportExcel "WHERE left(tag_no,3)='TRL' and status='" & cstatus.Text & "'"
        Case 18
            ExportExcel "WHERE left(tag_no,3)='TCO' and status='" & cstatus.Text & "'"
        Case 19
            ExportExcel "WHERE left(tag_no,3)='TCL' and status='" & cstatus.Text & "'"
        Case 20
            ExportExcel "WHERE left(tag_no,3)='TMH' and status='" & cstatus.Text & "'"
        Case 21
            ExportExcel "WHERE left(tag_no,3)='TGS' and status='" & cstatus.Text & "'"
        Case 22
            ExportExcel "WHERE left(tag_no,3)='TCC' and status='" & cstatus.Text & "'"
        Case 23
            ExportExcel "WHERE left(tag_no,3)='TAS' and status='" & cstatus.Text & "'"
        Case 24
            ExportExcel "WHERE left(tag_no,2)='TB' and status='" & cstatus.Text & "'"
        
    End Select
    
End Sub

Private Sub Form_Activate()

cdepartment.Clear
cdepartment.AddItem "All Item"
cdepartment.AddItem "A"
cdepartment.AddItem "B"
cdepartment.AddItem "C"
cdepartment.AddItem "D"
cdepartment.AddItem "E"
cdepartment.AddItem "F"
cdepartment.AddItem "H"
cdepartment.AddItem "I"
cdepartment.AddItem "J"
cdepartment.AddItem "K"
cdepartment.AddItem "T"
cdepartment.AddItem "U"
cdepartment.AddItem "S"
cdepartment.AddItem "R1"
cdepartment.AddItem "R2"
cdepartment.AddItem "R3"
cdepartment.AddItem "R LUAR"
cdepartment.AddItem "CONTAINER"
cdepartment.AddItem "CSD-ROOM"
cdepartment.AddItem "MAIN WH"
cdepartment.AddItem "GAS"
cdepartment.AddItem "CONT. CONS"
cdepartment.AddItem "ASSET"
cdepartment.AddItem "TAG BLANK"

cdepartment.ListIndex = 0
cstatus.ListIndex = 0
cdepartment.SetFocus
End Sub

