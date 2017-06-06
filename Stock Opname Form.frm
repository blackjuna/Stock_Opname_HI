VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Stock Opname Form"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbstatus 
      Height          =   315
      ItemData        =   "Stock Opname Form.frx":0000
      Left            =   6120
      List            =   "Stock Opname Form.frx":000A
      TabIndex        =   48
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox tsearch 
      Height          =   375
      Left            =   9360
      TabIndex        =   46
      Top             =   240
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Input Tag Stock Opname"
      Height          =   2415
      Left            =   7680
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox tno_input 
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox tlocation_input 
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox tnopart_input 
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox tpartname_input 
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox tcategory_input 
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox tgrup_input 
         Height          =   285
         Left            =   5520
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox tqty_input 
         Height          =   285
         Left            =   5520
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox ctagcode_input 
         Height          =   315
         Left            =   5520
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cstatus_input 
         Height          =   315
         Left            =   5520
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Tag No"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Part No"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Part Name"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Grup"
         Height          =   195
         Left            =   4560
         TabIndex        =   35
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tag Code"
         Height          =   195
         Left            =   4560
         TabIndex        =   34
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   4560
         TabIndex        =   33
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   195
         Left            =   4560
         TabIndex        =   32
         Top             =   1560
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stock Opname Tag Detail"
      Height          =   3855
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   7215
      Begin VB.TextBox tcategory 
         Height          =   285
         Left            =   5520
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox twidth 
         Height          =   285
         Left            =   1200
         TabIndex        =   55
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton ccancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   5280
         TabIndex        =   54
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox tlength 
         Height          =   285
         Left            =   1200
         TabIndex        =   53
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox tsatuan 
         Height          =   285
         Left            =   5520
         TabIndex        =   51
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox tsize 
         Height          =   285
         Left            =   1200
         TabIndex        =   49
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox tqtyslh 
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox tqtyadm 
         Height          =   285
         Left            =   5520
         TabIndex        =   41
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cstatus 
         Height          =   315
         ItemData        =   "Stock Opname Form.frx":0016
         Left            =   5520
         List            =   "Stock Opname Form.frx":0018
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox ctag_code 
         Height          =   315
         Left            =   5520
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox tqty 
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox tthickness 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox tclass 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox tpart_name 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox tpart_no 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox tlocation 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox ttag_no 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   3360
         Width           =   420
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Thickness"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Satuan"
         Height          =   195
         Left            =   4560
         TabIndex        =   52
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Inch"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Qty Selisih"
         Height          =   195
         Left            =   4560
         TabIndex        =   44
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Qty Actual"
         Height          =   195
         Left            =   4560
         TabIndex        =   42
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Qty Admin"
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   4560
         TabIndex        =   17
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tag Code"
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sch"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Part Name"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Part No"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tag No"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   8705
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cdepartment 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label23 
      Caption         =   "Status :"
      Height          =   255
      Left            =   5280
      TabIndex        =   47
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Search :"
      Height          =   255
      Left            =   8400
      TabIndex        =   45
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Silahkan Pilih Department"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu progess 
      Caption         =   "Progress"
   End
   Begin VB.Menu tools 
      Caption         =   "Tools"
      Begin VB.Menu export 
         Caption         =   "Export To Excel"
      End
      Begin VB.Menu Import 
         Caption         =   "Import From Excel"
      End
      Begin VB.Menu print 
         Caption         =   "Print Tag No"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sql, str As String
Public tes As Integer
Sub filter_lv(strstatus As String)
If cdepartment.Text = "All Item" Then
        Dim Lst As ListItem, nmr As Integer
        If rs_so.State = 0 Then
            rs_so.Open "select *,(qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where status like '%" & strstatus & "%'", conn
        End If
        lihat = "select *,(qty_actual-(ISNULL(qty_admin,0)))  as qty_selisih  from tag_stock_opname where status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    ElseIf cdepartment.Text = "A" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRA' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRA' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "B" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRB' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRB' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "C" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRC' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRC' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "D" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRD' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRD' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "E" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRE' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRE' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "F" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRF' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRF' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "H" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRH' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRH' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "I" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRI' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRI' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "J" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRJ' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRJ' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "K" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRK' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRK' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "T" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRT' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRT' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "U" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRU' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRU' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "S" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRS' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRS' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "R1" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TR1' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TR1' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "R2" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TR2' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TR' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "R3" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TR3' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TR3' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "R LUAR" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRL' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TRL' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
     
     ElseIf cdepartment.Text = "CONTAINER" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TCO' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TCO' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "CSD-ROOM" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TCL' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TCL' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "MAIN WH" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TMH' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TMH' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "GAS" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TGS' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TGS' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    
    ElseIf cdepartment.Text = "CONT. CONS" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TCC' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TCC' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "ASSET" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TAS' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,3)='TAS' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
        
    ElseIf cdepartment.Text = "TAG BLANK" Then
        If rs_so.State = 0 Then
            rs_so.Open "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,2)='TB' and status like '%" & strstatus & "%'", conn
        End If
        lihat = "select*, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where left(tag_no,2)='TB' and status like '%" & strstatus & "%'"
        Set rs_so = conn.Execute(lihat)
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Call lvItem
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus

    End If

End Sub
Sub update()
If Left(ttag_no.Text, 2) = "TB" Then
    ubah = "update tag_stock_opname set location='" & CheckCharacter(tlocation.Text) & "', " & _
    "part_no='" & tpart_no.Text & "', part_name='" & CheckCharacter(tpart_name.Text) & "', " & _
    "inch='" & CheckCharacter(tsize.Text) & "', sch='" & CheckCharacter(tclass.Text) & "'," & _
    "thickness='" & CheckCharacter(tthickness.Text) & "', length='" & CheckCharacter(tlength.Text) & "'," & _
        "width='" & CheckCharacter(twidth.Text) & "'," & _
    "category='" & CheckCharacter(tcategory.Text) & "', " & _
    "tag_code='" & CheckCharacter(ctag_code.Text) & "',status='" & CheckCharacter(cstatus.Text) & "', " & _
    "satuan='" & CheckCharacter(tsatuan.Text) & "', qty_admin='" & Val(tqty.Text) & "', qty_actual='" & Val(tqtyadm.Text) & "'," & _
    "variance='" & Val(tqtyslh.Text) & "' where tag_no='" & ttag_no.Text & "'"
Else
    ubah = "update tag_stock_opname set location='" & CheckCharacter(tlocation.Text) & "', " & _
        "part_no='" & tpart_no.Text & "', part_name='" & CheckCharacter(tpart_name.Text) & "', " & _
        "inch='" & CheckCharacter(tsize.Text) & "', sch='" & CheckCharacter(tclass.Text) & "'," & _
        "thickness='" & CheckCharacter(tthickness.Text) & "', length='" & CheckCharacter(tlength.Text) & "'," & _
        "width='" & CheckCharacter(twidth.Text) & "'," & _
        "category='" & CheckCharacter(tcategory.Text) & "', " & _
        "tag_code='" & CheckCharacter(ctag_code.Text) & "',status='" & CheckCharacter(cstatus.Text) & "', " & _
        "qty_actual='" & Val(tqtyadm.Text) & "'," & _
        "variance='" & Val(tqtyslh.Text) & "' where tag_no='" & ttag_no.Text & "'"
End If
Set rs_so = conn.Execute(ubah)
End Sub

Sub bersih()
For Each A In Me
    If TypeOf A Is TextBox Then A.Text = ""
Next A
ctag_code.Text = ""
cstatus.Text = ""
tqtyslh.Text = ""

End Sub
Sub Warna_List()
Dim I As Long

For I = 1 To ListView1.ListItems.Count
If ListView1.ListItems(I).SubItems(12) = "OK" And Val(ListView1.ListItems(I).SubItems(16)) < 0 Then
    ListView1.ListItems(I).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(1).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(2).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(3).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(4).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(5).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(6).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(7).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(8).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(9).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(10).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(11).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(12).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(13).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(14).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(15).ForeColor = vbRed
    ListView1.ListItems(I).ListSubItems(16).ForeColor = vbRed
  '  ListView1.ListItems(I).ListSubItems(17).ForeColor = vbRed
ElseIf ListView1.ListItems(I).SubItems(12) = "OK" And Val(ListView1.ListItems(I).SubItems(16)) > 0 Then
    ListView1.ListItems(I).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(1).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(2).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(3).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(4).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(5).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(6).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(7).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(8).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(9).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(10).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(11).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(12).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(13).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(14).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(15).ForeColor = vbGreen
    ListView1.ListItems(I).ListSubItems(16).ForeColor = vbGreen
   ' ListView1.ListItems(I).ListSubItems(17).ForeColor = vbGreen
ElseIf ListView1.ListItems(I).SubItems(12) = "OK" And Val(ListView1.ListItems(I).SubItems(12)) = 0 Then
    ListView1.ListItems(I).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(1).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(2).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(3).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(4).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(5).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(6).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(7).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(8).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(9).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(10).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(11).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(12).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(13).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(14).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(15).ForeColor = vbBlue
    ListView1.ListItems(I).ListSubItems(16).ForeColor = vbBlue
    'ListView1.ListItems(I).ListSubItems(17).ForeColor = vbBlue
Else
    ListView1.ListItems(I).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(1).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(2).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(3).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(4).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(5).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(6).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(7).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(8).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(9).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(10).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(11).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(12).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(13).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(14).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(15).ForeColor = vbBlack
    ListView1.ListItems(I).ListSubItems(16).ForeColor = vbBlack
'    ListView1.ListItems(I).ListSubItems(17).ForeColor = vbBlack
End If
Next

End Sub

Public Sub SetLV()
    With ListView1
        .Gridlines = True
        .View = lvwReport
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .MultiSelect = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "Tag No", 0
        .ColumnHeaders.Add 2, , "Tag No", 1000
        .ColumnHeaders.Add 3, , "Part No", 1500
        .ColumnHeaders.Add 4, , "Part Name", 5000
        .ColumnHeaders.Add 5, , "Inch", 1000
        .ColumnHeaders.Add 6, , "Sch", 1000
        .ColumnHeaders.Add 7, , "Thickness", 1000
        .ColumnHeaders.Add 8, , "Length", 1000
        .ColumnHeaders.Add 9, , "Width", 1000
        .ColumnHeaders.Add 10, , "Category", 1500
        .ColumnHeaders.Add 11, , "Location", 1000
        .ColumnHeaders.Add 12, , "Tag Code", 1100
        .ColumnHeaders.Add 13, , "Status", 1100
        .ColumnHeaders.Add 14, , "Satuan", 1100
        .ColumnHeaders.Add 15, , "Qty Admin", 1100
        .ColumnHeaders.Add 16, , "Qty Actual", 1100
        .ColumnHeaders.Add 17, , "Qty Selisih", 1100
        .ColumnHeaders.Add 18, , "qty sheet", 1000
        .Width = 18500
    End With
End Sub
Sub TplGrid()
    Dim Lst As ListItem, nmr As Integer
    If rs_so.State = 0 Then
        rs_so.Open "select *,(qty_actual-qty_admin) as qty_selisih from tag_stock_opname", conn
    End If
    lihat = "select * from tag_stock_opname"
    Set rs_so = conn.Execute(lihat)
    With rs_so
    ListView1.ListItems.Clear
    Do While Not rs_so.EOF
    Call lvItem
    rs_so.MoveNext
    Loop
    End With
   
End Sub

Private Sub cbstatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        filter_lv (cbstatus.Text)
    End If
End Sub

Private Sub ccancel_Click()
bersih
ttag_no.SetFocus
'Form1.PrintForm
End Sub

Private Sub cdepartment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    filter_lv (cbstatus.Text)
    'current_progress
End If


End Sub

Private Sub Command1_Click()
    Dim vexcel As Excel.Application
    Dim obook As Object
    Dim osheet As Object

   'Start a new workbook in Excel
   Set oexcel = CreateObject("Excel.Application")
'   Set oBook =
    oexcel.Workbooks.Add

   'Add data to cells of the first worksheet in the new workbook
   Set osheet = obook.Worksheets(1)
   osheet.Range("A1").Value = "Last Name"
   osheet.Range("B1").Value = "First Name"
   osheet.Range("A1:B1").Font.Bold = True
   osheet.Range("A2").Value = "Doe"
   osheet.Range("B2").Value = "John"

   'Save the Workbook and Quit Excel
   obook.SaveAs "C:\Book1.xls"
   oexcel.Quit
End Sub

Private Sub cstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tsatuan.SelStart = 0
    tsatuan.SelLength = Len(tsatuan.Text)
    tsatuan.SetFocus
End If
End Sub

Private Sub cstatus_Scroll()
If tpart_name.Text = "" Then
    tlocation.SetFocus
End If

End Sub

Private Sub ctag_code_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cstatus.SelStart = 0
    cstatus.SelLength = Len(cstatus.Text)
    cstatus.SetFocus
End If

End Sub

Private Sub export_Click()
    Form3.Show vbModal
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

cdepartment.SetFocus
cstatus.Clear
cstatus.AddItem "OK"
cstatus.AddItem "NO"

cstatus_input.AddItem "OK"
cstatus_input.AddItem "NO"

'ctagcode_input.AddItem "TR"
'ctagcode_input.AddItem "TW"
'ctagcode_input.AddItem "TB"



End Sub

Private Sub Form_Load()
Call db
Call SetLV
If rs_so.State = 1 Then rs_so.Close
rs_so.Open "Select * from tag_stock_opname", conn, adOpenDynamic, adLockOptimistic
'Call TplGrid
Call Warna_List
rs_so.Close
'Set rs_so = Nothing
End Sub

Private Sub Import_Click()
    Form4.Show vbModal
End Sub

Private Sub ListView1_DblClick()
    SendKeys "{Enter}", True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        KeyAscii = 0
    On Error GoTo this
        str = ListView1.SelectedItem.SubItems(1)
        ttag_no.Text = str
        ttag_no.SetFocus
        SendKeys "{Enter}", True
this:
Exit Sub
    End Select
End Sub

Private Sub progess_Click()
Form2.Show vbModal
End Sub

Private Sub tcategory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ctag_code.SelStart = 0
    ctag_code.SelLength = Len(ctag_code.Text)
    ctag_code.SetFocus
End If

End Sub

Private Sub tclass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tthickness.SelStart = 0
        tthickness.SelLength = Len(tthickness.Text)
        tthickness.SetFocus
    End If
End Sub

Private Sub tlength_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        twidth.SelStart = 0
        twidth.SelLength = Len(twidth.Text)
        twidth.SetFocus
    End If
End Sub

Private Sub tlocation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tpart_no.SelStart = 0
        tpart_no.SelLength = Len(tpart_no.Text)
        tpart_no.SetFocus
    End If
End Sub

Private Sub tpart_name_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tsize.SelStart = 0
    tsize.SelLength = Len(tsize.Text)
    tsize.SetFocus
End If

End Sub

Private Sub tpart_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tpart_name.SelStart = 0
    tpart_name.SelLength = Len(tpart_name.Text)
    tpart_name.SetFocus
End If
End Sub

Private Sub tqty_GotFocus()
    If Left(ttag_no.Text, 2) <> "TB" Then
        tqty.Enabled = False
        tqtyadm.SelStart = 0
        tqtyadm.SelLength = Len(tqtyadm.Text)
        tqtyadm.SetFocus
    Else
        tqty.Enabled = True
    End If
End Sub

Private Sub tqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tqtyadm.SelStart = 0
    tqtyadm.SelLength = Len(tqtyadm.Text)
    tqtyadm.SetFocus
End If

End Sub

Private Sub tqty_LostFocus()
    tqty.Enabled = True
End Sub

Private Sub tqtyadm_Change()
'    tqtyslh.Text = Format(Val(tqty.Text), "###,##0.00") - Format(Val(tqtyadm.Text), "###,##0.00")
    'tqtyslh.Text = Format(Val(tqtyadm.Text), "###,##0.00") - Format(Val(tqty.Text), "###,##0.00")
    tqtyslh.Text = Format(Val(tqtyadm.Text) - Val(tqty.Text), "###,##0.00")
'    tqtyslh.Text = Format(tes, "###,##0.00")
    'tes = tqtyslh.Text
End Sub

Private Sub tqtyadm_GotFocus()
    'tqtyslh.Text = Format(Val(tqtyadm.Text), "###,##0.00") - Format(Val(tqty.Text), "###,##0.00")
End Sub

Private Sub tqtyadm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim r As Long
        Dim I As Long
        Dim strtag As String
        If IsNumeric(tqtyadm.Text) = False Then
            MsgBox "Harap memasukkan angka", vbInformation + vbOKOnly, "Information"
            Exit Sub
        End If
        
        Call update
        
        strtag = ttag_no.Text
        Call bersih
        
        Dim Lst As ListItem, nmr As Integer
        'If rs_so.State = 1 Then rs_so.State = 0
        If Mid$(strtag, 4, 1) <> "-" Then
            lihat = "select * from tag_stock_opname where left(tag_no,2)='" & Left(strtag, 2) & "' "
        Else
            lihat = "select * from tag_stock_opname where left(tag_no,3)='" & Left(strtag, 3) & "' "
        End If
        Set rs_so = conn.Execute(lihat)
        
        With rs_so
            ListView1.ListItems.Clear
            Do While Not rs_so.EOF
                Call lvItem
                rs_so.MoveNext
            Loop
        End With
        'Call TplGrid
        
        Call Warna_List
        'Call filter_lv(cbstatus)
       
        With ListView1
             For r = 1 To .ListItems.Count
                 I = Len(Trim(Left(.ListItems(r).SubItems(1), _
                    InStr(1, .ListItems(r).SubItems(1), Trim(strtag), 1))))
                 If I <> 0 Then
                    Set itm = .FindItem(.ListItems(r).SubItems(1))
                    If Not itm Is Nothing Then
                       .ListItems(itm.Index).Selected = True
                        itm.EnsureVisible
                       .SetFocus
                        SendKeys "{LEFT}", True
                    End If
                 End If
             Next r
        End With
        
        'ttag_no.SetFocus
        
        'current_progress
    End If
End Sub

Private Sub tsatuan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tqty.SelStart = 0
    tqty.SelLength = Len(tqty.Text)
    tqty.SetFocus
End If
End Sub

Private Sub tsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Searching(tsearch.Text)
        ctag_code.SetFocus
    End If
End Sub

Private Sub tsize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tclass.SelStart = 0
        tclass.SelLength = Len(tclass.Text)
        tclass.SetFocus
    End If
End Sub

Private Sub ttag_no_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call db
        Cari = "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih from tag_stock_opname where tag_no='" & ttag_no.Text & "'"
        Set rs_so = conn.Execute(Cari)
        If Not rs_so.EOF Then
            tlocation.Text = Format(IIf(IsNull(rs_so!Location), "", rs_so!Location))
            tpart_no.Text = Format(IIf(IsNull(rs_so!part_no), "", rs_so!part_no))
            tpart_name.Text = Format(IIf(IsNull(rs_so!part_name), "", rs_so!part_name))
            tsize.Text = Format(IIf(IsNull(rs_so!INCH), "", rs_so!INCH))
            tclass.Text = Format(IIf(IsNull(rs_so!sch), "", rs_so!sch))
            tthickness.Text = Format(IIf(IsNull(rs_so!Thickness), "", rs_so!Thickness))
            tlength.Text = Format(IIf(IsNull(rs_so!Length), "", rs_so!Length))
            twidth.Text = Format(IIf(IsNull(rs_so!Width), "", rs_so!Width))
            tcategory.Text = Format(IIf(IsNull(rs_so.Fields("category")), "", rs_so.Fields("category")))
            ctag_code.Text = Format(IIf(IsNull(rs_so!tag_code), "", rs_so!tag_code))
            cstatus.Text = Format(IIf(IsNull(rs_so!Status), "", rs_so!Status))
            tsatuan.Text = Format(IIf(IsNull(rs_so!satuan), "", rs_so!satuan))
            tqty.Text = Format(IIf(IsNull(rs_so!qty_admin), 0, rs_so!qty_admin))
            tqtyadm.Text = Format(IIf(IsNull(rs_so.Fields!qty_actual), 0, rs_so.Fields!qty_actual))
            tqtyslh.Text = Format(IIf(IsNull(rs_so!variance), 0, rs_so!variance))
            cstatus.SetFocus
            If Left(ttag_no.Text, 2) = "TB" Or Left(ttag_no.Text, 2) = "TB" Then
                tlocation.SelStart = 0
                tlocation.SelLength = Len(tlocation.Text)
                tlocation.SetFocus
            End If
        Else
            MsgBox "Tag Belum terdaftar mohon laporkan ke sekretariat", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Public Sub Searching(Cari As String)
    Dim Lst As ListItem, nmr As Integer
    sql = "select *, (qty_actual-(ISNULL(qty_admin,0))) as qty_selisih  from tag_stock_opname where tag_no like'%" & Cari & "%' or location like'%" & Cari & "%' or " & _
        "part_no like'%" & Cari & "%' or part_name like'%" & Cari & "%' or category like'%" & Cari & "%' or " & _
        "category like'%" & Cari & "%' or tag_code like'%" & Cari & "%'"
    If rs_so.State = 1 Then rs_so.Close
    rs_so.Open sql, conn, adOpenDynamic, adLockOptimistic
    
    If Not rs_so.EOF Then
        With rs_so
        ListView1.ListItems.Clear
        Do While Not rs_so.EOF
            Set Lst = ListView1.ListItems.Add
            Lst.Text = rs_so!tag_no
            Lst.SubItems(1) = Format(IIf(IsNull(rs_so!tag_no), "", rs_so!tag_no))
            Lst.SubItems(2) = Format(IIf(IsNull(rs_so!part_no), "", rs_so!part_no))
            Lst.SubItems(3) = Format(IIf(IsNull(rs_so!part_name), "", rs_so!part_name))
            Lst.SubItems(4) = Format(IIf(IsNull(rs_so!INCH), "", rs_so!INCH))
            Lst.SubItems(5) = Format(IIf(IsNull(rs_so!sch), "", rs_so!sch))
            Lst.SubItems(6) = Format(IIf(IsNull(rs_so!Thickness), "", rs_so!Thickness))
            Lst.SubItems(7) = Format(IIf(IsNull(rs_so!Length), "", rs_so!Length))
            Lst.SubItems(8) = Format(IIf(IsNull(rs_so!Width), "", rs_so!Width))
            Lst.SubItems(9) = Format(IIf(IsNull(rs_so!Category), "", rs_so!Category))
            Lst.SubItems(10) = Format(IIf(IsNull(rs_so!Location), "", rs_so!Location))
            Lst.SubItems(11) = Format(IIf(IsNull(rs_so!tag_code), "", rs_so!tag_code))
            Lst.SubItems(12) = Format(IIf(IsNull(rs_so!Status), "", rs_so!Status))
            Lst.SubItems(13) = Format(IIf(IsNull(rs_so!satuan), "", rs_so!satuan))
            Lst.SubItems(14) = Format(IIf(IsNull(rs_so!qty_actual), "", rs_so!qty_actual), "###,##0.00")
            Lst.SubItems(15) = Format(IIf(IsNull(rs_so!qty_admin), "", rs_so!qty_admin), "###,##0.00")
            Lst.SubItems(16) = Format(IIf(IsNull(rs_so!variance), "", rs_so!variance), "###,##0.00")
            rs_so.MoveNext
        Loop
        End With
        Call Warna_List
        ttag_no.SetFocus
    Else
        MsgBox "Data tidak ada", vbOKOnly + vbInformation, "Informasi"
    End If

End Sub

Public Sub lvItem()
    Set Lst = ListView1.ListItems.Add
    Lst.Text = rs_so!tag_no
    Lst.SubItems(1) = Format(IIf(IsNull(rs_so!tag_no), "", rs_so!tag_no))
    Lst.SubItems(2) = Format(IIf(IsNull(rs_so!part_no), "", rs_so!part_no))
    Lst.SubItems(3) = Format(IIf(IsNull(rs_so!part_name), "", rs_so!part_name))
    Lst.SubItems(4) = Format(IIf(IsNull(rs_so!INCH), "", rs_so!INCH))
    Lst.SubItems(5) = Format(IIf(IsNull(rs_so!sch), "", rs_so!sch))
    Lst.SubItems(6) = Format(IIf(IsNull(rs_so!Thickness), "", rs_so!Thickness))
    Lst.SubItems(7) = Format(IIf(IsNull(rs_so!Length), "", rs_so!Length))
    Lst.SubItems(8) = Format(IIf(IsNull(rs_so!Width), "", rs_so!Width))
    Lst.SubItems(9) = Format(IIf(IsNull(rs_so!Category), "", rs_so!Category))
    Lst.SubItems(10) = Format(IIf(IsNull(rs_so!Location), "", rs_so!Location))
    Lst.SubItems(11) = Format(IIf(IsNull(rs_so!tag_code), "", rs_so!tag_code))
    Lst.SubItems(12) = Format(IIf(IsNull(rs_so!Status), "", rs_so!Status))
    Lst.SubItems(13) = Format(IIf(IsNull(rs_so!satuan), "", rs_so!satuan))
    Lst.SubItems(14) = Format(IIf(IsNull(rs_so.Fields("qty_admin")), "", rs_so.Fields("qty_admin")), "###,##0.00")
    Lst.SubItems(15) = Format(IIf(IsNull(rs_so.Fields("qty_actual")), "", rs_so.Fields("qty_actual")), "###,##0.00")
    Lst.SubItems(16) = Format(IIf(IsNull(rs_so.Fields("variance")), "", rs_so.Fields("variance")), "###,##0.00")
End Sub

Private Sub ttag_no_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tthickness_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tlength.SelStart = 0
        tlength.SelLength = Len(tlength.Text)
        tlength.SetFocus
    End If
End Sub

Private Sub twidth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tcategory.SelStart = 0
        tcategory.SelLength = Len(tcategory.Text)
        tcategory.SetFocus
    End If
End Sub

