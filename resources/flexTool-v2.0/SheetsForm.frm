VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetsForm 
   Caption         =   "Sheets"
   ClientHeight    =   12480
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   3000
   OleObjectBlob   =   "SheetsForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "SheetsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 




' Sheet selection userform
' Written: November 11, 2017
' Author:  Simo Rissanen
'
' Resize userform
' Written: February 14, 2011
' Author:  Leith Ross
' Modified by: Simo Rissanen

#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function GetWindowLong _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long) _
  As Long
               
Private Declare PtrSafe Function SetWindowLong _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) _
  As Long

#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) _
  As Long
               
Private Declare Function SetWindowLong _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) _
  As Long
#End If

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

Public Sub InitForm()
    Call UserForm_Initialize
    Call UserForm_Activate
    Call SetInitHeight
    Call UserForm_Resize
End Sub

Public Sub UserForm_Resize()
    Call scaleContent(Me.Height)
End Sub


Public Sub MakeFormResizable()

  Dim lStyle As Long
  Dim hwnd As Long
  Dim RetVal
  
    hwnd = FindWindow("ThunderDFrame", SheetsForm.Caption)
  
    'Get the basic window style
     lStyle = GetWindowLong(hwnd, GWL_STYLE) Or WS_THICKFRAME

    'Set the basic window styles
     RetVal = SetWindowLong(hwnd, GWL_STYLE, lStyle)
End Sub

Private Function IsInArray(valToBeFound As Variant, Arr As Variant) As Boolean

Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In Arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Public Sub SetInitHeight()
' Scale size based on content
Dim availableHeight As Double
Dim sum As Double

   ListBox1.IntegralHeight = False
   ListBox2.IntegralHeight = False
   ListBox3.IntegralHeight = False
   ListBox4.IntegralHeight = False
   
   sum = ListBox1.ListCount + ListBox2.ListCount + ListBox3.ListCount + ListBox4.ListCount
   targeth = sum * 13 + TextBox1.Height + TextBox2.Height + TextBox3.Height + TextBox4.Height
   If targeth < 100 Then
       targeth = 100
   ElseIf targeth > Application.Height * 0.9 Then
       targeth = Application.Height * 0.9
   End If
   
   Me.Height = targeth
   availableHeight = Me.Height - TextBox1.Height - TextBox2.Height - TextBox3.Height - TextBox4.Height - 20

   If availableHeight < 100 Then
       availableHeight = 100
   End If
   
   ListBox1.Height = availableHeight * ListBox1.ListCount / sum
   ListBox2.Height = availableHeight * ListBox2.ListCount / sum
   ListBox3.Height = availableHeight * ListBox3.ListCount / sum
   ListBox4.Height = availableHeight * ListBox4.ListCount / sum
   ListBox1.IntegralHeight = True
   ListBox2.IntegralHeight = True
   ListBox3.IntegralHeight = True
   ListBox4.IntegralHeight = True
   
   TextBox1.Top = 0
   ListBox1.Top = TextBox1.Top + TextBox1.Height
   TextBox2.Top = ListBox1.Top + ListBox1.Height
   ListBox2.Top = TextBox2.Top + TextBox2.Height
   TextBox3.Top = ListBox2.Top + ListBox2.Height
   ListBox3.Top = TextBox3.Top + TextBox3.Height
   TextBox4.Top = ListBox3.Top + ListBox3.Height
   ListBox4.Top = TextBox4.Top + TextBox4.Height
   
   TextBox1.Width = Me.Width - 11
   ListBox1.Width = Me.Width - 11
   TextBox2.Width = Me.Width - 11
   ListBox2.Width = Me.Width - 11
   TextBox3.Width = Me.Width - 11
   ListBox3.Width = Me.Width - 11
   TextBox4.Width = Me.Width - 11
   ListBox4.Width = Me.Width - 11
End Sub

Private Sub scaleContent(frameH As Double)

Dim availableHeight As Double
Dim sum As Double

   TextBox1.Height = 16
   TextBox2.Height = 16
   TextBox3.Height = 16
   TextBox4.Height = 16

   ListBox1.IntegralHeight = False
   ListBox2.IntegralHeight = False
   ListBox3.IntegralHeight = False
   ListBox4.IntegralHeight = False
   availableHeight = frameH - TextBox1.Height - TextBox2.Height - TextBox3.Height - TextBox4.Height - 20
   
   If availableHeight < 100 Then
       availableHeight = 100
   End If
   
   sum = ListBox1.ListCount + ListBox2.ListCount + ListBox3.ListCount + ListBox4.ListCount
   targeth = availableHeight * ListBox1.ListCount / sum
   
   ListBox1.Height = availableHeight * ListBox1.ListCount / sum
   ListBox2.Height = availableHeight * ListBox2.ListCount / sum
   ListBox3.Height = availableHeight * ListBox3.ListCount / sum
   ListBox4.Height = availableHeight * ListBox4.ListCount / sum
   ListBox1.IntegralHeight = True
   ListBox2.IntegralHeight = True
   ListBox3.IntegralHeight = True
   ListBox4.IntegralHeight = True
   
   totalboxh = ListBox1.Height + ListBox2.Height + ListBox3.Height + ListBox4.Height _
               + TextBox1.Height + TextBox2.Height + TextBox3.Height + TextBox4.Height
               
   diff = frameH - totalboxh - 28

   If diff > 4 Then
       TextBox1.Height = TextBox1.Height + (diff / 4)
       TextBox2.Height = TextBox2.Height + (diff / 4)
       TextBox3.Height = TextBox3.Height + Int(diff / 4)
       TextBox4.Height = TextBox4.Height + Int(diff / 4)
   End If
   
   TextBox1.Top = 0
   ListBox1.Top = TextBox1.Top + TextBox1.Height
   TextBox2.Top = ListBox1.Top + ListBox1.Height
   ListBox2.Top = TextBox2.Top + TextBox2.Height
   TextBox3.Top = ListBox2.Top + ListBox2.Height
   ListBox3.Top = TextBox3.Top + TextBox3.Height
   TextBox4.Top = ListBox3.Top + ListBox3.Height
   ListBox4.Top = TextBox4.Top + TextBox4.Height
   
   TextBox1.Width = Me.Width - 11
   ListBox1.Width = Me.Width - 11
   TextBox2.Width = Me.Width - 11
   ListBox2.Width = Me.Width - 11
   TextBox3.Width = Me.Width - 11
   ListBox3.Width = Me.Width - 11
   TextBox4.Width = Me.Width - 11
   ListBox4.Width = Me.Width - 11
   
End Sub


Private Sub UserForm_Activate()
    
    With Me
      .StartUpPosition = 0
      .Top = Application.Top + 25
      .Left = Application.Left + Application.Width - Me.Width - 25
      .Height = Application.Height * 0.9

   End With
   
   TextBox1.Height = 16
   TextBox2.Height = 16
   TextBox3.Height = 16
   TextBox4.Height = 16
   
   ThisWorkbook.Activate
   Call MakeFormResizable
   DoEvents
  
End Sub

Private Sub UserForm_Initialize()
    Dim N As Long
    Dim shortestname As String
    Dim added() As String
    
    Call setVariables
    ThisWorkbook.Activate
    
    ListBox1.Clear
    ListBox2.Clear
    ListBox3.Clear
    ListBox4.Clear
    
    TextBox1.Value = "FLEXIBILITY"
    TextBox1.BackColor = vbGreen
    TextBox2.Value = "OPERATIONS"
    TextBox2.BackColor = vbCyan
    TextBox3.Value = "COSTS"
    TextBox3.BackColor = vbMagenta
    TextBox4.Value = "NODES"
    TextBox4.BackColor = vbRed
    
    ReDim added(1)
       
    For N = 1 To ThisWorkbook.Sheets.count
        If IsInArray(ThisWorkbook.Sheets(N).Name, added) Then
           Debug.Print ("sheet already added")
        Else
            If InStr(ThisWorkbook.Sheets(N).Name, "summary") Or _
                ThisWorkbook.Sheets(N).Name = "node" Or _
                ThisWorkbook.Sheets(N).Name = "events" Or _
                ThisWorkbook.Sheets(N).Name = "node_plot" Then
                ListBox1.AddItem ThisWorkbook.Sheets(N).Name
                added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                ReDim Preserve added(UBound(added) + 1)
                
            ElseIf InStr(ThisWorkbook.Sheets(N).Name, "genType") Or _
                   InStr(ThisWorkbook.Sheets(N).Name, "storageContent") Or _
                   InStr(ThisWorkbook.Sheets(N).Name, "onlineUnit") Or _
                   InStr(ThisWorkbook.Sheets(N).Name, "inertiaUnit") Or _
                   InStr(ThisWorkbook.Sheets(N).Name, "reserveUnit") Or _
                   ThisWorkbook.Sheets(N).Name = "transfers_t" Or _
                   InStr(ThisWorkbook.Sheets(N).Name, "genUnit") Then
                ListBox2.AddItem ThisWorkbook.Sheets(N).Name
                added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                ReDim Preserve added(UBound(added) + 1)

            ' costs
            ElseIf ThisWorkbook.Sheets(N).Name = "costs" Or _
                       InStr(ThisWorkbook.Sheets(N).Name, "invest") Or _
                       InStr(ThisWorkbook.Sheets(N).Name, "costs") Then
                    ListBox3.AddItem ThisWorkbook.Sheets(N).Name
                    added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                    ReDim Preserve added(UBound(added) + 1)
            ElseIf InStr(ThisWorkbook.Sheets(N).Name, "node_t") Then
                    ListBox4.AddItem ThisWorkbook.Sheets(N).Name
                    added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                    ReDim Preserve added(UBound(added) + 1)
            Else
        
            For G = 1 To UBound(gridfilter)
                ' Flexibility
                If ThisWorkbook.Sheets(N).Name = "rampRoom_1h_" & gridfilter(G) Or _
                   ThisWorkbook.Sheets(N).Name = "rampRoom_4h_" & gridfilter(G) Or _
                   ThisWorkbook.Sheets(N).Name = "rampRoom_1h_" & gridfilter(G) & "_plot" Or _
                   ThisWorkbook.Sheets(N).Name = "rampRoom_4h_" & gridfilter(G) & "_plot" Or _
                   ThisWorkbook.Sheets(N).Name = "duration_" & gridfilter(G) Or _
                   ThisWorkbook.Sheets(N).Name = "duration_" & gridfilter(G) & "_plot" Or _
                   ThisWorkbook.Sheets(N).Name = "durationRamp_" & gridfilter(G) Or _
                   ThisWorkbook.Sheets(N).Name = "durationRamp_" & gridfilter(G) & "_plot" Then
                   ListBox1.AddItem ThisWorkbook.Sheets(N).Name
                   added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                   ReDim Preserve added(UBound(added) + 1)

                ' operations
                ElseIf ThisWorkbook.Sheets(N).Name = "units_" & gridfilter(G) Or _
                       ThisWorkbook.Sheets(N).Name = "units_" & gridfilter(G) & "_plot" Or _
                       ThisWorkbook.Sheets(N).Name = "transfers_" & gridfilter(G) Or _
                       ThisWorkbook.Sheets(N).Name = "transfers_" & gridfilter(G) & "_plot" Or _
                       ThisWorkbook.Sheets(N).Name = "grid_t_" & gridfilter(G) Then
                    ListBox2.AddItem ThisWorkbook.Sheets(N).Name
                    added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                    ReDim Preserve added(UBound(added) + 1)

                'nodes
                ElseIf InStr(ThisWorkbook.Sheets(N).Name, "rampRoom_1h_" & gridfilter(G)) Or _
                   InStr(ThisWorkbook.Sheets(N).Name, "rampRoom_4h_" & gridfilter(G)) Then
                    ListBox4.AddItem ThisWorkbook.Sheets(N).Name
                    added(UBound(added)) = ThisWorkbook.Sheets(N).Name
                    ReDim Preserve added(UBound(added) + 1)
                Else
                    Debug.Print (ThisWorkbook.Sheets(N).Name & " not added to sheetsform")
                End If
            
            Next G
            End If
        End If
    Next N

    ListBox1.SetFocus
    ListBox2.SetFocus
    ListBox3.SetFocus
    ListBox4.SetFocus
    
    Call MakeFormResizable
    
End Sub

Private Sub ListBox1_Click()
    ThisWorkbook.Activate
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    Me.ListBox2.MultiSelect = fmMultiSelectSingle
    Me.ListBox3.MultiSelect = fmMultiSelectSingle
    Me.ListBox4.MultiSelect = fmMultiSelectSingle
    Me.ListBox2.Value = ""
    Me.ListBox3.Value = ""
    Me.ListBox4.Value = ""
    ThisWorkbook.Sheets(ListBox1.List(ListBox1.ListIndex)).Select
    
End Sub
Private Sub ListBox2_Click()
    ThisWorkbook.Activate
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    Me.ListBox2.MultiSelect = fmMultiSelectSingle
    Me.ListBox3.MultiSelect = fmMultiSelectSingle
    Me.ListBox4.MultiSelect = fmMultiSelectSingle
    Me.ListBox1.Value = ""
    Me.ListBox3.Value = ""
    Me.ListBox4.Value = ""
    ThisWorkbook.Sheets(ListBox2.List(ListBox2.ListIndex)).Select
End Sub
Private Sub ListBox3_Click()
    ThisWorkbook.Activate
    Me.ListBox2.MultiSelect = fmMultiSelectSingle
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    Me.ListBox3.MultiSelect = fmMultiSelectSingle
    Me.ListBox4.MultiSelect = fmMultiSelectSingle
    Me.ListBox2.Value = ""
    Me.ListBox1.Value = ""
    Me.ListBox4.Value = ""
    ThisWorkbook.Sheets(ListBox3.List(ListBox3.ListIndex)).Select
End Sub
Private Sub ListBox4_Click()
    ThisWorkbook.Activate
    Me.ListBox2.MultiSelect = fmMultiSelectSingle
    Me.ListBox3.MultiSelect = fmMultiSelectSingle
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    Me.ListBox4.MultiSelect = fmMultiSelectSingle
    Me.ListBox2.Value = ""
    Me.ListBox3.Value = ""
    Me.ListBox1.Value = ""
    ThisWorkbook.Sheets(ListBox4.List(ListBox4.ListIndex)).Select
End Sub

Private Sub TextBox1_Enter()
    UserForm_Initialize
End Sub
Private Sub TextBox2_Enter()
    UserForm_Initialize
End Sub
Private Sub TextBox3_Enter()
    UserForm_Initialize
End Sub
Private Sub TextBox4_Enter()
    UserForm_Initialize
End Sub

Private Sub Listbox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 40 Then
        If ListBox1.ListIndex + 1 = ListBox1.ListCount Then
            ThisWorkbook.Sheets(Me.ListBox2.List(Me.ListBox2.TopIndex)).Select
            Me.ListBox2.SetFocus
        End If
    End If
End Sub

Private Sub Listbox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 40 Then
        If ListBox2.ListIndex + 1 = ListBox2.ListCount Then
            ThisWorkbook.Sheets(Me.ListBox3.List(Me.ListBox3.TopIndex)).Select
            Me.ListBox3.SetFocus
        End If
    
    ElseIf KeyCode = 38 Then
        If ListBox2.ListIndex = ListBox2.TopIndex Then
            ThisWorkbook.Sheets(Me.ListBox1.List(Me.ListBox1.ListCount - 1)).Select
            Me.ListBox1.SetFocus
        End If
    End If
End Sub

Private Sub Listbox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 40 Then
        If ListBox3.ListIndex + 1 = ListBox3.ListCount Then
            ThisWorkbook.Sheets(Me.ListBox4.List(Me.ListBox4.TopIndex)).Select
            Me.ListBox4.SetFocus
        End If

    
    ElseIf KeyCode = 38 Then
        If ListBox3.ListIndex = ListBox3.TopIndex Then
            ThisWorkbook.Sheets(Me.ListBox2.List(Me.ListBox2.ListCount - 1)).Select
            Me.ListBox2.SetFocus
        End If
    End If
End Sub

Private Sub Listbox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 38 Then
        If ListBox4.ListIndex = ListBox4.TopIndex Then
            ThisWorkbook.Sheets(Me.ListBox3.List(Me.ListBox3.ListCount - 1)).Select
            Me.ListBox3.SetFocus
        End If
    End If
End Sub


