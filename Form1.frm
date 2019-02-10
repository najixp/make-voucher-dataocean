VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrimaryRS As Recordset

Private Sub Command1_Click()
Dim MyRec As ADODB.Recordset
Dim MyRec1 As ADODB.Recordset
Dim MyRec2 As ADODB.Recordset
Dim MyRec1acc As ADODB.Recordset
Dim MyRec2acc As ADODB.Recordset
Dim x As String
Dim vDrirection As String
Dim sDrirection1, sDrirection2

List1.Clear

fOpenRSSQLacc MyRec1acc, "select * from VoucherHead"
fOpenRSSQLacc MyRec2acc, "select * from VoucherDetail"

fOpenRSSQL MyRec, "select * from MVouchers"
Do While Not MyRec.EOF



fOpenRSSQL MyRec1, "select * from trnhead where trnid=" & Val(MyRec.Fields("TrnID").Value) & " and najipost is null order by headid"
If MyRec1.EOF = False Then
    Do While Not MyRec1.EOF
        MyRec1acc.AddNew
        MyRec1acc.Fields("TrnID").Value = MyRec.Fields("AccTrnID").Value
        MyRec1acc.Fields("Number").Value = MyRec1.Fields("StrID").Value & MyRec1.Fields("invno").Value
        MyRec1acc.Fields("SrAmount").Value = Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)
        MyRec1acc.Fields("RegDate").Value = MyRec1.Fields("RegDate").Value
        MyRec1acc.Fields("RegUser").Value = MyRec1.Fields("RegUser").Value
        MyRec1acc.Fields("VchDate").Value = MyRec1.Fields("InvDate").Value
        MyRec1acc.Update
Dim rs2 As New ADODB.Recordset
rs2.Open "SELECT     MAX(id) AS Expr1 " & _
                "FROM VoucherHead ", CN, adOpenDynamic, adLockOptimistic
                
    If rs2.EOF = False Then
        If IsNull(rs2.Fields("Expr1")) = False Then
            x = rs2.Fields("Expr1")
        End If
    End If
    rs2.Close
vDrirection = ""

        
If Trim(MyRec.Fields("trncode").Value) = "3" Then
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 26
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value))
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 4168
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value * -1
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 3043
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value * -1
        MyRec2acc.Update
End If
If Trim(MyRec.Fields("trncode").Value) = "4" Then
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 2842
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 3042
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 26
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)) * -1
        MyRec2acc.Update
End If
If Trim(MyRec.Fields("trncode").Value) = "1" Then
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value))
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 4168
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value * -1
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 3043
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value * -1
        MyRec2acc.Update
End If
If Trim(MyRec.Fields("trncode").Value) = "2" Then
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 2842
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = 3042
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value
        MyRec2acc.Update
        MyRec2acc.AddNew
        MyRec2acc.Fields("ID").Value = x
        MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
        MyRec2acc.Fields("SubCostID").Value = 0
        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
        MyRec2acc.Fields("Reverse").Value = 1
        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)) * -1
        MyRec2acc.Update
End If
'If Trim(MyRec.Fields("trncode").Value) = "5" Then
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 2778
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 3042
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 26
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)) * -1
'        MyRec2acc.Update
'End If
'If Trim(MyRec.Fields("trncode").Value) = "6" Then
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 26
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value))
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 2799
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value * -1
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 3043
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value * -1
'        MyRec2acc.Update
'End If
'If Trim(MyRec.Fields("trncode").Value) = "A" Then
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 2778
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 3042
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)) * -1
'        MyRec2acc.Update
'
'End If
'
'If Trim(MyRec.Fields("trncode").Value) = "B" Then
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value))
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 2793
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value * -1
'        MyRec2acc.Update
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        MyRec2acc.Fields("AccountID").Value = 3043
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value * -1
'        MyRec2acc.Update
'End If
        
        
    
    
    
    
    MyRec1.Fields("najipost").Value = x
    MyRec1.Update
    MyRec1.MoveNext
    Me.Caption = x
    DoEvents
    
    Loop
End If
MyRec1.Close


List1.AddItem Trim(MyRec.Fields("TrnID").Value) & " - " & Trim(MyRec.Fields("trncode").Value) & " - " & Trim(MyRec.Fields("trnname").Value) & " - " & Trim(MyRec.Fields("AccTrnID").Value) & " - " & Trim(MyRec.Fields("AccTrnCode").Value) & " - " & Trim(MyRec.Fields("AccTrnName").Value) & " - " & Trim(MyRec.Fields("TrnSign").Value) & " - " & Trim(MyRec.Fields("AccDebit1").Value) & " - " & Trim(MyRec.Fields("AccCredit1").Value) & " - " & Trim(MyRec.Fields("AccDebit2").Value) & " - " & Trim(MyRec.Fields("AccCredit2").Value)

MyRec.MoveNext
DoEvents
Loop
MyRec.Close
MsgBox "done"


'        MyRec2acc.Fields("ID").Value = x
'        If IsNull(MyRec.Fields("AccDebit1ID").Value) = False Then
'            MyRec2acc.Fields("AccountID").Value = MyRec.Fields("AccDebit1ID").Value
'        Else
'            MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
'        End If
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        If IsNull(MyRec.Fields("AccDebit2ID").Value) = False Then
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value
'        vDrirection = "D"
'        Else
'        MyRec2acc.Fields("SrAmount").Value = Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)
'        End If
'
'        MyRec2acc.Update
'
'        MyRec2acc.AddNew
'        MyRec2acc.Fields("ID").Value = x
'        If IsNull(MyRec.Fields("AccCredit1ID").Value) = False Then
'            MyRec2acc.Fields("AccountID").Value = MyRec.Fields("AccCredit1ID").Value
'        Else
'            MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
'        End If
'        MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'        MyRec2acc.Fields("SubCostID").Value = 0
'        MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'        MyRec2acc.Fields("Reverse").Value = 1
'        If IsNull(MyRec.Fields("AccCredit2ID").Value) = False Then
'        MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvNetAmnt").Value * -1
'        vDrirection = "C"
'        Else
'        MyRec2acc.Fields("SrAmount").Value = Val(Val(MyRec1.Fields("InvTotalVatValue").Value) + Val(MyRec1.Fields("InvNetAmnt").Value)) * -1
'        End If
'        MyRec2acc.Update
'
'        If vDrirection = "D" Then
'            MyRec2acc.AddNew
'            MyRec2acc.Fields("ID").Value = x
'            If IsNull(MyRec.Fields("AccDebit2ID").Value) = False Then
'                MyRec2acc.Fields("AccountID").Value = MyRec.Fields("AccDebit2ID").Value
'            Else
'                MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
'            End If
'            MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'            MyRec2acc.Fields("SubCostID").Value = 0
'            MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'            MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'            MyRec2acc.Fields("Reverse").Value = 1
'            MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value * -1
'            MyRec2acc.Update
'        End If
'        If vDrirection = "C" Then
'            MyRec2acc.AddNew
'            MyRec2acc.Fields("ID").Value = x
'            If IsNull(MyRec.Fields("AccCredit2ID").Value) = False Then
'                MyRec2acc.Fields("AccountID").Value = MyRec.Fields("AccCredit2ID").Value
'            Else
'                MyRec2acc.Fields("AccountID").Value = MyRec1.Fields("AcctID").Value
'            End If
'            MyRec2acc.Fields("CostID").Value = MyRec1.Fields("StrID").Value
'            MyRec2acc.Fields("SubCostID").Value = 0
'            MyRec2acc.Fields("Detailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'            MyRec2acc.Fields("ArbDetailes").Value = Trim(MyRec.Fields("trnname").Value) & " " & MyRec1.Fields("invno").Value
'            MyRec2acc.Fields("Reverse").Value = 1
'            MyRec2acc.Fields("SrAmount").Value = MyRec1.Fields("InvTotalVatValue").Value
'            MyRec2acc.Update
'        End If


End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
MKConnStrSQL
ConnectToServer
End Sub

