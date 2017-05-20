Attribute VB_Name = "Module1"
Option Explicit
Dim cnf As Integer




Public Sub ex()
cnf = MsgBox("Are you sure?", vbYesNo, "Exit")
If (cnf = vbYes) Then
Form1.Hide
End If
End Sub


Public Function mylogic()

'If value X is true in all buttons.


If Form1.Command1.Caption = "X" And Form1.Command2.Caption = "X" And Form1.Command3.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"
ElseIf Form1.Command4.Caption = "X" And Form1.Command5.Caption = "X" And Form1.Command6.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"
ElseIf Form1.Command7.Caption = "X" And Form1.Command8.Caption = "X" And Form1.Command9.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"
ElseIf Form1.Command1.Caption = "X" And Form1.Command4.Caption = "X" And Form1.Command7.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"
ElseIf Form1.Command2.Caption = "X" And Form1.Command5.Caption = "X" And Form1.Command8.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"
ElseIf Form1.Command3.Caption = "X" And Form1.Command6.Caption = "X" And Form1.Command9.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"

'if value O is true in all buttons.

ElseIf Form1.Command1.Caption = "O" And Form1.Command2.Caption = "O" And Form1.Command3.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"
ElseIf Form1.Command4.Caption = "O" And Form1.Command5.Caption = "O" And Form1.Command6.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"
ElseIf Form1.Command7.Caption = "O" And Form1.Command8.Caption = "O" And Form1.Command9.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"
ElseIf Form1.Command1.Caption = "O" And Form1.Command4.Caption = "O" And Form1.Command7.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"
ElseIf Form1.Command2.Caption = "O" And Form1.Command5.Caption = "O" And Form1.Command8.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"
ElseIf Form1.Command3.Caption = "O" And Form1.Command6.Caption = "O" And Form1.Command9.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"

'for cross value O


ElseIf Form1.Command1.Caption = "O" And Form1.Command5.Caption = "O" And Form1.Command9.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"
ElseIf Form1.Command3.Caption = "O" And Form1.Command5.Caption = "O" And Form1.Command7.Caption = "O" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text1.Text, vbInformation, "Winner"


'for cross value tru with X


ElseIf Form1.Command1.Caption = "X" And Form1.Command5.Caption = "X" And Form1.Command9.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"
ElseIf Form1.Command3.Caption = "X" And Form1.Command5.Caption = "X" And Form1.Command7.Caption = "X" Then
MsgBox "And the winner is Miss/Mr. " & Form3.Text2.Text, vbInformation, "Winner"



ElseIf (Form1.Command1.Enabled = False) And (Form1.Command2.Enabled = False) And (Form1.Command3.Enabled = False) And (Form1.Command4.Enabled = False) And (Form1.Command5.Enabled = False) And (Form1.Command6.Enabled = False) And (Form1.Command7.Enabled = False) And (Form1.Command8.Enabled = False) And (Form1.Command9.Enabled = False) Then
MsgBox "Nobudy is Winner , And my Auther is nobudy LOL...", vbInformation
End If
End Function



'A Successfull End Of Project.
'Thank You God for giving me the ability to complete this.














