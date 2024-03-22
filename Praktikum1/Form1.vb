Imports System.Globalization

Public Class FrmHitungNilai
    Protected Attribute As Char = Nothing

    Private Sub KeyboardValidation(sender As Object, e As KeyPressEventArgs) Handles TxtJumlah.KeyPress
        e.Handled = True
    End Sub

    Private Sub InputNum(sender As Object, e As MouseEventArgs) Handles NumDot.MouseClick, Num0.MouseClick, Num1.MouseClick, Num2.MouseClick, Num3.MouseClick, Num4.MouseClick, Num5.MouseClick, Num6.MouseClick, Num7.MouseClick, Num8.MouseClick, Num9.MouseClick
        Dim clickedNum As String = DirectCast(sender, Control).Tag
        Dim digit As Char = clickedNum

        If TxtJumlah.Text = "0" AndAlso digit <> "." Then
            TxtJumlah.Text = digit
        ElseIf TxtJumlah.TextLength < 12 Then
            TxtJumlah.AppendText(digit)
        End If
    End Sub

    Private Sub DeleteNum(sender As Object, e As MouseEventArgs) Handles BtDelete.MouseClick
        If TxtJumlah.TextLength > 0 Then
            TxtJumlah.Text = TxtJumlah.Text.Remove(TxtJumlah.TextLength - 1, 1)
        End If
    End Sub

    Private Sub ClearNum(sender As Object, e As MouseEventArgs) Handles BtClear.MouseClick
        ClearNum()
    End Sub

    Private Sub ClearNum()
        TxtJumlah.Clear()
    End Sub

    Private Sub SetDefaultZero(sender As Object, e As EventArgs) Handles TxtJumlah.TextChanged
        If TxtJumlah.TextLength = 0 Then
            TxtJumlah.Text = "0"
        End If
    End Sub

    Private Function SetValueTotalSementara(values() As Double) As String
        Dim result As String = ""

        If values.Length > 1 AndAlso Not String.IsNullOrEmpty(GetAttribute().ToString) Then
            result = values(0)

            For i As Integer = 1 To values.Length - 1
                result &= $" {GetAttribute()} {values(i)}"
            Next
        End If

        Return result
    End Function

    Private Function CalculateTotal(values() As Double) As Double
        Dim total As Double = values(0)

        If values.Length > 1 Then
            For i As Integer = 1 To values.Length - 1
                Dim item As Double = values(i)
                Select Case GetAttribute().ToString()
                    Case "+"
                        total += item
                    Case "-"
                        total = item - total
                    Case "*"
                        total *= item
                    Case "/"
                        If item <> 0 Then
                            total = item / total
                        End If
                End Select
            Next
        End If

        Return total
    End Function

    Private Sub SetAttribute(Attribute As Char)
        Me.Attribute = Attribute
    End Sub

    Private Function GetAttribute() As Char
        Return Attribute
    End Function

    Private Sub OperatorButtonClick(sender As Object, e As MouseEventArgs) Handles BtTambah.MouseClick, BtKurang.MouseClick, BtKali.MouseClick, BtBagi.MouseClick
        Dim operatorButton As Button = DirectCast(sender, Button)
        SetAttribute(operatorButton.Tag()(0))
        JumlahSementara.Text = TxtJumlah.Text
        ClearNum()
    End Sub

    Private Sub CalculateResult(sender As Object, e As MouseEventArgs) Handles BtJumlah.MouseClick
        Dim values() As Double
        Dim result As String = "0"

        If Not String.IsNullOrEmpty(GetAttribute().ToString) AndAlso Not String.IsNullOrEmpty(JumlahSementara.Text) AndAlso IsNumeric(JumlahSementara.Text) Then
            values = {CDbl(Val(TxtJumlah.Text)), CDbl(Val(JumlahSementara.Text))}
            result = CalculateTotal(values)
            JumlahSementara.Text = SetValueTotalSementara(values)
        End If

        TxtJumlah.Text = result
    End Sub

    Private Sub ClearAll(sender As Object, e As MouseEventArgs) Handles BtClearAll.MouseClick
        ClearNum()
        JumlahSementara.Text = ""
    End Sub
End Class