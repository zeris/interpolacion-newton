Public Class Form1
    Dim n, k, l, i As Integer
    Dim m, s As Single
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        tbN.Text = ""
        tbValInterpolar.Text = ""
        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Close()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        n = tbN.Text
        Dim x(n), y(n), vaInterpolar As Single
        k = 0
        vaInterpolar = tbValInterpolar.Text
        l = (n * 2) - 1
        While k < l
            DataGridView1.Rows.Add()
            k = k + 1
        End While
        k = 0
        While k < 3
            l = 0
            If k = 0 Then
                While l < n
                    If l = 0 Then
                        DataGridView1.Rows(l).Cells(k).Value = l
                    Else
                        DataGridView1.Rows(l + l).Cells(k).Value = l
                    End If
                    l = l + 1
                End While
            ElseIf k = 1 Then
                While l < n
                    x(l) = InputBox("Dame el valor de x" + l.ToString())
                    If l = 0 Then
                        DataGridView1.Rows(l).Cells(k).Value = x(l)
                    Else
                        DataGridView1.Rows(l + l).Cells(k).Value = x(l)
                        If vaInterpolar > x(l - 1) And vaInterpolar < x(l) Then
                            i = l - 1
                            DataGridView1.Rows(l + (l - 1)).Cells(k).Value = vaInterpolar
                        End If
                    End If
                    l = l + 1
                End While
            Else
                While l < n
                    y(l) = InputBox("Dame el valor de y" + l.ToString())
                    If l = 0 Then
                        DataGridView1.Rows(l).Cells(k).Value = y(l)
                    Else
                        DataGridView1.Rows(l + l).Cells(k).Value = y(l)
                    End If
                    l = l + 1
                End While
            End If
            k = k + 1
        End While
        m = n - (i + 1)
        s = (vaInterpolar - x(i)) / (x(i + 1) - x(i))
        Label1.Text = "Se observa que i = " + i.ToString() + ", i+1 = " + (i + 1).ToString()
        Label2.Text = "Por lo tanto: Xi = " + x(i).ToString() + ", Xi+1 = " + x(i + 1).ToString()
        Label3.Text = "m = " + m.ToString()
        Label4.Text = "s = " + Math.Round(s, 4).ToString()
        Dim Csk(m) As Single
        Csk(0) = 1
        Csk(1) = s
        For k = 2 To m
            Csk(k) = 1
            For h = 1 To k
                Csk(k) = Csk(k) * (((s) - (h - 1)) / h)
            Next
        Next


        Dim coef, sumad, delta(m) As Single
        delta(0) = y(i)
        For d = 1 To m
            sumad = 0
            For k = 0 To d
                coef = 1
                For p = 1 To k
                    coef = coef * (d - k + p) / p
                Next
                sumad = sumad + coef * -1 ^ k * y(i + d - k)
            Next
            delta(d) = sumad
        Next

        For d = 0 To m
            DataGridView1.Rows.Add(delta(d))
        Next

        For k = 0 To m
            DataGridView1.Rows.Add(Math.Round(s, 4).ToString() + "/" + k.ToString, Math.Round(Csk(k), 4))
        Next
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs)


    End Sub
End Class
