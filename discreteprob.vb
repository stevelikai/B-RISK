Public Class frmDiscreteProb
    Dim spr1prob As Decimal 'probability 1 sprinkler required to suppress/control fire
    Dim spr2prob As Decimal 'probability 2 sprinklers required to suppress/control fire
    Dim spr3prob As Decimal 'probability 3 sprinklers required to suppress/control fire
    Dim spr4prob As Decimal 'probability 4 sprinklers required to suppress/control fire
    Dim total As Decimal

    Private Sub txtSpr1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpr1.Leave
        If Not IsNumeric(txtSpr1.Text) Then
            txtSpr1.Text = 0D
        End If
    End Sub

    Private Sub txtSpr1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSpr1.TextChanged

        If IsNumeric(txtSpr1.Text) Then spr1prob = CDec(txtSpr1.Text)
        If IsNumeric(txtSpr2.Text) Then spr2prob = CDec(txtSpr2.Text)
        If IsNumeric(txtSpr3.Text) Then spr3prob = CDec(txtSpr3.Text)
        If IsNumeric(txtSpr4.Text) Then spr4prob = CDec(txtSpr4.Text)

        total = spr1prob + spr2prob + spr3prob + spr4prob
        TextBox1.Text = total

    End Sub
    Private Function IsValidData() As Boolean
        Return Validator.IsPresent(txtSpr1) AndAlso _
               Validator.IsPresent(txtSpr2) AndAlso _
               Validator.IsPresent(txtSpr3) AndAlso _
                Validator.IsPresent(txtSpr4) AndAlso _
                Validator.IsWithinRange(txtSpr1, 0, 1) AndAlso _
                 Validator.IsWithinRange(txtSpr2, 0, 1) AndAlso _
         Validator.IsWithinRange(txtSpr3, 0, 1) AndAlso _
         Validator.IsWithinRange(txtSpr4, 0, 1)
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsValidData() Then

            If total <> 1D Then
                MsgBox("The probabilities must total 1.000")
                txtSpr4.Focus()
            Else

                sprnum_prob(0) = spr1prob
                sprnum_prob(1) = spr2prob
                sprnum_prob(2) = spr3prob
                sprnum_prob(3) = spr4prob 'probability 4 sprinklers required to suppress/control fire

                Me.Close()
            End If
        End If
    End Sub

    Private Sub txtSpr2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpr2.Leave
        If Not IsNumeric(txtSpr2.Text) Then
            txtSpr2.Text = 0D
        End If
    End Sub

    Private Sub txtSpr2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSpr2.TextChanged

        If IsNumeric(txtSpr1.Text) Then spr1prob = CDec(txtSpr1.Text)
        If IsNumeric(txtSpr2.Text) Then spr2prob = CDec(txtSpr2.Text)
        If IsNumeric(txtSpr3.Text) Then spr3prob = CDec(txtSpr3.Text)
        If IsNumeric(txtSpr4.Text) Then spr4prob = CDec(txtSpr4.Text)

        total = spr1prob + spr2prob + spr3prob + spr4prob
        TextBox1.Text = total

    End Sub

    Private Sub txtSpr3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpr3.Leave
        If Not IsNumeric(txtSpr3.Text) Then
            txtSpr3.Text = 0D
        End If
    End Sub

    Private Sub txtSpr3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSpr3.TextChanged


        If IsNumeric(txtSpr1.Text) Then spr1prob = CDec(txtSpr1.Text)
        If IsNumeric(txtSpr2.Text) Then spr2prob = CDec(txtSpr2.Text)
        If IsNumeric(txtSpr3.Text) Then spr3prob = CDec(txtSpr3.Text)
        If IsNumeric(txtSpr4.Text) Then spr4prob = CDec(txtSpr4.Text)

        total = spr1prob + spr2prob + spr3prob + spr4prob
        TextBox1.Text = total

    End Sub

    Private Sub txtSpr4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpr4.Leave
        If Not IsNumeric(txtSpr4.Text) Then
            txtSpr4.Text = 0D
        End If
    End Sub

    Private Sub txtSpr4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSpr4.TextChanged


        If IsNumeric(txtSpr1.Text) Then spr1prob = CDec(txtSpr1.Text)
        If IsNumeric(txtSpr2.Text) Then spr2prob = CDec(txtSpr2.Text)
        If IsNumeric(txtSpr3.Text) Then spr3prob = CDec(txtSpr3.Text)
        If IsNumeric(txtSpr4.Text) Then spr4prob = CDec(txtSpr4.Text)

        total = spr1prob + spr2prob + spr3prob + spr4prob
        TextBox1.Text = total

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Call Me.Close()
    End Sub
End Class