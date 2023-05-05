Public Class Form1
    Dim negativo, positivo, valor, tolerancia, total, h, tolerancia1 As Decimal
    Public Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.SelectedItem

        Select Case ComboBox1.SelectedIndex
            Case 0
                Button1.BackColor = Color.Brown
                TextBox5.Text = "1"
            Case 1
                Button1.BackColor = Color.Red
                TextBox5.Text = "2"
            Case 2
                Button1.BackColor = Color.Orange
                TextBox5.Text = "3"
            Case 3
                Button1.BackColor = Color.Yellow
                TextBox5.Text = "4"
            Case 4
                Button1.BackColor = Color.Green
                TextBox5.Text = "5"
            Case 5
                Button1.BackColor = Color.Blue
                TextBox5.Text = "6"
            Case 6
                Button1.BackColor = Color.Purple
                TextBox5.Text = "7"
            Case 7
                Button1.BackColor = Color.Gray
                TextBox5.Text = "8"
            Case 8
                Button1.BackColor = Color.White
                TextBox5.Text = "9"
        End Select

    End Sub



    Public Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox2.Text = ComboBox2.SelectedItem
        Select Case ComboBox2.SelectedIndex
            Case 0
                Button2.BackColor = Color.Black
                TextBox5.Text += "0"
            Case 1
                Button2.BackColor = Color.Brown
                TextBox5.Text += "1"
            Case 2
                Button2.BackColor = Color.Red
                TextBox5.Text += "2"
            Case 3
                Button2.BackColor = Color.Orange
                TextBox5.Text += "3"
            Case 4
                Button2.BackColor = Color.Yellow
                TextBox5.Text += "4"
            Case 5
                Button2.BackColor = Color.Green
                TextBox5.Text += "5"

            Case 6
                TextBox5.Text += "6"
                Button2.BackColor = Color.Blue
            Case 7
                Button2.BackColor = Color.Purple
                TextBox5.Text += "7"
            Case 8
                Button2.BackColor = Color.Gray
                TextBox5.Text += "8"
            Case 9
                Button2.BackColor = Color.White
                TextBox5.Text += "9"
        End Select

    End Sub

    Public Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox3.Text = ComboBox3.SelectedItem
        Select Case ComboBox3.SelectedIndex

            Case 0
                Button3.BackColor = Color.Black
                TextBox5.Text += ""
                Label7.Text = "Ω  Error Calculado"
            Case 1
                Button3.BackColor = Color.Brown
                TextBox5.Text += "0"
                Label7.Text = "Ω  Error Calculado"

            Case 2
                Button3.BackColor = Color.Red
                TextBox5.Text += "00"
                Label7.Text = "Ω  Error Calculado"
            Case 3
                Button3.BackColor = Color.Orange
                TextBox6.Text += "000"
                Label7.Text = "Ω  Error Calculado"
            Case 4
                Button3.BackColor = Color.Yellow
                TextBox5.Text += "0000"
                Label7.Text = "Ω  Error Calculado"
            Case 5
                Button3.BackColor = Color.Green
                TextBox5.Text += "00000"
                Label7.Text = "Ω  Error Calculado"
            Case 6
                Button3.BackColor = Color.Blue
                TextBox5.Text += "0000000"
                Label7.Text = "Ω  Error Calculado"
            Case 7
                Button3.BackColor = Color.Purple
                TextBox5.Text += "0000000"
                Label7.Text = "Ω   Error Calculado"
            Case 8
                Button3.BackColor = Color.Gray
                TextBox5.Text += "00000000"
                Label7.Text = "Ω   Error Calculado"
            Case 9
                Button3.BackColor = Color.White
                TextBox5.Text += "000000000"
                Label7.Text = "Ω   Error Calculado"
            Case 10
                Button3.BackColor = Color.Gold
                TextBox5.Text = TextBox6.Text * 0.1
                Label7.Text = "Ω   Error Calculado"
            Case 11
                Button3.BackColor = Color.Silver
                TextBox5.Text = TextBox6.Text * 0.01
                Label7.Text = "Ω   Error Calculado"
        End Select
    End Sub

    Public Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        TextBox4.Text = ComboBox4.SelectedItem
        Select Case ComboBox4.SelectedIndex

            Case 0
                Button4.BackColor = Color.Brown
                TextBox7.Text = TextBox5.Text * 1 / 100
                Label5.Text = "Ω + -1%"
                Label12.Text = " + -1%"

            Case 1
                Button4.BackColor = Color.Red
                TextBox7.Text = TextBox5.Text / 100 * 2
                Label5.Text = "Ω + -2%"
                Label12.Text = " + -2%"
            Case 2
                Button4.BackColor = Color.Green
                TextBox7.Text = TextBox5.Text / 100 * 0.5
                Label5.Text = "Ω + -0.5%"
                Label12.Text = " + -0.5%"
            Case 3
                Button4.BackColor = Color.Blue
                TextBox7.Text = TextBox5.Text / 100 * 0.25
                Label12.Text = "Ω + -0.25%"

                Label6.Text += " + -0.25%"
            Case 4
                Button4.BackColor = Color.Purple
                TextBox7.Text = TextBox5.Text / 100 * 0.1
                Label5.Text = "Ω + -0.10%"
                Label12.Text = " + -0.10%"
            Case 5
                Button4.BackColor = Color.Gray
                TextBox7.Text = TextBox5.Text / 100 * 0.5
                Label5.Text = "Ω + -0.5%"
                Label12.Text = " + -0.5%"
            Case 6
                Button4.BackColor = Color.Gold
                TextBox7.Text = TextBox5.Text / 100 * 5
                Label5.Text = "Ω + -5%"
                Label12.Text = " + -5%"
            Case 7
                Button4.BackColor = Color.Silver
                TextBox7.Text = TextBox5.Text / 100 * 10
                Label5.Text = "Ω + -10%"
                Label12.Text = " + -10%"
        End Select


        valor = TextBox5.Text
        If TextBox5.Text < 1000 Then
            TextBox6.Text = TextBox5.Text

        Else
            If TextBox5.Text > 1000000000 Then
                total = valor / 1000000000


                Label6.Text = "G Ω"

            Else
                If TextBox5.Text > 100000000 Then
                    total = valor / 1000000

                    Label6.Text = "M Ω"


                Else
                    If TextBox5.Text > 10000000 Then
                        total = valor / 1000000


                        Label6.Text = "M Ω"

                    Else
                        If TextBox5.Text > 1000000 Then
                            total = valor / 1000000


                            Label6.Text = "M Ω"


                        Else
                            If TextBox5.Text > 100000 Then
                                total = valor / 1000

                                Label6.Text = "KΩ"

                            Else
                                If TextBox5.Text > 10000 Then
                                    total = valor / 1000
                                    Label6.Text = "KΩ"

                                Else
                                    If TextBox5.Text >= 1000 Then
                                        total = valor / 1000
                                        Label6.Text = "KΩ"



                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        TextBox8.Text = total

        negativo = valor - TextBox7.Text
        If negativo < 1000 Then
            TextBox6.Text = negativo
            tolerancia = negativo
            Label9.Text = " Ω  Y"

        Else
            If negativo > 1000000000 Then
                tolerancia = negativo / 1000000000


                Label9.Text = "G Ω  Y"

            Else
                If negativo > 100000000 Then
                    tolerancia = negativo / 1000000

                    Label9.Text = "M Ω  Y"


                Else
                    If negativo > 10000000 Then
                        tolerancia = negativo / 1000000


                        Label9.Text = "M Ω  Y"

                    Else
                        If negativo > 1000000 Then
                            tolerancia = negativo / 1000000


                            Label9.Text = "M Ω  Y"


                        Else
                            If negativo > 100000 Then
                                tolerancia = negativo / 1000

                                Label9.Text = "KΩ  Y"

                            Else
                                If negativo > 10000 Then
                                    tolerancia = negativo / 1000
                                    Label9.Text = "KΩ  Y"

                                Else
                                    If negativo >= 1000 Then
                                        tolerancia = negativo / 1000
                                        Label9.Text = "KΩ  Y"



                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        TextBox6.Text = tolerancia



        positivo = valor + TextBox7.Text
        If positivo < 1000 Then
            TextBox6.Text = positivo
            tolerancia1 = positivo
            Label10.Text = "Ω "

        Else
            If positivo > 1000000000 Then
                tolerancia1 = positivo / 1000000000


                Label10.Text = "G Ω  "

            Else
                If positivo > 100000000 Then
                    tolerancia1 = positivo / 1000000

                    Label10.Text = "M Ω  "


                Else
                    If positivo > 10000000 Then
                        tolerancia1 = positivo / 1000000


                        Label10.Text = "M Ω  "

                    Else
                        If positivo > 1000000 Then
                            tolerancia1 = positivo / 1000000


                            Label10.Text = "M Ω"


                        Else
                            If positivo > 100000 Then
                                tolerancia1 = positivo / 1000

                                Label10.Text = "KΩ "

                            Else
                                If positivo > 10000 Then
                                    tolerancia1 = positivo / 1000
                                    Label10.Text = "KΩ "

                                Else
                                    If positivo >= 1000 Then
                                        tolerancia1 = positivo / 1000
                                        Label10.Text = "KΩ "



                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        TextBox9.Text = tolerancia1


    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        GroupBox1.Visible = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""

        TextBox8.Text = ""
        TextBox9.Text = ""
        Label5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
        Label8.Text = ""
        Label9.Text = ""
        Label10.Text = ""
        Label11.Text = ""
        Label12.Text = ""
        Button3.BackColor = Color.WhiteSmoke
        Button4.BackColor = Color.WhiteSmoke
        Button1.BackColor = Color.WhiteSmoke
        Button2.BackColor = Color.WhiteSmoke


    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        GroupBox1.Visible = True
        TextBox6.Visible = True
        TextBox7.Visible = True
        TextBox5.Visible = True
        TextBox8.Visible = True
        TextBox9.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label12.Visible = True

    End Sub
End Class