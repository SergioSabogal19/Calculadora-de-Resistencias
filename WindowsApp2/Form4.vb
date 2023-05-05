Public Class Form4
    Public resistencia, b1, b2, b3, b4, b5, valor, suma, mo, gr, bl As Integer
    Dim calculo, tt2, tolerancia As Decimal
    Dim residuo, t1, t3, t4, t5, t6, t7, t8, t9, tt1, t, tt3, tt4, tt5, tt6, tt7, tt8, tt9 As Decimal
    Dim ss As Char

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub



    Public Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Text = 0
        TextBox9.Text = 0
        TextBox11.Text = 0
        TextBox9.Visible = 0
        Button3.BackColor = Color.WhiteSmoke
        Button4.BackColor = Color.WhiteSmoke
        Button5.BackColor = Color.WhiteSmoke
        Button6.BackColor = Color.WhiteSmoke
        Button7.BackColor = Color.WhiteSmoke
        TextBox6.Visible = False
        TextBox7.Visible = False
        TextBox8.Visible = False
        GroupBox1.Visible = False
        TextBox9.Visible = False
        TextBox11.Visible = False
        Label5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Public Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.SelectedItem

        Select Case ComboBox1.SelectedIndex
            Case 0
                Button3.BackColor = Color.Brown
                TextBox6.Text = "1"
            Case 1
                Button3.BackColor = Color.Red
                TextBox6.Text = "2"
            Case 2
                Button3.BackColor = Color.Orange
                TextBox6.Text = "3"
            Case 3
                Button3.BackColor = Color.Yellow
                TextBox6.Text = "4"
            Case 4
                Button3.BackColor = Color.Green
                TextBox6.Text = "5"
            Case 5
                Button3.BackColor = Color.Blue
                TextBox6.Text = "6"
            Case 6
                Button3.BackColor = Color.Purple
                TextBox6.Text = "7"
            Case 7
                Button3.BackColor = Color.Gray
                TextBox6.Text = "8"
            Case 8
                Button3.BackColor = Color.White
                TextBox6.Text = "9"
        End Select
    End Sub

    Public Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox2.Text = ComboBox2.SelectedItem
        Select Case ComboBox2.SelectedIndex
            Case 0
                Button4.BackColor = Color.Black
                TextBox6.Text += "0"
            Case 1
                Button4.BackColor = Color.Brown
                TextBox6.Text += "1"
            Case 2
                Button4.BackColor = Color.Red
                TextBox6.Text += "2"
            Case 3
                Button4.BackColor = Color.Orange
                TextBox6.Text += "3"
            Case 4
                Button4.BackColor = Color.Yellow
                TextBox6.Text += "4"
            Case 5
                Button4.BackColor = Color.Green
                TextBox6.Text += "5"

            Case 6
                TextBox6.Text += "6"
                Button4.BackColor = Color.Blue
            Case 7
                Button4.BackColor = Color.Purple
                TextBox6.Text += "7"
            Case 8
                Button4.BackColor = Color.Gray
                TextBox6.Text += "8"
            Case 9
                Button4.BackColor = Color.White
                TextBox6.Text += "9"
        End Select

    End Sub

    Public Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox3.Text = ComboBox3.SelectedItem
        Select Case ComboBox3.SelectedIndex
            Case 0
                Button5.BackColor = Color.Black
                TextBox6.Text += "0"
            Case 1
                Button5.BackColor = Color.Brown
                TextBox6.Text += "1"
            Case 2
                Button5.BackColor = Color.Red
                TextBox6.Text += "2"
            Case 3
                Button5.BackColor = Color.Orange
                TextBox6.Text += "3"
            Case 4
                Button5.BackColor = Color.Yellow
                TextBox6.Text += "4"
            Case 5
                Button5.BackColor = Color.Green
                TextBox6.Text += "5"
            Case 6
                Button5.BackColor = Color.Blue
                TextBox6.Text += "6"
            Case 7
                Button5.BackColor = Color.Purple
                TextBox6.Text += "7"
            Case 8
                Button5.BackColor = Color.Gray
                TextBox6.Text += "8"
            Case 9
                Button5.BackColor = Color.White
                TextBox6.Text += "9"
        End Select

    End Sub

    Public Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        TextBox4.Text = ComboBox4.SelectedItem
        Select Case ComboBox4.SelectedIndex
            Case 0
                Button6.BackColor = Color.Black
                TextBox6.Text += ""
                Label7.Text = "Error Calculado"
            Case 1
                Button6.BackColor = Color.Brown
                TextBox6.Text += "0"
                Label7.Text = "Error Calculado"

            Case 2
                Button6.BackColor = Color.Red
                TextBox6.Text += "00"
                Label7.Text = "Error Calculado"
            Case 3
                Button6.BackColor = Color.Orange
                TextBox6.Text += "000"
                Label7.Text = "Error Calculado"
            Case 4
                Button6.BackColor = Color.Yellow
                TextBox6.Text += "0000"
                Label7.Text = "Error Calculado"
            Case 5
                Button6.BackColor = Color.Green
                TextBox6.Text += "00000"
                Label7.Text = "Error Calculado"
            Case 6
                Button6.BackColor = Color.Blue
                TextBox6.Text += "0000000"
                Label7.Text = "Error Calculado"
            Case 7
                Button6.BackColor = Color.Purple
                TextBox6.Text += "0000000"
                Label7.Text = "Error Calculado"
            Case 8
                Button6.BackColor = Color.Gray
                TextBox6.Text += "00000000"
                Label7.Text = "Error Calculado"
            Case 9
                Button6.BackColor = Color.White
                TextBox6.Text += "000000000"
                Label7.Text = "Error Calculado"
            Case 10
                Button6.BackColor = Color.Gold
                TextBox6.Text = TextBox6.Text * 0.1
                Label7.Text = "Error Calculado"
            Case 11
                Button6.BackColor = Color.Silver
                TextBox6.Text = TextBox6.Text * 0.01
                Label7.Text = "Error Calculado"
        End Select

    End Sub
    Public Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

        TextBox9.Text = TextBox6.Text


        If TextBox6.Text < 1000 Then
            TextBox9.Text = TextBox6.Text
        Else
            If TextBox6.Text > 1000000000 Then
                calculo = TextBox6.Text / 1000000000
                Label6.Text = "G"
                TextBox9.Text = calculo
            Else
                If TextBox6.Text > 100000000 Then
                    calculo = TextBox6.Text / 1000000
                    Label6.Text = "M"
                    TextBox9.Text = calculo
                Else
                    If TextBox6.Text > 10000000 Then
                        calculo = TextBox6.Text / 1000000
                        Label6.Text = "M"
                        TextBox9.Text = calculo
                    Else
                        If TextBox6.Text > 1000000 Then
                            calculo = TextBox6.Text / 1000000
                            Label6.Text = "M"
                            TextBox9.Text = calculo
                        Else
                            If TextBox6.Text > 100000 Then
                                calculo = TextBox6.Text / 1000
                                Label6.Text = "K"
                                TextBox9.Text = calculo
                            Else
                                If TextBox6.Text > 10000 Then
                                    calculo = TextBox6.Text / 1000
                                    Label6.Text = "K"
                                    TextBox9.Text = calculo
                                Else
                                    If TextBox6.Text >= 1000 Then
                                        calculo = TextBox6.Text / 1000
                                        Label6.Text = "K"
                                        TextBox9.Text = calculo

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        TextBox9.Text = calculo


    End Sub

    Public Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        TextBox5.Text = ComboBox5.SelectedItem

        Select Case ComboBox5.SelectedIndex

            Case 0
                Button7.BackColor = Color.Brown
                TextBox8.Text = TextBox6.Text * 1 / 100
                t1 = TextBox6.Text - TextBox8.Text
                tt1 = TextBox6.Text + TextBox8.Text
                TextBox7.Text = t1
                TextBox7.Text += "-"
                TextBox7.Text += tt1
                Label5.Text += "Ω + -1%"
                Label6.Text += "Ω + -1%"




            Case 1
                Button7.BackColor = Color.Red
                TextBox8.Text = TextBox6.Text / 100 * 2
                t1 = TextBox6.Text - TextBox8.Text
                tt1 = TextBox6.Text + TextBox8.Text
                TextBox7.Text = t1
                TextBox7.Text += "-"
                TextBox7.Text += tt1
                Label5.Text += "Ω + -2%"
                Label6.Text += "Ω + -2%"

            Case 2
                Button7.BackColor = Color.Green
                TextBox8.Text = TextBox6.Text / 100 * 0.5
                t1 = TextBox6.Text - TextBox8.Text
                tt1 = TextBox6.Text + TextBox8.Text
                TextBox7.Text = t1
                TextBox7.Text += "-"
                TextBox7.Text += tt1
                Label5.Text += "Ω + -0.5%"
                Label6.Text += "Ω + -0.5%"


            Case 3
                Button7.BackColor = Color.Blue
                TextBox8.Text = TextBox6.Text / 100 * 0.25
                t1 = TextBox6.Text - TextBox8.Text
                tt1 = TextBox6.Text + TextBox8.Text
                TextBox7.Text = t1
                TextBox7.Text += "-"
                TextBox7.Text += tt1
                Label5.Text += "Ω + -0.25%"
                Label6.Text += "Ω + -0.25%"


            Case 4
                Button7.BackColor = Color.Purple
                TextBox8.Text = TextBox6.Text / 100 * 0.1
                t1 = TextBox6.Text - TextBox8.Text
                tt1 = TextBox6.Text + TextBox8.Text
                'TextBox7.Text = t1
                'TextBox7.Text += "-"
                'TextBox7.Text += tt1
                Label5.Text += "Ω + -0.10%"
                Label6.Text += "Ω + -0.10%"

            Case 5
                Button7.BackColor = Color.Gray
                TextBox8.Text = TextBox6.Text / 100 * 0.5
                't1 = TextBox6.Text - TextBox8.Text
                'tt1 = TextBox6.Text + TextBox8.Text
                'TextBox7.Text = t1
                'TextBox7.Text += "-"
                'TextBox7.Text += tt1
                Label5.Text += "Ω + -0.5%"
                Label6.Text += "Ω + -0.5%"


            Case 6
                Button7.BackColor = Color.Gold
                TextBox8.Text = TextBox6.Text / 100 * 5
                't1 = TextBox6.Text - TextBox8.Text
                'tt1 = TextBox6.Text + TextBox8.Text
                'TextBox7.Text = t1
                'TextBox7.Text += "-"
                'TextBox7.Text += tt1
                Label5.Text += "Ω + -5%"
                Label6.Text += "Ω + -5%"


            Case 7
                Button7.BackColor = Color.Silver
                TextBox8.Text = TextBox6.Text / 100 * 10
                TextBox7.Text = TextBox6.Text - TextBox8.Text



                ' tt1 = TextBox6.Text + TextBox8.Text
                'TextBox7.Text = t1
                ' TextBox7.Text += "-"
                'TextBox7.Text += tt1
                Label5.Text += "Ω + -10%"
                Label6.Text += "Ω + -10%"


        End Select
        TextBox7.Text = TextBox6.Text - TextBox8.Text
        TextBox10.Text = ((TextBox8.Text) + (TextBox6.Text))
        If TextBox7.Text < 1000 Then
            TextBox7.Text = TextBox7.Text
            TextBox11.Text = TextBox10.Text
        Else
            If TextBox7.Text > 1000000000 Then
                tt2 = TextBox7.Text / 1000000000
                tt3 = TextBox10.Text / 1000000000
                TextBox11.Text = tt3
                Label8.Text = "G Ω  Y"
                Label9.Text = "G Ω"

            Else
                If TextBox7.Text > 100000000 Then
                    tt2 = TextBox7.Text / 1000000
                    tt3 = TextBox10.Text / 1000000
                    Label8.Text = "MΩ  Y"
                    Label9.Text = "M Ω"
                    TextBox11.Text = tt3

                Else
                    If TextBox7.Text > 10000000 Then
                        tt2 = TextBox7.Text / 1000000
                        tt3 = TextBox10.Text / 1000000
                        Label8.Text = "MΩ  Y"
                        Label9.Text = "M Ω"
                        TextBox11.Text = tt3

                    Else
                        If TextBox7.Text > 1000000 Then
                            tt2 = TextBox7.Text / 1000000
                            tt3 = TextBox10.Text / 1000000
                            Label8.Text = "MΩ  Y"
                            Label9.Text = "M Ω"
                            TextBox11.Text = tt3

                        Else
                            If TextBox7.Text > 100000 Then
                                tt2 = TextBox7.Text / 1000
                                tt3 = TextBox10.Text / 1000
                                Label8.Text = "KΩ  Y"
                                Label9.Text = "KΩ"
                                TextBox11.Text = tt3

                            Else
                                If TextBox7.Text > 10000 Then
                                    tt2 = TextBox7.Text / 1000
                                    tt3 = TextBox10.Text / 1000
                                    Label8.Text = "KΩ  Y"
                                    Label9.Text = "KΩ"
                                    TextBox11.Text = tt3

                                Else
                                    If TextBox7.Text >= 1000 Then
                                        tt2 = TextBox7.Text / 1000
                                        tt3 = TextBox10.Text / 1000
                                        Label8.Text = "KΩ  Y"
                                        Label9.Text = "KΩ"
                                        TextBox11.Text = tt3


                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        TextBox7.Text = tt2
        TextBox11.Text = tt3

        'positivo del error 

        tt7 = TextBox6.Text
        tt8 = TextBox8.Text
        tt9 = tt7 + tt8
        TextBox10.Text = tt9
        If TextBox10.Text < 1000 Then
            TextBox10.Text = tt9
            TextBox11.Text = tt9
        Else
            If tt9 > 1000000000 Then

                tt5 = tt9 / 1000000000


            Else
                If tt9 > 100000000 Then

                    tt5 = tt9 / 1000000



                Else
                    If tt9 > 10000000 Then

                        tt5 = TextBox10.Text / 1000000



                    Else
                        If tt9 > 1000000 Then

                            tt5 = TextBox10.Text / 1000000



                        Else
                            If tt9 > 100000 Then

                                tt5 = TextBox10.Text / 1000


                            Else
                                If tt9 > 10000 Then

                                    tt5 = TextBox10.Text / 1000


                                Else
                                    If tt9 >= 1000 Then

                                        tt5 = TextBox10.Text / 1000

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        TextBox11.Text = tt5

    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GroupBox1.Visible = True

        TextBox6.Visible = True
        TextBox7.Visible = True
        TextBox8.Visible = True

        TextBox9.Visible = True
        Label5.Visible = True
        Label7.Visible = True
        Label6.Visible = True

        TextBox11.Visible = True
    End Sub
End Class
