Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim km, pago, viaje, redon As Integer
        Dim total, total1, total2, tot, totre, puntos, pa, pare As Single

        km = Val(TextBox1.Text)
        redon = km * 2 'KILOMETROS REDONDOS
        tot = km * 1.48 'km x precio
        totre = redon * 1.48 'km redondos x precio
        pago = Val(TextBox2.Text)
        viaje = Val(TextBox3.Text)

        If km > 0 Then ' KM MAYORES A 0

            If pago = 1 Then ' PAGO EFECTIVO

                If viaje = 1 Then 'VIAJE SENCILLO


                    If km > 1200 Then
                        total = tot - (tot * 0.15)
                        puntos = km * 0.3
                    ElseIf km > 800 Then
                        total = tot - (tot * 0.09)
                        puntos = km * 0.18
                    ElseIf km > 600 Then
                        total = tot - (tot * 0.05)
                        puntos = km * 0.15
                    ElseIf km > 400 Then
                        total = tot - (tot * 0.04)
                        puntos = km * 0.1
                    ElseIf km <= 400 Then
                        total = tot - (tot * 0.03)
                        puntos = km * 0.08
                    End If



                ElseIf viaje = 2 Then 'VIAJE REDONDO



                    If redon > 1200 Then
                        total = totre - (totre * 0.2) '20% 
                        puntos = redon * 0.3
                    ElseIf redon > 800 Then
                        total = (totre - (totre * 0.1)) '10% 
                        puntos = redon * 0.25
                    ElseIf redon > 600 Then
                        total1 = totre - (totre * 0.05)
                        total = total1 - (total1 * 0.07) '7% adicional
                        puntos = redon * 0.2
                    ElseIf redon > 400 Then
                        total1 = totre - (totre * 0.04)
                        total = (total1 - total1 * 0.06) '6% adicional
                        puntos = redon * 0.15
                    ElseIf redon <= 400 Then
                        total1 = totre - (totre * 0.03)
                        total = total1 - (total * 0.04) '4% adicional
                        puntos = redon * 0.08
                    End If



                Else
                    MsgBox("VIAJE NO VALIDO, ELIGE 1 ó 2", MsgBoxStyle.Critical)

                End If ' VIN DE VIAJE



            ElseIf pago = 2 Then ' PAGO TARJETA  ------ 2% ADICIONAL

                pa = km * 0.02                  '  puntos adcionales 
                pare = redon * 0.02             '  puntos adcionales en km redondos


                If viaje = 1 Then 'VIAJE SENCILLO


                    If km > 1200 Then
                        total1 = tot - (tot * 0.15)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.3) + pa
                    ElseIf km > 800 Then
                        total1 = tot - (tot * 0.09)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.18) + pa
                    ElseIf km > 600 Then
                        total1 = tot - (tot * 0.05)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.15) + pa
                    ElseIf km > 400 Then
                        total1 = tot - (tot * 0.04)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.1) + pa
                    ElseIf km <= 400 Then
                        total1 = tot - (tot * 0.03)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.08) + pa
                    End If



                ElseIf viaje = 2 Then 'VIAJE REDONDO



                    If redon > 1200 Then
                        total1 = totre - (totre * 0.2) '20%
                        total = total1 - (total * 0.02)
                        puntos = (redon * 0.3) + pare
                    ElseIf redon > 800 Then
                        total1 = (totre - (totre * 0.1)) '10%
                        total = total1 - (total * 0.02)
                        puntos = (redon * 0.25) + pare
                    ElseIf redon > 600 Then
                        total1 = totre - (totre * 0.05)
                        total2 = total1 - (total1 * 0.07) '7% adicional
                        total = total2 - (total2 * 0.02)
                        puntos = (redon * 0.2) + pare
                    ElseIf redon > 400 Then
                        total1 = totre - (totre * 0.04)
                        total2 = (total1 - total1 * 0.06) '6% adicional
                        total = total2 - (total2 * 0.02)
                        puntos = (redon * 0.15) + pare
                    ElseIf redon <= 400 Then
                        total1 = totre - (totre * 0.03)
                        total2 = total1 - (total * 0.04) '4% adicional
                        total = total2 - (total2 * 0.02)
                        puntos = (redon * 0.08) + pare
                    End If



                Else
                    MsgBox("Viaje no valido, Elige 1 ó 2", MsgBoxStyle.Critical, " Advertencia")

                End If ' VIN DE VIAJE




            Else
                MsgBox("Pago no Valido, Elige 1 ó 2 ", MsgBoxStyle.Critical, " Advertencia")


            End If ' FIN DE PAGO


        Else 'ELSE KM
            MsgBox("Ingresa KM reales ", MsgBoxStyle.Critical, " Advertencia")
        End If ' FIN DE KM
        Label4.Text = "TOTAL A PAGAR:  " & total & "           PUNTOS PROXIMO VIAJE:  " & puntos
    End Sub
    
End Class
