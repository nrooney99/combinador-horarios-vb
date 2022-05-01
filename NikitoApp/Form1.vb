Public Class Form1
    Dim fileReader, materia(6), lineas(), h1() As String
    Dim numTit(6), inicioLin(0), i, j, ac, ac2(5), ac3(5), numCom1(1), numCom2(1), numCom3(1), numCom4(1), numCom5(1) As Integer
    
    
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'fileReader = My.Computer.FileSystem.ReadAllText("E:/test.txt")
        ComboBox1.Items.Clear()
        fileReader = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "/test.txt")
        ListBox1.Items.Clear()
        lineas = Split(fileReader, vbNewLine)
        ac = 1
        ac2(1) = 1
        ac2(2) = 1
        ac2(3) = 1
        ac2(4) = 1
        ac2(5) = 1
        inicioLin(0) = 1
        For Me.i = LBound(lineas) To UBound(lineas) - 1
            ReDim Preserve inicioLin(i + 1)
            If inicioLin(i) < 1 Then
                inicioLin(i + 1) = InStr(1, fileReader, lineas(i + 1), CompareMethod.Text)
            Else
                inicioLin(i + 1) = InStr(inicioLin(i), fileReader, lineas(i + 1), CompareMethod.Text)
            End If
            'Busco numero de linea de cada materia y almaceno en numTit

            If Mid(lineas(i), 1, 5) = "Comis" Then
                numTit(ac) = i - 1
                ac = ac + 1
            End If

            'ListBox1.Items.Add(lineas(i))
            'ListBox1.Items.Add(inicioLin(i))
        Next
        For Me.i = LBound(lineas) To UBound(lineas) - 1
            If Len(lineas(i)) < 3 Then
                Select Case i
                    Case Is < numTit(2)
                        numCom1(ac2(1)) = i
                        ac2(1) = ac2(1) + 1
                        ReDim Preserve numCom1(UBound(numCom1) + 1)
                    Case Is < numTit(3)
                        numCom2(ac2(2)) = i
                        ac2(2) = ac2(2) + 1
                        ReDim Preserve numCom2(UBound(numCom2) + 1)
                    Case Is < numTit(4)
                        numCom3(ac2(3)) = i
                        ac2(3) = ac2(3) + 1
                        ReDim Preserve numCom3(UBound(numCom3) + 1)
                    Case Is < numTit(5)
                        numCom4(ac2(4)) = i
                        ac2(4) = ac2(4) + 1
                        ReDim Preserve numCom4(UBound(numCom4) + 1)
                    Case Is > numTit(5)
                        numCom5(ac2(5)) = i
                        ac2(5) = ac2(5) + 1
                        ReDim Preserve numCom5(UBound(numCom5) + 1)
                End Select
            End If
        Next
        For Me.i = 1 To 5
            'Asigno al texto segun linea
            materia(i) = lineas(numTit(i))
            ComboBox1.Items.Add(materia(i))
        Next

        ac3(1) = 1
        ac3(2) = 1
        ac3(3) = 1
        ac3(4) = 1
        ac3(5) = 1
        ReDim Preserve h1(UBound(numCom1))
        For Me.i = LBound(lineas) To UBound(lineas) - 1
            Select Case i
                Case Is < numTit(2)
                    Select Case Mid(lineas(i), 1, InStr(lineas(i), " ", CompareMethod.Text))
                        Case "Lunes"
                            h1(ac3(1)) = h1(ac3(1)) & "L"
                            ac3(1) = ac3(1) + 1
                        Case "Martes"
                            h1(ac3(1)) = h1(ac3(1)) & "M"
                            ac3(1) = ac3(1) + 1
                        Case "Miércoles"
                            h1(ac3(1)) = h1(ac3(1)) & "X"
                            ac3(1) = ac3(1) + 1
                        Case "Jueves"
                            h1(ac3(1)) = h1(ac3(1)) & "J"
                            ac3(1) = ac3(1) + 1
                        Case "Viernes"
                            h1(ac3(1)) = h1(ac3(1)) & "V"
                            ac3(1) = ac3(1) + 1
                    End Select
                Case Is < numTit(3)

                Case Is < numTit(4)
                    ReDim Preserve numCom3(UBound(numCom3) + 1)
                Case Is < numTit(5)
                    ReDim Preserve numCom4(UBound(numCom4) + 1)
                Case Is > numTit(5)
                    ReDim Preserve numCom5(UBound(numCom5) + 1)
            End Select
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
        Select Case ComboBox1.Text
            Case materia(1)
                For Me.i = 1 To UBound(numCom1) - 1
                    ListBox1.Items.Add(lineas(numCom1(i)))
                Next
            Case materia(2)
                For Me.i = 1 To UBound(numCom2) - 1
                    ListBox1.Items.Add(lineas(numCom2(i)))
                Next
            Case materia(3)
                For Me.i = 1 To UBound(numCom3) - 1
                    ListBox1.Items.Add(lineas(numCom3(i)))
                Next
            Case materia(4)
                For Me.i = 1 To UBound(numCom4) - 1
                    ListBox1.Items.Add(lineas(numCom4(i)))
                Next
            Case materia(5)
                For Me.i = 1 To UBound(numCom5) - 1
                    ListBox1.Items.Add(lineas(numCom5(i)))
                Next
        End Select
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'MsgBox(Mid(lineas(5), 1, InStr(lineas(5), " ", CompareMethod.Text)))
        MsgBox(h1(1))
    End Sub
End Class
'fileReader = My.Computer.FileSystem.ReadAllText(Application.ExecutablePath & "test.txt")
'inicioTit(1) = 1
'finalTit(1) = InStr(fileReader, "comisi", CompareMethod.Text) - 1
'For i = 1 To 5
'finalTit(i + 1) = InStr(finalTit(i) + 10, fileReader, "comisi", CompareMethod.Text) - 1
'If finalTit(i + 1) > 50 Then
'inicioTit(i + 1) = InStr(finalTit(i + 1) - 50, fileReader, ".", CompareMethod.Text) - 2
'Else
'inicioTit(i + 1) = InStr(finalTit(i + 1) + 2, fileReader, ".", CompareMethod.Text) - 2
'End If
'materia(i) = Mid(fileReader, inicioTit(i), finalTit(i) - inicioTit(i))
'ComboBox1.Items.Add(materia(i))
'MsgBox(posicion(i) & "    " & final(i))

'Next