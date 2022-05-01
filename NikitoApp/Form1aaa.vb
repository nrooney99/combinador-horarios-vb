Imports System.Text
Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions

Public Class Menu1
    Dim prueba(,,) As Long
    Dim fileReader, test, lineas(), vH(1), a, profesores(,) As String
    Dim numMat(), i, j, k, r, s, h, a1, a2, ac, ac4, ac5, ac6, aux1, ac7, ac8(1000), ac9, numH(), hb(), posh, pos(4), ph, q, w, n, v, x, y, z, p, o, cont(50000), posta(50000, 5, 21), Numcom(,), cm, contCom(), d(), c(), f(), ac3(), numHor(,,), infHor(,,,), vd() As Integer
    Private Sub cargarDatos()
        'fileReader = My.Computer.FileSystem.ReadAllText(Application.ExecutablePath & "test.txt")
        'My.Computer.FileSystem.WriteAllText(Application.StartupPath & "/test.txt", fileReader)

        'limpio mostrador de numeros
        ListBox1.Items.Clear()

        'leo txt con materias cargadas
        fileReader = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "/Materias.txt")

        'separo cada linea
        lineas = Split(fileReader, vbNewLine)


        'ubico el numero de linea de cada materia
        ac = 1
        For Me.i = LBound(lineas) To UBound(lineas) - 1
            If Mid(lineas(i), 1, 5) = "Comis" Then
                ReDim Preserve numMat(ac)
                numMat(ac) = i - 1
                ac = ac + 1
            End If
        Next
        ac = ac - 1
        cm = ac
        Label20.Text = cm & " materias cargadas"
        ReDim Preserve profesores(ac, 9), Numcom(ac, 9), contCom(ac), d(ac), c(ac), f(ac), ac3(ac), numHor(ac, 9, 3), infHor(ac, 9, 3, 3), vd(ac)
        'limpio contador para poder usarlo
        For Me.i = 1 To UBound(contCom)
            contCom(i) = 0
        Next

        'ubico numero de linea de cada Comisión en numcom(numMat, numCom)
        For Me.i = LBound(lineas) To UBound(lineas) - 1
            If Len(lineas(i)) < 5 Then
                Select Case cm
                    Case Is > 6
                        Select Case i
                            Case Is > numMat(7)
                                contCom(7) = contCom(7) + 1
                                Numcom(7, contCom(7)) = i
                            Case Is > numMat(6)
                                contCom(6) = contCom(6) + 1
                                Numcom(6, contCom(6)) = i
                            Case Is > numMat(5)
                                contCom(5) = contCom(5) + 1
                                Numcom(5, contCom(5)) = i
                            Case Is > numMat(4)
                                contCom(4) = contCom(4) + 1
                                Numcom(4, contCom(4)) = i
                            Case Is > numMat(3)
                                contCom(3) = contCom(3) + 1
                                Numcom(3, contCom(3)) = i
                            Case Is > numMat(2)
                                contCom(2) = contCom(2) + 1
                                Numcom(2, contCom(2)) = i
                            Case Is < numMat(2)
                                contCom(1) = contCom(1) + 1
                                Numcom(1, contCom(1)) = i
                        End Select
                    Case Is > 5
                        Select Case i
                            Case Is > numMat(6)
                                contCom(6) = contCom(6) + 1
                                Numcom(6, contCom(6)) = i
                            Case Is > numMat(5)
                                contCom(5) = contCom(5) + 1
                                Numcom(5, contCom(5)) = i
                            Case Is > numMat(4)
                                contCom(4) = contCom(4) + 1
                                Numcom(4, contCom(4)) = i
                            Case Is > numMat(3)
                                contCom(3) = contCom(3) + 1
                                Numcom(3, contCom(3)) = i
                            Case Is > numMat(2)
                                contCom(2) = contCom(2) + 1
                                Numcom(2, contCom(2)) = i
                            Case Is < numMat(2)
                                contCom(1) = contCom(1) + 1
                                Numcom(1, contCom(1)) = i
                        End Select
                    Case Is > 4
                        Select Case i
                            Case Is > numMat(5)
                                contCom(5) = contCom(5) + 1
                                Numcom(5, contCom(5)) = i
                            Case Is > numMat(4)
                                contCom(4) = contCom(4) + 1
                                Numcom(4, contCom(4)) = i
                            Case Is > numMat(3)
                                contCom(3) = contCom(3) + 1
                                Numcom(3, contCom(3)) = i
                            Case Is > numMat(2)
                                contCom(2) = contCom(2) + 1
                                Numcom(2, contCom(2)) = i
                            Case Is < numMat(2)
                                contCom(1) = contCom(1) + 1
                                Numcom(1, contCom(1)) = i
                        End Select
                    Case Is > 3
                        Select Case i
                            Case Is > numMat(4)
                                contCom(4) = contCom(4) + 1
                                Numcom(4, contCom(4)) = i
                            Case Is > numMat(3)
                                contCom(3) = contCom(3) + 1
                                Numcom(3, contCom(3)) = i
                            Case Is > numMat(2)
                                contCom(2) = contCom(2) + 1
                                Numcom(2, contCom(2)) = i
                            Case Is < numMat(2)
                                contCom(1) = contCom(1) + 1
                                Numcom(1, contCom(1)) = i
                        End Select
                    Case Is > 2
                        Select Case i
                            Case Is > numMat(3)
                                contCom(3) = contCom(3) + 1
                                Numcom(3, contCom(3)) = i
                            Case Is > numMat(2)
                                contCom(2) = contCom(2) + 1
                                Numcom(2, contCom(2)) = i
                            Case Is < numMat(2)
                                contCom(1) = contCom(1) + 1
                                Numcom(1, contCom(1)) = i
                        End Select
                    Case Is > 1
                        Select Case i
                            Case Is > numMat(2)
                                contCom(2) = contCom(2) + 1
                                Numcom(2, contCom(2)) = i
                            Case Is < numMat(2)
                                contCom(1) = contCom(1) + 1
                                Numcom(1, contCom(1)) = i
                        End Select
                    Case Is > 0
                        contCom(1) = contCom(1) + 1
                        Numcom(1, contCom(1)) = i
                End Select
            End If
        Next
        For Me.i = 1 To cm
            For Me.j = 1 To 9
                profesores(i, j) = ""
                For Me.k = 0 To 6
                    'On Error GoTo GestionaError
                    'GestionaError:
                    '                    MsgBox(Len(lineas(Numcom(i, j) + k)))
                    'MsgBox(Len(lineas(Numcom(i, j) + k)))
                    For Me.o = 1 To Len(lineas(Numcom(i, j) + k))
                        If Mid(lineas(Numcom(i, j) + k), o, 1) = "," Then
                            profesores(i, j) = profesores(i, j) & ", " & Mid(lineas(Numcom(i, j) + k), 1, o - 1)
                            'MsgBox(profesores(i, j))
                        End If
                    Next
                Next
            Next
        Next
        For Me.i = 1 To cm 'materiasprofesores(i, j)
            For Me.j = 1 To 9 'Comisiones
                For Me.k = 1 To 3 'num de lineas
                    numHor(i, j, k) = 0
                Next
            Next
        Next

        For Me.i = LBound(numHor, 1) To UBound(numHor, 1)
            For Me.j = LBound(numHor, 2) To UBound(numHor, 2)
                If Numcom(i, j) <> 0 Then
                    a = Trim(Mid(lineas(Numcom(i, j) + 1), 1, InStr(lineas(Numcom(i, j) + 1), " ", CompareMethod.Text)))
                    If (a = "Lunes" Or a = "Martes" Or Mid(a, 4, 6) = "rcoles" Or a = "Jueves" Or a = "Viernes") Then
                        numHor(i, j, 1) = Numcom(i, j) + 1
                    End If
                    a = Trim(Mid(lineas(Numcom(i, j) + 2), 1, InStr(lineas(Numcom(i, j) + 2), " ", CompareMethod.Text)))
                    If a = "Lunes" Or a = "Martes" Or Mid(a, 4, 6) = "rcoles" Or a = "Jueves" Or a = "Viernes" Then
                        numHor(i, j, 2) = Numcom(i, j) + 2
                    End If
                    a = Trim(Mid(lineas(Numcom(i, j) + 3), 1, InStr(lineas(Numcom(i, j) + 3), " ", CompareMethod.Text)))
                    If a = "Lunes" Or a = "Martes" Or Mid(a, 4, 6) = "rcoles" Or a = "Jueves" Or a = "Viernes" Then
                        numHor(i, j, 3) = Numcom(i, j) + 3
                    End If
                Else
                    numHor(i, j, 3) = 0
                    numHor(i, j, 2) = 0
                    numHor(i, j, 1) = 0
                End If
            Next
        Next
        'For Me.i = LBound(numHor, 1) To UBound(numHor, 1)
        '    For Me.j = LBound(numHor, 2) To UBound(numHor, 2)
        '        MsgBox(numHor(i, j, 1) & " " & i & " " & j)
        '    Next
        'Next

        For Me.i = LBound(numHor, 1) To UBound(numHor, 1)
            For Me.j = LBound(numHor, 2) To UBound(numHor, 2)
                For Me.k = LBound(numHor, 3) To UBound(numHor, 3)
                    If numHor(i, j, k) <> 0 Then
                        pos(1) = InStr(lineas(numHor(i, j, k)), " ", CompareMethod.Text)
                        Select Case Trim(Mid(lineas(numHor(i, j, k)), 1, pos(1)))
                            Case "Lunes"
                                infHor(i, j, k, 1) = 1
                            Case "Martes"
                                infHor(i, j, k, 1) = 2
                            Case "Jueves"
                                infHor(i, j, k, 1) = 4
                            Case "Viernes"
                                infHor(i, j, k, 1) = 5
                            Case Else
                                If Mid(Trim(Mid(lineas(numHor(i, j, k)), 1, pos(1))), 4, 6) = "rcoles" Then
                                    infHor(i, j, k, 1) = 3
                                End If
                        End Select
                        pos(2) = pos(1) + 1
                        infHor(i, j, k, 2) = Trim(Mid(lineas(numHor(i, j, k)), pos(2), 2))
                        pos(3) = pos(2) + 8
                        infHor(i, j, k, 3) = Trim(Mid(lineas(numHor(i, j, k)), pos(3), 2))
                    End If
                Next
            Next
        Next
        ac4 = 1
        MsgBox(Numcom(7, 1))
        For Me.j = 1 To cm
            For Me.i = 1 To 9
                If Numcom(j, i) = 0 Then
                    vd(j) = i - 1
                    Exit For
                Else
                    If Numcom(j, 9) <> 0 Then
                        vd(j) = 9

                        Exit For
                    End If
                End If
            Next
            ac4 = ac4 * vd(j)
        Next

        ReDim Preserve prueba(ac4, 5, 21)



        For Me.i = 1 To UBound(prueba, 1)
            For Me.j = 1 To UBound(prueba, 2)
                For Me.k = 1 To UBound(prueba, 3)
                    prueba(i, j, k) = 0
                Next
            Next
        Next
        Select Case cm
            Case 7
                Call AlgoritmoMagico7()
            Case 6
                Call AlgoritmoMagico6()
            Case 5
                Call AlgoritmoMagico5()
            Case 4
                Call AlgoritmoMagico4()
            Case 3
                Call AlgoritmoMagico3()
            Case 2
                Call AlgoritmoMagico2()
            Case 1
                Call AlgoritmoMagico1()
            Case 0
                MsgBox("No hay materias cargadas")
        End Select

        For Me.i = 1 To UBound(prueba, 1)
            cont(i) = 1
        Next


        ac7 = 1
        ac6 = 0
        Dim cont2(UBound(prueba, 1), 5)
        ' filtro dia libre
        If CheckBox1.Checked = True Then
            For Me.i = 1 To UBound(prueba, 1)
                For Me.j = 1 To 5
                    cont2(i, j) = 1
                Next
            Next
            For Me.i = 1 To UBound(prueba, 1)
                For Me.o = 1 To 5
                    For Me.p = 8 To 21
                        If prueba(i, o, p) <> 0 Then
                            cont2(i, o) = 0
                        End If
                    Next
                Next
            Next
            For Me.i = 1 To UBound(prueba, 1)
                For Me.j = 1 To 5
                    If cont2(i, j) = 1 Then
                        cont(i) = 1
                        Exit For
                    Else
                        cont(i) = 0
                    End If
                Next
            Next
        End If



        ac = 0
        ac4 = 0
        ' filtro baches
        'If CheckBox2.Checked = True Then
        '    For Me.i = 1 To UBound(prueba, 1)
        '        For Me.o = 1 To 5
        '            For Me.p = 8 To 20
        '                If prueba(i, o, p) <> 0 Then
        '                    ac = ac + 1
        '                    If ac4 = 1 Then
        '                        cont(i) = 0
        '                    End If
        '                Else
        '                    If ac > 0 Then
        '                        If ac4 = 0 Then
        '                            ac4 = 1
        '                        End If
        '                    End If
        '                End If
        '            Next
        '            ac4 = 0
        '            ac = 0
        '        Next
        '    Next
        'End If

        ' filtro salir antes de tal hora
        If CheckBox3.Checked = True Then
            For Me.i = 1 To UBound(prueba, 1)
                For Me.o = 1 To 5
                    For Me.p = 8 To 21
                        If prueba(i, o, p) <> 0 And p > TextBox1.Text - 1 Then cont(i) = 0
                    Next
                Next
            Next
        End If

        ac = 0
        If CheckBox4.Checked = True Then 'filtro a las 8
            For Me.i = 1 To UBound(prueba, 1)
                For Me.o = 1 To 5
                    For Me.p = 8 To 21
                        'MsgBox(i & " " & o & " " & p & " " & prueba(i, o, p) & " " & ac)
                        If prueba(i, o, p) = 0 And p = 8 Then
                            ac = 1
                        End If
                        If prueba(i, o, p) <> 0 And ac = 1 Then cont(i) = 0
                    Next
                    ac = 0
                Next
            Next
        End If



        For Me.i = 1 To UBound(prueba, 1)
            For Me.o = 1 To 5
                For Me.p = 8 To 21
                    If Len(CStr(prueba(i, o, p))) > 2 Then
                        cont(i) = 0
                        Exit For
                    End If
                Next
                If cont(i) = 0 Then
                    Exit For
                End If
            Next
        Next


        ac6 = 0
        For Me.i = 1 To UBound(prueba, 1)
            For Me.j = 1 To 5
                For Me.k = 8 To 20
                    posta(i, j, k) = 0
                Next
            Next
        Next
        For Me.i = 1 To UBound(prueba, 1)
            For Me.j = 1 To 5
                For Me.k = 8 To 21
                    If cont(i) = 1 And prueba(i, j, k) <> 0 Then
                        posta(ac7, j, k) = prueba(i, j, k)
                        'If ac7 = 226 Then MsgBox(i)
                        ac6 = 1
                    End If
                Next
            Next
            If ac6 = 1 Then
                ac7 = ac7 + 1
                ac6 = 0
            End If
        Next
        If RadioButton1.Checked = True Then 'Ordenar por baches
            ReDim hb(ac7 - 1)
            For Me.k = 1 To ac7 - 1
                For Me.i = 1 To 5
                    For Me.j = 8 To 21
                        If a1 = 1 Then
                            If posta(k, i, j) = 0 Then
                                a2 = a2 + 1
                            Else
                                hb(k) = hb(k) + a2
                                a2 = 0
                            End If
                        Else
                            If posta(k, i, j) <> 0 Then
                                a1 = 1
                            End If
                        End If
                    Next
                    a1 = 0
                    a2 = 0
                Next
            Next

            For Me.i = 1 To ac7 - 2
                For Me.j = i + 1 To ac7 - 1
                    If hb(j) < hb(i) Then
                        aux1 = hb(i)
                        hb(i) = hb(j)
                        hb(j) = aux1
                        For Me.k = 1 To 5
                            For Me.o = 8 To 21
                                aux1 = posta(i, k, o)
                                posta(i, k, o) = posta(j, k, o)
                                posta(j, k, o) = aux1
                            Next
                        Next
                    End If
                Next
            Next
            For Me.i = 1 To ac7 - 1 'imprimo horarios
                ListBox1.Items.Add("Horario número " & i & " (" & hb(i) & " hs.)")
            Next
        Else
            For Me.i = 1 To ac7 - 1 'imprimo horarios
                ListBox1.Items.Add("Horario número " & i)
            Next
        End If



    End Sub
    Private Sub mostrarDatos()
        Lunes8.Text = ""
        Lunes9.Text = ""
        Lunes10.Text = ""
        Lunes11.Text = ""
        Lunes12.Text = ""
        Lunes13.Text = ""
        Lunes14.Text = ""
        Lunes15.Text = ""
        Lunes16.Text = ""
        Lunes17.Text = ""
        Lunes18.Text = ""
        Lunes19.Text = ""
        Lunes20.Text = ""
        Martes8.Text = ""
        Martes9.Text = ""
        Martes10.Text = ""
        Martes11.Text = ""
        Martes12.Text = ""
        Martes13.Text = ""
        Martes14.Text = ""
        Martes15.Text = ""
        Martes16.Text = ""
        Martes17.Text = ""
        Martes18.Text = ""
        Martes19.Text = ""
        Martes20.Text = ""
        Miercoles8.Text = ""
        Miercoles9.Text = ""
        Miercoles10.Text = ""
        Miercoles11.Text = ""
        Miercoles12.Text = ""
        Miercoles13.Text = ""
        Miercoles14.Text = ""
        Miercoles15.Text = ""
        Miercoles16.Text = ""
        Miercoles17.Text = ""
        Miercoles18.Text = ""
        Miercoles19.Text = ""
        Miercoles20.Text = ""
        Jueves8.Text = ""
        Jueves9.Text = ""
        Jueves10.Text = ""
        Jueves11.Text = ""
        Jueves12.Text = ""
        Jueves13.Text = ""
        Jueves14.Text = ""
        Jueves15.Text = ""
        Jueves16.Text = ""
        Jueves17.Text = ""
        Jueves18.Text = ""
        Jueves19.Text = ""
        Jueves20.Text = ""
        Viernes8.Text = ""
        Viernes9.Text = ""
        Viernes10.Text = ""
        Viernes11.Text = ""
        Viernes12.Text = ""
        Viernes13.Text = ""
        Viernes14.Text = ""
        Viernes15.Text = ""
        Viernes16.Text = ""
        Viernes17.Text = ""
        Viernes18.Text = ""
        Viernes19.Text = ""
        Viernes20.Text = ""

        Lunes8.BackColor = Color.Silver
        Lunes9.BackColor = Color.Silver
        Lunes10.BackColor = Color.Silver
        Lunes11.BackColor = Color.Silver
        Lunes12.BackColor = Color.Silver
        Lunes13.BackColor = Color.Silver
        Lunes14.BackColor = Color.Silver
        Lunes15.BackColor = Color.Silver
        Lunes16.BackColor = Color.Silver
        Lunes17.BackColor = Color.Silver
        Lunes18.BackColor = Color.Silver
        Lunes19.BackColor = Color.Silver
        Lunes20.BackColor = Color.Silver
        Lunes21.BackColor = Color.Silver
        Martes8.BackColor = Color.Silver
        Martes9.BackColor = Color.Silver
        Martes10.BackColor = Color.Silver
        Martes11.BackColor = Color.Silver
        Martes12.BackColor = Color.Silver
        Martes13.BackColor = Color.Silver
        Martes14.BackColor = Color.Silver
        Martes15.BackColor = Color.Silver
        Martes16.BackColor = Color.Silver
        Martes17.BackColor = Color.Silver
        Martes18.BackColor = Color.Silver
        Martes19.BackColor = Color.Silver
        Martes20.BackColor = Color.Silver
        Martes21.BackColor = Color.Silver
        Miercoles8.BackColor = Color.Silver
        Miercoles9.BackColor = Color.Silver
        Miercoles10.BackColor = Color.Silver
        Miercoles11.BackColor = Color.Silver
        Miercoles12.BackColor = Color.Silver
        Miercoles13.BackColor = Color.Silver
        Miercoles14.BackColor = Color.Silver
        Miercoles15.BackColor = Color.Silver
        Miercoles16.BackColor = Color.Silver
        Miercoles17.BackColor = Color.Silver
        Miercoles18.BackColor = Color.Silver
        Miercoles19.BackColor = Color.Silver
        Miercoles20.BackColor = Color.Silver
        Miercoles21.BackColor = Color.Silver
        Jueves8.BackColor = Color.Silver
        Jueves9.BackColor = Color.Silver
        Jueves10.BackColor = Color.Silver
        Jueves11.BackColor = Color.Silver
        Jueves12.BackColor = Color.Silver
        Jueves13.BackColor = Color.Silver
        Jueves14.BackColor = Color.Silver
        Jueves15.BackColor = Color.Silver
        Jueves16.BackColor = Color.Silver
        Jueves17.BackColor = Color.Silver
        Jueves18.BackColor = Color.Silver
        Jueves19.BackColor = Color.Silver
        Jueves20.BackColor = Color.Silver
        Jueves21.BackColor = Color.Silver
        Viernes8.BackColor = Color.Silver
        Viernes9.BackColor = Color.Silver
        Viernes10.BackColor = Color.Silver
        Viernes11.BackColor = Color.Silver
        Viernes12.BackColor = Color.Silver
        Viernes13.BackColor = Color.Silver
        Viernes14.BackColor = Color.Silver
        Viernes15.BackColor = Color.Silver
        Viernes16.BackColor = Color.Silver
        Viernes17.BackColor = Color.Silver
        Viernes18.BackColor = Color.Silver
        Viernes19.BackColor = Color.Silver
        Viernes20.BackColor = Color.Silver
        Viernes21.BackColor = Color.Silver
        For Me.i = 1 To ac7
            'If ComboBox2.Text = "" Then
            '    ComboBox2.Text = "1"
            'End If
            If ListBox1.SelectedIndex + 1 = i Then
                For Me.k = 1 To 5
                    For Me.j = 8 To 21
                        If posta(i, k, j) <> 0 Then
                            'MsgBox(Mid(posta(i, k, j), 1, 1) & " " & i & " " & k & " " & j)
                            Select Case Mid(posta(i, k, j), 1, 1)
                                Case 1
                                    Mat1.Text = lineas(numMat(1)) & ": Comisión " & lineas(Numcom(1, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(1, Mid(posta(i, k, j), 2, 1)), 2)
                                Case 2
                                    Mat2.Text = lineas(numMat(2)) & ": Comisión " & lineas(Numcom(2, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(2, Mid(posta(i, k, j), 2, 1)), 2)
                                Case 3
                                    Mat3.Text = lineas(numMat(3)) & ": Comisión " & lineas(Numcom(3, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(3, Mid(posta(i, k, j), 2, 1)), 2)
                                Case 4
                                    Mat4.Text = lineas(numMat(4)) & ": Comisión " & lineas(Numcom(4, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(4, Mid(posta(i, k, j), 2, 1)), 2)
                                Case 5
                                    Mat5.Text = lineas(numMat(5)) & ": Comisión " & lineas(Numcom(5, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(5, Mid(posta(i, k, j), 2, 1)), 2)
                                Case 6
                                    Mat6.Text = lineas(numMat(6)) & ": Comisión " & lineas(Numcom(6, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(6, Mid(posta(i, k, j), 2, 1)), 2)
                                Case 7
                                    Mat7.Text = lineas(numMat(7)) & ": Comisión " & lineas(Numcom(7, Mid(posta(i, k, j), 2, 1))) & " - " & "Profesores: " & Mid(profesores(7, Mid(posta(i, k, j), 2, 1)), 2)
                            End Select
                            Select Case k
                                Case 1
                                    Select Case j
                                        Case 8
                                            Lunes8.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes8.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes8.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes8.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes8.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes8.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes8.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes8.BackColor = Color.Pink
                                            End Select
                                        Case 9
                                            Lunes9.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes9.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes9.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes9.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes9.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes9.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes9.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes9.BackColor = Color.Pink
                                            End Select

                                        Case 10
                                            Lunes10.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes10.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes10.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes10.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes10.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes10.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes10.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes10.BackColor = Color.Pink
                                            End Select

                                        Case 11
                                            Lunes11.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes11.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes11.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes11.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes11.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes11.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes11.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes11.BackColor = Color.Pink
                                            End Select

                                        Case 12
                                            Lunes12.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes12.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes12.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes12.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes12.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes12.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes12.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes12.BackColor = Color.Pink
                                            End Select

                                        Case 13
                                            Lunes13.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes13.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes13.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes13.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes13.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes13.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes13.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes13.BackColor = Color.Pink
                                            End Select

                                        Case 14
                                            Lunes14.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes14.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes14.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes14.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes14.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes14.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes14.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes14.BackColor = Color.Pink
                                            End Select

                                        Case 15
                                            Lunes15.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes15.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes15.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes15.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes15.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes15.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes15.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes15.BackColor = Color.Pink
                                            End Select

                                        Case 16
                                            Lunes16.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes16.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes16.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes16.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes16.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes16.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes16.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes16.BackColor = Color.Pink
                                            End Select

                                        Case 17
                                            Lunes17.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes17.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes17.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes17.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes17.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes17.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes17.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes17.BackColor = Color.Pink
                                            End Select

                                        Case 18
                                            Lunes18.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes18.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes18.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes18.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes18.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes18.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes18.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes18.BackColor = Color.Pink
                                            End Select

                                        Case 19
                                            Lunes19.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes19.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes19.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes19.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes19.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes19.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes19.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes19.BackColor = Color.Pink
                                            End Select

                                        Case 20
                                            Lunes20.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes20.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes20.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes20.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes20.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes20.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes20.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes20.BackColor = Color.Pink
                                            End Select
                                        Case 21
                                            Lunes21.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Lunes21.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Lunes21.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Lunes21.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Lunes21.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Lunes21.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Lunes21.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Lunes21.BackColor = Color.Pink
                                            End Select
                                    End Select
                                Case 2
                                    Select Case j
                                        Case 8
                                            Martes8.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes8.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes8.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes8.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes8.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes8.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes8.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes8.BackColor = Color.Pink
                                            End Select
                                        Case 9
                                            Martes9.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes9.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes9.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes9.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes9.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes9.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes9.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes9.BackColor = Color.Pink
                                            End Select

                                        Case 10
                                            Martes10.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes10.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes10.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes10.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes10.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes10.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes10.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes10.BackColor = Color.Pink
                                            End Select

                                        Case 11
                                            Martes11.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes11.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes11.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes11.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes11.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes11.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes11.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes11.BackColor = Color.Pink
                                            End Select

                                        Case 12
                                            Martes12.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes12.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes12.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes12.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes12.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes12.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes12.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes12.BackColor = Color.Pink
                                            End Select

                                        Case 13
                                            Martes13.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes13.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes13.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes13.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes13.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes13.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes13.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes13.BackColor = Color.Pink
                                            End Select

                                        Case 14
                                            Martes14.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes14.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes14.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes14.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes14.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes14.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes14.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes14.BackColor = Color.Pink
                                            End Select

                                        Case 15
                                            Martes15.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes15.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes15.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes15.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes15.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes15.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes15.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes15.BackColor = Color.Pink
                                            End Select

                                        Case 16
                                            Martes16.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes16.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes16.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes16.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes16.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes16.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes16.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes16.BackColor = Color.Pink
                                            End Select

                                        Case 17
                                            Martes17.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes17.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes17.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes17.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes17.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes17.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes17.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes17.BackColor = Color.Pink
                                            End Select

                                        Case 18
                                            Martes18.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes18.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes18.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes18.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes18.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes18.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes18.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes18.BackColor = Color.Pink
                                            End Select

                                        Case 19
                                            Martes19.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes19.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes19.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes19.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes19.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes19.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes19.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes19.BackColor = Color.Pink
                                            End Select

                                        Case 20
                                            Martes20.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes20.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes20.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes20.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes20.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes20.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes20.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes20.BackColor = Color.Pink
                                            End Select
                                        Case 21
                                            Martes21.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Martes21.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Martes21.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Martes21.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Martes21.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Martes21.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Martes21.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Martes21.BackColor = Color.Pink
                                            End Select
                                    End Select
                                Case 3
                                    Select Case j
                                        Case 8
                                            Miercoles8.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles8.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles8.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles8.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles8.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles8.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles8.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles8.BackColor = Color.Pink
                                            End Select
                                        Case 9
                                            Miercoles9.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles9.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles9.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles9.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles9.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles9.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles9.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles9.BackColor = Color.Pink
                                            End Select

                                        Case 10
                                            Miercoles10.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles10.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles10.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles10.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles10.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles10.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles10.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles10.BackColor = Color.Pink
                                            End Select

                                        Case 11
                                            Miercoles11.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles11.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles11.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles11.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles11.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles11.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles11.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles11.BackColor = Color.Pink
                                            End Select

                                        Case 12
                                            Miercoles12.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles12.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles12.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles12.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles12.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles12.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles12.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles12.BackColor = Color.Pink
                                            End Select

                                        Case 13
                                            Miercoles13.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles13.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles13.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles13.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles13.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles13.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles13.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles13.BackColor = Color.Pink
                                            End Select

                                        Case 14
                                            Miercoles14.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles14.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles14.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles14.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles14.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles14.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles14.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles14.BackColor = Color.Pink
                                            End Select

                                        Case 15
                                            Miercoles15.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles15.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles15.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles15.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles15.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles15.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles15.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles15.BackColor = Color.Pink
                                            End Select

                                        Case 16
                                            Miercoles16.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles16.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles16.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles16.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles16.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles16.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles16.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles16.BackColor = Color.Pink
                                            End Select

                                        Case 17
                                            Miercoles17.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles17.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles17.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles17.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles17.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles17.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles17.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles17.BackColor = Color.Pink
                                            End Select

                                        Case 18
                                            Miercoles18.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles18.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles18.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles18.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles18.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles18.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles18.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles18.BackColor = Color.Pink
                                            End Select

                                        Case 19
                                            Miercoles19.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles19.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles19.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles19.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles19.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles19.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles19.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles19.BackColor = Color.Pink
                                            End Select

                                        Case 20
                                            Miercoles20.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles20.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles20.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles20.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles20.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles20.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles20.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles20.BackColor = Color.Pink
                                            End Select
                                        Case 21
                                            Miercoles21.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Miercoles21.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Miercoles21.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Miercoles21.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Miercoles21.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Miercoles21.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Miercoles21.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Miercoles21.BackColor = Color.Pink
                                            End Select
                                    End Select
                                Case 4
                                    Select Case j
                                        Case 8
                                            Jueves8.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves8.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves8.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves8.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves8.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves8.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves8.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves8.BackColor = Color.Pink
                                            End Select
                                        Case 9
                                            Jueves9.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves9.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves9.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves9.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves9.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves9.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves9.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves9.BackColor = Color.Pink
                                            End Select

                                        Case 10
                                            Jueves10.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves10.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves10.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves10.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves10.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves10.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves10.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves10.BackColor = Color.Pink
                                            End Select

                                        Case 11
                                            Jueves11.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves11.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves11.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves11.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves11.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves11.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves11.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves11.BackColor = Color.Pink
                                            End Select

                                        Case 12
                                            Jueves12.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves12.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves12.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves12.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves12.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves12.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves12.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves12.BackColor = Color.Pink
                                            End Select

                                        Case 13
                                            Jueves13.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves13.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves13.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves13.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves13.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves13.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves13.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves13.BackColor = Color.Pink
                                            End Select

                                        Case 14
                                            Jueves14.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves14.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves14.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves14.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves14.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves14.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves14.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves14.BackColor = Color.Pink
                                            End Select

                                        Case 15
                                            Jueves15.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves15.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves15.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves15.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves15.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves15.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves15.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves15.BackColor = Color.Pink
                                            End Select

                                        Case 16
                                            Jueves16.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves16.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves16.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves16.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves16.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves16.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves16.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves16.BackColor = Color.Pink
                                            End Select

                                        Case 17
                                            Jueves17.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves17.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves17.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves17.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves17.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves17.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves17.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves17.BackColor = Color.Pink
                                            End Select

                                        Case 18
                                            Jueves18.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves18.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves18.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves18.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves18.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves18.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves18.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves18.BackColor = Color.Pink
                                            End Select

                                        Case 19
                                            Jueves19.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves19.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves19.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves19.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves19.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves19.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves19.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves19.BackColor = Color.Pink
                                            End Select

                                        Case 20
                                            Jueves20.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves20.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves20.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves20.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves20.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves20.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves20.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves20.BackColor = Color.Pink
                                            End Select
                                        Case 21
                                            Jueves21.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Jueves21.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Jueves21.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Jueves21.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Jueves21.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Jueves21.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Jueves21.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Jueves21.BackColor = Color.Pink
                                            End Select
                                    End Select
                                Case 5
                                    Select Case j
                                        Case 8
                                            Viernes8.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes8.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes8.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes8.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes8.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes8.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes8.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes8.BackColor = Color.Pink
                                            End Select
                                        Case 9
                                            Viernes9.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes9.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes9.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes9.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes9.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes9.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes9.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes9.BackColor = Color.Pink
                                            End Select

                                        Case 10
                                            Viernes10.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes10.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes10.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes10.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes10.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes10.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes10.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes10.BackColor = Color.Pink
                                            End Select

                                        Case 11
                                            Viernes11.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes11.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes11.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes11.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes11.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes11.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes11.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes11.BackColor = Color.Pink
                                            End Select

                                        Case 12
                                            Viernes12.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes12.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes12.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes12.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes12.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes12.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes12.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes12.BackColor = Color.Pink
                                            End Select

                                        Case 13
                                            Viernes13.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes13.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes13.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes13.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes13.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes13.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes13.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes13.BackColor = Color.Pink
                                            End Select

                                        Case 14
                                            Viernes14.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes14.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes14.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes14.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes14.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes14.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes14.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes14.BackColor = Color.Pink
                                            End Select

                                        Case 15
                                            Viernes15.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes15.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes15.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes15.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes15.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes15.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes15.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes15.BackColor = Color.Pink
                                            End Select

                                        Case 16
                                            Viernes16.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes16.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes16.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes16.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes16.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes16.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes16.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes16.BackColor = Color.Pink
                                            End Select

                                        Case 17
                                            Viernes17.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes17.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes17.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes17.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes17.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes17.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes17.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes17.BackColor = Color.Pink
                                            End Select

                                        Case 18
                                            Viernes18.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes18.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes18.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes18.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes18.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes18.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes18.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes18.BackColor = Color.Pink
                                            End Select

                                        Case 19
                                            Viernes19.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes19.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes19.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes19.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes19.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes19.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes19.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes19.BackColor = Color.Pink
                                            End Select

                                        Case 20
                                            Viernes20.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes20.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes20.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes20.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes20.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes20.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes20.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes20.BackColor = Color.Pink
                                            End Select
                                        Case 21
                                            Viernes21.Text = (lineas(numMat(Mid(posta(i, k, j), 1, 1))))
                                            Select Case Mid(posta(i, k, j), 1, 1)
                                                Case 1
                                                    Viernes21.BackColor = Color.DarkKhaki
                                                Case 2
                                                    Viernes21.BackColor = Color.CornflowerBlue
                                                Case 3
                                                    Viernes21.BackColor = Color.DarkSalmon
                                                Case 4
                                                    Viernes21.BackColor = Color.DarkSeaGreen
                                                Case 5
                                                    Viernes21.BackColor = Color.LightSkyBlue
                                                Case 6
                                                    Viernes21.BackColor = Color.MediumOrchid
                                                Case 7
                                                    Viernes21.BackColor = Color.Pink
                                            End Select
                                    End Select
                            End Select
                        End If
                    Next
                Next
            End If
        Next

    End Sub
    Private Sub AlgoritmoMagico7()
        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia
            For Me.w = 1 To vd(2)
                For Me.x = 1 To vd(3)
                    For Me.y = 1 To vd(4)
                        For Me.z = 1 To vd(5)
                            For Me.r = 1 To vd(6)
                                For Me.s = 1 To vd(7)
                                    For Me.o = 1 To 3
                                        d(1) = infHor(1, v, o, 1) 'dia de la semana
                                        c(1) = infHor(1, v, o, 2)
                                        f(1) = infHor(1, v, o, 3)
                                        If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                                        If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                                            prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                                            prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                                            If f(1) - c(1) > 2 Then
                                                prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                                                If f(1) - c(1) > 3 Then
                                                    prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                                                End If
                                            End If
                                        End If

                                        d(2) = infHor(2, w, o, 1)
                                        c(2) = infHor(2, w, o, 2)
                                        f(2) = infHor(2, w, o, 3)
                                        If Len(CStr(prueba(ac7, d(2), c(2)))) < 7 Then
                                            prueba(ac7, d(2), c(2)) = prueba(ac7, d(2), c(2)) & "2" & w
                                            prueba(ac7, d(2), c(2) + 1) = prueba(ac7, d(2), c(2) + 1) & "2" & w
                                            If f(2) - c(2) > 2 Then
                                                prueba(ac7, d(2), c(2) + 2) = prueba(ac7, d(2), c(2) + 2) & "2" & w
                                                If f(2) - c(2) > 3 Then
                                                    prueba(ac7, d(2), c(2) + 3) = prueba(ac7, d(2), c(2) + 3) & "2" & w
                                                End If
                                            End If
                                        End If


                                        d(3) = infHor(3, x, o, 1)
                                        c(3) = infHor(3, x, o, 2)
                                        f(3) = infHor(3, x, o, 3)
                                        If Len(CStr(prueba(ac7, d(3), c(3)))) < 7 Then
                                            prueba(ac7, d(3), c(3)) = prueba(ac7, d(3), c(3)) & "3" & x
                                            prueba(ac7, d(3), c(3) + 1) = prueba(ac7, d(3), c(3) + 1) & "3" & x
                                            If f(3) - c(3) > 2 Then
                                                prueba(ac7, d(3), c(3) + 2) = prueba(ac7, d(3), c(3) + 2) & "3" & x
                                                If f(3) - c(3) > 3 Then
                                                    prueba(ac7, d(3), c(3) + 3) = prueba(ac7, d(3), c(3) + 3) & "3" & x
                                                End If
                                            End If
                                        End If
                                        d(4) = infHor(4, y, o, 1)
                                        c(4) = infHor(4, y, o, 2)
                                        f(4) = infHor(4, y, o, 3)
                                        If Len(CStr(prueba(ac7, d(4), c(4)))) < 7 Then
                                            prueba(ac7, d(4), c(4)) = prueba(ac7, d(4), c(4)) & "4" & y
                                            prueba(ac7, d(4), c(4) + 1) = prueba(ac7, d(4), c(4) + 1) & "4" & y
                                            If f(4) - c(4) > 2 Then
                                                prueba(ac7, d(4), c(4) + 2) = prueba(ac7, d(4), c(4) + 2) & "4" & y
                                                If f(4) - c(4) > 3 Then
                                                    prueba(ac7, d(4), c(4) + 3) = prueba(ac7, d(4), c(4) + 3) & "4" & y
                                                End If
                                            End If
                                        End If
                                        d(5) = infHor(5, z, o, 1)
                                        c(5) = infHor(5, z, o, 2)
                                        f(5) = infHor(5, z, o, 3)
                                        If Len(CStr(prueba(ac7, d(5), c(5)))) < 7 Then
                                            prueba(ac7, d(5), c(5)) = prueba(ac7, d(5), c(5)) & "5" & z
                                            prueba(ac7, d(5), c(5) + 1) = prueba(ac7, d(5), c(5) + 1) & "5" & z
                                            If f(5) - c(5) > 2 Then
                                                prueba(ac7, d(5), c(5) + 2) = prueba(ac7, d(5), c(5) + 2) & "5" & z
                                                If f(5) - c(5) > 3 Then
                                                    prueba(ac7, d(5), c(5) + 3) = prueba(ac7, d(5), c(5) + 3) & "5" & z
                                                End If
                                            End If
                                        End If
                                        d(6) = infHor(6, r, o, 1)
                                        c(6) = infHor(6, r, o, 2)
                                        f(6) = infHor(6, r, o, 3)
                                        If Len(CStr(prueba(ac7, d(6), c(6)))) < 7 Then

                                            prueba(ac7, d(6), c(6)) = prueba(ac7, d(6), c(6)) & "6" & r
                                            prueba(ac7, d(6), c(6) + 1) = prueba(ac7, d(6), c(6) + 1) & "6" & r
                                            If f(6) - c(6) > 2 Then
                                                prueba(ac7, d(6), c(6) + 2) = prueba(ac7, d(6), c(6) + 2) & "6" & r
                                                If f(6) - c(6) > 3 Then
                                                    prueba(ac7, d(6), c(6) + 3) = prueba(ac7, d(6), c(6) + 3) & "6" & r
                                                End If
                                            End If
                                        End If
                                        d(7) = infHor(7, s, o, 1)
                                        c(7) = infHor(7, s, o, 2)
                                        f(7) = infHor(7, s, o, 3)
                                        If Len(CStr(prueba(ac7, d(7), c(7)))) < 7 Then

                                            prueba(ac7, d(7), c(7)) = prueba(ac7, d(7), c(7)) & "7" & s
                                            prueba(ac7, d(7), c(7) + 1) = prueba(ac7, d(7), c(7) + 1) & "7" & s
                                            If f(7) - c(7) > 2 Then
                                                prueba(ac7, d(7), c(7) + 2) = prueba(ac7, d(7), c(7) + 2) & "7" & s
                                                If f(7) - c(7) > 3 Then
                                                    prueba(ac7, d(7), c(7) + 3) = prueba(ac7, d(7), c(7) + 3) & "7" & s
                                                End If
                                            End If
                                        End If
                                    Next
                                    ac7 = ac7 + 1
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    End Sub
    Private Sub AlgoritmoMagico6()

        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia 
            For Me.w = 1 To vd(2)
                For Me.x = 1 To vd(3)
                    For Me.y = 1 To vd(4)
                        For Me.z = 1 To vd(5)
                            For Me.r = 1 To vd(6)
                                For Me.o = 1 To 3
                                    d(1) = infHor(1, v, o, 1) 'dia de la semana
                                    c(1) = infHor(1, v, o, 2)
                                    f(1) = infHor(1, v, o, 3)
                                    If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                                    If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                                        prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                                        prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                                        If f(1) - c(1) > 2 Then
                                            prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                                            If f(1) - c(1) > 3 Then
                                                prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                                            End If
                                        End If
                                    End If

                                    d(2) = infHor(2, w, o, 1)
                                    c(2) = infHor(2, w, o, 2)
                                    f(2) = infHor(2, w, o, 3)
                                    If Len(CStr(prueba(ac7, d(2), c(2)))) < 7 Then
                                        prueba(ac7, d(2), c(2)) = prueba(ac7, d(2), c(2)) & "2" & w
                                        prueba(ac7, d(2), c(2) + 1) = prueba(ac7, d(2), c(2) + 1) & "2" & w
                                        If f(2) - c(2) > 2 Then
                                            prueba(ac7, d(2), c(2) + 2) = prueba(ac7, d(2), c(2) + 2) & "2" & w
                                            If f(2) - c(2) > 3 Then
                                                prueba(ac7, d(2), c(2) + 3) = prueba(ac7, d(2), c(2) + 3) & "2" & w
                                            End If
                                        End If
                                    End If


                                    d(3) = infHor(3, x, o, 1)
                                    c(3) = infHor(3, x, o, 2)
                                    f(3) = infHor(3, x, o, 3)
                                    If Len(CStr(prueba(ac7, d(3), c(3)))) < 7 Then
                                        prueba(ac7, d(3), c(3)) = prueba(ac7, d(3), c(3)) & "3" & x
                                        prueba(ac7, d(3), c(3) + 1) = prueba(ac7, d(3), c(3) + 1) & "3" & x
                                        If f(3) - c(3) > 2 Then
                                            prueba(ac7, d(3), c(3) + 2) = prueba(ac7, d(3), c(3) + 2) & "3" & x
                                            If f(3) - c(3) > 3 Then
                                                prueba(ac7, d(3), c(3) + 3) = prueba(ac7, d(3), c(3) + 3) & "3" & x
                                            End If
                                        End If
                                    End If
                                    d(4) = infHor(4, y, o, 1)
                                    c(4) = infHor(4, y, o, 2)
                                    f(4) = infHor(4, y, o, 3)
                                    If Len(CStr(prueba(ac7, d(4), c(4)))) < 7 Then
                                        prueba(ac7, d(4), c(4)) = prueba(ac7, d(4), c(4)) & "4" & y
                                        prueba(ac7, d(4), c(4) + 1) = prueba(ac7, d(4), c(4) + 1) & "4" & y
                                        If f(4) - c(4) > 2 Then
                                            prueba(ac7, d(4), c(4) + 2) = prueba(ac7, d(4), c(4) + 2) & "4" & y
                                            If f(4) - c(4) > 3 Then
                                                prueba(ac7, d(4), c(4) + 3) = prueba(ac7, d(4), c(4) + 3) & "4" & y
                                            End If
                                        End If
                                    End If
                                    d(5) = infHor(5, z, o, 1)
                                    c(5) = infHor(5, z, o, 2)
                                    f(5) = infHor(5, z, o, 3)
                                    If Len(CStr(prueba(ac7, d(5), c(5)))) < 7 Then
                                        prueba(ac7, d(5), c(5)) = prueba(ac7, d(5), c(5)) & "5" & z
                                        prueba(ac7, d(5), c(5) + 1) = prueba(ac7, d(5), c(5) + 1) & "5" & z
                                        If f(5) - c(5) > 2 Then
                                            prueba(ac7, d(5), c(5) + 2) = prueba(ac7, d(5), c(5) + 2) & "5" & z
                                            If f(5) - c(5) > 3 Then
                                                prueba(ac7, d(5), c(5) + 3) = prueba(ac7, d(5), c(5) + 3) & "5" & z
                                            End If
                                        End If
                                    End If
                                    d(6) = infHor(6, r, o, 1)
                                    c(6) = infHor(6, r, o, 2)
                                    f(6) = infHor(6, r, o, 3)
                                    If Len(CStr(prueba(ac7, d(6), c(6)))) < 7 Then

                                        prueba(ac7, d(6), c(6)) = prueba(ac7, d(6), c(6)) & "6" & r
                                        prueba(ac7, d(6), c(6) + 1) = prueba(ac7, d(6), c(6) + 1) & "6" & r
                                        If f(6) - c(6) > 2 Then
                                            prueba(ac7, d(6), c(6) + 2) = prueba(ac7, d(6), c(6) + 2) & "6" & r
                                            If f(6) - c(6) > 3 Then
                                                prueba(ac7, d(6), c(6) + 3) = prueba(ac7, d(6), c(6) + 3) & "6" & r
                                            End If
                                        End If
                                    End If
                                Next
                                ac7 = ac7 + 1
                            Next
                        Next
                    Next
                Next
            Next
        Next
    End Sub
    Private Sub AlgoritmoMagico5()

        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia 
            For Me.w = 1 To vd(2)
                For Me.x = 1 To vd(3)
                    For Me.y = 1 To vd(4)
                        For Me.z = 1 To vd(5)
                            For Me.o = 1 To 3
                                d(1) = infHor(1, v, o, 1) 'dia de la semana
                                c(1) = infHor(1, v, o, 2)
                                f(1) = infHor(1, v, o, 3)
                                If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                                If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                                    prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                                    prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                                    If f(1) - c(1) > 2 Then
                                        prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                                        If f(1) - c(1) > 3 Then
                                            prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                                        End If
                                    End If
                                End If

                                d(2) = infHor(2, w, o, 1)
                                c(2) = infHor(2, w, o, 2)
                                f(2) = infHor(2, w, o, 3)
                                If Len(CStr(prueba(ac7, d(2), c(2)))) < 7 Then
                                    prueba(ac7, d(2), c(2)) = prueba(ac7, d(2), c(2)) & "2" & w
                                    prueba(ac7, d(2), c(2) + 1) = prueba(ac7, d(2), c(2) + 1) & "2" & w
                                    If f(2) - c(2) > 2 Then
                                        prueba(ac7, d(2), c(2) + 2) = prueba(ac7, d(2), c(2) + 2) & "2" & w
                                        If f(2) - c(2) > 3 Then
                                            prueba(ac7, d(2), c(2) + 3) = prueba(ac7, d(2), c(2) + 3) & "2" & w
                                        End If
                                    End If
                                End If


                                d(3) = infHor(3, x, o, 1)
                                c(3) = infHor(3, x, o, 2)
                                f(3) = infHor(3, x, o, 3)
                                If Len(CStr(prueba(ac7, d(3), c(3)))) < 7 Then
                                    prueba(ac7, d(3), c(3)) = prueba(ac7, d(3), c(3)) & "3" & x
                                    prueba(ac7, d(3), c(3) + 1) = prueba(ac7, d(3), c(3) + 1) & "3" & x
                                    If f(3) - c(3) > 2 Then
                                        prueba(ac7, d(3), c(3) + 2) = prueba(ac7, d(3), c(3) + 2) & "3" & x
                                        If f(3) - c(3) > 3 Then
                                            prueba(ac7, d(3), c(3) + 3) = prueba(ac7, d(3), c(3) + 3) & "3" & x
                                        End If
                                    End If
                                End If
                                d(4) = infHor(4, y, o, 1)
                                c(4) = infHor(4, y, o, 2)
                                f(4) = infHor(4, y, o, 3)
                                If Len(CStr(prueba(ac7, d(4), c(4)))) < 7 Then
                                    prueba(ac7, d(4), c(4)) = prueba(ac7, d(4), c(4)) & "4" & y
                                    prueba(ac7, d(4), c(4) + 1) = prueba(ac7, d(4), c(4) + 1) & "4" & y
                                    If f(4) - c(4) > 2 Then
                                        prueba(ac7, d(4), c(4) + 2) = prueba(ac7, d(4), c(4) + 2) & "4" & y
                                        If f(4) - c(4) > 3 Then
                                            prueba(ac7, d(4), c(4) + 3) = prueba(ac7, d(4), c(4) + 3) & "4" & y
                                        End If
                                    End If
                                End If
                                d(5) = infHor(5, z, o, 1)
                                c(5) = infHor(5, z, o, 2)
                                f(5) = infHor(5, z, o, 3)
                                If Len(CStr(prueba(ac7, d(5), c(5)))) < 7 Then
                                    prueba(ac7, d(5), c(5)) = prueba(ac7, d(5), c(5)) & "5" & z
                                    prueba(ac7, d(5), c(5) + 1) = prueba(ac7, d(5), c(5) + 1) & "5" & z
                                    If f(5) - c(5) > 2 Then
                                        prueba(ac7, d(5), c(5) + 2) = prueba(ac7, d(5), c(5) + 2) & "5" & z
                                        If f(5) - c(5) > 3 Then
                                            prueba(ac7, d(5), c(5) + 3) = prueba(ac7, d(5), c(5) + 3) & "5" & z
                                        End If
                                    End If
                                End If
                            Next
                            ac7 = ac7 + 1
                        Next
                    Next
                Next
            Next
        Next
    End Sub
    Private Sub AlgoritmoMagico4()

        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia 
            For Me.w = 1 To vd(2)
                For Me.x = 1 To vd(3)
                    For Me.y = 1 To vd(4)
                        For Me.o = 1 To 3
                            d(1) = infHor(1, v, o, 1) 'dia de la semana
                            c(1) = infHor(1, v, o, 2)
                            f(1) = infHor(1, v, o, 3)
                            If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                            If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                                prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                                prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                                If f(1) - c(1) > 2 Then
                                    prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                                    If f(1) - c(1) > 3 Then
                                        prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                                    End If
                                End If
                            End If

                            d(2) = infHor(2, w, o, 1)
                            c(2) = infHor(2, w, o, 2)
                            f(2) = infHor(2, w, o, 3)
                            If Len(CStr(prueba(ac7, d(2), c(2)))) < 7 Then
                                prueba(ac7, d(2), c(2)) = prueba(ac7, d(2), c(2)) & "2" & w
                                prueba(ac7, d(2), c(2) + 1) = prueba(ac7, d(2), c(2) + 1) & "2" & w
                                If f(2) - c(2) > 2 Then
                                    prueba(ac7, d(2), c(2) + 2) = prueba(ac7, d(2), c(2) + 2) & "2" & w
                                    If f(2) - c(2) > 3 Then
                                        prueba(ac7, d(2), c(2) + 3) = prueba(ac7, d(2), c(2) + 3) & "2" & w
                                    End If
                                End If
                            End If


                            d(3) = infHor(3, x, o, 1)
                            c(3) = infHor(3, x, o, 2)
                            f(3) = infHor(3, x, o, 3)
                            If Len(CStr(prueba(ac7, d(3), c(3)))) < 7 Then
                                prueba(ac7, d(3), c(3)) = prueba(ac7, d(3), c(3)) & "3" & x
                                prueba(ac7, d(3), c(3) + 1) = prueba(ac7, d(3), c(3) + 1) & "3" & x
                                If f(3) - c(3) > 2 Then
                                    prueba(ac7, d(3), c(3) + 2) = prueba(ac7, d(3), c(3) + 2) & "3" & x
                                    If f(3) - c(3) > 3 Then
                                        prueba(ac7, d(3), c(3) + 3) = prueba(ac7, d(3), c(3) + 3) & "3" & x
                                    End If
                                End If
                            End If
                            d(4) = infHor(4, y, o, 1)
                            c(4) = infHor(4, y, o, 2)
                            f(4) = infHor(4, y, o, 3)
                            If Len(CStr(prueba(ac7, d(4), c(4)))) < 7 Then
                                prueba(ac7, d(4), c(4)) = prueba(ac7, d(4), c(4)) & "4" & y
                                prueba(ac7, d(4), c(4) + 1) = prueba(ac7, d(4), c(4) + 1) & "4" & y
                                If f(4) - c(4) > 2 Then
                                    prueba(ac7, d(4), c(4) + 2) = prueba(ac7, d(4), c(4) + 2) & "4" & y
                                    If f(4) - c(4) > 3 Then
                                        prueba(ac7, d(4), c(4) + 3) = prueba(ac7, d(4), c(4) + 3) & "4" & y
                                    End If
                                End If
                            End If
                        Next
                        ac7 = ac7 + 1
                    Next
                Next
            Next
        Next
    End Sub
    Private Sub AlgoritmoMagico1()

        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia 
            For Me.o = 1 To 3
                d(1) = infHor(1, v, o, 1) 'dia de la semana
                c(1) = infHor(1, v, o, 2)
                f(1) = infHor(1, v, o, 3)
                If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                    prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                    prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                    If f(1) - c(1) > 2 Then
                        prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                        If f(1) - c(1) > 3 Then
                            prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                        End If
                    End If
                End If
            Next
            ac7 = ac7 + 1
        Next
    End Sub
    Private Sub AlgoritmoMagico3()

        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia 
            For Me.w = 1 To vd(2)
                For Me.x = 1 To vd(3)
                        For Me.o = 1 To 3
                            d(1) = infHor(1, v, o, 1) 'dia de la semana
                            c(1) = infHor(1, v, o, 2)
                            f(1) = infHor(1, v, o, 3)
                            If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                            If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                                prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                                prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                                If f(1) - c(1) > 2 Then
                                    prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                                    If f(1) - c(1) > 3 Then
                                        prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                                    End If
                                End If
                            End If

                            d(2) = infHor(2, w, o, 1)
                            c(2) = infHor(2, w, o, 2)
                            f(2) = infHor(2, w, o, 3)
                            If Len(CStr(prueba(ac7, d(2), c(2)))) < 7 Then
                                prueba(ac7, d(2), c(2)) = prueba(ac7, d(2), c(2)) & "2" & w
                                prueba(ac7, d(2), c(2) + 1) = prueba(ac7, d(2), c(2) + 1) & "2" & w
                                If f(2) - c(2) > 2 Then
                                    prueba(ac7, d(2), c(2) + 2) = prueba(ac7, d(2), c(2) + 2) & "2" & w
                                    If f(2) - c(2) > 3 Then
                                        prueba(ac7, d(2), c(2) + 3) = prueba(ac7, d(2), c(2) + 3) & "2" & w
                                    End If
                                End If
                            End If


                            d(3) = infHor(3, x, o, 1)
                            c(3) = infHor(3, x, o, 2)
                            f(3) = infHor(3, x, o, 3)
                            If Len(CStr(prueba(ac7, d(3), c(3)))) < 7 Then
                                prueba(ac7, d(3), c(3)) = prueba(ac7, d(3), c(3)) & "3" & x
                                prueba(ac7, d(3), c(3) + 1) = prueba(ac7, d(3), c(3) + 1) & "3" & x
                                If f(3) - c(3) > 2 Then
                                    prueba(ac7, d(3), c(3) + 2) = prueba(ac7, d(3), c(3) + 2) & "3" & x
                                    If f(3) - c(3) > 3 Then
                                        prueba(ac7, d(3), c(3) + 3) = prueba(ac7, d(3), c(3) + 3) & "3" & x
                                    End If
                                End If
                            End If
                        Next
                        ac7 = ac7 + 1
                    Next
                Next
            Next
    End Sub
    Private Sub AlgoritmoMagico2()

        ac7 = 1
        For Me.v = 1 To vd(1) 'recorre Comisiones de cada materia 
            For Me.w = 1 To vd(2)
                For Me.o = 1 To 3
                    d(1) = infHor(1, v, o, 1) 'dia de la semana
                    c(1) = infHor(1, v, o, 2)
                    f(1) = infHor(1, v, o, 3)
                    If ac7 > ac4 Or d(1) > 5 Or c(1) > 20 Then MsgBox(ac7 & " " & d(1) & " " & c(1))
                    If Len(CStr(prueba(ac7, d(1), c(1)))) < 7 Then
                        prueba(ac7, d(1), c(1)) = prueba(ac7, d(1), c(1)) & "1" & v
                        prueba(ac7, d(1), c(1) + 1) = prueba(ac7, d(1), c(1) + 1) & "1" & v
                        If f(1) - c(1) > 2 Then
                            prueba(ac7, d(1), c(1) + 2) = prueba(ac7, d(1), c(1) + 2) & "1" & v
                            If f(1) - c(1) > 3 Then
                                prueba(ac7, d(1), c(1) + 3) = prueba(ac7, d(1), c(1) + 3) & "1" & v
                            End If
                        End If
                    End If

                    d(2) = infHor(2, w, o, 1)
                    c(2) = infHor(2, w, o, 2)
                    f(2) = infHor(2, w, o, 3)
                    If Len(CStr(prueba(ac7, d(2), c(2)))) < 7 Then
                        prueba(ac7, d(2), c(2)) = prueba(ac7, d(2), c(2)) & "2" & w
                        prueba(ac7, d(2), c(2) + 1) = prueba(ac7, d(2), c(2) + 1) & "2" & w
                        If f(2) - c(2) > 2 Then
                            prueba(ac7, d(2), c(2) + 2) = prueba(ac7, d(2), c(2) + 2) & "2" & w
                            If f(2) - c(2) > 3 Then
                                prueba(ac7, d(2), c(2) + 3) = prueba(ac7, d(2), c(2) + 3) & "2" & w
                            End If
                        End If
                    End If
                Next
                ac7 = ac7 + 1
            Next
        Next
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        mostrarDatos()
    End Sub
    Private Sub Menu1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call cargarDatos()
        'If Today > "10/6/2019" Then
        '    MsgBox("Versión vencida, contacte al developer Rooney para conseguir la nueva.")
        '    End
        'End If
        test = ""
        For Me.i = 1 To 5
            For Me.j = 8 To 21
                test = test & " " & posta(1, i, j)
            Next
        Next
        'MsgBox(test)
        'MsgBox(numHor(6, 1, 1))
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        cargarDatos()
    End Sub

    'Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)
    '    cargarDatos()
    'End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        cargarDatos()
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        cargarDatos()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        cargarDatos()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        cargarDatos()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("Esta es la versión Beta, para que funcione optimamente se deben cargar como máximo 6 materias de 9 comisiones cada una. El nombre de las comisiones no pueden tener mas de dos letras, por ejemplo a la comisión 'MNPE' de Mate III hay que acortarle el nombre a 'MN'.")
    End Sub
End Class