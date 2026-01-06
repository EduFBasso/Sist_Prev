Sub OrdenarResultados(arr As Variant, coluna As Long)

    Dim i As Long, j As Long
    Dim temp1, temp2, temp3

    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)

            If LCase(arr(i, coluna)) > LCase(arr(j, coluna)) Then

                ' Trocar ID
                temp1 = arr(i, 1)
                arr(i, 1) = arr(j, 1)
                arr(j, 1) = temp1

                ' Trocar Nome
                temp2 = arr(i, 2)
                arr(i, 2) = arr(j, 2)
                arr(j, 2) = temp2

                ' Trocar CPF
                temp3 = arr(i, 3)
                arr(i, 3) = arr(j, 3)
                arr(j, 3) = temp3

            End If

        Next j
    Next i

End Sub
