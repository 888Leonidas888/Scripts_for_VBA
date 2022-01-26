Attribute VB_Name = "M�dulo1_udf"
Function createPassword(Optional ByVal lenght = 6) As String
      Dim password$, character$, n%, i%, j%, a%
      Dim character_excluded(0 To 12) As Byte
      'n�meros del 48 al 57
      'letras may�sculas 65 al 90
      'letras min�sculas 97 al 122
   
      Rem Application.Volatile False
      
      For i = 58 To 64
            character_excluded(n) = i
            n = n + 1
      Next i
      
      For i = 91 To 96
            character_excluded(n) = i
            n = n + 1
      Next i

      For a = 1 To lenght
otra_vez:
            Randomize
            character = Int((122 - 48 + 1) * Rnd + 48)
            
            For j = 0 To 12
                  If character = character_excluded(j) Then
                        GoTo otra_vez
                  End If
            Next j
            password = password & Chr(character)
      Next a
      
      createPassword = password
      rem estamos agregando un comentario para probar nuestro rama que hemos en creado en nuestro
      rem repositorio
End Function






