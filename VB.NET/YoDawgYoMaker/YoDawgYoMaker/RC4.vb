Module RC4
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' RC4 encryption and decryption.
    ''' </summary>
    ''' <param name="plaintxt"></param>
    ''' <param name="password"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' [jamie] 8/18/2005 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function RC4EnDeCrypt(ByVal plaintxt As String, ByVal password As String) As String
        Dim temp As Int32
        Dim a, i, j, k As Int32
        Dim cipherby As Int32
        Dim cipher As String
        Dim sbox As Int32()
        Dim key As Int32()
        i = 0
        j = 0
        RC4Initialize(password, sbox, key)
        For a = 1 To Len(plaintxt)
            i = (i + 1) Mod 256
            j = (j + sbox(i)) Mod 256
            temp = sbox(i)
            sbox(i) = sbox(j)
            sbox(j) = temp
            k = sbox((sbox(i) + sbox(j)) Mod 256)
            cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
            cipher = cipher & Chr(cipherby)
        Next
        Return cipher
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This routine called by EnDeCrypt function. Initializes the
    ''' sbox and the key array)
    ''' </summary>
    ''' <param name="password"></param>
    ''' <param name="sbox"></param>
    ''' <param name="key"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' [jamie] 8/18/2005 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub RC4Initialize(ByVal password As String, ByRef sbox() As Int32, ByRef key() As Int32)
        Dim tempSwap As Int32
        Dim intLength, a, b As Int32
        intLength = Len(password)
        ReDim sbox(255)
        ReDim key(255)
        For a = 0 To 255
            key(a) = Asc(Mid(password, (a Mod intLength) + 1, 1))
            sbox(a) = a
        Next
        b = 0
        For a = 0 To 255
            b = (b + sbox(a) + key(a)) Mod 256
            tempSwap = sbox(a)
            sbox(a) = sbox(b)
            sbox(b) = tempSwap
        Next
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This routine called by EnDeCrypt function. Initializes the
    ''' sbox and the key array)
    ''' </summary>
    ''' <param name="password"></param>
    ''' <param name="sbox"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    '''  [jamie] 8/18/2005 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub RC4Initialize2(ByVal key As Byte(), ByRef sbox As Int32())
        Dim tempSwap As Int32
        Dim i, j, l As Int32
        l = key.Length

        For i = 0 To 255
            sbox(i) = i
        Next
        j = 0
        For i = 0 To 255
            j = (j + sbox(i) + key(i Mod l)) Mod 256
            tempSwap = sbox(i)
            sbox(i) = sbox(j)
            sbox(j) = tempSwap
        Next
    End Sub
End Module
