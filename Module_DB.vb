Module Module_DB
    Public cnsqlclient As New System.Data.SqlClient.SqlConnection
    Public source(4) As String
    Public catalog(4) As String

    Public Function DB_INIT()
        Call DB_INIT_sub(1)
        'Call DB_INIT_sub(2)
        'Call DB_INIT_sub(3)
        'Call DB_INIT_sub(4)
    End Function

    Sub DB_INIT_sub(ByVal seq)
        Dim sr As System.io.StreamReader
        Select Case seq
            Case Is = 1
                sr = New System.IO.StreamReader("bicdb.ini")
            Case Is = 2
                sr = New System.IO.StreamReader("best_wrn_data2.ini")
            Case Is = 3
                sr = New System.IO.StreamReader("best_intl.ini")
            Case Is = 4
                sr = New System.IO.StreamReader("best_wrn_temp.ini")
        End Select

        Dim line As String
        Dim line_len As Integer
        Dim eq_pos As Integer
        Dim line_key As String
        Do
            line = sr.ReadLine()
            line_len = Len(line)
            If line_len = 0 Then
            Else
                eq_pos = InStr(1, line, "=", 1)
                line_key = Mid(line, 1, eq_pos - 1)
                If line_key = "source" Then
                    source(seq) = Mid(line, eq_pos + 1, line_len - eq_pos)
                End If
                If line_key = "catalog" Then
                    catalog(seq) = Mid(line, eq_pos + 1, line_len - eq_pos)
                End If
            End If
        Loop Until line Is Nothing
        sr.Close()
    End Sub

    Public Function DB_OPEN(ByVal file) As Boolean
        DB_OPEN = False

        '******  接続文字列を作成して接続を開始する  *****
        Select Case file
            Case Is = "bicdb"
                cnsqlclient.ConnectionString = "integrated security=SSPI;data source=" & source(1) & ";" &
                                               "persist security info=False;initial catalog=" & catalog(1)
            Case Is = "best_wrn_data2"
                cnsqlclient.ConnectionString = "integrated security=SSPI;data source=" & source(2) & ";" &
                                               "persist security info=False;initial catalog=" & catalog(2)
            Case Is = "best_intl"
                cnsqlclient.ConnectionString = "integrated security=SSPI;data source=" & source(3) & ";" &
                                               "persist security info=False;initial catalog=" & catalog(3)
            Case Is = "best_wrn_temp"
                cnsqlclient.ConnectionString = "integrated security=SSPI;data source=" & source(4) & ";" &
                                               "persist security info=False;initial catalog=" & catalog(4)
        End Select


        Try
            '*****  Connectionが接続されているかチェック  *****
            If cnsqlclient.State = ConnectionState.Closed Then
                cnsqlclient.Open()
            End If
        Catch
            MsgBox(Err.Description, 16, "接続エラー")
            DB_OPEN = False
            Exit Function
        End Try

        DB_OPEN = True
    End Function

    Public Sub DB_CLOSE()
        '接続を終了する
        cnsqlclient.Close()
    End Sub
End Module
