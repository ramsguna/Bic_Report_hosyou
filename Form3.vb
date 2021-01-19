Public Class Form3
    Dim 受付日 As Integer
    Dim strSQL As String
    Dim MainstrSQL As String
    Dim WK_DATE As Date
    Private objMutex As System.Threading.Mutex
    Dim SqlCmd1 As SqlClient.SqlCommand
    Dim DaList1 = New SqlClient.SqlDataAdapter
    Dim DsExport As New DataSet
    Dim DtView1 As DataView

    Dim waitDlg As WaitDlg   ''進行状況フォームクラス  

    Dim inz_F As String
    Dim r, i As Integer

    Dim year As Integer
    Dim month As Integer
    Dim days As Integer
    Sub XLS_OUT()
        waitDlg.MainMsg = "総合補償で出力をしています。"              ' 進行状況ダイアログのメーターを設定
        waitDlg.ProgressMsg = "データ出力準備中"    ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                      ' メッセージ処理を促して表示を更新する
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        DsExport.Clear()
        'strSQL = "If EXISTS(SELECT * FROM sys.objects "
        'strSQL = strSQL & " WHERE object_id = OBJECT_ID(N'[dbo].[dbo_txt_data_all1]') "
        'strSQL = strSQL & " And type in (N'U')) "
        'strSQL = strSQL & " DROP TABLE [dbo].[dbo_txt_data_all1] "
        MainstrSQL = " SELECT  txt_data_all.WRN_DATE, txt_data_all.WRN_NO, txt_data_all.SHOP_CODE, txt_data_all.ITEM_CODE, txt_data_all.MODEL, "
        MainstrSQL = MainstrSQL & " txt_data_all.CAT_CODE, txt_data_all.CAT_NAME, txt_data_all.MKR_CODE, txt_data_all.MKR_NAME, txt_data_all.PRICE,  "
        MainstrSQL = MainstrSQL & " txt_data_all.WRN_PRICE, txt_data_all.WRN_PRD, txt_data_all.SALE_STS, txt_data_all.CRT_DATE, txt_data_all.CLS_MNTH,"
        MainstrSQL = MainstrSQL & " txt_data_all.PNT_NO, txt_data_all.CUST_NAME, txt_data_all.ZIP1, txt_data_all.ZIP2, txt_data_all.ADRS1, txt_data_all.ADRS2,"
        MainstrSQL = MainstrSQL & " txt_data_all.SEX, txt_data_all.BRTH_DATE, txt_data_all.TEL_NO, txt_data_all.CNT_NO, txt_data_all.IMPT_FILE "
        ' strSQL = strSQL & " INTO dbo_txt_data_all1"
        MainstrSQL = MainstrSQL & " FROM txt_data_all"
        'strSQL = strSQL & " WHERE (((txt_data_all.IMPT_FILE)>='CYOKI.200729'  And (txt_data_all.IMPT_FILE)<='CYOKI.200729'))"
        MainstrSQL = MainstrSQL & " WHERE (((txt_data_all.IMPT_FILE)>='" & テキスト01.Text & "'  And (txt_data_all.IMPT_FILE)<='" & テキスト19.Text & "'))"


        'SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        'DaList1.SelectCommand = SqlCmd1
        'DB_OPEN("bicdb")
        'SqlCmd1.CommandTimeout = 600
        'SqlCmd1.ExecuteNonQuery()
        'DB_CLOSE()

        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='00')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        ' waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "総合補償_txt_data_" & テキスト6.Text & "_19.csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If


        '03
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='03')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV1")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV1"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        ' waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証03_txt_data_" & テキスト6.Text & "_19.csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        '05
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='05')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV2")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV2"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        '  Me.Activate()                   ' いったんオーナーをアクティブにする
        ' waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証05_txt_data_" & テキスト6.Text & "_19.csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        '10
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='10')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV3")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV3"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        '  Me.Activate()                   ' いったんオーナーをアクティブにする
        ' waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証10_txt_data_" & テキスト6.Text & "_19.csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        ' MsgBox("出力しました", , "")
        ' Call Deleteobjects()
    End Sub
    Sub XLS_OUT1()
        waitDlg.MainMsg = "総合補償で出力をしています。"              ' 進行状況ダイアログのメーターを設定
        waitDlg.ProgressMsg = "データ出力準備中"    ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                      ' メッセージ処理を促して表示を更新する
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        DsExport.Clear()
        'strSQL = "If EXISTS(SELECT * FROM sys.objects "
        'strSQL = strSQL & " WHERE object_id = OBJECT_ID(N'[dbo].[dbo_txt_data_all1]') "
        'strSQL = strSQL & " And type in (N'U')) "
        'strSQL = strSQL & " DROP TABLE [dbo].[dbo_txt_data_all1] "
        MainstrSQL = " SELECT txt_data_all.WRN_DATE, txt_data_all.WRN_NO, txt_data_all.SHOP_CODE, txt_data_all.ITEM_CODE, txt_data_all.MODEL, "
        MainstrSQL = MainstrSQL & " txt_data_all.CAT_CODE, txt_data_all.CAT_NAME, txt_data_all.MKR_CODE, txt_data_all.MKR_NAME, txt_data_all.PRICE,  "
        MainstrSQL = MainstrSQL & " txt_data_all.WRN_PRICE, txt_data_all.WRN_PRD, txt_data_all.SALE_STS, txt_data_all.CRT_DATE, txt_data_all.CLS_MNTH,"
        MainstrSQL = MainstrSQL & " txt_data_all.PNT_NO, txt_data_all.CUST_NAME, txt_data_all.ZIP1, txt_data_all.ZIP2, txt_data_all.ADRS1, txt_data_all.ADRS2,"
        MainstrSQL = MainstrSQL & " txt_data_all.SEX, txt_data_all.BRTH_DATE, txt_data_all.TEL_NO, txt_data_all.CNT_NO, txt_data_all.IMPT_FILE "
        ' strSQL = strSQL & " INTO dbo_txt_data_all1"
        MainstrSQL = MainstrSQL & " FROM txt_data_all"
        'strSQL = strSQL & " WHERE (((txt_data_all.IMPT_FILE)>='CYOKI.200729'  And (txt_data_all.IMPT_FILE)<='CYOKI.200729'))"
        MainstrSQL = MainstrSQL & " WHERE (((txt_data_all.IMPT_FILE)>='" & テキスト20.Text & "'  And (txt_data_all.IMPT_FILE)<='" & テキスト30.Text & "'))"


        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN("bicdb")
        SqlCmd1.CommandTimeout = 6000
        SqlCmd1.ExecuteNonQuery()
        DB_CLOSE()

        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='00')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV4")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV4"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        'waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "総合補償_txt_data_" & テキスト10.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If


        '03
        strSQL = MainstrSQL & " And ((([txt_data_all].[WRN_PRD])='03')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV5")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV5"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        'waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証03_txt_data_" & テキスト10.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        '05
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='05')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV6")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV6"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        '  Me.Activate()                   ' いったんオーナーをアクティブにする
        ' waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証05_txt_data_" & テキスト10.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        '10
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='10')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 3600
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV7")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV7"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        'waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証10_txt_data_" & テキスト10.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        ' C
        ' Call Deleteobjects()
    End Sub
    Sub XLS_OUT2()
        waitDlg.MainMsg = "総合補償で出力をしています。"              ' 進行状況ダイアログのメーターを設定
        waitDlg.ProgressMsg = "データ出力準備中"    ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                      ' メッセージ処理を促して表示を更新する
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        DsExport.Clear()
        'strSQL = "If EXISTS(SELECT * FROM sys.objects "
        'strSQL = strSQL & " WHERE object_id = OBJECT_ID(N'[dbo].[dbo_txt_data_all1]') "
        'strSQL = strSQL & " And type in (N'U')) "
        'strSQL = strSQL & " DROP TABLE [dbo].[dbo_txt_data_all1] "
        MainstrSQL = " SELECT  txt_data_all.WRN_DATE, txt_data_all.WRN_NO, txt_data_all.SHOP_CODE, txt_data_all.ITEM_CODE, txt_data_all.MODEL, "
        MainstrSQL = MainstrSQL & " txt_data_all.CAT_CODE, txt_data_all.CAT_NAME, txt_data_all.MKR_CODE, txt_data_all.MKR_NAME, txt_data_all.PRICE,  "
        MainstrSQL = MainstrSQL & " txt_data_all.WRN_PRICE, txt_data_all.WRN_PRD, txt_data_all.SALE_STS, txt_data_all.CRT_DATE, txt_data_all.CLS_MNTH,"
        MainstrSQL = MainstrSQL & " txt_data_all.PNT_NO, txt_data_all.CUST_NAME, txt_data_all.ZIP1, txt_data_all.ZIP2, txt_data_all.ADRS1, txt_data_all.ADRS2,"
        MainstrSQL = MainstrSQL & " txt_data_all.SEX, txt_data_all.BRTH_DATE, txt_data_all.TEL_NO, txt_data_all.CNT_NO, txt_data_all.IMPT_FILE "
        ' strSQL = strSQL & " INTO dbo_txt_data_all1"
        MainstrSQL = MainstrSQL & " FROM txt_data_all"
        'strSQL = strSQL & " WHERE (((txt_data_all.IMPT_FILE)>='CYOKI.200729'  And (txt_data_all.IMPT_FILE)<='CYOKI.200729'))"
        MainstrSQL = MainstrSQL & " WHERE (((txt_data_all.IMPT_FILE)>='" & テキスト01.Text & "'  And (txt_data_all.IMPT_FILE)<='" & テキスト30.Text & "'))"


        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN("bicdb")
        SqlCmd1.CommandTimeout = 6000
        SqlCmd1.ExecuteNonQuery()
        DB_CLOSE()

        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='00')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV8")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV8"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        'waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "総合補償_txt_data_" & テキスト6.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If


        '03
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='03')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV9")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV9"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        'Me.Activate()                   ' いったんオーナーをアクティブにする
        'waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証03_txt_data_" & テキスト6.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        '05
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='05')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV10")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV10"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        '  Me.Activate()                   ' いったんオーナーをアクティブにする
        ' waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証05_txt_data_" & テキスト6.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        '10
        strSQL = MainstrSQL & " and ((([txt_data_all].[WRN_PRD])='10')) ORDER BY txt_data_all.WRN_DATE, txt_data_all.WRN_NO "

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 6000
        DB_OPEN("bicdb")
        r = DaList1.Fill(DsExport, "CSV11")
        DB_CLOSE()

        'If r = 0 Then
        '    MessageBox.Show("該当するデータがありません", "エクスポート", MessageBoxButtons.OK)
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        '    Exit Sub
        'End If

        waitDlg.ProgressMsg = "CSV出力実行中"           ' 進行状況ダイアログのメーターを設定
        Application.DoEvents()                          ' メッセージ処理を促して表示を更新する

        'DtView1 = New DataView(DsExport.Tables("CSV"), "avlbty is Null", "", DataViewRowState.CurrentRows)
        'For i = 0 To DtView1.Count - 1
        '    DtView1(i)("GRP") = "対象外"
        'Next

        'ファイルに出力
        'Dim sw As System.IO.StreamWriter  'StreamWriterオブジェクト
        'Dim sbuf As String                'ファイルに出力するデータ

        sw = New System.IO.StreamWriter(Application.StartupPath & "\temp", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        sbuf = "WRN_DATE,WRN_NO,SHOP_CODE,ITEM_CODE,MODEL,CAT_CODE,CAT_NAME,MKR_CODE,MKR_NAME,PRICE,WRN_PRICE,WRN_PRD,SALE_STS,CRT_DATE,CLS_MNTH,PNT_NO,CUST_NAME,ZIP1,ZIP2,ADRS1,ADRS2,SEX,BRTH_DATE,TEL_NO,CNT_NO,IMPT_FILE"
        sw.WriteLine(sbuf)

        DtView1 = New DataView(DsExport.Tables("CSV11"), "", "", DataViewRowState.CurrentRows)

        waitDlg.ProgressMax = DtView1.Count         ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0                   ' 最初の件数を設定

        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & (i + 1) & "/" & DtView1.Count & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%　"

            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            sbuf = DtView1(i)("WRN_DATE")
            sbuf += "," & DtView1(i)("WRN_NO")
            sbuf += "," & DtView1(i)("SHOP_CODE")
            sbuf += "," & DtView1(i)("ITEM_CODE")
            sbuf += "," & DtView1(i)("MODEL")
            sbuf += "," & DtView1(i)("CAT_CODE")
            sbuf += "," & DtView1(i)("CAT_NAME")
            sbuf += "," & DtView1(i)("MKR_CODE")
            sbuf += "," & DtView1(i)("MKR_NAME")
            sbuf += "," & DtView1(i)("PRICE")
            sbuf += "," & DtView1(i)("WRN_PRICE")
            sbuf += "," & DtView1(i)("WRN_PRD")
            sbuf += "," & DtView1(i)("SALE_STS")
            sbuf += "," & DtView1(i)("CRT_DATE")
            sbuf += "," & DtView1(i)("CLS_MNTH")
            sbuf += "," & DtView1(i)("PNT_NO")
            sbuf += "," & DtView1(i)("CUST_NAME")
            sbuf += "," & DtView1(i)("ZIP1")
            sbuf += "," & DtView1(i)("ZIP2")
            sbuf += "," & DtView1(i)("ADRS1")
            sbuf += "," & DtView1(i)("ADRS2")
            sbuf += "," & DtView1(i)("SEX")
            sbuf += "," & DtView1(i)("BRTH_DATE")
            sbuf += "," & DtView1(i)("TEL_NO")
            sbuf += "," & DtView1(i)("CNT_NO")
            sbuf += "," & DtView1(i)("IMPT_FILE")

            sw.WriteLine(sbuf)
        Next
        sw.Close()

        ' Me.Activate()                   ' いったんオーナーをアクティブにする
        waitDlg.Close()                 ' 進行状況ダイアログを閉じる
        Me.Enabled = False               ' オーナーのフォームを有効にする

        '［名前を付けて保存］ダイアログボックスを表示
        SaveFileDialog1.FileName = "長期保証10_txt_data_" & テキスト6.Text & "_" & days & ".csv"
        SaveFileDialog1.Filter = "CSVファイル|*.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Microsoft.VisualBasic.FileSystem.Kill(Application.StartupPath & "\temp")
        Else
            If System.IO.File.Exists(SaveFileDialog1.FileName) = False And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(SaveFileDialog1.FileName) And System.IO.File.Exists(Application.StartupPath & "\temp") Then
                Microsoft.VisualBasic.FileSystem.Kill(SaveFileDialog1.FileName)
                Microsoft.VisualBasic.FileSystem.Rename(Application.StartupPath & "\temp", SaveFileDialog1.FileName)
            ElseIf System.IO.File.Exists(Application.StartupPath & "\temp") = False Then
                MessageBox.Show("アプリケーションエラー", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        ' MsgBox("出力しました", , "")
        ' Call Deleteobjects()
    End Sub
    'Sub Deleteobjects()
    '    strSQL = "Drop table dbo_txt_data_all1 "
    '    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
    '    DaList1.SelectCommand = SqlCmd1
    '    DB_OPEN("bicdb")
    '    SqlCmd1.CommandTimeout = 600
    '    SqlCmd1.ExecuteNonQuery()
    '    DB_CLOSE()
    'End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        WaitDlg = New WaitDlg        ' 進行状況ダイアログ
        WaitDlg.Owner = Me              ' ダイアログのオーナーを設定する
        waitDlg.MainMsg = Nothing       ' 処理の概要を表示
        waitDlg.ProgressMax = 0         ' 全体の処理件数を設定
        waitDlg.ProgressMin = 0         ' 処理件数の最小値を設定（0件から開始）
        waitDlg.ProgressStep = 1        ' 何件ごとにメータを進めるかを設定
        waitDlg.ProgressValue = 0       ' 最初の件数を設定
        Me.Enabled = False              ' オーナーのフォームを無効にする
        waitDlg.Show()                  ' 進行状況ダイアログを表示する
        Call XLS_OUT()
        Me.Enabled = False
        Call XLS_OUT1()
        Me.Enabled = False
        Call XLS_OUT2()

        MsgBox("出力しました", , "")
        Me.Enabled = True
        Application.Exit()
    End Sub

    Private Sub Button99_Click(sender As Object, e As EventArgs) Handles Button99.Click
        Application.Exit()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call DB_INIT()

        受付日 = Date.Today.Year
        ' 受付日 = Format(DateAdd("m", -1, Date.Today.Year), "yyyy/MM")
        If Date.Today.Month - 1 <> 0 Then
            TextBox2.Text = 受付日 & "/" & Date.Today.Month - 1
        Else
            TextBox2.Text = Date.Today.Year - 1 & "12"
        End If
        WK_DATE = TextBox2.Text & "/01"

        Label1.Text = Format(WK_DATE, "yyyy")
        Label3.Text = Format(WK_DATE, "MM")

        year = Label1.Text
        month = Label3.Text
        days = System.DateTime.DaysInMonth(year, month)

        テキスト6.Text = Format(WK_DATE, "yyyyMMdd")
        テキスト8.Text = Format(DateAdd("d", -1, DateAdd("m", 1, WK_DATE)), "yyyyMMdd")
        テキスト9.Text = Format(WK_DATE, "yyyyMM") & "19"
        テキスト10.Text = Format(WK_DATE, "yyyyMM") & "20"

        テキスト01.Text = "CYOKI." & Format(WK_DATE, "yyMMdd")
        テキスト19.Text = "CYOKI." & Format(WK_DATE, "yyMM") & "19"
        テキスト20.Text = "CYOKI." & Format(WK_DATE, "yyMM") & "20"
        テキスト30.Text = "CYOKI." & Mid(テキスト8.Text, 3, 6)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If TextBox2.Text = Nothing Then
            MSG.Text = "処理年を入力してください。"
            Button5.Enabled = False

        Else
            Try
                MSG.Text = ""
                Button5.Enabled = True
                WK_DATE = TextBox2.Text & "/01"
                Label1.Text = Format(WK_DATE, "yyyy")
                Label3.Text = Format(WK_DATE, "MM")

                year = Label1.Text
                month = Label3.Text
                days = System.DateTime.DaysInMonth(year, month)

                テキスト6.Text = Format(WK_DATE, "yyyyMMdd")
                テキスト8.Text = Format(DateAdd("d", -1, DateAdd("m", 1, WK_DATE)), "yyyyMMdd")
                テキスト9.Text = Format(WK_DATE, "yyyyMM") & "19"
                テキスト10.Text = Format(WK_DATE, "yyyyMM") & "20"

                テキスト01.Text = "CYOKI." & Format(WK_DATE, "yyMMdd")
                テキスト19.Text = "CYOKI." & Format(WK_DATE, "yyMM") & "19"
                テキスト20.Text = "CYOKI." & Format(WK_DATE, "yyMM") & "20"
                テキスト30.Text = "CYOKI." & Mid(テキスト8.Text, 3, 6)
            Catch ex As System.Exception
                MSG.Text = "処理年を入力してください。"
                Button5.Enabled = False
            End Try
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 46 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class