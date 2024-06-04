Attribute VB_Name = "Module1"

' フィルタリングされたデータを取得する関数
' ws: 検索対象のシート
' filterWork: 検索ワード
Function GetFilteredData(ws As Worksheet, filterWork As String) As Range
    ' データが存在する範囲を自動的に検出する
    Dim lastRow As Long

    ' 今回は「F」列に「金融機関」種別が入っているのでその部分に対して検索する。
    ' 「金融機関」の位置がFからずれた場合、ここを変更する
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' F列のデータをフィルタリングする
    ' F列の内容がfilterWorkと同じ行を取得する
    ws.Range("$A$1:$F$" & lastRow).AutoFilter Field:=6, Criteria1:=filterWork

    ' フィルタリングされたデータを取得する
    ' Offsetメソッドで1行ずらすことで、ヘッダー行を除外する
    ' エラーハンドリングを追加して、フィルタリング結果がヘッダー行のみの場合を考慮する
    On Error Resume Next
    Set GetFilteredData = ws.Range("$A$2:$F$" & lastRow).SpecialCells(xlCellTypeVisible)
    If Err.Number <> 0 Then
        Set GetFilteredData = Nothing
    End If
    On Error GoTo 0
End Function

' 金融機関名, 学校名 を受け取り一致する行の一覧を返す
Function GetFilteredDataFromFinancialInstitution(fiName As String, sheetName As String) As Range
    ' シート "sheetName" のタブを開く
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' フィルタリングされたデータを取得する
    Set GetFilteredDataFromFinancialInstitution = GetFilteredData(ws, fiName)

    ' フィルタリングを解除する
    ws.AutoFilterMode = False
End Function

' 結果の件数を取得する
Function GetRowCount(rngData As Range) As Long
    Dim rowCnt As Long: rowCnt = 0
    Dim rng As Range
    For Each rng In rngData.Areas
        rowCnt = rowCnt + rng.EntireRow.Count
    Next
    GetRowCount = rowCnt
End Function

' 数字を取得する
Function GetAmount(target As String, ws As Worksheet) As Long
    GetAmount = ws.Range(target).Value
End Function


' ===========================
'  ここから東邦銀行関係の処理
' ===========================

' テンプレートにデータを追加する関数
Sub AppendTohoDataToTemplate(filteredData As Range,lastRowTemplate As Long, amount As String, teachAmount As String, transferDate As String, wsTemplate As Worksheet, branchDict As Dictionary)
    ' フィルタリングされたデータをテンプレートに追記する
    Dim cell As Range
    For Each cell In filteredData.Rows
        wsTemplate.Cells(lastRowTemplate, "D").Value = cell.Cells(1, "G").Value ' 口座名義(漢字)
        wsTemplate.Cells(lastRowTemplate, "E").Value = cell.Cells(1, "H").Value ' 口座名義(カナ)
        wsTemplate.Cells(lastRowTemplate, "G").Value = "東邦銀行"
        wsTemplate.Cells(lastRowTemplate, "H").Value = cell.Cells(1, "I").Value ' 支店名(漢字)
        wsTemplate.Cells(lastRowTemplate, "I").Value = branchDict(Replace(cell.Cells(1, "I").Value, "支店", "")) ' 支店名から支店番号を取得(支店名に'支店'の文字があれば削除)        
        wsTemplate.Cells(lastRowTemplate, "J").Value = "普通"
        wsTemplate.Cells(lastRowTemplate, "K").Value = cell.Cells(1, "J").Value '
        
        ' 給食費: 教師の場合 (学年が 7 ) は教師向けの金額を入れる
        If cell.Cells(1, "B").Value = "7"  Then
            wsTemplate.Cells(lastRowTemplate, "L").Value = teachAmount
        Else 
            wsTemplate.Cells(lastRowTemplate, "L").Value = amount
        End If

        wsTemplate.Cells(lastRowTemplate, "M").Value = transferDate ' 振替日(東邦銀行)
        wsTemplate.Cells(lastRowTemplate, "N").Value = cell.Cells(1, "K").Value ' 住所

        lastRowTemplate = lastRowTemplate + 1
    Next cell
End Sub

' 東邦銀行の支店名と支店番号を紐つけるデータを読み込み
Function CreateToHoBranchDictionary() As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("東邦銀行_支店情報") ' 支店情報の書かれているシート

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastRow
        Dim key As String
        key = ws.Cells(i, 1)

        ' 値は支店番号
        Dim value As String
        value = ws.Cells(i, 3)

        ' データに追加
        dict.Add key, value
    Next i

    ' 紐つけデータを返す
    Set CreateToHoBranchDictionary = dict

End Function

Sub ExecuteToho()


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("マクロ")

    ' 自動の金額
    Dim elementaryAmount As String ' 小学
    Dim juniorHighAmount As String ' 中学
    ' 教師等の金額 
    Dim elementaryTeachAmount As String ' 小学
    Dim juniorHighTeachAmount As String ' 中学

    elementaryAmount = GetAmount("C9", ws)
    juniorHighAmount = GetAmount("C10", ws)

    elementaryTeachAmount = GetAmount("C11", ws)
    juniorHighTeachAmount = GetAmount("C12", ws)

    ' 振替日
    Dim transferDate As String 
    transferDate = GetAmount("C15", ws)


    ' テンプレートファイルを開く
    Dim wbTemplate As Workbook
    Set wbTemplate = Workbooks.Open(ThisWorkbook.Path & "\templates\" &  "toho.xlsx") 'テンプレートのパスを指定してください。

    ' テンプレートの最初のシートを取得する
    Dim wsTemplate As Worksheet
    Set wsTemplate = wbTemplate.Sheets(1)

    ' branchDictを初期化します
    Dim branchDict As Dictionary
    Set branchDict = CreateToHoBranchDictionary()
    
    ' データを追記する行をして石います。4からなのは1-3がヘッダーだからです。
    Dim offset As Long: offset = 4

    ' 笈川
    Dim rngOikawa As Range
    Set rngOikawa = GetFilteredDataFromFinancialInstitution("東邦", "笈川")
    AppendTohoDataToTemplate rngOikawa, offset, elementaryAmount, elementaryTeachAmount, transferDate, wsTemplate, branchDict

    ' 勝常
    Dim rngShojo As Range
    Set rngShojo = GetFilteredDataFromFinancialInstitution("東邦", "勝常")
    AppendTohoDataToTemplate rngShojo, offset, elementaryAmount, elementaryTeachAmount, transferDate, wsTemplate, branchDict

    ' 湯川中
    Dim rngYugawa As Range
    Set rngYugawa = GetFilteredDataFromFinancialInstitution("東邦", "湯川中")
    AppendTohoDataToTemplate rngYugawa, offset, juniorHighAmount, juniorHighTeachAmount, transferDate, wsTemplate, branchDict

    ' テンプレートを保存する
    wbTemplate.SaveAs fileName := ThisWorkbook.Path & "\result\" & "toho.xlsx"
    wbTemplate.Close savechanges := False
End Sub

' ===========================
'  ここまで東邦銀行関係の処理
' ===========================


' ===========================
'  ここからJAよつば関係の処理
' ===========================

' テンプレートにデータを追加する関数
Sub AppendJaDataToTemplate(filteredData As Range, inputDescription As String, lastRowTemplate As Long, amount As String, teachAmount As String, wsTemplate As Worksheet, branchDict As Dictionary)
    ' フィルタリングされたデータをテンプレートに追記する
    Dim cell As Range
    For Each cell In filteredData.Rows
        wsTemplate.Cells(lastRowTemplate, "A").Value = branchDict(Replace(cell.Cells(1, "I").Value, "支店", "")) ' 支店名から支店番号を取得(支店名に'支店'の文字があれば削除)  
        wsTemplate.Cells(lastRowTemplate, "B").Value = cell.Cells(1, "J").Value '
        wsTemplate.Cells(lastRowTemplate, "C").Value = cell.Cells(1, "H").Value '

        ' 教師の場合 (学年が 7 ) は教師向けの金額を入れる
        If cell.Cells(1, "B").Value = "7"  Then
            wsTemplate.Cells(lastRowTemplate, "D").Value = teachAmount
        Else 
            wsTemplate.Cells(lastRowTemplate, "D").Value = amount
        End If

        wsTemplate.Cells(lastRowTemplate, "E").Value = inputDescription
        lastRowTemplate = lastRowTemplate + 1
    Next cell
End Sub



' JAよつばの支店名と支店番号を紐つけるデータを読み込み
Function CreateJABranchDictionary() As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("JAよつば_支店情報") ' 支店情報の書かれているシート

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastRow
        Dim key As String
        key = ws.Cells(i, 1)

        ' 値は支店番号
        Dim value As String
        value = ws.Cells(i, 2) '２列目

        ' データに追加
        dict.Add key, value
    Next i

    ' 紐つけデータを返す
    Set CreateJABranchDictionary = dict
End Function

Sub ExecuteJA()
    Dim ws As Worksheet
    Dim inputDescription As String
     
    ' ｷｭｳｼｮｸﾋ5ｶﾞﾂﾌﾞﾝなどの通帳コメントを入力
    inputDescription = InputBox("通帳のコメントを入力してください:", "コメントの入力")

    ' もし入力が空の場合、処理を終了する
    If inputDescription = "" Then
        MsgBox "入力されていないため、処置はキャンセルされました。", vbInformation, "キャンセル"
        Exit Sub
    End If

    ' 確認ダイアログを表示
    msgResponse = MsgBox("以下の内容で生成しますか？" & vbNewLine & inputDescription, vbYesNo + vbQuestion, "確認")

    Set ws = ThisWorkbook.Sheets("マクロ")
    ' 自動の金額
    Dim elementaryAmount As String ' 小学
    Dim juniorHighAmount As String ' 中学
    ' 教師等の金額 
    Dim elementaryTeachAmount As String ' 小学
    Dim juniorHighTeachAmount As String ' 中学

    elementaryAmount = GetAmount("C9", ws)
    juniorHighAmount = GetAmount("C10", ws)

    elementaryTeachAmount = GetAmount("C11", ws)
    juniorHighTeachAmount = GetAmount("C12", ws)

    ' もし「いいえ」が選択された場合
    If msgResponse = vbNo Then
        Exit Sub
    End If

    ' テンプレートファイルを開く
    Dim wbTemplate As Workbook
    Set wbTemplate = Workbooks.Open(ThisWorkbook.Path & "\templates\" &  "ja.xlsx") 'テンプレートのパスを指定してください。

    ' テンプレートの最初のシートを取得する
    Dim wsTemplate As Worksheet
    Set wsTemplate = wbTemplate.Sheets(1)

    ' 口座番号紐つけデータ
    Dim branchDict As Dictionary
    Set branchDict = CreateJABranchDictionary()
    
    ' データを追記する行をして石います。2からなのは1-3がヘッダーだからです。
    Dim offset As Long: offset = 2

    ' 笈川
    Dim rngOikawa As Range
    Set rngOikawa = GetFilteredDataFromFinancialInstitution("JA", "笈川")
    If Not rngOikawa Is Nothing Then
        AppendJaDataToTemplate rngOikawa, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If
    Set rngOikawa = GetFilteredDataFromFinancialInstitution("ＪＡ会津よつば", "笈川")
    If Not rngOikawa Is Nothing Then
        AppendJaDataToTemplate rngOikawa, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If

    ' 勝常
    Dim rngShojo As Range
    Set rngShojo = GetFilteredDataFromFinancialInstitution("JA", "勝常")
    If Not rngShojo Is Nothing Then
        AppendJaDataToTemplate rngShojo, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If
    Set rngShojo = GetFilteredDataFromFinancialInstitution("ＪＡ会津よつば", "勝常")
    If Not rngShojo Is Nothing Then
        AppendJaDataToTemplate rngShojo, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If

    ' 湯川中
    Dim rngYugawa As Range
    Set rngYugawa = GetFilteredDataFromFinancialInstitution("JA", "湯川中")
    If Not rngYugawa Is Nothing Then
        AppendJaDataToTemplate rngYugawa, inputDescription, offset, juniorHighAmount, juniorHighTeachAmount, wsTemplate, branchDict
    End If
    Set rngYugawa = GetFilteredDataFromFinancialInstitution("ＪＡ会津よつば", "湯川中")
    If Not rngYugawa Is Nothing Then
        AppendJaDataToTemplate rngYugawa, inputDescription, offset, juniorHighAmount, juniorHighTeachAmount, wsTemplate, branchDict
    End If

    ' テンプレートを保存する
    wbTemplate.SaveAs fileName := ThisWorkbook.Path & "\result\" & "ja.xlsx"
    wbTemplate.Close savechanges := False
End Sub




' ===========================
'  ここから学年移行の処理
' ===========================

Sub ExecuteMigration()
    Dim inputNumber As String
    ' 新入生の人数を入力
    inputNumber = InputBox("新入生の人数を入力してください:", "新入生の人数")

    ' もし入力が空の場合、処理を終了する
    If inputNumber = "" Then
        MsgBox "入力されていないため、処置はキャンセルされました。", vbInformation, "キャンセル"
        Exit Sub
    End If

    ' 確認ダイアログを表示
    msgResponse = MsgBox("以下の内容で学年を更新しますか？" & vbNewLine & inputNumber, vbYesNo + vbQuestion, "確認")
    
    ' もし「いいえ」が選択された場合
    If msgResponse = vbNo Then
        Exit Sub
    End If

    Dim i As Long

    'シートを設定
    Dim wsOikawa As Worksheet, wsShojo As Worksheet, wsYugawa As Worksheet
    Set wsOikawa = ThisWorkbook.Sheets("笈川")
    Set wsShojo = ThisWorkbook.Sheets("勝常")
    Set wsYugawa = ThisWorkbook.Sheets("湯川中")

    '最後の行を取得
    Dim LastRowOikawa As Long
    Dim LastRowShojo As Long
    Dim LastRowYugawa As Long
    LastRowOikawa = wsOikawa.Cells(wsOikawa.Rows.Count, "A").End(xlUp).Row
    LastRowShojo = wsShojo.Cells(wsShojo.Rows.Count, "A").End(xlUp).Row
    LastRowYugawa = wsYugawa.Cells(wsYugawa.Rows.Count, "A").End(xlUp).Row

    ' ===========================
    '  ここから湯川中の処理
    ' ===========================

    '湯川中、学年が3の行を削除
    For i = LastRowYugawa To 2 Step -1 ' 2行目から開始
        If wsYugawa.Cells(i, 2).Value = 3 Then
            wsYugawa.Rows(i).Delete
        End If
    Next i

    '湯川中、学年の更新
    For i = 2 To LastRowYugawa
        If wsYugawa.Cells(i, 2).Value = 1 Then
            wsYugawa.Cells(i, 2).Value = 2
        ElseIf wsYugawa.Cells(i, 2).Value = 2 Then
            wsYugawa.Cells(i, 2).Value = 3
        End If
    Next i

    ' ===========================
    '  ここから笈川小の処理
    ' ===========================

    '笈川小から学年が6の行を湯川中シートの先頭にコピー
    For i = LastRowOikawa To 2 Step -1 ' 2行目から開始
        If wsOikawa.Cells(i, 2).Value = 6 Then ' 学年が６のもの
            'コピーする範囲を設定
            Set rngToCopy = wsOikawa.Rows(i)
            '挿入する位置を設定
            Set rngDest = wsYugawa.Rows(2)
            '行を挿入
            rngDest.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            'データをコピー
            rngToCopy.Copy Destination:=wsYugawa.Rows(2)
            '笈川シートから行を削除
            wsOikawa.Rows(i).Delete
        End If
    Next i

    '笈川小、学年の更新
    For i = 2 To LastRowOikawa
        If wsOikawa.Cells(i, 2).Value = 1 Then
            wsOikawa.Cells(i, 2).Value = 2
        ElseIf wsOikawa.Cells(i, 2).Value = 2 Then
            wsOikawa.Cells(i, 2).Value = 3
        ElseIf wsOikawa.Cells(i, 2).Value = 3 Then
            wsOikawa.Cells(i, 2).Value = 4
        ElseIf wsOikawa.Cells(i, 2).Value = 4 Then
            wsOikawa.Cells(i, 2).Value = 5
        ElseIf wsOikawa.Cells(i, 2).Value = 5 Then
            wsOikawa.Cells(i, 2).Value = 6
        End If
    Next i

    '先頭に20行の空白を追加
    wsOikawa.Rows("2:" & inputNumber).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    ' ===========================
    '  ここから勝常小の処理
    ' ===========================

    '勝常小から学年が6の行を湯川中シートの先頭にコピー
    For i = LastRowShojo To 2 Step -1 ' 2行目から開始
        If wsShojo.Cells(i, 2).Value = 6 Then ' 学年が６のもの
            'コピーする範囲を設定
            Set rngToCopy = wsShojo.Rows(i)
            '挿入する位置を設定
            Set rngDest = wsYugawa.Rows(2)
            '行を挿入
            rngDest.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            'データをコピー
            rngToCopy.Copy Destination:=wsYugawa.Rows(2)
            '笈川シートから行を削除
            wsShojo.Rows(i).Delete
        End If
    Next i

    '勝常小、学年の更新
    For i = 2 To LastRowShojo
        If wsShojo.Cells(i, 2).Value = 1 Then
            wsShojo.Cells(i, 2).Value = 2
        ElseIf wsShojo.Cells(i, 2).Value = 2 Then
            wsShojo.Cells(i, 2).Value = 3
        ElseIf wsShojo.Cells(i, 2).Value = 3 Then
            wsShojo.Cells(i, 2).Value = 4
        ElseIf wsShojo.Cells(i, 2).Value = 4 Then
            wsShojo.Cells(i, 2).Value = 5
        ElseIf wsShojo.Cells(i, 2).Value = 5 Then
            wsShojo.Cells(i, 2).Value = 6
        End If
    Next i

    '湯川中、6年生を1年生に
    For i = 2 To LastRowYugawa
        If wsYugawa.Cells(i, 2).Value = 6 Then
            wsYugawa.Cells(i, 2).Value = 1
        End If
    Next i


    '先頭に20行の空白を追加
    wsShojo.Rows("2:" & inputNumber).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

   '別名で保存
    ThisWorkbook.SaveAs fileName := ThisWorkbook.Path & "\new_" & ThisWorkbook.Name

End Sub