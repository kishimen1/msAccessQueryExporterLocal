' =============================================================
' Access クエリ一括エクスポート (JSON出力版)
' =============================================================
' 使用方法:
'   AccessQueryExporter.bat にファイルをドラッグ＆ドロップ
'   または: cscript ExtractToJSON.vbs "C:\path\to\database.accdb"
' =============================================================
Option Explicit

Dim engine, db, qdef, tdef, rel, fld
Dim fso, outFile, txtFile
Dim dbPath, outputDir, jsonPath, txtPath, baseName
Dim queryCount, tableCount, relCount
Dim jsonStr, txtLines
Dim exportedAt

' ===== 引数チェック =====
If WScript.Arguments.Count = 0 Then
    WScript.Echo "使用方法: AccessQueryExporter.bat にファイルをドラッグ＆ドロップしてください。"
    WScript.Quit 1
End If

dbPath = WScript.Arguments(0)

' ===== ファイル存在チェック =====
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(dbPath) Then
    WScript.Echo "エラー: ファイルが見つかりません: " & dbPath
    WScript.Quit 1
End If

' ===== 出力パスを生成 =====
baseName = fso.GetBaseName(dbPath)
outputDir = fso.GetParentFolderName(WScript.ScriptFullName)
jsonPath = outputDir & "\" & baseName & "_result.json"
txtPath = outputDir & "\" & baseName & "_クエリ一覧.txt"
exportedAt = FormatDateTime(Now(), 0)

' ===== DAO でデータベースを開く =====
On Error Resume Next
Set engine = CreateObject("DAO.DBEngine.120")
If Err.Number <> 0 Then
    Err.Clear
    Set engine = CreateObject("DAO.DBEngine.36")
    If Err.Number <> 0 Then
        WScript.Echo "エラー: DAO.DBEngine が見つかりません。"
        WScript.Echo "Microsoft Access Database Engine をインストールしてください。"
        WScript.Quit 1
    End If
End If

Set db = engine.OpenDatabase(dbPath)
If Err.Number <> 0 Then
    WScript.Echo "エラー: データベースを開けませんでした。"
    WScript.Echo "パス: " & dbPath
    WScript.Echo Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

' ===== JSON文字列エスケープ関数 =====
Function JsonEscape(s)
    Dim result
    result = s
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    JsonEscape = result
End Function

' ===== クエリタイプラベル関数 =====
Function GetQueryTypeLabel(qtype)
    Select Case qtype
        Case 0: GetQueryTypeLabel = "選択"
        Case 1: GetQueryTypeLabel = "クロス集計"
        Case 2: GetQueryTypeLabel = "削除"
        Case 3: GetQueryTypeLabel = "更新"
        Case 4: GetQueryTypeLabel = "追加"
        Case 5: GetQueryTypeLabel = "テーブル作成"
        Case 6: GetQueryTypeLabel = "データ定義"
        Case 7: GetQueryTypeLabel = "パススルー"
        Case 8: GetQueryTypeLabel = "ユニオン"
        Case 9: GetQueryTypeLabel = "サブフォーム/サブレポート"
        Case Else: GetQueryTypeLabel = "その他(" & qtype & ")"
    End Select
End Function

' ===== テキスト出力用の行バッファ =====
txtLines = ""

Sub AddTxtLine(line)
    txtLines = txtLines & line & vbCrLf
End Sub

' ===== テキストヘッダー =====
AddTxtLine "=============================================="
AddTxtLine "データベース: " & fso.GetFileName(dbPath)
AddTxtLine "出力日時: " & exportedAt
AddTxtLine "=============================================="
AddTxtLine ""

' ===== JSON構築開始 =====
jsonStr = "{" & vbCrLf
jsonStr = jsonStr & "  ""filename"": """ & JsonEscape(fso.GetFileName(dbPath)) & """," & vbCrLf
jsonStr = jsonStr & "  ""exported_at"": """ & JsonEscape(CStr(exportedAt)) & """," & vbCrLf

' ----- テーブル一覧 -----
jsonStr = jsonStr & "  ""tables"": [" & vbCrLf
AddTxtLine "■■■ テーブル一覧 ■■■"
AddTxtLine ""

tableCount = 0
For Each tdef In db.TableDefs
    If Left(tdef.Name, 4) <> "MSys" And Left(tdef.Name, 1) <> "~" Then
        If tableCount > 0 Then
            jsonStr = jsonStr & "," & vbCrLf
        End If
        jsonStr = jsonStr & "    {""name"": """ & JsonEscape(tdef.Name) & """, ""fields"": ["

        Dim fieldCount
        fieldCount = 0
        Dim fieldList
        fieldList = ""
        For Each fld In tdef.Fields
            If fieldCount > 0 Then
                jsonStr = jsonStr & ", "
                fieldList = fieldList & ", "
            End If
            jsonStr = jsonStr & """" & JsonEscape(fld.Name) & """"
            fieldList = fieldList & fld.Name
            fieldCount = fieldCount + 1
        Next
        jsonStr = jsonStr & "]}"

        AddTxtLine "  ・" & tdef.Name
        AddTxtLine "    フィールド: " & fieldList
        tableCount = tableCount + 1
    End If
Next
jsonStr = jsonStr & vbCrLf & "  ]," & vbCrLf
AddTxtLine ""

' ----- リレーションシップ -----
jsonStr = jsonStr & "  ""relationships"": [" & vbCrLf
AddTxtLine "■■■ リレーションシップ ■■■"
AddTxtLine ""

relCount = 0
For Each rel In db.Relations
    If relCount > 0 Then
        jsonStr = jsonStr & "," & vbCrLf
    End If
    jsonStr = jsonStr & "    {""table"": """ & JsonEscape(rel.Table) & """, ""foreign_table"": """ & JsonEscape(rel.ForeignTable) & """, ""fields"": ["

    Dim relFieldCount
    relFieldCount = 0
    Dim relFieldInfo
    relFieldInfo = ""
    Dim f
    For Each f In rel.Fields
        If relFieldCount > 0 Then
            jsonStr = jsonStr & ", "
            relFieldInfo = relFieldInfo & ", "
        End If
        jsonStr = jsonStr & "{""from_field"": """ & JsonEscape(f.Name) & """, ""to_field"": """ & JsonEscape(f.ForeignName) & """}"
        relFieldInfo = relFieldInfo & f.Name & "=" & f.ForeignName
        relFieldCount = relFieldCount + 1
    Next
    jsonStr = jsonStr & "]}"

    AddTxtLine "  " & rel.Table & " → " & rel.ForeignTable & "  (" & relFieldInfo & ")"
    relCount = relCount + 1
Next
jsonStr = jsonStr & vbCrLf & "  ]," & vbCrLf
AddTxtLine ""

' ----- クエリ一覧とSQL -----
jsonStr = jsonStr & "  ""queries"": [" & vbCrLf
AddTxtLine "■■■ クエリ一覧とSQL定義 ■■■"
AddTxtLine ""

queryCount = 0
For Each qdef In db.QueryDefs
    If Left(qdef.Name, 1) <> "~" Then
        If queryCount > 0 Then
            jsonStr = jsonStr & "," & vbCrLf
        End If
        Dim qTypeLabel
        qTypeLabel = GetQueryTypeLabel(qdef.Type)
        jsonStr = jsonStr & "    {""name"": """ & JsonEscape(qdef.Name) & """, ""sql"": """ & JsonEscape(qdef.SQL) & """, ""type"": """ & JsonEscape(qTypeLabel) & """}"

        queryCount = queryCount + 1
        AddTxtLine "【" & queryCount & "】 " & qdef.Name & "  [" & qTypeLabel & "]"
        AddTxtLine String(50, "-")
        AddTxtLine qdef.SQL
        AddTxtLine ""
        AddTxtLine ""
    End If
Next
jsonStr = jsonStr & vbCrLf & "  ]" & vbCrLf
jsonStr = jsonStr & "}"

AddTxtLine "=============================================="
AddTxtLine "総クエリ数: " & queryCount
AddTxtLine "=============================================="

db.Close

' ===== JSON ファイル出力 (UTF-8) =====
Dim adoStream
Set adoStream = CreateObject("ADODB.Stream")
adoStream.Type = 2 ' adTypeText
adoStream.Charset = "UTF-8"
adoStream.Open
adoStream.WriteText jsonStr
adoStream.SaveToFile jsonPath, 2 ' adSaveCreateOverWrite
adoStream.Close
Set adoStream = Nothing

' ===== テキストファイル出力 (UTF-8) =====
Set adoStream = CreateObject("ADODB.Stream")
adoStream.Type = 2
adoStream.Charset = "UTF-8"
adoStream.Open
adoStream.WriteText txtLines
adoStream.SaveToFile txtPath, 2
adoStream.Close
Set adoStream = Nothing

' ===== 完了メッセージ =====
WScript.Echo "出力完了！"
WScript.Echo "テーブル数: " & tableCount
WScript.Echo "リレーション数: " & relCount
WScript.Echo "クエリ数: " & queryCount
WScript.Echo ""
WScript.Echo "JSON出力: " & jsonPath
WScript.Echo "テキスト出力: " & txtPath
