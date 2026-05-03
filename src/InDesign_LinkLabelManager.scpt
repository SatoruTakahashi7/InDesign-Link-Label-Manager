(*
	InDesign_LinkLabelManager.scpt	
	または
    InDesignの配置に色付け.scpt
    Version: 1.0.0
    Updated: 2026-05-01
    GYAHTEI Design Laboratory
    @gyahtei_satoru
    ChatGPT と共同開発

    ----------------------------------------

    InDesign 配置リンク Finderラベル管理ツール

    ----------------------------------------

    機能:
    - 配置リンク元ファイルに Finder ラベル付与
    - AI(.ai) はオレンジ
    - その他画像は赤
    - 使用中リンクのラベル解除
    - ドキュメント保存先フォルダ配下のみ対象
    - 同一元ファイルは1回だけ処理
    - 埋め込み・リンク切れは安全にスキップ

    対応:
    - Adobe InDesign 2022〜2026
    - macOS Finderラベル

    注意:
    - 必ず複製データで検証してください
    - Finderラベルはファイル単位です
    - 複数ドキュメントで共通素材を使用している場合、
      他ドキュメント側のラベル状態にも影響します

    用途:
    - 未使用画像整理
    - Linksフォルダ整理
    - Illustratorリンク確認補助
    - 案件納品前整理

*)


set dlg1 to display dialog "実行する処理を選択してください" buttons {"キャンセル", "ラベルを消す", "ラベルを付ける"} default button "ラベルを付ける" with title "InDesign 配置リンクラベル"

set actionButton to button returned of dlg1

if actionButton is "キャンセル" then return

if actionButton is "ラベルを付ける" then
	set modeName to "set"
else
	set modeName to "clear"
end if


tell application id "com.adobe.InDesign"
	activate
	
	if (count of documents) is 0 then
		my showFrontDialog("InDesignでドキュメントが開かれていません。")
		return
	end if
	
	tell document 1
		
		if saved is false then
			my showFrontDialog("ドキュメントが未保存です。" & return & "保存してから実行してください。")
			return
		end if
		
		set docFilePath to file path
		set docFolderPOSIX to my parentFolderPOSIX(docFilePath)
		
		set mygraphics to all graphics
		set imgcount to count mygraphics
		
		if imgcount is 0 then
			my showFrontDialog("配置画像が見つかりませんでした。")
			return
		end if
		
		set processedPaths to {}
		
		set okCount to 0
		set aiCount to 0
		set normalCount to 0
		set skipCount to 0
		set ngCount to 0
		set outsideCount to 0
		
		set errList to {}
		
		repeat with i from 1 to imgcount
			
			try
				
				set g to item i of mygraphics
				
				-- リンク取得
				try
					set theLink to item link of g
				on error
					set skipCount to skipCount + 1
					set end of errList to ("[" & i & "] リンクを取得できないためスキップ（埋め込み等）")
					error number -128
				end try
				
				-- file path取得
				try
					set theFile to file path of theLink
				on error
					set skipCount to skipCount + 1
					set end of errList to ("[" & i & "] file pathを取得できないためスキップ")
					error number -128
				end try
				
				if theFile is missing value then
					set skipCount to skipCount + 1
					set end of errList to ("[" & i & "] file pathが存在しないためスキップ")
					error number -128
				end if
				
				set pathKey to my normalizePathKey(theFile)
				
				-- ドキュメント保存先フォルダ配下だけ対象
				if my isUnderFolder(pathKey, docFolderPOSIX) is false then
					set outsideCount to outsideCount + 1
					error number -129
				end if
				
				-- 同じ元ファイルは1回だけ
				if processedPaths contains pathKey then
					-- 処理済み
				else
					
					set ext to my getExtension(pathKey)
					
					if modeName is "set" then
						
						if ext is "ai" then
							my setFinderLabel(theFile, 1) -- オレンジ
							set aiCount to aiCount + 1
						else
							my setFinderLabel(theFile, 2) -- 赤
							set normalCount to normalCount + 1
						end if
						
					else
						
						my setFinderLabel(theFile, 0) -- ラベルなし
						
						if ext is "ai" then
							set aiCount to aiCount + 1
						else
							set normalCount to normalCount + 1
						end if
						
					end if
					
					set end of processedPaths to pathKey
					set okCount to okCount + 1
					
				end if
				
			on error errMsg number errNum
				
				if errNum is -128 then
					-- スキップ済み
					
				else if errNum is -129 then
					-- 保存先フォルダ外
					
				else
					set ngCount to ngCount + 1
					set end of errList to ("[" & i & "] " & errMsg & " (" & errNum & ")")
				end if
				
			end try
			
		end repeat
		
	end tell
end tell


if modeName is "set" then
	set titleText to "ラベル付与 完了"
	set detailText to "AI＝オレンジ / その他＝赤"
else
	set titleText to "ラベル解除 完了"
	set detailText to "使用中ファイルのラベルを解除"
end if

set scopeText to "対象範囲: ドキュメント保存先フォルダ配下のみ"

my showResult(titleText, detailText, scopeText, docFolderPOSIX, okCount, aiCount, normalCount, skipCount, outsideCount, ngCount, errList)


-- =========================================
-- Finderラベル設定
-- =========================================

on setFinderLabel(theFile, labelnum)
	
	set targetAlias to my pathToAlias(theFile)
	
	tell application "Finder"
		set label index of targetAlias to labelnum
	end tell
	
end setFinderLabel


-- =========================================
-- alias変換
-- =========================================

on pathToAlias(v)
	
	if class of v is alias then
		
		return v
		
	else if class of v is string then
		
		if v starts with "/" then
			return (POSIX file v) as alias
		else
			return v as alias
		end if
		
	else
		
		try
			return v as alias
		on error
			error "aliasに変換できないパス形式です。"
		end try
		
	end if
	
end pathToAlias


-- =========================================
-- パス正規化
-- =========================================

on normalizePathKey(v)
	
	if class of v is alias then
		
		return POSIX path of v
		
	else if class of v is string then
		
		if v starts with "/" then
			return v
		else
			try
				return POSIX path of (v as alias)
			on error
				return v
			end try
		end if
		
	else
		
		try
			return POSIX path of (v as alias)
		on error
			return v as string
		end try
		
	end if
	
end normalizePathKey


-- =========================================
-- 親フォルダ取得
-- =========================================

on parentFolderPOSIX(v)
	
	set p to my normalizePathKey(v)
	
	if p ends with "/" then
		return p
	end if
	
	set oldTID to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "/"
	set parts to text items of p
	set AppleScript's text item delimiters to oldTID
	
	if (count of parts) < 2 then return p
	
	set folderPath to ""
	
	repeat with i from 1 to ((count of parts) - 1)
		
		set partText to item i of parts
		
		if i is 1 then
			set folderPath to folderPath & partText
		else
			set folderPath to folderPath & "/" & partText
		end if
		
	end repeat
	
	if folderPath does not start with "/" then
		set folderPath to "/" & folderPath
	end if
	
	if folderPath does not end with "/" then
		set folderPath to folderPath & "/"
	end if
	
	return folderPath
	
end parentFolderPOSIX


-- =========================================
-- 指定フォルダ配下判定
-- =========================================

on isUnderFolder(filePathPOSIX, folderPathPOSIX)
	
	if folderPathPOSIX does not end with "/" then
		set folderPathPOSIX to folderPathPOSIX & "/"
	end if
	
	if filePathPOSIX starts with folderPathPOSIX then
		return true
	else
		return false
	end if
	
end isUnderFolder


-- =========================================
-- 拡張子取得
-- =========================================

on getExtension(p)
	
	set oldTID to AppleScript's text item delimiters
	
	set AppleScript's text item delimiters to "."
	set pItems to text items of p
	set AppleScript's text item delimiters to oldTID
	
	if (count of pItems) < 2 then return ""
	
	set ext to item -1 of pItems
	
	try
		return do shell script "printf %s " & quoted form of ext & " | tr '[:upper:]' '[:lower:]'"
	on error
		return ext
	end try
	
end getExtension


-- =========================================
-- 結果表示
-- =========================================

on showResult(titleText, detailText, scopeText, baseFolderText, okCount, aiCount, normalCount, skipCount, outsideCount, ngCount, errList)
	
	set reportText to titleText & return & ¬
		detailText & return & ¬
		scopeText & return & ¬
		"基準フォルダ: " & baseFolderText & return & return & ¬
		"処理成功: " & okCount & "件" & return & ¬
		"  うちAI: " & aiCount & "件" & return & ¬
		"  うち通常画像: " & normalCount & "件" & return & ¬
		"スキップ: " & skipCount & "件" & return & ¬
		"対象外（保存先フォルダ外）: " & outsideCount & "件" & return & ¬
		"失敗: " & ngCount & "件"
	
	if (skipCount + ngCount) > 0 then
		
		set previewText to ""
		set maxLines to 12
		set listCount to count of errList
		
		if listCount < maxLines then
			set showLines to listCount
		else
			set showLines to maxLines
		end if
		
		repeat with n from 1 to showLines
			set previewText to previewText & item n of errList & return
		end repeat
		
		if listCount > maxLines then
			set previewText to previewText & "…以下省略"
		end if
		
		my showFrontDialog(reportText & return & return & "詳細:" & return & previewText)
		
	else
		
		my showFrontDialog(reportText)
		
	end if
	
end showResult


-- =========================================
-- 前面ダイアログ表示
-- =========================================

on showFrontDialog(messageText)
	
	tell application "System Events"
		activate
		
		display dialog messageText buttons {"OK"} default button "OK" with title "InDesign 配置リンクラベル"
	end tell
	
end showFrontDialog
