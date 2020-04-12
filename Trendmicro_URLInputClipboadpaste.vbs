Option Explicit

'クリップボードへファイル情報貼り付け

' FileSystemObject オブジェクトを作成する
    Dim objFso
      Set objFso = CreateObject("Scripting.FileSystemObject")

' C:\work\test\clip.batが存在する場合
    If objFso.FileExists("C:\work\test\clip.bat") Then

' WshShellオブジェクトを作成する
    Dim WshShell, RtnCode
    Set WshShell = WScript.CreateObject("WScript.Shell")

' batファイルを実行し、batファイルからの戻り値をRtnCodeに設定される
    RtnCode = WshShell.Run("C:\work\test\clip.bat",0,1)

' 戻り値が 0 の場合
    If RtnCode = 0 Then
        ' 正常終了 のメッセージを表示する
        WScript.Echo "bat  clipcopy 正常終了"
' 戻り値が 0 以外の場合
    Else
        ' 異常終了 のメッセージを表示する
        WScript.Echo " bat  clipcopy 異常終了"
    End If
' オブジェクトを開放する
    Set WshShell = Nothing

' C:\work\test\clip.bat が存在しない場合
   Else

' batファイルが存在しません のメッセージを表示する
    WScript.Echo "batファイルが存在しません"

   End If

' オブジェクトを開放する
    Set objFso = Nothing
    
'Webサイト入力フォームへクリップボードから情報入力

  Dim objIE  
  
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

 'InternetExplorerフォームページを起動
  'objIE.navigate "https://www.google.co.jp"  ' Chrome
  objIE.navigate2 "https://global.sitesafety.trendmicro.com/" ' URL
  
   'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100        
    Loop
  
'クリップボードよりURLを自動入力する※複数のクリップボードを一行ずつ貼り付ける方法が不明
  
  Dim mystr
  Dim objInpTxt

     mystr = OBJIE.document.parentWindow.clipboardData.GetData("text")
      Set objInpTxt = objIE.document.getElementsByName("urlname")(0)
      objInpTxt.Value = mystr

  Dim objbutton
    
   'button要素をコレクションとして取得
    Set objbutton = objIE.document.getElementById("getinfo")     
    
    objbutton.click
        
   'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100        
    Loop
    
    
    '評価をテキストファイルに書き出す
    OutputText objIE.document.GetElementsByClassName("labeltitleresult")(0).innerText
    OutputText objIE.document.GetElementsByClassName("labelinfo columnholder2")(0).innerText
      
'テキストファイルへ出力
  Function OutputText(ByVal strMsg)
 
    Dim objFSO     ' FileSystemObject
    Dim objText    ' ファイル書き込み用
 
    'ファイルシステムオブジェクト
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    'テキストファイルを開く
    Set objText = objFSO.OpenTextFile("サイト情報.txt", 8, True, -1)
    
    objText.write vbCrLf
    objText.write strMsg '改行
    
    objText.close
    
    'オブジェクト変数をクリア
    Set objFSO = Nothing
    Set objText = Nothing 
    
     End Function
 
 'InternetExplorerを閉じる
    objIE.Quit
    
    
