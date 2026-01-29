# ExcelAutomation
売り上げ管理用ローカルアプリ

# アプリ実行手順
1. App.config でDB接続、出力先ファイルパスを設定
2. ユーザー認証API（CenterSystem）を起動
3. ログインウィンドウでユーザー名、パスワードを入力してログイン
	管理者
		ユーザー名：admin001、パスワード：Admin001
	ゲスト
		ユーザー名：guest001、パスワード：Guest001
4. SalesSampleフォルダのExcelファイルを利用して売上登録（管理者のみ可能）
5. 4で登録した売上履歴を検索