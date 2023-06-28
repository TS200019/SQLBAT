# SQL Batch
WindowsでODBCを利用したSQL実行専用のバッチツールです

# インストール
　①フォルダごと配置する。
　②run.batを編集する。
　　接続情報を書き換える。

　　C:\Windows\SysWOW64\CScript //nologo bin\sql.vbs "{接続情報}" {SQLリスト}

　　　接続情報　→ "DSN={データソース名};UID={ユーザ名};PWD={パスワード}"
　　　SQLリスト → prmフォルダのリストファイル名

SQLをリストに格納して実行するバッチシステムを作成してみた
