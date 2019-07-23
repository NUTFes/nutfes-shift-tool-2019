## Setup

**実行環境**  
- python3.6

**cloneとライブラリのインストール**  
```bash
$ git clone git@github.com:NUTFes/nutfes-shift-tool-2019.git
$ cd nutfes-shift-tool-2019
$ pip install -r requirements.txt
```

**Sheets APIの承認情報を取得**  
参考：[PythonとSheets API v4でGoogleスプレッドシートを読み書きする - kumilog.net](https://www.kumilog.net/entry/2018/03/22/090000)  
上記記事を参考に，Sheets APIを有効化し，承認情報を作成する．  
作成してダウンロードした`client_secret_XXX.json`を`client_secret.json`にリネームし，`nutfes-shift-tool-2019`ディレクトリにコピーする．  

**アカウント承認**  
```bash
$ python spreadsheet_api.py
```

これを実行するとブラウザが開き，Googleアカウントログイン画面が表示されるので，技大祭googleアカウントでログインする．  
使用を承認すると，今のディレクトリに`.credentials`ディレクトリが作成される．  
GoogleAPIを使用するには，`client_secret.json`と`./credentials`の2つが必要になる．  

**設定ファイルの作成**  
設定ファイルのテンプレをコピーする．  

```bash
$ cp config_exsample.py config.py
```

`config.py`に必要な情報を埋める．


**起動**  
```bash
$ python app.py
```
