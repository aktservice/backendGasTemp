# backendGasTemp

## GAS を LOCAL で構築する為のテンプレートです

### 使用しているソフト等

- VsCode(editor)
  - VsCode Extension: more(debug)
- TypeScript(js is also fine)
- Nodejs(npm or yarn)
  - @types/google-apps-script
  - Google/Clasp(global)
- git(use git bash)

### ざっくりとした流れ

src フォルダで開発 → clasp push で 反映

### 使い方

Vscode をインストール（User インストールの場合は管理者権限は不要です）  
[Vscode](https://code.visualstudio.com/)  
nodejs(npm) をインストール(要管理者権限)  
[nodejs](https://nodejs.org/ja/)  
nodejs と一緒にインストールされますが一応  
[npm](https://www.npmjs.com/)  
git をインストール(要管理者権限)  
[git](https://git-scm.com/)  
3 点をインストール後に

```bash

#bash or powershell
# git bash を開いて下記のコマンドを打ち込みして下さい（コピペＯＫ）
git clone https://github.com/yoshikoro/backendGasTemp.git
cd backendGasTemp
npm install
code ./
#ここでVsCodeが開くのでそこから開発
# clasp コマンドをGlobalにインストールした状態で
#Googleアカウントへログイン
clasp login
#ログイン成功後にプロジェクトのScriptIDを.clasp.jsonのscriptIdに記入して
clasp open
# .clasp.jsonに記載されたIdを開くので確認後
clasp push
```
