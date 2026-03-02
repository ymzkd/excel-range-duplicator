# Range Grid Duplicator

Excel の選択範囲をグリッド状に縦横へ複製する Office アドイン。

## 機能

1. Excel 上でセル範囲を選択する
2. タスクペインで横方向・縦方向の複製数を指定する
3. 「複製」ボタンを押すと、値・書式・数式がグリッド状にコピーされる

- 選択範囲はリアルタイムで自動取得（ボタン操作不要）
- `Range.copyFrom` API により値・数値書式・フォント・罫線・結合セル等を一括コピー
- Ctrl+Z で操作を取り消し可能

## 技術スタック

- React 18 + Fluent UI v9（`@fluentui/react-components`）
- webpack 5（babel-loader で JSX をトランスパイル）
- Office JavaScript API（ExcelApi 1.9+）

## プロジェクト構成

```
├── manifest.xml          # 開発用マニフェスト（localhost:3000）
├── manifest.prod.xml     # 本番用マニフェスト（GitHub Pages）
├── package.json
├── webpack.config.js
├── src/
│   ├── taskpane/
│   │   ├── taskpane.jsx  # メインコンポーネント + Excel API ロジック
│   │   ├── taskpane.html
│   │   └── taskpane.css
│   └── commands/
│       ├── commands.js   # マニフェスト要求のスタブ
│       └── commands.html
└── assets/
    └── icon-*.png        # リボンアイコン（16/32/64/80px）
```

## 開発環境のセットアップ

### 前提条件

- Node.js 18 以上
- Excel デスクトップ版（Microsoft 365）または Excel Online

### インストール

```bash
npm install
```

### HTTPS 証明書のインストール（初回のみ）

Office アドインは開発時も HTTPS が必須です。

```bash
npx office-addin-dev-certs install
```

### 開発サーバーの起動

```bash
# dev server 起動 + Excel にサイドロード（デスクトップ版）
npm start

# dev server のみ起動（Excel Online で使う場合）
npm run dev-server
```

`https://localhost:3000` で配信されます。

### Excel Online でのサイドロード

1. `npm run dev-server` を実行
2. Excel Online でブックを開く
3. ホーム → アドイン → マイ アドインのアップロード → `manifest.xml` を選択

### 停止

```bash
npm stop
```

## デプロイ（GitHub Pages）

### ビルド

```bash
npm run build
```

`dist/` フォルダにビルド成果物が出力されます。

### GitHub Pages へのデプロイ

```bash
# gh-pages ブランチに切り替えてビルド成果物をコミット
git checkout gh-pages
cp -r dist/* .
git add -A
git commit -m "Deploy"
git push origin gh-pages
git checkout master
```

デプロイ先: `https://<username>.github.io/<repo>/`

`manifest.prod.xml` 内の URL がデプロイ先と一致していることを確認してください。

## 常時利用の設定（Windows デスクトップ Excel）

dev server を起動せずに、GitHub Pages から配信されるアドインを常時使うための設定です。

### 1. 共有フォルダの作成

1. `C:\AddinManifests` フォルダを作成
2. `manifest.prod.xml` を `C:\AddinManifests\manifest.xml` としてコピー
3. フォルダを右クリック → プロパティ → 共有タブ → 共有 → 自分のユーザーを追加して共有

### 2. Excel への登録

1. Excel → ファイル → オプション → セキュリティ センター → セキュリティ センターの設定
2. 信頼できるアドイン カタログ
3. カタログの URL に `\\localhost\AddinManifests` を入力 → カタログの追加
4. 「メニューに表示」にチェック → OK
5. Excel を再起動
6. ホーム → アドイン → 共有フォルダ → Range Grid Duplicator を追加

以後、Excel を開くだけでアドインが利用可能になります。

## ライセンス

MIT
