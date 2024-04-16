## 概要

帳票基盤のSVFで出力したPDFをPDFBoxで取り扱う場合について、以下の問題を解消する機能を有するユーティリティクラスです。
* MSゴシック、MS明朝、MSPゴシック、MSP明朝フォントを使用しているPDFのフォント名が不正な状態で保存されてしまう
* PDFのページに文字列を追加する場合、DPI（=400）を理由として、指定した位置に描画できない

## 各メソッドの使用例

### フォント名の正常化

```
PDFBoxUtil.fontNormalize("./pdf/input.pdf", "./pdf/output.pdf");
```

### PDFページへの文字列の描画

```
try {
	// ファイルを開いて1ページ目を取得
	InputStream is = new FileInputStream("./pdf/target.pdf");
	PDDocument doc = PDDocument.load(is);
	PDPage page = doc.getPage(0);
	
	// ページの一番下、中央に文字列を追加
	EnumSet<PDFBoxUtil.TextWritePosition> pos = EnumSet.of(PDFBoxUtil.TextWritePosition.BOTTOM, PDFBoxUtil.TextWritePosition.CENTER);
	PDFBoxUtil.writeText(doc, page, "Center text", PDType1Font.HELVETICA, 10, pos, 0, 0);
	
	// PDFを保存
	doc.save(outfile);
	doc.close();
} catch (Exception e) {
	e.printStackTrace();
}
```
