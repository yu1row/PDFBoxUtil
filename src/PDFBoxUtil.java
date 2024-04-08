import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.pdfbox.contentstream.operator.Operator;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdfparser.PDFStreamParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.PDPageContentStream.AppendMode;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.font.PDCIDFont;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDFontDescriptor;
import org.apache.pdfbox.pdmodel.font.PDType0Font;

/**
 * PDFBoxのユーティリティクラス。
 * <h3>PDFBoxのフォント名の正常化</h3>
 * <p>
 * SVFの出力するPDFのフォント名がSJISのコード参照の場合、
 * PDFBoxでの読み込みで文字化けが発生し、
 * 保存したPDFのフォントが不正となる場合がある（以下例）。
 * </p>
 * <pre>
 * 例：「ＭＳ明朝」の場合
 * "#82l#82r#96#BE#92#A9" -> "?l?r??’?"（文字化け）
 * </pre>
 * <p>
 * SJISの上位バイトの#82,#82,#96,#92などはC1制御コードとして
 * Unicodeに変換されてしまう（PDFBoxが改修される可能性はある）。
 * このクラスの {@link #fontNormalize(PDDocument)} メソッドでは、
 * C1制御コードを逆変換してSJIS文字列とした場合に元の文字列から
 * 変換し、さらにフォント名を英名表現に変換して差し替えを行う。
 * </p>
 * <p>
 * 変換に対応するフォントは以下の通り。
 * </p>
 * <ul>
 * <li>MS明朝</li>
 * <li>MSP明朝</li>
 * <li>MSゴシック</li>
 * <li>MSPゴシック</li>
 * </ul>
 * <h3>PDFページへの文字列の描画</h3>
 * <p>
 * SVFの出力するPDFページはデフォルトではDPI=400として出力する。
 * PDFBoxでテキストを追記する場合などはこれが原因で位置計算や
 * フォントサイズの指定などに問題が発生してしまうことがある。
 * このクラスの {@link #writeText(PDDocument, PDPage, String, PDFont, float, EnumSet, float, float)}、
 * {@link #writeText(PDDocument, PDPage, String, PDFont, float, EnumSet, float, float, boolean)}
 * メソッドでは、必要ならページの初期状態のグラフィックス状態を保持し
 * 計算の狂いが生じないように、指定したフォント、フォントサイズ、
 * 縦位置、横位置で文字列を追加する手段を提供する。
 * </p>
 * @author yu1row
 */
public final class PDFBoxUtil {

	/** UnicodeからC1制御コードへの変換マップ */
	private static final Map<Integer, Integer> UNICODE_TO_C1_MAP;
	/** 日本語フォント名から英名への変換マップ */
	private static final Map<String, String> FONT_NAME_MAP;
	static {
		// UnicodeからC1制御コードへの変換マップを初期化
		// ※0x81, 0x8d, 0x8f, 0x90, 0x9dは変換対象外
		UNICODE_TO_C1_MAP = new HashMap<Integer, Integer>();
		UNICODE_TO_C1_MAP.put(0x201A, 0x82);
		UNICODE_TO_C1_MAP.put(0x0192, 0x83);
		UNICODE_TO_C1_MAP.put(0x201E, 0x84);
		UNICODE_TO_C1_MAP.put(0x2026, 0x85);
		UNICODE_TO_C1_MAP.put(0x2020, 0x86);
		UNICODE_TO_C1_MAP.put(0x2021, 0x87);
		UNICODE_TO_C1_MAP.put(0x02C6, 0x88);
		UNICODE_TO_C1_MAP.put(0x2030, 0x89);
		UNICODE_TO_C1_MAP.put(0x0160, 0x8A);
		UNICODE_TO_C1_MAP.put(0x2039, 0x8B);
		UNICODE_TO_C1_MAP.put(0x0152, 0x8C);
		UNICODE_TO_C1_MAP.put(0x2018, 0x91);
		UNICODE_TO_C1_MAP.put(0x2019, 0x92);
		UNICODE_TO_C1_MAP.put(0x201C, 0x93);
		UNICODE_TO_C1_MAP.put(0x201D, 0x94);
		UNICODE_TO_C1_MAP.put(0x2022, 0x95);
		UNICODE_TO_C1_MAP.put(0x2013, 0x96);
		UNICODE_TO_C1_MAP.put(0x2014, 0x97);
		UNICODE_TO_C1_MAP.put(0x02DC, 0x98);
		UNICODE_TO_C1_MAP.put(0x2122, 0x99);
		UNICODE_TO_C1_MAP.put(0x0161, 0x9A);
		UNICODE_TO_C1_MAP.put(0x203A, 0x9B);
		UNICODE_TO_C1_MAP.put(0x0153, 0x9C);
		UNICODE_TO_C1_MAP.put(0x0178, 0x9F);
		// 日本語フォント名から英名への変換マップを初期化
		FONT_NAME_MAP = new HashMap<String, String>();
		FONT_NAME_MAP.put("ＭＳ明朝", "MS Mincho");
		FONT_NAME_MAP.put("ＭＳ 明朝", "MS Mincho");
		FONT_NAME_MAP.put("@ＭＳ明朝", "MS Mincho");
		FONT_NAME_MAP.put("@ＭＳ 明朝", "MS Mincho");
		FONT_NAME_MAP.put("ＭＳＰ明朝", "MS PMincho");
		FONT_NAME_MAP.put("ＭＳ Ｐ明朝", "MS PMincho");
		FONT_NAME_MAP.put("@ＭＳＰ明朝", "MS PMincho");
		FONT_NAME_MAP.put("@ＭＳ Ｐ明朝", "MS PMincho");
		FONT_NAME_MAP.put("ＭＳゴシック", "MS Gothic");
		FONT_NAME_MAP.put("ＭＳ ゴシック", "MS Gothic");
		FONT_NAME_MAP.put("@ＭＳゴシック", "MS Gothic");
		FONT_NAME_MAP.put("@ＭＳ ゴシック", "MS Gothic");
		FONT_NAME_MAP.put("ＭＳＰゴシック", "MS PGothic");
		FONT_NAME_MAP.put("ＭＳ Ｐゴシック", "MS PGothic");
		FONT_NAME_MAP.put("@ＭＳＰゴシック", "MS PGothic");
		FONT_NAME_MAP.put("@ＭＳ Ｐゴシック", "MS PGothic");
	}

	/**
	 * テキスト出力位置の列挙.
	 */
	public enum TextWritePosition {
		/** 上下中央（{@link TextWritePosition#TOP}, {@link TextWritePosition#BOTTOM} より優先） */
		MIDDLE,
		/** 上端（{@link TextWritePosition#BOTTOM} より優先） */
		TOP,
		/** 下端 */
		BOTTOM,
		/** 左右中央（{@link TextWritePosition#LEFT}, {@link TextWritePosition#RIGHT} より優先） */
		CENTER,
		/** 右端（{@link TextWritePosition#RIGHT} より優先） */
		RIGHT,
		/** 左端 */
		LEFT,
	}

	/**
	 * インスタンス化不可クラス。
	 */
	private PDFBoxUtil() {
	}

	/**
	 * PDFドキュメントのフォント名の正常化ができるかを返す。
	 * @param doc PDFドキュメント
	 * @return フォント名の正常化ができる場合：true
	 * @throws IOException 入出力時の例外
	 */
	public static boolean checkFontNormalizable(PDDocument doc) throws IOException {
		boolean ret = false;
		// フォント一覧を取得して処理
		for (PDFont font : getFontList(doc)) {
			// Type 0のフォントのみ対象
			if (font instanceof PDType0Font) {
				// 修正後のフォント名から英名取得成功時、変換対象
				if (FONT_NAME_MAP.get(getFixedFontName(font)) != null) {
					ret = true;
					break;
				}
			}
		}
		return ret;
	}

	/**
	 * PDFファイルのフォント名を正常化して保存する。
	 * @param pathInput 正常化対象のPDFファイルのパス
	 * @param pathOutput 保存されるPDFファイルのパス
	 * @throws IOException 入出力時の例外
	 */
	public static void fontNormalize(String pathInput, String pathOutput) throws IOException {
		InputStream is = new FileInputStream(pathInput);
		PDDocument doc = PDDocument.load(is);
		fontNormalize(doc);
		doc.save(pathOutput);
		doc.close();
	}

	/**
	 * PDFBoxのフォント名を正常化する。
	 * @param doc PDFドキュメント
	 * @return フォント名の変換が行われた場合：true
	 * @throws IOException 入出力時の例外
	 */
	public static boolean fontNormalize(PDDocument doc) throws IOException {
		boolean ret = false;
		// フォント一覧を取得して処理
		for (PDFont font : getFontList(doc)) {
			// Type 0のフォントのみ対象
			if (font instanceof PDType0Font) {
				// 修正後のフォント名を取得し、英名変換
				String newName = FONT_NAME_MAP.get(getFixedFontName(font));
				// 英名取得成功時、ベース／子孫(CID)／属性のフォント名を修正
				if (newName != null) {
					PDType0Font type0font = (PDType0Font)font;
					type0font.getCOSObject().setString(COSName.BASE_FONT, newName);
					if (type0font.getCOSObject().containsKey(COSName.NAME)) {
						type0font.getCOSObject().setString(COSName.NAME, newName);
					}
					PDCIDFont descendantFont = type0font.getDescendantFont();
					descendantFont.getCOSObject().setString(COSName.BASE_FONT, newName);
					PDFontDescriptor fontDescriptor = descendantFont.getFontDescriptor();
					fontDescriptor.setFontName(newName);
					// 変換が行われたことを返す
					ret = true;
				}
			}
		}
		return ret;
	}

	/**
	 * PDFドキュメントで使用されているフォント一覧を、重複を除いて取得する。
	 * @param doc PDFドキュメント
	 * @return PDFドキュメントで使用されているフォント一覧
	 * @throws IOException 入出力時の例外
	 */
	private static Set<PDFont> getFontList(PDDocument doc) throws IOException {
		Set<PDFont> ret = new HashSet<PDFont>();
		for (PDPage page : doc.getPages()) {
			PDResources res = page.getResources();
			for (COSName fontName : res.getFontNames()) {
				PDFont font = res.getFont(fontName);
				if (!ret.contains(font)) {
					ret.add(font);
				}
			}
		}
		return ret;
	}

	/**
	 * C1制御コードの逆変換を利用し、フォント名を修正して返す。
	 * @param font 対象のフォント
	 * @return 変換後のフォント名、または変換されない場合は空文字列
	 */
	private static String getFixedFontName(PDFont font) {
		// 文字化けが発生している可能性のあるフォント名を取得
		String fontName = font.getName();
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		// フォント名を１文字ずつ処理
		for (int i=0; i<fontName.length(); i++) {
			int c = fontName.codePointAt(i);
			// １文字ずつバイト配列として取得
			if (UNICODE_TO_C1_MAP.containsKey(c)) {
				// 制御コードの逆変換ができる場合、変換
				bos.write(UNICODE_TO_C1_MAP.get(c));
			} else {
				bos.write(c);
			}
		}

		String fixedName = "";
		try {
		// バイト配列をSJIS文字列として取得
			String tempName = String.format(new String(bos.toByteArray(), "SJIS"));
			// フォント名に変更がある場合のみ戻り値に設定
			if (!fontName.equals(tempName)) {
				fixedName = tempName;
			}
		} catch (UnsupportedEncodingException e) {
			// 「SJIS」固定で変換するため、この例外は発生しない想定
		}
		return fixedName;
	}

	/**
	 * PDFページの指定位置に文字列を描画する。
	 * <p>SVF対策（デフォルトのDPI=400）としてグラフィックス状態を分離させた描画を行う。</p>
	 * @param doc PDFドキュメント
	 * @param page PDFページ
	 * @param contentStream 
	 * @param text 出力する文字列
	 * @param font フォント
	 * @param fontSize フォントサイズ
	 * @param position 文字列の出力位置
	 * （ページ下中央の例：<pre>EnumSet.of(PDFBoxUtil.TextWritePosition.BOTTOM, PDFBoxUtil.TextWritePosition.CENTER)</pre>）
	 * @param offsetX 横位置のオフセット（単位はピクセル、RIGHT指定ありの場合は右端からの距離）
	 * @param offsetY 縦位置のオフセット（単位はピクセル、TOP指定ありの場合は上端からの距離）
	 * @throws IOException 入出力時の例外
	 */
	public static void writeText(PDDocument doc, PDPage page, String text, PDFont font, float fontSize, EnumSet<TextWritePosition> position, float offsetX, float offsetY) throws IOException {
		writeText(doc, page, text, font, fontSize, position, offsetX, offsetY, true);
	}

	/**
	 * PDFページの指定位置に文字列を描画する。
	 * @param doc PDFドキュメント
	 * @param page PDFページ
	 * @param contentStream 
	 * @param text 出力する文字列
	 * @param font フォント
	 * @param fontSize フォントサイズ
	 * @param position 文字列の出力位置
	 * （ページ下中央の例：<pre>EnumSet.of(PDFBoxUtil.TextWritePosition.BOTTOM, PDFBoxUtil.TextWritePosition.CENTER)</pre>）
	 * @param offsetX 横位置のオフセット（単位はピクセル、RIGHT指定ありの場合は右端からの距離）
	 * @param offsetY 縦位置のオフセット（単位はピクセル、TOP指定ありの場合は上端からの距離）
	 * @param isSeparateGraphicsState グラフィックス状態を分離させるか：trueにするとページ先頭でグラフィックス状態を保存し、テキスト描画前に復元する（SVF対策）
	 * @throws IOException 入出力時の例外
	 */
	public static void writeText(PDDocument doc, PDPage page, String text, PDFont font, float fontSize, EnumSet<TextWritePosition> position, float offsetX, float offsetY, boolean isSeparateGraphicsState) throws IOException {
		// ページ全体がグラフィックス状態保存されているか
		boolean existPageGS = checkStartWithSavedGraphicsState(page);

		// X座標位置計算（用紙左下が原点）
		float x = 0f;
		float pageWidth = page.getMediaBox().getWidth();
		float textWidth = font.getStringWidth(text) / 1000f * fontSize;
		if (position.contains(TextWritePosition.CENTER)) {
			x = (pageWidth - textWidth) / 2f + offsetX;
		} else if (position.contains(TextWritePosition.LEFT)) {
			x = offsetX;
		} else {
			x = pageWidth - textWidth - offsetX;
		}

		// Y座標位置計算（用紙左下が原点）
		float y = 0f;
		float pageHeight = page.getMediaBox().getHeight();
		float textHeight = font.getFontDescriptor().getCapHeight() / 1000f * fontSize;
		if (position.contains(TextWritePosition.MIDDLE)) {
			y = (pageHeight - textHeight) / 2f + offsetY;
		} else if (position.contains(TextWritePosition.BOTTOM)) {
			y = offsetY;
		} else {
			y = pageHeight - textHeight - offsetY;
		}

		// ページ初期のグラフィックス状態を先頭で保持、テキスト描画直前で復帰する（オプション）
		PDPageContentStream contentStream;
		if (isSeparateGraphicsState && !existPageGS) {
			contentStream = new PDPageContentStream(doc, page, AppendMode.PREPEND, true);
			contentStream.saveGraphicsState();
			contentStream.close();
		}
		contentStream = new PDPageContentStream(doc, page, AppendMode.APPEND, true);
		if (isSeparateGraphicsState && !existPageGS) {
			contentStream.restoreGraphicsState();
		}

		// テキスト描画
		contentStream.saveGraphicsState();
		contentStream.beginText();
		contentStream.setFont(font, fontSize);
		contentStream.newLineAtOffset(x, y);
		contentStream.showText(text);
		contentStream.endText();
		contentStream.restoreGraphicsState();
		contentStream.close();
	}

	/**
	 * ページがグラフィックス状態保存から始まっているかを返す。
	 * @param page PDFページ
	 * @return ページがグラフィックス状態保存から始まっている場合：true
	 * @throws IOException 入出力時の例外
	 */
	private static boolean checkStartWithSavedGraphicsState(PDPage page) throws IOException {
		boolean ret = false;
		// PDFページの解析
		PDFStreamParser parser = new PDFStreamParser(page);
		parser.parse();
		// 解析結果のトークンリストを取得
		List<Object> tokens = parser.getTokens();
		// トークンが1件以上で、グラフィックス状態保存（qコマンド）から始まっているかを調べる
		if (0 < tokens.size()) {
			Object o = tokens.get(0);
			if (o instanceof Operator) {
				if ("q".equals(((Operator)o).getName())) {
					ret = true;
				}
			}
		}
		return ret;
	}
}
