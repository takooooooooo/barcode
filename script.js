// --- ライブラリ・HTML要素の前提 ---
// - jsPDF, PapaParse, opentype.js, JSZip, FileSaver.js が読み込まれていること
// - HTMLに以下のIDを持つ要素が存在すること:
//   - barcodeInputArea (textarea)
//   - generateFromTextButton (button)
//   - csvFile (input type=file)
//   - zipFilenameInput (input type=text)
//   - status (div or similar for messages)
// --- ---

// 定数
const WIDTH = 29.832; // mm
const HEIGHT = 18.552; // mm
const BAR_WIDTH = 0.264; // mm
const BAR_HEIGHT = 15.296; // mm
const GUARD_BAR_HEIGHT = 2.992; // mm
const EDGE = "101";
const MIDDLE = "01010";
const CODES = {
    "A": ["0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011"],
    "B": ["0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111"],
    "C": ["1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100"]
};
const LEFT_PATTERN = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"];
const PT_TO_MM = 25.4 / 72; // 参考

// EAN-13バーコード生成関数 (変更なし - 前回のバージョン)
function generateEan13Barcode(number) {
    function calculateCheckDigit(num) {
        // チェックデジット計算用に先頭と末尾を除いた11桁を取得 (入力が12桁の場合を想定)
        const digits = num.slice(0, 12).split('').map(n => parseInt(n)); // 12桁目がチェックデジットとして計算に使われる
        if (digits.some(isNaN)) throw new Error(`Invalid non-numeric characters for checksum: ${num.slice(0, 12)}`);
        // EAN-13のチェックデジット計算 (奇数桁 * 1 + 偶数桁 * 3) - ※0ベースインデックスでは逆
        const totalSum = digits.reduce((sum, n, i) => sum + (i % 2 !== 0 ? n * 3 : n), 0); // 0始まりなので i%2 !== 0 が偶数番目(2,4,6...)
        const checkDigit = (10 - (totalSum % 10)) % 10;
        if (isNaN(checkDigit)) throw new Error(`Checksum calculation resulted in NaN for: ${num}`);
        return checkDigit;
    }
    function encode(num) {
        if (typeof num !== 'string') throw new Error(`Invalid input type: ${typeof num}`);
        let codeToProcess = num.trim();
        if (codeToProcess.length === 12 && /^\d+$/.test(codeToProcess)) {
            try {
                const checkDigit = calculateCheckDigit(codeToProcess); // 12桁からチェックデジット計算
                codeToProcess += checkDigit; // 計算したチェックデジットを付与
             }
            catch (e) { throw new Error(`Checksum calculation failed for ${num}: ${e.message}`); }
        }
        if (codeToProcess.length !== 13 || !/^\d+$/.test(codeToProcess)) throw new Error(`Invalid JAN code format: '${codeToProcess}' (original: '${num}')`);
        const firstDigit = parseInt(codeToProcess[0]);
        if (isNaN(firstDigit) || firstDigit < 0 || firstDigit >= LEFT_PATTERN.length) throw new Error(`Invalid first digit: ${codeToProcess[0]}`);
        const pattern = LEFT_PATTERN[firstDigit];
        let encoded = EDGE;
        for (let i = 0; i < 6; i++) {
            const patternType = pattern[i]; const digit = parseInt(codeToProcess[i + 1]);
            if (isNaN(digit) || !CODES[patternType] || digit < 0 || digit >= CODES[patternType].length) throw new Error(`Invalid digit/pattern for left part: pattern=${patternType}, digit=${codeToProcess[i+1]}`);
            encoded += CODES[patternType][digit];
        }
        encoded += MIDDLE;
        for (let i = 7; i < 13; i++) { // 右側は7桁目から13桁目(インデックス6から12)
            const digit = parseInt(codeToProcess[i]);
             if (isNaN(digit) || !CODES["C"] || digit < 0 || digit >= CODES["C"].length) throw new Error(`Invalid digit for right part: ${codeToProcess[i]}`);
            encoded += CODES["C"][digit];
        }
        encoded += EDGE;
        if (typeof encoded !== 'string' || encoded.length !== 95) throw new Error("Internal encoding error occurred.");
        return encoded;
    }
    try {
        const result = encode(number);
        if (typeof result !== 'string') throw new Error("Encoding function did not return a string.");
        return result;
    } catch (e) { console.error(`Error generating EAN-13 for "${number}": ${e.message}`); throw e; }
}


// PDF生成関数 (★ 文字送り幅の計算を修正 & ★ 白背景追加)
async function generateBarcodePDF(barcodeNumber) {
    console.log(`[generateBarcodePDF] STARTED for: ${barcodeNumber}`);
    try {
        // --- ライブラリチェック ---
        if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') throw new Error("jsPDF library is not loaded.");
        if (typeof window.opentype === 'undefined') throw new Error("opentype.js library is not loaded.");
        const { jsPDF } = window.jspdf;

        const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: [WIDTH, HEIGHT] });

        // ★★★ 追加: 白い背景を描画 ★★★
        doc.setFillColor(255, 255, 255); // 塗りつぶし色を白に設定 (RGB: 255, 255, 255)
        doc.rect(0, 0, WIDTH, HEIGHT, 'F'); // PDFのサイズ全体 (0,0) から (WIDTH, HEIGHT) までを塗りつぶす ('F')
        console.log('[generateBarcodePDF] White background drawn.');
        // ★★★ 追加 ここまで ★★★

        // --- バーコード生成と検証 ---
        console.log(`[generateBarcodePDF] Calling generateEan13Barcode for ${barcodeNumber}...`);
        let barcode;
        try { barcode = generateEan13Barcode(barcodeNumber); }
        catch(barcodeGenError) { throw new Error(`Failed during barcode string generation for ${barcodeNumber}: ${barcodeGenError.message}`); }
        if (typeof barcode !== 'string' || barcode.length !== 95) throw new Error(`Internal error: Failed to get valid barcode string (received ${typeof barcode}, length ${barcode?.length}) for ${barcodeNumber}`); // barcode.length チェックを95に変更
        console.log(`[generateBarcodePDF] Barcode data validated: ${barcode.length} chars`);

        // --- バーコード描画 ---
        let x = 2.9;
        doc.setFillColor(0, 0, 0); // バーの色を黒に戻す
        for (let i = 0; i < barcode.length; i++) { if (barcode[i] === "1") doc.rect(x, 0.264, BAR_WIDTH, BAR_HEIGHT, 'F'); x += BAR_WIDTH; }
        console.log('[generateBarcodePDF] Barcode bars drawn.');

        // --- ガードバー描画 ---
        x = 2.9;
        const guardBar = "10100000000000000000000000000000000000000000001010000000000000000000000000000000000000000000101";
        doc.setFillColor(0, 0, 0); // ガードバーの色も黒に設定
        for (let i = 0; i < guardBar.length; i++) { if (guardBar[i] === "1") doc.rect(x, 15.56, BAR_WIDTH, GUARD_BAR_HEIGHT, 'F'); x += BAR_WIDTH; }
        console.log('[generateBarcodePDF] Guard bars drawn.');

        // --- フォント読み込み ---
        let font;
        try { font = await opentype.load('MyriadPro.ttf'); } // ★フォントパス確認
        catch (fontLoadError) { throw new Error('Font loading failed: ' + fontLoadError.message); }
        if (!font) throw new Error('Font loading resulted in null/undefined');
        console.log('[generateBarcodePDF] Font loaded successfully.');

        // --- パラメータ設定 ---
        const fontSizePt = 8.75;
        const letterSpacing = 0.15;    // 文字間の *追加* スペース
        const baselineYOffset = 0.523;
        const fontSizeScaleFactor = 0.35;
        const finalFontSizeForGetPath = fontSizePt * fontSizeScaleFactor;
        // barcodeNumber は 13桁のはずなので、インデックスでアクセス
        const parts = [
            { text: barcodeNumber[0], startX: 0.052 },
            { text: barcodeNumber.slice(1, 7), startX: 4.082 }, // 1から6 (index 1 to 6)
            { text: barcodeNumber.slice(7, 13), startX: 16.34 } // 7から12 (index 7 to 12) - チェックデジット含む
        ];
        const baselineY = HEIGHT - baselineYOffset;

        // --- 文字描画 ---
        console.log('[generateBarcodePDF] Processing text parts...');
        doc.setFillColor(0, 0, 0); // 文字の色を黒に設定
        for (const part of parts) {
            let currentX = part.startX;
            const text = part.text;
            for (let i = 0; i < text.length; i++) {
                const char = text[i];
                try {
                    const textPath = font.getPath(char, 0, 0, finalFontSizeForGetPath);
                    if (!textPath || !textPath.commands || textPath.commands.length === 0) {
                         console.warn(`[generateBarcodePDF]   Skipping char '${char}' due to no path data.`);
                         continue;
                    }

                    const pathData = textPath.commands.map(cmd => {
                        const newCmd = { op: cmd.type.toLowerCase() }; newCmd.c = [];
                        if (cmd.type === 'M' || cmd.type === 'L') { newCmd.c = [cmd.x + currentX, cmd.y + baselineY]; }
                        else if (cmd.type === 'Q') { newCmd.c = [cmd.x1 + currentX, cmd.y1 + baselineY, cmd.x + currentX, cmd.y + baselineY]; }
                        else if (cmd.type === 'C') { newCmd.c = [cmd.x1 + currentX, cmd.y1 + baselineY, cmd.x2 + currentX, cmd.y2 + baselineY, cmd.x + currentX, cmd.y + baselineY]; }
                        else if (cmd.type === 'Z') { newCmd.op = 'h'; }
                        else { return null; }
                        return newCmd;
                    }).filter(Boolean);
                    if (pathData.length > 0) { doc.path(pathData); doc.fill('black'); } // fill('black') を使用

                    // ★★★ 文字送り幅の計算方法を変更 ★★★
                    const advanceWidthScaled = font.getAdvanceWidth(char, finalFontSizeForGetPath);
                    if (typeof advanceWidthScaled !== 'number' || isNaN(advanceWidthScaled)) {
                         console.warn(`[generateBarcodePDF]   Invalid advance width for char '${char}'. Using fallback width.`);
                         const bbox = textPath.getBoundingBox();
                         currentX += (bbox ? (bbox.x2 - bbox.x1) : (BAR_WIDTH * 4)) + letterSpacing; // Fallback
                    } else {
                         // フォントの送り幅に追加の文字間隔を加える
                         currentX += advanceWidthScaled + letterSpacing;
                    }
                    // ★★★ ここまで変更 ★★★

                } catch (charError) {
                     throw new Error(`Error processing character '${char}': ${charError.message}`);
                }
            }
        }
        console.log('[generateBarcodePDF] Text parts processed.');

        // --- PDFデータ生成 ---
        const pdfData = doc.output('arraybuffer');
        console.log(`[generateBarcodePDF] FINISHED for: ${barcodeNumber}`);
        return { filename: `${barcodeNumber}.ai`, data: pdfData };

    } catch (error) {
        // --- 関数全体のエラーハンドリング ---
        console.error(`[generateBarcodePDF] CAUGHT error during processing for ${barcodeNumber}. Message: ${error.message}`);
        console.error("[generateBarcodePDF] Error Stack:", error.stack);
        return null; // エラー発生時は null を返す
    }
}


// --- UI要素の取得 (グローバルスコープ) ---
const statusElement = document.getElementById('status');
const csvFileInput = document.getElementById('csvFile');
const barcodeTextArea = document.getElementById('barcodeInputArea');
const generateTextButton = document.getElementById('generateFromTextButton');
const zipFilenameInput = document.getElementById('zipFilenameInput');

// --- ユーティリティ関数 ---
function disableUI(processing = true) {
    if (csvFileInput) csvFileInput.disabled = processing;
    if (barcodeTextArea) barcodeTextArea.disabled = processing;
    if (generateTextButton) generateTextButton.disabled = processing;
    // zipFilenameInput は処理中でも編集可能にするかもしれないので、ここでは無効化しない
}

function updateStatus(message) {
    console.log(`[Status] ${message}`);
    if (statusElement) statusElement.textContent = message;
}

// --- テキストエリア処理関数 ---
async function handleTextarea() {
    updateStatus("テキスト入力を処理中...");
    disableUI(true);

    if (!barcodeTextArea) {
        updateStatus("エラー: テキスト入力エリアが見つかりません。");
        disableUI(false); return;
    }
    const inputText = barcodeTextArea.value;
    if (!inputText.trim()) {
        updateStatus("テキストエリアにバーコード番号（12桁または13桁）を入力してください。");
        disableUI(false); return;
    }

    const barcodeList = inputText.split('\n')
        .map(line => line.trim())
        // 12桁または13桁の数字にマッチするかどうかでフィルタリング
        .filter(line => /^\d{12,13}$/.test(line));

    if (barcodeList.length === 0) {
        updateStatus("有効な12桁または13桁のバーコード番号が見つかりませんでした。");
        disableUI(false); return;
    }

    console.log(`[handleTextarea] Found ${barcodeList.length} valid barcode inputs.`);
    await generateAndZip(barcodeList);
}

// --- CSVファイル処理関数 ---
async function handleFile(event) {
    updateStatus("CSVファイルを処理中...");
    disableUI(true);
    const file = event.target.files[0];
    if (!file) {
        updateStatus("ファイルが選択されていません。");
        disableUI(false); return;
    }
    if (typeof window.Papa === 'undefined') {
         updateStatus("エラー: CSV解析ライブラリ(PapaParse)が見つかりません。");
         disableUI(false); return;
     }

    console.log("[handleFile] Calling Papa.parse...");
    Papa.parse(file, {
        complete: async function(results) {
            console.log("[handleFile] Papa.parse complete.");
            const data = results.data;
            const barcodeList = [];
            data.forEach((row, index) => {
                 // CSVの各行からバーコード番号を取得（ここでは2列目を想定）
                 if (Array.isArray(row) && row.length >= 2) {
                     const barcodeValue = row[1]; // 2列目 (インデックス 1)
                     if (barcodeValue !== null && typeof barcodeValue !== 'undefined') {
                         const barcodeNumber = String(barcodeValue).trim();
                         // 12桁または13桁の数字かチェック
                         if (/^\d{12,13}$/.test(barcodeNumber)) {
                             barcodeList.push(barcodeNumber);
                         } else {
                              console.warn(`[handleFile] Skipping invalid format in CSV row ${index + 1}, column 2: '${barcodeNumber}'`);
                         }
                     } else {
                          console.warn(`[handleFile] Skipping empty value in CSV row ${index + 1}, column 2.`);
                     }
                 } else {
                     console.warn(`[handleFile] Skipping row ${index + 1} due to insufficient columns or invalid format.`);
                 }
            });

            if (barcodeList.length === 0) {
                updateStatus("CSVファイルの2列目に有効な12桁または13桁のJANコードが見つかりませんでした。");
                disableUI(false);
                if (csvFileInput) csvFileInput.value = ''; // ファイル選択をクリア
                return;
            }
            console.log(`[handleFile] Found ${barcodeList.length} valid barcode inputs from CSV.`);
            await generateAndZip(barcodeList);
            if (csvFileInput) csvFileInput.value = ''; // 処理完了後クリア

        },
        header: false, skipEmptyLines: true,
        error: function(error, file) {
            console.error("[handleFile] Papa.parse error:", error);
            updateStatus(`CSVファイルの読み込みエラー: ${error.message}`);
            disableUI(false);
            if (csvFileInput) csvFileInput.value = '';
        }
    });
}

// --- 共通PDF生成＆ZIP化関数 ---
async function generateAndZip(barcodeList) {
    console.log(`[generateAndZip] Starting process for ${barcodeList.length} barcodes.`);
    updateStatus(`${barcodeList.length} 件のバーコード生成を開始...`);

    // ライブラリチェック
    if (typeof window.JSZip === 'undefined') {
         updateStatus("エラー: ZIP生成ライブラリ(JSZip)が見つかりません。");
         disableUI(false); return;
     }
     if (typeof window.saveAs === 'undefined') {
         updateStatus("エラー: ファイル保存ライブラリ(FileSaver.js)が見つかりません。");
         disableUI(false); return;
     }

    try {
        // 各バーコードに対して generateBarcodePDF を呼び出すプロミスの配列を作成
        // generateBarcodePDF はエラー時に null を返す可能性がある
        const pdfPromises = barcodeList.map(barcodeNumber =>
             generateBarcodePDF(barcodeNumber).catch(err => {
                 // 個別の generateBarcodePDF 内で catch 済みだが、念のためここでも catch
                 console.error(`[generateAndZip] Error in promise for ${barcodeNumber}: ${err.message}`);
                 return null; // エラーが発生した場合は null を返すようにする
             })
        );

        const zip = new JSZip();
        const barcodeNumbersProcessed = [...barcodeList]; // 元のリストをコピー

        updateStatus(`${barcodeList.length} 件のバーコードを処理中... (これには時間がかかる場合があります)`);
        // すべてのPDF生成プロミスが完了するのを待つ (成功/失敗問わず)
        const pdfResults = await Promise.allSettled(pdfPromises);
        console.log("[generateAndZip] All PDF generation promises settled.");

        let addedFileCount = 0;
        let failedBarcodes = [];

        pdfResults.forEach((result, index) => {
            const originalBarcode = barcodeNumbersProcessed[index] || `(unknown at index ${index})`;
            if (result.status === 'fulfilled' && result.value && result.value.filename && result.value.data) {
                // プロミスが成功し、かつ generateBarcodePDF が有効なオブジェクト ({filename, data}) を返した場合
                zip.file(result.value.filename, result.value.data);
                addedFileCount++;
            } else {
                // プロミスが失敗したか、generateBarcodePDF が null を返した場合
                failedBarcodes.push(originalBarcode);
                const reason = result.status === 'rejected'
                    ? (result.reason ? result.reason.message : '(No specific reason provided)')
                    : '(PDF generation function returned null or invalid data)';
                console.error(`[generateAndZip] PDF generation failed for '${originalBarcode}'. Status: ${result.status}. Reason: ${reason}`);
            }
        });

        if (failedBarcodes.length > 0) {
            console.warn(`[generateAndZip] Failed to generate PDF for ${failedBarcodes.length} barcodes: ${failedBarcodes.join(', ')}`);
            // 必要であればユーザーに警告を表示
            // alert(`警告: ${failedBarcodes.length}件のバーコード生成に失敗しました。\n${failedBarcodes.join(', ')}`);
        }

        if (addedFileCount > 0) {
            updateStatus(`ZIPファイルを生成中 (${addedFileCount}件)...`);
            const zipBlob = await zip.generateAsync({
                type: "blob",
                compression: "DEFLATE",
                compressionOptions: { level: 6 } // 圧縮レベル (0-9, 9が最高圧縮)
            });
            console.log("[generateAndZip] ZIP Blob generated.");

            let finalZipFilename = "barcodes.zip";
            if (zipFilenameInput && zipFilenameInput.value.trim()) {
                finalZipFilename = zipFilenameInput.value.trim();
                // 拡張子 .zip がなければ追加
                if (!finalZipFilename.toLowerCase().endsWith('.zip')) {
                    finalZipFilename += ".zip";
                }
            }
            console.log(`[generateAndZip] Determined ZIP filename: ${finalZipFilename}`);

            saveAs(zipBlob, finalZipFilename); // FileSaver.js の saveAs 関数でダウンロードをトリガー

            let successMessage = `完了: ${finalZipFilename} (${addedFileCount}件) をダウンロードしました。`;
            if (failedBarcodes.length > 0) {
                successMessage += ` (${failedBarcodes.length}件のエラーあり)`;
            }
            updateStatus(successMessage);

        } else {
            console.warn("[generateAndZip] No PDF files were successfully generated to add to ZIP.");
            if (failedBarcodes.length > 0) {
                 // 入力はあったが、すべて失敗した場合
                 updateStatus(`エラー: ${failedBarcodes.length}件すべてのバーコード生成に失敗しました。詳細はコンソールを確認してください。`);
            } else {
                 // 有効な入力がなかった、または予期せぬ理由でファイルが生成されなかった場合
                 updateStatus('エラー: ZIPに追加できる有効なバーコードファイルが生成されませんでした。');
            }
        }

    } catch (error) {
        // generateAndZip 関数自体の予期せぬエラー
        console.error("[generateAndZip] Unexpected error during the zipping process:", error);
        updateStatus(`致命的なエラーが発生しました: ${error.message}`);
    } finally {
        console.log("[generateAndZip] Processing finished.");
        disableUI(false); // 最後に必ずUIを有効化
        // 必要であれば数秒後にステータスメッセージをクリア
        // setTimeout(() => { updateStatus(''); }, 7000);
    }
}

// --- イベントリスナーの設定 ---
// ページ読み込み完了時にイベントリスナーを設定
document.addEventListener('DOMContentLoaded', () => {
    // ライブラリ存在チェック (任意)
    if (!window.jspdf) console.error("警告: jsPDFが見つかりません。");
    if (!window.Papa) console.error("警告: PapaParseが見つかりません。");
    if (!window.opentype) console.error("警告: opentype.jsが見つかりません。");
    if (!window.JSZip) console.error("警告: JSZipが見つかりません。");
    if (!window.saveAs) console.error("警告: FileSaver.js (saveAs)が見つかりません。");

    // ボタンやファイル入力へのイベントリスナー設定
    if (generateTextButton) {
        generateTextButton.addEventListener('click', handleTextarea);
    } else {
        console.error("Error: Button with ID 'generateFromTextButton' not found.");
    }

    if (csvFileInput) {
        csvFileInput.addEventListener('change', handleFile);
    } else {
        console.error("Error: File input with ID 'csvFile' not found.");
    }

    updateStatus("準備完了"); // 初期ステータス
});
