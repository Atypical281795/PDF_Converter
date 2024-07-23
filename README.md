# PDF Converter

這是一個使用 Python 開發的 PDF 轉換工具，透過GUI點選button，將 PDF 文件轉換為 PNG、JPG、SVG 或 Word 格式。
=======
如果要直接使用，則安裝後前往路徑==>...\PDF_Converter\dist\PDF_Converter\執行PDF_Converter.exe即可
## 功能

- 支持將 PDF 轉換為 PNG、JPG、SVG 和 Word 格式
- 直觀的圖形用戶界面
- 支持拖放 PDF 文件
- 自動創建輸出文件夾

## 系統要求

- Python 3.6 或更高版本
- Windows、macOS 或 Linux 操作系統

## 安裝步驟

1. 克隆此儲存庫或下載源代碼：
   git clone https:github.com/Atypical281795/PDF_Converter.git
2. 創建並激活虛擬環境（可選但推薦）：
   python -m venv venv
   source venv/bin/activate #在windows上使用venv\Scripts\activate
3. 安裝所需的依賴項：
   pip install -r requirements.txt
4. 安裝 Poppler（pdf2image 的依賴項）：
- Windows：下載 Poppler 二進制文件並將其添加到系統路徑
- macOS：使用 Homebrew 安裝 `brew install poppler`
- Linux：使用包管理器安裝，例如 `sudo apt-get install poppler-utils`

## 使用說明

1. 運行程序：
   python GUI_PdfConvert.py
2. 在打開的 GUI 窗口中，通過拖放或點擊"瀏覽"按鈕選擇 PDF 文件。

3. 選擇所需的輸出格式（PNG、JPG、SVG 或 Word）。

4. 點擊"轉換"按鈕開始轉換過程。

5. 轉換完成後，輸出窗口將顯示結果和保存位置。

## 注意事項

- 轉換大型 PDF 文件可能需要一些時間，請耐心等待。
- 確保您有足夠的磁盤空間來保存轉換後的文件。
- 某些複雜的 PDF 可能無法完美轉換，特別是在轉換為 Word 格式時。

## 疑難解答

如果您在運行程序時遇到問題：

1. 確保所有依賴項都已正確安裝。
2. 檢查 Python 版本是否兼容。
3. 確保 Poppler 已正確安裝並添加到系統路徑。

如果問題仍然存在，請提交一個 issue，並附上錯誤消息和系統信息。

## 貢獻

歡迎提交 pull requests。對於重大更改，請先開一個 issue 討論您想要改變的內容。

## 許可證

[MIT](https://choosealicense.com/licenses/mit/)
