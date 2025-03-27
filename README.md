# YouTube 影片預覽工具

這是一個用於預覽和分析 YouTube 影片數據的工具。

## 功能特點

- 支援上傳 Excel 文件分析數據
- YouTube 影片預覽顯示
- 數據可視化展示
- 分享功能

## 已知問題和解決方案

### 1. Excel 欄位讀取問題

#### 問題描述
- A1 欄位內容沒有自動顯示在標題
- 第一欄的值顯示格式不正確
- 部分數值欄位（花費總和、曝光總和、CPM）沒有正確顯示

#### 解決方案

1. A1 標題顯示
```javascript
// 在讀取 Excel 文件時，獲取並設置 A1 的值為標題
const a1Value = firstSheet['A1'] ? firstSheet['A1'].v : '';
document.getElementById('pageTitle').textContent = a1Value;
```

2. 年月格式轉換
```javascript
// 自動將 6 位數字轉換為年月格式
if (/^\d{6}$/.test(String(firstColumnValue))) {
    const year = String(firstColumnValue).substring(0, 4);
    const month = String(firstColumnValue).substring(4, 6);
    formattedDate = `${year}年${month}月`;
}
```

3. 數值格式化
```javascript
// 改進數字格式化函數
function formatNumber(num) {
    if (num === null || num === undefined || isNaN(num)) return '0';
    const value = parseFloat(String(num).replace(/,/g, ''));
    if (isNaN(value)) return '0';
    return new Intl.NumberFormat('zh-TW').format(Math.round(value));
}
```

### 2. Excel 欄位映射

#### 正確的欄位對應
```
舊欄位 -> 新欄位
- 暗稿GG_An_安裝優化_YT版位_YT影片Top 3 (花費) -> 顯示A1欄位
- 素材名稱 -> 聯播網
- 花費總和 -> 花費總和（保持不變）
- 曝光總和 -> 曝光總和（保持不變）
- 安裝總和 -> 鉤子
- 安裝佔比 -> 風格
- IR -> IR（保持不變）
- CPI -> CPI（保持不變）
- 平均IR -> CPM
- 平均CPI -> 秒數
- url -> url（保持不變）
```

## 使用說明

1. Excel 文件格式要求：
   - 確保欄位名稱與上述對應表一致
   - A1 欄位將作為頁面標題顯示
   - 第一欄如果是 6 位數字（如：202401）將自動轉換為年月格式

2. 數據格式要求：
   - 花費總和、曝光總和、CPI、CPM：數字格式，會自動添加千分位
   - IR：數字格式，會自動轉換為百分比並保留 3 位小數
   - 鉤子、風格、秒數：文字格式
   - url：需要是有效的 YouTube 影片連結

## 注意事項

1. 數值處理：
   - 所有數字欄位會自動移除逗號後再處理
   - 無效的數值會顯示為 0
   - 百分比值會自動格式化

2. 顯示格式：
   - 數字會自動加上千分位分隔符
   - 百分比會保留 3 位小數
   - 年月格式會自動轉換（例：202401 -> 2024年01月）

## 開發者注意事項

1. 調試輸出：
   - 已添加 console.log 來追蹤數據處理過程
   - 可以通過瀏覽器控制台查看數據轉換過程

2. 錯誤處理：
   - 所有數據處理都有適當的錯誤處理機制
   - 無效數據會有預設值處理 