/**
 * ==============================================================================
 * 嘉義縣立永慶高級中學 - 品牌永慶・標配未來 2.0 (YCSH Future Standard 2.0)
 * Google Apps Script (GAS) 後端主程式 (Code.gs)
 * ==============================================================================
 * 
 * 【部署操作指南 / Deployment Instructions】
 * 1. 在 Google 雲端硬碟建立「Google Apps Script」專案。
 * 2. 將本檔案內容完整貼至專案預設的「程式碼.gs (Code.gs)」。
 * 3. 點選左側「+」號新增 HTML 檔案，命名為「index」，並貼上網頁 HTML 程式碼。
 * 4. 點選右上角「部署 (Deploy)」 > 「新建部署 (New deployment)」。
 * 5. 類型選擇「網頁應用程式 (Web app)」：
 *    - 說明 (Description)：永慶高中品牌宣傳網頁 2.0
 *    - 執行身分 (Execute as)：我 (Me)
 *    - 誰可以存取 (Who has access)：任何人 (Anyone)
 * 6. 點擊「部署」，即可獲得專屬公開網址！
 */

/**
 * 處理 GET 請求並動態渲染 HTML 網頁
 * 支援透過 URL 參數切換初始語系，例如：
 * ?lang=en (預設英文) | ?lang=zh (繁體中文) | ?lang=ja (日本語) | ?lang=ko (한국어)
 * 
 * @param {Object} e HTTP GET 事件參數物件
 * @returns {HtmlOutput} 評估並配置完成之 HTML 頁面
 */
function doGet(e) {
  try {
    // 建立 HTML 範本實例（GAS 專案內檔案名稱為 index 時，引數填 'index'）
    var template = HtmlService.createTemplateFromFile('index');
    
    // 讀取 URL 傳入的 lang 參數（若無則預設為英文 'en'）
    var requestedLang = (e && e.parameter && e.parameter.lang) ? e.parameter.lang.toLowerCase() : 'en';
    var validLangs = ['en', 'zh', 'ja', 'ko'];
    template.initialLang = (validLangs.indexOf(requestedLang) !== -1) ? requestedLang : 'en';
    
    // 渲染並設定網頁標頭（GAS addMetaTag 僅支援 'viewport'，勿加 'description'）
    return template.evaluate()
      .setTitle('Chiayi County Yung Ching High School | Future Standard 2.0')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    // 錯誤防護：若發生異常則提供乾淨的除錯訊息
    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; padding: 2rem; text-align: center;">' +
      '<h2>網頁載入發生狀況 / Page Loading Error</h2>' +
      '<p>請確認 GAS 專案中已新增名稱為 <code>index.html</code> 的 HTML 檔案。</p>' +
      '<p style="color: #666; font-size: 0.9em;">' + error.toString() + '</p>' +
      '</div>'
    ).setTitle('Error - YCSH Web App');
  }
}

/**
 * 模組化載入輔助函式
 * 供 index.html 內部使用 <?!= include('檔案名稱'); ?> 語法載入子組件或樣式
 * 
 * @param {string} filename 要載入的 HTML 檔案名稱 (不含 .html 副檔名)
 * @returns {string} 檔案原始字串內容
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    Logger.log('Include error for file ' + filename + ': ' + err.toString());
    return '<!-- Error loading ' + filename + ' -->';
  }
}

/**
 * 伺服器端資料介面：提供學校基本資訊查詢（可用於前端 google.script.run）
 * 
 * @returns {Object} 包含校名、聯絡資訊與重要連結之物件
 */
function getSchoolMetadata() {
  return {
    schoolNameZh: "嘉義縣立永慶高級中學",
    schoolNameEn: "Chiayi County Yung Ching High School",
    address: "61248 嘉義縣太保市信義二路 1 號",
    phone: "+886-5-362-7777",
    fax: "+886-5-362-7778",
    logoUrl: "https://lh3.googleusercontent.com/d/11oeojtlq1WpAm8mZ47LzemGM_sdodQL4",
    officialWebsite: "https://www.ycsh.cyc.edu.tw/"
  };
}