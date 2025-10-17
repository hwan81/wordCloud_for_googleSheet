/**
 * @fileoverview Code.gs
 * Google Sheets URL을 받아 B열 데이터를 처리하고 단어 빈도수를 계산합니다.
 */

function doGet() {
  // index.html 파일을 템플릿으로 로드하여 웹 앱으로 제공합니다.
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Google Sheets 워드클라우드 데이터 처리');
}

/**
 * Google Sheets URL에서 데이터를 읽고 단어 빈도수를 계산합니다.
 * @param {string} url - Google Sheets의 URL
 * @returns {object} - 시트 제목과 단어 빈도수 데이터 또는 오류 메시지
 */
function processSheetData(url) {
  // 1. URL에서 시트 ID 추출
  const sheetIdMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!sheetIdMatch || sheetIdMatch.length < 2) {
    return { error: "유효하지 않은 시트 URL입니다. URL 형식이 올바른지 확인하십시오." };
  }
  const sheetId = sheetIdMatch[1];

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    // 첫 번째 시트를 가져옵니다.
    const sheet = ss.getSheets()[0];
    
    // 2. B1 셀에서 제목 읽기
    const title = sheet.getRange('B1').getValue();
    
    // 3. B2부터 마지막 행까지의 데이터 읽기
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { error: "시트에 B2 셀 이후에 분석할 텍스트 데이터가 없습니다." };
    }
    
    // B열, 2행부터 마지막 행까지 (lastRow - 1 개 행, 1개 열)
    const range = sheet.getRange(2, 2, lastRow - 1, 1); 
    const values = range.getValues().flat(); // 2차원 배열을 1차원 배열로 변환
    
    // 4. 단어 빈도수 계산
    const text = values.join(' ');
    
    // 간단한 전처리: 소문자 변환, 구두점 제거, 다중 공백 단일 공백으로
    const cleanedText = text
      .toLowerCase()
      .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()'"“”‘’\r\n]/g, ' ') 
      .replace(/\s+/g, ' ') 
      .trim();

    // 단어 분리 및 한 글자 단어 제외
    const words = cleanedText.split(' ').filter(word => word.length > 1); 
    
    const wordFrequency = {};
    words.forEach(word => {
      wordFrequency[word] = (wordFrequency[word] || 0) + 1;
    });
    
    // 결과 배열로 변환 및 빈도수 기준 내림차순 정렬
    const frequencyData = Object.entries(wordFrequency)
      .map(([word, count]) => ({ item: word, frequency: count }))
      .sort((a, b) => b.frequency - a.frequency);
      
    return { title: title, frequencyData: frequencyData };

  } catch (e) {
    // 시트 접근 오류 또는 기타 오류 처리 (예: 권한 문제)
    return { error: `시트 처리 중 오류가 발생했습니다: ${e.message}. 시트가 '링크가 있는 모든 사용자에게 보기' 권한으로 공유되었는지 확인하십시오.` };
  }
}
