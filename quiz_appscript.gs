/**
 * 퀴즈 결과 저장 스크립트 (5지선다용)
 *
 * 사용법: 결과를 모을 구글 스프레드시트를 하나 만들고,
 *        [확장 프로그램] → [Apps Script] 에서 이 코드를 통째로 붙여넣은 뒤 배포하세요.
 *        (자세한 순서는 가이드 3단계 참고)
 *
 * 시트는 자동으로 두 개 생깁니다.
 *   - "요약"   : 제출 1건 = 1줄 (이름/점수/등급)
 *   - "문항별" : 문항마다 1줄 (무엇을 골라 맞았는지/틀렸는지)
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = e.parameter;

  var name    = p.name    || '이름없음';
  var week    = p.week    || '-';           // 퀴즈 이름(QUIZ_ID)
  var correct = p.correct || '0';
  var total   = p.total   || '0';
  var pct     = parseInt(p.pct || '0');
  var grade   = p.grade   || '-';
  var sel     = (p.sel     || '').split(',');   // 학생이 고른 답 번호들
  var cor     = (p.cor     || '').split(',');   // 정답 번호들
  var qtexts  = (p.qtexts  || '').split('|');   // 문항 요약
  var optexts = (p.optexts || '').split('|');   // 고른 보기 요약
  var now = new Date();

  // ── 시트1: 요약 ──────────────────────────────────
  var summary = ss.getSheetByName('요약');
  if (!summary) {
    summary = ss.insertSheet('요약');
    summary.appendRow(['제출시간', '이름', '퀴즈', '점수(%)', '정답수', '전체문항', '등급']);
    summary.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#ccfbf1');
  }
  summary.appendRow([now, name, week, pct, correct, total, grade]);
  var sRow = summary.getLastRow();
  var color = pct >= 80 ? '#dcfce7' : (pct >= 60 ? '#fef3c7' : '#fee2e2');
  summary.getRange(sRow, 1, 1, 7).setBackground(color);

  // ── 시트2: 문항별 ───────────────────────────────
  var detail = ss.getSheetByName('문항별');
  if (!detail) {
    detail = ss.insertSheet('문항별');
    detail.appendRow(['제출시간', '이름', '퀴즈', '문항', '질문(요약)', '고른답', '정답', '정오']);
    detail.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#ccfbf1');
  }
  var marks = ['①', '②', '③', '④', '⑤', '⑥'];
  for (var i = 0; i < parseInt(total); i++) {
    var selIdx = parseInt(sel[i]);
    var corIdx = parseInt(cor[i]);
    var ok = selIdx === corIdx;
    var selMark = (!isNaN(selIdx) && selIdx >= 0) ? marks[selIdx] : '?';
    var corMark = (!isNaN(corIdx) && corIdx >= 0) ? marks[corIdx] : '?';
    detail.appendRow([
      now, name, week, 'Q' + (i + 1),
      qtexts[i] || '',
      selMark + ' ' + (optexts[i] || '-'),
      corMark,
      ok ? '⭕' : '❌'
    ]);
    var dRow = detail.getLastRow();
    detail.getRange(dRow, 8).setBackground(ok ? '#dcfce7' : '#fee2e2');
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
