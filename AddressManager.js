/** AddressManager.gs
 *
 * 주소 변환:
 * - 입력: F4(블록 시작행) 문자열에서 지번까지(base)만 카카오 검색
 * - 결과: F4 = "지번 (도로명)"
 * - 추가: 지번 뒤 나머지(호/층 등) = F6에 저장
 * - 지도: F8에 Kakao map URL
 * - ✅ 완료/취소 블록은 스킵
 * - ✅ 연속 빈 블록 N개면 중단
 */

function updateAddressesBatch(isSilent) {
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return { summary: "데이터 없음", failedList: [] };

  var stopCtl = makeStopController_();

  var successCount = 0;
  var skipCount = 0;
  var failCount = 0;
  var failedList = [];

  var cleanKey = (KAKAO_API_KEY || "").toString().trim();

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    var nameVal = sheet.getRange(r + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(nameVal)) { continue; }

    if (isClosedBlock_(sheet, r)) { skipCount++; continue; }

    var noVal = sheet.getRange(r + CONFIG.POS_NO.row, CONFIG.POS_NO.col).getDisplayValue();
    var projectLabel = (noVal ? noVal : "") + " " + nameVal;

    var addrCell = sheet.getRange(r + CONFIG.POS_ADDR.row, CONFIG.POS_ADDR.col);         // F4
    var extraCell = sheet.getRange(r + CONFIG.POS_ADDR_EXTRA.row, CONFIG.POS_ADDR_EXTRA.col); // F6
    var rawAddress = (addrCell.getDisplayValue() || "").toString().trim();
    if (!rawAddress) continue;

    var mapCell = sheet.getRange(r + CONFIG.POS_MAP.row, CONFIG.POS_MAP.col); // F8
    var currentMapUrl = (mapCell.getDisplayValue() || "").toString().trim();

    // 이미 "지번 (도로명)" + 지도URL 있으면 스킵
    if (rawAddress.includes("(") && currentMapUrl) {
      skipCount++;
      continue;
    }

    // ✅ 지번까지만 검색 + extra 추출
    var parts = splitAddressExtra_(rawAddress);
    var baseQuery = (parts.base || "").trim();
    var extra = (parts.extra || "").trim();

    if (!baseQuery) continue;

    try {
      var url = "https://dapi.kakao.com/v2/local/search/address.json?query=" + encodeURIComponent(baseQuery);
      var response = UrlFetchApp.fetch(url, {
        headers: { "Authorization": "KakaoAK " + cleanKey },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.documents && json.documents.length > 0) {
          var doc = json.documents[0];

          var jibun = cleanPrefix_(doc.address ? doc.address.address_name : (doc.address_name || ""));
          var road = doc.road_address
            ? cleanPrefix_(doc.road_address.address_name).replace(/^[가-힣]+구\s+/, "")
            : "";

          var finalAddr = jibun + (road ? " (" + road + ")" : "");
          addrCell.setValue(finalAddr);

          // extra는 기존 값이 비어있을 때만 채움(수동 편집 보호)
          if (extra && (extraCell.getDisplayValue() || "").toString().trim() === "") {
            extraCell.setValue(extra);
          }

          mapCell.setValue("https://map.kakao.com/?q=" + encodeURIComponent(finalAddr));
          successCount++;
        } else {
          failCount++;
          failedList.push(projectLabel + " (주소 불분명)");
        }
      } else {
        failCount++;
        failedList.push(projectLabel + " (API 응답코드 " + response.getResponseCode() + ")");
      }
    } catch (e) {
      failCount++;
      failedList.push(projectLabel + " (시스템 오류: " + e.message + ")");
    }
  }

  if (!isSilent) SpreadsheetApp.getUi().alert("✅ 주소 변환 완료");
  return {
    summary: "신규 " + successCount + "건 / 이미완료 " + skipCount + "건 / 실패 " + failCount + "건",
    failedList: failedList
  };
}
