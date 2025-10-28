// ========================================
// 📊 스마트워치 데이터 수집 프로젝트
// Google Apps Script - 양방향 연동 백엔드
// ========================================

// ========================================
// 웹앱 GET 요청 핸들러 (페이지 제공)
// ========================================
function doGet(e) {
  var page = e.parameter.page || "participant";

  if (page === "manager") {
    // 관리자 대시보드 데이터 API
    return getManagerData();
  } else {
    // 참가자 페이지 - POST 요청으로 처리
    return ContentService.createTextOutput(
      JSON.stringify({
        success: true,
        message: "참가자 페이지입니다. POST 요청으로 로그인해주세요.",
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========================================
// 웹앱 POST 요청 핸들러 (상태 업데이트)
// ========================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    // 로그 기록
    Logger.log("POST 요청 받음 - Action: " + action);
    Logger.log("데이터: " + JSON.stringify(data));

    if (action === "updateStatus") {
      return updateParticipantStatus(data);
    } else if (action === "getParticipant") {
      return getParticipantInfo(data.id, data.name);
    } else if (action === "updateMeasureDate") {
      return updateMeasureDate(data);
    }

    Logger.log("알 수 없는 요청: " + action);
    return createResponse(false, "알 수 없는 요청입니다: " + action);
  } catch (error) {
    Logger.log("에러 발생: " + error.message);
    Logger.log("스택: " + error.stack);
    return createResponse(false, "오류가 발생했습니다: " + error.message);
  }
}

// ========================================
// 참가자 정보 조회
// ========================================
function getParticipantInfo(participantId, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("진행현황");
  var data = sheet.getDataRange().getValues();

  // 참가자 찾기
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === participantId && data[i][1] === name) {
      var participant = {
        id: data[i][0],
        name: data[i][1],
        device: data[i][2] || "-",
        status: data[i][3] || "대기중",
        shipDate: formatDate(data[i][4]),
        receiveDate: formatDate(data[i][5]),
        syncDate: formatDate(data[i][6]),
        collectStartDate: formatDate(data[i][7]),
        daysElapsed: calculateDaysElapsed(data[i][4]),
        collectDays: calculateCollectDays(data[i][7]),
        pickupDate: formatDate(data[i][9]) || "조율 중", // I→I 위치 변경 없음 (수집종료예정일 제거로 한 칸 앞당겨짐)
      };

      return createResponse(true, "정보를 가져왔습니다.", participant);
    }
  }

  return createResponse(
    false,
    "참가자 정보를 찾을 수 없습니다. ID와 이름을 확인해주세요."
  );
}

// ========================================
// 참가자 상태 업데이트
// ========================================
function updateParticipantStatus(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("진행현황");
  var sheetData = sheet.getDataRange().getValues();

  // 참가자 찾기
  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] === data.id) {
      var today = new Date();
      var row = i + 1;

      // 상태 업데이트
      sheet.getRange(row, 4).setValue(data.status); // D열 (현재상태)

      // 날짜 자동 기록
      if (data.status === "수령완료") {
        sheet.getRange(row, 6).setValue(today); // F열 (수령일)
        // 알림: 연동 가이드 전송 필요
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "택배를 수령했습니다."
        );
      } else if (data.status === "수집중") {
        sheet.getRange(row, 7).setValue(today); // G열 (연동완료일)
        // H열 (수집시작일)은 사용자가 직접 선택하도록 자동 입력 제거

        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "기기 연동을 완료했습니다. 측정 예정일을 선택해주세요."
        );
      } else if (data.status === "수집완료") {
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "측정을 완료했습니다. 데이터 확인 단계입니다."
        );
      } else if (data.status === "데이터확인완료") {
        sheet.getRange(row, 9).setValue(today); // I열 (데이터확인일) - J→I 이동
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "데이터 확인을 완료했습니다. 설문 작성 대기 중입니다."
        );
      } else if (data.status === "설문완료") {
        sheet.getRange(row, 10).setValue(today); // J열 (설문완료일) - K→J 이동
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "설문을 완료했습니다. 데이터 제출 대기 중입니다."
        );
      } else if (data.status === "데이터제출완료") {
        sheet.getRange(row, 11).setValue(today); // K열 (데이터제출일) - L→K 이동
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "데이터를 제출했습니다. 매니저 확인이 필요합니다."
        );
      } else if (data.status === "매니저확인완료") {
        sheet.getRange(row, 12).setValue(today); // L열 (매니저확인일) - M→L 이동
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "매니저가 데이터를 확인 완료했습니다. 기기 반납 단계로 진행합니다."
        );
      } else if (data.status === "회수대기") {
        sendNotificationToManager(
          data.id,
          sheetData[i][1],
          "반납 준비가 완료되었습니다. 택배 회수 일정을 안내해주세요."
        );
      }

      // 로그 기록
      logAction(
        data.id,
        sheetData[i][1],
        "상태 변경: " + data.status,
        "참가자"
      );

      return createResponse(true, "상태가 업데이트되었습니다.");
    }
  }

  return createResponse(false, "참가자를 찾을 수 없습니다.");
}

// ========================================
// 관리자 대시보드 데이터
// ========================================
function getManagerData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("진행현황");
  var data = sheet.getDataRange().getValues();

  var participants = [];
  var stats = {
    total: 150,
    completed: 0,
    inProgress: 0,
    urgent: 0,
    availableDevices: 10,
  };

  for (var i = 1; i < data.length; i++) {
    var participant = {
      id: data[i][0],
      name: data[i][1],
      device: data[i][2],
      status: data[i][3],
      priority: data[i][12] || "정상",
      daysElapsed: calculateDaysElapsed(data[i][4]),
      action: data[i][11] || "-",
    };

    participants.push(participant);

    // 통계 계산
    if (data[i][3] === "회수완료") stats.completed++;
    if (data[i][3] === "수집중") stats.inProgress++;
    if (data[i][12] === "긴급") stats.urgent++;
  }

  stats.availableDevices = 10 - stats.inProgress;

  // 기기 현황 데이터 가져오기
  var devices = getDeviceStatus();

  var result = {
    participants: participants,
    stats: stats,
    devices: devices,
    lastUpdate: new Date().toISOString(),
  };

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ========================================
// 기기 현황 조회
// ========================================
function getDeviceStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deviceSheet = ss.getSheetByName("기기현황");

  // 기기현황 시트가 없으면 생성
  if (!deviceSheet) {
    deviceSheet = ss.insertSheet("기기현황");
    deviceSheet.appendRow([
      "기기번호",
      "현재상태",
      "현재사용자",
      "사용시작일",
      "예상반환일",
    ]);

    // 기본 10대 기기 데이터 생성
    for (var i = 1; i <= 10; i++) {
      deviceSheet.appendRow([i, "사용가능", "-", "", ""]);
    }
  }

  var deviceData = deviceSheet.getDataRange().getValues();
  var devices = [];

  for (var i = 1; i < deviceData.length; i++) {
    var device = {
      number: deviceData[i][0],
      status: deviceData[i][1] || "사용가능",
      currentUser: deviceData[i][2] || "-",
      startDate: formatDate(deviceData[i][3]) || "-",
      returnDate: formatDate(deviceData[i][4]) || "-",
    };
    devices.push(device);
  }

  return devices;
}

// ========================================
// 관리자에게 알림 전송
// ========================================
function sendNotificationToManager(participantId, participantName, message) {
  var 수신자 = "your-email@example.com"; // 관리자 이메일

  var 제목 =
    "[참가자 업데이트] " + participantName + " (" + participantId + ")";
  var 본문 = '<html><body style="font-family: Arial, sans-serif;">';
  본문 += '<h2 style="color: #667eea;">📢 참가자 상태 업데이트</h2>';
  본문 +=
    "<p><strong>참가자:</strong> " +
    participantName +
    " (" +
    participantId +
    ")</p>";
  본문 += "<p><strong>업데이트:</strong> " + message + "</p>";
  본문 +=
    "<p><strong>시간:</strong> " + new Date().toLocaleString("ko-KR") + "</p>";
  본문 += "<hr>";
  본문 +=
    '<p><a href="' +
    SpreadsheetApp.getActiveSpreadsheet().getUrl() +
    '" style="background: #667eea; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">스프레드시트 확인하기</a></p>';
  본문 += "</body></html>";

  try {
    MailApp.sendEmail({
      to: 수신자,
      subject: 제목,
      htmlBody: 본문,
    });
  } catch (error) {
    Logger.log("메일 전송 실패: " + error.message);
  }
}

// ========================================
// 액션 로그 기록
// ========================================
function logAction(participantId, participantName, action, actor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("액션로그") || ss.insertSheet("액션로그");

  // 헤더가 없으면 생성
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "날짜",
      "시간",
      "참가자ID",
      "참가자명",
      "액션",
      "실행자",
    ]);
  }

  var now = new Date();
  logSheet.appendRow([
    Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd"),
    Utilities.formatDate(now, "GMT+9", "HH:mm:ss"),
    participantId,
    participantName,
    action,
    actor,
  ]);
}

// ========================================
// 유틸리티: 경과 일수 계산
// ========================================
function calculateDaysElapsed(startDate) {
  if (!startDate) return 0;

  var today = new Date();
  var start = new Date(startDate);
  var diff = today - start;
  var days = Math.floor(diff / (1000 * 60 * 60 * 24));

  return days >= 0 ? days : 0;
}

// ========================================
// 측정일 업데이트
// ========================================
function updateMeasureDate(data) {
  try {
    Logger.log("=== 측정일 업데이트 시작 ===");
    Logger.log("요청 데이터: " + JSON.stringify(data));
    Logger.log("ID 타입: " + typeof data.id + ", 값: " + data.id);
    Logger.log("Date 타입: " + typeof data.date + ", 값: " + data.date);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("진행현황");

    if (!sheet) {
      Logger.log("❌ 진행현황 시트를 찾을 수 없습니다.");
      return createResponse(false, "진행현황 시트를 찾을 수 없습니다.");
    }

    var sheetData = sheet.getDataRange().getValues();
    Logger.log("✓ 시트 데이터 행 수: " + sheetData.length);

    // 참가자 찾기
    for (var i = 1; i < sheetData.length; i++) {
      var sheetId = String(sheetData[i][0]).trim();
      var requestId = String(data.id).trim();

      Logger.log(
        "비교 [" +
          i +
          "] - 시트 ID: '" +
          sheetId +
          "' vs 요청 ID: '" +
          requestId +
          "'"
      );

      if (sheetId === requestId) {
        var row = i + 1;
        Logger.log("✓ 참가자 발견! 행 번호: " + row);

        // 날짜 파싱 및 검증
        var measureDate;
        try {
          measureDate = new Date(data.date);
          Logger.log("✓ 날짜 파싱 성공: " + measureDate);

          if (isNaN(measureDate.getTime())) {
            Logger.log("❌ 유효하지 않은 날짜");
            return createResponse(false, "유효하지 않은 날짜 형식입니다.");
          }
        } catch (dateError) {
          Logger.log("❌ 날짜 파싱 에러: " + dateError.message);
          return createResponse(
            false,
            "날짜를 처리할 수 없습니다: " + dateError.message
          );
        }

        // 기존값 확인
        var currentValue = sheet.getRange(row, 8).getValue();
        Logger.log("현재 H열 값: " + currentValue);

        // H열(수집시작일)에 측정 예정일 저장
        try {
          sheet.getRange(row, 8).setValue(measureDate);
          Logger.log("✓ H열에 측정일 저장 완료");
        } catch (saveError) {
          Logger.log("❌ 저장 에러: " + saveError.message);
          return createResponse(false, "저장 중 오류: " + saveError.message);
        }

        // 저장 후 확인
        var newValue = sheet.getRange(row, 8).getValue();
        Logger.log("저장 후 H열 값: " + newValue);

        // 로그 기록
        try {
          logAction(
            data.id,
            sheetData[i][1],
            "측정 예정일 설정: " + data.date,
            "참가자"
          );
          Logger.log("✓ 로그 기록 완료");
        } catch (logError) {
          Logger.log("⚠️ 로그 기록 실패 (진행 계속): " + logError.message);
        }

        Logger.log("=== 측정일 업데이트 성공 완료 ===");
        return createResponse(true, "측정 예정일이 설정되었습니다.");
      }
    }

    Logger.log("❌ 참가자를 찾을 수 없습니다: " + data.id);
    Logger.log(
      "시트에 있는 모든 ID: " +
        sheetData
          .slice(1)
          .map(function (row) {
            return row[0];
          })
          .join(", ")
    );
    return createResponse(false, "참가자를 찾을 수 없습니다: " + data.id);
  } catch (error) {
    Logger.log("❌ 측정일 업데이트 치명적 에러: " + error.message);
    Logger.log("에러 스택: " + error.stack);
    return createResponse(false, "측정일 업데이트 중 오류: " + error.message);
  }
}

// ========================================
// 유틸리티: 수집 일수 계산
// ========================================
function calculateCollectDays(collectStartDate) {
  if (!collectStartDate) return 0;

  var today = new Date();
  var start = new Date(collectStartDate);

  // 날짜만 비교 (시간 제외)
  today.setHours(0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);

  var diff = today - start;
  var days = Math.floor(diff / (1000 * 60 * 60 * 24));

  // 측정일이 오늘이거나 지났으면 1 (완료), 아니면 0 (대기중)
  return days >= 0 ? 1 : 0;
}

// ========================================
// 유틸리티: 날짜 포맷
// ========================================
function formatDate(date) {
  if (!date) return null;

  try {
    return Utilities.formatDate(new Date(date), "GMT+9", "yyyy-MM-dd");
  } catch (error) {
    return null;
  }
}

// ========================================
// 유틸리티: 응답 생성
// ========================================
function createResponse(success, message, data) {
  var response = {
    success: success,
    message: message,
    data: data || null,
  };

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ========================================
// 일일 자동 체크 (기존 함수 유지)
// ========================================
function 일일체크() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("진행현황");
  var data = sheet.getDataRange().getValues();

  var 긴급알림목록 = [];
  var 주의알림목록 = [];

  for (var i = 1; i < data.length; i++) {
    var 참가자ID = data[i][0];
    var 이름 = data[i][1];
    var 현재상태 = data[i][3];
    var 발송일 = data[i][4];
    var 경과일수 = calculateDaysElapsed(발송일);

    // 발송 후 2일 경과 - 수령 확인 필요
    if (현재상태 === "발송완료" && 경과일수 >= 2) {
      긴급알림목록.push(
        `🚨 ${이름} (${참가자ID}) - 택배 수령 확인 필요 (발송 후 ${경과일수}일)`
      );
      sheet.getRange(i + 1, 4).setValue("수령확인필요");
      sheet.getRange(i + 1, 13).setValue("긴급");
    }

    // 측정 완료 확인 - 자동으로 수집완료로 변경하지 않음 (참가자가 직접 업데이트)
    if (현재상태 === "수집중" && 경과일수 >= 3) {
      주의알림목록.push(
        `📦 ${이름} (${참가자ID}) - 측정 예정일로부터 ${경과일수}일 경과. 참가자 확인 필요`
      );
    }
  }

  if (긴급알림목록.length > 0 || 주의알림목록.length > 0) {
    알림메일발송(긴급알림목록, 주의알림목록);
  }
}

// ========================================
// 알림 메일 발송 (기존 함수)
// ========================================
function 알림메일발송(긴급목록, 주의목록) {
  var 수신자 = "your-email@example.com";
  var 제목 =
    "[데이터수집] 일일 액션 리포트 - " +
    Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");

  var 본문 = '<html><body style="font-family: Arial, sans-serif;">';
  본문 += '<h2 style="color: #667eea;">📊 오늘의 액션 리포트</h2>';

  if (긴급목록.length > 0) {
    본문 +=
      '<h3 style="color: #ff6b6b;">🚨 긴급 액션 (' +
      긴급목록.length +
      "건)</h3>";
    본문 += "<ul>";
    긴급목록.forEach(function (item) {
      본문 += '<li style="margin: 10px 0;">' + item + "</li>";
    });
    본문 += "</ul>";
  }

  if (주의목록.length > 0) {
    본문 +=
      '<h3 style="color: #ffa500;">⚠️ 주의 액션 (' +
      주의목록.length +
      "건)</h3>";
    본문 += "<ul>";
    주의목록.forEach(function (item) {
      본문 += '<li style="margin: 10px 0;">' + item + "</li>";
    });
    본문 += "</ul>";
  }

  본문 += "<hr>";
  본문 +=
    '<p><a href="' +
    SpreadsheetApp.getActiveSpreadsheet().getUrl() +
    '" style="background: #667eea; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">스프레드시트 열기</a></p>';
  본문 += "</body></html>";

  MailApp.sendEmail({
    to: 수신자,
    subject: 제목,
    htmlBody: 본문,
  });
}

// ========================================
// 트리거 설정
// ========================================
function 트리거설정() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger("일일체크").timeBased().atHour(9).everyDays(1).create();

  Logger.log("트리거 설정 완료");
}

// ========================================
// 참가자에게 개인 링크 전송 (선택 기능)
// ========================================
function 참가자링크전송() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("참가자마스터");
  var data = sheet.getDataRange().getValues();

  // 웹앱 URL (배포 후 여기에 입력)
  var webAppUrl = "YOUR_WEB_APP_URL_HERE";

  for (var i = 1; i < data.length; i++) {
    var 참가자ID = data[i][0];
    var 이름 = data[i][1];
    var 연락처 = data[i][2];
    var 이메일 = data[i][3];

    if (!이메일) continue;

    // 개인 링크 생성
    var 개인링크 =
      webAppUrl + "?id=" + 참가자ID + "&name=" + encodeURIComponent(이름);

    // 이메일 전송
    var 제목 = "[수면 데이터 수집] 참가자 페이지 안내";
    var 본문 = '<html><body style="font-family: Arial, sans-serif;">';
    본문 += "<h2>안녕하세요, " + 이름 + "님!</h2>";
    본문 += "<p>수면 데이터 수집 프로젝트에 참여해주셔서 감사합니다.</p>";
    본문 +=
      "<p>아래 링크에서 진행 현황을 확인하고 상태를 업데이트하실 수 있습니다:</p>";
    본문 +=
      '<p><a href="' +
      개인링크 +
      '" style="background: #667eea; color: white; padding: 15px 30px; text-decoration: none; border-radius: 10px; display: inline-block; margin: 20px 0;">내 진행 현황 보기</a></p>';
    본문 +=
      '<p style="color: #666; font-size: 14px;">참가자 ID: <strong>' +
      참가자ID +
      "</strong></p>";
    본문 += "<p>문의사항이 있으시면 언제든 연락주세요!</p>";
    본문 += "</body></html>";

    try {
      MailApp.sendEmail({
        to: 이메일,
        subject: 제목,
        htmlBody: 본문,
      });

      Logger.log("링크 전송 완료: " + 이름);
      Utilities.sleep(1000); // API 제한 방지
    } catch (error) {
      Logger.log("전송 실패: " + 이름 + " - " + error.message);
    }
  }
}

// ========================================
// 사용 가이드:
//
// 1. Apps Script 편집기에서 이 코드 붙여넣기
// 2. '배포' > '새 배포' 클릭
// 3. 유형: '웹 앱' 선택
// 4. 실행 사용자: '나'
// 5. 액세스 권한: '모든 사용자'
// 6. 배포 후 웹 앱 URL 복사
// 7. participant_interface.html의 SCRIPT_URL에 URL 붙여넣기
// 8. 관리자 이메일 주소 변경
// 9. 트리거설정() 실행
// ========================================
