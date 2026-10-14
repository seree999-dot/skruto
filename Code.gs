/**
 * =====================================================================================
 * ระบบการขอเดินทางไปราชการของหน่วยงาน (Official Travel Request System)
 * Backend Script สำหรับ Google Apps Script เชื่อมต่อ Google Sheets
 * =====================================================================================
 */

const SHEET_REQUESTS = "Requests";
const SHEET_REPORTS = "TravelReports";
const SHEET_CONFIG = "SystemConfig";

/**
 * ฟังก์ชันเริ่มต้นสร้างชีตและหัวตารางอัตโนมัติ (รันครั้งแรกเพียงครั้งเดียว)
 */
function setupInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. ตารางคำขอไปราชการ (Requests)
  let sheetReq = ss.getSheetByName(SHEET_REQUESTS);
  if (!sheetReq) {
    sheetReq = ss.insertSheet(SHEET_REQUESTS);
  }
  const reqHeaders = [
    "RequestID", "CreatedAt", "FullName", "Position", "Department",
    "Phone", "Email", "PurposeCategory", "Topic", "Destination",
    "Province", "StartDate", "EndDate", "TotalDays", "Companions",
    "TransportType", "VehiclePlate", "BudgetType", "EstimatedBudget",
    "DocUrl", "Status", "ApproverName", "ApprovalDate", "ApprovalComment"
  ];
  sheetReq.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
  sheetReq.getRange(1, 1, 1, reqHeaders.length).setFontWeight("bold").setBackground("#e0f2fe");

  // 2. ตารางรายงานผลหลังกลับ (TravelReports)
  let sheetRep = ss.getSheetByName(SHEET_REPORTS);
  if (!sheetRep) {
    sheetRep = ss.insertSheet(SHEET_REPORTS);
  }
  const repHeaders = [
    "ReportID", "RequestID", "ReportDate", "SummaryOutcome",
    "Benefits", "ActualExpense", "AttachmentUrl", "ReporterName"
  ];
  sheetRep.getRange(1, 1, 1, repHeaders.length).setValues([repHeaders]);
  sheetRep.getRange(1, 1, 1, repHeaders.length).setFontWeight("bold").setBackground("#dcfce7");

  // 3. ตารางการตั้งค่าระบบ (SystemConfig)
  let sheetCfg = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetCfg) {
    sheetCfg = ss.insertSheet(SHEET_CONFIG);
    const cfgHeaders = ["ConfigKey", "ConfigValue"];
    sheetCfg.getRange(1, 1, 1, cfgHeaders.length).setValues([cfgHeaders]);
    sheetCfg.getRange(1, 1, 1, cfgHeaders.length).setFontWeight("bold").setBackground("#fef3c7");

    const defaultConfigs = [
      ["AdminPIN", "1234"],
      ["OrgName", "สำนักงานส่วนกลางและบริการเทคโนโลยี"],
      ["Departments", "กลุ่มบริหารทั่วไป,กลุ่มแผนงานและงบประมาณ,กลุ่มส่งเสริมและพัฒนา,กลุ่มเทคโนโลยีสารสนเทศ,กลุ่มตรวจสอบภายใน"],
      ["Vehicles", "กข-1234 กทม. (รถตู้ส่วนกลาง),ฮภ-5678 กทม. (รถเก๋งประจำกอง),นข-9999 กทม. (รถกระบะตรวจงาน)"]
    ];
    sheetCfg.getRange(2, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
  }

  Logger.log("ติดตั้งตารางข้อมูลและโครงสร้างเรียบร้อยแล้ว");
}

/**
 * จัดการ HTTP GET Request
 */
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "getData";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "ping") {
      return responseJSON({ status: "success", message: "GAS Web App API online" });
    }

    // ดึงข้อมูลทั้งหมด
    const requests = getSheetData(ss.getSheetByName(SHEET_REQUESTS));
    const reports = getSheetData(ss.getSheetByName(SHEET_REPORTS));
    const configData = getSheetData(ss.getSheetByName(SHEET_CONFIG));

    const config = {};
    if (configData && configData.length > 0) {
      configData.forEach(row => {
        if (row.ConfigKey) {
          config[row.ConfigKey] = row.ConfigValue;
        }
      });
    }

    return responseJSON({
      status: "success",
      requests: requests || [],
      reports: reports || [],
      config: config
    });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

/**
 * จัดการ HTTP POST Request
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. สร้างคำขอใหม่ (Create Request)
    if (action === "createRequest") {
      const sheet = ss.getSheetByName(SHEET_REQUESTS);
      if (!sheet) throw new Error("ไม่พบแผ่นงาน " + SHEET_REQUESTS);

      const now = new Date();
      const yearMonth = Utilities.formatDate(now, "Asia/Bangkok", "yyyyMM");
      const lastRow = sheet.getLastRow();
      const runningNum = ("000" + (lastRow)).slice(-4);
      const requestId = "TR-" + yearMonth + "-" + runningNum;

      const newRow = [
        requestId,
        Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss"),
        payload.fullName || "",
        payload.position || "",
        payload.department || "",
        payload.phone || "",
        payload.email || "",
        payload.purposeCategory || "",
        payload.topic || "",
        payload.destination || "",
        payload.province || "",
        payload.startDate || "",
        payload.endDate || "",
        payload.totalDays || 1,
        payload.companions || "",
        payload.transportType || "",
        payload.vehiclePlate || "",
        payload.budgetType || "",
        payload.estimatedBudget || 0,
        payload.docUrl || "",
        "รอพิจารณา", // Status เริ่มต้น
        "", // ApproverName
        "", // ApprovalDate
        ""  // ApprovalComment
      ];

      sheet.appendRow(newRow);
      return responseJSON({ status: "success", requestId: requestId, message: "บันทึกคำขอเรียบร้อยแล้ว" });
    }

    // 2. อนุมัติ / ปฏิเสธคำขอ (Approve / Reject Request)
    if (action === "updateStatus") {
      const sheet = ss.getSheetByName(SHEET_REQUESTS);
      const data = sheet.getDataRange().getValues();
      const requestId = payload.requestId;
      let rowIndex = -1;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === requestId) {
          rowIndex = i + 1; // 1-indexed in Sheets
          break;
        }
      }

      if (rowIndex === -1) {
        return responseJSON({ status: "error", message: "ไม่พบรหัสคำขอ: " + requestId });
      }

      const now = new Date();
      const approvalDateStr = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

      // คอลัมน์: U = Status(21), V = ApproverName(22), W = ApprovalDate(23), X = ApprovalComment(24)
      sheet.getRange(rowIndex, 21).setValue(payload.status || "อนุมัติแล้ว");
      sheet.getRange(rowIndex, 22).setValue(payload.approverName || "ผู้อนุมัติ");
      sheet.getRange(rowIndex, 23).setValue(approvalDateStr);
      sheet.getRange(rowIndex, 24).setValue(payload.approvalComment || "");

      return responseJSON({ status: "success", message: "บันทึกผลการพิจารณาเรียบร้อยแล้ว" });
    }

    // 3. บันทึกรายงานผลหลังกลับ (Submit Post-Travel Report)
    if (action === "submitReport") {
      const repSheet = ss.getSheetByName(SHEET_REPORTS);
      const now = new Date();
      const yearMonth = Utilities.formatDate(now, "Asia/Bangkok", "yyyyMM");
      const lastRow = repSheet.getLastRow();
      const reportId = "RP-" + yearMonth + "-" + ("000" + lastRow).slice(-4);

      const reportRow = [
        reportId,
        payload.requestId || "",
        Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss"),
        payload.summaryOutcome || "",
        payload.benefits || "",
        payload.actualExpense || 0,
        payload.attachmentUrl || "",
        payload.reporterName || ""
      ];
      repSheet.appendRow(reportRow);

      // อัปเดตสถานะในตาราง Requests ให้เป็น "เสร็จสิ้นภารกิจ"
      const reqSheet = ss.getSheetByName(SHEET_REQUESTS);
      const reqData = reqSheet.getDataRange().getValues();
      for (let i = 1; i < reqData.length; i++) {
        if (reqData[i][0] === payload.requestId) {
          reqSheet.getRange(i + 1, 21).setValue("เสร็จสิ้นภารกิจ");
          break;
        }
      }

      return responseJSON({ status: "success", reportId: reportId, message: "บันทึกรายงานผลเรียบร้อยแล้ว" });
    }

    return responseJSON({ status: "error", message: "Action ไม่ถูกต้อง" });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

/**
 * ดึงข้อมูลในแผ่นงานออกมาเป็น Array of Objects
 */
function getSheetData(sheet) {
  if (!sheet) return [];
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  const results = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const item = {};
    for (let c = 0; c < headers.length; c++) {
      let val = row[c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
      }
      item[headers[c]] = val;
    }
    results.push(item);
  }
  return results;
}

/**
 * ส่งค่า JSON กลับ
 */
function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
