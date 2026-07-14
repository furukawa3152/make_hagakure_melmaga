const SHEET_NAMES = {
    main: "main",
    testAddress: "test_address",
    clientAddress: "client_address",
    unsubscribeList: "unsubscribe_list",
    sendLog: "send_log",
  };
  
  const MAIL_CONFIG = {
    subject: "【HAGAKURE PROGRAMMING塾】活動レポート",
    fallbackBody: "HTMLメールをご確認ください",
  };
  
  const UNSUBSCRIBE_SUBJECT = "配信停止希望";
  const UNSUBSCRIBE_BODY = "配信停止希望(このまま送信してください。)";
  const CLIENT_SEND_BATCH_SIZE = 100;
  const CLIENT_SEND_STATUS = {
    sent: "sent",
    unsubscribe: "unsubscribe",
    error: "error",
  };
  const NEWSLETTER_TRIGGER_CONFIG = {
    timezone: "Asia/Tokyo",
    startHour: 10,
    endHour: 12,
    intervalMinutes: 5,
    dailyHandler: "startNewsletterSendWindow",
    continuationHandler: "executeNewsletterSendByTrigger",
  };
  
  const ENV_KEYS = {
    newsletterFromEmail: "NEWSLETTER_FROM_EMAIL",
    newsletterFromName: "NEWSLETTER_FROM_NAME",
    unsubscribeEmail: "NEWSLETTER_UNSUBSCRIBE_EMAIL",
    unsubscribeTargetAlias: "NEWSLETTER_UNSUBSCRIBE_TARGET_ALIAS",
  };
  
  function getSheetNames_() {
    if (typeof SHEET_NAMES !== "undefined" && SHEET_NAMES) {
      return SHEET_NAMES;
    }
  
    return {
      main: "main",
      testAddress: "test_address",
      clientAddress: "client_address",
      unsubscribeList: "unsubscribe_list",
      sendLog: "send_log",
    };
  }
  
  /**
   * 必要な4シートを作成し、ヘッダーを初期化する。
   */
  function setupNewsletterSheets() {
    const sheetNames = getSheetNames_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetDefinitions = [
      { name: sheetNames.main, headers: ["mail_body", "test_sended", "client_sended"] },
      { name: sheetNames.testAddress, headers: ["address"] },
      {
        name: sheetNames.clientAddress,
        headers: ["address", "unsubscribe", "send_status", "sent_at", "error", "retry_count", "send_campaign"],
      },
      { name: sheetNames.unsubscribeList, headers: ["ymd", "address", "value"] },
      { name: sheetNames.sendLog, headers: ["ymd", "address", "status", "message", "executed_at"] },
    ];
  
    sheetDefinitions.forEach((definition) => {
      const sheet = ss.getSheetByName(definition.name) || ss.insertSheet(definition.name);
      ensureSheetHeaders_(sheet, definition.headers);
    });
  }
  
  /**
   * メルマガ送信のメイン実行関数。
   * 1回目: test_address に送信し、main.test_sended を更新。
   * 2回目: client_address に送信し、main.client_sended を更新。
   */
  function executeNewsletterSend() {
    const sheetNames = getSheetNames_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getRequiredSheet_(ss, sheetNames.main);
    const mainHeaders = getHeaderMap_(mainSheet);
    const mainRow = getLatestDataRow_(mainSheet, mainHeaders.mail_body);
  
    const mailBody = String(mainSheet.getRange(mainRow, mainHeaders.mail_body).getValue() || "").trim();
    if (!mailBody) {
      throw new Error("main.mail_body が空です。本文を入力してから実行してください。");
    }
  
    const testSendedValue = mainSheet.getRange(mainRow, mainHeaders.test_sended).getValue();
    const clientSendedValue = mainSheet.getRange(mainRow, mainHeaders.client_sended).getValue();
    if (clientSendedValue) {
      throw new Error("本送信は既に実行済みです。新しい行を作成して再実行してください。");
    }
  
    if (!testSendedValue) {
      const testAddresses = getAddressListFromSheet_(ss, sheetNames.testAddress);
      if (!testAddresses.length) {
        throw new Error("test_address.address に送信先がありません。");
      }
  
      sendBulkMail_(testAddresses, mailBody);
      mainSheet.getRange(mainRow, mainHeaders.test_sended).setValue(formatYmdJst_(new Date()));
      return {
        phase: "test",
        sentCount: testAddresses.length,
        remainingCount: 0,
        errorCount: 0,
      };
    }
  
    const result = sendToClients_(ss, mailBody, `main_row_${mainRow}`);
    Logger.log(
      `client送信: sent=${result.sentCount}, unsubscribed=${result.unsubscribedCount}, errors=${result.errorCount}, remaining=${result.remainingCount}`,
    );
  
    if (result.remainingCount === 0 && result.errorCount === 0) {
      mainSheet.getRange(mainRow, mainHeaders.client_sended).setValue(formatYmdJst_(new Date()));
    } else if (result.remainingCount === 0 && result.errorCount > 0) {
      throw new Error("本送信にエラーがあります。client_address.error を確認し、必要に応じて retryErroredClientSends() を実行してください。");
    }
  
    return result;
  }
  
  /**
   * 時間主導型トリガーから呼び出すバッチ送信用関数。
   * 本送信完了またはエラー発生時は、同じトリガーを停止する。
   */
  function startNewsletterSendWindow() {
    executeNewsletterSendByTrigger();
  }
  
  function executeNewsletterSendByTrigger() {
    if (!isNewsletterSendWindow_()) {
      return;
    }
  
    try {
      const result = executeNewsletterSend();
      if (isLatestNewsletterClientSendCompleted_()) {
        deleteNewsletterBatchTriggers_();
      } else if (result && result.remainingCount > 0) {
        installNewsletterContinuationTrigger_();
      }
    } catch (error) {
      deleteNewsletterBatchTriggers_();
      throw error;
    }
  }
  
  function installNewsletterBatchTrigger() {
    deleteNewsletterBatchTriggers_();
    ScriptApp.newTrigger(NEWSLETTER_TRIGGER_CONFIG.dailyHandler)
      .timeBased()
      .atHour(NEWSLETTER_TRIGGER_CONFIG.startHour)
      .everyDays(1)
      .inTimezone(NEWSLETTER_TRIGGER_CONFIG.timezone)
      .create();
  }
  
  function stopNewsletterBatchTrigger() {
    deleteNewsletterBatchTriggers_();
  }
  
  function installNewsletterContinuationTrigger_() {
    deleteNewsletterContinuationTriggers_();
  
    if (!canScheduleNextNewsletterContinuation_()) {
      return;
    }
  
    ScriptApp.newTrigger(NEWSLETTER_TRIGGER_CONFIG.continuationHandler)
      .timeBased()
      .after(NEWSLETTER_TRIGGER_CONFIG.intervalMinutes * 60 * 1000)
      .create();
  }
  
  function isNewsletterSendWindow_() {
    const hour = Number(Utilities.formatDate(new Date(), NEWSLETTER_TRIGGER_CONFIG.timezone, "H"));
    return hour >= NEWSLETTER_TRIGGER_CONFIG.startHour && hour < NEWSLETTER_TRIGGER_CONFIG.endHour;
  }
  
  function canScheduleNextNewsletterContinuation_() {
    const nextRun = new Date(Date.now() + NEWSLETTER_TRIGGER_CONFIG.intervalMinutes * 60 * 1000);
    const nextHour = Number(Utilities.formatDate(nextRun, NEWSLETTER_TRIGGER_CONFIG.timezone, "H"));
    return nextHour >= NEWSLETTER_TRIGGER_CONFIG.startHour && nextHour < NEWSLETTER_TRIGGER_CONFIG.endHour;
  }
  
  function retryErroredClientSends() {
    const sheetNames = getSheetNames_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getRequiredSheet_(ss, sheetNames.main);
    const mainHeaders = getHeaderMap_(mainSheet);
    const mainRow = getLatestDataRow_(mainSheet, mainHeaders.mail_body);
    const campaignId = `main_row_${mainRow}`;
    const clientSheet = getRequiredSheet_(ss, sheetNames.clientAddress);
    ensureSheetHeaders_(clientSheet, [
      "address",
      "unsubscribe",
      "send_status",
      "sent_at",
      "error",
      "retry_count",
      "send_campaign",
    ]);
    const clientHeaders = getHeaderMap_(clientSheet);
    const lastRow = clientSheet.getLastRow();
    let resetCount = 0;
  
    if (lastRow < 2) {
      return;
    }
  
    const rows = clientSheet.getRange(2, 1, lastRow - 1, clientSheet.getLastColumn()).getValues();
    rows.forEach((row, index) => {
      const rowNumber = index + 2;
      const rowCampaignId = String(row[clientHeaders.send_campaign - 1] || "").trim();
      const status = String(row[clientHeaders.send_status - 1] || "").trim().toLowerCase();
      if (rowCampaignId === campaignId && status === CLIENT_SEND_STATUS.error) {
        clientSheet.getRange(rowNumber, clientHeaders.send_status).setValue("");
        clientSheet.getRange(rowNumber, clientHeaders.error).setValue("");
        resetCount += 1;
      }
    });
  
    SpreadsheetApp.getActiveSpreadsheet().toast(`再送対象に戻しました: ${resetCount}件`, "Newsletter", 5);
  }
  
  function isLatestNewsletterClientSendCompleted_() {
    const sheetNames = getSheetNames_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getRequiredSheet_(ss, sheetNames.main);
    const mainHeaders = getHeaderMap_(mainSheet);
    const mainRow = getLatestDataRow_(mainSheet, mainHeaders.mail_body);
    return !!mainSheet.getRange(mainRow, mainHeaders.client_sended).getValue();
  }
  
  function deleteNewsletterBatchTriggers_() {
    ScriptApp.getProjectTriggers().forEach((trigger) => {
      if (
        trigger.getHandlerFunction() === NEWSLETTER_TRIGGER_CONFIG.dailyHandler ||
        trigger.getHandlerFunction() === NEWSLETTER_TRIGGER_CONFIG.continuationHandler
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }
  
  function deleteNewsletterContinuationTriggers_() {
    ScriptApp.getProjectTriggers().forEach((trigger) => {
      if (trigger.getHandlerFunction() === NEWSLETTER_TRIGGER_CONFIG.continuationHandler) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }
  
  function sendToClients_(ss, mailBody, campaignId) {
    const sheetNames = getSheetNames_();
    const clientSheet = getRequiredSheet_(ss, sheetNames.clientAddress);
    ensureSheetHeaders_(clientSheet, [
      "address",
      "unsubscribe",
      "send_status",
      "sent_at",
      "error",
      "retry_count",
      "send_campaign",
    ]);
    const clientHeaders = getHeaderMap_(clientSheet);
    const unsubscribeSet = getUnsubscribeAddressSet_(ss);
    const lastRow = clientSheet.getLastRow();
  
    if (lastRow < 2) {
      throw new Error("client_address.address に送信先がありません。");
    }
  
    const numRows = lastRow - 1;
    const rows = clientSheet.getRange(2, 1, numRows, clientSheet.getLastColumn()).getValues();
    const htmlBody = buildMailHtml_(mailBody);
    const mailOptions = getMailOptions_();
    const mailConfig = getNewsletterMailConfig_();
    const logRows = [];
    let unsubscribedCount = 0;
    let sentCount = 0;
    let errorCount = 0;
    let remainingCount = 0;
  
    rows.forEach((row, index) => {
      const rowNumber = index + 2;
      const address = normalizeAddress_(row[clientHeaders.address - 1]);
      const rowCampaignId = String(row[clientHeaders.send_campaign - 1] || "").trim();
      const currentStatus =
        rowCampaignId === campaignId ? String(row[clientHeaders.send_status - 1] || "").trim().toLowerCase() : "";
  
      if (!address || isFinishedClientStatus_(currentStatus)) {
        return;
      }
  
      if (unsubscribeSet.has(address)) {
        markClientRow_(clientSheet, clientHeaders, rowNumber, {
          unsubscribe: CLIENT_SEND_STATUS.unsubscribe,
          sendStatus: CLIENT_SEND_STATUS.unsubscribe,
          sentAt: "",
          error: "",
          retryCount: row[clientHeaders.retry_count - 1] || "",
          sendCampaign: campaignId,
        });
        logRows.push(createSendLogRow_(address, CLIENT_SEND_STATUS.unsubscribe, "unsubscribe_list に存在するため送信除外"));
        unsubscribedCount += 1;
        return;
      }
  
      if (sentCount >= CLIENT_SEND_BATCH_SIZE) {
        return;
      }
  
      const retryCount = rowCampaignId === campaignId ? Number(row[clientHeaders.retry_count - 1] || 0) : 0;
      try {
        GmailApp.sendEmail(address, mailConfig.subject, mailConfig.fallbackBody, mailOptions(htmlBody));
        markClientRow_(clientSheet, clientHeaders, rowNumber, {
          unsubscribe: "",
          sendStatus: CLIENT_SEND_STATUS.sent,
          sentAt: formatDateTimeJst_(new Date()),
          error: "",
          retryCount: retryCount,
          sendCampaign: campaignId,
        });
        logRows.push(createSendLogRow_(address, CLIENT_SEND_STATUS.sent, "送信成功"));
        sentCount += 1;
      } catch (error) {
        const message = error && error.message ? error.message : String(error);
        markClientRow_(clientSheet, clientHeaders, rowNumber, {
          unsubscribe: "",
          sendStatus: CLIENT_SEND_STATUS.error,
          sentAt: "",
          error: message,
          retryCount: retryCount + 1,
          sendCampaign: campaignId,
        });
        logRows.push(createSendLogRow_(address, CLIENT_SEND_STATUS.error, message));
        errorCount += 1;
      }
    });
  
    appendSendLogs_(ss, logRows);
    const summary = getClientSendSummary_(clientSheet, clientHeaders, campaignId);
    remainingCount += summary.pendingCount;
    errorCount = summary.errorCount;
  
    if (sentCount === 0 && unsubscribedCount === 0 && remainingCount === 0 && errorCount === 0) {
      throw new Error("client_address.address に有効な送信先がありません。");
    }
  
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `本送信バッチ完了: 送信${sentCount}件 / 除外${unsubscribedCount}件 / 残り${remainingCount}件 / エラー${errorCount}件`,
      "Newsletter",
      8,
    );
  
    return {
      sentCount: sentCount,
      unsubscribedCount: unsubscribedCount,
      errorCount: errorCount,
      remainingCount: remainingCount,
    };
  }
  
  function isFinishedClientStatus_(status) {
    return [
      CLIENT_SEND_STATUS.sent,
      CLIENT_SEND_STATUS.unsubscribe,
      CLIENT_SEND_STATUS.error,
    ].includes(status);
  }
  
  function markClientRow_(sheet, headers, rowNumber, values) {
    sheet.getRange(rowNumber, headers.unsubscribe).setValue(values.unsubscribe);
    sheet.getRange(rowNumber, headers.send_status).setValue(values.sendStatus);
    sheet.getRange(rowNumber, headers.sent_at).setValue(values.sentAt);
    sheet.getRange(rowNumber, headers.error).setValue(values.error);
    sheet.getRange(rowNumber, headers.retry_count).setValue(values.retryCount);
    sheet.getRange(rowNumber, headers.send_campaign).setValue(values.sendCampaign);
  }
  
  function getClientSendSummary_(sheet, headers, campaignId) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        pendingCount: 0,
        errorCount: 0,
      };
    }
  
    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    return rows.reduce(
      (summary, row) => {
        const address = normalizeAddress_(row[headers.address - 1]);
        if (!address) {
          return summary;
        }
  
        const rowCampaignId = String(row[headers.send_campaign - 1] || "").trim();
        const status =
          rowCampaignId === campaignId ? String(row[headers.send_status - 1] || "").trim().toLowerCase() : "";
        if (status === CLIENT_SEND_STATUS.error) {
          summary.errorCount += 1;
        } else if (!isFinishedClientStatus_(status)) {
          summary.pendingCount += 1;
        }
  
        return summary;
      },
      {
        pendingCount: 0,
        errorCount: 0,
      },
    );
  }
  
  function createSendLogRow_(address, status, message) {
    const now = new Date();
    return [formatYmdJst_(now), address, status, message, formatDateTimeJst_(now)];
  }
  
  function appendSendLogs_(ss, logRows) {
    if (!logRows.length) {
      return;
    }
  
    const sheetNames = getSheetNames_();
    const logSheet = ss.getSheetByName(sheetNames.sendLog) || ss.insertSheet(sheetNames.sendLog);
    ensureSheetHeaders_(logSheet, ["ymd", "address", "status", "message", "executed_at"]);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }
  
  function sendBulkMail_(addresses, mailBody) {
    const htmlBody = buildMailHtml_(mailBody);
    const mailOptions = getMailOptions_();
    const mailConfig = getNewsletterMailConfig_();
    let sentCount = 0;
    addresses.forEach((address) => {
      GmailApp.sendEmail(address, mailConfig.subject, mailConfig.fallbackBody, mailOptions(htmlBody));
      sentCount += 1;
    });
    SpreadsheetApp.getActiveSpreadsheet().toast(`送信処理完了: ${sentCount}件`, "Newsletter", 5);
  }
  
  function getMailOptions_() {
    const mailConfig = getNewsletterMailConfig_();
  
    return function (htmlBody) {
      return {
        name: mailConfig.name,   // 表示名だけ使う
        htmlBody: htmlBody,
      };
    };
  }
  function buildMailHtml_(mailBody) {
    const unsubscribeLink = buildUnsubscribeLink_();
  
    return `
  <div style="background:#f6fbf7; padding:20px 0; font-family: Arial, 'Hiragino Sans', Meiryo, sans-serif;">
  
  <table align="center" width="600" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:10px; overflow:hidden; border:1px solid #e0efe5;">
  
    <!-- ヘッダー -->
    <tr>
      <td style="background:#e8f5e9; text-align:center; padding:25px 20px; border-bottom:3px solid #a5d6a7;">
        
        <img src="https://hagakurepgm.net/static/images/logo.bd5050763f1f.png" width="90" style="display:block; margin:0 auto 10px;">
  
        <h1 style="margin:0; font-size:22px; color:#2e7d32;">
          HAGAKURE PROGRAMMING塾 通信
        </h1>
  
        <p style="margin:8px 0 0; font-size:16px; font-weight:bold; color:#388e3c;">
          2026年5月号
        </p>
      </td>
    </tr>
  
    <!-- 挨拶（ござる） -->
    <tr>
      <td style="padding:25px 20px;">
  
        <table width="100%">
          <tr>
            <td style="vertical-align:top; padding-right:10px;">
  
              <p style="margin-top:0; font-size:17px; line-height:1.8; color:#2e7d32; font-weight:bold;">
                いつも見守っていただき、誠にありがとうでござる。<br>
                このメールでは、我らの活動や最近の話題についてお届けするでござる。
              </p>
  
              <p style="font-size:15px; line-height:1.8; color:#333;">
                最近 Hagakure Programming塾 の Slack は開いたでござるか？<br>
                多くの話題で盛り上がっているので、久しぶりの者は是非 Slack を開いてみて欲しいでござる。
              </p>
  
              <table cellpadding="0" cellspacing="0" style="margin:12px 0;">
                <tr>
                  <td style="background:#81c784; border-radius:6px;">
                    <a href="https://app.slack.com/client/T02T5N5LD61/C02T1VAQRU6"
                       style="display:inline-block; padding:10px 20px; color:#ffffff; text-decoration:none; font-weight:bold; font-size:14px;">
                      Slackを開くでござる →
                    </a>
                  </td>
                </tr>
              </table>
  
              <p>
                どうぞ最後までご覧くだされ！
              </p>
  
            </td>
  
            <td width="120" style="text-align:center;">
              <img src="https://hagakurepgm.net/static/images/voice/voice-samurai.73a6e84d19cc.png" width="100">
            </td>
          </tr>
        </table>
  
        <!-- 本文 -->
        <div style="margin-top:20px; padding:16px; background:#f1f8f4; border-left:5px solid #81c784; border-radius:6px;">
          ${mailBody}
        </div>
  
      </td>
    </tr>
  
    <!-- リンク集 -->
    <tr>
      <td style="padding:24px 25px; background:#f6fbf7; border-top:1px solid #e0efe5; text-align:center;">
  
        <h2 style="margin:0 0 16px; font-size:16px; color:#2e7d32;">
          リンク集
        </h2>
  
        <table align="center" cellpadding="0" cellspacing="0" width="100%" style="max-width:320px;">
          <tr>
            <td style="padding:6px 0;">
              <table align="center" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td style="background:#81c784; border-radius:6px; text-align:center;">
                    <a href="https://www.facebook.com/profile.php?id=100085682160764&amp;locale=ja_JP"
                       style="display:block; padding:12px 20px; color:#ffffff; text-decoration:none; font-weight:bold; font-size:14px;">
                      公式 Facebook
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:6px 0;">
              <table align="center" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td style="background:#81c784; border-radius:6px; text-align:center;">
                    <a href="https://x.com/hagakurepgm"
                       style="display:block; padding:12px 20px; color:#ffffff; text-decoration:none; font-weight:bold; font-size:14px;">
                      公式 X
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:6px 0;">
              <table align="center" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td style="background:#81c784; border-radius:6px; text-align:center;">
                    <a href="https://hagakurepgm.net/"
                       style="display:block; padding:12px 20px; color:#ffffff; text-decoration:none; font-weight:bold; font-size:14px;">
                      Webサイト
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:6px 0;">
              <table align="center" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td style="background:#81c784; border-radius:6px; text-align:center;">
                    <a href="https://hagakurepgm.net/blog/"
                       style="display:block; padding:12px 20px; color:#ffffff; text-decoration:none; font-weight:bold; font-size:14px;">
                      ブログ Scroll
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
  
      </td>
    </tr>
  
    <!-- フッター -->
    <tr>
      <td style="padding:15px; text-align:center; font-size:12px; color:#777; border-top:1px solid #e0efe5;">
  
        <p style="margin:0; font-weight:bold; color:#2e7d32;">
          HAGAKURE PROGRAMMING塾
        </p>
  
        <p style="margin:10px 0 0; font-size:11px; color:#888; line-height:1.6;">
          このメルマガの、参加者以外への転送はお控えくださいでござる。
        </p>
  
        <p style="margin:8px 0;">
          <a href="${unsubscribeLink}" style="color:#66bb6a; text-decoration:none;">
            配信停止はこちらでござる
          </a>
        </p>
  
      </td>
    </tr>
  
  </table>
  </div>
  
  
  
    `;
  }
  
  function buildUnsubscribeLink_() {
    const subject = encodeURIComponent(UNSUBSCRIBE_SUBJECT);
    const body = encodeURIComponent(UNSUBSCRIBE_BODY);
    const unsubscribeEmail = getRequiredScriptProperty_(ENV_KEYS.unsubscribeEmail);
    return `mailto:${unsubscribeEmail}?subject=${subject}&body=${body}`;
  }
  
  function getNewsletterMailConfig_() {
    return {
      from: getRequiredScriptProperty_(ENV_KEYS.newsletterFromEmail),
      name: getRequiredScriptProperty_(ENV_KEYS.newsletterFromName),
      subject: MAIL_CONFIG.subject,
      fallbackBody: MAIL_CONFIG.fallbackBody,
    };
  }
  
  function getUnsubscribeTargetAlias_() {
    return getRequiredScriptProperty_(ENV_KEYS.unsubscribeTargetAlias);
  }
  
  function getRequiredScriptProperty_(key) {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    if (!value) {
      throw new Error(`スクリプトプロパティ '${key}' が未設定です。`);
    }
    return value;
  }
  
  function getUnsubscribeAddressSet_(ss) {
    const sheetNames = getSheetNames_();
    const sheet = getRequiredSheet_(ss, sheetNames.unsubscribeList);
    const headers = getHeaderMap_(sheet);
    const lastRow = sheet.getLastRow();
    const set = new Set();
  
    if (lastRow < 2) {
      return set;
    }
  
    const values = sheet.getRange(2, headers.address, lastRow - 1, 1).getValues();
    values.forEach((row) => {
      const address = normalizeAddress_(row[0]);
      if (address) {
        set.add(address);
      }
    });
  
    return set;
  }
  
  function getAddressListFromSheet_(ss, sheetName) {
    const sheet = getRequiredSheet_(ss, sheetName);
    const headers = getHeaderMap_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
  
    const values = sheet.getRange(2, headers.address, lastRow - 1, 1).getValues();
    return values
      .map((row) => normalizeAddress_(row[0]))
      .filter((value) => !!value);
  }
  
  function normalizeAddress_(value) {
    return String(value || "").trim().toLowerCase();
  }
  
  function getLatestDataRow_(sheet, requiredColumn) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error("main シートにデータ行がありません。");
    }
  
    const values = sheet.getRange(2, requiredColumn, lastRow - 1, 1).getValues();
    for (let i = values.length - 1; i >= 0; i -= 1) {
      if (String(values[i][0] || "").trim()) {
        return i + 2;
      }
    }
  
    throw new Error("main.mail_body にデータがありません。");
  }
  
  function getRequiredSheet_(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません。`);
    }
    return sheet;
  }
  
  function getHeaderMap_(sheet) {
    const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {};
    headerValues.forEach((value, index) => {
      const key = String(value || "").trim();
      if (key) {
        headerMap[key] = index + 1;
      }
    });
  
    validateHeader_(sheet.getName(), headerMap);
    return headerMap;
  }
  
  function ensureSheetHeaders_(sheet, requiredHeaders) {
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const headers = headerValues.map((value) => String(value || "").trim());
    const hasHeader = headers.some((value) => !!value);
    let nextColumn = hasHeader ? lastColumn + 1 : 1;
  
    requiredHeaders.forEach((header) => {
      if (!headers.includes(header)) {
        sheet.getRange(1, nextColumn).setValue(header);
        headers.push(header);
        nextColumn += 1;
      }
    });
  }
  
  function validateHeader_(sheetName, headerMap) {
    const sheetNames = getSheetNames_();
    const requiredBySheet = {
      [sheetNames.main]: ["mail_body", "test_sended", "client_sended"],
      [sheetNames.testAddress]: ["address"],
      [sheetNames.clientAddress]: [
        "address",
        "unsubscribe",
        "send_status",
        "sent_at",
        "error",
        "retry_count",
        "send_campaign",
      ],
      [sheetNames.unsubscribeList]: ["ymd", "address", "value"],
      [sheetNames.sendLog]: ["ymd", "address", "status", "message", "executed_at"],
    };
  
    const requiredHeaders = requiredBySheet[sheetName] || [];
    const missing = requiredHeaders.filter((name) => !headerMap[name]);
    if (missing.length) {
      throw new Error(`${sheetName} シートに必須カラムがありません: ${missing.join(", ")}`);
    }
  }
  
  function formatYmdJst_(date) {
    return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd");
  }
  
  function formatDateTimeJst_(date) {
    return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
  }
  
  
  
  function collectUnsubscribeEmails() {
    const LABEL_NAME = "processed_unsubscribe";
  
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("unsubscribe_list");
  
    // ラベル取得 or 作成
    let label = GmailApp.getUserLabelByName(LABEL_NAME);
    if (!label) {
      label = GmailApp.createLabel(LABEL_NAME);
    }
  
    // ★ 自分宛の未処理メールを取得（エイリアス不要）
    const query = `is:unread -label:${LABEL_NAME} newer_than:7d`;
    const threads = GmailApp.search(query);
  
    const lastRow = sheet.getLastRow();
    let existingEmails = [];
  
    if (lastRow >= 2) {
      existingEmails = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    }
  
    const newData = [];
  
    threads.forEach((thread) => {
      const messages = thread.getMessages();
  
      messages.forEach((message) => {
        if (!message.isUnread()) return;
  
        const subject = message.getSubject() || "";
        const body = message.getPlainBody() || "";
  
        // ★ 配信停止判定（強化）
        const isUnsubscribe =
          subject.match(/配信停止|unsubscribe/i) ||
          body.match(/配信停止|unsubscribe/i);
  
        if (!isUnsubscribe) return;
  
        const from = message.getFrom();
        const date = message.getDate();
  
        // ★ メールアドレス抽出強化
        const emailMatch = from.match(/[\w\.-]+@[\w\.-]+\.\w+/);
        const email = emailMatch ? emailMatch[0].toLowerCase() : from.toLowerCase();
  
        // 重複チェック
        if (!existingEmails.includes(email)) {
          newData.push([date, email, subject]);
          existingEmails.push(email);
        }
  
        // 既読化
        message.markRead();
      });
  
      // ラベル付与（再処理防止）
      thread.addLabel(label);
    });
  
    // 書き込み
    if (newData.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newData.length, 3).setValues(newData);
    }
  
    Logger.log(`配信停止取得: ${newData.length}件`);
  }