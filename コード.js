/**
 * LIFF研究参加登録システム
 *
 * 必要なスクリプトプロパティ
 * --------------------------------
 * SPREADSHEET_ID
 * LINE_LOGIN_CHANNEL_ID
 * LIFF_ID
 * LINE_ADD_FRIEND_URL
 *
 * participantsシートの必須列
 * --------------------------------
 * participant_id
 * consent_status
 * participant_status
 * line_user_id
 * linked_at
 * language
 * registration_status
 * app_consent_status
 * app_consent_at
 * app_consent_version
 * email
 * note
 */

const CONFIG = {
  PARTICIPANTS_SHEET: 'participants',
  REGISTRATION_LOG_SHEET: 'registration_logs',
  LINE_VERIFY_URL: 'https://api.line.me/oauth2/v2.1/verify',
};

/**
 * フロントエンド(GitHub Pages)向けの設定情報を返すAPI
 *
 * GASのWebアプリはHtmlServiceで返すと内部で
 * iframeにラップされ、LIFFの初期化が完了しない事象が
 * あったため、画面はGitHub Pagesの静的ページに移し、
 * GASはJSON APIとしてのみ利用する。
 */
function doGet() {
  const properties =
    PropertiesService.getScriptProperties();

  const config = {
    liffId:
      properties.getProperty('LIFF_ID') || '',
    lineAddFriendUrl:
      properties.getProperty('LINE_ADD_FRIEND_URL') || '',
  };

  return ContentService
    .createTextOutput(JSON.stringify(config))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 参加者登録API
 *
 * GitHub Pages側のfetch()から呼び出される。
 * プリフライトを避けるため、リクエストの
 * Content-Typeはtext/plainで送られてくる想定。
 */
function doPost(e) {
  let result;

  try {
    const request =
      JSON.parse(e.postData.contents);

    result =
      registerParticipant(request);
  } catch (error) {
    result = {
      success: false,
      code: 'INVALID_REQUEST',
      message:
        'リクエストを処理できませんでした。',
    };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 参加者登録処理
 *
 * Index.htmlのgoogle.script.runから呼び出される
 *
 * @param {Object} request
 * @param {string} request.participantId
 * @param {string} request.idToken
 * @param {string} request.language
 * @param {boolean} request.consentAccepted
 * @param {string} request.consentVersion
 * @return {Object}
 */
function registerParticipant(request) {
  const lock =
    LockService.getScriptLock();

  let normalizedParticipantId = '';
  let verifiedLineUserId = '';

  try {
    lock.waitLock(10000);

    validateRequest_(request);

    normalizedParticipantId =
      normalizeParticipantId_(
        request.participantId
      );

    /**
     * LINE IDトークンをサーバー側で検証
     */
    const lineProfile =
      verifyLineIdToken_(
        request.idToken
      );

    verifiedLineUserId =
      String(lineProfile.sub || '').trim();

    const email =
      String(lineProfile.email || '').trim();

    if (!verifiedLineUserId) {
      throw new Error(
        'LINEユーザー情報を取得できませんでした。'
      );
    }

    const sheet =
      getParticipantsSheet_();

    const dataRange =
      sheet.getDataRange();

    const values =
      dataRange.getValues();

    if (values.length < 2) {
      throw new Error(
        '参加者台帳に登録データがありません。'
      );
    }

    const headers =
      values[0].map(function (header) {
        return String(header).trim();
      });

    const indexes =
      getColumnIndexes_(
        headers,
        [
          'participant_id',
          'consent_status',
          'participant_status',
          'line_user_id',
          'linked_at',
          'language',
          'registration_status',
          'app_consent_status',
          'app_consent_at',
          'app_consent_version',
          'email',
          'note',
        ]
      );

    let targetRowNumber = -1;
    let sameLineUserRowNumber = -1;

    /**
     * 参加IDとLINEユーザーIDの重複確認
     */
    for (
      let rowIndex = 1;
      rowIndex < values.length;
      rowIndex++
    ) {
      const row =
        values[rowIndex];

      const registeredParticipantId =
        normalizeParticipantId_(
          row[indexes.participant_id]
        );

      const registeredLineUserId =
        String(
          row[indexes.line_user_id] || ''
        ).trim();

      if (
        registeredParticipantId ===
        normalizedParticipantId
      ) {
        targetRowNumber =
          rowIndex + 1;
      }

      if (
        registeredLineUserId &&
        registeredLineUserId ===
        verifiedLineUserId
      ) {
        sameLineUserRowNumber =
          rowIndex + 1;
      }
    }

    /**
     * 参加IDが台帳に存在しない
     */
    if (targetRowNumber === -1) {
      logRegistration_({
        participantId:
          normalizedParticipantId,
        lineUserId:
          verifiedLineUserId,
        email,
        result:
          'ERROR',
        message:
          '参加IDが存在しない',
      });

      return {
        success: false,
        code: 'PARTICIPANT_NOT_FOUND',
        message:
          '参加IDを確認できませんでした。' +
          '書面に記載された番号を再度ご確認ください。',
      };
    }

    const targetRowValues =
      sheet
        .getRange(
          targetRowNumber,
          1,
          1,
          headers.length
        )
        .getValues()[0];

    const consentStatus =
      String(
        targetRowValues[
        indexes.consent_status
        ] || ''
      ).trim();

    const participantStatus =
      String(
        targetRowValues[
        indexes.participant_status
        ] || ''
      ).trim();

    const currentLineUserId =
      String(
        targetRowValues[
        indexes.line_user_id
        ] || ''
      ).trim();

    /**
     * 書面同意取得済みか確認
     */
    if (
      !isConsentConfirmed_(
        consentStatus
      )
    ) {
      logRegistration_({
        participantId:
          normalizedParticipantId,
        lineUserId:
          verifiedLineUserId,
        email,
        result:
          'ERROR',
        message:
          '書面同意未確認',
      });

      return {
        success: false,
        code: 'CONSENT_NOT_CONFIRMED',
        message:
          '書面での同意取得を確認できませんでした。' +
          '研究担当者へお問い合わせください。',
      };
    }

    /**
     * 参加停止、撤回、終了を確認
     */
    if (
      isInactiveParticipant_(
        participantStatus
      )
    ) {
      logRegistration_({
        participantId:
          normalizedParticipantId,
        lineUserId:
          verifiedLineUserId,
        email,
        result:
          'ERROR',
        message:
          '参加者ステータス無効',
      });

      return {
        success: false,
        code: 'PARTICIPANT_INACTIVE',
        message:
          '現在、この参加IDではLINE登録を行えません。' +
          '研究担当者へお問い合わせください。',
      };
    }

    /**
     * 参加IDが別のLINEアカウントに登録済み
     */
    if (
      currentLineUserId &&
      currentLineUserId !==
      verifiedLineUserId
    ) {
      logRegistration_({
        participantId:
          normalizedParticipantId,
        lineUserId:
          verifiedLineUserId,
        email,
        result:
          'ERROR',
        message:
          '参加IDが別LINEユーザーと登録済み',
      });

      return {
        success: false,
        code:
          'PARTICIPANT_ALREADY_LINKED',
        message:
          'この参加IDは、すでに別のLINEアカウントと登録されています。' +
          '研究担当者へお問い合わせください。',
      };
    }

    /**
     * 同じLINEアカウントが別参加IDに登録済み
     */
    if (
      sameLineUserRowNumber !== -1 &&
      sameLineUserRowNumber !==
      targetRowNumber
    ) {
      logRegistration_({
        participantId:
          normalizedParticipantId,
        lineUserId:
          verifiedLineUserId,
        email,
        result:
          'ERROR',
        message:
          'LINEユーザーが別参加IDと登録済み',
      });

      return {
        success: false,
        code:
          'LINE_USER_ALREADY_LINKED',
        message:
          'このLINEアカウントは、すでに別の参加IDと登録されています。' +
          '研究担当者へお問い合わせください。',
      };
    }

    const now =
      new Date();

    const language =
      sanitizeLanguage_(
        request.language
      );

    const consentVersion =
      String(
        request.consentVersion ||
        '1.0'
      ).trim();

    /**
     * 同じ参加ID・同じLINEユーザーの場合
     * 同意情報やメールアドレスだけ再更新
     */
    if (
      currentLineUserId ===
      verifiedLineUserId
    ) {
      updateParticipantRow_({
        sheet,
        rowNumber:
          targetRowNumber,
        indexes,
        lineUserId:
          verifiedLineUserId,
        linkedAt:
          targetRowValues[
          indexes.linked_at
          ] || now,
        language,
        registrationStatus:
          '登録完了',
        appConsentStatus:
          '同意済み',
        appConsentAt:
          now,
        appConsentVersion:
          consentVersion,
        email,
      });

      logRegistration_({
        participantId:
          normalizedParticipantId,
        lineUserId:
          verifiedLineUserId,
        email,
        result:
          'SUCCESS',
        message:
          '既登録ユーザー再確認',
      });

      return {
        success: true,
        alreadyRegistered: true,
        participantId:
          normalizedParticipantId,
        emailCollected:
          Boolean(email),
        message:
          'すでに登録が完了しています。',
      };
    }

    /**
     * 新規登録
     */
    updateParticipantRow_({
      sheet,
      rowNumber:
        targetRowNumber,
      indexes,
      lineUserId:
        verifiedLineUserId,
      linkedAt:
        now,
      language,
      registrationStatus:
        '登録完了',
      appConsentStatus:
        '同意済み',
      appConsentAt:
        now,
      appConsentVersion:
        consentVersion,
      email,
    });

    logRegistration_({
      participantId:
        normalizedParticipantId,
      lineUserId:
        verifiedLineUserId,
      email,
      result:
        'SUCCESS',
      message:
        'LIFF登録完了',
    });

    return {
      success: true,
      alreadyRegistered: false,
      participantId:
        normalizedParticipantId,
      emailCollected:
        Boolean(email),
      message:
        '研究参加登録が完了しました。',
    };
  } catch (error) {
    console.error(
      'registerParticipant error:',
      error
    );

    logRegistration_({
      participantId:
        normalizedParticipantId ||
        request?.participantId ||
        '',
      lineUserId:
        verifiedLineUserId,
      email:
        '',
      result:
        'SYSTEM_ERROR',
      message:
        error.message ||
        String(error),
    });

    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message:
        '登録処理中にエラーが発生しました。' +
        '時間をおいて再度お試しください。',
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      console.warn(
        'Lock release error:',
        error
      );
    }
  }
}

/**
 * participantsシートの対象行を更新
 *
 * @param {Object} params
 */
function updateParticipantRow_(params) {
  const {
    sheet,
    rowNumber,
    indexes,
    lineUserId,
    linkedAt,
    language,
    registrationStatus,
    appConsentStatus,
    appConsentAt,
    appConsentVersion,
    email,
  } = params;

  sheet
    .getRange(
      rowNumber,
      indexes.line_user_id + 1
    )
    .setValue(
      lineUserId
    );

  sheet
    .getRange(
      rowNumber,
      indexes.linked_at + 1
    )
    .setValue(
      linkedAt
    );

  sheet
    .getRange(
      rowNumber,
      indexes.language + 1
    )
    .setValue(
      language
    );

  sheet
    .getRange(
      rowNumber,
      indexes.registration_status + 1
    )
    .setValue(
      registrationStatus
    );

  sheet
    .getRange(
      rowNumber,
      indexes.app_consent_status + 1
    )
    .setValue(
      appConsentStatus
    );

  sheet
    .getRange(
      rowNumber,
      indexes.app_consent_at + 1
    )
    .setValue(
      appConsentAt
    );

  sheet
    .getRange(
      rowNumber,
      indexes.app_consent_version + 1
    )
    .setValue(
      appConsentVersion
    );

  /**
   * メールアドレスが取得できた場合のみ更新
   * 未承認・未取得の場合は既存値を消さない
   */
  if (email) {
    sheet
      .getRange(
        rowNumber,
        indexes.email + 1
      )
      .setValue(
        email
      );
  }
}

/**
 * LINE IDトークンを検証
 *
 * @param {string} idToken
 * @return {Object}
 */
function verifyLineIdToken_(idToken) {
  const properties =
    PropertiesService.getScriptProperties();

  const channelId =
    properties.getProperty(
      'LINE_LOGIN_CHANNEL_ID'
    );

  if (!channelId) {
    throw new Error(
      'LINE_LOGIN_CHANNEL_IDが設定されていません。'
    );
  }

  const response =
    UrlFetchApp.fetch(
      CONFIG.LINE_VERIFY_URL,
      {
        method: 'post',
        payload: {
          id_token:
            idToken,
          client_id:
            channelId,
        },
        muteHttpExceptions:
          true,
      }
    );

  const statusCode =
    response.getResponseCode();

  const responseText =
    response.getContentText();

  let responseData;

  try {
    responseData =
      JSON.parse(
        responseText
      );
  } catch (error) {
    console.error(
      'LINE verify response:',
      responseText
    );

    throw new Error(
      'LINE認証結果の解析に失敗しました。'
    );
  }

  if (statusCode !== 200) {
    console.error(
      'LINE verify error:',
      responseText
    );

    throw new Error(
      'LINE認証に失敗しました。' +
      'LINEアプリから登録画面を開き直してください。'
    );
  }

  /**
   * IDトークンの送信先確認
   */
  if (
    String(responseData.aud) !==
    String(channelId)
  ) {
    throw new Error(
      'LINE認証情報の送信先が一致しません。'
    );
  }

  /**
   * 有効期限確認
   */
  const currentUnixTime =
    Math.floor(
      Date.now() / 1000
    );

  if (
    responseData.exp &&
    Number(responseData.exp) <
    currentUnixTime
  ) {
    throw new Error(
      'LINE認証の有効期限が切れています。' +
      '画面を再読み込みしてください。'
    );
  }

  /**
   * ユーザーID確認
   */
  if (!responseData.sub) {
    throw new Error(
      'LINEユーザーIDを取得できませんでした。'
    );
  }

  return responseData;
}

/**
 * リクエスト内容を検証
 *
 * @param {Object} request
 */
function validateRequest_(request) {
  if (
    !request ||
    typeof request !== 'object'
  ) {
    throw new Error(
      '登録情報が送信されていません。'
    );
  }

  if (!request.participantId) {
    throw new Error(
      '参加IDを入力してください。'
    );
  }

  if (!request.idToken) {
    throw new Error(
      'LINE認証情報を取得できませんでした。'
    );
  }

  if (
    request.consentAccepted !== true
  ) {
    throw new Error(
      '情報の取得および使用目的への同意が確認できません。'
    );
  }

  const participantId =
    normalizeParticipantId_(
      request.participantId
    );

  if (
    !participantId ||
    participantId.length > 20
  ) {
    throw new Error(
      '参加IDの形式が正しくありません。'
    );
  }

  /**
   * EP0001形式を想定
   * 必要に応じて正規表現を変更
   */
  if (
    !/^EP\d{4,10}$/.test(
      participantId
    )
  ) {
    throw new Error(
      '参加IDは「EP0001」の形式で入力してください。'
    );
  }
}

/**
 * 参加IDを正規化
 *
 * 全角英数字を半角化
 * 大文字へ変換
 * 空白・ハイフンを削除
 *
 * @param {*} value
 * @return {string}
 */
function normalizeParticipantId_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(
      /[Ａ-Ｚａ-ｚ０-９]/g,
      function (character) {
        return String.fromCharCode(
          character.charCodeAt(0) -
          0xFEE0
        );
      }
    )
    .replace(
      /[\s　\-ー]/g,
      ''
    );
}

/**
 * 使用言語を制限
 *
 * @param {string} language
 * @return {string}
 */
function sanitizeLanguage_(language) {
  const allowedLanguages = [
    'ja',
    'en',
    'ms',
  ];

  if (
    allowedLanguages.includes(
      language
    )
  ) {
    return language;
  }

  return 'ja';
}

/**
 * 書面同意済み判定
 *
 * @param {*} value
 * @return {boolean}
 */
function isConsentConfirmed_(value) {
  const normalized =
    String(value || '')
      .trim()
      .toUpperCase();

  const acceptedValues = [
    '済',
    '同意済み',
    '同意済',
    'YES',
    'TRUE',
    '1',
  ];

  return acceptedValues.includes(
    normalized
  );
}

/**
 * 参加不可状態を判定
 *
 * @param {*} value
 * @return {boolean}
 */
function isInactiveParticipant_(value) {
  const normalized =
    String(value || '')
      .trim();

  const inactiveStatuses = [
    '撤回',
    '辞退',
    '終了',
    '停止',
    '除外',
  ];

  return inactiveStatuses.includes(
    normalized
  );
}

/**
 * participantsシートを取得
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getParticipantsSheet_() {
  const spreadsheet =
    getSpreadsheet_();

  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.PARTICIPANTS_SHEET
    );

  if (!sheet) {
    throw new Error(
      `${CONFIG.PARTICIPANTS_SHEET}シートがありません。`
    );
  }

  return sheet;
}

/**
 * スプレッドシート取得
 *
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet_() {
  const properties =
    PropertiesService.getScriptProperties();

  const spreadsheetId =
    properties.getProperty(
      'SPREADSHEET_ID'
    );

  if (!spreadsheetId) {
    throw new Error(
      'SPREADSHEET_IDが設定されていません。'
    );
  }

  return SpreadsheetApp.openById(
    spreadsheetId
  );
}

/**
 * 必須列のインデックスを取得
 *
 * @param {string[]} headers
 * @param {string[]} requiredColumns
 * @return {Object}
 */
function getColumnIndexes_(
  headers,
  requiredColumns
) {
  const indexes = {};

  requiredColumns.forEach(
    function (columnName) {
      const columnIndex =
        headers.indexOf(
          columnName
        );

      if (columnIndex === -1) {
        throw new Error(
          `participantsシートに「${columnName}」列がありません。`
        );
      }

      indexes[columnName] =
        columnIndex;
    }
  );

  return indexes;
}

/**
 * 登録処理ログを保存
 *
 * @param {Object} data
 */
function logRegistration_(data) {
  try {
    const spreadsheet =
      getSpreadsheet_();

    let sheet =
      spreadsheet.getSheetByName(
        CONFIG.REGISTRATION_LOG_SHEET
      );

    if (!sheet) {
      sheet =
        spreadsheet.insertSheet(
          CONFIG.REGISTRATION_LOG_SHEET
        );

      sheet.appendRow([
        'logged_at',
        'participant_id',
        'line_user_id',
        'email',
        'result',
        'message',
      ]);

      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.participantId || '',
      data.lineUserId || '',
      data.email || '',
      data.result || '',
      data.message || '',
    ]);
  } catch (error) {
    console.error(
      'logRegistration error:',
      error
    );
  }
}

/**
 * 初期シート作成用
 *
 * 初回のみGASエディタから手動実行
 * 既存participantsシートがある場合は上書きしない
 */
function setupSheets() {
  const spreadsheet =
    getSpreadsheet_();

  let participantsSheet =
    spreadsheet.getSheetByName(
      CONFIG.PARTICIPANTS_SHEET
    );

  if (!participantsSheet) {
    participantsSheet =
      spreadsheet.insertSheet(
        CONFIG.PARTICIPANTS_SHEET
      );

    participantsSheet.appendRow([
      'participant_id',
      'consent_status',
      'participant_status',
      'line_user_id',
      'linked_at',
      'language',
      'registration_status',
      'app_consent_status',
      'app_consent_at',
      'app_consent_version',
      'email',
      'note',
    ]);

    participantsSheet.setFrozenRows(1);
  }

  let logSheet =
    spreadsheet.getSheetByName(
      CONFIG.REGISTRATION_LOG_SHEET
    );

  if (!logSheet) {
    logSheet =
      spreadsheet.insertSheet(
        CONFIG.REGISTRATION_LOG_SHEET
      );

    logSheet.appendRow([
      'logged_at',
      'participant_id',
      'line_user_id',
      'email',
      'result',
      'message',
    ]);

    logSheet.setFrozenRows(1);
  }
}

/**
 * 設定値確認用
 *
 * GASエディタから手動実行するとログへ出力
 * IDトークンなどの秘密情報は出力しない
 */
function checkConfiguration() {
  const properties =
    PropertiesService.getScriptProperties();

  const configStatus = {
    SPREADSHEET_ID:
      Boolean(
        properties.getProperty(
          'SPREADSHEET_ID'
        )
      ),

    LINE_LOGIN_CHANNEL_ID:
      Boolean(
        properties.getProperty(
          'LINE_LOGIN_CHANNEL_ID'
        )
      ),

    LIFF_ID:
      Boolean(
        properties.getProperty(
          'LIFF_ID'
        )
      ),

    LINE_ADD_FRIEND_URL:
      Boolean(
        properties.getProperty(
          'LINE_ADD_FRIEND_URL'
        )
      ),
  };

  console.log(
    JSON.stringify(
      configStatus,
      null,
      2
    )
  );

  return configStatus;
}