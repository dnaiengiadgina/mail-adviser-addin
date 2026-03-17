/*
 * メール送信前チェック - Smart Alerts イベントハンドラ
 * OnMessageSend イベントで起動し、宛先・件名・添付ファイルを検査する
 */

// Office.js の初期化 (event-based activation では Office.onReady は不要)
// ただし Office.actions.associate は必須
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

/**
 * メール送信時に呼び出されるハンドラ
 * @param {Office.AddinCommands.Event} event
 */
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // 並列で宛先・件名を取得
  Promise.all([
    getSubject(item),
    getRecipients(item.to),
    getRecipients(item.cc),
    getRecipients(item.bcc),
    getAttachments(item)
  ])
  .then(([subject, toList, ccList, bccList, attachments]) => {
    const warnings = [];

    // ① 件名チェック
    if (!subject || subject.trim() === '') {
      warnings.push('件名が入力されていません。');
    }

    // ② 宛先チェック
    const allRecipients = [...toList, ...ccList, ...bccList];
    if (allRecipients.length === 0) {
      warnings.push('宛先が設定されていません。');
    }

    // ③ 社外ドメインチェック
    const INTERNAL_DOMAIN = 'fujilogi.co.jp';
    const externalRecipients = allRecipients.filter(r =>
      r.emailAddress && !r.emailAddress.toLowerCase().endsWith('@' + INTERNAL_DOMAIN)
    );
    if (externalRecipients.length > 0) {
      const extEmails = externalRecipients.map(r => r.emailAddress).join(', ');
      warnings.push(`社外への送信が含まれています:\n${extEmails}`);
    }

    // ④ 添付ファイルキーワードチェック
    // 本文に「添付」「別紙」「ファイル」などのキーワードがあるのに添付がない場合
    checkBodyKeywords(item, attachments, warnings).then(() => {
      if (warnings.length > 0) {
        // softBlock: ユーザーに確認を促す（「送信する」「戻る」の選択肢を表示）
        const message = '以下の点をご確認ください:\n\n' +
          warnings.map((w, i) => `${i + 1}. ${w}`).join('\n\n');

        event.completed({
          allowEvent: false,
          cancelLabel: '戻って確認する',
          sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
          errorMessage: message
        });
      } else {
        // 問題なし: 送信を許可
        event.completed({ allowEvent: true });
      }
    });
  })
  .catch(err => {
    // エラー時は送信を許可（ブロックしない）
    console.error('MailAdviser error:', err);
    event.completed({ allowEvent: true });
  });
}

/** 件名を取得 */
function getSubject(item) {
  return new Promise(resolve => {
    item.subject.getAsync(r => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '');
    });
  });
}

/** 宛先リストを取得 */
function getRecipients(field) {
  return new Promise(resolve => {
    field.getAsync(r => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}

/** 添付ファイルリストを取得 */
function getAttachments(item) {
  return new Promise(resolve => {
    item.getAttachmentsAsync(r => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}

/** 本文に「添付」関連キーワードがあるのに添付なしの場合に警告 */
function checkBodyKeywords(item, attachments, warnings) {
  return new Promise(resolve => {
    if (attachments.length > 0) {
      // 添付あり: スキップ
      resolve();
      return;
    }
    item.body.getAsync(Office.CoercionType.Text, r => {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        const body = r.value || '';
        const keywords = ['添付', '別紙', 'ファイルを', 'ファイルを送', 'お送りします', 'attached', 'attachment'];
        const found = keywords.some(kw => body.includes(kw));
        if (found) {
          warnings.push('本文に添付ファイルを示す言葉がありますが、添付ファイルがありません。');
        }
      }
      resolve();
    });
  });
}
