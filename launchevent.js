/* global Office */
/**
 * 送信前 宛先チェック（OnMessageSend / Smart Alerts）
 *
 * 動作:
 *   1. 送信ボタン押下で To / Cc / Bcc を取得
 *   2. 宛先一覧をダイアログに表示し、利用者に確認を求める
 *   3. 社外アドレスが含まれる場合は警告を先頭に表示
 *   4. 利用者は「このまま送信」または「宛先を修正する」を選択
 */

// ===== 設定 =====
// 自社ドメイン（サブドメインも社内扱い）。空配列の場合は Outlook が判定した
// recipientType（ExternalUser）を用いて社外判定します。
const INTERNAL_DOMAINS = ["japan-systems.co.jp"];
// 一覧に表示する最大件数（超過分は「他 N 件」と表示）
const MAX_LIST = 30;

Office.onReady(() => {
  // イベントベース ランタイムでは特に初期化不要
});

function getRecipientsAsync(field, label) {
  return new Promise((resolve) => {
    if (!field) return resolve([]);
    field.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return resolve([]);
      resolve((result.value || []).map((r) => ({ ...r, field: label })));
    });
  });
}

function isExternal(recipient) {
  const addr = (recipient.emailAddress || "").toLowerCase();
  const domain = addr.split("@")[1] || "";
  if (INTERNAL_DOMAINS.length > 0) {
    const internal = INTERNAL_DOMAINS.map((d) => d.toLowerCase());
    return !internal.some((d) => domain === d || domain.endsWith("." + d));
  }
  return recipient.recipientType === Office.MailboxEnums.RecipientType.ExternalUser;
}

function formatRecipient(r) {
  const name = r.displayName && r.displayName !== r.emailAddress ? `${r.displayName} ` : "";
  return `${name}<${r.emailAddress}>`;
}

function buildMessages(recipients) {
  const external = recipients.filter(isExternal);
  const internal = recipients.filter((r) => !isExternal(r));

  const listLines = (list) => {
    const shown = list.slice(0, MAX_LIST).map((r) => `- [${r.field}] ${formatRecipient(r)}`);
    if (list.length > MAX_LIST) shown.push(`- …他 ${list.length - MAX_LIST} 件`);
    return shown;
  };

  const md = [];
  const plain = [];

  if (external.length > 0) {
    md.push(`**⚠ 社外のアドレスが ${external.length} 件含まれています。送信して問題ないか確認してください。**`, "");
    plain.push(`【警告】社外のアドレスが ${external.length} 件含まれています。送信して問題ないか確認してください。`, "");
    md.push("**社外宛先**", ...listLines(external), "");
    plain.push("■社外宛先", ...listLines(external).map((l) => l.replace(/^- /, "・")), "");
  } else {
    md.push(`宛先を確認してください（合計 ${recipients.length} 件、社外なし）。`, "");
    plain.push(`宛先を確認してください（合計 ${recipients.length} 件、社外なし）。`, "");
  }

  if (internal.length > 0) {
    md.push("**社内宛先**", ...listLines(internal));
    plain.push("■社内宛先", ...listLines(internal).map((l) => l.replace(/^- /, "・")));
  }

  return {
    hasExternal: external.length > 0,
    markdown: md.join("\n"),
    plain: plain.join("\n"),
  };
}

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const [to, cc, bcc] = await Promise.all([
      getRecipientsAsync(item.to, "To"),
      getRecipientsAsync(item.cc, "Cc"),
      getRecipientsAsync(item.bcc, "Bcc"),
    ]);
    const recipients = [...to, ...cc, ...bcc];

    if (recipients.length === 0) {
      // 宛先なしは Outlook 側で弾かれるため通過
      event.completed({ allowEvent: true });
      return;
    }

    const msg = buildMessages(recipients);

    // allowEvent:false → 確認ダイアログを表示。SendMode=PromptUser のため
    // 利用者は「このまま送信」を選べる。cancelLabel は修正側ボタンの文言。
    event.completed({
      allowEvent: false,
      errorMessage: msg.plain,            // Markdown 非対応クライアント向け
      errorMessageMarkdown: msg.markdown, // Mailbox 1.14 以上
      cancelLabel: "宛先を修正する",
    });
  } catch (e) {
    // 予期せぬエラー時は送信を止めない（運用方針に応じて allowEvent:false に変更可）
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
