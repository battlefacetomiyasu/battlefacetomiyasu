# バトルフェイス富安 ライブ予約フォーム 運用マニュアル

カスタムHTMLフォームからGAS（Google Apps Script）へ送信し、Googleスプレッドシートへの記録と予約完了メールの自動送信を行う仕組み。Googleフォーム本体は不要。

## 構成

- `forms/live-YYYYMMDD.html`：各ライブの予約ページ
- `forms/thanks.html`：予約完了ページ
- `assets/css/style.css`：共通スタイル
- Googleスプレッドシート ＋ GAS Webアプリ：記録・完了メール送信（URLは共通で使い回し）

---

## 新しいライブの予約ページを追加する手順

### 1. フォームHTMLを複製する
既存のライブHTML（例：`forms/live-20261213.html`）をコピーして、新しい日付などでファイルを作る。
- 例：`forms/live-20270115.html`

### 2. 画面に表示される公演情報を書き換える
- `<title>`
- `<h1>`（イベント名）
- `<dl class="event-meta">` 内の公演日・開催場所・OPEN/START・出演・チケット情報

### 3. 送信用の `hidden` 項目を書き換える
`<form>` 直下の `input[type="hidden"]` を新しいライブのものに変更する。
**※ `sheetName` にスプレッドシートで使えない記号（`\` `/` `?` `*` `[` `]` `:`）は入れないこと。**

```html
<input type="hidden" name="sheetName" value="2027-01-15_会場名">
<input type="hidden" name="eventTitle" value="イベント名">
<input type="hidden" name="eventDate" value="2027.01.15(金)">
<input type="hidden" name="eventVenue" value="会場名">
<input type="hidden" name="eventTime" value="OPEN 18:00 / START 18:30">
```

### 4. イベント一覧（index.html)を編集する
- 新しいイベントの追加・既存ライブの編集（受付の停止）を行う。

# ライブ予約 GAS（Google Apps Script）仕様・運用メモ

カスタムHTMLフォームからの送信を受け取り、Googleスプレッドシートへの追記（シート自動作成・ヘッダー保護）と、予約者宛て完了メール／主催者宛て通知メールの送信を行うスクリプト。

## 役割と処理の流れ

1. `LockService` で同時送信を排他制御（最大10秒待機）し、記録の競合を防ぐ
2. HTMLフォームから送信された値（`name`, `email`, `ticketCount`, `sheetName`, `eventTitle`, `eventDate`, `eventVenue`, `eventTime`）を受け取る
3. `sheetName` のシートを探し、**なければ作成して1行目にヘッダーを追加**、**あれば末尾に1行追記**する
4. `MailApp.sendEmail()` で予約者へ予約完了メールを送信する
5. `ADMIN_EMAIL` が設定されていれば、主催者（自分）宛てにも新規予約通知を送る

---

## GASコード全文

```javascript
const ADMIN_EMAIL = "bftomiyasu@gmail.com";
const HEADERS = ["送信日時", "公演日", "公演名", "会場", "お名前", "メールアドレス", "予約枚数"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = parsePayload(e);

    const name = String(data.name || "").trim();
    const email = String(data.email || "").trim();
    const ticketCount = String(data.ticketCount || "").trim();

    const eventTitle = String(data.eventTitle || "HeartLand pre.HOMEGAME vol.20").trim();
    const eventDate = String(data.eventDate || "2026.12.13(日)").trim();
    const eventVenue = String(data.eventVenue || "愛知県 新栄ハートランド").trim();
    const eventTime = String(data.eventTime || "OPEN 17:30 / START 18:00").trim();

    const rawSheetName = String(data.sheetName || eventDate || "予約一覧").trim();
    const sheetName = sanitizeSheetName(rawSheetName);

    if (!name || !email || !ticketCount) {
      return jsonResponse({ ok: false, error: "必須項目が不足しています。" });
    }

    // 1. シート取得（なければ作成）、あればそのまま追記
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(HEADERS);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    sheet.appendRow([new Date(), eventDate, eventTitle, eventVenue, name, email, ticketCount]);

    // 2. 予約者へ完了メール送信
    const subject = `【予約完了】${eventTitle}（${eventDate}）`;
    const body = [
      `${name} 様`,
      ``,
      `ライブのご予約ありがとうございます。以下の内容で受け付けました。`,
      ``,
      `■ 公演名：${eventTitle}`,
      `■ 公演日：${eventDate}`,
      `■ 会場：${eventVenue}`,
      `■ 時間：${eventTime}`,
      `■ 予約枚数：${ticketCount}枚`,
      ``,
      `当日会場でお会いできるのを楽しみにしています！`,
      `※変更やキャンセルがある場合は、このメールまたは下記までご連絡ください。`,
      `連絡先：${ADMIN_EMAIL || "bftomiyasu@gmail.com"}`,
      ``,
      `バトルフェイス富安`,
    ].join("\n");

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: "バトルフェイス富安 ライブ予約",
      replyTo: ADMIN_EMAIL || undefined,
    });

    // 3. 主催者（自分）宛て通知（任意）
    if (ADMIN_EMAIL) {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: `【新規予約】${eventDate} ${eventTitle}（${name}さん / ${ticketCount}枚）`,
        body: [
          `新しい予約が入りました。`,
          ``,
          `記録シート：${sheetName}`,
          `公演日：${eventDate}`,
          `公演名：${eventTitle}`,
          `会場：${eventVenue}`,
          `時間：${eventTime}`,
          `お名前：${name}`,
          `メール：${email}`,
          `枚数：${ticketCount}枚`,
        ].join("\n"),
      });
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function sanitizeSheetName(name) {
  const cleaned = String(name)
    .replace(/[\\/?*[\]:]/g, "-")
    .trim();
  return (cleaned || "予約一覧").slice(0, 90);
}

function parsePayload(e) {
  if (e && e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  return {};
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
