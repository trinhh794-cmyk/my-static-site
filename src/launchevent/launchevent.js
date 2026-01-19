
/* CẢNH BÁO GỬI RA NGOÀI DOANH NGHIỆP (Smart Alerts - OnMessageSend)
 * - ES5-compatible (tránh optional chaining, spread, v.v.) để chạy tốt trên classic Outlook
 * - Quét recipients (To/Cc/Bcc); nếu có ngoài DN => cảnh báo
 */

// ===== (Tùy chọn) Nhiều domain nội bộ =====
// null => tự lấy domain từ email người gửi.
// Hoặc điền mảng: var INTERNAL_DOMAINS = ["contoso.com","subsidiary.contoso.com"];
var INTERNAL_DOMAINS = null;

function getInternalDomains() {
  if (Object.prototype.toString.call(INTERNAL_DOMAINS) === "[object Array]" && INTERNAL_DOMAINS.length) {
    var lower = [];
    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) lower.push(String(INTERNAL_DOMAINS[i]).toLowerCase());
    return lower;
  }
  var me = (Office.context && Office.context.mailbox && Office.context.mailbox.userProfile)
           ? Office.context.mailbox.userProfile.emailAddress : "";
  var parts = String(me).split("@");
  var d = parts.length > 1 ? parts[1] : "";
  return d ? [d.toLowerCase()] : [];
}

function isExternalEmail(addr) {
  if (!addr) return false;
  var doms = getInternalDomains();
  var parts = String(addr).split("@");
  var domain = parts.length > 1 ? String(parts[1]).toLowerCase() : "";
  if (!domain) return false;
  for (var i = 0; i < doms.length; i++) {
    if (domain === doms[i]) return false;
  }
  return true; // không trùng domain nội bộ => external
}

// Lấy recipients (compose mode phải dùng getAsync)  ── Microsoft docs
// https://learn.microsoft.com/.../get-set-or-add-recipients [4](https://stackoverflow.com/questions/76483367/microsoft-office-launchevent-onmessagesend-not-working-for-windows-outlook-add-i)
function getAllRecipientsAsync(cb) {
  var item = Office.context.mailbox.item;
  var result = { to: [], cc: [], bcc: [] };
  var pending = 3;

  function done() { pending--; if (pending === 0) cb(result); }

  item.to.getAsync(function(r){
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) result.to = r.value;
    done();
  });
  item.cc.getAsync(function(r){
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) result.cc = r.value;
    done();
  });
  item.bcc.getAsync(function(r){
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) result.bcc = r.value;
    done();
  });
}

function hasExternalRecipient(recips) {
  var arrays = [recips.to || [], recips.cc || [], recips.bcc || []];
  for (var a = 0; a < arrays.length; a++) {
    var arr = arrays[a];
    for (var i = 0; i < arr.length; i++) {
      var email = arr[i] && arr[i].emailAddress ? arr[i].emailAddress : "";
      if (isExternalEmail(email)) return true;
    }
  }
  return false;
}

function onMessageSendHandler(event) {
  // LẤY RECIPIENTS
  getAllRecipientsAsync(function(recips){
    var isExternal = hasExternalRecipient(recips);
    if (isExternal) {
      // CẢNH BÁO: gửi ra ngoài DN → hiện Smart Alerts + "Send anyway"
      event.completed({
        allowEvent: false,
        errorMessage: "Bạn sắp gửi email ra ngoài doanh nghiệp. Hãy kiểm tra nội dung và người nhận.",
        // Ép PromptUser để có nút "Send anyway" (Mailbox ≥ 1.14)
        // https://learn.microsoft.com/.../office.mailboxenums.sendmodeoverride [5](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/outlook/smart-alerts-onmessagesend-walkthrough.md)
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
      });
      return;
    }
    // Không external → cho gửi
    event.completed({ allowEvent: true });
  });
}

// BẮT BUỘC: map tên handler trong manifest ↔ hàm JS (đặc biệt classic Outlook)
// https://learn.microsoft.com/.../smart-alerts-onmessagesend-walkthrough [1](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/apis)
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
