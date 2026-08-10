/*
 * Calgary Counselling Centre — Recipient Check
 * Outlook Smart Alerts add-in (OnMessageSend).
 * Shows To / Cc / Bcc separately, flags external addresses, guarded so it
 * never hangs. Nothing is added to the message; recipients never see this.
 */

// Domains that count as "internal" (case-insensitive, includes sub-domains).
var INTERNAL_DOMAINS = ["calgarycounselling.com"];

// false = prompt ONLY when there is an external recipient.
var PROMPT_ON_MULTIPLE_RECIPIENTS = false;
var MULTIPLE_RECIPIENT_THRESHOLD = 2;

// Max addresses to list per field before "and N more".
var MAX_LISTED = 8;

// If reading recipients stalls this long (ms), allow the send.
var SAFETY_TIMEOUT_MS = 4000;

Office.onReady(function () {});

function domainMatches(domain, internal) {
  if (domain === internal) return true;
  var dotted = "." + internal;
  if (domain.length <= dotted.length) return false;
  return domain.lastIndexOf(dotted) === (domain.length - dotted.length);
}

function domainOf(address) {
  if (!address) return "";
  var at = address.lastIndexOf("@");
  if (at === -1) return "";
  return address.substring(at + 1).toLowerCase().replace(/^\s+|\s+$/g, "");
}

function isInternal(address) {
  var domain = domainOf(address);
  if (!domain) return false;
  for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
    if (domainMatches(domain, INTERNAL_DOMAINS[i].toLowerCase())) return true;
  }
  return false;
}

function externalIn(list) {
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var email = list[i] && list[i].emailAddress ? list[i].emailAddress : "";
    if (!isInternal(email)) out.push(email);
  }
  return out;
}

// Format one field ("To"/"Cc"/"Bcc"), tagging external addresses. null if empty.
function formatField(label, list) {
  if (!list || !list.length) return null;
  var parts = [];
  var n = list.length < MAX_LISTED ? list.length : MAX_LISTED;
  for (var i = 0; i < n; i++) {
    var email = list[i] && list[i].emailAddress ? list[i].emailAddress : "";
    var shown = email || (list[i] && list[i].displayName) || "(unknown address)";
    parts.push(isInternal(email) ? shown : shown + " (external)");
  }
  var s = parts.join(", ");
  if (list.length > MAX_LISTED) s += ", and " + (list.length - MAX_LISTED) + " more";
  return label + ": " + s;
}

function onMessageSendHandler(event) {
  var settled = false;
  function finish(allow, message) {
    if (settled) return;
    settled = true;
    try {
      if (allow) event.completed({ allowEvent: true });
      else event.completed({ allowEvent: false, errorMessage: message });
    } catch (e) {}
  }

  try { setTimeout(function () { finish(true); }, SAFETY_TIMEOUT_MS); } catch (e) {}

  try {
    var item = Office.context.mailbox.item;
    var buckets = { to: [], cc: [], bcc: [] };
    var pending = 3;
    var failed = false;

    function collect(key) {
      return function (result) {
        try {
          if (result && result.status === Office.AsyncResultStatus.Succeeded) {
            buckets[key] = result.value || [];
          } else { failed = true; }
        } catch (e) { failed = true; }
        pending--;
        if (pending === 0) {
          if (failed) finish(true);
          else evaluate(finish, buckets);
        }
      };
    }

    item.to.getAsync(collect("to"));
    item.cc.getAsync(collect("cc"));
    item.bcc.getAsync(collect("bcc"));
  } catch (e) {
    finish(true);
  }
}

function evaluate(finish, buckets) {
  try {
    var to = buckets.to, cc = buckets.cc, bcc = buckets.bcc;
    var total = to.length + cc.length + bcc.length;

    var externalCount = externalIn(to).length + externalIn(cc).length + externalIn(bcc).length;
    var triggerExternal = externalCount >= 1;
    var triggerMultiple = PROMPT_ON_MULTIPLE_RECIPIENTS && total >= MULTIPLE_RECIPIENT_THRESHOLD;

    if (!triggerExternal && !triggerMultiple) { finish(true); return; }

    var lines = [];
    lines.push("Please double-check who this email is going to.");
    lines.push("");

    var toLine = formatField("To", to);
    var ccLine = formatField("Cc", cc);
    var bccLine = formatField("Bcc", bcc);
    if (toLine) lines.push(toLine);
    if (ccLine) lines.push(ccLine);
    if (bccLine) lines.push(bccLine);

    if (cc.length > 0 || bcc.length > 0) {
      lines.push("");
      lines.push("Ensure that your Cc and Bcc lists are correct. Cc'd users are visible to all addressees, Bcc'd users are not.");
    }

    var closing = "If your email is addressed correctly, select 'Send anyway'. Otherwise, select 'Don't send' and adjust the recipients.";
    var body = lines.join("\n");
    var message = body + "\n\n" + closing;
    if (message.length > 480) {
      var room = 480 - closing.length - 5;
      if (room < 0) { room = 0; }
      body = body.substring(0, room) + "...";
      message = body + "\n\n" + closing;
    }

    finish(false, message);
  } catch (e) {
    finish(true);
  }
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
