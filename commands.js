/*
 * Calgary Counselling Centre — Recipient Check
 * Outlook Smart Alerts add-in (OnMessageSend).
 * Hardened: any error or stall resolves quickly instead of hanging.
 */

// Domains that count as "internal" (case-insensitive, includes sub-domains).
var INTERNAL_DOMAINS = ["calgarycounselling.com"];

// false = prompt ONLY when there is an external recipient.
var PROMPT_ON_MULTIPLE_RECIPIENTS = false;
var MULTIPLE_RECIPIENT_THRESHOLD = 2;

// How many external addresses to list in the dialog before truncating.
var MAX_LISTED = 8;

// When true, if there is an EXTERNAL recipient in the Cc field, the dialog also
// asks whether they should be in Bcc instead (so external people can't see the
// other recipients' email addresses). Internal-only Cc does not trigger this.
var SUGGEST_BCC_WHEN_CC_USED = true;

// If reading recipients stalls this long (ms), allow the send rather than
// leaving the sender stuck. Normal reads take a few milliseconds.
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

// Return the external addresses found in a single recipient list (To/Cc/Bcc).
function externalIn(list) {
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var email = list[i] && list[i].emailAddress ? list[i].emailAddress : "";
    if (!isInternal(email)) {
      out.push(email || (list[i] && list[i].displayName) || "(unknown address)");
    }
  }
  return out;
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
    var total = buckets.to.length + buckets.cc.length + buckets.bcc.length;
    var bccCount = buckets.bcc.length;

    var ccExternal = externalIn(buckets.cc);
    var external = [].concat(externalIn(buckets.to), ccExternal, externalIn(buckets.bcc));

    var triggerExternal = external.length >= 1;
    var triggerMultiple = PROMPT_ON_MULTIPLE_RECIPIENTS && total >= MULTIPLE_RECIPIENT_THRESHOLD;
    var triggerCc = SUGGEST_BCC_WHEN_CC_USED && ccExternal.length >= 1;

    if (!triggerExternal && !triggerMultiple && !triggerCc) { finish(true); return; }

    var lines = [];
    lines.push("Please double-check who this email is going to.");
    lines.push("");
    lines.push("Recipients: " + total + (bccCount ? " (" + bccCount + " in Bcc)" : ""));

    if (external.length > 0) {
      var listed = external.slice(0, MAX_LISTED).join(", ");
      if (external.length > MAX_LISTED) listed += ", and " + (external.length - MAX_LISTED) + " more";
      lines.push("Outside Calgary Counselling (" + external.length + "): " + listed);
    } else {
      lines.push("All recipients are internal.");
    }

    if (triggerCc) {
      lines.push("");
      lines.push(ccExternal.length + (ccExternal.length === 1 ? " external recipient is" : " external recipients are")
        + " in Cc, visible to everyone. Should they be in Bcc instead, so recipients can't see each other's addresses?");
    }

    var closing = "If this is correct, choose 'Send anyway'. Otherwise choose 'Don't send' to fix the recipients.";
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
