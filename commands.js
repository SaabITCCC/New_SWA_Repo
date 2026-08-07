/*
 * Calgary Counselling Centre — Recipient Check
 * Outlook Smart Alerts add-in (OnMessageSend).
 *
 * When a message is sent, this checks the recipients and — if the message is
 * going to anyone outside the organization OR to more than one recipient —
 * stops the send and shows a short summary the sender must acknowledge.
 *
 * Nothing here is added to the message, so external recipients never see it.
 *
 * ---------------------------------------------------------------------------
 * CONFIGURATION — edit the three values below to suit your organization.
 * ---------------------------------------------------------------------------
 */

// Domains that count as "internal". Add every accepted domain your org sends
// from (e.g. a secondary brand domain). Matching is case-insensitive and also
// covers sub-domains (e.g. "mail.calgarycounselling.com").
var INTERNAL_DOMAINS = ["calgarycounselling.com"];

// Prompt when the total number of recipients is at least this many.
// 2 = prompt whenever there is more than one recipient. Raise it (e.g. 5) if
// you only want to warn on larger audiences.
var MULTIPLE_RECIPIENT_THRESHOLD = 2;

// How many external addresses to spell out in the dialog before truncating.
var MAX_LISTED = 8;

/* ------------------------------------------------------------------------- */

Office.onReady(function () {
  // No initialization needed for event-based activation.
});

/**
 * Main handler wired to the manifest's LaunchEvent (FunctionName).
 */
function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var buckets = { to: [], cc: [], bcc: [] };
  var pending = 3;
  var failed = false;

  function collect(key) {
    return function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        buckets[key] = result.value || [];
      } else {
        failed = true;
      }
      pending--;
      if (pending === 0) {
        // If we couldn't read the recipients for some reason, fail open
        // (allow the send) rather than trapping the user.
        if (failed) {
          event.completed({ allowEvent: true });
        } else {
          evaluate(event, buckets);
        }
      }
    };
  }

  item.to.getAsync(collect("to"));
  item.cc.getAsync(collect("cc"));
  item.bcc.getAsync(collect("bcc"));
}

/**
 * Return the lowercased domain portion of an email address.
 */
function domainOf(address) {
  if (!address) return "";
  var at = address.lastIndexOf("@");
  if (at === -1) return "";
  return address.substring(at + 1).toLowerCase().trim();
}

/**
 * True when an address belongs to one of the configured internal domains.
 * Unknown / malformed addresses are treated as external (the safer default).
 */
function isInternal(address) {
  var domain = domainOf(address);
  if (!domain) return false;
  for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
    var internal = INTERNAL_DOMAINS[i].toLowerCase();
    if (domain === internal || domain.endsWith("." + internal)) {
      return true;
    }
  }
  return false;
}

/**
 * Decide whether to prompt, and build the message shown to the sender.
 */
function evaluate(event, buckets) {
  var all = [].concat(buckets.to, buckets.cc, buckets.bcc);
  var total = all.length;
  var bccCount = buckets.bcc.length;

  var external = [];
  for (var i = 0; i < all.length; i++) {
    var email = all[i].emailAddress || "";
    if (!isInternal(email)) {
      external.push(email || (all[i].displayName || "(unknown address)"));
    }
  }

  var triggerExternal = external.length >= 1;
  var triggerMultiple = total >= MULTIPLE_RECIPIENT_THRESHOLD;

  // Neither condition met — let it send with no interruption.
  if (!triggerExternal && !triggerMultiple) {
    event.completed({ allowEvent: true });
    return;
  }

  var lines = [];
  lines.push("Please double-check who this email is going to.");
  lines.push("");
  lines.push("Recipients: " + total + (bccCount ? " (" + bccCount + " in Bcc)" : ""));

  if (external.length > 0) {
    var listed = external.slice(0, MAX_LISTED).join(", ");
    if (external.length > MAX_LISTED) {
      listed += ", and " + (external.length - MAX_LISTED) + " more";
    }
    lines.push("Outside Calgary Counselling (" + external.length + "): " + listed);
  } else {
    lines.push("All recipients are internal.");
  }

  lines.push("");
  lines.push("If this is correct, choose Send anyway. Otherwise choose Don't send to fix the recipients.");

  var message = lines.join("\n");
  // Smart Alerts caps the message length; keep it comfortably short.
  if (message.length > 480) {
    message = message.substring(0, 477) + "...";
  }

  event.completed({
    allowEvent: false,
    errorMessage: message
  });
}

// Register the handler for event-based activation.
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
