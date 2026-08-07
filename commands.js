/*
 * Calgary Counselling Centre — Recipient Check
 * Outlook Smart Alerts add-in (OnMessageSend).
 *
 * When a message is sent, this checks the recipients and — if the message is
 * going to anyone outside the organization — stops the send and shows a short
 * summary the sender must acknowledge.
 *
 * Nothing here is added to the message, so external recipients never see it.
 *
 * ---------------------------------------------------------------------------
 * CONFIGURATION — edit the values below to suit your organization.
 * ---------------------------------------------------------------------------
 */

// Domains that count as "internal". Add every accepted domain your org sends
// from (e.g. a secondary brand domain). Matching is case-insensitive and also
// covers sub-domains (e.g. "mail.calgarycounselling.com").
var INTERNAL_DOMAINS = ["calgarycounselling.com"];

// Set to true to also prompt on messages with many recipients (even if all
// internal). Set to false to prompt ONLY when there is an external recipient.
var PROMPT_ON_MULTIPLE_RECIPIENTS = false;

// Only used when PROMPT_ON_MULTIPLE_RECIPIENTS is true: prompt when the total
// number of recipients is at least this many (2 = more than one recipient).
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

  var triggerExternal =
