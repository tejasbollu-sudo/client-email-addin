/*
 * Client Email add-in - runs automatically when a new email is started.
 * Contains NO client information. Clients come from the mailbox's
 * "Tax Clients" contacts folder; topic and body text are stored in the
 * mailbox (roaming settings). Nothing about clients is in these files.
 *
 * Keep this file self-contained (no imports): classic Outlook on Windows
 * loads it directly in a JavaScript-only runtime.
 */

var CET_SETTINGS_KEY = "cet_settings";
var CET_NOTICE_KEY = "cet_pick_client";
var CET_PANE_COMMAND = "cetOpenPaneButton";

function cetGetSettings() {
  var s = null;
  try {
    s = Office.context.roamingSettings.get(CET_SETTINGS_KEY);
  } catch (e) {
    s = null;
  }
  if (!s || typeof s !== "object") s = {};
  if (!Array.isArray(s.clients)) s.clients = [];
  if (typeof s.topic !== "string") s.topic = "";
  if (typeof s.bodyText !== "string") s.bodyText = "";
  return s;
}

function cetEscapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function cetTextToHtml(text) {
  var parts = String(text).replace(/\r\n/g, "\n").split("\n");
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    out.push("<div>" + (parts[i] ? cetEscapeHtml(parts[i]) : "<br>") + "</div>");
  }
  return out.join("") + "<div><br></div>";
}

// Calls back with true only for a brand-new email (not reply/forward).
function cetIsNewMail(callback) {
  var item = Office.context.mailbox.item;
  if (typeof item.getComposeTypeAsync === "function") {
    item.getComposeTypeAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        callback(result.value.composeType === "newMail");
      } else {
        callback(!item.conversationId);
      }
    });
  } else {
    callback(!item.conversationId);
  }
}

function cetSetSubject(settings, next) {
  if (!settings.topic) { next(); return; }
  Office.context.mailbox.item.subject.setAsync(settings.topic + " - [Client]", function () { next(); });
}

function cetAddBody(settings, next) {
  if (!settings.bodyText) { next(); return; }
  var body = Office.context.mailbox.item.body;
  body.getTypeAsync(function (typeResult) {
    var isHtml = typeResult.status === Office.AsyncResultStatus.Succeeded &&
      typeResult.value === Office.CoercionType.Html;
    var content = isHtml ? cetTextToHtml(settings.bodyText) : settings.bodyText + "\n\n";
    body.prependAsync(
      content,
      { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text },
      function () { next(); }
    );
  });
}

function cetShowNotice(settings, next) {
  var message = "Choose the client(s) for this email.";
  Office.context.mailbox.item.notificationMessages.addAsync(
    CET_NOTICE_KEY,
    {
      type: "insightMessage",
      message: message,
      icon: "Icon.16x16",
      actions: [
        {
          actionType: "showTaskPane",
          actionText: "Choose client",
          commandId: CET_PANE_COMMAND,
          contextData: "{}"
        }
      ]
    },
    function () { next(); }
  );
}

function onNewMessageComposeHandler(event) {
  var done = function () { event.completed(); };
  try {
    cetIsNewMail(function (isNew) {
      if (!isNew) { done(); return; }
      var settings = cetGetSettings();
      cetSetSubject(settings, function () {
        cetAddBody(settings, function () {
          cetShowNotice(settings, done);
        });
      });
    });
  } catch (e) {
    done();
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
