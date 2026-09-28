/*
 * Client Email add-in - side panel.
 * The client list, topic and body template are stored in the mailbox's
 * roaming settings. Nothing about clients is stored in these files.
 */

var SETTINGS_KEY = "cet_settings";
var NOTICE_KEY = "cet_pick_client";
var settings = { clients: [], topic: "", bodyText: "" };

function $(id) { return document.getElementById(id); }

function loadSettings() {
  var s = Office.context.roamingSettings.get(SETTINGS_KEY);
  if (!s || typeof s !== "object") s = {};
  settings.clients = Array.isArray(s.clients) ? s.clients : [];
  settings.topic = typeof s.topic === "string" ? s.topic : "";
  settings.bodyText = typeof s.bodyText === "string" ? s.bodyText : "";
}

function setStatus(id, text, kind) {
  var el = $(id);
  el.textContent = text || "";
  el.className = "status" + (kind ? " " + kind : "");
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* ---------- Views ---------- */

function showPick() {
  $("settingsView").hidden = true;
  $("pickView").hidden = false;
  renderClients("");
  $("clientFilter").value = "";
  setStatus("pickStatus", "");
  if (!settings.clients.length) {
    setStatus("pickStatus", "No clients yet. Use “Edit client list & template” below.", "error");
  }
  initTopicFromSubject();
  $("clientFilter").focus();
}

function showSettings() {
  $("pickView").hidden = true;
  $("settingsView").hidden = false;
  $("clientsText").value = settings.clients.join("\n");
  $("defaultTopic").value = settings.topic;
  $("bodyText").value = settings.bodyText;
  setStatus("settingsStatus", "");
  $("clientsText").focus();
}

/* ---------- Pick view ---------- */

function renderClients(filter) {
  var select = $("clientSelect");
  var previous = select.value;
  var f = filter.trim().toLowerCase();
  select.innerHTML = "";
  settings.clients.forEach(function (name) {
    if (f && name.toLowerCase().indexOf(f) === -1) return;
    var opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    select.appendChild(opt);
  });
  if (previous) select.value = previous;
  if (select.selectedIndex < 0 && select.options.length === 1) select.selectedIndex = 0;
  updatePreview();
}

function initTopicFromSubject() {
  $("topicInput").value = settings.topic;
  updatePreview();
  Office.context.mailbox.item.subject.getAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return;
    var subject = (r.value || "").trim();
    if (!subject) return;
    // Strip a placeholder or a client name that was already applied.
    var topic = subject.replace(/\s*-\s*\[Client\]\s*$/i, "");
    settings.clients.forEach(function (name) {
      var suffix = " - " + name;
      if (topic.length > suffix.length &&
          topic.slice(-suffix.length).toLowerCase() === suffix.toLowerCase()) {
        topic = topic.slice(0, -suffix.length);
      } else if (topic.toLowerCase() === name.toLowerCase()) {
        topic = "";
      }
    });
    $("topicInput").value = topic;
    updatePreview();
  });
}

function buildSubject() {
  var client = $("clientSelect").value;
  var topic = $("topicInput").value.trim();
  if (!client) return "";
  return topic ? topic + " - " + client : client;
}

function updatePreview() {
  var s = buildSubject();
  $("subjectPreview").textContent = s || "(pick a client)";
  $("applyBtn").disabled = !$("clientSelect").value;
}

function replaceClientInBody(client, done) {
  var body = Office.context.mailbox.item.body;
  body.getTypeAsync(function (t) {
    var isHtml = t.status === Office.AsyncResultStatus.Succeeded && t.value === Office.CoercionType.Html;
    var coercion = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
    body.getAsync(coercion, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { done(); return; }
      var content = r.value || "";
      if (!/\[Client\]/i.test(content)) { done(); return; }
      var replacement = isHtml ? escapeHtml(client) : client;
      var updated = content.replace(/\[Client\]/gi, function () { return replacement; });
      body.setAsync(updated, { coercionType: coercion }, function () { done(); });
    });
  });
}

function applyClient() {
  var client = $("clientSelect").value;
  if (!client) {
    setStatus("pickStatus", "Please choose a client.", "error");
    return;
  }
  var subject = buildSubject();
  $("applyBtn").disabled = true;
  setStatus("pickStatus", "Applying…");
  var item = Office.context.mailbox.item;
  item.subject.setAsync(subject, function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("pickStatus", "Couldn’t set the subject: " + r.error.message, "error");
      $("applyBtn").disabled = false;
      return;
    }
    replaceClientInBody(client, function () {
      item.notificationMessages.removeAsync(NOTICE_KEY, function () {
        setStatus("pickStatus", "Done: " + subject, "ok");
        $("applyBtn").disabled = false;
        setTimeout(function () {
          try { Office.context.ui.closeContainer(); } catch (e) { /* not supported: leave open */ }
        }, 800);
      });
    });
  });
}

/* ---------- Settings view ---------- */

function saveSettings() {
  var seen = {};
  var clients = $("clientsText").value
    .split(/\r?\n/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) {
      var key = s.toLowerCase();
      if (!s || seen[key]) return false;
      seen[key] = true;
      return true;
    })
    .sort(function (a, b) { return a.localeCompare(b, undefined, { sensitivity: "base" }); });

  var next = {
    clients: clients,
    topic: $("defaultTopic").value.trim(),
    bodyText: $("bodyText").value.replace(/\s+$/, "")
  };

  $("saveBtn").disabled = true;
  setStatus("settingsStatus", "Saving…");
  Office.context.roamingSettings.set(SETTINGS_KEY, next);
  Office.context.roamingSettings.saveAsync(function (r) {
    $("saveBtn").disabled = false;
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("settingsStatus", "Couldn’t save: " + r.error.message, "error");
      return;
    }
    settings = next;
    setStatus("settingsStatus", "Saved " + clients.length + " client" + (clients.length === 1 ? "" : "s") + ".", "ok");
    setTimeout(showPick, 700);
  });
}

/* ---------- Start ---------- */

Office.onReady(function () {
  loadSettings();

  $("clientFilter").addEventListener("input", function (e) { renderClients(e.target.value); });
  $("clientFilter").addEventListener("keydown", function (e) {
    if (e.key === "ArrowDown") { $("clientSelect").focus(); e.preventDefault(); }
    if (e.key === "Enter" && $("clientSelect").options.length) {
      if ($("clientSelect").selectedIndex < 0) $("clientSelect").selectedIndex = 0;
      applyClient();
    }
  });
  $("clientSelect").addEventListener("change", updatePreview);
  $("clientSelect").addEventListener("dblclick", applyClient);
  $("clientSelect").addEventListener("keydown", function (e) { if (e.key === "Enter") applyClient(); });
  $("topicInput").addEventListener("input", updatePreview);
  $("applyBtn").addEventListener("click", applyClient);
  $("openSettings").addEventListener("click", function (e) { e.preventDefault(); showSettings(); });
  $("saveBtn").addEventListener("click", saveSettings);
  $("cancelBtn").addEventListener("click", showPick);

  if (settings.clients.length) showPick(); else showSettings();
});
