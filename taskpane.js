/*
 * Client Email add-in - side panel (v2).
 * Clients are read live from the signed-in mailbox's "Tax Clients" contacts
 * folder via Microsoft Graph. Topic and body template are stored in the
 * mailbox (roaming settings). Nothing about clients is stored in these files.
 */

/* ================= CONFIG ================= */
var CLIENT_ID = "71615059-79c7-49a8-baef-40a21a80f532";
var TENANT_ID = "3047d88f-75f5-431b-b373-7cd08d655a4c";
var FOLDER_NAME = "Tax Clients";
var FALLBACK_REDIRECT = "https://tejasbollu-sudo.github.io/client-email-addin/index.html";
/* ========================================== */

var SETTINGS_KEY = "cet_settings";
var NOTICE_KEY = "cet_pick_client";
var SCOPES = ["Contacts.Read"];
var GRAPH = "https://graph.microsoft.com/v1.0";

var settings = { topic: "", bodyText: "" };
var clients = [];          // [{ id, name, email }]
var selected = [];         // client ids, in the order picked
var pca = null;

function $(id) { return document.getElementById(id); }

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

function joinNames(names) {
  if (names.length <= 1) return names.join("");
  if (names.length === 2) return names[0] + " & " + names[1];
  return names.slice(0, -1).join(", ") + " & " + names[names.length - 1];
}

function byId(id) {
  for (var i = 0; i < clients.length; i++) if (clients[i].id === id) return clients[i];
  return null;
}

function officeAsync(fn) {
  return new Promise(function (resolve, reject) {
    fn(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(new Error(r.error ? r.error.message : "Outlook request failed"));
    });
  });
}

/* ---------- Settings (topic + body) ---------- */

function loadSettings() {
  var s = Office.context.roamingSettings.get(SETTINGS_KEY);
  if (!s || typeof s !== "object") s = {};
  settings.topic = typeof s.topic === "string" ? s.topic : "";
  settings.bodyText = typeof s.bodyText === "string" ? s.bodyText : "";
}

function saveSettings() {
  var next = {
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
    setStatus("settingsStatus", "Saved.", "ok");
    setTimeout(showPick, 600);
  });
}

/* ---------- Sign-in (nested app authentication) ---------- */

async function getPca() {
  if (pca) return pca;
  pca = await msal.createNestablePublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/" + TENANT_ID,
      redirectUri: FALLBACK_REDIRECT
    }
  });
  return pca;
}

async function getToken(interactive) {
  var app = await getPca();
  var request = { scopes: SCOPES };
  if (!interactive) {
    var result = await app.acquireTokenSilent(request);
    return result.accessToken;
  }
  var popup = await app.acquireTokenPopup(request);
  return popup.accessToken;
}

/* ---------- Microsoft Graph ---------- */

async function graphGet(token, url) {
  var res = await fetch(url.indexOf("https://") === 0 ? url : GRAPH + url, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!res.ok) {
    var detail = "";
    try { detail = (await res.json()).error.message; } catch (e) { /* ignore */ }
    throw new Error("Microsoft Graph error " + res.status + (detail ? ": " + detail : ""));
  }
  return res.json();
}

async function findFolder(token) {
  var filter = "?$filter=" + encodeURIComponent("displayName eq '" + FOLDER_NAME.replace(/'/g, "''") + "'");
  var direct = await graphGet(token, "/me/contactFolders" + filter);
  if (direct.value && direct.value.length) return direct.value[0].id;

  // Look one and two levels down, in case the folder is nested.
  var top = await graphGet(token, "/me/contactFolders?$top=100&$select=id,displayName");
  var queue = (top.value || []).map(function (f) { return { id: f.id, depth: 1 }; });
  while (queue.length) {
    var f = queue.shift();
    var kids = await graphGet(token, "/me/contactFolders/" + f.id + "/childFolders?$top=100&$select=id,displayName");
    for (var i = 0; i < (kids.value || []).length; i++) {
      var k = kids.value[i];
      if ((k.displayName || "").toLowerCase() === FOLDER_NAME.toLowerCase()) return k.id;
      if (f.depth < 2) queue.push({ id: k.id, depth: f.depth + 1 });
    }
  }
  return null;
}

async function loadContacts(token) {
  var folderId = await findFolder(token);
  if (!folderId) throw new Error("NOFOLDER");
  var url = "/me/contactFolders/" + folderId +
    "/contacts?$select=id,displayName,companyName,emailAddresses&$top=500";
  var out = [];
  while (url) {
    var page = await graphGet(token, url);
    (page.value || []).forEach(function (c) {
      var email = (c.emailAddresses && c.emailAddresses[0] && c.emailAddresses[0].address) || "";
      var name = (c.displayName || c.companyName || email || "").trim();
      if (name) out.push({ id: c.id, name: name, email: email.trim() });
    });
    url = page["@odata.nextLink"] || null;
  }
  out.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { sensitivity: "base" }); });
  return out;
}

async function loadClients(interactive) {
  $("signInBtn").hidden = true;
  $("pickControls").hidden = true;
  setStatus("loadStatus", "Loading clients from “" + FOLDER_NAME + "”…");
  var token;
  try {
    token = await getToken(interactive);
  } catch (e) {
    if (!interactive) {
      setStatus("loadStatus", "Sign in once to let the add-in read your “" + FOLDER_NAME + "” contacts.");
      $("signInBtn").hidden = false;
      return;
    }
    setStatus("loadStatus", "Sign-in didn’t complete: " + (e.message || e), "error");
    $("signInBtn").hidden = false;
    return;
  }
  try {
    clients = await loadContacts(token);
  } catch (e) {
    if (e.message === "NOFOLDER") {
      setStatus("loadStatus", "Couldn’t find a contacts folder named “" + FOLDER_NAME + "” in this mailbox.", "error");
    } else {
      setStatus("loadStatus", "Couldn’t load clients: " + (e.message || e), "error");
    }
    return;
  }
  selected = selected.filter(function (id) { return byId(id); });
  if (!clients.length) {
    setStatus("loadStatus", "The “" + FOLDER_NAME + "” folder has no contacts yet.", "error");
    return;
  }
  setStatus("loadStatus", "");
  $("pickControls").hidden = false;
  renderList();
  renderSelection();
  $("clientFilter").focus();
}

/* ---------- Pick view ---------- */

function renderList() {
  var f = $("clientFilter").value.trim().toLowerCase();
  var list = $("clientList");
  list.innerHTML = "";
  var shown = 0;
  clients.forEach(function (c) {
    if (f && c.name.toLowerCase().indexOf(f) === -1 && c.email.toLowerCase().indexOf(f) === -1) return;
    shown++;
    var row = document.createElement("label");
    row.className = "client-row";
    var box = document.createElement("input");
    box.type = "checkbox";
    box.checked = selected.indexOf(c.id) !== -1;
    box.addEventListener("change", function () { toggle(c.id, box.checked); });
    var text = document.createElement("span");
    text.className = "client-text";
    var nm = document.createElement("span");
    nm.className = "client-name";
    nm.textContent = c.name;
    var em = document.createElement("span");
    em.className = "client-email";
    em.textContent = c.email || "(no email)";
    text.appendChild(nm);
    text.appendChild(em);
    row.appendChild(box);
    row.appendChild(text);
    list.appendChild(row);
  });
  if (!shown) {
    var empty = document.createElement("div");
    empty.className = "client-empty";
    empty.textContent = "No matches.";
    list.appendChild(empty);
  }
}

function toggle(id, on) {
  var i = selected.indexOf(id);
  if (on && i === -1) selected.push(id);
  if (!on && i !== -1) selected.splice(i, 1);
  renderSelection();
}

function selectedClients() {
  return selected.map(byId).filter(Boolean);
}

function renderSelection() {
  var sel = selectedClients();
  $("selectedBox").hidden = !sel.length;
  $("selectedNames").textContent = joinNames(sel.map(function (c) { return c.name; }));

  // Hints: other clients sharing an email with a selected client.
  var hints = $("sharedHints");
  hints.innerHTML = "";
  var seenEmails = {};
  sel.forEach(function (c) {
    var key = c.email.toLowerCase();
    if (!key || seenEmails[key]) return;
    seenEmails[key] = true;
    var others = clients.filter(function (o) {
      return o.email.toLowerCase() === key && selected.indexOf(o.id) === -1;
    });
    if (!others.length) return;
    var box = document.createElement("div");
    box.className = "shared-hint";
    box.appendChild(document.createTextNode(
      joinNames(others.map(function (o) { return o.name; })) +
      (others.length === 1 ? " also uses " : " also use ") + c.email + "."));
    box.appendChild(document.createElement("br"));
    var btn = document.createElement("button");
    btn.textContent = others.length === 1 ? "Add them" : "Add all";
    btn.addEventListener("click", function () {
      others.forEach(function (o) { if (selected.indexOf(o.id) === -1) selected.push(o.id); });
      renderList();
      renderSelection();
    });
    box.appendChild(btn);
    hints.appendChild(box);
  });

  updatePreview();
}

function buildSubject() {
  var sel = selectedClients();
  if (!sel.length) return "";
  var names = joinNames(sel.map(function (c) { return c.name; }));
  var topic = $("topicInput").value.trim();
  return topic ? topic + " - " + names : names;
}

function updatePreview() {
  var s = buildSubject();
  $("subjectPreview").textContent = s || "(pick a client)";
  $("applyBtn").disabled = !selectedClients().length;
}

function isAllClientNames(text) {
  var parts = text.split(/\s*(?:,|&)\s*/);
  if (!parts.length) return false;
  return parts.every(function (p) {
    var q = p.trim().toLowerCase();
    return q && clients.some(function (c) { return c.name.toLowerCase() === q; });
  });
}

async function initTopicFromSubject() {
  $("topicInput").value = settings.topic;
  updatePreview();
  try {
    var subject = ((await officeAsync(function (cb) { Office.context.mailbox.item.subject.getAsync(cb); })) || "").trim();
    if (!subject) return;
    var topic = subject.replace(/\s*-\s*\[Client\]\s*$/i, "");
    var cut = topic.lastIndexOf(" - ");
    if (topic === subject && cut > -1 && isAllClientNames(topic.slice(cut + 3))) {
      topic = topic.slice(0, cut);
    }
    $("topicInput").value = topic;
    updatePreview();
  } catch (e) { /* keep default topic */ }
}

async function replaceClientInBody(namesText) {
  var body = Office.context.mailbox.item.body;
  var type = await officeAsync(function (cb) { body.getTypeAsync(cb); });
  var isHtml = type === Office.CoercionType.Html;
  var coercion = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
  var content = await officeAsync(function (cb) { body.getAsync(coercion, cb); });
  if (!/\[Client\]/i.test(content || "")) return;
  var replacement = isHtml ? escapeHtml(namesText) : namesText;
  var updated = content.replace(/\[Client\]/gi, function () { return replacement; });
  await officeAsync(function (cb) { body.setAsync(updated, { coercionType: coercion }, cb); });
}

async function addRecipients(sel) {
  var item = Office.context.mailbox.item;
  var existing = await officeAsync(function (cb) { item.to.getAsync(cb); });
  var have = {};
  (existing || []).forEach(function (r) { if (r.emailAddress) have[r.emailAddress.toLowerCase()] = true; });
  var toAdd = [];
  sel.forEach(function (c) {
    var key = c.email.toLowerCase();
    if (key && !have[key]) { have[key] = true; toAdd.push(c.email); }
  });
  if (toAdd.length) await officeAsync(function (cb) { item.to.addAsync(toAdd, cb); });
  return toAdd.length;
}

async function applyClients() {
  var sel = selectedClients();
  if (!sel.length) { setStatus("pickStatus", "Please choose at least one client.", "error"); return; }
  var subject = buildSubject();
  var namesText = joinNames(sel.map(function (c) { return c.name; }));
  $("applyBtn").disabled = true;
  setStatus("pickStatus", "Applying…");
  var item = Office.context.mailbox.item;
  try {
    await officeAsync(function (cb) { item.subject.setAsync(subject, cb); });
    var added = await addRecipients(sel);
    await replaceClientInBody(namesText);
    try { await officeAsync(function (cb) { item.notificationMessages.removeAsync(NOTICE_KEY, cb); }); } catch (e) { /* no notice */ }
    var noEmail = sel.filter(function (c) { return !c.email; }).map(function (c) { return c.name; });
    var msg = "Done: " + subject + (added ? " · " + added + " address" + (added === 1 ? "" : "es") + " added to To" : "");
    if (noEmail.length) msg += ". No email saved for " + joinNames(noEmail) + ".";
    setStatus("pickStatus", msg, "ok");
    if (!noEmail.length) {
      setTimeout(function () {
        try { Office.context.ui.closeContainer(); } catch (e) { /* leave open */ }
      }, 900);
    }
  } catch (e) {
    setStatus("pickStatus", "Something went wrong: " + (e.message || e), "error");
  }
  $("applyBtn").disabled = false;
}

/* ---------- Views ---------- */

function showPick() {
  $("settingsView").hidden = true;
  $("pickView").hidden = false;
  setStatus("pickStatus", "");
  initTopicFromSubject();
}

function showSettings() {
  $("pickView").hidden = true;
  $("settingsView").hidden = false;
  $("defaultTopic").value = settings.topic;
  $("bodyText").value = settings.bodyText;
  setStatus("settingsStatus", "");
  $("defaultTopic").focus();
}

/* ---------- Start ---------- */

Office.onReady(function () {
  loadSettings();

  $("clientFilter").addEventListener("input", renderList);
  $("clientFilter").addEventListener("keydown", function (e) {
    if (e.key !== "Enter") return;
    var boxes = $("clientList").querySelectorAll("input[type=checkbox]");
    if (boxes.length === 1) {
      boxes[0].checked = !boxes[0].checked;
      boxes[0].dispatchEvent(new Event("change"));
    } else if (selectedClients().length) {
      applyClients();
    }
  });
  $("clearSel").addEventListener("click", function (e) {
    e.preventDefault(); selected = []; renderList(); renderSelection();
  });
  $("topicInput").addEventListener("input", updatePreview);
  $("applyBtn").addEventListener("click", applyClients);
  $("signInBtn").addEventListener("click", function () { loadClients(true); });
  $("refreshLink").addEventListener("click", function (e) { e.preventDefault(); showPick(); loadClients(false); });
  $("openSettings").addEventListener("click", function (e) { e.preventDefault(); showSettings(); });
  $("saveBtn").addEventListener("click", saveSettings);
  $("cancelBtn").addEventListener("click", showPick);

  showPick();
  loadClients(false);
});
