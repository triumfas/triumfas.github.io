// Settings logic is duplicated in launchevent.js (the event runtime cannot share modules); keep in sync.
const SETTINGS_KEY = "meetingCostSettings";
const CURRENCY = "DKK";
const HOURLY_RATE = 400;
// Set to the real feedback page URL when available; the button stays disabled while empty.
const FEEDBACK_URL = "";
const DEFAULT_SETTINGS = { includeOptional: false };
function loadSettings() {
  const s = Office.context.roamingSettings.get(SETTINGS_KEY) || {};
  return { includeOptional: typeof s.includeOptional === "boolean" ? s.includeOptional : DEFAULT_SETTINGS.includeOptional };
}
const norm = (v) => (v || "").trim().toLowerCase();
let settings;

const $ = (id) => document.getElementById(id);
const fmt = (n) => n.toLocaleString("en-IE");

function getAsync(fn, label, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${label} timed out`)), timeoutMs);
    fn((result) => {
      clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

// Rooms are recognised via the meeting's structured location (rooms picked in Room Finder / Teams rooms are Room-type locations).
async function getRoomIdentities(item) {
  const emails = new Set();
  const names = new Set();
  try {
    const locations = await getAsync((cb) => item.enhancedLocation.getAsync(cb), "enhancedLocation");
    (locations || []).forEach((l) => {
      if (l.locationIdentifier && l.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
        if (l.emailAddress) emails.add(norm(l.emailAddress));
        if (l.displayName) names.add(norm(l.displayName));
      }
    });
  } catch (e) {
    console.log("[MeetingCost] enhancedLocation unavailable", e);
  }
  return { emails, names };
}
// Room mailboxes are named like "103 name [50]/name": "[capacity]/CC-Site".
const ROOM_NAME_PATTERN = /\[\d+\]\s*\/\s*[A-Za-z]{2}-/;
const isRoom = (a, rooms) =>
  rooms.emails.has(norm(a.emailAddress)) || rooms.names.has(norm(a.displayName)) || ROOM_NAME_PATTERN.test(a.displayName || "");

async function getOrganizer(item) {
  try {
    const o = await getAsync((cb) => item.organizer.getAsync(cb), "organizer");
    return { name: o.displayName, email: o.emailAddress };
  } catch (e) {
    const p = Office.context.mailbox.userProfile;
    return { name: p.displayName, email: p.emailAddress };
  }
}

async function render() {
  const item = Office.context.mailbox.item;
  const activeSettings = settings;
  try {
    const [required, optional, start, end, organizer, rooms] = await Promise.all([
      getAsync((cb) => item.requiredAttendees.getAsync(cb), "requiredAttendees"),
      getAsync((cb) => item.optionalAttendees.getAsync(cb), "optionalAttendees"),
      getAsync((cb) => item.start.getAsync(cb), "start"),
      getAsync((cb) => item.end.getAsync(cb), "end"),
      getOrganizer(item),
      getRoomIdentities(item),
    ]);

    const people = new Map();
    const add = (name, email) => {
      if (isRoom({ displayName: name, emailAddress: email }, rooms)) return;
      const key = norm(email || name);
      if (key && !people.has(key)) people.set(key, { name: name || email, rate: HOURLY_RATE });
    };
    add(organizer.name, organizer.email);
    [...required, ...(activeSettings.includeOptional ? optional : [])].forEach((a) => add(a.displayName, a.emailAddress));

    const hours = Math.max(0, (end.getTime() - start.getTime()) / 3600000);
    const hourlyTotal = [...people.values()].reduce((sum, p) => sum + p.rate, 0);
    const total = Math.round(hourlyTotal * hours);
    const average = people.size ? Math.round(total / people.size) : 0;
    const hoursText = Math.round(hours * 100) / 100;

    $("amount").textContent = `${CURRENCY} ${fmt(total)}`;
    $("formula").textContent = `${people.size} people × ${hoursText} h · ${CURRENCY} ${hourlyTotal}/h combined`;
    $("people").textContent = people.size;
    $("hours").textContent = hoursText;
    $("perPerson").textContent = `${CURRENCY} ${fmt(average)}`;

    const list = $("list");
    list.replaceChildren(
      ...[...people.values()].map((p) => {
        const li = document.createElement("li");
        const a = document.createElement("span");
        a.textContent = p.name;
        const b = document.createElement("span");
        b.textContent = `${CURRENCY} ${fmt(Math.round(p.rate * hours))} (${p.rate}/h)`;
        li.append(a, b);
        return li;
      })
    );
    $("error").hidden = true;
    const message =
      `\u26a0\ufe0f Estimated meeting cost: ${CURRENCY} ${fmt(total)} ` +
      `(${people.size} participants, ${hoursText} h, ${CURRENCY} ${hourlyTotal}/h combined)`;
    try {
      await updateBanner(item, message);
    } catch (error) {
      $("error").textContent = "Could not update meeting banner: " + (error.message || "unknown error");
      $("error").hidden = false;
    }
  } catch (error) {
    const e = error || {};
    $("error").textContent = "Could not calculate: " + ([e.code, e.message].filter(Boolean).join(": ") || "unknown error");
    $("error").hidden = false;
  }
}

async function updateBanner(item, message) {
  try {
    await getAsync((cb) => item.notificationMessages.replaceAsync("meetingCost", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message,
      icon: "Icon.16x16",
      actions: [{
        actionText: "Details",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: "meetingCostPaneButton",
        contextData: "{}",
      }],
    }, cb), "insightNotification");
  } catch (error) {
    await getAsync((cb) => item.notificationMessages.replaceAsync("meetingCost", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.16x16",
      persistent: false,
    }, cb), "informationalNotification");
  }
}

function updateUnsavedState() {
  const changed = $("includeOptional").checked !== settings.includeOptional;
  $("includeOptional").classList.toggle("unsaved", changed);
  $("save").classList.toggle("unsaved", changed);
}

function fillSettingsForm(formSettings = settings) {
  $("includeOptional").checked = formSettings.includeOptional;
  updateUnsavedState();
}

function saveSettings() {
  const next = { includeOptional: $("includeOptional").checked };
  const status = $("saveStatus");
  Office.context.roamingSettings.set(SETTINGS_KEY, next);
  Office.context.roamingSettings.saveAsync((res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      settings = next;
      updateUnsavedState();
      status.textContent = "Saved.";
      render();
    } else {
      status.textContent = "Save failed: " + (res.error && res.error.message);
    }
  });
}

function resetSettings() {
  fillSettingsForm(DEFAULT_SETTINGS);
  $("saveStatus").textContent = "Defaults restored - press Save to apply.";
}

Office.onReady(() => {
  settings = loadSettings();
  const item = Office.context.mailbox.item;
  fillSettingsForm();
  $("includeOptional").onchange = updateUnsavedState;
  $("save").onclick = saveSettings;
  $("reset").onclick = resetSettings;
  $("feedback").disabled = !FEEDBACK_URL;
  $("feedback").title = FEEDBACK_URL ? "Send feedback" : "Feedback link coming soon";
  $("feedback").onclick = () => {
    if (FEEDBACK_URL) window.open(FEEDBACK_URL, "_blank");
  };
  render();
  // Live updates while the pane stays open.
  item.addHandlerAsync(Office.EventType.RecipientsChanged, render);
  item.addHandlerAsync(Office.EventType.AppointmentTimeChanged, render);
});
