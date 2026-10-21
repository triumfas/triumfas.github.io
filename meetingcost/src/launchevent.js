// Single self-contained file: classic Outlook's JS-only runtime does not support imports.
// Settings logic is duplicated in taskpane.js (the event runtime cannot share modules); keep in sync.
const SETTINGS_KEY = "meetingCostSettings";
const CURRENCY = "DKK";
const HOURLY_RATE = 400;
const DEFAULT_SETTINGS = { includeOptional: false };
function loadSettings() {
  const s = Office.context.roamingSettings.get(SETTINGS_KEY) || {};
  return { includeOptional: typeof s.includeOptional === "boolean" ? s.includeOptional : DEFAULT_SETTINGS.includeOptional };
}
const norm = (v) => (v || "").trim().toLowerCase();

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
    LOG("enhancedLocation unavailable:", describeError(e));
  }
  return { emails, names };
}
const isRoom = (a, rooms) => rooms.emails.has(norm(a.emailAddress)) || rooms.names.has(norm(a.displayName));
const NOTIFICATION_KEY = "meetingCost";

const LOG = (...args) => console.log("[MeetingCost]", ...args);

function describeError(e) {
  if (!e) return "unknown error";
  return [e.code, e.name, e.message].filter(Boolean).join(": ") || JSON.stringify(e);
}

// The timeout turns a callback that never fires into a named error instead of a hung command.
function getAsync(fn, label = "call", timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${label} timed out`)), timeoutMs);
    try {
      fn((result) => {
        clearTimeout(timer);
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    } catch (e) {
      clearTimeout(timer);
      reject(e);
    }
  });
}

// Classic includes the organizer in the attendee lists, web/new Outlook only when editing, so add it explicitly and dedupe.
async function getOrganizerEmail(item) {
  try {
    const organizer = await getAsync((cb) => item.organizer.getAsync(cb), "organizer");
    return organizer.emailAddress;
  } catch (e) {
    return Office.context.mailbox.userProfile.emailAddress;
  }
}

async function calculateCost() {
  const item = Office.context.mailbox.item;
  LOG("calculateCost start, itemType:", item && item.itemType);
  const [required, optional, start, end, organizerEmail, rooms] = await Promise.all([
    getAsync((cb) => item.requiredAttendees.getAsync(cb), "requiredAttendees"),
    getAsync((cb) => item.optionalAttendees.getAsync(cb), "optionalAttendees"),
    getAsync((cb) => item.start.getAsync(cb), "start"),
    getAsync((cb) => item.end.getAsync(cb), "end"),
    getOrganizerEmail(item),
    getRoomIdentities(item),
  ]);

  const settings = loadSettings();
  const people = [...required, ...(settings.includeOptional ? optional : [])].filter((a) => !isRoom(a, rooms));
  const participants = new Set(people.map((a) => norm(a.emailAddress)));
  if (organizerEmail) {
    participants.add(norm(organizerEmail));
  }
  rooms.emails.forEach((e) => participants.delete(e));

  LOG("inputs", { required: required.length, optional: optional.length, rooms: rooms.emails.size, start, end, organizerEmail });

  const hours = Math.max(0, (end.getTime() - start.getTime()) / 3600000);
  const hourlyTotal = participants.size * HOURLY_RATE;
  return { count: participants.size, hours, hourlyTotal, cost: Math.round(hourlyTotal * hours), currency: CURRENCY };
}

async function recalc(event) {
  LOG("recalc fired");
  try {
    const { count, hours, hourlyTotal, cost, currency } = await calculateCost();
    const message =
      `\u26a0\ufe0f Estimated meeting cost: ${currency} ${cost.toLocaleString("en-IE")} ` +
      `(${count} participants, ${Math.round(hours * 100) / 100} h, ${currency} ${hourlyTotal}/h combined)`;

    const notifications = Office.context.mailbox.item.notificationMessages;
    try {
      // Insight banner with a "Details" button that opens the task pane (commandId = ribbon control id in the manifest).
      await getAsync(
        (cb) =>
          notifications.replaceAsync(
            NOTIFICATION_KEY,
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
              message,
              icon: "Icon.16x16",
              actions: [
                {
                  actionText: "Details",
                  actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
                  commandId: "meetingCostPaneButton",
                  contextData: "{}",
                },
              ],
            },
            cb
          ),
        "insightNotification"
      );
      LOG("insight notification set:", message);
    } catch (insightError) {
      LOG("insight notification failed, falling back to informational:", describeError(insightError));
      await getAsync(
        (cb) =>
          notifications.replaceAsync(
            NOTIFICATION_KEY,
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
              message,
              icon: "Icon.16x16",
              // persistent notifications are rejected on unsaved drafts (error 9028)
              persistent: false,
            },
            cb
          ),
        "informationalNotification"
      );
      LOG("informational notification set:", message);
    }
  } catch (error) {
    console.error("[MeetingCost] calculation failed", error);
    try {
      await getAsync((cb) =>
        Office.context.mailbox.item.notificationMessages.replaceAsync(
          NOTIFICATION_KEY,
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            message: ("Meeting cost failed: " + describeError(error)).slice(0, 150),
          },
          cb
        )
      );
    } catch (e2) {
      console.error("[MeetingCost] could not show error notification", e2);
    }
  } finally {
    event.completed();
  }
}

LOG("script loaded, Office:", typeof Office);
// The browser runtime throws if associate() runs before Office.js has initialised, so retry from onReady.
// The classic JS-only runtime never fires onReady for handlers, so the direct call must stay.
let associated = false;
function associate() {
  if (associated) return;
  Office.actions.associate("recalc", recalc);
  associated = true;
  LOG("recalc associated");
}
try {
  associate();
} catch (e) {
  LOG("early associate failed, waiting for onReady:", describeError(e));
}
Office.onReady((info) => {
  LOG("Office.onReady", info);
  associate();
});
