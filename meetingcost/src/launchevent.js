// Single self-contained file: classic Outlook's JS-only runtime does not support imports.
const HOURLY_RATE = 50;
const CURRENCY = "EUR";
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
  const [required, optional, start, end, organizerEmail] = await Promise.all([
    getAsync((cb) => item.requiredAttendees.getAsync(cb), "requiredAttendees"),
    getAsync((cb) => item.optionalAttendees.getAsync(cb), "optionalAttendees"),
    getAsync((cb) => item.start.getAsync(cb), "start"),
    getAsync((cb) => item.end.getAsync(cb), "end"),
    getOrganizerEmail(item),
  ]);

  const participants = new Set(
    [...required, ...optional].map((a) => a.emailAddress.toLowerCase())
  );
  if (organizerEmail) {
    participants.add(organizerEmail.toLowerCase());
  }

  LOG("inputs", { required: required.length, optional: optional.length, start, end, organizerEmail });

  const hours = Math.max(0, (end.getTime() - start.getTime()) / 3600000);
  return { count: participants.size, hours, cost: Math.round(participants.size * hours * HOURLY_RATE) };
}

async function recalc(event) {
  LOG("recalc fired");
  try {
    const { count, hours, cost } = await calculateCost();
    const message =
      `Estimated meeting cost: ${CURRENCY} ${cost.toLocaleString("en-IE")} ` +
      `(${count} participants x ${Math.round(hours * 100) / 100} h x ${HOURLY_RATE}/h)`;

    await getAsync((cb) =>
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        NOTIFICATION_KEY,
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message,
          icon: "Icon.16x16",
          // persistent notifications are rejected on unsaved drafts (error 9028)
          persistent: false,
        },
        cb
      )
    );
    LOG("notification set:", message);
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
