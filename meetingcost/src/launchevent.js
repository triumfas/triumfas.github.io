// Single self-contained file: classic Outlook's JS-only runtime does not support imports.
const HOURLY_RATE = 50;
const CURRENCY = "EUR";
const NOTIFICATION_KEY = "meetingCost";

function getAsync(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

// Classic includes the organizer in the attendee lists, web/new Outlook only when editing, so add it explicitly and dedupe.
async function getOrganizerEmail(item) {
  try {
    const organizer = await getAsync((cb) => item.organizer.getAsync(cb));
    return organizer.emailAddress;
  } catch (e) {
    return Office.context.mailbox.userProfile.emailAddress;
  }
}

async function calculateCost() {
  const item = Office.context.mailbox.item;
  const [required, optional, start, end, organizerEmail] = await Promise.all([
    getAsync((cb) => item.requiredAttendees.getAsync(cb)),
    getAsync((cb) => item.optionalAttendees.getAsync(cb)),
    getAsync((cb) => item.start.getAsync(cb)),
    getAsync((cb) => item.end.getAsync(cb)),
    getOrganizerEmail(item),
  ]);

  const participants = new Set(
    [...required, ...optional].map((a) => a.emailAddress.toLowerCase())
  );
  if (organizerEmail) {
    participants.add(organizerEmail.toLowerCase());
  }

  const hours = Math.max(0, (end.getTime() - start.getTime()) / 3600000);
  return { count: participants.size, hours, cost: Math.round(participants.size * hours * HOURLY_RATE) };
}

async function recalc(event) {
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
          persistent: true,
        },
        cb
      )
    );
  } catch (error) {
    console.error("Meeting cost calculation failed", error);
  } finally {
    event.completed();
  }
}

Office.actions.associate("recalc", recalc);
