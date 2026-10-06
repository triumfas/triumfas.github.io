// Keep in sync with launchevent.js (the event runtime cannot share modules).
const RATES_BY_SUFFIX = [
  [".lt", 25],
  [".com", 30],
  [".dk", 35],
];
const DEFAULT_RATE = 50;
const rateFor = (email) => {
  const e = (email || "").toLowerCase();
  const match = RATES_BY_SUFFIX.find(([suffix]) => e.endsWith(suffix));
  return match ? match[1] : DEFAULT_RATE;
};
const CURRENCY = "EUR";

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
  try {
    const [required, optional, start, end, organizer] = await Promise.all([
      getAsync((cb) => item.requiredAttendees.getAsync(cb), "requiredAttendees"),
      getAsync((cb) => item.optionalAttendees.getAsync(cb), "optionalAttendees"),
      getAsync((cb) => item.start.getAsync(cb), "start"),
      getAsync((cb) => item.end.getAsync(cb), "end"),
      getOrganizer(item),
    ]);

    const people = new Map();
    const add = (name, email) => {
      const key = (email || name || "").toLowerCase();
      if (key && !people.has(key)) people.set(key, { name: name || email, rate: rateFor(email) });
    };
    add(organizer.name, organizer.email);
    [...required, ...optional].forEach((a) => add(a.displayName, a.emailAddress));

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
  } catch (error) {
    const e = error || {};
    $("error").textContent = "Could not calculate: " + ([e.code, e.message].filter(Boolean).join(": ") || "unknown error");
    $("error").hidden = false;
  }
}

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  render();
  // Live updates while the pane stays open.
  item.addHandlerAsync(Office.EventType.RecipientsChanged, render);
  item.addHandlerAsync(Office.EventType.AppointmentTimeChanged, render);
});
