// Settings logic is duplicated in launchevent.js (the event runtime cannot share modules); keep in sync.
const SETTINGS_KEY = "meetingCostSettings";
const DEFAULT_SETTINGS = {
  currency: "EUR",
  defaultRate: 30,
  rules: [
    { suffix: ".lt", rate: 35 },
    { suffix: ".com", rate: 50 },
    { suffix: ".dk", rate: 45 },
  ],
};
function loadSettings() {
  const s = Office.context.roamingSettings.get(SETTINGS_KEY) || {};
  return {
    currency: typeof s.currency === "string" && s.currency.trim() ? s.currency.trim().toUpperCase() : DEFAULT_SETTINGS.currency,
    defaultRate: Number.isFinite(s.defaultRate) ? s.defaultRate : DEFAULT_SETTINGS.defaultRate,
    rules: Array.isArray(s.rules) ? s.rules : DEFAULT_SETTINGS.rules,
  };
}
const rateFor = (email, settings) => {
  const e = (email || "").toLowerCase();
  const match = settings.rules.find((r) => r.suffix && e.endsWith(r.suffix.toLowerCase()));
  return match ? match.rate : settings.defaultRate;
};
let settings;
let savedRuleValues = new WeakMap();

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
  const CURRENCY = settings.currency;
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
      if (key && !people.has(key)) people.set(key, { name: name || email, rate: rateFor(email, settings) });
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

function updateUnsavedState() {
  const currencyChanged = $("currency").value.trim().toUpperCase() !== settings.currency;
  const rateChanged = $("defaultRate").value.trim() === "" || Number($("defaultRate").value) !== settings.defaultRate;
  $("currency").classList.toggle("unsaved", currencyChanged);
  $("defaultRate").classList.toggle("unsaved", rateChanged);
  const rows = [...document.querySelectorAll("#rules .rule")];
  let changed = currencyChanged || rateChanged || rows.length !== settings.rules.length;
  rows.forEach((row) => {
    const [suffix, rate] = row.querySelectorAll("input");
    const saved = savedRuleValues.get(row);
    const suffixChanged = !saved || suffix.value.trim() !== saved.suffix;
    const ruleRateChanged = !saved || rate.value.trim() === "" || Number(rate.value) !== saved.rate;
    suffix.classList.toggle("unsaved", suffixChanged);
    rate.classList.toggle("unsaved", ruleRateChanged);
    changed = changed || suffixChanged || ruleRateChanged;
  });
  $("save").classList.toggle("unsaved", changed);
}

function addRuleRow(suffix = "", rate = "", savedRule) {
  const row = document.createElement("div");
  row.className = "rule";
  const s = document.createElement("input");
  s.type = "text";
  s.placeholder = ".example.com";
  s.value = suffix;
  const r = document.createElement("input");
  r.type = "number";
  r.min = "0";
  r.step = "any";
  r.placeholder = "rate/h";
  r.value = rate;
  const x = document.createElement("button");
  x.type = "button";
  x.textContent = "\u00d7";
  x.title = "Remove";
  x.onclick = () => {
    row.remove();
    updateUnsavedState();
  };
  s.oninput = updateUnsavedState;
  r.oninput = updateUnsavedState;
  row.append(s, r, x);
  if (savedRule) savedRuleValues.set(row, savedRule);
  $("rules").append(row);
  updateUnsavedState();
}

function fillSettingsForm(formSettings = settings) {
  $("currency").value = formSettings.currency;
  $("defaultRate").value = formSettings.defaultRate;
  savedRuleValues = new WeakMap();
  $("rules").replaceChildren();
  formSettings.rules.forEach((rule, index) => addRuleRow(rule.suffix, rule.rate, settings.rules[index]));
  updateUnsavedState();
}

function saveSettings() {
  const savedRules = new Map([...document.querySelectorAll("#rules .rule")]
    .map((row) => {
      const [s, r] = row.querySelectorAll("input");
      return [row, { suffix: s.value.trim(), rate: Number(r.value) }];
    })
    .filter(([, rule]) => rule.suffix && Number.isFinite(rule.rate) && rule.rate >= 0));
  const rules = [...savedRules.values()];
  const defaultRate = Number($("defaultRate").value);
  const status = $("saveStatus");
  const currency = $("currency").value.trim().toUpperCase();
  if (!currency) {
    status.textContent = "Enter a currency.";
    return;
  }
  if (!Number.isFinite(defaultRate) || defaultRate < 0) {
    status.textContent = "Enter a valid default rate.";
    return;
  }
  const next = { currency, defaultRate, rules };
  Office.context.roamingSettings.set(SETTINGS_KEY, next);
  Office.context.roamingSettings.saveAsync((res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      settings = next;
      savedRuleValues = new WeakMap(savedRules);
      if ($("currency").value.trim().toUpperCase() === settings.currency) $("currency").value = settings.currency;
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
  $("currency").oninput = updateUnsavedState;
  $("defaultRate").oninput = updateUnsavedState;
  $("currencyPreset").onchange = () => {
    $("currency").value = $("currencyPreset").value;
    $("currencyPreset").value = "";
    updateUnsavedState();
  };
  $("addRule").onclick = () => addRuleRow();
  $("save").onclick = saveSettings;
  $("reset").onclick = resetSettings;
  render();
  // Live updates while the pane stays open.
  item.addHandlerAsync(Office.EventType.RecipientsChanged, render);
  item.addHandlerAsync(Office.EventType.AppointmentTimeChanged, render);
});
