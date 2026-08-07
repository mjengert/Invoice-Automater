// Theme toggle — persisted in localStorage (this is a real locally-hosted
// app in the user's own browser, not a sandboxed artifact, so localStorage
// is fine here).
(function () {
  const root = document.documentElement;
  const KEY = "funkys-dashboard-theme";
  const toggle = document.getElementById("themeToggle");
  const label = document.getElementById("themeToggleLabel");

  function apply(theme) {
    root.setAttribute("data-theme", theme);
    if (label) label.textContent = theme === "dark" ? "Light mode" : "Dark mode";
    if (toggle) {
      const moon = toggle.querySelector(".ic-moon");
      const sun = toggle.querySelector(".ic-sun");
      if (moon && sun) {
        moon.style.display = theme === "dark" ? "none" : "";
        sun.style.display = theme === "dark" ? "" : "none";
      }
    }
  }

  const saved = localStorage.getItem(KEY) ||
    (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  apply(saved);

  if (toggle) {
    toggle.addEventListener("click", function () {
      const next = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
      localStorage.setItem(KEY, next);
      apply(next);
    });
  }
})();

// Palette (color theme) picker — a separate axis from light/dark mode.
// Four electrical-company-flavored accent palettes (classic/voltage/
// hazard/copper), each swatch a diagonal preview of that palette's own
// colors. Persisted in localStorage the same way as the light/dark mode.
(function () {
  const root = document.documentElement;
  const KEY = "funkys-dashboard-palette";
  const VALID = ["classic", "voltage", "hazard", "copper"];
  const swatches = document.querySelectorAll(".palette-swatch");

  function apply(palette) {
    root.setAttribute("data-palette", palette);
    swatches.forEach(function (btn) {
      btn.classList.toggle("active", btn.dataset.palette === palette);
    });
  }

  const saved = localStorage.getItem(KEY);
  apply(VALID.includes(saved) ? saved : "classic");

  swatches.forEach(function (btn) {
    btn.addEventListener("click", function () {
      const palette = btn.dataset.palette;
      localStorage.setItem(KEY, palette);
      apply(palette);
    });
  });
})();

// Small fetch helper shared by the wizard / invoices pages.
async function apiFetch(url, opts) {
  const res = await fetch(url, Object.assign({
    headers: { "Content-Type": "application/json" },
  }, opts || {}));
  let data = null;
  try { data = await res.json(); } catch (e) { /* not JSON */ }
  if (!res.ok) {
    const msg = (data && data.error) || `Request failed (${res.status})`;
    throw new Error(msg);
  }
  return data;
}

function money(n) {
  return "$" + Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Prev/next arrows on the Monthly Log + Payroll Log "month switcher" —
// steps the <input type="month"> a month at a time and resubmits its form,
// same as picking a month from the native calendar dropdown would.
function shiftMonthInput(inputId, delta) {
  const input = document.getElementById(inputId);
  if (!input || !input.value) return;
  const [y, m] = input.value.split("-").map(Number);
  const d = new Date(y, m - 1 + delta, 1);
  input.value = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
  input.form.submit();
}

// ── Mark Paid modal (shared across Invoices list + Customer history) ────────
// The modal markup itself lives once in base.html; any page just calls
// openMarkPaidModal() and gets a callback when the user confirms or cancels.
let __markPaidCtx = null;
let __markPaidCategory = null;

function openMarkPaidModal(customer, number, onDone) {
  __markPaidCtx = { mode: "single", customer, number, onDone };
  __markPaidCategory = null;
  document.getElementById("markPaidWho").textContent = customer + " — Invoice #" + number;
  document.getElementById("markPaidDate").value = new Date().toISOString().slice(0, 10);
  document.getElementById("mpBtnG").classList.remove("active");
  document.getElementById("mpBtnE").classList.remove("active");
  document.getElementById("mpConfirmBtn").disabled = true;
  document.getElementById("markPaidModal").classList.add("show");
}

// Bulk variant — same modal, same one date + one G/E category, applied to
// every invoice in `items` (each {customer, number, total}). Used by the
// Invoices list's checkbox multi-select ("Mark Selected Paid").
function openBulkMarkPaidModal(items, onDone) {
  __markPaidCtx = { mode: "bulk", items, onDone };
  __markPaidCategory = null;
  const total = items.reduce((s, it) => s + (it.total || 0), 0);
  document.getElementById("markPaidWho").textContent =
    items.length + " invoice" + (items.length === 1 ? "" : "s") + " selected — " + money(total);
  document.getElementById("markPaidDate").value = new Date().toISOString().slice(0, 10);
  document.getElementById("mpBtnG").classList.remove("active");
  document.getElementById("mpBtnE").classList.remove("active");
  document.getElementById("mpConfirmBtn").disabled = true;
  document.getElementById("markPaidModal").classList.add("show");
}

function closeMarkPaidModal() {
  document.getElementById("markPaidModal").classList.remove("show");
  __markPaidCtx = null;
}

function selectMarkPaidCategory(cat) {
  __markPaidCategory = cat;
  document.getElementById("mpBtnG").classList.toggle("active", cat === "G");
  document.getElementById("mpBtnE").classList.toggle("active", cat === "E");
  document.getElementById("mpConfirmBtn").disabled = false;
}

async function confirmMarkPaid() {
  if (!__markPaidCtx || !__markPaidCategory) return;
  const date = document.getElementById("markPaidDate").value;
  const btn = document.getElementById("mpConfirmBtn");
  btn.disabled = true;

  if (__markPaidCtx.mode === "bulk") {
    const { items, onDone } = __markPaidCtx;
    const results = { ok: [], failed: [] };
    // Sequential, one at a time — each request reuses the same atomic
    // per-invoice mark-paid endpoint the single-item flow already relies
    // on, so a bulk mark-paid can't race or double-count any differently
    // than clicking "Mark Paid" on each row by hand would.
    for (const it of items) {
      try {
        const data = await apiFetch("/api/invoices/" + it.number + "/mark-paid", {
          method: "POST",
          body: JSON.stringify({ customer: it.customer, category: __markPaidCategory, date }),
        });
        results.ok.push(Object.assign({}, it, { month_label: data.month_label }));
      } catch (e) {
        results.failed.push(Object.assign({}, it, { error: e.message }));
      }
    }
    closeMarkPaidModal();
    if (onDone) onDone(null, results);
    return;
  }

  const { customer, number, onDone } = __markPaidCtx;
  try {
    const data = await apiFetch("/api/invoices/" + number + "/mark-paid", {
      method: "POST",
      body: JSON.stringify({ customer, category: __markPaidCategory, date }),
    });
    closeMarkPaidModal();
    if (onDone) onDone(null, data);
  } catch (e) {
    btn.disabled = false;
    if (onDone) onDone(e, null);
  }
}

// ── Bulk Email modal (Invoices list — "Email Selected") ────────────────────
// Sends each selected invoice's PDF as its own individual, immediately-sent
// email to the saved "bulk email address" setting. No manual send step, so
// this modal exists purely as an explicit confirmation gate before firing.
let __bulkEmailCtx = null;

function openBulkEmailModal(items, toAddress, onDone) {
  __bulkEmailCtx = { items, onDone };
  document.getElementById("bulkEmailWho").textContent =
    "Send " + items.length + " invoice" + (items.length === 1 ? "" : "s") + " to " + toAddress + "?";
  document.getElementById("beConfirmBtn").disabled = false;
  document.getElementById("bulkEmailModal").classList.add("show");
}

function closeBulkEmailModal() {
  document.getElementById("bulkEmailModal").classList.remove("show");
  __bulkEmailCtx = null;
}

async function confirmBulkEmail() {
  if (!__bulkEmailCtx) return;
  const { items, onDone } = __bulkEmailCtx;
  const btn = document.getElementById("beConfirmBtn");
  btn.disabled = true;
  try {
    const data = await apiFetch("/api/invoices/bulk-email", {
      method: "POST",
      body: JSON.stringify({ items }),
    });
    closeBulkEmailModal();
    if (onDone) onDone(null, data);
  } catch (e) {
    btn.disabled = false;
    if (onDone) onDone(e, null);
  }
}

// ── Email Report PDF modal (Monthly Log + Payroll Log — shared) ────────────
// Same confirm-before-send pattern as bulk invoice email: emails the
// current month's PDF right away, no review step in Mail.app, so this
// modal exists purely as the explicit confirmation gate before firing.
let __reportEmailCtx = null;

function openReportEmailModal(kind, month, monthLabel, toAddress, onDone) {
  // kind is 'monthly' or 'payroll' — picks the API endpoint and label.
  __reportEmailCtx = { kind, month, onDone };
  const label = kind === "monthly" ? "Monthly Log" : "Payroll Report";
  document.getElementById("reportEmailWho").textContent =
    "Send the " + monthLabel + " " + label + " to " + toAddress + "?";
  document.getElementById("reConfirmBtn").disabled = false;
  document.getElementById("reportEmailModal").classList.add("show");
}

function closeReportEmailModal() {
  document.getElementById("reportEmailModal").classList.remove("show");
  __reportEmailCtx = null;
}

async function confirmReportEmail() {
  if (!__reportEmailCtx) return;
  const { kind, month, onDone } = __reportEmailCtx;
  const btn = document.getElementById("reConfirmBtn");
  btn.disabled = true;
  try {
    const url = kind === "monthly" ? "/api/monthly/email-pdf" : "/api/payroll/email-pdf";
    const data = await apiFetch(url, {
      method: "POST",
      body: JSON.stringify({ month }),
    });
    closeReportEmailModal();
    if (onDone) onDone(null, data);
  } catch (e) {
    btn.disabled = false;
    if (onDone) onDone(e, null);
  }
}

// ── Google Places address autocomplete (Customers / New Document) ─────────
// Wires up every input.address-autocomplete field on the current page.
// This is the `callback=` target of the Google Maps script tag base.html
// adds when a google_places_api_key setting is present — Google calls this
// once the API has loaded, regardless of what page we're on, so it stays
// a no-op on pages with no matching input.
function initGooglePlacesAutocomplete() {
  document.querySelectorAll("input.address-autocomplete").forEach(function (input) {
    if (input.dataset.placesInit) return; // don't double-init on repeat calls
    input.dataset.placesInit = "1";

    const ac = new google.maps.places.Autocomplete(input, {
      types: ["address"],
      componentRestrictions: { country: "us" },
      fields: ["address_components"],
    });

    ac.addListener("place_changed", function () {
      const place = ac.getPlace();
      const comps = place.address_components || [];
      const part = (type, short) => {
        const c = comps.find((c) => c.types.includes(type));
        return c ? (short ? c.short_name : c.long_name) : "";
      };
      const street = [part("street_number"), part("route")].filter(Boolean).join(" ");
      const city = part("locality") || part("sublocality") || part("postal_town");
      const state = part("administrative_area_level_1", true);
      const zip = part("postal_code");

      // Built as "Street, City, ST Zip" — exactly the 3-comma-segment shape
      // normalize_address() (invoice_engine.py) already knows how to split
      // into the stored two-line address, so no backend change is needed.
      if (street && city && state) {
        input.value = street + ", " + city + ", " + state + (zip ? " " + zip : "");
      }
      // If Google couldn't resolve a full address (component missing),
      // leave whatever the user already typed/selected alone rather than
      // overwriting it with a partial result.
    });
  });
}
