// FILE: /Resale-Application/Resale Application.en-US.customjs.js
// PURPOSE:
// - Canonicalize URL to ?id=<applicationGuid>
// - Populate application details
// - Filter Applicants grid by Application Number and rewrite links
// - Financial Assets editable grids + save
// - Fix lookup binding for vms_financialassets using associatednavigationproperty annotations
//
// Notes:
// - Safe to include on other pages (modules auto-disable if required DOM not present)

(() => {
  "use strict";

  // Polyfill: CSS.escape (used for attribute selectors in older browsers)
  if (!window.CSS) window.CSS = {};
  if (!window.CSS.escape) {
    window.CSS.escape = function (value) {
      return String(value).replace(/[^a-zA-Z0-9_\-]/g, function (ch) {
        const hex = ch.codePointAt(0).toString(16).toUpperCase();
        return "\\" + hex + " ";
      });
    };
  }


  /* ===================================================================
   * Shared helpers
   * =================================================================== */

  const GUID_RX =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  const qs = (sel, root = document) => root.querySelector(sel);
  const qsa = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const byId = (id) => document.getElementById(id);

  const norm = (s) =>
    String(s ?? "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const parseGuid = (raw) => {
    const s = norm(raw).replace(/[{}()]/g, "");
    return GUID_RX.test(s) ? s.toLowerCase() : "";
  };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /* ===================================================================
   * Portal API (Power Pages Web API)
   * =================================================================== */

  const PortalApi = (() => {
    const TOKEN_INPUT_NAME = "__RequestVerificationToken";
    const TOKEN_META_NAME = "pp-request-verification-token";

    const HDR_PRIMARY = "__RequestVerificationToken";
    const HDR_FALLBACK = "RequestVerificationToken";

    async function getShellToken() {
      if (window.shell?.getTokenDeferred) {
        try {
          return await new Promise((resolve) =>
            window.shell.getTokenDeferred().done(resolve).fail(() => resolve(""))
          );
        } catch {
          return "";
        }
      }
      return "";
    }

    async function ensureTokenMeta() {
      const hidden = qs(`input[name="${TOKEN_INPUT_NAME}"]`)?.value || "";
      const shellTok = await getShellToken();
      const tok = shellTok || hidden || "";

      let m = qs(`meta[name="${TOKEN_META_NAME}"]`);
      if (!m) {
        m = document.createElement("meta");
        m.name = TOKEN_META_NAME;
        document.head.appendChild(m);
      }
      m.content = tok;
      return tok;
    }

    async function getToken() {
      await ensureTokenMeta();
      return (
        qs(`input[name="${TOKEN_INPUT_NAME}"]`)?.value ||
        qs(`meta[name="${TOKEN_META_NAME}"]`)?.content ||
        (await getShellToken()) ||
        ""
      );
    }

    function parseDataverseError(text) {
      if (!text) return { message: "" };
      try {
        const j = JSON.parse(text);
        return {
          message: j?.error?.message || "",
          code: j?.error?.code || "",
          cdscode: j?.error?.cdscode || "",
          inner: j?.error?.innererror?.message || "",
        };
      } catch {
        return { message: String(text) };
      }
    }

    function extractGuid(s) {
      const m = String(s || "").match(/[0-9a-f-]{36}/i);
      return m ? m[0].toLowerCase() : "";
    }

    function readHeader(headers, name) {
      try {
        return headers?.get?.(name) || "";
      } catch {
        return "";
      }
    }

    async function request(url, { method = "GET", body = null, prefer = "return=minimal" } = {}) {
      const token = await getToken();

      const headersBase = {
        Accept: "application/json",
        "OData-Version": "4.0",
        "OData-MaxVersion": "4.0",
        "X-Requested-With": "XMLHttpRequest",
        Prefer: prefer,
      };

      const doFetch = async (tokenHeader) => {
        const headers = {
          ...headersBase,
          ...(body ? { "Content-Type": "application/json; charset=utf-8" } : {}),
          ...(token ? { [tokenHeader]: token } : {}),
        };

        return fetch(url, {
          method,
          credentials: "same-origin",
          headers,
          body: body ? JSON.stringify(body) : null,
        });
      };

      let res = await doFetch(HDR_PRIMARY);
      if (res.status === 401 || res.status === 403) res = await doFetch(HDR_FALLBACK);

      const text = await res.text();
      if (!res.ok) {
        const parsed = parseDataverseError(text);
        const msg = parsed.inner || parsed.message || `${res.status} ${res.statusText}`;
        const e = new Error(msg);
        e.status = res.status;
        e.body = text;
        e.parsed = parsed;
        throw e;
      }

      return { res, text };
    }

    async function requestJson(url, opts = {}) {
      const { res, text } = await request(url, opts);
      if (!text) return { res, json: null };
      try {
        return { res, json: JSON.parse(text) };
      } catch {
        return { res, json: null };
      }
    }

    return {
      request,
      requestJson,
      parseDataverseError,
      extractGuid,
      readHeader,
    };
  })();

  // Optional for debugging in console
  window.PortalApi = PortalApi;

  /* ===================================================================
   * Lease / Resale Application header + Applicants grid
   * =================================================================== */

  (() => {
    const CFG = {
      applicationSetName: "vms_applicationheaders",
      dashboardPath: "/dashboard",
      newApplicantPath: "/New-Applicant/",
      sessionAppIdKey: "la_current_application_id",
      sessionAppNumKey: "la_current_application_number",

      toastId: "la-toast",
      backId: "la-back",
      saveBtnId: "la-save",
      saveCloseBtnId: "la-finish",

      appNumberElementId: "la-appnum",
      unitElementId: "la-unit",
      mutualElementId: "la-mutual",
      addressElementId: "la-address",
      appDateElementId: "la-appdate",
      statusPillId: "la-status-pill",

      functionAllocationSetName: "msdyn_functionallocations",
      functionAllocationAddressField: "msdyn_address1",

      applicantsRootId: "la-applicants",
      applicantsTableSelector: ".entity-grid table, .entitylist table, table.table",

      waitTimeoutMs: 20000,
      pollMs: 250,

      // After create, app # can be assigned asynchronously; poll briefly.
      appNumberWaitMs: 15000,
      appNumberPollMs: 500,
    };

    const log = (...a) =>
      console.log("%c[Resale-Application]", "color:#2563eb;font-weight:800", ...a);

    const stripSortNoise = (s) =>
      norm(s)
        .replace(/\.\s*sort\s+(ascending|descending)\s*$/i, "")
        .replace(/\s*sort\s+(ascending|descending)\s*$/i, "")
        .trim();

    function pageHasLeaseUi() {
      return !!byId(CFG.appNumberElementId) || !!byId(CFG.applicantsRootId);
    }

    function isPlaceholder(v) {
      const s = norm(v);
      return (
        !s ||
        s === "—" ||
        s === "-" ||
        s.toLowerCase() === "null" ||
        s.toLowerCase() === "undefined"
      );
    }

    function toast(msg, ok = true) {
      const el = byId(CFG.toastId);
      if (!el) return;
      el.textContent = msg || "";
      el.style.display = msg ? "block" : "none";
      el.classList.toggle("ok", !!ok);
      el.classList.toggle("bad", !ok);
    }

    function resolveAppGuid() {
      const url = new URL(window.location.href);
      const fromUrl = parseGuid(
        url.searchParams.get("id") || url.searchParams.get("appid")
      );
      if (fromUrl) {
        sessionStorage.setItem(CFG.sessionAppIdKey, fromUrl);
        return fromUrl;
      }
      return parseGuid(sessionStorage.getItem(CFG.sessionAppIdKey));
    }

    function canonicalizeUrl(appGuid) {
      if (!appGuid) return;
      const url = new URL(window.location.href);
      url.searchParams.set("id", appGuid);
      url.searchParams.delete("appid");
      window.history.replaceState({}, "", url.toString());
    }

    function setText(id, v) {
      const el = byId(id);
      if (!el) return;
      const value = isPlaceholder(v) ? "—" : String(v);
      if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement || el instanceof HTMLSelectElement) {
        el.value = value;
      } else {
        el.textContent = value;
      }
      el.title = value;
    }

    function readText(id) {
      const el = byId(id);
      if (!el) return "";
      if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement || el instanceof HTMLSelectElement) {
        return norm(el.value);
      }
      return norm(el.textContent);
    }

    function formatted(row, logicalName) {
      return row?.[`${logicalName}@OData.Community.Display.V1.FormattedValue`];
    }

    async function apiGetApp(appGuid, selectCsv) {
      const select = encodeURIComponent(selectCsv);
      const { json } = await PortalApi.requestJson(
        `/_api/${CFG.applicationSetName}(${appGuid})?$select=${select}`,
        { method: "GET" }
      );
      return json || null;
    }

    async function readFirstExistingField(appGuid, candidates) {
      for (const field of candidates) {
        try {
          const row = await apiGetApp(appGuid, field);
          if (!row) continue;
          const v = formatted(row, field) ?? row[field];
          if (!isPlaceholder(v)) return v;
        } catch {
          // field likely doesn't exist on this environment; ignore
        }
      }
      return "";
    }

    async function readLookupGuid(appGuid, candidates) {
      for (const field of candidates) {
        try {
          const row = await apiGetApp(appGuid, field);
          const v = row?.[field];
          const id = parseGuid(v);
          if (id) return id;
        } catch {
          // field likely doesn't exist on this environment; ignore
        }
      }
      return "";
    }

    function formatDate(value) {
      if (isPlaceholder(value)) return "";
      const date = new Date(value);
      if (Number.isNaN(date.getTime())) return String(value);
      const mm = String(date.getMonth() + 1).padStart(2, "0");
      const dd = String(date.getDate()).padStart(2, "0");
      const yyyy = date.getFullYear();
      return `${mm}/${dd}/${yyyy}`;
    }

    function setStatusPill(status) {
      const el = byId(CFG.statusPillId);
      if (!el) return;

      const normalized = norm(status).toLowerCase();
      el.classList.remove("la-badge--ok", "la-badge--warn", "la-badge--muted");
      el.classList.add("la-badge");

      let variant = "la-badge--muted";
      if (normalized.includes("active") || normalized.includes("approved") || normalized.includes("complete")) {
        variant = "la-badge--ok";
      } else if (normalized.includes("pending") || normalized.includes("review") || normalized.includes("in progress")) {
        variant = "la-badge--warn";
      }
      el.classList.add(variant);
      el.textContent = isPlaceholder(status) ? "—" : String(status);
    }

    async function getLocationAddress(locationId) {
      if (!locationId) return "";
      try {
        const select = encodeURIComponent(CFG.functionAllocationAddressField);
        const { json } = await PortalApi.requestJson(
          `/_api/${CFG.functionAllocationSetName}(${locationId})?$select=${select}`,
          { method: "GET" }
        );
        return json?.[CFG.functionAllocationAddressField] || "";
      } catch (e) {
        log("Location address lookup failed:", e.message);
        return "";
      }
    }

    async function populateApplicationDetails(appGuid) {
      try {
        // Always read safe base fields first (avoid "property not found" faults)
        const base = await apiGetApp(appGuid, "vms_applicationnumber,statuscode");
        if (!base) return;

        const appNum =
          base.vms_applicationnumber || formatted(base, "vms_applicationnumber") || "";

        // Unit/Mutual can vary by environment (lookup vs text)
        const unit = await readFirstExistingField(appGuid, [
          "_vms_location_value",
          "_vms_unit_value",
          "vms_unit",
          "vms_location",
        ]);

        const mutual = await readFirstExistingField(appGuid, [
          "_vms_mutual_value",
          "vms_mutual",
        ]);

        // Status (prefer formatted label)
        const status =
          formatted(base, "statuscode") ||
          base.statuscode ||
          (await readFirstExistingField(appGuid, ["vms_applicationstatus"])) ||
          "";

        setText(CFG.appNumberElementId, appNum);
        setText(CFG.unitElementId, unit);
        setText(CFG.mutualElementId, mutual);
        setStatusPill(status);

        const appDateRaw = await readFirstExistingField(appGuid, [
          "vms_applicationdate",
          "vms_application_date",
          "createdon",
        ]);
        setText(CFG.appDateElementId, formatDate(appDateRaw));

        const locationId = await readLookupGuid(appGuid, [
          "_vms_location_value",
          "_vms_unit_value",
        ]);
        const address = await getLocationAddress(locationId);
        setText(CFG.addressElementId, address);

        sessionStorage.setItem(CFG.sessionAppIdKey, appGuid);
        if (!isPlaceholder(appNum)) {
          sessionStorage.setItem(CFG.sessionAppNumKey, String(appNum));
        }
      } catch (e) {
        log("populateApplicationDetails failed:", e.message);
      }
    }

    function readAppNumber() {
      const fromDom = readText(CFG.appNumberElementId);
      if (!isPlaceholder(fromDom)) return fromDom;

      const fromSess = norm(sessionStorage.getItem(CFG.sessionAppNumKey));
      if (!isPlaceholder(fromSess)) return fromSess;

      return "";
    }

    async function ensureAppNumber(appGuid) {
      if (readAppNumber()) return true;

      const started = Date.now();
      while (Date.now() - started < CFG.appNumberWaitMs) {
        try {
          const row = await apiGetApp(appGuid, "vms_applicationnumber");
          const appNum =
            row?.vms_applicationnumber ||
            formatted(row, "vms_applicationnumber") ||
            "";
          if (!isPlaceholder(appNum)) {
            setText(CFG.appNumberElementId, appNum);
            sessionStorage.setItem(CFG.sessionAppNumKey, String(appNum));
            return true;
          }
        } catch {
          // transient read error; keep polling briefly
        }
        await sleep(CFG.appNumberPollMs);
      }
      return false;
    }

    // ----- Applicants table helpers -----

    function findApplicantsTable(root) {
      return qs(CFG.applicantsTableSelector, root) || null;
    }

    function applicationColumnIndex(table) {
      const ths = qsa("thead th", table);
      const headers = ths.map((th) => stripSortNoise(th.textContent).toLowerCase());

      // Prefer exact-ish matches
      const exact = headers.findIndex(
        (h) => h === "application" || h === "application #" || h === "application number"
      );
      if (exact >= 0) return exact;

      // Fallback: contains "application"
      const contains = headers.findIndex((h) => h.includes("application"));
      return contains >= 0 ? contains : null;
    }

    function filterApplicantsTable(table, appNumber) {
      const idx = applicationColumnIndex(table);
      if (idx === null) return;

      const rows = qsa("tbody tr", table);
      let kept = 0;

      for (const tr of rows) {
        const tds = qsa("td", tr);
        const v = norm(tds[idx]?.textContent || "");
        const ok = v === appNumber;
        tr.style.display = ok ? "" : "none";
        if (ok) kept++;
      }

      // If nothing matched, don't hide everything (view/column mismatch).
      if (rows.length && kept === 0) {
        for (const tr of rows) tr.style.display = "";
      }
    }

    function isNewLink(a) {
      const t = norm(a.textContent).toLowerCase();
      const aria = norm(a.getAttribute("aria-label")).toLowerCase();
      return /(new|add|create)/.test(t) || /(new|add|create)/.test(aria);
    }

    function readApplicantIdFromRow(a) {
      const tr = a.closest("tr");
      if (!tr) return "";

      const attrKeys = [
        "data-id",
        "data-entity-id",
        "data-entityid",
        "data-record-id",
        "data-recordid",
        "data-guid",
      ];
      for (const k of attrKeys) {
        const v = parseGuid(tr.getAttribute(k));
        if (v) return v;
      }

      const anyGuid = (tr.getAttribute("id") || "").match(/[0-9a-f-]{36}/i);
      return anyGuid ? parseGuid(anyGuid[0]) : "";
    }

    function rewriteApplicantLinks(root, appGuid) {
      const anchors = qsa('a[href]', root);

      for (const a of anchors) {
        const href = a.getAttribute("href") || "";
        if (!href) continue;

        // Only touch New-Applicant links
        if (!/\/new-applicant\/?/i.test(href)) continue;

        let u;
        try {
          u = new URL(href, window.location.origin);
        } catch {
          continue;
        }

        // Always carry parent app id
        u.searchParams.set("id", appGuid);
        u.searchParams.delete("appid");

        // If this is an edit link, try to keep/derive applicantid
        const existingApplicant =
          parseGuid(u.searchParams.get("applicantid")) ||
          parseGuid(u.searchParams.get("recordid")) ||
          "";

        if (!isNewLink(a)) {
          const fromRow = readApplicantIdFromRow(a);
          const applicantId = existingApplicant || fromRow;
          if (applicantId) u.searchParams.set("applicantid", applicantId);
        }

        a.setAttribute("href", u.pathname + "?" + u.searchParams.toString());
      }
    }

    function bindFooterButtons() {
      const back = byId(CFG.backId);
      if (back) back.href = CFG.dashboardPath;

      const save = byId(CFG.saveBtnId);
      if (save) {
        save.addEventListener(
          "click",
          (e) => {
            e.preventDefault();
            e.stopPropagation();
            const submit = qs('form input[type="submit"]') || qs('form button[type="submit"]');
            if (submit) submit.click();
            toast("Saved.", true);
            setTimeout(() => toast(""), 1500);
          },
          true
        );
      }

      const finish = byId(CFG.saveCloseBtnId);
      if (finish) {
        finish.addEventListener(
          "click",
          (e) => {
            e.preventDefault();
            e.stopPropagation();
            window.location.assign(CFG.dashboardPath);
          },
          true
        );
      }
    }

    async function waitForApplicantsRoot() {
      const started = Date.now();
      while (Date.now() - started < CFG.waitTimeoutMs) {
        const root = byId(CFG.applicantsRootId);
        if (root) return root;
        await sleep(CFG.pollMs);
      }
      return null;
    }

    function startApplicantsObserver(root, appGuid) {
      const apply = () => {
        const appNumber = readAppNumber();
        const table = findApplicantsTable(root);

        // Only filter once we have a real application number
        if (table && appNumber) filterApplicantsTable(table, appNumber);

        // Always rewrite links so New/Edit carry the parent app id
        rewriteApplicantLinks(root, appGuid);
      };

      apply();

      const obs = new MutationObserver(() => apply());
      obs.observe(root, { childList: true, subtree: true });

      // Re-apply after pagination/sorting clicks
      root.addEventListener("click", () => setTimeout(apply, 50), true);
    }

    async function init() {
      if (!pageHasLeaseUi()) return;

      const appGuid = resolveAppGuid();
      if (!appGuid) {
        toast("Missing application id (?id=<guid>).", false);
        return;
      }

      canonicalizeUrl(appGuid);
      bindFooterButtons();

      await populateApplicationDetails(appGuid);
      await ensureAppNumber(appGuid);

      const root = await waitForApplicantsRoot();
      if (!root) return;

      startApplicantsObserver(root, appGuid);

      log("Ready", { appGuid, appNumber: readAppNumber() });
    }

    if (document.readyState === "loading")
      document.addEventListener("DOMContentLoaded", init);
    else init();
  })();

/* ===================================================================
   * Lookup binder for vms_financialassets (annotation-based, no guessing)
   * =================================================================== */

  (() => {
    const CFG = {
      assetSet: "vms_financialassets",
      assetId: "vms_financialassetid",

      appSet: "vms_applicationheaders",
      catSet: "vms_categorylineitems",

      // Lookup COLUMN logical names on vms_financialasset:
      appLookupLogical: "vms_assetapplicationlookup",
      catLookupLogical: "vms_assetcategorylookup",

      // Cache nav prop bind keys per session
      storageKey: "vms_financialasset_bindnavs_annotations_v1",
    };

    const lower = (s) => String(s || "").toLowerCase();
    const isGuid = (s) => GUID_RX.test(String(s || "").trim());

    const lookupValueName = (lookupLogical) => `_${lookupLogical}_value`;

    function findAssociatedNavProp(obj, lookupLogical) {
      const base = lookupValueName(lookupLogical);

      const k1 = `${base}@Microsoft.Dynamics.CRM.associatednavigationproperty`;
      if (typeof obj?.[k1] === "string" && obj[k1]) return obj[k1];

      // Case-insensitive fallback
      const want = lower(k1);
      for (const [k, v] of Object.entries(obj || {})) {
        if (lower(k) === want && typeof v === "string" && v) return v;
      }
      return "";
    }

    async function getAnyAssetIdForProbe(appId) {
      // Try current application first (best)
      const filter1 = encodeURIComponent(`${lookupValueName(CFG.appLookupLogical)} eq ${appId}`);
      const url1 = `/_api/${CFG.assetSet}?$select=${CFG.assetId}&$filter=${filter1}&$top=1`;

      const r1 = await PortalApi.requestJson(url1, { method: "GET" });
      const id1 = r1?.json?.value?.[0]?.[CFG.assetId];
      if (id1 && isGuid(id1)) return String(id1).toLowerCase();

      // Otherwise: pick a record where BOTH lookups are populated (so nav annotations exist)
      const filter2 = encodeURIComponent(
        `${lookupValueName(CFG.appLookupLogical)} ne null and ${lookupValueName(CFG.catLookupLogical)} ne null`
      );
      const url2 = `/_api/${CFG.assetSet}?$select=${CFG.assetId}&$filter=${filter2}&$top=1`;

      const r2 = await PortalApi.requestJson(url2, { method: "GET" });
      const id2 = r2?.json?.value?.[0]?.[CFG.assetId];
      if (id2 && isGuid(id2)) return String(id2).toLowerCase();

      // Last resort: any record (still within permissions)
      const url3 = `/_api/${CFG.assetSet}?$select=${CFG.assetId}&$top=1`;
      const r3 = await PortalApi.requestJson(url3, { method: "GET" });
      const id3 = r3?.json?.value?.[0]?.[CFG.assetId];
      if (id3 && isGuid(id3)) return String(id3).toLowerCase();

      return "";
    }

    async function discoverBindProps(appId, catId) {
      const cached = sessionStorage.getItem(CFG.storageKey);
      if (cached) {
        try {
          const j = JSON.parse(cached);
          const ok =
            typeof j?.appBindProp === "string" &&
            typeof j?.catBindProp === "string" &&
            j.appBindProp.endsWith("@odata.bind") &&
            j.catBindProp.endsWith("@odata.bind") &&
            !j.appBindProp.startsWith("_") &&
            !j.catBindProp.startsWith("_") &&
            !j.appBindProp.includes("_value") &&
            !j.catBindProp.includes("_value");
          if (ok) return j;
          sessionStorage.removeItem(CFG.storageKey);
        } catch {
          sessionStorage.removeItem(CFG.storageKey);
        }
      }

if (!isGuid(appId)) throw new Error("Binder: invalid application GUID.");
      if (!isGuid(catId)) throw new Error("Binder: invalid category GUID.");

      const probeId = await getAnyAssetIdForProbe(appId);
      if (!probeId) {
        throw new Error(
          "Binder: cannot discover lookup nav properties because there are no readable vms_financialasset records to probe. " +
            "Create/read at least one record (or grant read permission) and try again."
        );
      }

      const select = encodeURIComponent(
        [
          CFG.assetId,
          lookupValueName(CFG.appLookupLogical),
          lookupValueName(CFG.catLookupLogical),
        ].join(",")
      );

      // Ask Dataverse to include the nav prop names for lookup columns
      const prefer =
        'odata.include-annotations="Microsoft.Dynamics.CRM.associatednavigationproperty"';

      const { json } = await PortalApi.requestJson(
        `/_api/${CFG.assetSet}(${probeId})?$select=${select}`,
        { method: "GET", prefer }
      );

      if (!json) throw new Error("Binder: probe GET returned no JSON.");

      const appNav = findAssociatedNavProp(json, CFG.appLookupLogical);
      const catNav = findAssociatedNavProp(json, CFG.catLookupLogical);

      
      // If probe record had a null lookup, annotations may be missing. Retry with a record that has both lookups.
      let appNav2 = appNav;
      let catNav2 = catNav;

      if (!appNav2 || !catNav2) {
        const filterAny = encodeURIComponent(
          `${lookupValueName(CFG.appLookupLogical)} ne null and ${lookupValueName(CFG.catLookupLogical)} ne null`
        );
        const urlAny = `/_api/${CFG.assetSet}?$select=${CFG.assetId}&$filter=${filterAny}&$top=1`;
        try {
          const rAny = await PortalApi.requestJson(urlAny, { method: "GET" });
          const anyId = rAny?.json?.value?.[0]?.[CFG.assetId];
          if (anyId && isGuid(anyId)) {
            const { json: j2 } = await PortalApi.requestJson(
              `/_api/${CFG.assetSet}(${String(anyId).toLowerCase()})?$select=${select}`,
              { method: "GET", prefer }
            );
            appNav2 = findAssociatedNavProp(j2, CFG.appLookupLogical) || appNav2;
            catNav2 = findAssociatedNavProp(j2, CFG.catLookupLogical) || catNav2;
          }
        } catch {
          // ignore and fall through to error
        }
      }

if (!appNav2 || !catNav2) {
        const keys = Object.keys(json).slice(0, 80).join(", ");
        throw new Error(
          `Binder: could not read associatednavigationproperty annotations. ` +
            `appNav="${appNav2}" catNav="${catNav2}". Returned keys (sample): ${keys}`
        );
      }

      const navs = {
        appBindProp: `${appNav2}@odata.bind`,
        catBindProp: `${catNav2}@odata.bind`,
      };

      sessionStorage.setItem(CFG.storageKey, JSON.stringify(navs));
      console.log("[Resale-FinancialAssets] ✅ bind props discovered:", navs);
      return navs;
    }

    async function bindLookups(payload, appId, catId) {
      const navs = await discoverBindProps(lower(appId), lower(catId));
      payload[navs.appBindProp] = `/${CFG.appSet}(${lower(appId)})`;
      payload[navs.catBindProp] = `/${CFG.catSet}(${lower(catId)})`;
      return payload;
    }

    window.ResaleFinancialAssetsBind = {
      bindLookups,
      discoverBindProps,
      resetCache: () => sessionStorage.removeItem(CFG.storageKey),
    };
  })();

  /* ===================================================================
   * Financial Assets grids
   * =================================================================== */

  (() => {
    const CFG = {
      appIdParam: "id",
      appIdParamFallback: "appid",

      applicationSet: "vms_applicationheaders",

      cat: {
        set: "vms_categorylineitems",
        id: "vms_categorylineitemid",
        name: "vms_categoryname",
        orderBy: "vms_categoryname",
        top: 2000,
      },

      asset: {
        set: "vms_financialassets",
        id: "vms_financialassetid",

        applicationLookup: "vms_assetapplicationlookup", // lookup COLUMN logical name
        categoryLookup: "vms_assetcategorylookup", // lookup COLUMN logical name

        name: "vms_assetname",
        institution: "vms_financialinstitution",
        description: "vms_assetdescription",
        lineTotal: "vms_linetotal",
        includeInTotal: "vms_includeintotal",

        // Additional fields used by drawers
        accountType: "vms_accounttype",
        accountLast4: "vms_accountlast4digits",
        address1: "vms_addressline1",
        city: "vms_city",
        state: "vms_state",
        annualAmount: "vms_annualamount",
        balanceOwing: "vms_balanceowing",
        quantity: "vms_quantity",

        top: 5000,
      },

      dom: {
        gridsRoot: "ra-fin-grids",
        saveBtn: "ra-save-assets",
        msg: "ra-assets-msg",
        totalYes: "ra-grand-total",
        totalAll: "ra-grand-total-all",

        drawerOverlayId: "ra-drawer-overlay",
        drawerId: "ra-drawer",
        drawerTitleId: "ra-drawer-title",
        drawerBodyId: "ra-drawer-body",
        drawerSaveId: "ra-drawer-save",
        drawerCloseId: "ra-drawer-close",
      },
    };

    const log = (...a) =>
      console.log("%c[Resale-FinancialAssets]", "color:#2563eb;font-weight:800", ...a);

    // -------- Money helpers (pennies-accurate display) --------
    const parseMoney = (v) => {
      const n = Number(String(v ?? "").replace(/[^0-9.\-]/g, ""));
      return Number.isFinite(n) ? n : 0;
    };

    const toCents = (v) => Math.round(parseMoney(v) * 100);

    const formatMoney2 = (v) => {
      if (v === null || v === undefined || v === "") return "";
      const n = Number(v);
      return Number.isFinite(n) ? n.toFixed(2) : "";
    };

    const moneyFromCents = (cents) =>
      (Number(cents || 0) / 100).toLocaleString(undefined, {
        style: "currency",
        currency: "USD",
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });

    const escapeHtml = (s) =>
      String(s ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");

    const lookupValueName = (lookupLogical) => `_${lookupLogical}_value`;

    function showMsg(text, ok) {
      const el = byId(CFG.dom.msg);
      if (!el) return;
      el.textContent = text || "";
      el.className = `ra-msg ${ok ? "ok" : "err"}`;
      el.style.display = text ? "block" : "none";
    }

    function pageHasUi() {
      return !!byId(CFG.dom.gridsRoot) && !!byId(CFG.dom.saveBtn);
    }

    function getAppId() {
      const url = new URL(window.location.href);
      return parseGuid(url.searchParams.get(CFG.appIdParam) || url.searchParams.get(CFG.appIdParamFallback)) ||
        parseGuid(sessionStorage.getItem("la_current_application_id"));
    }

    // ---------------- Categories & templates ----------------

    const ACCOUNT_TYPE_OPTIONS = [
      { value: 0, label: "Checking" },
      { value: 1, label: "Brokerage" },
    ];

    const TEMPLATE_OTHER = "other";
    const TEMPLATE = {
      life: {
        institutionLabel: "Insurance company",
        descriptionLabel: "Notes",
        lineTotalLabel: "Net cash value",
        additional: [
          { key: "annualAmount", label: "Premiums / Year", type: "money" },
          { key: "balanceOwing", label: "Policy loans", type: "money" },
          { key: "accountLast4", label: "Policy number (last 4)", type: "last4" },
        ],
        attachmentLabel: "Upload policy / statement",
      },
      bank: {
        institutionLabel: "Bank name",
        descriptionLabel: "Description",
        lineTotalLabel: "Account balance",
        additional: [
          { key: "accountType", label: "Account type", type: "accountType" },
          { key: "accountLast4", label: "Account number (last 4)", type: "last4" },
          { key: "address1", label: "Address / Branch", type: "text" },
          { key: "city", label: "City", type: "text" },
          { key: "state", label: "State", type: "text" },
        ],
        attachmentLabel: "Upload bank statement",
      },
      residential: {
        institutionLabel: "Institution",
        descriptionLabel: "Notes",
        lineTotalLabel: "Estimated equity",
        additional: [
          { key: "address1", label: "Address", type: "text" },
          { key: "city", label: "City", type: "text" },
          { key: "state", label: "State", type: "text" },
          { key: "annualAmount", label: "Annual payments", type: "money" },
          { key: "balanceOwing", label: "Balance owing", type: "money" },
        ],
        attachmentLabel: "Upload mortgage statement / appraisal",
      },
      income: {
        institutionLabel: "Institution",
        descriptionLabel: "Notes",
        lineTotalLabel: "Estimated equity",
        additional: [
          { key: "address1", label: "Address", type: "text" },
          { key: "city", label: "City", type: "text" },
          { key: "state", label: "State", type: "text" },
          { key: "annualAmount", label: "Net income / Year", type: "money" },
          { key: "balanceOwing", label: "Balance owing", type: "money" },
        ],
        attachmentLabel: "Upload rent roll / income statement",
      },
      stocks: {
        institutionLabel: "Brokerage / Institution",
        descriptionLabel: "Holdings / companies",
        lineTotalLabel: "Market value",
        additional: [{ key: "quantity", label: "Shares / quantity", type: "number" }],
        attachmentLabel: "Upload brokerage statement / holdings schedule",
      },
      cd: {
        institutionLabel: "Institution",
        descriptionLabel: "CD details",
        lineTotalLabel: "Market value",
        additional: [
          { key: "accountLast4", label: "CD number (last 4)", type: "last4" },
          { key: "annualAmount", label: "Annual interest / income", type: "money" },
        ],
        attachmentLabel: "Upload CD statement",
      },
      gov: {
        institutionLabel: "Issuing agency",
        descriptionLabel: "Bond details",
        lineTotalLabel: "Maturity value",
        additional: [
          { key: "quantity", label: "Quantity", type: "number" },
          { key: "annualAmount", label: "Annual interest / income", type: "money" },
          { key: "accountLast4", label: "Bond ID (last 4)", type: "last4" },
        ],
        attachmentLabel: "Upload bond schedule",
      },
      other: {
        institutionLabel: "Institution / Holder",
        descriptionLabel: "Description",
        lineTotalLabel: "Market value",
        additional: [
          { key: "annualAmount", label: "Annual income / year", type: "money" },
          { key: "balanceOwing", label: "Balance owing", type: "money" },
          { key: "accountLast4", label: "Identifier (last 4)", type: "last4" },
          { key: "address1", label: "Address", type: "text" },
          { key: "city", label: "City", type: "text" },
          { key: "state", label: "State", type: "text" },
        ],
        attachmentLabel: "Attach schedule / supporting docs",
      },
    };

    function templateKeyForCategoryName(categoryName) {
      const n = norm(categoryName).toLowerCase();

      if (!n) return TEMPLATE_OTHER;
      if (n === "categories") return TEMPLATE_OTHER;

      if (n.includes("life")) return "life";
      if (n.includes("bank")) return "bank";
      if (n.includes("residential")) return "residential";
      if (n.includes("income")) return "income";
      if (n.includes("stocks") || n.includes("bonds")) return "stocks";
      if (n.includes("certificate") || n.includes("deposit")) return "cd";
      if (n.includes("government")) return "gov";
      if (n.includes("other")) return "other";

      return TEMPLATE_OTHER;
    }

    // ---------------- Grid HTML ----------------

    function gridHtml(cat, idx) {
      const collapsed = idx > 0;
      const chevron = collapsed ? "▸" : "▾";
      const expanded = collapsed ? "false" : "true";
      const cardClass = collapsed ? "ra-grid-card is-collapsed" : "ra-grid-card";

      return `
        <div class="${cardClass}" data-cat="${escapeHtml(cat.id)}" data-catname="${escapeHtml(cat.name)}">
          <div class="ra-grid-hd">
            <div class="ra-grid-title">${escapeHtml(cat.name)}</div>
            <div class="ra-grid-actions" style="display:flex; gap:8px; align-items:center;">
              <button type="button" class="ra-collapse" aria-label="Toggle ${escapeHtml(cat.name)}" aria-expanded="${expanded}">${chevron}</button>
              <button type="button" class="btn btn-default ra-add-row" data-add="${escapeHtml(cat.id)}" data-catname="${escapeHtml(cat.name)}">Add row</button>
            </div>
          </div>
          <div class="ra-grid-bd">
            <table class="table ra-grid">
              <thead>
                <tr>
                  <th>Asset Name</th>
                  <th>Institution</th>
                  <th>Description</th>
                  <th style="width:140px;">Line Total</th>
                  <th style="width:120px;">Include</th>
                  <th style="width:140px;"></th>
                </tr>
              </thead>
              <tbody data-body="${escapeHtml(cat.id)}"></tbody>
            </table>
          </div>
        </div>
      `;
    }

    function encodeExtras(extras) {
      try {
        return encodeURIComponent(JSON.stringify(extras || {}));
      } catch {
        return "";
      }
    }

    function decodeExtras(raw) {
      if (!raw) return {};
      try {
        return JSON.parse(decodeURIComponent(raw));
      } catch {
        return {};
      }
    }

    function ensureExtrasShape(x) {
      const e = x && typeof x === "object" ? x : {};
      return {
        accountType: e.accountType ?? null,
        accountLast4: e.accountLast4 ?? null,
        address1: e.address1 ?? null,
        city: e.city ?? null,
        state: e.state ?? null,
        annualAmount: e.annualAmount ?? null,
        balanceOwing: e.balanceOwing ?? null,
        quantity: e.quantity ?? null,
      };
    }

    function rowHtml(r) {
      const includeVal = String(r.include ?? "Yes");
      const extras = encodeExtras(ensureExtrasShape(r.extras));

      return `
        <tr data-row="1"
            data-id="${escapeHtml(r.id || "")}"
            data-cat="${escapeHtml(r.catId || "")}"
            data-catname="${escapeHtml(r.catName || "")}"
            data-extra="${extras}"
            data-dirty="${r.dirty ? "1" : "0"}">
          <td><input class="form-control" data-k="name" value="${escapeHtml(r.name || "")}"></td>
          <td><input class="form-control" data-k="institution" value="${escapeHtml(r.institution || "")}"></td>
          <td><input class="form-control" data-k="description" value="${escapeHtml(r.description || "")}"></td>
          <td><input class="form-control" data-k="lineTotal" data-money="1" inputmode="decimal" value="${escapeHtml(formatMoney2(r.lineTotal))}"></td>
          <td>
            <select class="form-control" data-k="include">
              <option value="Yes" ${includeVal === "Yes" ? "selected" : ""}>Yes</option>
              <option value="No"  ${includeVal === "No" ? "selected" : ""}>No</option>
            </select>
          </td>
          <td>
            <div class="ra-row-actions">
              <button type="button" class="ra-link ra-details-row">Details</button>
              <button type="button" class="ra-link ra-del-row">Remove</button>
            </div>
          </td>
        </tr>
      `;
    }

    // ---------------- Totals ----------------

    function recalcTotals() {
      const root = byId(CFG.dom.gridsRoot);
      if (!root) return;

      let yesCents = 0;
      let allCents = 0;

      for (const tr of qsa('tr[data-row="1"]', root)) {
        const lineTotalValue = qs('[data-k="lineTotal"]', tr)?.value;
        const lineCents = toCents(lineTotalValue);

        allCents += lineCents;

        const inc = String(qs('[data-k="include"]', tr)?.value || "Yes");
        if (inc === "Yes") yesCents += lineCents;
      }

      const yesEl = byId(CFG.dom.totalYes);
      const allEl = byId(CFG.dom.totalAll);
      if (yesEl) yesEl.textContent = moneyFromCents(yesCents);
      if (allEl) allEl.textContent = moneyFromCents(allCents);
    }

    // ---------------- Data loading ----------------

    async function loadCategories() {
      const select = encodeURIComponent([CFG.cat.id, CFG.cat.name].join(","));
      const url =
        `/_api/${CFG.cat.set}` +
        `?$select=${select}` +
        `&$orderby=${encodeURIComponent(CFG.cat.orderBy + " asc")}` +
        `&$top=${CFG.cat.top}`;

      const { json } = await PortalApi.requestJson(url, { method: "GET" });
      return (json?.value || []).map((r) => ({
        id: r[CFG.cat.id],
        name: r[CFG.cat.name] || "(Unnamed category)",
      }));
    }

    async function loadAssets(appId) {
      const select = encodeURIComponent(
        [
          CFG.asset.id,
          lookupValueName(CFG.asset.applicationLookup),
          lookupValueName(CFG.asset.categoryLookup),
          CFG.asset.name,
          CFG.asset.institution,
          CFG.asset.description,
          CFG.asset.lineTotal,
          CFG.asset.includeInTotal,
          // Drawer fields
          CFG.asset.accountType,
          CFG.asset.accountLast4,
          CFG.asset.address1,
          CFG.asset.city,
          CFG.asset.state,
          CFG.asset.annualAmount,
          CFG.asset.balanceOwing,
          CFG.asset.quantity,
        ].join(",")
      );

      const filter = encodeURIComponent(`${lookupValueName(CFG.asset.applicationLookup)} eq ${appId}`);

      const url = `/_api/${CFG.asset.set}?$select=${select}&$filter=${filter}&$top=${CFG.asset.top}`;

      const { json } = await PortalApi.requestJson(url, { method: "GET" });
      return (json?.value || []).map((r) => ({
        id: r[CFG.asset.id],
        catId: r[lookupValueName(CFG.asset.categoryLookup)],
        name: r[CFG.asset.name] || "",
        institution: r[CFG.asset.institution] || "",
        description: r[CFG.asset.description] || "",
        lineTotal: formatMoney2(r[CFG.asset.lineTotal] ?? ""),
        include: r[CFG.asset.includeInTotal] ? "Yes" : "No",
        extras: ensureExtrasShape({
          accountType: r[CFG.asset.accountType] ?? null,
          accountLast4: r[CFG.asset.accountLast4] ?? null,
          address1: r[CFG.asset.address1] ?? null,
          city: r[CFG.asset.city] ?? null,
          state: r[CFG.asset.state] ?? null,
          annualAmount: r[CFG.asset.annualAmount] ?? null,
          balanceOwing: r[CFG.asset.balanceOwing] ?? null,
          quantity: r[CFG.asset.quantity] ?? null,
        }),
        dirty: false,
      }));
    }

    function render(categories) {
      const root = byId(CFG.dom.gridsRoot);
      if (!root) return;
      root.innerHTML = categories.map((c, i) => gridHtml(c, i)).join("");
    }

    // ---------------- Row read/write ----------------

    function readRowExtras(tr) {
      return ensureExtrasShape(decodeExtras(tr.getAttribute("data-extra")));
    }

    function writeRowExtras(tr, extras) {
      tr.setAttribute("data-extra", encodeExtras(ensureExtrasShape(extras)));
      tr.setAttribute("data-dirty", "1");
    }

    function readRows() {
      const root = byId(CFG.dom.gridsRoot);
      const out = [];

      for (const tr of qsa('tr[data-row="1"]', root)) {
        const get = (k) => qs(`[data-k="${k}"]`, tr)?.value;

        const id = parseGuid(tr.getAttribute("data-id"));
        const catId = parseGuid(tr.getAttribute("data-cat"));
        const catName = norm(tr.getAttribute("data-catname"));
        const dirty = tr.getAttribute("data-dirty") === "1";

        const row = {
          tr,
          id: id || "",
          catId: catId || "",
          catName,
          dirty,
          name: norm(get("name")),
          institution: norm(get("institution")),
          description: norm(get("description")),
          lineTotal: parseMoney(get("lineTotal")),
          include: String(get("include") || "Yes") === "Yes",
          extras: readRowExtras(tr),
        };

        const meaningful =
          !!row.name || !!row.institution || !!row.description || row.lineTotal !== 0 ||
          Object.values(row.extras).some((v) => v !== null && v !== "" && v !== 0);

        if (meaningful) out.push(row);
      }

      return out;
    }

    function buildScalarPayload(row) {
      const e = ensureExtrasShape(row.extras);

      const moneyOrNull = (v) => {
        if (v === null || v === undefined || v === "") return null;
        const n = Number(v);
        return Number.isFinite(n) ? n : null;
      };

      const numOrNull = (v) => {
        if (v === null || v === undefined || v === "") return null;
        const n = Number(String(v).replace(/[^0-9.\-]/g, ""));
        return Number.isFinite(n) ? n : null;
      };

      const intOrNull = (v) => {
        if (v === null || v === undefined || v === "") return null;
        const n = Number(v);
        return Number.isFinite(n) ? Math.trunc(n) : null;
      };

      return {
        [CFG.asset.name]: row.name || null,
        [CFG.asset.institution]: row.institution || null,
        [CFG.asset.description]: row.description || null,
        [CFG.asset.lineTotal]: row.lineTotal,
        [CFG.asset.includeInTotal]: row.include,

        // Drawer fields (nullable)
        [CFG.asset.accountType]: intOrNull(e.accountType),
        [CFG.asset.accountLast4]: e.accountLast4 ? String(e.accountLast4) : null,
        [CFG.asset.address1]: e.address1 ? String(e.address1) : null,
        [CFG.asset.city]: e.city ? String(e.city) : null,
        [CFG.asset.state]: e.state ? String(e.state) : null,
        [CFG.asset.annualAmount]: moneyOrNull(e.annualAmount),
        [CFG.asset.balanceOwing]: moneyOrNull(e.balanceOwing),
        [CFG.asset.quantity]: numOrNull(e.quantity),
      };
    }

    function extractCreatedId(res, json) {
      const fromBody = json?.[CFG.asset.id];
      if (fromBody && GUID_RX.test(String(fromBody))) return String(fromBody).toLowerCase();

      const entityId = PortalApi.readHeader(res.headers, "OData-EntityId");
      return PortalApi.extractGuid(entityId);
    }

    async function saveAll(appId) {
      try {
        showMsg("", true);

        const binder = window.ResaleFinancialAssetsBind;
        if (!binder?.bindLookups) throw new Error("Lookup binder missing on page.");

        const rows = readRows();
        const toSave = rows.filter((r) => !r.id || r.dirty);

        if (!toSave.length) {
          showMsg("No changes to save.", true);
          return;
        }

        for (const r of toSave) {
          if (!r.catId) throw new Error("Missing category id on a row.");

          if (r.id) {
            const payload = buildScalarPayload(r);
            await PortalApi.requestJson(`/_api/${CFG.asset.set}(${r.id})`, {
              method: "PATCH",
              body: payload,
            });
            r.tr.setAttribute("data-dirty", "0");
            continue;
          }

          const payload = buildScalarPayload(r);
          await binder.bindLookups(payload, appId, r.catId);

          const { res, json } = await PortalApi.requestJson(`/_api/${CFG.asset.set}`, {
            method: "POST",
            prefer: "return=representation",
            body: payload,
          });

          const newId = extractCreatedId(res, json);
          if (!newId) throw new Error("Create succeeded but could not read new record id.");

          r.tr.setAttribute("data-id", newId);
          r.tr.setAttribute("data-dirty", "0");
        }

        showMsg("Saved financial assets.", true);

        // Normalize Line Total display to 2 decimals immediately after save
        const root = byId(CFG.dom.gridsRoot);
        if (root) {
          for (const tr of qsa('tr[data-row="1"]', root)) {
            const inp = qs('[data-k="lineTotal"]', tr);
            if (inp && inp.value !== "") inp.value = formatMoney2(parseMoney(inp.value));
          }
        }

        recalcTotals();
              return true;
      }
      catch (e) {
        console.error("[Resale-FinancialAssets] save failed:", e);
        showMsg(`Save failed: ${e.message}`, false);
              return false;
      }
    }

    // ---------------- Drawer UI ----------------

    function ensureDrawerShell() {
      if (byId(CFG.dom.drawerOverlayId)) return;

      const overlay = document.createElement("div");
      overlay.id = CFG.dom.drawerOverlayId;
      overlay.className = "ra-drawer-overlay";
      overlay.innerHTML = `
        <div class="ra-drawer" id="${CFG.dom.drawerId}" role="dialog" aria-modal="true" aria-label="Asset details">
          <div class="ra-drawer-hd">
            <div class="ra-drawer-title" id="${CFG.dom.drawerTitleId}">Asset details</div>
            <button type="button" class="ra-drawer-x" id="${CFG.dom.drawerCloseId}" aria-label="Close">×</button>
          </div>
          <div class="ra-drawer-bd" id="${CFG.dom.drawerBodyId}"></div>
          <div class="ra-drawer-ft">
            <button type="button" class="btn btn-default" id="${CFG.dom.drawerSaveId}">Save & Close</button>
            <button type="button" class="btn btn-default ra-drawer-close">Close</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);

      overlay.addEventListener("click", (e) => {
        if (e.target === overlay) closeDrawer();
      });

      overlay.addEventListener("keydown", (e) => {
        if (e.key === "Escape") closeDrawer();
      });

      byId(CFG.dom.drawerCloseId)?.addEventListener("click", closeDrawer);
      overlay.querySelector(".ra-drawer-close")?.addEventListener("click", closeDrawer);
    }

    let drawerContext = null; // { tr, appId }

    function hasAdditionalValues(tpl, extras) {
      const e = ensureExtrasShape(extras);
      for (const f of tpl.additional || []) {
        const v = e[f.key];
        if (v !== null && v !== undefined && String(v).trim() !== "") return true;
      }
      return false;
    }

    function setDetailsAccordionExpanded(expanded) {
      const body = byId(CFG.dom.drawerBodyId);
      if (!body) return;

      const btn = qs('.ra-acc-btn[data-acc="details"]', body);
      const panel = qs('[data-acc-panel="details"]', body);
      if (!btn || !panel) return;

      btn.setAttribute("aria-expanded", expanded ? "true" : "false");
      panel.hidden = !expanded;
    }


    function openDrawerForRow(tr, appId) {
      ensureDrawerShell();

      const catName = norm(tr.getAttribute("data-catname"));
      const key = templateKeyForCategoryName(catName);
      const t = TEMPLATE[key] || TEMPLATE.other;

      const title = `${catName || "Asset"} details`;
      byId(CFG.dom.drawerTitleId).textContent = title;

      const body = byId(CFG.dom.drawerBodyId);
      if (!body) return;

      const gridGet = (k) => qs(`[data-k="${k}"]`, tr)?.value || "";
      const extras = readRowExtras(tr);
      const expandDetails = hasAdditionalValues(t, extras);

      body.innerHTML = buildDrawerHtml(catName, t, {
        name: gridGet("name"),
        institution: gridGet("institution"),
        description: gridGet("description"),
        lineTotal: gridGet("lineTotal"),
        include: gridGet("include") || "Yes",
        id: tr.getAttribute("data-id") || "",
        expandDetails,
        extras,
      });

      drawerContext = { tr, appId, templateKey: key };

      bindDrawerEvents(tr, t);
      const overlay = byId(CFG.dom.drawerOverlayId);
      overlay.classList.add("open");

      // focus first input
      setTimeout(() => body.querySelector("input,select,textarea,button")?.focus?.(), 20);
    }


    function formatBytes(bytes) {
      const n = Number(bytes || 0);
      if (!Number.isFinite(n) || n <= 0) return "";
      const units = ["B", "KB", "MB", "GB"];
      let v = n;
      let i = 0;
      while (v >= 1024 && i < units.length - 1) {
        v /= 1024;
        i++;
      }
      return `${v.toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
    }

    function fileToBase64(file) {
      return new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onerror = () => reject(new Error("Failed to read file."));
        r.onload = () => {
          const s = String(r.result || "");
          const comma = s.indexOf(",");
          resolve(comma >= 0 ? s.slice(comma + 1) : s);
        };
        r.readAsDataURL(file);
      });
    }

    function base64ToBlob(base64, mime) {
      const bin = atob(base64 || "");
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return new Blob([bytes], { type: mime || "application/octet-stream" });
    }

    async function listAssetAttachments(assetId) {
      const filter = `_objectid_value eq ${assetId} and isdocument eq true`;
      const q = `annotations?$select=annotationid,filename,filesize,createdon,mimetype&$filter=${encodeURIComponent(
        filter
      )}&$orderby=createdon desc`;
      const { json } = await PortalApi.requestJson(`/_api/${q}`, { method: "GET" });
      return Array.isArray(json?.value) ? json.value : [];
    }

    async function uploadAssetAttachment(assetId, file) {
      const maxMb = 8;
      if (file.size > maxMb * 1024 * 1024) {
        throw new Error(`File too large. Max ${maxMb} MB.`);
      }

      const body64 = await fileToBase64(file);
      const payload = {
        subject: file.name,
        filename: file.name,
        mimetype: file.type || "application/octet-stream",
        documentbody: body64,
        isdocument: true,
        notetext: "Uploaded via portal.",
        "objectid_vms_financialasset@odata.bind": `/${CFG.asset.set}(${assetId})`,
      };

      await PortalApi.requestJson(`/_api/annotations`, { method: "POST", body: payload });
    }

    async function deleteAssetAttachment(annotationId) {
      await PortalApi.requestJson(`/_api/annotations(${annotationId})`, { method: "DELETE" });
    }

    async function downloadAssetAttachment(annotationId) {
      const { json } = await PortalApi.requestJson(
        `/_api/annotations(${annotationId})?$select=documentbody,filename,mimetype`,
        { method: "GET" }
      );
      const filename = json?.filename || "download";
      const mime = json?.mimetype || "application/octet-stream";
      const blob = base64ToBlob(json?.documentbody || "", mime);

      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        a.remove();
      }, 250);
    }

    function renderAttachments(listEl, items) {
      if (!listEl) return;

      if (!items.length) {
        listEl.innerHTML = '<div class="ra-drawer-help">No attachments yet.</div>';
        return;
      }

      listEl.innerHTML = `
        <div class="ra-attach-items">
          ${items
            .map((it) => {
              const size = formatBytes(it.filesize);
              const when = it.createdon ? new Date(it.createdon).toLocaleString() : "";
              return `
                <div class="ra-attach-item">
                  <div class="ra-attach-meta">
                    <div class="ra-attach-name">${escapeHtml(it.filename || "Attachment")}</div>
                    <div class="ra-attach-sub">${escapeHtml([size, when].filter(Boolean).join(" • "))}</div>
                  </div>
                  <div class="ra-attach-actions">
                    <button type="button" class="btn btn-default ra-attach-action" data-attach-act="dl" data-attach-id="${escapeHtml(
                      it.annotationid
                    )}">Download</button>
                    <button type="button" class="btn btn-default ra-attach-action" data-attach-act="del" data-attach-id="${escapeHtml(
                      it.annotationid
                    )}">Remove</button>
                  </div>
                </div>
              `;
            })
            .join("")}
        </div>
      `;
    }

    async function loadAttachmentsIntoDrawer(tr) {
      const body = byId(CFG.dom.drawerBodyId);
      if (!body) return;

      const help = qs("[data-attach-help]", body);
      const listEl = qs("[data-attach-list]", body);
      const fileEl = qs("[data-attach-file]", body);
      const uploadBtn = qs(".ra-attach-upload", body);

      const assetId = parseGuid(tr?.getAttribute("data-id") || "");
      if (!assetId) {
        if (help) help.textContent = "Save this asset to enable uploads.";
        if (listEl) listEl.innerHTML = "";
        if (fileEl) fileEl.disabled = true;
        if (uploadBtn) uploadBtn.disabled = true;
        return;
      }

      if (help) help.textContent = "";
      if (fileEl) fileEl.disabled = false;
      if (uploadBtn) uploadBtn.disabled = false;

      try {
        const items = await listAssetAttachments(assetId);
        renderAttachments(listEl, items);
        if (items.length) setDetailsAccordionExpanded(true);
      } catch (e) {
        if (help) help.textContent = `Uploads not available: ${e.message}`;
        if (listEl) listEl.innerHTML = "";
      }
    }

    function closeDrawer() {
      const overlay = byId(CFG.dom.drawerOverlayId);
      if (!overlay) return;
      overlay.classList.remove("open");
      drawerContext = null;
    }

    function buildDrawerHtml(categoryName, tpl, model) {
      const instLabel = tpl.institutionLabel || "Institution";
      const descLabel = tpl.descriptionLabel || "Description";
      const ltLabel = tpl.lineTotalLabel || "Line Total";
      const attachLabel = tpl.attachmentLabel || "Attach schedule";

      const addFieldsHtml = tpl.additional
        .map((f) => drawerFieldHtml(f, model.extras))
        .join("");

      return `
        <div class="ra-drawer-form">
          <div class="ra-drawer-row">
            <div class="ra-drawer-field">
              <label class="ra-drawer-label">Asset Name</label>
              <input class="form-control" data-dk="name" value="${escapeHtml(model.name)}">
            </div>
            <div class="ra-drawer-field">
              <label class="ra-drawer-label">${escapeHtml(instLabel)}</label>
              <input class="form-control" data-dk="institution" value="${escapeHtml(model.institution)}">
            </div>
          </div>

          <div class="ra-drawer-row">
            <div class="ra-drawer-field ra-drawer-field-wide">
              <label class="ra-drawer-label">${escapeHtml(descLabel)}</label>
              <input class="form-control" data-dk="description" value="${escapeHtml(model.description)}">
            </div>
          </div>

          <div class="ra-drawer-row">
            <div class="ra-drawer-field">
              <label class="ra-drawer-label">${escapeHtml(ltLabel)}</label>
              <input class="form-control" data-dk="lineTotal" inputmode="decimal" value="${escapeHtml(formatMoney2(model.lineTotal))}">
            </div>
            <div class="ra-drawer-field">
              <label class="ra-drawer-label">Include</label>
              <select class="form-control" data-dk="include">
                <option value="Yes" ${String(model.include) === "Yes" ? "selected" : ""}>Yes</option>
                <option value="No"  ${String(model.include) === "No" ? "selected" : ""}>No</option>
              </select>
            </div>
          </div>

          <button type="button" class="ra-acc-btn" data-acc="details" aria-expanded="${model.expandDetails ? "true" : "false"}">
            Additional details
          </button>
          <div class="ra-acc-panel" data-acc-panel="details" ${model.expandDetails ? "" : "hidden"}>
            <div class="ra-acc-inner">
              ${addFieldsHtml || `<div class="ra-drawer-help">No additional fields for this category.</div>`}
              <div class="ra-drawer-attach">
                <div class="ra-drawer-attach-title">${escapeHtml(attachLabel)}</div>
                <div class="ra-drawer-help" data-attach-help></div>
                <div class="ra-drawer-attach-row">
                  <input type="file" class="form-control ra-attach-file" data-attach-file>
                  <button type="button" class="btn btn-default ra-attach-upload">Upload</button>
                </div>
                <div class="ra-attach-list" data-attach-list></div>
              </div>
            </div>
          </div>
        </div>
      `;
    }

    function drawerFieldHtml(field, extras) {
      const e = ensureExtrasShape(extras);
      const v = e[field.key];

      if (field.type === "accountType") {
        const opts = ACCOUNT_TYPE_OPTIONS.map((o) => {
          const sel = String(o.value) === String(v) ? "selected" : "";
          return `<option value="${o.value}" ${sel}>${escapeHtml(o.label)}</option>`;
        }).join("");
        return `
          <div class="ra-drawer-field ra-drawer-field-wide">
            <label class="ra-drawer-label">${escapeHtml(field.label)}</label>
            <select class="form-control" data-ek="${escapeHtml(field.key)}">
              <option value="" ${v === null || v === "" ? "selected" : ""}>—</option>
              ${opts}
            </select>
          </div>
        `;
      }

      if (field.type === "money") {
        return `
          <div class="ra-drawer-field">
            <label class="ra-drawer-label">${escapeHtml(field.label)}</label>
            <input class="form-control" data-money="1" inputmode="decimal" data-ek="${escapeHtml(field.key)}" value="${escapeHtml(formatMoney2(v))}">
          </div>
        `;
      }

      if (field.type === "last4") {
        return `
          <div class="ra-drawer-field">
            <label class="ra-drawer-label">${escapeHtml(field.label)}</label>
            <input class="form-control" inputmode="numeric" maxlength="4" data-ek="${escapeHtml(field.key)}" value="${escapeHtml(v ?? "")}">
          </div>
        `;
      }

      if (field.type === "number") {
        return `
          <div class="ra-drawer-field">
            <label class="ra-drawer-label">${escapeHtml(field.label)}</label>
            <input class="form-control" inputmode="decimal" data-ek="${escapeHtml(field.key)}" value="${escapeHtml(v ?? "")}">
          </div>
        `;
      }

      return `
        <div class="ra-drawer-field ra-drawer-field-wide">
          <label class="ra-drawer-label">${escapeHtml(field.label)}</label>
          <input class="form-control" data-ek="${escapeHtml(field.key)}" value="${escapeHtml(v ?? "")}">
        </div>
      `;
    }

    function bindDrawerEvents(tr, tpl) {
      const body = byId(CFG.dom.drawerBodyId);
      if (!body) return;

      // Toggle accordion (collapsed by default)
      body.querySelectorAll(".ra-acc-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
          const expanded = btn.getAttribute("aria-expanded") === "true";
          btn.setAttribute("aria-expanded", expanded ? "false" : "true");
          const panel = body.querySelector(`[data-acc-panel="${btn.getAttribute("data-acc")}"]`);
          if (panel) panel.hidden = expanded;
        });
      });

      const markDirty = () => tr.setAttribute("data-dirty", "1");

      const syncGrid = (k, v) => {
        const inp = qs(`[data-k="${k}"]`, tr);
        if (inp) inp.value = v;
        markDirty();
      };

      // Top-level sync to grid
      const nameEl = body.querySelector('[data-dk="name"]');
      const instEl = body.querySelector('[data-dk="institution"]');
      const descEl = body.querySelector('[data-dk="description"]');
      const ltEl = body.querySelector('[data-dk="lineTotal"]');
      const incEl = body.querySelector('[data-dk="include"]');

      nameEl?.addEventListener("input", () => syncGrid("name", nameEl.value));
      instEl?.addEventListener("input", () => syncGrid("institution", instEl.value));
      descEl?.addEventListener("input", () => syncGrid("description", descEl.value));

      incEl?.addEventListener("change", () => {
        syncGrid("include", incEl.value);
        recalcTotals();
      });

      ltEl?.addEventListener("input", () => {
        syncGrid("lineTotal", ltEl.value);
        recalcTotals();
      });

      ltEl?.addEventListener("blur", () => {
        ltEl.value = formatMoney2(parseMoney(ltEl.value));
        syncGrid("lineTotal", ltEl.value);
        recalcTotals();
      });

      // Additional details -> extras
      const updateExtra = (key, value) => {
        const ex = readRowExtras(tr);
        ex[key] = value;
        writeRowExtras(tr, ex);
      };

      body.querySelectorAll("[data-ek]").forEach((el) => {
        const key = el.getAttribute("data-ek");
        const isMoney = el.getAttribute("data-money") === "1";

        const onUpdate = () => {
          let v = el.value;
          if (key === "accountType") v = v === "" ? null : Number(v);
          if (key === "annualAmount" || key === "balanceOwing") v = v === "" ? null : parseMoney(v);
          if (key === "quantity") v = v === "" ? null : Number(String(v).replace(/[^0-9.\-]/g, ""));
          if (key === "accountLast4") v = v ? String(v).replace(/\D+/g, "").slice(-4) : null;
          if (key === "state") v = v ? String(v).trim().toUpperCase().slice(0, 2) : null;
          if (key === "city" || key === "address1") v = v ? String(v) : null;

          updateExtra(key, v);
        };

        el.addEventListener("input", () => {
          if (key === "accountLast4") el.value = String(el.value || "").replace(/\D+/g, "").slice(0, 4);
          onUpdate();
        });

        el.addEventListener("change", onUpdate);

        if (isMoney) {
          el.addEventListener("blur", () => {
            el.value = formatMoney2(parseMoney(el.value));
            onUpdate();
          });
        }
      });

      
      // Attachments (Notes) - load list + bind upload/delete/download
      loadAttachmentsIntoDrawer(tr);

      const attachBody = byId(CFG.dom.drawerBodyId);
      const fileEl = attachBody ? qs("[data-attach-file]", attachBody) : null;
      const uploadBtn = attachBody ? qs(".ra-attach-upload", attachBody) : null;
      const listEl = attachBody ? qs("[data-attach-list]", attachBody) : null;

      uploadBtn?.addEventListener("click", async () => {
        try {
          const file = fileEl?.files?.[0];
          if (!file) return;

          const appId = drawerContext?.appId;
          if (!appId) return;

          // Ensure this row has an id for attachment binding
          if (!parseGuid(tr.getAttribute("data-id") || "")) {
            const ok = await saveAll(appId);
            if (!ok) return;
          }

          const assetId = parseGuid(tr.getAttribute("data-id") || "");
          if (!assetId) throw new Error("Please save this asset first.");

          await uploadAssetAttachment(assetId, file);
          if (fileEl) fileEl.value = "";
          showMsg("Uploaded attachment.", true);
          await loadAttachmentsIntoDrawer(tr);
        } catch (e) {
          console.error("[Resale-FinancialAssets] upload failed:", e);
          showMsg(`Upload failed: ${e.message}`, false);
        }
      });

      listEl?.addEventListener("click", async (e) => {
        const btn = e.target?.closest?.("[data-attach-act]");
        if (!btn) return;

        const act = btn.getAttribute("data-attach-act");
        const id = parseGuid(btn.getAttribute("data-attach-id") || "");
        if (!id) return;

        try {
          if (act === "dl") {
            await downloadAssetAttachment(id);
            return;
          }

          if (act === "del") {
            if (!confirm("Remove this attachment?")) return;
            await deleteAssetAttachment(id);
            showMsg("Removed attachment.", true);
            await loadAttachmentsIntoDrawer(tr);
          }
        } catch (err) {
          console.error("[Resale-FinancialAssets] attachment action failed:", err);
          showMsg(`Attachment failed: ${err.message}`, false);
        }
      });

// Drawer save button
      const saveBtn = byId(CFG.dom.drawerSaveId);
      saveBtn?.addEventListener("click", async () => {
        if (!drawerContext?.appId) return;
        const ok = await saveAll(drawerContext.appId);
        if (ok) {
          const tr0 = drawerContext?.tr;
          closeDrawer();
          if (tr0) setTimeout(() => tr0.scrollIntoView({ block: "center" }), 50);
        }
      });
    }

    // ---------------- UI bindings ----------------

    function bindUi(appId) {
      const root = byId(CFG.dom.gridsRoot);
      if (!root) return;

      root.addEventListener(
        "input",
        (e) => {
          const tr = e.target?.closest?.('tr[data-row="1"]');
          if (tr) tr.setAttribute("data-dirty", "1");

          if (e.target?.matches?.('[data-k="lineTotal"]') || e.target?.matches?.('[data-k="include"]')) {
            recalcTotals();
          }
        },
        true
      );

      // Normalize Line Total to 2 decimals when leaving the field
      root.addEventListener(
        "blur",
        (e) => {
          if (!e.target?.matches?.('[data-k="lineTotal"]')) return;
          e.target.value = formatMoney2(parseMoney(e.target.value));
          recalcTotals();
        },
        true
      );

      root.addEventListener("click", async (e) => {
        const collapseBtn = e.target.closest?.(".ra-collapse");
        if (collapseBtn) {
          const card = collapseBtn.closest?.(".ra-grid-card");
          if (card) {
            card.classList.toggle("is-collapsed");
            const collapsed = card.classList.contains("is-collapsed");
            collapseBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
            collapseBtn.textContent = collapsed ? "▸" : "▾";
          }
          return;
        }

        const addBtn = e.target.closest?.(".ra-add-row");
        if (addBtn) {
          const card = addBtn.closest?.(".ra-grid-card");
          if (card && card.classList.contains("is-collapsed")) {
            card.classList.remove("is-collapsed");
            const btn = card.querySelector?.(".ra-collapse");
            if (btn) {
              btn.setAttribute("aria-expanded", "true");
              btn.textContent = "▾";
            }
          }

          const catIdRaw = addBtn.getAttribute("data-add");
          const catName = addBtn.getAttribute("data-catname") || "";
          const catId = parseGuid(catIdRaw);

          const tbody = root.querySelector(`tbody[data-body="${CSS.escape(catIdRaw)}"]`);
          if (tbody && catId) {
            tbody.insertAdjacentHTML(
              "beforeend",
              rowHtml({
                id: "",
                catId,
                catName,
                name: "",
                institution: "",
                description: "",
                lineTotal: "",
                include: "Yes",
                extras: {},
                dirty: true,
              })
            );
            recalcTotals();
          }
          return;
        }

        const detailsBtn = e.target.closest?.(".ra-details-row");
        if (detailsBtn) {
          const tr = detailsBtn.closest('tr[data-row="1"]');
          if (tr) openDrawerForRow(tr, appId);
          return;
        }

        const delBtn = e.target.closest?.(".ra-del-row");
        if (delBtn) {
          const tr = delBtn.closest('tr[data-row="1"]');
          const id = parseGuid(tr?.getAttribute("data-id"));

          try {
            if (id) await PortalApi.requestJson(`/_api/${CFG.asset.set}(${id})`, { method: "DELETE" });
          } catch (err) {
            console.error("[Resale-FinancialAssets] delete failed:", err);
            showMsg(`Delete failed: ${err.message}`, false);
            return;
          }

          tr?.remove();
          recalcTotals();
        }
      });

      const saveBtn = byId(CFG.dom.saveBtn);
      if (saveBtn) saveBtn.addEventListener("click", () => saveAll(appId));
    }

    // ---------------- Init ----------------

    async function init() {
      if (!pageHasUi()) return;

      const appId = getAppId();
      if (!appId) {
        showMsg("Missing application context. Open as: .../Resale-Application/?id=<guid>", false);
        return;
      }

      try {
        const categories = await loadCategories();
        const catNameById = new Map(categories.map((c) => [String(c.id).toLowerCase(), c.name]));

        render(categories);

        const assets = await loadAssets(appId);

        for (const cat of categories) {
          const tbody = qs(`tbody[data-body="${cat.id}"]`);
          if (!tbody) continue;

          const catRows = assets
            .filter((a) => String(a.catId || "").toLowerCase() === String(cat.id).toLowerCase())
            .map((a) => ({ ...a, catName: cat.name }));

          if (!catRows.length) {
            tbody.insertAdjacentHTML(
              "beforeend",
              rowHtml({
                id: "",
                catId: cat.id,
                catName: cat.name,
                name: "",
                institution: "",
                description: "",
                lineTotal: "",
                include: "Yes",
                extras: {},
                dirty: true,
              })
            );
            continue;
          }

          for (const r of catRows) tbody.insertAdjacentHTML("beforeend", rowHtml(r));
        }

        bindUi(appId);
        recalcTotals();

        log("Ready", { appId, assets: assets.length, categories: categories.length });
      } catch (e) {
        console.error("[Resale-FinancialAssets] init failed:", e);
        showMsg(`Init failed: ${e.message}`, false);
      }
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
  })();
})();

/* ===================================================================
   Resale Application - Escrow Details (compact rows + drawer editing)
   =================================================================== */
(() => {
  "use strict";

  const CFG = {
    rootId: "la-header-form",
    appSet: "vms_applicationheaders",
    appIdParam: "id",
    roles: [
      { key: "escrow", title: "Escrow", companyNav: "vms_escrowcompany", agentNav: "vms_escrowagent", companyTargetSet: "accounts", agentTargetSet: "contacts" },
      { key: "seller", title: "Seller", companyNav: "vms_sellerscompany", agentNav: "vms_sellersagent", companyTargetSet: "accounts", agentTargetSet: "contacts" },
      { key: "buyer",  title: "Buyer",  companyNav: "vms_buyerscompany",  agentNav: "vms_buyersagent",  companyTargetSet: "accounts", agentTargetSet: "contacts" },
    ],
    txFields: [
      { logical: "vms_escrownumber", label: "Escrow Number", type: "text" },
      { logical: "vms_purchaseamount", label: "Purchase Amount", type: "money" },
      { logical: "vms_financingtype", label: "Financing Type", type: "optionset" }, // option set displays as formatted value
      { logical: "vms_financingamount", label: "Financing Amount", type: "money" },
      { logical: "vms_escrowloanamount", label: "Escrow Loan Amount", type: "money" },
      { logical: "vms_estimatedclosedate", label: "Estimated Close Date", type: "date" },
      { logical: "vms_resaleinspectiondate", label: "Resale Inspection Date", type: "date" },
    ],
    maxSuggestions: 12,
    debounceMs: 180,
  };

  const byId = (id) => document.getElementById(id);
  const qs = (sel, root = document) => root.querySelector(sel);
  const escapeHtml = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const isGuid = (s) => /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(String(s || ""));
  const parseGuid = (s) => (isGuid(s) ? String(s).toLowerCase() : "");
  const debounce = (fn, ms) => {
    let t = 0;
    return (...args) => {
      window.clearTimeout(t);
      t = window.setTimeout(() => fn(...args), ms);
    };
  };

  function getAppId() {
    const url = new URL(window.location.href);
    return parseGuid(url.searchParams.get(CFG.appIdParam));
  }

  function createdFlow() {
    const url = new URL(window.location.href);
    return url.searchParams.get("created") === "1";
  }

  function showToast(root, text, ok = false) {
    const el = qs("[data-es-toast]", root);
    if (!el) return;
    if (!text) {
      el.style.display = "none";
      el.textContent = "";
      el.classList.remove("alert-danger", "alert-success");
      return;
    }
    el.style.display = "block";
    el.classList.toggle("alert-success", ok);
    el.classList.toggle("alert-danger", !ok);
    el.textContent = text;
  }

  function buildHeaderUi(root) {
    const hd = qs(".la-card__hd", root);
    const bd = qs(".la-card__bd", root);
    if (hd) {
      hd.innerHTML = `
        <div class="la-hd-left">
          <div class="la-hd-title">Escrow Details</div>
          <span class="la-badge la-badge--muted" data-es-badge>Status</span>
        </div>
        <div class="la-hd-meta" data-es-meta></div>
      `;
    }
    if (bd) {
      bd.innerHTML = `
        <div class="alert" data-es-toast style="display:none; margin-bottom: 10px;"></div>
        <div class="es-grid-wrap">
          <table class="table es-grid">
            <thead>
              <tr>
                <th style="width:120px;">Party</th>
                <th>Company</th>
                <th>Agent</th>
                <th style="width:220px;">Email</th>
                <th style="width:140px;">Office</th>
                <th style="width:140px;">Mobile</th>
                <th style="width:120px;"></th>
              </tr>
            </thead>
            <tbody data-es-body></tbody>
          </table>
        </div>
        <div class="es-hint">Agents are filtered by the selected company. Clear the company to clear the agent.</div>
      `;
    }
  }

  function ensureDrawer() {
    let overlay = byId("es-drawer-overlay");
    if (overlay) return overlay;

    overlay = document.createElement("div");
    overlay.id = "es-drawer-overlay";
    overlay.className = "es-drawer-overlay";
    overlay.innerHTML = `
      <div class="es-drawer" id="es-drawer" role="dialog" aria-modal="true" aria-label="Escrow Details">
        <div class="es-drawer__hd">
          <div class="es-drawer__title" data-es-drawer-title>Details</div>
          <button type="button" class="btn btn-default btn-sm" data-es-close>Close</button>
        </div>
        <div class="es-drawer__bd" data-es-drawer-body></div>
        <div class="es-drawer__ft">
          <div class="es-uploading" data-es-saving style="display:none;"><span class="es-spinner"></span><span>Saving…</span></div>
          <div style="display:flex; gap:10px; margin-left:auto;">
            <button type="button" class="btn btn-primary" data-es-saveclose>Save &amp; Close</button>
            <button type="button" class="btn btn-default" data-es-cancel>Cancel</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) closeDrawer();
    });
    qs("[data-es-close]", overlay)?.addEventListener("click", closeDrawer);
    qs("[data-es-cancel]", overlay)?.addEventListener("click", closeDrawer);

    return overlay;
  }

  function openDrawer(title, bodyHtml) {
    const overlay = ensureDrawer();
    qs("[data-es-drawer-title]", overlay).textContent = title;
    qs("[data-es-drawer-body]", overlay).innerHTML = bodyHtml;
    overlay.classList.add("open");
  }

  function closeDrawer() {
    const overlay = byId("es-drawer-overlay");
    if (!overlay) return;
    overlay.classList.remove("open");
    qs("[data-es-drawer-body]", overlay).innerHTML = "";
    showToast(byId(CFG.rootId), "");
  }

  function odataQuote(s) {
    return String(s ?? "").replace(/'/g, "''");
  }

  async function readHeaderRecord(appId) {
  const select = [
    ...CFG.roles.flatMap((r) => [`_${r.companyNav}_value`, `_${r.agentNav}_value`]),
    ...CFG.txFields.map((f) => f.logical),
  ].join(",");
  const url = `/_api/${CFG.appSet}(${appId})?$select=${select}`;
  const { json } = await window.PortalApi.requestJson(url, {
    method: "GET",
    prefer: 'odata.include-annotations="*",return=representation',
  });
  return json || {};
}

  async function readContacts(ids) {
  const uniq = [...new Set(ids.filter(Boolean))];
  const out = new Map();

  for (const id of uniq) {
    try {
      const { json: c } = await window.PortalApi.requestJson(
        `/_api/contacts(${id})?$select=contactid,fullname,emailaddress1,telephone1,mobilephone`,
        { method: "GET", prefer: 'odata.include-annotations="*",return=representation' }
      );
      if (!c) continue;

      out.set(id, {
        id: parseGuid(c.contactid),
        name: c.fullname || c?.[`fullname@OData.Community.Display.V1.FormattedValue`] || "",
        email: c.emailaddress1 || "",
        office: c.telephone1 || "",
        mobile: c.mobilephone || "",
      });
    } catch (e) {
      // Non-fatal: contact might be invisible to this user.
      console.warn("[Escrow] contact read skipped:", id, e?.message || e);
    }
  }

  return out;
}

  function fmtMoney(v) {
    if (v === null || v === undefined || v === "") return "—";
    const n = Number(v);
    if (Number.isNaN(n)) return String(v);
    return n.toLocaleString(undefined, { style: "currency", currency: "USD" });
  }

  function fmtDate(v) {
    if (!v) return "—";
    try {
      const d = new Date(v);
      if (Number.isNaN(d.getTime())) return String(v);
      return d.toLocaleDateString();
    } catch {
      return String(v);
    }
  }

  function safeFmt(v) {
    return v ? String(v) : "—";
  }

  function extractLookup(rec, nav) {
    const id = parseGuid(rec?.[`_${nav}_value`]);
    const name = rec?.[`_${nav}_value@OData.Community.Display.V1.FormattedValue`] || "";
    return { id, name };
  }

  function computeBadge(model) {
    const any = CFG.roles.some((r) => model[r.key].company.id || model[r.key].agent.id);
    if (!any) return { text: "Not started", cls: "la-badge--muted" };

    const invalid = CFG.roles.some((r) => model[r.key].agent.id && !model[r.key].company.id);
    if (invalid) return { text: "Needs review", cls: "la-badge--warn" };

    return { text: "In progress", cls: "la-badge--ok" };
  }

  function buildGridRows(model) {
    const rows = [];

    for (const r of CFG.roles) {
      const party = model[r.key];
      const agent = party.agent.id ? model.contacts.get(party.agent.id) : null;

      rows.push(`
        <tr data-es-role="${escapeHtml(r.key)}">
          <td class="es-role">${escapeHtml(r.title)}</td>
          <td>${escapeHtml(party.company.name || "—")}</td>
          <td>${escapeHtml(party.agent.name || "—")}</td>
          <td>${escapeHtml(agent?.email || "—")}</td>
          <td>${escapeHtml(agent?.office || "—")}</td>
          <td>${escapeHtml(agent?.mobile || "—")}</td>
          <td class="es-actions">
            <button type="button" class="btn btn-default btn-sm es-inline-btn" data-es-open="${escapeHtml(r.key)}">Details</button>
          </td>
        </tr>
      `);
    }

    // Transaction summary row
    const txPairs = CFG.txFields.map((f) => {
      const raw = model.tx[f.logical];
      const formatted =
        f.type === "money" ? fmtMoney(raw)
        : f.type === "date" ? fmtDate(raw)
        : (model.txFmt[f.logical] || safeFmt(raw));
      return `<span class="es-kpi"><b>${escapeHtml(f.label)}:</b> ${escapeHtml(formatted)}</span>`;
    }).join("");

    rows.push(`
      <tr data-es-role="tx">
        <td class="es-role">Transaction</td>
        <td colspan="5"><div class="es-kpis">${txPairs}</div></td>
        <td class="es-actions">
          <button type="button" class="btn btn-default btn-sm es-inline-btn" data-es-open="tx">Details</button>
        </td>
      </tr>
    `);

    return rows.join("");
  }

  async function hydrate(root, appId) {
    const rec = await readHeaderRecord(appId);

    const model = {
      escrow: { company: extractLookup(rec, "vms_escrowcompany"), agent: extractLookup(rec, "vms_escrowagent") },
      seller: { company: extractLookup(rec, "vms_sellerscompany"), agent: extractLookup(rec, "vms_sellersagent") },
      buyer:  { company: extractLookup(rec, "vms_buyerscompany"),  agent: extractLookup(rec, "vms_buyersagent") },
      tx: {},
      txFmt: {},
      contacts: new Map(),
    };

    for (const f of CFG.txFields) {
      model.tx[f.logical] = rec?.[f.logical] ?? null;
      model.txFmt[f.logical] = rec?.[`${f.logical}@OData.Community.Display.V1.FormattedValue`] || "";
    }

    const agentIds = CFG.roles.map((r) => model[r.key].agent.id).filter(Boolean);
    model.contacts = await readContacts(agentIds);

    // header badge/meta
    const badge = computeBadge(model);
    const badgeEl = qs("[data-es-badge]", root);
    badgeEl.textContent = badge.text;
    badgeEl.classList.remove("la-badge--ok", "la-badge--warn", "la-badge--muted");
    badgeEl.classList.add(badge.cls);

    const metaEl = qs("[data-es-meta]", root);
    const filled = CFG.roles.filter((r) => model[r.key].company.id || model[r.key].agent.id).length;
    metaEl.textContent = filled ? `${filled} party${filled === 1 ? "" : "ies"} set` : "No parties set";

    // grid
    qs("[data-es-body]", root).innerHTML = buildGridRows(model);

    return model;
  }

  const ES_NAV_CACHE_PREFIX = "es.navprop.";

  // Case-sensitive navigation property overrides for vms_applicationheaders lookups (Power Pages)
  const HEADER_LOOKUP_NAV_OVERRIDES = {
    vms_escrowcompany: "vms_EscrowCompany",
    vms_escrowagent: "vms_EscrowAgent",
    vms_buyerscompany: "vms_BuyersCompany",
    vms_buyersagent: "vms_BuyersAgent",
    vms_sellerscompany: "vms_SellersCompany",
    vms_sellersagent: "vms_SellersAgent",
  };

async function resolveAssociatedNavProp(appId, lookupLogicalName) {
  const forced = HEADER_LOOKUP_NAV_OVERRIDES[lookupLogicalName];
  if (forced) return forced;

  const cacheKey = `${ES_NAV_CACHE_PREFIX}${lookupLogicalName}`;
  const cached = sessionStorage.getItem(cacheKey);
  if (cached) return cached;

  const sel = `_${lookupLogicalName}_value`;
  try {
    const { json } = await window.PortalApi.requestJson(
      `/_api/${CFG.appSet}(${appId})?$select=${sel}`,
      { method: "GET", prefer: 'odata.include-annotations="*"' }
    );

    const annKey = `${sel}@Microsoft.Dynamics.CRM.associatednavigationproperty`;
    const nav = json?.[annKey];
    const use = typeof nav === "string" && nav.trim() ? nav.trim() : lookupLogicalName;
    sessionStorage.setItem(cacheKey, use);
    return use;
  } catch {
    // If the column isn't allowed in Web API fields, we fallback and let the caller surface the real error.
    return lookupLogicalName;
  }
}

async function setLookupRef(appId, lookupLogicalName, targetSet, targetId) {
  const navProp = await resolveAssociatedNavProp(appId, lookupLogicalName);

  // PATCH + @odata.bind is the most compatible on Power Pages
  try {
    const bindKey = `${navProp}@odata.bind`;
    const bindVal = targetId ? `/${targetSet}(${targetId})` : null;

    await window.PortalApi.requestJson(`/_api/${CFG.appSet}(${appId})`, {
      method: "PATCH",
      body: { [bindKey]: bindVal },
      prefer: "return=minimal",
    });
    return;
  } catch (e) {
    // fall through to $ref
  }

  const refUrl = `/_api/${CFG.appSet}(${appId})/${navProp}/$ref`;

  if (targetId) {
    // Relative @odata.id is safest in Power Pages
    const odataId = `/_api/${targetSet}(${targetId})`;
    await window.PortalApi.requestJson(refUrl, {
      method: "PUT",
      body: { "@odata.id": odataId },
      prefer: "return=minimal",
    });
    return;
  }

  await window.PortalApi.requestJson(refUrl, { method: "DELETE", prefer: "return=minimal" });
}

async function patchHeader(appId, payload) {
    const url = `/_api/${CFG.appSet}(${appId})`;
    await window.PortalApi.requestJson(url, { method: "PATCH", body: payload, prefer: "return=minimal" });
  }

  function roleDrawerHtml(roleKey, model) {
    const role = CFG.roles.find((r) => r.key === roleKey);
    const party = model[roleKey];
    const agent = party.agent.id ? model.contacts.get(party.agent.id) : null;

    return `
      <div data-es-roleform data-role="${escapeHtml(roleKey)}" style="position:relative;">
        <div class="es-field" style="position:relative;">
          <label class="es-label">Company</label>
          <input class="es-input" type="text" autocomplete="off" data-es-company-input value="${escapeHtml(party.company.name || "")}" placeholder="Search company...">
          <input type="hidden" data-es-company-id value="${escapeHtml(party.company.id || "")}">
          <div class="es-menu" data-es-company-menu style="display:none;"></div>
        </div>

        <div class="es-field" style="position:relative;">
          <label class="es-label">Agent</label>
          <input class="es-input" type="text" autocomplete="off" data-es-agent-input value="${escapeHtml(party.agent.name || "")}" placeholder="${party.company.id ? "Search agent..." : "Select a company first"}" ${party.company.id ? "" : "disabled"}>
          <input type="hidden" data-es-agent-id value="${escapeHtml(party.agent.id || "")}">
          <div class="es-menu" data-es-agent-menu style="display:none;"></div>
          <div class="es-hint">Press Enter to show filtered agents.</div>
        </div>

        <div class="es-field">
          <label class="es-label">Agent details</label>
          <div class="es-hint">Agent contact info is read-only here.</div>
        </div>
      </div>
    `;
  }

  function txDrawerHtml(model) {
    const fields = CFG.txFields.map((f) => {
      const raw = model.tx[f.logical] ?? "";
      const val = f.type === "date" ? (raw ? String(raw).slice(0, 10) : "") : String(raw ?? "");
      const typeAttr = f.type === "money" ? 'inputmode="decimal"' : "";
      const inputType = f.type === "date" ? "date" : "text";
      const isReadOnly = f.type === "optionset";

      return `
        <div class="es-field">
          <label class="es-label">${escapeHtml(f.label)}</label>
          <input class="es-input" type="${inputType}" ${typeAttr} data-es-tx="${escapeHtml(f.logical)}" value="${escapeHtml(isReadOnly ? (model.txFmt[f.logical] || "") : val)}" placeholder="—" ${isReadOnly ? "disabled" : ""}>
        </div>
      `;
    }).join("");

    return `<div data-es-txform>${fields}</div>`;
  }

  async function searchAccounts(q) {
  const term = odataQuote(q);
  const url = `/_api/accounts?$select=accountid,name&$filter=statecode eq 0 and contains(name,'${term}')&$orderby=name asc&$top=${CFG.maxSuggestions}`;
  const { json } = await window.PortalApi.requestJson(url, { method: "GET" });
  return (json?.value || []).map((r) => ({ id: parseGuid(r.accountid), name: r.name || "" }));
}

  async function searchContacts(companyId, q) {
  if (!companyId) return [];
  const term = odataQuote(q);
  const url = `/_api/contacts?$select=contactid,fullname,emailaddress1,telephone1,mobilephone&$filter=statecode eq 0 and _parentcustomerid_value eq ${companyId} and (contains(fullname,'${term}') or contains(emailaddress1,'${term}'))&$orderby=fullname asc&$top=${CFG.maxSuggestions}`;
  const { json } = await window.PortalApi.requestJson(url, { method: "GET" });
  return (json?.value || []).map((r) => ({
    id: parseGuid(r.contactid),
    name: r.fullname || "",
    email: r.emailaddress1 || "",
    office: r.telephone1 || "",
    mobile: r.mobilephone || "",
  }));
}

  function wireTypeahead(scope, input, menu, fetcher, onPick) {
    let items = [];
    let active = -1;

    const renderMenu = () => {
      if (!items.length) {
        menu.style.display = "none";
        menu.innerHTML = "";
        return;
      }
      menu.style.display = "block";
      menu.innerHTML = items.map((it, i) => `
        <div class="es-menu-item" role="option" aria-selected="${i === active ? "true" : "false"}" data-id="${escapeHtml(it.id)}" data-name="${escapeHtml(it.name)}">
          ${escapeHtml(it.name)}
          ${it.email ? `<small>${escapeHtml(it.email)}</small>` : ""}
        </div>
      `).join("");
    };

    const pickByIndex = (i) => {
      if (i < 0 || i >= items.length) return;
      const it = items[i];
      onPick(it);
      items = [];
      active = -1;
      renderMenu();
    };

    const run = debounce(async () => {
      const q = input.value.trim();
      if (!q) {
        items = [];
        active = -1;
        renderMenu();
        return;
      }
      items = await fetcher(q);
      active = items.length ? 0 : -1;
      renderMenu();
    }, CFG.debounceMs);

    input.addEventListener("input", run);
    input.addEventListener("focus", run);

    input.addEventListener("keydown", async (e) => {
      if (e.key === "ArrowDown") {
        if (!items.length) await run();
        if (items.length) {
          e.preventDefault();
          active = Math.min(active + 1, items.length - 1);
          renderMenu();
        }
      } else if (e.key === "ArrowUp") {
        if (items.length) {
          e.preventDefault();
          active = Math.max(active - 1, 0);
          renderMenu();
        }
      } else if (e.key === "Enter") {
        if (!items.length) {
          await run();
          if (!items.length) return;
        }
        e.preventDefault();
        pickByIndex(active >= 0 ? active : 0);
      } else if (e.key === "Escape") {
        items = [];
        active = -1;
        renderMenu();
      }
    });

    menu.addEventListener("mousedown", (e) => {
      const item = e.target.closest?.(".es-menu-item");
      if (!item) return;
      e.preventDefault();
      const id = parseGuid(item.getAttribute("data-id"));
      const name = item.getAttribute("data-name") || "";
      onPick({ id, name, email: "" });
      items = [];
      active = -1;
      renderMenu();
    });

    document.addEventListener("click", (e) => {
      if (!scope.contains(e.target)) {
        items = [];
        active = -1;
        renderMenu();
      }
    });
  }

  function setSaving(saving) {
    const overlay = byId("es-drawer-overlay");
    if (!overlay) return;
    qs("[data-es-saving]", overlay).style.display = saving ? "inline-flex" : "none";
    qs("[data-es-saveclose]", overlay).disabled = saving;
    qs("[data-es-cancel]", overlay).disabled = saving;
    qs("[data-es-close]", overlay).disabled = saving;
  }

  async function openRoleDrawer(root, appId, model, roleKey) {
    const role = CFG.roles.find((r) => r.key === roleKey);
    openDrawer(`${role.title} Details`, roleDrawerHtml(roleKey, model));

    const overlay = byId("es-drawer-overlay");
    const form = qs("[data-es-roleform]", overlay);
    const companyInput = qs("[data-es-company-input]", form);
    const companyMenu = qs("[data-es-company-menu]", form);
    const companyIdEl = qs("[data-es-company-id]", form);

    const agentInput = qs("[data-es-agent-input]", form);
    const agentMenu = qs("[data-es-agent-menu]", form);
    const agentIdEl = qs("[data-es-agent-id]", form);

    const setCompany = (it) => {
      companyInput.value = it?.name || "";
      companyIdEl.value = it?.id || "";
      if (!it?.id) {
        agentInput.value = "";
        agentIdEl.value = "";
      }
      agentInput.disabled = !it?.id;
      agentInput.placeholder = it?.id ? "Search agent..." : "Select a company first";
    };

    const setAgent = (it) => {
      agentInput.value = it?.name || "";
      agentIdEl.value = it?.id || "";
    };

    wireTypeahead(form, companyInput, companyMenu, (q) => searchAccounts(q), setCompany);
    wireTypeahead(form, agentInput, agentMenu, (q) => searchContacts(parseGuid(companyIdEl.value), q), (it) => setAgent(it));

    // Backspace-clearing: if user clears company, clear agent too (on input)
    companyInput.addEventListener("input", () => {
      if (!companyInput.value.trim()) setCompany({ id: "", name: "" });
    });
    agentInput.addEventListener("input", () => {
      if (!agentInput.value.trim()) setAgent({ id: "", name: "" });
    });

    const save = async () => {
      const companyId = parseGuid(companyIdEl.value);
      const agentId = parseGuid(agentIdEl.value);

      if (agentId && !companyId) {
        showToast(root, `${role.title}: Company is required when an Agent is selected.`, false);
        return false;
      }

      setSaving(true);
      try {
        await setLookupRef(appId, role.companyNav, role.companyTargetSet, companyId || "");
        await setLookupRef(appId, role.agentNav, role.agentTargetSet, agentId || "");
        showToast(root, "", true);

        const nextModel = await hydrate(root, appId);
        return nextModel;
      } catch (e) {
        const msg = window.PortalApi.parseDataverseError?.(e.body || "")?.message || e.message || "Save failed";
        showToast(root, `Save failed: ${msg}`, false);
        return false;
      } finally {
        setSaving(false);
      }
    };

    qs("[data-es-saveclose]", overlay).onclick = async () => {
      const res = await save();
      if (res) closeDrawer();
    };
  }

  async function openTxDrawer(root, appId, model) {
    openDrawer("Transaction Details", txDrawerHtml(model));
    const overlay = byId("es-drawer-overlay");
    const form = qs("[data-es-txform]", overlay);

    const save = async () => {
      const payload = {};
      for (const f of CFG.txFields) {
        const el = qs(`[data-es-tx="${f.logical}"]`, form);
        if (!el) continue;
        if (f.type === "optionset") continue;

        const raw = el.value.trim();
        if (!raw) {
          payload[f.logical] = null;
          continue;
        }

        if (f.type === "money") {
          const n = Number(raw.replace(/[^0-9.\-]/g, ""));
          payload[f.logical] = Number.isNaN(n) ? null : n;
        } else if (f.type === "date") {
          payload[f.logical] = raw; // yyyy-mm-dd
        } else {
          payload[f.logical] = raw;
        }
      }

      setSaving(true);
      try {
        await patchHeader(appId, payload);
        const nextModel = await hydrate(root, appId);
        return nextModel;
      } catch (e) {
        const msg = window.PortalApi.parseDataverseError?.(e.body || "")?.message || e.message || "Save failed";
        showToast(root, `Save failed: ${msg}`, false);
        return false;
      } finally {
        setSaving(false);
      }
    };

    qs("[data-es-saveclose]", overlay).onclick = async () => {
      const res = await save();
      if (res) closeDrawer();
    };
  }

  function bindGridActions(root, appId, getModel) {
    root.addEventListener("click", async (e) => {
      const btn = e.target.closest?.("[data-es-open]");
      if (!btn) return;
      const key = btn.getAttribute("data-es-open");
      const model = getModel();
      if (!model) return;

      showToast(root, "");
      if (key === "tx") await openTxDrawer(root, appId, model);
      else await openRoleDrawer(root, appId, model, key);
    });
  }

  function updateApplicantsBadge() {
    const card = byId("la-applicants");
    if (!card) return;

    const hd = qs(".la-card__hd", card);
    if (!hd || hd.querySelector("[data-app-badge]")) return;

    const count = card.querySelectorAll("tbody tr").length;
    const badge = document.createElement("span");
    badge.className = "la-badge la-badge--muted";
    badge.setAttribute("data-app-badge", "1");
    badge.textContent = `${count} Applicant${count === 1 ? "" : "s"}`;

    hd.appendChild(badge);
  }

  async function init() {
    const root = byId(CFG.rootId);
    const appId = getAppId();
    if (!root || !appId) return;

    buildHeaderUi(root);

    let model = null;
    const getModel = () => model;

    bindGridActions(root, appId, getModel);

    try {
      model = await hydrate(root, appId);
      updateApplicantsBadge();
      if (createdFlow()) {
        // reduce clicks: jump into edit immediately for newly created apps
        await openRoleDrawer(root, appId, model, "escrow");
      }
    } catch (e) {
      console.error("[Resale-Escrow] init failed:", e);
      showToast(root, `Init failed: ${e.message}`, false);
    }
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();

/* ===================================================================
 * Agreement Section (vms_agreementrecord) — mimics Financial Assets UX
 * =================================================================== */
(() => {
  "use strict";

  if (!window.PortalApi?.requestJson) return;

  const CFG = {
    agreementSet: "vms_agreementrecords",
    appSet: "vms_applicationheaders",
    sessionAppIdKey: "la_current_application_id",
    dom: {
      finGridsId: "ra-fin-grids",
      applicantsRootId: "la-applicants",
      rootId: "ra-agreements",
      groupsId: "ra-agreement-grids",
      msgId: "ra-agreement-msg",
      saveBtnId: "ra-save-agreements",
      statusBadgeId: "ra-agreement-status",

      drawerOverlayId: "ag-drawer-overlay",
      drawerTitleId: "ag-drawer-title",
      drawerBodyId: "ag-drawer-body",
      drawerSaveId: "ag-drawer-save",
      drawerCancelId: "ag-drawer-cancel",
      drawerCloseId: "ag-drawer-close",
    },
    top: 5000,
    prefer: 'odata.include-annotations="*"',
  };

  const byId = (id) => document.getElementById(id);
  const norm = (s) => (s ?? "").toString().trim();
  const isGuid = (v) =>
    /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(
      String(v || "")
    );
  const parseGuid = (v) => {
    const s = norm(v).toLowerCase();
    return isGuid(s) ? s : "";
  };
  const escapeHtml = (s) =>
    String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  const fmtLookup = (rec, lookupLogicalName) => {
    const key = `_${lookupLogicalName}_value@OData.Community.Display.V1.FormattedValue`;
    return norm(rec?.[key]);
  };

  const getAppId = () => {
    const usp = new URLSearchParams(window.location.search || "");
    const q = parseGuid(usp.get("id") || usp.get("appid") || "");
    if (q) return q;
    return parseGuid(sessionStorage.getItem(CFG.sessionAppIdKey) || "");
  };

  function ensureAgreementSection() {
    let root = byId(CFG.dom.rootId);
    if (root) return root;

    // Anchor after Financial Assets card if possible
    const fin = byId(CFG.dom.finGridsId);
    const finCard = fin?.closest?.(".la-card") || fin?.closest?.("section") || fin?.parentElement;

    // Fallback: insert before Applicants card, else append to main content.
    const applicants = byId(CFG.dom.applicantsRootId);
    const appCard = applicants?.closest?.(".la-card") || applicants?.closest?.("section") || applicants;

    root = document.createElement("section");
    root.className = "la-card";
    root.id = CFG.dom.rootId;
    root.innerHTML = `
      <div class="la-card__hd">
        <span>Agreements</span>
        <span class="ra-badge" id="${CFG.dom.statusBadgeId}">0 / 0 Signed</span>
      </div>
      <div class="la-card__bd">
        <div class="ra-note">
          Review each statement and check the box to sign. Use <b>Details</b> to read the full statement.
        </div>

        <div id="${CFG.dom.groupsId}"></div>

        <div class="la-ftr" style="margin-top:12px; display:flex; gap:10px; align-items:center;">
          <button type="button" class="btn btn-primary" id="${CFG.dom.saveBtnId}">Save Agreements</button>
          <div id="${CFG.dom.msgId}" class="ra-msg" style="display:none;"></div>
        </div>
      </div>
    `;

    if (finCard && finCard.parentElement) finCard.insertAdjacentElement("afterend", root);
    else if (appCard && appCard.parentElement) appCard.parentElement.insertBefore(root, appCard);
    else (document.querySelector("main") || document.body).appendChild(root);

    return root;
  }

  function ensureDrawerShell() {
    if (byId(CFG.dom.drawerOverlayId)) return;

    const overlay = document.createElement("div");
    overlay.id = CFG.dom.drawerOverlayId;
    overlay.className = "ra-drawer-overlay";
    overlay.innerHTML = `
      <div class="ra-drawer" role="dialog" aria-modal="true" aria-labelledby="${CFG.dom.drawerTitleId}">
        <div class="ra-drawer-hd">
          <div class="ra-drawer-title" id="${CFG.dom.drawerTitleId}">Agreement</div>
          <button type="button" class="ra-drawer-x" id="${CFG.dom.drawerCloseId}" aria-label="Close">×</button>
        </div>
        <div class="ra-drawer-bd" id="${CFG.dom.drawerBodyId}"></div>
        <div class="ra-drawer-ft">
          <button type="button" class="btn btn-default" id="${CFG.dom.drawerCancelId}">Cancel</button>
          <button type="button" class="btn btn-primary" id="${CFG.dom.drawerSaveId}">Save & Close</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const close = () => overlay.classList.remove("open");
    byId(CFG.dom.drawerCloseId)?.addEventListener("click", close);
    byId(CFG.dom.drawerCancelId)?.addEventListener("click", close);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && overlay.classList.contains("open")) close();
    });
  }

  const STATE = {
    appId: "",
    groups: new Map(),
    rowsById: new Map(),
    dirty: new Set(),
  };

  function showMsg(text, ok = true) {
    const el = byId(CFG.dom.msgId);
    if (!el) return;
    el.style.display = "block";
    el.classList.toggle("ra-msg--ok", !!ok);
    el.classList.toggle("ra-msg--err", !ok);
    el.textContent = text;
  }

  function hideMsg() {
    const el = byId(CFG.dom.msgId);
    if (!el) return;
    el.style.display = "none";
    el.textContent = "";
  }

  function signerKind(vm) {
    return vm.contactId ? "Contact" : vm.applicantId ? "Applicant" : "Signer";
  }

  function isContact(vm) {
    return !!vm.contactId;
  }

  function groupKeyFor(rec) {
    const byFx = norm(rec?.vms_agreementtypefx);
    const byLookup = fmtLookup(rec, "vms_agreementtype");
    return byFx || byLookup || "Agreement";
  }

  function computeCounts() {
    let total = 0;
    let signed = 0;

    STATE.groups.forEach((g) => {
      g.total = g.records.length;
      g.signed = g.records.filter((r) => !!r.signed).length;
      total += g.total;
      signed += g.signed;
    });

    const badge = byId(CFG.dom.statusBadgeId);
    if (badge) badge.textContent = `${signed} / ${total} Signed`;
  }

  function normalizeRecords(recs) {
    STATE.groups.clear();
    STATE.rowsById.clear();
    STATE.dirty.clear();

    for (const r of recs) {
      const id = parseGuid(r?.vms_agreementrecordid);
      if (!id) continue;

      const contactId = parseGuid(r?._vms_contact_value);
      const applicantId = parseGuid(r?._vms_applicant_value);
      const signed = contactId ? !!r?.vms_contactsigned : applicantId ? !!r?.vms_applicantsigned : false;

      const vm = {
        id,
        recordNumber: norm(r?.vms_recordnumber),
        agreementTypeName: groupKeyFor(r),
        statement: norm(r?.vms_statement),
        contactId,
        applicantId,
        signerName:
          (contactId && fmtLookup(r, "vms_contact")) ||
          (applicantId && fmtLookup(r, "vms_applicant")) ||
          "—",
        signed,
      };

      const gKey = vm.agreementTypeName;
      if (!STATE.groups.has(gKey)) STATE.groups.set(gKey, { key: gKey, title: gKey, records: [] });
      STATE.groups.get(gKey).records.push(vm);
      STATE.rowsById.set(vm.id, vm);
    }

    STATE.groups.forEach((g) => {
      g.records.sort((a, b) => (a.signerName + a.statement).localeCompare(b.signerName + b.statement));
    });

    computeCounts();
  }

  function statementSnippet(s) {
    const t = norm(s);
    if (t.length <= 70) return t;
    return `${t.slice(0, 67)}…`;
  }

  function rowHtml(vm) {
    const kind = signerKind(vm);
    const checked = vm.signed ? "checked" : "";
    const disabled = !vm.contactId && !vm.applicantId ? "disabled" : "";
    return `
      <tr data-agid="${escapeHtml(vm.id)}">
        <td>
          <div style="font-weight:600;">${escapeHtml(vm.signerName)}</div>
          <div><small class="text-muted">${escapeHtml(kind)}</small></div>
        </td>
        <td>${escapeHtml(statementSnippet(vm.statement) || "—")}</td>
        <td><input type="checkbox" class="ra-ag-signed" ${checked} ${disabled} /></td>
        <td>
          <div class="ra-row-actions">
            <button type="button" class="ra-link ra-ag-details">Details</button>
          </div>
        </td>
      </tr>
    `;
  }

  function groupCardHtml(g, idx) {
    const collapsed = idx > 0;
    const chevron = collapsed ? "▸" : "▾";
    const expanded = collapsed ? "false" : "true";
    const cardClass = collapsed ? "ra-grid-card is-collapsed" : "ra-grid-card";

    return `
      <div class="${cardClass}" data-agtype="${escapeHtml(g.key)}">
        <div class="ra-grid-hd">
          <div class="ra-grid-title">${escapeHtml(g.title)}</div>
          <div class="ra-grid-actions" style="display:flex; gap:10px; align-items:center;">
            <span class="ra-badge" data-agbadge="${escapeHtml(g.key)}">${g.signed} / ${g.total} Signed</span>
            <button type="button" class="ra-collapse" aria-label="Toggle ${escapeHtml(g.title)}" aria-expanded="${expanded}" data-agcollapse="${escapeHtml(g.key)}">${chevron}</button>
          </div>
        </div>
        <div class="ra-grid-bd">
          <div class="table-responsive">
            <table class="table table-striped">
              <thead>
                <tr>
                  <th style="width:220px;">Signer</th>
                  <th>Statement</th>
                  <th style="width:110px;">Signed</th>
                  <th style="width:120px;"></th>
                </tr>
              </thead>
              <tbody data-agbody="${escapeHtml(g.key)}"></tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }

  function renderAll() {
    const groupsRoot = byId(CFG.dom.groupsId);
    if (!groupsRoot) return;

    const groups = Array.from(STATE.groups.values());
    groups.forEach((g) => {
      g.total = g.records.length;
      g.signed = g.records.filter((r) => !!r.signed).length;
    });

    groupsRoot.innerHTML = groups.map((g, idx) => groupCardHtml(g, idx)).join("");

    for (const g of groups) {
      const body = groupsRoot.querySelector(`tbody[data-agbody="${CSS.escape(g.key)}"]`);
      if (!body) continue;
      body.innerHTML = g.records.map((r) => rowHtml(r)).join("");
    }

    computeCounts();
  }

  function updateGroupBadge(gKey) {
    const g = STATE.groups.get(gKey);
    if (!g) return;
    g.total = g.records.length;
    g.signed = g.records.filter((r) => !!r.signed).length;

    const badge = document.querySelector(`[data-agbadge="${CSS.escape(gKey)}"]`);
    if (badge) badge.textContent = `${g.signed} / ${g.total} Signed`;

    computeCounts();
  }

  function openDrawerForRecord(vm) {
    ensureDrawerShell();

    const overlay = byId(CFG.dom.drawerOverlayId);
    const titleEl = byId(CFG.dom.drawerTitleId);
    const bodyEl = byId(CFG.dom.drawerBodyId);
    const saveBtn = byId(CFG.dom.drawerSaveId);
    if (!overlay || !titleEl || !bodyEl || !saveBtn) return;

    titleEl.textContent = `${vm.agreementTypeName} — ${vm.signerName}`;

    bodyEl.innerHTML = `
      <div class="ra-note" style="margin-bottom:12px;">
        <div style="font-weight:700; margin-bottom:4px;">${escapeHtml(vm.recordNumber || "Agreement")}</div>
        <div><small class="text-muted">${escapeHtml(signerKind(vm))}</small></div>
      </div>

      <div style="margin-bottom:10px;">
        <label style="font-weight:700; display:block; margin-bottom:6px;">Statement</label>
        <div style="border:1px solid var(--slds-border,#d8dde6); border-radius:12px; padding:12px; background:#fff; white-space:pre-wrap;">
          ${escapeHtml(vm.statement || "—")}
        </div>
      </div>

      <div style="margin-top:12px;">
        <label style="display:flex; gap:10px; align-items:center;">
          <input type="checkbox" id="ag-drawer-signed" ${vm.signed ? "checked" : ""} />
          <span>I have read and agree</span>
        </label>
      </div>
    `;

    const signedEl = bodyEl.querySelector("#ag-drawer-signed");
    if (signedEl) {
      signedEl.addEventListener("change", () => setSigned(vm.id, !!signedEl.checked, { fromDrawer: true }));
    }

    saveBtn.onclick = async () => {
      await saveOne(vm.id);
      overlay.classList.remove("open");
    };

    overlay.classList.add("open");
  }

  function setSigned(id, checked, { fromDrawer = false } = {}) {
    const vm = STATE.rowsById.get(id);
    if (!vm) return;
    vm.signed = !!checked;
    STATE.dirty.add(id);

    if (fromDrawer) {
      const tr = document.querySelector(`tr[data-agid="${CSS.escape(id)}"]`);
      const cb = tr?.querySelector?.("input.ra-ag-signed");
      if (cb) cb.checked = vm.signed;
    }

    updateGroupBadge(vm.agreementTypeName);
  }

  async function saveOne(id) {
    const vm = STATE.rowsById.get(id);
    if (!vm) return;

    hideMsg();

    const patch = isContact(vm)
      ? { vms_contactsigned: !!vm.signed }
      : { vms_applicantsigned: !!vm.signed };

    await window.PortalApi.requestJson(`/_api/${CFG.agreementSet}(${vm.id})`, {
      method: "PATCH",
      body: patch,
      prefer: "return=minimal",
    });

    STATE.dirty.delete(id);
    showMsg("Saved.", true);
  }

  async function saveAll() {
    const ids = Array.from(STATE.dirty.values());
    if (!ids.length) return showMsg("No changes to save.", true);

    const btn = byId(CFG.dom.saveBtnId);
    if (btn) btn.disabled = true;
    hideMsg();

    try {
      for (const id of ids) await saveOne(id);
      showMsg(`Saved ${ids.length} agreement${ids.length === 1 ? "" : "s"}.`, true);
    } catch (e) {
      console.error("[Resale-Agreements] save failed:", e);
      showMsg(`Save failed: ${e.message}`, false);
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  async function loadAgreements(appId) {
    const select = [
      "vms_agreementrecordid",
      "vms_recordnumber",
      "vms_agreementtypefx",
      "vms_statement",
      "vms_applicantsigned",
      "vms_contactsigned",
      "_vms_application_value",
      "_vms_agreementtype_value",
      "_vms_agreementgroup_value",
      "_vms_contact_value",
      "_vms_applicant_value",
      "createdon",
      "modifiedon",
      "statecode",
      "statuscode",
    ].join(",");

    const filter = `_vms_application_value eq ${appId}`;

    const url =
      `/_api/${CFG.agreementSet}` +
      `?$select=${encodeURIComponent(select)}` +
      `&$filter=${encodeURIComponent(filter)}` +
      `&$orderby=${encodeURIComponent("createdon asc")}` +
      `&$top=${CFG.top}`;

    const { json } = await window.PortalApi.requestJson(url, { method: "GET", prefer: CFG.prefer });
    return json?.value || [];
  }

  function wireEvents() {
    const groupsRoot = byId(CFG.dom.groupsId);
    if (!groupsRoot) return;

    groupsRoot.addEventListener("click", (e) => {
      const collapseBtn = e.target.closest?.("button.ra-collapse[data-agcollapse]");
      if (collapseBtn) {
        const card = collapseBtn.closest?.(".ra-grid-card");
        if (!card) return;
        const isCollapsed = card.classList.toggle("is-collapsed");
        collapseBtn.textContent = isCollapsed ? "▸" : "▾";
        collapseBtn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
        return;
      }

      const detailsBtn = e.target.closest?.("button.ra-ag-details");
      if (detailsBtn) {
        const tr = detailsBtn.closest?.("tr[data-agid]");
        const id = parseGuid(tr?.getAttribute?.("data-agid") || "");
        const vm = STATE.rowsById.get(id);
        if (vm) openDrawerForRecord(vm);
      }
    });

    groupsRoot.addEventListener("change", (e) => {
      const cb = e.target.closest?.("input.ra-ag-signed");
      if (!cb) return;
      const tr = cb.closest?.("tr[data-agid]");
      const id = parseGuid(tr?.getAttribute?.("data-agid") || "");
      if (id) setSigned(id, !!cb.checked);
    });

    byId(CFG.dom.saveBtnId)?.addEventListener("click", saveAll);
  }

  async function init() {
    const root = ensureAgreementSection();
    if (!root) return;

    const appId = getAppId();
    if (!appId) return showMsg("Missing application id in URL. Open an application from the dashboard.", false);

    STATE.appId = appId;

    try {
      const recs = await loadAgreements(appId);
      normalizeRecords(recs);
      renderAll();
      wireEvents();
      if (!recs.length) showMsg("No agreements found for this application.", true);
    } catch (e) {
      console.error("[Resale-Agreements] init failed:", e);
      showMsg(`Agreements failed to load: ${e.message}`, false);
    }
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
