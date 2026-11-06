"use strict";
/* Облако (Supabase): вход по email и паролю, таблица банок и закрытое хранилище фото.
   Публичный (publishable) ключ можно держать в коде: доступ к данным ограничивают правила в базе. */

const Cloud = (() => {
  const URL_BASE = "https://nmttxdfozkdxiarqchrs.supabase.co";
  const KEY = "sb_publishable_lUJ9jRJoptJXeu0iKyNxIg_SXONAXLA";
  const BUCKET = "can-photos";
  const LS = "cans-session-v1";
  const SITE = location.origin + location.pathname;

  const listeners = [];
  let session = null;
  try { session = JSON.parse(localStorage.getItem(LS) || "null"); } catch (e) { session = null; }

  function setSession(s) {
    session = s;
    try { s ? localStorage.setItem(LS, JSON.stringify(s)) : localStorage.removeItem(LS); } catch (e) { /* storage blocked */ }
    listeners.forEach((f) => { try { f(session); } catch (e) { console.error(e); } });
  }
  function fromToken(j) {
    return {
      access_token: j.access_token,
      refresh_token: j.refresh_token,
      expires_at: j.expires_at || Math.floor(Date.now() / 1000) + (+j.expires_in || 3600),
      user: { id: j.user.id, email: j.user.email },
    };
  }
  function fail(code, message, status) { const e = new Error(message || code); e.code = code; e.status = status; return e; }

  async function request(path, { method = "GET", body, headers = {}, auth = true, raw = false, retried = false } = {}) {
    const h = { apikey: KEY, ...headers };
    if (auth) {
      const t = await accessToken();
      if (!t) throw fail("not_signed_in", "Нужно войти");
      h.Authorization = "Bearer " + t;
    }
    let payload = body;
    if (body !== undefined && !(body instanceof Blob)) { h["Content-Type"] = "application/json"; payload = JSON.stringify(body); }
    let res;
    try { res = await fetch(URL_BASE + path, { method, headers: h, body: payload, cache: "no-store" }); }
    catch (e) { throw fail("offline", "Нет связи с сервером"); }
    if (res.status === 401 && auth && !retried && session) {
      await refresh(true);
      return request(path, { method, body, headers, auth, raw, retried: true });
    }
    if (!res.ok) {
      let j = {};
      try { j = await res.json(); } catch (e) { /* not json */ }
      throw fail(j.error_code || j.code || j.error || "http_" + res.status, j.msg || j.message || j.error_description || res.statusText, res.status);
    }
    if (raw) return res;
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  let refreshing = null;
  function refresh(force) {
    if (!session) return Promise.resolve(null);
    if (!force && session.expires_at - 60 > Date.now() / 1000) return Promise.resolve(session);
    if (!refreshing) {
      refreshing = (async () => {
        try {
          const j = await request("/auth/v1/token?grant_type=refresh_token", { method: "POST", auth: false, body: { refresh_token: session.refresh_token } });
          setSession(fromToken(j));
        } catch (e) {
          // offline: keep the session and the cached collection; a rejected token: sign out
          if (e.code !== "offline") setSession(null);
          throw e;
        } finally { refreshing = null; }
        return session;
      })();
    }
    return refreshing;
  }
  async function accessToken() { if (!session) return null; await refresh(false); return session && session.access_token; }

  return {
    user: () => (session ? session.user : null),
    onChange: (f) => listeners.push(f),

    async signIn(email, password) {
      const j = await request("/auth/v1/token?grant_type=password", { method: "POST", auth: false, body: { email, password } });
      setSession(fromToken(j));
    },
    // returns true when the account still has to be confirmed by email
    async signUp(email, password) {
      const j = await request("/auth/v1/signup?redirect_to=" + encodeURIComponent(SITE), { method: "POST", auth: false, body: { email, password } });
      if (j && j.access_token) { setSession(fromToken(j)); return false; }
      return true;
    },
    async sendReset(email) {
      await request("/auth/v1/recover?redirect_to=" + encodeURIComponent(SITE), { method: "POST", auth: false, body: { email } });
    },
    async setPassword(password) { await request("/auth/v1/user", { method: "PUT", body: { password } }); },
    async signOut() {
      try { await request("/auth/v1/logout", { method: "POST" }); } catch (e) { /* signing out locally is enough */ }
      setSession(null);
    },
    // links from emails come back as #access_token=…&type=…  or  #error=…
    async fromLink() {
      const h = new URLSearchParams(location.hash.slice(1));
      if (!h.has("access_token") && !h.has("error") && !h.has("error_code")) return null;
      history.replaceState(null, "", location.pathname + location.search);
      if (h.get("error") || h.get("error_code")) return { error: h.get("error_code") || h.get("error"), message: h.get("error_description") };
      const user = await request("/auth/v1/user", { auth: false, headers: { Authorization: "Bearer " + h.get("access_token") } });
      setSession(fromToken({ access_token: h.get("access_token"), refresh_token: h.get("refresh_token"), expires_at: +h.get("expires_at") || 0, expires_in: h.get("expires_in"), user }));
      return { type: h.get("type") };
    },

    list: () => request("/rest/v1/cans?select=*&order=num.asc"),
    upsert: (rows) => request("/rest/v1/cans?on_conflict=id", { method: "POST", body: rows, headers: { Prefer: "resolution=merge-duplicates,return=minimal" } }),
    remove: (ids) => request("/rest/v1/cans?id=in.(" + ids.map((i) => '"' + i + '"').join(",") + ")", { method: "DELETE" }),

    upload: (path, blob) => request(`/storage/v1/object/${BUCKET}/${path}`, { method: "POST", body: blob, headers: { "Content-Type": "image/jpeg", "x-upsert": "true", "cache-control": "31536000" } }),
    download: async (path) => (await request(`/storage/v1/object/authenticated/${BUCKET}/${path}`, { raw: true })).blob(),
    removeFiles: (paths) => request(`/storage/v1/object/${BUCKET}`, { method: "DELETE", body: { prefixes: paths } }),

    // readable messages for the sign-in form
    explain(e) {
      const c = String(e && e.code || ""), m = String(e && e.message || "").toLowerCase();
      if (c === "offline") return "Нет интернета. Проверьте подключение и попробуйте снова.";
      if (c === "invalid_credentials" || m.includes("invalid login")) return "Неверный email или пароль.";
      if (c === "email_not_confirmed" || m.includes("not confirmed")) return "Email ещё не подтверждён. Откройте ссылку из письма и войдите снова.";
      if (c === "user_already_exists" || m.includes("already registered")) return "Такой email уже зарегистрирован. Войдите с паролем или восстановите его.";
      if (c === "weak_password" || m.includes("password should")) return "Пароль слишком простой: нужно не меньше 6 символов.";
      if (c === "validation_failed" || m.includes("valid email") || m.includes("invalid format")) return "Проверьте, правильно ли написан email.";
      if (c === "over_email_send_rate_limit" || c === "over_request_rate_limit" || e && e.status === 429) return "Слишком много попыток. Подождите минуту и попробуйте снова.";
      if (c === "same_password") return "Новый пароль совпадает со старым.";
      if (c === "otp_expired") return "Ссылка из письма устарела. Запросите новую.";
      return "Не получилось: " + (e && e.message || "неизвестная ошибка");
    },
  };
})();
