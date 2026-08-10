/* ==========================================================================
   weather-bg.js — 定位 + 当日天气信息（精简版 v2）
   --------------------------------------------------------------------------
   变更（2026-08-10）：移除 Canvas 动态天气背景渲染，仅保留
   IP 定位 → 当日天气 → 左下角信息 pill（页面恢复原有静态背景）。
   链路：ipwho.is → ipinfo.io → freeipapi.com → ipapi.co（四级 fallback）
        → Open-Meteo 天气（WMO code 映射中文）
   特性：
   - 全部 API 失败时兜底「未知位置 · 晴」，绝不阻塞页面
   - pill 显示：城市 · 天气 · 温度 · IP（IPv6 截断）
   ========================================================================== */
(function () {
  "use strict";
  if (typeof document === "undefined") return;

  /* ---------- 1. 工具 ---------- */
  function fetchJson(url, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
      var timer = setTimeout(function () { if (ctrl) ctrl.abort(); reject(new Error("timeout")); }, timeoutMs || 6000);
      fetch(url, ctrl ? { signal: ctrl.signal } : {})
        .then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); })
        .then(function (d) { clearTimeout(timer); resolve(d); })
        .catch(function (e) { clearTimeout(timer); reject(e); });
    });
  }

  /* ---------- 2. IP 定位（四级 fallback，2026-08-10 实测修正） ---------- */
  var IP_PROVIDERS = [
    { url: "https://ipwho.is/", parse: function (d) {
        if (!d || !d.success) throw new Error("ipwho fail");
        return { lat: parseFloat(d.latitude), lon: parseFloat(d.longitude), city: d.city || d.region || "本地", ip: d.ip || "" };
      } },
    { url: "https://ipinfo.io/json", parse: function (d) {
        if (!d || !d.loc) throw new Error("ipinfo fail");
        var p = String(d.loc).split(",");
        return { lat: parseFloat(p[0]), lon: parseFloat(p[1]), city: d.city || d.region || "本地", ip: d.ip || "" };
      } },
    { url: "https://freeipapi.com/api/json", parse: function (d) {
        if (!d || d.latitude == null) throw new Error("freeipapi fail");
        return { lat: parseFloat(d.latitude), lon: parseFloat(d.longitude), city: d.cityName || d.countryName || "本地", ip: d.ipAddress || "" };
      } },
    { url: "https://ipapi.co/json/", parse: function (d) {
        if (!d || (!d.latitude && !d.lat)) throw new Error("ipapi fail");
        return {
          lat: parseFloat(d.latitude != null ? d.latitude : d.lat),
          lon: parseFloat(d.longitude != null ? d.longitude : d.lon),
          city: d.city || d.region || d.province || "本地",
          ip: d.ip || ""
        };
      } }
  ];

  function locate() {
    var chain = IP_PROVIDERS.slice();
    function next() {
      if (!chain.length) return Promise.resolve({ lat: null, lon: null, city: "未知位置", ip: "", fallback: true });
      var p = chain.shift();
      return fetchJson(p.url, 5000)
        .then(function (d) { return p.parse(d); })
        .catch(function () { return next(); });
    }
    return next();
  }

  /* ---------- 3. 天气查询（Open-Meteo） ---------- */
  var WMO = {
    clear: [0],
    cloudy: [1, 2, 3],
    fog: [45, 48],
    drizzle: [51, 53, 55, 56, 57],
    rain: [61, 63, 65, 66, 67, 80, 81, 82],
    snow: [71, 73, 75, 77, 85, 86],
    thunder: [95, 96, 99]
  };
  function wmoToKey(code) {
    for (var k in WMO) { if (WMO[k].indexOf(code) >= 0) return k; }
    return "cloudy";
  }
  var WEATHER_CN = { clear: "晴", cloudy: "多云", fog: "雾", drizzle: "毛毛雨", rain: "雨", snow: "雪", thunder: "雷阵雨" };

  function getWeather(lat, lon) {
    if (lat == null || lon == null) return Promise.resolve({ key: "clear", temp: null });
    var url = "https://api.open-meteo.com/v1/forecast?latitude=" + lat + "&longitude=" + lon +
              "&current_weather=true&timezone=auto&forecast_days=1";
    return fetchJson(url, 7000)
      .then(function (d) {
        var cw = d && d.current_weather;
        if (!cw) throw new Error("no weather");
        return { key: wmoToKey(cw.weathercode), temp: Math.round(cw.temperature) };
      })
      .catch(function () { return { key: "clear", temp: null, fallback: true }; });
  }

  /* ---------- 4. 左下角信息 pill ---------- */
  function shortIp(ip) {
    if (!ip) return "";
    if (ip.indexOf(":") >= 0 && ip.length > 20) {
      // IPv6 截断：保留前 8 与后 4
      return ip.slice(0, 8) + "…" + ip.slice(-4);
    }
    return ip;
  }

  function addPill(geo, weather) {
    var old = document.getElementById("weatherPill");
    if (old) old.remove();

    var pill = document.createElement("div");
    pill.id = "weatherPill";
    var emoji = { clear: "☀️", cloudy: "⛅", fog: "🌫️", drizzle: "🌦️", rain: "🌧️", snow: "🌨️", thunder: "⛈️" }[weather.key] || "☀️";
    var tempTxt = weather.temp != null ? " · " + weather.temp + "°C" : "";
    var ipTxt = geo.ip ? ' · <span class="wp-ip">' + shortIp(geo.ip) + "</span>" : "";
    pill.innerHTML = '<span class="wp-ico">' + emoji + '</span><span class="wp-txt">' + geo.city +
      " · " + (WEATHER_CN[weather.key] || "晴") + tempTxt + ipTxt + "</span>";

    var st = document.createElement("style");
    st.textContent = [
      "#weatherPill{position:fixed;left:22px;bottom:22px;z-index:960;display:flex;align-items:center;gap:8px;",
      "background:rgba(255,255,255,.72);backdrop-filter:blur(14px) saturate(1.1);-webkit-backdrop-filter:blur(14px) saturate(1.1);",
      "border:1px solid rgba(255,255,255,.8);border-radius:999px;padding:8px 16px;",
      "font-family:'Space Mono',Consolas,monospace;font-size:11px;font-weight:700;letter-spacing:.06em;color:#4A6173;",
      "box-shadow:inset 0 1px 0 rgba(255,255,255,.7),0 20px 50px -32px rgba(20,19,16,.4);",
      "transition:transform .2s cubic-bezier(.2,.85,.25,1);}",
      "#weatherPill:hover{transform:translateY(-2px);}",
      "#weatherPill .wp-ico{font-size:14px;line-height:1;}",
      "#weatherPill .wp-ip{opacity:.65;font-weight:400;}",
      "@media (max-width:768px){#weatherPill{left:14px;bottom:14px;max-width:calc(100vw - 28px);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}",
      "@media (prefers-color-scheme:dark){#weatherPill{background:rgba(22,30,38,.8);border-color:rgba(255,255,255,.14);color:#A9BCCB;}}"
    ].join("");
    document.head.appendChild(st);
    document.body.appendChild(pill);
  }

  /* ---------- 5. 启动 ---------- */
  function init() {
    locate()
      .then(function (geo) {
        return getWeather(geo.lat, geo.lon).then(function (w) { return { geo: geo, weather: w }; });
      })
      .then(function (res) { addPill(res.geo, res.weather); })
      .catch(function () {
        addPill({ city: "未知位置", ip: "" }, { key: "clear", temp: null });
      });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
