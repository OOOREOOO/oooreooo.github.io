/* ==========================================================================
   weather-bg.js — 定位 + 当日天气信息（精简版 v3.5）
   --------------------------------------------------------------------------
   变更（2026-08-11）：
   v3.5 移除全部地理信息请求（Geolocation 授权弹窗/权限预检/反向地理编码），仅保留手动城市覆盖 + IP 定位。
   v3.4 评审优化：pill 加 ⌄ 可点暗示；prefers-reduced-motion 关闭面板动画。
   v3.3 自动物理定位强化：
     1. 权限预检（navigator.permissions.query）：已授权(granted) → 静默自动
        获取物理位置；未决定(prompt) → 正常请求授权弹窗；被拒/不支持 →
        立即降级 IP 定位（不做无谓等待）。
     2. 高精度定位 enableHighAccuracy=true（Windows 位置服务/蜂窝/GPS），
        超时放宽至 10s；位置缓存 5 分钟。
     3. 渐进策略保持：IP 先快速显示 pill，物理定位成功后静默刷新覆盖。
   v3.2 本机定位优先：浏览器 Geolocation 优先，失败降级 IP。
   v3.1 安全与稳健：textContent 构建、z-index 90/100、style 单例等。
   v3.0 手动城市覆盖：点击 pill 搜索城市，localStorage 记忆坐标。
   定位优先级：手动覆盖 > 物理定位(Geolocation) > IP 定位 > 「未知位置 · 晴」兜底。
   ========================================================================== */
(function () {
  "use strict";
  if (typeof document === "undefined") return;

  var LS_KEY = "oreo_weather_city"; // {city, lat, lon, ts}
  var STYLE_ID = "weatherPillStyle";
  var PANEL_ID = "weatherPanel";
  var SRC_MANUAL = "manual", SRC_GPS = "gps", SRC_IP = "ip";

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

  function lsGet() {
    try {
      var v = JSON.parse(localStorage.getItem(LS_KEY) || "null");
      if (!v || typeof v.city !== "string" || typeof v.lat !== "number" || typeof v.lon !== "number") return null;
      return v;
    } catch (e) { return null; }
  }
  function lsSet(v) {
    try {
      if (v) localStorage.setItem(LS_KEY, JSON.stringify(v));
      else localStorage.removeItem(LS_KEY);
    } catch (e) { /* 隐私模式等场景忽略 */ }
  }

  /* ---------- 2. IP 定位（四级 fallback；v3.5 起为自动定位唯一来源，不再请求本机地理信息） ---------- */
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
      if (!chain.length) return Promise.resolve({ lat: null, lon: null, city: "未知位置", ip: "", fallback: true, source: SRC_IP });
      var p = chain.shift();
      return fetchJson(p.url, 5000)
        .then(function (d) { var r = p.parse(d); r.source = SRC_IP; return r; })
        .catch(function () { return next(); });
    }
    return next();
  }

  /* ---------- 4. 天气查询（Open-Meteo） ---------- */
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
  var EMOJI = { clear: "☀️", cloudy: "⛅", fog: "🌫️", drizzle: "🌦️", rain: "🌧️", snow: "🌨️", thunder: "⛈️" };

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

  /* ---------- 5. 城市搜索（Open-Meteo Geocoding） ---------- */
  function searchCity(name) {
    var url = "https://geocoding-api.open-meteo.com/v1/search?name=" +
              encodeURIComponent(name) + "&count=5&language=zh&format=json";
    return fetchJson(url, 7000).then(function (d) {
      if (!d || !d.results || !d.results.length) throw new Error("no city");
      return d.results.map(function (r) {
        return {
          name: r.name,
          admin: r.admin1 || "",
          country: r.country || "",
          lat: parseFloat(r.latitude),
          lon: parseFloat(r.longitude)
        };
      });
    });
  }

  /* ---------- 6. 样式（单例，避免重复追加） ---------- */
  var CSS = [
    "#weatherPill{position:fixed;left:22px;bottom:22px;z-index:90;display:flex;align-items:center;gap:8px;",
    "background:rgba(255,255,255,.72);backdrop-filter:blur(14px) saturate(1.1);-webkit-backdrop-filter:blur(14px) saturate(1.1);",
    "border:1px solid rgba(255,255,255,.8);border-radius:999px;padding:8px 16px;cursor:pointer;",
    "font-family:'Space Mono',Consolas,monospace;font-size:11px;font-weight:700;letter-spacing:.06em;color:#4A6173;",
    "box-shadow:inset 0 1px 0 rgba(255,255,255,.7),0 20px 50px -32px rgba(20,19,16,.4);",
    "transition:transform .2s cubic-bezier(.2,.85,.25,1);}",
    "#weatherPill:hover{transform:translateY(-2px);}",
    "#weatherPill:focus-visible{outline:2px solid #0D99FF;outline-offset:2px;}",
    "#weatherPill .wp-ico{font-size:14px;line-height:1;}",
    "#weatherPill .wp-ip{opacity:.65;font-weight:400;}",
    "#weatherPill .wp-pin{font-size:11px;margin-right:2px;}",
    "#weatherPill::after{content:'\u2304';font-size:10px;margin-left:2px;opacity:.55;color:#4A6173;}",
        "@media (max-width:768px){#weatherPill{left:14px;bottom:14px;max-width:calc(100vw - 28px);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}",
    "@media (prefers-color-scheme:dark){#weatherPill{background:rgba(22,30,38,.8);border-color:rgba(255,255,255,.14);color:#A9BCCB;}}",
    /* --- 城市选择 popover --- */
    "#weatherPanel{position:fixed;left:22px;bottom:64px;z-index:100;width:280px;max-height:calc(100vh - 120px);",
    "display:flex;flex-direction:column;",
    "background:rgba(255,255,255,.92);backdrop-filter:blur(18px) saturate(1.15);-webkit-backdrop-filter:blur(18px) saturate(1.15);",
    "border:1px solid rgba(255,255,255,.9);border-radius:18px;padding:14px;",
    "box-shadow:0 26px 54px -30px rgba(20,32,43,.42);",
    "font-family:'Space Mono',Consolas,monospace;color:#14202B;",
    "animation:wpIn .18s cubic-bezier(.2,.7,.2,1);}",
    "@keyframes wpIn{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:none;}}",
    "#weatherPanel .wp-title{font-size:11px;font-weight:700;letter-spacing:.08em;color:#4A6173;margin:0 0 8px;}",
    "#weatherPanel .wp-search{display:flex;gap:6px;}",
    "#weatherPanel input{flex:1;min-width:0;border:1px solid rgba(20,32,43,.18);border-radius:10px;",
    "padding:7px 10px;font-family:inherit;font-size:12px;color:#14202B;background:rgba(255,255,255,.8);outline:none;}",
    "#weatherPanel input:focus{border-color:#0D99FF;}",
    "#weatherPanel .wp-go{border:0;border-radius:10px;background:#F0531C;color:#fff;font-family:inherit;",
    "font-size:12px;font-weight:700;padding:0 12px;cursor:pointer;transition:background .15s;}",
    "#weatherPanel .wp-go:hover{background:#D2410E;}",
    "#weatherPanel .wp-go:disabled{opacity:.55;cursor:default;}",
    "#weatherPanel .wp-err{font-size:11px;color:#D2410E;margin:8px 2px 0;}",
    "#weatherPanel .wp-list{margin:10px 0 0;padding:0;list-style:none;overflow:auto;flex:1;min-height:0;}",
    "#weatherPanel .wp-item{border:0;background:none;width:100%;text-align:left;border-radius:10px;",
    "padding:8px 10px;font-family:inherit;font-size:12px;color:#14202B;cursor:pointer;}",
    "#weatherPanel .wp-item:hover{background:rgba(13,153,255,.1);}",
    "#weatherPanel .wp-item .wp-sub{display:block;font-size:10px;color:#60788B;margin-top:2px;}",
    "#weatherPanel .wp-auto{margin-top:10px;width:100%;border:1px solid rgba(20,32,43,.18);border-radius:10px;",
    "background:none;padding:7px 0;font-family:inherit;font-size:11px;font-weight:700;color:#4A6173;cursor:pointer;}",
    "#weatherPanel .wp-auto:hover{border-color:#0D99FF;color:#0D99FF;}",
    "@media (max-width:768px){#weatherPanel{left:14px;bottom:56px;width:calc(100vw - 28px);max-width:280px;}}",
    "@media (prefers-color-scheme:dark){#weatherPanel{background:rgba(22,30,38,.94);border-color:rgba(255,255,255,.14);color:#E6EEF4;}",
    "@media (prefers-reduced-motion:reduce){#weatherPanel{animation:none;}#weatherPill{transition:none;}#weatherPill::after{display:none;}}",
    "#weatherPanel .wp-title{color:#8FA6B8;}",
    "#weatherPanel input{background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.18);color:#E6EEF4;}",
    "#weatherPanel .wp-item{color:#E6EEF4;}",
    "#weatherPanel .wp-item:hover{background:rgba(13,153,255,.18);}",
    "#weatherPanel .wp-auto{border-color:rgba(255,255,255,.2);color:#8FA6B8;}}"
  ].join("");

  function ensureStyle() {
    var old = document.getElementById(STYLE_ID);
    if (old) old.remove();
    var st = document.createElement("style");
    st.id = STYLE_ID;
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  /* ---------- 7. 左下角信息 pill（textContent 构建，无注入面） ---------- */
  function shortIp(ip) {
    if (!ip) return "";
    if (ip.indexOf(":") >= 0) {
      if (ip.length > 20) return "IPv6 " + ip.slice(0, 8) + "…" + ip.slice(-4);
      return ip;
    }
    return ip;
  }

  function addPill(geo, weather) {
    var old = document.getElementById("weatherPill");
    if (old) old.remove();

    var manual = geo.source === SRC_MANUAL;
    var cityName = geo.city || "未知位置";
    var ipTxt = "";
    if (geo.source === SRC_IP) {
      var s = shortIp(geo.ip);
      if (s) ipTxt = " · " + s;
    }

    var pill = document.createElement("div");
    pill.id = "weatherPill";
    pill.title = manual ? "手动定位：" + cityName + "（点击管理）" : "点击手动选择城市";
    pill.setAttribute("role", "button");
    pill.setAttribute("tabindex", "0");
    pill.addEventListener("click", togglePanel);
    pill.addEventListener("keydown", function (e) {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); togglePanel(); }
    });

    var ico = document.createElement("span");
    ico.className = "wp-ico";
    ico.textContent = EMOJI[weather.key] || "☀️";
    pill.appendChild(ico);

    var txt = document.createElement("span");
    txt.className = "wp-txt";
    if (manual) {
      var pin = document.createElement("span");
      pin.className = "wp-pin";
      pin.textContent = "📍";
      txt.appendChild(pin);
    }
    txt.appendChild(document.createTextNode(cityName + " · " + (WEATHER_CN[weather.key] || "晴") +
      (weather.temp != null ? " · " + weather.temp + "°C" : "") + ipTxt));
    pill.appendChild(txt);

    document.body.appendChild(pill);
  }

  /* ---------- 8. 城市选择面板 ---------- */
  function togglePanel() {
    var p = document.getElementById(PANEL_ID);
    if (p) { closePanel(p); return; }
    openPanel();
  }

  function closePanel(panel) {
    if (panel && panel.parentNode) panel.parentNode.removeChild(panel);
  }

  function openPanel() {
    var panel = document.createElement("div");
    panel.id = PANEL_ID;
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-label", "手动选择城市");

    var manual = lsGet();
    var title = document.createElement("p");
    title.className = "wp-title";
    title.textContent = manual ? "手动定位：" + manual.city + "（改选城市或恢复自动）" : "选择城市（手动定位）";
    panel.appendChild(title);

    var row = document.createElement("div");
    row.className = "wp-search";
    var input = document.createElement("input");
    input.type = "text";
    input.placeholder = "输入城市名，如：咸阳";
    input.value = manual ? manual.city : "";
    var go = document.createElement("button");
    go.className = "wp-go";
    go.textContent = "搜索";
    row.appendChild(input);
    row.appendChild(go);
    panel.appendChild(row);

    var err = document.createElement("div");
    err.className = "wp-err";
    err.style.display = "none";
    panel.appendChild(err);

    var list = document.createElement("ul");
    list.className = "wp-list";
    panel.appendChild(list);

    var auto = document.createElement("button");
    auto.className = "wp-auto";
    auto.textContent = "恢复自动定位";
    auto.addEventListener("click", function () {
      lsSet(null);
      closePanel(panel);
      init();
    });
    panel.appendChild(auto);

    var searching = false;
    function doSearch() {
      var q = input.value.trim();
      if (!q || searching) return;
      searching = true;
      err.style.display = "none";
      list.innerHTML = "";
      go.disabled = true;
      go.textContent = "…";
      searchCity(q).then(function (results) {
        results.forEach(function (r) {
          var li = document.createElement("li");
          var b = document.createElement("button");
          b.className = "wp-item";
          var name = document.createTextNode(r.name);
          var sub = document.createElement("span");
          sub.className = "wp-sub";
          sub.textContent = [r.admin, r.country].filter(Boolean).join(" · ") +
            "（" + r.lat.toFixed(2) + ", " + r.lon.toFixed(2) + "）";
          b.appendChild(name);
          b.appendChild(sub);
          b.addEventListener("click", function () {
            var label = r.name + (r.admin && r.admin !== r.name ? "·" + r.admin : "");
            lsSet({ city: label, lat: r.lat, lon: r.lon, ts: Date.now() });
            closePanel(panel);
            init();
          });
          li.appendChild(b);
          list.appendChild(li);
        });
      }).catch(function () {
        err.textContent = "未找到该城市，请检查名称或网络";
        err.style.display = "block";
      }).then(function () {
        searching = false;
        go.disabled = false;
        go.textContent = "搜索";
      });
    }
    go.addEventListener("click", doSearch);
    input.addEventListener("keydown", function (e) { if (e.key === "Enter") doSearch(); });

    document.body.appendChild(panel);
    setTimeout(function () { input.focus(); }, 50);

    // Escape 关闭
    panel.addEventListener("keydown", function (e) {
      if (e.key === "Escape") closePanel(panel);
    });

    // 点击外部关闭（一次性监听，自清理）
    setTimeout(function () {
      document.addEventListener("click", function handler(e) {
        if (!panel.contains(e.target) && e.target.id !== "weatherPill") {
          closePanel(panel);
          document.removeEventListener("click", handler);
        }
      });
    }, 0);
  }

  /* ---------- 9. 启动 ---------- */
  function init() {
    ensureStyle();
    var oldPanel = document.getElementById(PANEL_ID);
    if (oldPanel) closePanel(oldPanel);

    // ① 手动覆盖（用户显式选择，最高优先级）
    var manual = lsGet();
    if (manual && manual.lat != null) {
      getWeather(manual.lat, manual.lon).then(function (w) {
        addPill({ city: manual.city, ip: "", source: SRC_MANUAL }, w);
      });
      return;
    }

    // ② IP 定位（v3.5：已移除本机 Geolocation 请求）
    locate().then(function (geo) {
      return getWeather(geo.lat, geo.lon).then(function (w) { return { geo: geo, weather: w }; });
    }).then(function (res) {
      addPill(res.geo, res.weather);
    }).catch(function () {
      addPill({ city: "未知位置", ip: "", source: SRC_IP }, { key: "clear", temp: null });
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
