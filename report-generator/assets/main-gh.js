/* ==========================================================================
   main-gh.js — GitHub Pages 静态部署 (无后端依赖)
   ========================================================================== */
(function () {
  "use strict";
  var $ = function(id) { return document.getElementById(id); };

  function init() {
    var stepNodes = document.querySelectorAll(".nav-item");
    var panels = document.querySelectorAll(".panel");
    var currentStep = 0;

    function showPanel(idx) {
      currentStep = idx;
      panels.forEach(function(p, i) { p.classList.toggle("active", i === idx); });
      stepNodes.forEach(function(el, i) {
        el.classList.remove("active", "done");
        if (i < idx) el.classList.add("done");
        else if (i === idx) el.classList.add("active");
      });
    }

    stepNodes.forEach(function(node) {
      node.addEventListener("click", function() {
        var step = parseInt(node.getAttribute("data-step"));
        if (step <= currentStep || step === currentStep + 1) { showPanel(step); } else { toast("请先完成上一步", true); }
      });
    });

    function toast(msg, isErr) {
      var t = $("toast");
      t.textContent = msg;
      t.setAttribute("role", "status");
      t.setAttribute("aria-live", "polite");
      t.className = "toast show" + (isErr ? " err" : "");
      clearTimeout(t._timer);
      t._timer = setTimeout(function() { t.className = "toast"; }, 4000);
    }

    function refreshUI() {
      var hasFile = !!KPDA.state.file;
      var hasDefects = KPDA.state.defects && KPDA.state.defects.length > 0;
      var count = hasDefects ? KPDA.state.defects.length : 0;
      $("fileInfo").style.display = hasFile ? "flex" : "none";
      if (hasFile) $("fileName").textContent = KPDA.state.file.name;
      $("btnExtract").disabled = !hasFile;
      $("exportCount").textContent = count;
      var dct = $("defectCountTitle");
      if (dct) dct.textContent = count;
      var side = $("sideGenCount");
      if (side) side.textContent = count;
      if (hasDefects) {
        KPDA.renderDefectCards(KPDA.state.defects, $("defectCards"));
        $("reportViewer").innerHTML = KPDA.buildReportHtml(KPDA.state.info, KPDA.state.defects);
      }
    }

    // 初始化
    KPDA.state.info = Object.assign({}, KPDA.DEFAULT_INFO);
    KPDA.state.defects = [];
    showPanel(0);

    // 下载模板
    $("btnDownloadTpl").onclick = function() {
      try { KPDA.downloadTemplate(); toast("数据模板已下载"); }
      catch(e) { toast("下载失败: " + e.message, true); }
    };

    // 文件上传
    var fileInput = $("fileInput");
    var dropZone = $("dropZone");

    function handleFile(f) {
      if (!f) return;
      var ext = (f.name.split(".").pop() || "").toLowerCase();
      if (ext !== "xlsx" && ext !== "xls" && ext !== "xlsm") { toast("请上传 .xlsx / .xls / .xlsm", true); return; }
      KPDA.state.file = f;
      $("fileName").textContent = f.name;
      $("fileInfo").style.display = "flex";
      $("btnExtract").disabled = false;
      toast("已加载: " + f.name);
    }

    fileInput.onchange = function() { handleFile(this.files[0]); };
    dropZone.onclick = function() { fileInput.click(); };
    // 键盘可达性：Enter/Space 触发文件选择
    dropZone.addEventListener("keydown", function(e) {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
    });
    dropZone.addEventListener("dragover", function(e) { e.preventDefault(); dropZone.classList.add("dragover"); });
    dropZone.addEventListener("dragleave", function(e) { e.preventDefault(); dropZone.classList.remove("dragover"); });
    dropZone.addEventListener("drop", function(e) { e.preventDefault(); dropZone.classList.remove("dragover"); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });

    $("btnClearFile").onclick = function() {
      KPDA.state.file = null;
      $("fileInfo").style.display = "none";
      fileInput.value = "";
      $("btnExtract").disabled = true;
      showPanel(0);
    };

    // 提取数据（客户端SheetJS）
    $("btnExtract").onclick = function() {
      if (!KPDA.state.file) { toast("请先上传 Excel", true); return; }
      var btn = $("btnExtract");
      var old = btn.innerHTML;
      btn.innerHTML = "解析中...";
      btn.disabled = true;
      KPDA.parseExcel(KPDA.state.file).then(function(parsed) {
        KPDA.state.info = parsed.info;
        KPDA.state.defects = parsed.defects;
        showPanel(1);
        refreshUI();
        toast("解析 " + parsed.defects.length + " 条记录");
      }).catch(function(e) {
        console.error(e);
        toast("解析失败: " + (e.message || e), true);
      }).finally(function() {
        btn.innerHTML = old;
        btn.disabled = false;
      });
    };

    // 加载示例
    $("btnLoadDemo").onclick = function() {
      var data = KPDA.getExampleData();
      KPDA.state.info = data.info;
      KPDA.state.defects = data.defects;
      showPanel(1);
      refreshUI();
      toast("已加载示例数据（" + data.defects.length + " 条）");
    };

    // 导航
    $("btnNext1").onclick = function() { showPanel(2); };
    $("btnNext2").onclick = function() { showPanel(3); };
    $("btnBack1").onclick = function() { showPanel(0); };
    $("btnBack2").onclick = function() { showPanel(1); };
    $("btnBack3").onclick = function() { showPanel(2); };

    // 下载Word（客户端生成）
    $("btnDownloadWord").onclick = function() {
      var btn = $("btnDownloadWord");
      var old = btn.innerHTML;
      btn.innerHTML = "生成中...";
      btn.disabled = true;
      if (!window.KPDADocx || !window.KPDADocx.generateBlob) {
        toast("Word生成模块未加载", true);
        btn.innerHTML = old; btn.disabled = false;
        return;
      }
      KPDADocx.generateBlob(KPDA.state.info, KPDA.state.defects).then(function(blob) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        // 导出文件名：标题-副标题（如「南网云南曲靖沾益大坡供电所架空线双端局放检测报告-10kV架空线双端局放定位仪巡检」）
        var _info = KPDA.state.info || {};
        var _t = String(_info["标题"] || "").trim();
        var _s = String(_info["副标题"] || "").trim();
        var _fn = [_t, _s].filter(Boolean).join("-").replace(/[\\\/:*?"<>|]/g, "_").trim();
        a.download = (_fn || "架空线双端局放检测报告") + ".docx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
        toast("Word 已下载");
      }).catch(function(e) {
        toast("生成失败: " + (e.message || e), true);
      }).finally(function() {
        btn.innerHTML = old;
        btn.disabled = false;
      });
    };

        // 打印 PDF（v71：复用页面完整样式 + 弹窗拦截兜底）
    $("btnPrintPdf").onclick = function() {
      var w = window.open("", "_blank");
      if (!w) { toast("请允许弹出窗口后重试", true); return; }
      var html = KPDA.buildReportHtml(KPDA.state.info, KPDA.state.defects);
      var cssHref = "";
      var lk = document.querySelector('link[href*="style.css"]');
      if (lk) cssHref = lk.href;
      w.document.write("<!DOCTYPE html><html><head><meta charset='utf-8'><title>检测报告</title>");
      if (cssHref) w.document.write("<link rel='stylesheet' href='" + cssHref + "'>");
      w.document.write("<style>body{padding:40px;color:#1b1b1b;line-height:1.8;background:#fff;}");
      w.document.write(".report-img img{max-width:100%;height:auto;}.report-img-row{display:flex;gap:20px;}.report-img-half{flex:1;min-width:0;}.report-img,.report-img-ph{page-break-inside:avoid;}");
      w.document.write("@media print{body{padding:0;}}</style></head><body>");
      w.document.write(html + "</body></html>");
      w.document.close();
      setTimeout(function() { w.print(); }, 500);
    };
    // v71：移除 forceFresh 自动跳转（每次访问双加载 2.85MB 的根因）；保留手动刷新按钮

    $("btnRefresh").onclick = function() {
      // 硬刷新：加时间戳绕过缓存
      var u = new URL(location.href);
      u.searchParams.set("_t", String(Date.now()));
      location.replace(u.toString());
    };

    // 重置
    $("btnReset").onclick = function() {
      if (!confirm("确认重置？")) return;
      KPDA.state.file = null;
      KPDA.state.info = Object.assign({}, KPDA.DEFAULT_INFO);
      KPDA.state.defects = [];
      $("fileInfo").style.display = "none"; fileInput.value = "";
      $("reportViewer").innerHTML = "";
      $("defectCards").innerHTML = '<div class="empty-state">暂无数据，请先导入 Excel 或加载示例</div>';
      $("exportCount").textContent = "0";
      var dct = $("defectCountTitle");
      if (dct) dct.textContent = "0";
      showPanel(0);
      $("btnExtract").disabled = true;
      toast("已重置");
    };
  }

  if (window.KPDA && window.XLSX) { init(); }
  else { alert("脚本加载失败，请刷新。"); }
})();
