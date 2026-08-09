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
        if (step <= currentStep || step === currentStep + 1) showPanel(step);
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
        a.download = "架空线双端局放检测报告.docx";
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

    // 打印 PDF
    $("btnPrintPdf").onclick = function() {
      var w = window.open("", "_blank");
      var html = KPDA.buildReportHtml(KPDA.state.info, KPDA.state.defects);
      w.document.write("<html><head><meta charset='utf-8'><title>报告</title><style>");
      w.document.write("body{font-family:'Microsoft YaHei',sans-serif;padding:40px;color:#1b1b1b;line-height:1.8;}");
      w.document.write(".report-cover{text-align:center;border-bottom:2px solid #1A5276;margin-bottom:30px;padding-bottom:20px;}");
      w.document.write(".report-cover .banner{font-size:48px;font-weight:700;color:#fff;background:#10407C;padding:20px;margin:30px 0;}");
      w.document.write(".report-h{font-size:16px;font-weight:700;color:#1A5276;border-bottom:2px solid #1A5276;margin:20px 0 10px;}");
      w.document.write("</style></head><body>" + html + "</body></html>");
      w.document.close();
      setTimeout(function() { w.print(); }, 500);
    };

    // 刷新页面（验证是否最新版本）
    // 强制刷新：每次加载都确保 index.html 是最新（绕过浏览器对主文档的缓存）
    (function forceFresh() {
      try {
        var u = new URL(location.href);
        if (!u.searchParams.has("_t")) {
          u.searchParams.set("_t", String(Date.now()));
          location.replace(u.toString());
        }
      } catch (e) {}
    })();

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
