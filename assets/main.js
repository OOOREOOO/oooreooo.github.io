/* ============================================================
   main.js — KPD-800T 报告系统 主控制逻辑
   步骤流程 / 事件绑定 / UI 更新
   GitHub Pages 纯静态部署就绪
   ============================================================ */
(function(){
  "use strict";
  var $ = function(id){ return document.getElementById(id); };

  // ===== 等待核心库加载 =====
  function waitFor(fn, max) {
    max = max || 50;
    var n = 0;
    return new Promise(function(resolve, reject){
      var timer = setInterval(function(){
        n++;
        if (fn()) { clearInterval(timer); resolve(); }
        else if (n >= max) { clearInterval(timer); reject(new Error("超时")); }
      }, 100);
    });
  }

  waitFor(function(){ return window.KPDA && window.XLSX; }).then(function(){ init(); }).catch(function(){
    alert("核心脚本加载失败 (KPDA/XLSX)，请刷新页面重试。");
  });

  function init() {
    // ===== 步骤处理器 =====
    var steps = [$("step1"),$("step2"),$("step3"),$("step4")];
    function setStep(idx) {
      steps.forEach(function(el, i){
        el.classList.remove("is-active","is-done");
        if (i < idx) el.classList.add("is-done");
        else if (i === idx) el.classList.add("is-active");
      });
    }

    // ===== 可见性更新 =====
    function updateVisibility() {
      var hasFile = !!KPDA.state.file;
      var hasDefects = KPDA.state.defects && KPDA.state.defects.length > 0;
      $("panelDataPreview").style.display = (hasFile || hasDefects) ? "" : "none";
      $("colRight").style.display = hasDefects ? "" : "none";
    }

    // ===== Toast =====
    function showToast(msg, isErr) {
      var t = $("extractToast");
      t.textContent = msg;
      t.classList.toggle("err", !!isErr);
      t.classList.add("show");
      clearTimeout(t._timer);
      t._timer = setTimeout(function(){ t.classList.remove("show"); }, 4500);
    }

    // ===== 刷新全部 =====
    function refreshDataTable() {
      if ($("panelDataPreview").style.display === "none") return;
      KPDA.renderDataTable(KPDA.state.defectRows, $("dataTable"));
      $("defectsCount").textContent = KPDA.state.defects.length || 0;
    }
    function renderReport() {
      var html = KPDA.buildReportHtml(KPDA.state.info, KPDA.state.defects);
      $("reportDoc").innerHTML = html;
    }
    function refreshAll(skipRender) {
      $("badgeImport").textContent = KPDA.state.defects.length ? "已完成" : (KPDA.state.file ? "已选择" : "待导入");
      $("badgeImport").className = "badge " + (KPDA.state.defects.length ? "ok" : (KPDA.state.file ? "run" : "idle"));
      $("reportPill").textContent = KPDA.state.defects.length ? ("已生成 " + KPDA.state.defects.length + " 处缺陷") : "等待提取";
      $("fileInfo").style.display = KPDA.state.file ? "flex" : "none";
      if (KPDA.state.file) $("fileName").textContent = KPDA.state.file.name + " · " + (KPDA.state.file.size / 1024).toFixed(1) + " KB";
      refreshDataTable();
      if (!skipRender) renderReport();
      updateVisibility();
    }

    // ===== 初始化 =====
    KPDA.state.info = KPDA.infoDefault();
    KPDA.state.defects = [];
    KPDA.state.defectRows = [];
    setStep(0);
    $("colRight").style.display = "none";
    $("panelDataPreview").style.display = "none";

    // ===== 1) 下载模板 =====
    function onDownloadTpl() {
      try {
        KPDA.downloadTemplate();
        showToast("数据模板已开始下载 (.xlsx)", false);
      } catch(e) {
        console.error("[downloadTemplate]", e);
        showToast("模板下载失败：" + (e.message||"未知错误"), true);
      }
    }
    $("btnDownloadTpl").onclick = onDownloadTpl;
    var linkTpl = $("linkDownloadTpl");
    if (linkTpl) linkTpl.addEventListener("click", function(e){ e.preventDefault(); onDownloadTpl(); });

    // ===== 2) 文件上传 =====
    function handleFile(f) {
      if (!f) return;
      var ext = (f.name.split(".").pop()||"").toLowerCase();
      if (ext !== "xlsx" && ext !== "xls") { showToast("请上传 .xlsx 或 .xls 文件", true); return; }
      KPDA.state.file = f;
      $("fileName").textContent = f.name + " · " + (f.size / 1024).toFixed(1) + " KB";
      $("fileInfo").style.display = "flex";
      $("badgeImport").textContent = "已选择";
      $("badgeImport").className = "badge ok";
      setStep(1);
      updateVisibility();
      showToast("已加载：" + f.name + "，请点击「开始提取并生成报告」", false);
    }
    $("fileInput").onchange = function(){ handleFile(this.files[0]); };
    var drop = $("drop");
    ["dragenter","dragover"].forEach(function(ev){ drop.addEventListener(ev, function(e){ e.preventDefault(); drop.classList.add("dragover"); }); });
    ["dragleave","drop"].forEach(function(ev){ drop.addEventListener(ev, function(e){ e.preventDefault(); drop.classList.remove("dragover"); }); });
    drop.addEventListener("drop", function(e){ var f = e.dataTransfer.files[0]; if (f) handleFile(f); });

    $("btnClearFile").onclick = function(){
      KPDA.state.file = null;
      $("fileInfo").style.display = "none";
      $("fileInput").value = "";
      setStep(0);
      updateVisibility();
    };

    // ===== 3) 提取数据 =====
    $("btnExtract").onclick = function(){
      if (!KPDA.state.file) { showToast("请先上传 Excel 文件", true); return; }
      $("badgeImport").textContent = "处理中...";
      $("badgeImport").className = "badge run";
      $("extractProgress").classList.add("show");
      showToast("正在解析 Excel 并映射报告字段...", false);
      setStep(2);
      KPDA.parseFile(KPDA.state.file).then(function(parsed){
        KPDA.state.info = parsed.info;
        KPDA.state.defects = parsed.defects;
        KPDA.state.defectRows = parsed.defectRows;
        KPDA.saveReport();
        setStep(3);
        updateVisibility();
        refreshAll();
        showToast("成功解析 " + parsed.defects.length + " 条缺陷记录", false);
      }).catch(function(e){
        console.error("[parseFile]", e);
        showToast("解析失败：" + (e.message||"未知错误"), true);
      }).finally(function(){
        $("badgeImport").textContent = "已完成";
        $("badgeImport").className = "badge ok";
        $("extractProgress").classList.remove("show");
      });
    };

    // ===== 4) 下载 Word =====
    $("btnDownloadWord").onclick = function(){
      var btn = $("btnDownloadWord"), old = btn.innerHTML;
      btn.innerHTML = '<span class="loading"></span><span>生成中...</span>';
      btn.disabled = true;
      KPDA.downloadWord("架空线双端局放检测报告.docx").then(function(){
        showToast("Word 报告已下载", false);
      }).catch(function(e){
        console.error("[downloadWord]", e);
        showToast("Word 生成失败：" + (e.message||"未知错误"), true);
      }).finally(function(){
        btn.innerHTML = old;
        btn.disabled = false;
      });
    };

    // ===== 5) 打印 PDF =====
    $("btnPrintPdf").onclick = function(){
      var style = '<style>body{font-family:"Microsoft YaHei","PingFang SC",serif;color:#1e293b;padding:24px;}.report-cover{text-align:center;padding:32px 24px;border-bottom:2px solid #1d4ed8;margin-bottom:28px;}.report-cover .banner{display:inline-block;background:#1d4ed8;color:#fff;font-size:22px;font-weight:700;letter-spacing:.6em;padding:14px 28px;border-radius:6px;margin:14px 0;}.report-cover .org{display:flex;justify-content:space-between;font-size:13px;color:#475569;margin-top:12px;}.report-h{font-size:16px;font-weight:700;color:#1e3a8a;margin:22px 0 12px;}.report-h2{font-size:14px;font-weight:700;color:#1e3a8a;margin:14px 0 8px;}.report-toc{border:1px solid #cbd5e1;border-radius:6px;padding:14px 18px;margin:12px 0 24px;}.report-toc .toc-line{display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px dashed #cbd5e1;}.report-toc .toc-num{width:80px;color:#1e3a8a;font-weight:600;}table.report-table{width:100%;border-collapse:collapse;margin:12px 0;font-size:12.5px;}table.report-table th,table.report-table td{border:1px solid #94a3b8;padding:6px 10px;}table.report-table th{background:#eff6ff;font-weight:600;color:#1e3a8a;width:22%;}table.report-table td{background:#fff;text-align:left;}.report-blue-bar{background:#1d4ed8;color:#fff;font-weight:700;text-align:center;padding:10px;margin:18px 0 12px;border-radius:4px;}.report-blue-bar.sub{font-size:13px;padding:8px;}.report-img-ph{display:flex;flex-direction:column;align-items:center;justify-content:center;background:repeating-linear-gradient(45deg,#eff6ff,#eff6ff 8px,#dbeafe 8px,#dbeafe 16px);border:1px dashed #93c5fd;color:#2563eb;border-radius:6px;height:180px;margin:10px 0;padding:12px;}.report-para{margin:10px 0;text-indent:2em;}.report-end{margin-top:28px;padding-top:14px;border-top:1px solid #cbd5e1;font-size:11.5px;color:#64748b;}</style>';
      var html = '<html><head><meta charset="utf-8"><title>检测报告</title>'+style+'</head><body>'+KPDA.buildReportHtml(KPDA.state.info, KPDA.state.defects)+'</body></html>';
      var w = window.open("","_blank");
      w.document.write(html);
      w.document.close();
      setTimeout(function(){ w.print(); }, 500);
    };

    // ===== 6) 示例 / 编辑 / 重置 =====
    $("btnLoadDemo").onclick = function(){
      KPDA.initDemo();
      setStep(3);
      updateVisibility();
      refreshAll();
      showToast("已加载示例数据（2 条缺陷记录）", false);
    };
    $("btnEditReport").onclick = function(){
      KPDA.saveReport();
      window.location.href = "edit.html";
    };
    $("btnReset").onclick = function(){
      if (!confirm("确认重置当前所有数据？")) return;
      KPDA.clearReport();
      KPDA.state.file = null;
      $("fileInfo").style.display = "none";
      $("fileInput").value = "";
      $("badgeImport").textContent = "待导入";
      $("badgeImport").className = "badge idle";
      $("colRight").style.display = "none";
      $("panelDataPreview").style.display = "none";
      setStep(0);
      $("reportDoc").innerHTML = "";
    };
    $("btnResetReport").onclick = function(){
      if (!confirm("重置报告为默认示例数据？")) return;
      KPDA.initDemo();
      setStep(3);
      updateVisibility();
      refreshAll();
    };

    window.__KPDA_refresh = refreshAll;
  }
})();
