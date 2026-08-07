/* ==========================================================================
   架空线局放 KPD-800T 报告生成系统 — 共享逻辑
   - 对齐 make_excel_template.py / fill_word_template.py / report_builder.py
   - Excel 28列 + 报告基本信息 4字段
   - 客户端 docx 下载 (docx-gen.js)
   ========================================================================== */
(function () {
  "use strict";

  // ============================================================
  // 1. 常量（对齐 make_excel_template.py 28列）
  // ============================================================
  const HEADERS = [
    "序号", "客户单位", "检测单位", "线路名称", "检测仪器", "检测时间", "检测人员",
    "天气", "温湿度",
    "前端设备悬挂杆号", "前端A相幅值(mV)", "前端B相幅值(mV)", "前端C相幅值(mV)",
    "后端设备悬挂杆号", "后端A相幅值(mV)", "后端B相幅值(mV)", "后端C相幅值(mV)",
    "定位分析描述", "结论", "检测结果分析", "疑似杆情况说明", "处理意见",
    "验证描述（含绝缘电阻值MΩ）",
    "最大相",
    "前端实时图谱图片路径", "后端实时图谱图片路径", "双端定位谱图图片路径", "现场疑似缺陷图片路径",
  ];

  // COL_MAP: Excel列名 → 内部字段名（对齐 fill_word_template.py）
  const COL_MAP = {
    "检测时间": "detect_time",
    "检测人员": "detect_person",
    "温湿度": "temp_humidity",
    "前端设备悬挂杆号": "ft_tower",
    "前端A相幅值(mV)": "ftA",
    "前端B相幅值(mV)": "ftB",
    "前端C相幅值(mV)": "ftC",
    "后端设备悬挂杆号": "bk_tower",
    "后端A相幅值(mV)": "bkA",
    "后端B相幅值(mV)": "bkB",
    "后端C相幅值(mV)": "bkC",
    "定位分析描述": "position_analysis",
    "结论": "conclusion",
    "检测结果分析": "result_analysis",
    "疑似杆情况说明": "pole_condition",
    "处理意见": "advice",
    "验证描述（含绝缘电阻值MΩ）": "verification",
    "最大相": "max_phase",
    "前端实时图谱图片路径": "frontend_img",
    "后端实时图谱图片路径": "backend_img",
    "双端定位谱图图片路径": "position_img",
    "现场疑似缺陷图片路径": "defect_img",
  };

  const BASIC_INFO = ["标题", "副标题", "出具单位", "报告日期"];
  // 内部字段名
  const INFO_MAP = { "标题": "report_title", "副标题": "subtitle", "出具单位": "company", "报告日期": "date" };

  const NOTES = [
    "【填写说明】",
    "",
    "1. 报告基本信息：每份报告填写一次，决定报告封面和页脚的信息。",
    "",
    "2. 缺陷检测记录：",
    "   - 每行代表一个检测到的缺陷点位",
    "   - 示例数据来源于 2026.8.4 会泽待补供电所现场检测，供参考",
    "   - 删除示例行后填入实际数据（在示例下方继续添加新行即可）",
    "   - 最大相：哪个相位幅值最大（填 A / B / C）",
    "",
    "3. 关键字段说明：",
    "   - 检测结果分析：完整描述文字，脚本会按段落拆分",
    "   - 验证描述（含绝缘电阻值MΩ）：验证环节的完整描述，需包含实测绝缘电阻值",
    "   - 定位分析描述：填入定位谱图分析文字",
    "   - 前端/后端 A、B、C 相幅值(mV)：现场实测三相幅值",
    "",
    "4. 图片路径列（可选）：",
    "   - 填入本地图片绝对路径，留空 = 显示占位符",
    "   - 支持 jpg/jpeg/png/bmp/gif",
    "",
    "5. 填写完成后，上传此 Excel 即可自动生成 DOCX 报告。",
    "",
    "6. 列顺序严格固定，请勿增删或调换列，以保证自动填充准确。",
  ];

  // ============================================================
  // 2. 工具函数
  // ============================================================
  const SKEY = "kpd800_report_v2";

  function escapeHtml(s) {
    if (s == null) return "";
    return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function hasVal(v) { return v !== undefined && v !== null && String(v).trim() !== ""; }
  function num(v) { const x = Number(v); return Number.isFinite(x) ? x : null; }
  function g(v, fallback) { return hasVal(v) ? v : (fallback || ""); }
  function splitSentences(s) {
    if (!s) return [];
    return String(s).split(/[。！？；\.\!\?\;]/).filter(Boolean);
  }

  function downloadBlob(blob, filename) {
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // ============================================================
  // 3. 状态
  // ============================================================
  const state = { info: {}, defects: [], defectRows: [], file: null };

  function infoDefault() {
    return {
      标题: "南网云南会泽待补供电所架空线双端局放检测报告",
      副标题: "10kV架空线双端局放定位仪巡检",
      出具单位: "咸亨国际科技股份有限公司",
      报告日期: "二〇二六年八月",
    };
  }

  // ============================================================
  // 4. 示例数据（对齐 make_excel_template.py EXAMPLES）
  // ============================================================
  function makeExamples() {
    const e1 = {};
    e1["序号"] = 1;
    e1["客户单位"] = "南网云南会泽待补供电所";
    e1["检测单位"] = "咸亨国际科技股份有限公司";
    e1["线路名称"] = "10kV野马线";
    e1["检测仪器"] = "架空线双端局放定位仪";
    e1["检测时间"] = "2026.8.4";
    e1["检测人员"] = "刘奥";
    e1["天气"] = "晴";
    e1["温湿度"] = "26℃ / 未知";
    e1["前端设备悬挂杆号"] = "10kV野马线T车路箐支线#16.5";
    e1["前端A相幅值(mV)"] = 7148;
    e1["前端B相幅值(mV)"] = 6198;
    e1["前端C相幅值(mV)"] = 7460;
    e1["后端设备悬挂杆号"] = "10kV野马线T车路箐支线#16.10";
    e1["后端A相幅值(mV)"] = 6148;
    e1["后端B相幅值(mV)"] = 6198;
    e1["后端C相幅值(mV)"] = 7510;
    e1["定位分析描述"] = "根据定位谱图，故障位置距离#2设备（10号杆）约56m，初步判断为#9号杆。";
    e1["结论"] = "10kV野马线T车路箐支线#16.9 C相绝缘子存在重度局部放电隐患";
    e1["检测结果分析"] = "根据现场测试数据可知，C相数值相比其他两相显著偏高且放电特征一致（其余两相信号疑似为感应耦合的高频信号），且图谱符合局部放电特性；10kV野马线#16.5断开后端供电后，发现疑似局放信号骤降95%左右，因此判断疑似局放来源后端杆塔；后续在10kV野马线#16.5与#16.10进行双端定位时，发现疑似局放信号来源于 #16.9，信号表明此处位置处数值幅值最大且图谱特征极为明显。";
    e1["疑似杆情况说明"] = "到现场杆塔下人耳可听见明显放电声，肉眼观察无明显异常的现象，后续用其他远程拍摄设备发现绝缘子内部存在疑似穿孔、绝缘子捆扎线存在明显放电痕迹。";
    e1["处理意见"] = "放电强度较强，人耳可听明显放电声，现场已确定疑似故障位置，需及时符合验证，并进行缺陷故障处理（停电更换或带电左右更换）。";
    e1["验证描述（含绝缘电阻值MΩ）"] = "经带电作业（或停电检修）流程，对设备定位的故障缺陷位置实施了精准消缺处理。现场拆解疑似缺陷绝缘子，可见明显穿孔击穿痕迹。采用兆欧表对该绝缘子进行绝缘电阻检测，实测绝缘电阻值为5MΩ。";
    e1["最大相"] = "C";
    e1["前端实时图谱图片路径"] = "";
    e1["后端实时图谱图片路径"] = "";
    e1["双端定位谱图图片路径"] = "";
    e1["现场疑似缺陷图片路径"] = "";

    const e2 = {};
    e2["序号"] = 2;
    e2["客户单位"] = "南网云南会泽待补供电所";
    e2["检测单位"] = "咸亨国际科技股份有限公司";
    e2["线路名称"] = "10kV野马线";
    e2["检测仪器"] = "架空线双端局放定位仪";
    e2["检测时间"] = "2026.8.4";
    e2["检测人员"] = "刘奥";
    e2["天气"] = "晴";
    e2["温湿度"] = "26℃ / 未知";
    e2["前端设备悬挂杆号"] = "10kV野马线#16.29";
    e2["前端A相幅值(mV)"] = 748;
    e2["前端B相幅值(mV)"] = 698;
    e2["前端C相幅值(mV)"] = 803;
    e2["后端设备悬挂杆号"] = "10kV野马线#16.25";
    e2["后端A相幅值(mV)"] = 1148;
    e2["后端B相幅值(mV)"] = 1020;
    e2["后端C相幅值(mV)"] = 1510;
    e2["定位分析描述"] = "根据定位谱图，故障位置距离#1设备（3号杆）约8m，初步判断为#16.29。";
    e2["结论"] = "10kV野马线T车路箐支线#16.29 C相绝缘子存在局部放电隐患";
    e2["检测结果分析"] = "根据现场测试数据可知，C相数值相比其他两相显著偏高且放电特征一致（其余两相信号疑似为感应耦合的高频信号），且图谱符合局部放电特性；10kV野马线#16.29，在电源侧与负载侧分别挂载时，其电源侧信号略大于负载侧，因而在 #16.25（小号侧，电缆供电杆塔）进行定位；后续在10kV野马线#16.25与#16.29进行双端定位时，发现疑似局放信号来源于 #16.29。";
    e2["疑似杆情况说明"] = "人耳未听见放电声，肉眼观察无明显异常的现象，待进一步跟进分析，初步判断为绝缘子内部缺陷，后续用其他远程拍摄设备发现绝缘子内部存在疑似放电产生的烧蚀痕迹。";
    e2["处理意见"] = "放电强度轻中度，人耳不可听见明显放电声，针对疑似故障位置进行多方位确认，条件允许时可考虑进行更换。";
    e2["验证描述（含绝缘电阻值MΩ）"] = "经带电作业（或停电检修）流程，对设备定位的故障缺陷位置实施了精准消缺处理。现场拆解疑似缺陷绝缘子，可见明显穿孔击穿痕迹。采用兆欧表对该绝缘子进行绝缘电阻检测，实测绝缘电阻值为10MΩ。";
    e2["最大相"] = "C";
    e2["前端实时图谱图片路径"] = "";
    e2["后端实时图谱图片路径"] = "";
    e2["双端定位谱图图片路径"] = "";
    e2["现场疑似缺陷图片路径"] = "";

    const result = [e1, e2].map((d, i) => {
      const obj = { idx: i + 1, display: Object.assign({}, d) };
      HEADERS.forEach(h => {
        if (d[h] != null) obj[h] = d[h];
        if (COL_MAP[h]) obj[COL_MAP[h]] = d[h];
      });
      return obj;
    });
    return result;
  }

  // ============================================================
  // 5. 下载 Excel 模板（对齐 make_excel_template.py）
  // ============================================================
  function downloadTemplate() {
    if (!window.XLSX || !window.XLSX.utils) {
      alert("XLSX 库未加载，请稍候重试或刷新页面。");
      return;
    }
    try {
      const wb = XLSX.utils.book_new();

      // --- Sheet 1: 报告基本信息 ---
      const infoRows = BASIC_INFO.map(k => [k, infoDefault()[k] || ""]);
      const ws1 = XLSX.utils.aoa_to_sheet([["报告基本信息（每份报告填写一次）"], ["", ""], ...infoRows]);
      ws1["!cols"] = [{ wch: 14 }, { wch: 52 }];
      ws1["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
      XLSX.utils.book_append_sheet(wb, ws1, "报告基本信息");

      // --- Sheet 2: 缺陷检测记录（28列） ---
      const examples = makeExamples();
      const exampleRows = examples.map((d, i) =>
        HEADERS.map((h, ci) => (ci === 0 ? (i + 1) : (d[h] != null ? d[h] : "")))
      );
      const ws2 = XLSX.utils.aoa_to_sheet([
        ["缺陷检测记录（每行 = 一个缺陷点位，请逐行填写，删除示例后填实数据）"],
        HEADERS,
        ...exampleRows,
        HEADERS.map(() => ""),
        HEADERS.map(() => ""),
        HEADERS.map(() => ""),
      ]);
      // 合并第一行标题
      ws2["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: HEADERS.length - 1 } }];
      // 美化表头
      for (let c = 0; c < HEADERS.length; c++) {
        const addr = XLSX.utils.encode_cell({ r: 1, c });
        if (ws2[addr] && typeof ws2[addr] === "object") {
          ws2[addr].s = { fill: { fgColor: { rgb: "1A5276" } }, font: { color: { rgb: "FFFFFF" }, bold: true }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
        }
      }
      ws2["!cols"] = HEADERS.map(() => ({ wch: 14 }));
      XLSX.utils.book_append_sheet(wb, ws2, "缺陷检测记录");

      // --- Sheet 3: 填写说明 ---
      const ws3 = XLSX.utils.aoa_to_sheet([NOTES.map(n => [n])].flat());
      ws3["!cols"] = [{ wch: 90 }];
      XLSX.utils.book_append_sheet(wb, ws3, "填写说明");

      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      downloadBlob(blob, "架空线双端局放检测_数据填写模板.xlsx");
    } catch (e) {
      console.error("[downloadTemplate]", e);
      alert("模板生成失败：" + (e && e.message ? e.message : e));
    }
  }

  // ============================================================
  // 6. Excel 解析
  // ============================================================
  function readWorkbook(wb) {
    const info = Object.assign(infoDefault(), {});
    const defects = [];
    const defectRows = [];

    const wsi = wb.Sheets["报告基本信息"];
    if (wsi) {
      const arr = XLSX.utils.sheet_to_json(wsi, { header: 1 });
      arr.forEach(r => {
        const k = (r[0] || "").toString().trim();
        if (BASIC_INFO.includes(k) && hasVal(r[1])) info[k] = String(r[1]).trim();
      });
    }

    const wsd = wb.Sheets["缺陷检测记录"];
    if (wsd) {
      const arr = XLSX.utils.sheet_to_json(wsd, { header: 1 });
      // 第1行是标题，第2行是表头，数据从第3行开始
      if (arr.length >= 3) {
        const headers = arr[1];
        for (let r = 2; r < arr.length; r++) {
          const row = arr[r];
          if (!row || row.every(c => !hasVal(c))) continue;
          const obj = { idx: defects.length + 1, display: {} };
          headers.forEach((h, i) => {
            const v = row[i];
            obj.display[h] = v;
            obj[h] = v;
            if (COL_MAP[h]) obj[COL_MAP[h]] = v;
          });
          defects.push(obj);
          defectRows.push(obj.display);
        }
      }
    }

    return { info, defects, defectRows };
  }

  function parseFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          resolve(readWorkbook(wb));
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // ============================================================
  // 7. 报告 HTML 渲染（对齐 report_builder.py 结构）
  // ============================================================
  function kvTable(rows) {
    const trs = rows.map(r => `<tr><th>${escapeHtml(r[0])}</th><td>${escapeHtml(r[1] || "")}</td></tr>`).join("");
    return `<table class="report-table"><tbody>${trs}</tbody></table>`;
  }

  function imgPH(label, hint) {
    return `<div class="report-img-ph">${escapeHtml(label)}${hint ? `<div class="tag">${escapeHtml(hint)}</div>` : ""}</div>`;
  }

  function reportImg(label, fileHint) {
    if (label && /^(assets\/|[a-zA-Z]:[\\/])/.test(label)) {
      return `<img src="${escapeHtml(label)}" alt="${escapeHtml(fileHint||'')}" style="max-width:100%;border-radius:6px;border:1px solid #cbd5e1;margin:10px 0;" />`;
    }
    return imgPH(label || "图", fileHint);
  }

  function buildReportHtml(info, defects) {
    const d = info || infoDefault();
    const list = (defects && defects.length) ? defects : [];

    const cover = `
      <section id="sec-cover" class="report-cover">
        <div class="sub">${escapeHtml(g(d.副标题, ""))}</div>
        <div class="banner">检 测 报 告</div>
        <div class="org">
          <span>${escapeHtml(g(d.出具单位, ""))}</span>
          <span>${escapeHtml(g(d.报告日期, ""))}</span>
        </div>
        <div style="font-size:20px;font-weight:700;color:#1e3a8a;margin-top:18px;">
          ${escapeHtml(g(d.标题, "架空线双端局放检测报告"))}
        </div>
      </section>`;

    const toc = `
      <section id="sec-toc" class="report-toc">
        <div class="toc-line"><span class="toc-num">一、</span><span>检测目的</span><span>···</span></div>
        <div class="toc-line"><span class="toc-num">二、</span><span>架空线双端局放定位检测原理</span><span>···</span></div>
        <div class="toc-line"><span class="toc-num">三、</span><span>本次使用设备</span><span>···</span></div>
        <div class="toc-line"><span class="toc-num">四、</span><span>缺陷综合检测报告</span><span>···</span></div>
        ${list.map((_, i) => `<div class="toc-line"><span></span><span>${i+1}、${escapeHtml(g(list[i].结论, '缺陷 '+(i+1))).slice(0,40)}</span><span>···</span></div>`).join("")}
      </section>`;

    const secPurpose = `
      <section id="sec-purpose">
        <h2 class="report-h">一、检测目的</h2>
        <p class="report-para">针对 ${escapeHtml(g(d.标题, "本次架空线"))}，采用 KPD-800T 架空线双端局放定位系统开展现场带电检测，识别并定位线路潜在局放缺陷，评估缺陷类别与严重等级，为线路检修与运行维护提供技术依据。</p>
      </section>`;

    const secPrinciple = `
      <section id="sec-principle">
        <h2 class="report-h">二、架空线双端局放定位检测原理</h2>
        <p class="report-para">KPD-800T 架空线双端局放定位系统基于双端行波时差法，通过在线路两端各布置一台检测装置（前端、后端），同步采集线路上的局部放电信号；利用两侧装置接收到同一放电信号的时间差，结合信号传播速度与线路长度，计算放电源距端点的精确位置。系统支持连续、宽频带、抗电磁干扰的高速采样，可识别表面放电、电晕放电、内部放电等多种类型。</p>
        ${reportImg("图  双端局放检测原理示意图", "可嵌入原理图")}
      </section>`;

    const secEquip = `
      <section id="sec-equipment">
        <h2 class="report-h">三、本次使用设备</h2>
        <div class="report-blue-bar">KPD-800T 双端局放定位仪</div>
        ${reportImg("图  KPD-800T 主机外观", "可嵌入产品图")}
        <div class="report-blue-bar sub">主要特性</div>
        <ol>
          <li>支持中压架空线局部放电带电检测：设备专为中压架空线路设计，可在不停电状态下实时监测局部放电情况，提前发现绝缘缺陷，避免停电检测带来的供电中断。</li>
          <li>支持中压架空线局部放电精准定位：采用双端检测原理，通过捕捉放电信号的传播时间差，实现对放电点的精准定位，精度可达0.5米或0.5%线缆长度。</li>
          <li>自动生成立位谱图，直接给出放电位置及强度：检测过程中自动生成立位谱图，直观显示放电位置和强度，便于快速分析和决策。</li>
          <li>采用电池供电，使用方便：设备采用电池供电，无需外部电源，适合户外或无电源环境。内置大容量电池可支持长时间连续工作，续航能力达3小时以上。</li>
        </ol>
      </section>`;

    const blocks = list.map((defect, i) => defectBlockHtml(i + 1, defect)).join("");
    const secDefects = `
      <section id="sec-defects">
        <h2 class="report-h">四、缺陷综合检测报告</h2>
        <p class="report-para">本次检测共发现 ${list.length} 处疑似缺陷，分类与定位信息详述如下：</p>
      </section>${blocks}`;

    const end = `
      <section id="sec-end">
        <div class="report-end">
          注：因无近距离及解体分析，结论为初步判断，请以最终解体或近距离确诊为准，合理检修。
          <br /><br />依据 DL/T 596—2021《电力设备预防性试验规程》关于10kV架空线路针式绝缘子绝缘电阻不应低于300MΩ的规定，判定该绝缘子已严重劣化，属零值（低值）绝缘子，绝缘性能已基本失效。
        </div>
      </section>`;

    return cover + toc + secPurpose + secPrinciple + secEquip + secDefects + end;
  }

  function defectBlockHtml(idx, defect) {
    const d = defect || {};
    const phase3 = ["A相", "B相", "C相"];
    const ftKeys = ["前端A相幅值(mV)", "前端B相幅值(mV)", "前端C相幅值(mV)"];
    const bkKeys = ["后端A相幅值(mV)", "后端B相幅值(mV)", "后端C相幅值(mV)"];

    const ftVals = ftKeys.map(k => num(d[k]));
    const bkVal = bkKeys.map(k => num(d[k]));

    // 计算最大相
    const allVals = [...ftVals, ...bkVal].filter(v => v != null);
    const maxVal = allVals.length ? Math.max(...allVals) : null;
    const maxPhaseLabel = hasVal(d["最大相"]) ? d["最大相"] + "相"
      : (maxVal != null ? (() => {
          const idx = [...ftVals, ...bkVal].indexOf(maxVal);
          return (idx < 3 ? "前端" : "后端") + phase3[idx % 3];
        })() : "");

    const lineName = g(d["线路名称"], "——");
    const faultPos = g(d["结论"], "——");

    return `
      <section id="sec-defect-${idx}">
        <div class="report-blue-bar sub">${idx}、${escapeHtml(lineName)} ${escapeHtml(g(d['结论'], ''))}</div>

        <h3 class="report-h2">1) 基本信息</h3>
        ${kvTable([
          ["客户单位", g(d['客户单位'], '——')],
          ["检测单位", g(d['检测单位'], '——')],
          ["线路名称", lineName],
          ["检测仪器", g(d['检测仪器'], '架空线双端局放定位仪')],
          ["检测时间", g(d['检测时间'], '——')],
          ["检测人员", g(d['检测人员'], '——')],
        ])}

        <h3 class="report-h2">2) 现场条件</h3>
        ${kvTable([
          ["天气", g(d['天气'], '——')],
          ["温湿度", g(d['温湿度'], '——')],
        ])}

        <h3 class="report-h2">3) 双端局放现场试验条件</h3>
        <p class="report-para"><b>前端设备悬挂杆号：</b>${escapeHtml(g(d['前端设备悬挂杆号'], '——'))}</p>
        <p class="report-para"><b>前端三相幅值：</b>A相 ${g(d['前端A相幅值(mV)'], '——')}mV，B相 ${g(d['前端B相幅值(mV)'], '——')}mV，C相 ${g(d['前端C相幅值(mV)'], '——')}mV</p>
        ${reportImg("前端设备实时图谱（3D柱状图）", "前端图谱")}

        <p class="report-para"><b>后端设备悬挂杆号：</b>${escapeHtml(g(d['后端设备悬挂杆号'], '——'))}</p>
        <p class="report-para"><b>后端三相幅值：</b>A相 ${g(d['后端A相幅值(mV)'], '——')}mV，B相 ${g(d['后端B相幅值(mV)'], '——')}mV，C相 ${g(d['后端C相幅值(mV)'], '——')}mV</p>
        ${reportImg("后端设备实时图谱（3D柱状图）", "后端图谱")}

        <p class="report-para"><b>最大相：</b>${escapeHtml(maxPhaseLabel || '——')}</p>
        ${reportImg("定位图谱 / 双端定位谱图", "定位图谱")}

        ${hasVal(d['定位分析描述']) ? `
        <h3 class="report-h2">分析</h3>
        <p class="report-para">${escapeHtml(g(d['定位分析描述'], '——'))}</p>` : ''}

        ${reportImg("疑似故障杆塔现场照片", "现场照片")}

        <h3 class="report-h2">4) 检测结果分析</h3>
        ${hasVal(d['结论']) ? `
        <h4 class="report-h2" style="font-size:13px;">① 结论</h4>
        <p class="report-para"><b>${escapeHtml(g(d['结论'], '——'))}</b></p>` : ''}

        ${hasVal(d['检测结果分析']) ? `
        <h4 class="report-h2" style="font-size:13px;">② 检测结果分析</h4>
        ${splitSentences(d['检测结果分析']).map(s => `<p class="report-para">${escapeHtml(s.trim())}</p>`).join("")}` : ''}

        ${hasVal(d['疑似杆情况说明']) ? `
        <h4 class="report-h2" style="font-size:13px;">③ 疑似杆情况说明</h4>
        <p class="report-para">${escapeHtml(g(d['疑似杆情况说明'], '——'))}</p>` : ''}

        ${hasVal(d['处理意见']) ? `
        <h4 class="report-h2" style="font-size:13px;">④ 处理意见</h4>
        <p class="report-para"><b>${escapeHtml(g(d['处理意见'], '——'))}</b></p>` : ''}

        ${hasVal(d['处理意见']) ? reportImg("绝缘子疑似缺陷位置 / 捆扎线破损位置", "缺陷照片") : ''}

        ${hasVal(d['验证描述（含绝缘电阻值MΩ）']) ? `
        <h4 class="report-h2" style="font-size:13px;">⑤ 验证</h4>
        ${splitSentences(d['验证描述（含绝缘电阻值MΩ）']).map(s => `<p class="report-para">${escapeHtml(s.trim())}</p>`).join("")}
        <p class="report-para" style="font-size:11px;color:#64748b;">依据DL/T 596—2021《电力设备预防性试验规程》关于10kV架空线路针式绝缘子绝缘电阻不应低于300MΩ的规定，判定该绝缘子已严重劣化，属零值（低值）绝缘子，绝缘性能已基本失效。</p>
        <p class="report-para" style="font-size:11px;color:#64748b;">上述检测结果与KPD-800T架空线局放双端定位仪所给出的局放定位图谱高度吻合，充分验证了该设备对架空线路潜伏性绝缘缺陷的双端定位准确性与检测有效性，为线路状态检修及供电可靠性提升提供了可靠的技术依据。</p>` : ''}
      </section>`;
  }

  function getReportSections(info, defects) {
    const list = defects || [];
    return [
      { id: "sec-cover", num: "Ⅰ", title: "封面", sub: "检 测 报 告" },
      { id: "sec-toc", num: "Ⅱ", title: "目录", sub: "章节导航" },
      { id: "sec-purpose", num: "一", title: "检测目的", sub: "本次检测技术依据" },
      { id: "sec-principle", num: "二", title: "定位原理", sub: "双端同步 · 行波时差" },
      { id: "sec-equipment", num: "三", title: "本次使用设备", sub: "KPD-800T 主要特性" },
    ].concat(
      list.length ? [{ id: "sec-defects", num: "四", title: "缺陷综合检测报告", sub: `共 ${list.length} 处缺陷` }] : []
    ).concat(
      list.map((d, i) => ({
        id: `sec-defect-${i+1}`,
        num: `${i+1}`,
        title: `缺陷 ${i+1}`,
        sub: `${g(d['线路名称'], "——")} · ${g(d['结论'], "").slice(0, 20)}`
      }))
    ).concat([
      { id: "sec-end", num: "Ⅴ", title: "结尾说明", sub: "依据与备注" }
    ]);
  }

  function renderDataTable(rows, container) {
    if (!rows || !rows.length) {
      container.innerHTML = `<div class="empty-state">尚未提取数据，请上传 Excel 或<a href="#" onclick="event.preventDefault();window.KPDA.downloadTemplate();" style="color:var(--accent-cyan);"> 下载数据模板</a>。</div>`;
      return;
    }
    const headers = HEADERS.slice(0, 24); // 显示前24列（不含图片路径）
    const thead = `<thead><tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>`;
    const tbody = `<tbody>${rows.map(r => `<tr>${headers.map(h => `<td>${escapeHtml(r[h] == null ? "" : r[h])}</td>`).join("")}</tr>`).join("")}</tbody>`;
    container.innerHTML = `<div class="data-table-wrap"><table class="data-table">${thead}${tbody}</table></div>`;
  }

  // ============================================================
  // 8. sessionStorage 持久化
  // ============================================================
  function saveReport(info, defects, editedHtml) {
    try {
      sessionStorage.setItem(SKEY, JSON.stringify({
        info: info || state.info,
        defects: defects || state.defects,
        defectRows: (defects || state.defects).map(d => d.display || {}),
        editedHtml: editedHtml || null,
        ts: Date.now()
      }));
    } catch (e) { /* ignore */ }
  }
  function loadReport() {
    try { const raw = sessionStorage.getItem(SKEY); return raw ? JSON.parse(raw) : null; }
    catch (e) { return null; }
  }
  function saveEditedHtml(html) {
    try {
      const p = loadReport() || { info: state.info, defects: state.defects, defectRows: state.defectRows };
      p.editedHtml = html;
      sessionStorage.setItem(SKEY, JSON.stringify(p));
    } catch (e) {}
  }
  function clearReport() {
    try { sessionStorage.removeItem(SKEY); } catch (e) {}
  }

  // ============================================================
  // 9. 客户端 docx 下载
  // ============================================================
  async function downloadWord(filename) {
    const data = state.defects.length ? state : (loadReport() || state);
    if (window.KPDADocx && typeof window.KPDADocx.generateDocxBlob === "function") {
      try {
        const blob = await window.KPDADocx.generateDocxBlob(data.info, data.defects);
        downloadBlob(blob, filename || "架空线双端局放检测报告.docx");
        return true;
      } catch (e) {
        console.warn("docx 生成失败，回退 HTML → .doc", e);
      }
    }
    const content = document.querySelector("#reportDoc")?.innerHTML || buildReportHtml(data.info, data.defects);
    const html = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset="utf-8"><title>检测报告</title></head><body style="font-family:'Microsoft YaHei';">${content}</body></html>`;
    const blob = new Blob(["\ufeff", html], { type: "application/msword" });
    downloadBlob(blob, (filename || "架空线双端局放检测报告") + ".doc");
    return false;
  }

  // ============================================================
  // 10. 初始化
  // ============================================================
  function initDemo() {
    state.info = infoDefault();
    state.defects = makeExamples();
    state.defectRows = state.defects.map(d => d.display || {});
    saveReport();
    return state;
  }

  // ============================================================
  // 暴露
  // ============================================================
  window.KPDA = {
    HEADERS, BASIC_INFO, NOTES, COL_MAP, SKEY, INFO_MAP,
    state, infoDefault, makeExamples,
    escapeHtml, hasVal, num, g, splitSentences,
    imgPH, reportImg, kvTable, downloadBlob,
    buildReportHtml, defectBlockHtml, getReportSections, renderDataTable,
    downloadTemplate, readWorkbook, parseFile,
    saveReport, loadReport, saveEditedHtml, clearReport,
    downloadWord, initDemo
  };
})();
