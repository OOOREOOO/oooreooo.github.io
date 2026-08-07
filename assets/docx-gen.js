﻿/* =============================================================================
 *  docx-gen.js v2.1 — browser-side .docx generation
 *  Aligned with report_builder.py (Python reference)
 *  Supports: WPS embedded images, Excel image paths → Word embedding
 * =========================================================================== */
(function () {
  "use strict";
  var W = window;
  if (!W.KPDA) { console.warn("[docx-gen] KPDA not loaded"); return; }
  var KPDA = W.KPDA;
  var D = W.docx;
  if (!D) { console.warn("[docx-gen] docx.iife.js not loaded"); return; }

  var Document = D.Document, Packer = D.Packer, Paragraph = D.Paragraph,
      TextRun = D.TextRun, Table = D.Table, TableRow = D.TableRow,
      TableCell = D.TableCell, ImageRun = D.ImageRun, Header = D.Header,
      Footer = D.Footer, AlignmentType = D.AlignmentType,
      BorderStyle = D.BorderStyle, WidthType = D.WidthType,
      VerticalAlign = D.VerticalAlign, ShadingType = D.ShadingType;

  // ============= helpers =============
  var tw = function(cm) { return Math.round(cm * 567); };
  var fill = function(hex) { return { type: ShadingType.CLEAR, fill: hex, color: "auto" }; };
  var tb = { style: BorderStyle.SINGLE, size: 4, color: "94A3B8" };
  var BORD = { top: tb, bottom: tb, left: tb, right: tb };
  var COVER_BLUE = "10407B", HEAD_BLUE = "1A5276", BAR_LIGHT = "D4E6F1", CELL_LABEL = "EBF5FB";
  var CN = "Microsoft YaHei", HT = "SimHei";

  var r = function(text, o) {
    o = o || {};
    return new TextRun({
      text: String(text == null ? "" : text),
      font: o.font || CN,
      size: Math.round((o.size || 10.5) * 2),
      bold: o.bold || false,
      color: o.color || "1B1B1B",
      italics: !!o.italic
    });
  };

  var para = function(txt, o) {
    o = o || {};
    var ch = typeof txt === "string" ? [r(txt, o)] : txt;
    var opts = { children: ch };
    if (o.align) opts.alignment = o.align;
    if (o.shading) opts.shading = o.shading;
    if (o.indent) opts.indent = o.indent;
    if (o.border) opts.border = o.border;
    opts.spacing = { before: o.spb || 0, after: o.spa !== undefined ? o.spa : 80 };
    return new Paragraph(opts);
  };

  var cell = function(text, o) {
    o = o || {};
    var al = o.align === "center" ? AlignmentType.CENTER
           : o.align === "right"  ? AlignmentType.RIGHT : AlignmentType.LEFT;
    return new TableCell({
      shading: o.fill ? fill(o.fill) : undefined,
      width: o.width != null ? { size: o.width, type: WidthType.DXA } : undefined,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: al, spacing: { before: 0, after: 0 },
        children: [r(text, { size: o.size || 10, bold: o.bold || false,
          color: o.color || (o.bold ? "1A5276" : "1B1B1B"), font: o.font || CN })]
      })]
    });
  };

  function kvTable(rows, w0, w1) {
    w0 = w0 || tw(3.0); w1 = w1 || tw(5.5);
    var trs = rows.map(function(kv) {
      var cells = [];
      for (var ci = 0; ci < kv.length; ci += 2) {
        cells.push(cell(kv[ci] || "", { fill: CELL_LABEL, bold: true, width: w0, size: 10 }));
        cells.push(cell(kv[ci+1] || "", { width: w1, size: 10 }));
      }
      return new TableRow({ children: cells });
    });
    return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: BORD, rows: trs });
  }

  function bar(text, bfill, txtColor, size) {
    return new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
      shading: fill(bfill),
      children: [r(text, { size: size || 10, bold: true, color: txtColor || "FFFFFF", font: CN })]
    });
  }

  // ===== Image block: matches report_builder.py add_img_placeholder =====
  // Creates a captioned image/placeholder block
  function imgBlock(caption, imgBuf, innerText, widthPx, heightPx) {
    var children = [];
    widthPx = widthPx || 400;
    heightPx = heightPx || 220;

    if (imgBuf) {
      // Real image embedded
      children.push(new ImageRun({ data: imgBuf, transformation: { width: widthPx, height: heightPx } }));
    } else {
      // Placeholder with inner text
      if (innerText) {
        children.push(r(innerText + "\n\n", { size: 10, color: "1B1B1B", font: CN }));
      }
      children.push(r("【" + (caption || "图片 - AI批量导入") + "】", { size: 10, color: "666666", font: CN }));
    }

    var elements = [];
    if (caption && !imgBuf) {
      // Only show caption above for non-image placeholders (or adjust as needed)
    }
    elements.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 120 },
      children: children
    }));

    return elements;
  }

  // ===== Fetch image by URL =====
  async function fetchImgBuf(url) {
    if (!url || typeof url !== "string" || !url.trim()) return null;
    url = url.trim();
    if (/^[a-zA-Z]:[\\/]/.test(url)) return null; // local paths unsupported in browser
    try {
      var resp = await fetch(url);
      if (!resp.ok) return null;
      return await resp.arrayBuffer();
    } catch(e) { return null; }
  }

  // =====================================================
  // Cover page (matches report_builder.py COVER)
  // =====================================================
  function buildCover(ri) {
    var el = [];
    el.push(new Paragraph({
      spacing: { before: 120, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } },
      children: [r("")]
    }));
    el.push(new Paragraph({
      spacing: { before: 240, after: 0 }, indent: { firstLine: 280 },
      children: [r(KPDA.g(ri["标题"], "架空线双端局放检测报告"), { size: 14, bold: true, font: HT })]
    }));
    el.push(new Paragraph({
      spacing: { before: 240, after: 0 }, indent: { firstLine: 280 },
      children: [r(KPDA.g(ri["副标题"], "10kV架空线双端局放定位仪巡检"), { size: 14, bold: true, font: HT })]
    }));
    for (var i = 0; i < 2; i++) el.push(new Paragraph({ children: [r("")] }));
    el.push(new Paragraph({
      alignment: AlignmentType.CENTER, shading: fill(COVER_BLUE),
      spacing: { before: 60, after: 60 },
      children: [r("检 测 报 告", { size: 60, bold: true, color: "FFFFFF", font: HT })]
    }));
    for (var j = 0; j < 10; j++) el.push(new Paragraph({ children: [r("")] }));
    el.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      children: [r(KPDA.g(ri["出具单位"], "咸亨国际科技股份有限公司"), { size: 15, font: HT })]
    }));
    el.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      children: [r(KPDA.g(ri["报告日期"], "二〇二六年八月"), { size: 15, font: HT })]
    }));
    return el;
  }

  // =====================================================
  // TOC with tab stops + dot leaders + page numbers
  // =====================================================
  function buildTOC(defects) {
    var el = [];
    el.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 240, after: 320 },
      children: [r("目  录", { size: 18, bold: true, font: HT })]
    }));
    var entries = [
      ["一、检测目的", "3"],
      ["二、架空线双端局放定位检测原理", "3"],
      ["三、本次使用设备", "4"],
    ];
    if (defects && defects.length) {
      entries.push(["四、缺陷综合检测报告", "5"]);
      defects.forEach(function(d, i) {
        var title = (i+1) + "、" + KPDA.g(d["结论"], "缺陷 " + (i+1)).slice(0, 40);
        entries.push([title, String(5 + i * 5)]);
      });
    }
    entries.forEach(function(entry) {
      el.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        tabStops: [{ type: "right", position: tw(14.0), leader: "dot" }],
        children: [r(entry[0], { size: 11.5, font: CN }), new TextRun({ text: "\t" }), r(entry[1], { size: 11.5, font: CN })]
      }));
    });
    return el;
  }

  // =====================================================
  // Defect block — matches report_builder.py defect_block
  // =====================================================
  function defectBlock(idx, d, isFirst) {
    var out = [];

    // Section heading (only on first defect)
    if (isFirst) {
      out.push(new Paragraph({
        spacing: { before: 360, after: 200 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2980B9", space: 1 } },
        children: [r("四、缺陷综合检测报告", { size: 16, bold: true, color: "1A5276", font: CN })]
      }));
    }

    // Sub-title: "1、{line} {conclusion}"
    var titleText = idx + "、" + KPDA.g(d["结论"],
      KPDA.g(d["线路名称"], "10kV") + " " + KPDA.g(d["前端设备悬挂杆号"], "") + " 绝缘子存在局部放电隐患");
    out.push(new Paragraph({
      spacing: { before: 200, after: 160 },
      children: [r(titleText, { size: 13, bold: true, color: "0E3A5F", font: CN })]
    }));

    // ===== 1）基本信息 =====
    out.push(bar("1）基本信息", HEAD_BLUE, "FFFFFF", 10));
    out.push(kvTable([
      ["客户单位", KPDA.g(d["客户单位"], "【填写客户单位】"),
       "检测单位", KPDA.g(d["检测单位"], "咸亨国际科技股份有限公司")],
      ["线路名称", KPDA.g(d["线路名称"], "【填写线路名称】"),
       "检测仪器", KPDA.g(d["检测仪器"], "架空线双端局放定位仪")],
      ["检测时间", KPDA.g(d["检测时间"], "【填写检测时间】"),
       "检测人员", KPDA.g(d["检测人员"], "【填写检测人员】")],
    ]));

    // ===== 2）现场条件 =====
    out.push(bar("2）现场条件", HEAD_BLUE, "FFFFFF", 10));
    out.push(kvTable([
      ["天气", KPDA.g(d["天气"], "【填写天气】"),
       "温湿度", KPDA.g(d["温湿度"], "【填写温度/湿度】")],
    ]));

    // ===== 3）双端局放现场试验条件 =====
    out.push(bar("3）双端局放现场试验条件", HEAD_BLUE, "FFFFFF", 10));

    // --- Front-end ---
    out.push(new Paragraph({
      spacing: { before: 60, after: 40 },
      children: [r("前端设备悬挂杆号：" + KPDA.g(d["前端设备悬挂杆号"], "【填写】"), { size: 10.5, font: CN })]
    }));

    // Calculate max phase values for caption
    var fa = KPDA.num(d["前端A相幅值(mV)"]), fb = KPDA.num(d["前端B相幅值(mV)"]), fc = KPDA.num(d["前端C相幅值(mV)"]);
    var ftHasVals = (fa != null && fb != null && fc != null);
    if (ftHasVals) {
      var fvs = { A: fa, B: fb, C: fc };
      var fmx = Object.keys(fvs).reduce(function(a,b){ return fvs[a] > fvs[b] ? a : b; });
      var fot = ["A","B","C"].filter(function(k){ return k !== fmx; });
      var ftInner = "设备实时图谱（三相幅值最大的" + fmx + "相数值为" + Math.round(fvs[fmx]) + "mV，其余两相分别为" + Math.round(fvs[fot[0]]) + "mV、" + Math.round(fvs[fot[1]]) + "mV）";
    } else {
      var ftInner = "设备实时图谱（三相幅值最大的【相】数值为____mV，其余两相分别为____mV、____mV）";
    }
    out = out.concat(imgBlock("前端设备实时图谱（3D柱状图）", d._frontendImgBuf, ftInner, 460, 280));

    // --- Back-end ---
    out.push(new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [r("后端设备悬挂杆号：" + KPDA.g(d["后端设备悬挂杆号"], "【填写】"), { size: 10.5, font: CN })]
    }));

    var ba = KPDA.num(d["后端A相幅值(mV)"]), bb = KPDA.num(d["后端B相幅值(mV)"]), bc = KPDA.num(d["后端C相幅值(mV)"]);
    var bkHasVals = (ba != null && bb != null && bc != null);
    if (bkHasVals) {
      var bvs = { A: ba, B: bb, C: bc };
      var bmx = Object.keys(bvs).reduce(function(a,b){ return bvs[a] > bvs[b] ? a : b; });
      var bot = ["A","B","C"].filter(function(k){ return k !== bmx; });
      var bkInner = "设备实时图谱（三相幅值最大的" + bmx + "相数值为" + Math.round(bvs[bmx]) + "mV，其余两相分别为" + Math.round(bvs[bot[0]]) + "mV、" + Math.round(bvs[bot[1]]) + "mV）";
    } else {
      var bkInner = "设备实时图谱（三相幅值最大的【相】数值为____mV，其余两相分别为____mV、____mV）";
    }
    out = out.concat(imgBlock("后端设备实时图谱（3D柱状图）", d._backendImgBuf, bkInner, 460, 280));

    // --- Max phase + positioning spectrum ---
    var maxPhase = KPDA.g(d["最大相"], "");
    out.push(new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [r("最大相：" + (maxPhase || "——"), { size: 10.5, font: CN })]
    }));
    out = out.concat(imgBlock("定位图谱（2D散点图）+ 双端定位谱图", d._positionImgBuf, "定位图谱 / 双端定位谱图", 460, 280));

    // --- Analysis ---
    if (KPDA.hasVal(d["定位分析描述"])) {
      out.push(bar("分析", BAR_LIGHT, "1B1B1B", 11));
      out.push(para(KPDA.g(d["定位分析描述"], "【填写】"), {
        size: 11, indent: { firstLine: 0 }, align: AlignmentType.JUSTIFIED, spb: 40, spa: 80
      }));
    }

    // --- Field photo ---
    out = out.concat(imgBlock("疑似故障杆塔现场照片", d._defectImgBuf, "疑似故障杆塔现场照片", 420, 280));

    // ===== 3）检测结果分析 (matches reference numbering) =====
    var hasAnalysisContent =
      KPDA.hasVal(d["结论"]) ||
      KPDA.hasVal(d["检测结果分析"]) ||
      KPDA.hasVal(d["疑似杆情况说明"]) ||
      KPDA.hasVal(d["处理意见"]) ||
      KPDA.hasVal(d["验证描述（含绝缘电阻值MΩ）"]);

    if (hasAnalysisContent) {
      out.push(bar("3）检测结果分析", HEAD_BLUE, "FFFFFF", 10));
    }

    // Conclusion
    if (KPDA.hasVal(d["结论"])) {
      out.push(bar("结论", BAR_LIGHT, "1B1B1B", 11));
      out.push(new Paragraph({
        spacing: { before: 40, after: 80 },
        children: [r(KPDA.g(d["结论"], "【填写结论】"), { size: 10.5, bold: true, font: CN })]
      }));
    }

    // Result analysis
    if (KPDA.hasVal(d["检测结果分析"])) {
      out.push(bar("检测结果分析", BAR_LIGHT, "1B1B1B", 11));
      KPDA.splitSentences(d["检测结果分析"]).forEach(function(s) {
        if (s.trim()) out.push(para(s.trim(), {
          size: 11, indent: { firstLine: 0 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 60
        }));
      });
    }

    // Pole condition
    if (KPDA.hasVal(d["疑似杆情况说明"])) {
      out.push(bar("疑似杆情况说明", BAR_LIGHT, "1B1B1B", 11));
      out.push(para(KPDA.g(d["疑似杆情况说明"], "【填写】"), {
        size: 11, indent: { firstLine: 0 }, align: AlignmentType.JUSTIFIED, spb: 40, spa: 60
      }));
    }

    // Advice
    if (KPDA.hasVal(d["处理意见"])) {
      out.push(bar("处理意见", BAR_LIGHT, "1B1B1B", 11));
      out.push(new Paragraph({
        spacing: { before: 40, after: 80 },
        children: [r(KPDA.g(d["处理意见"], "【填写处理意见】"), { size: 10.5, bold: true, font: CN })]
      }));
    }

    // ===== Defect photos (2-column layout) =====
    if (KPDA.hasVal(d["处理意见"])) {
      out.push(bar("现场疑似缺陷图片", BAR_LIGHT, "1B1B1B", 11));
      // Two-column image layout: left=绝缘子疑似缺陷位置, right=捆扎线破损位置
      var defectImgCells = [];
      // Left cell
      defectImgCells.push(new TableCell({
        shading: fill(CELL_LABEL),
        margins: { top: 30, bottom: 30, left: 50, right: 50 },
        verticalAlign: VerticalAlign.CENTER,
        width: { size: tw(8.0), type: WidthType.DXA },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
            children: [r("绝缘子疑似缺陷位置", { size: 10, bold: true, font: CN })]
          }),
          d._defectImgBuf
            ? new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: d._defectImgBuf, transformation: { width: 320, height: 240 } })] })
            : new Paragraph({ alignment: AlignmentType.CENTER, children: [r("【绝缘子疑似缺陷照片】", { size: 10, color: "666666", font: CN })] })
        ]
      }));
      // Right cell
      defectImgCells.push(new TableCell({
        shading: fill(CELL_LABEL),
        margins: { top: 30, bottom: 30, left: 50, right: 50 },
        verticalAlign: VerticalAlign.CENTER,
        width: { size: tw(8.0), type: WidthType.DXA },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
            children: [r("捆扎线破损位置", { size: 10, bold: true, font: CN })]
          }),
          d._defectImgBuf2
            ? new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: d._defectImgBuf2, transformation: { width: 320, height: 240 } })] })
            : new Paragraph({ alignment: AlignmentType.CENTER, children: [r("【捆扎线破损位置照片】", { size: 10, color: "666666", font: CN })] })
        ]
      }));
      out.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: BORD,
        rows: [new TableRow({ children: defectImgCells })]
      }));
    }

    // ===== Verification ===== (only shown when data exists)
    if (KPDA.hasVal(d["验证描述（含绝缘电阻值MΩ）"])) {
      out.push(bar("验证", BAR_LIGHT, "1B1B1B", 11));
      KPDA.splitSentences(d["验证描述（含绝缘电阻值MΩ）"]).forEach(function(s) {
        if (s.trim()) out.push(para(s.trim(), {
          size: 11, indent: { firstLine: 0 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 60
        }));
      });
      out.push(para("依据DL/T 596—2021《电力设备预防性试验规程》关于10kV架空线路针式绝缘子绝缘电阻不应低于300MΩ的规定，判定该绝缘子已严重劣化，属零值（低值）绝缘子，绝缘性能已基本失效。", {
        size: 11, indent: { firstLine: 0 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 60, color: "475569"
      }));
      out.push(para("上述检测结果与KPD-800T架空线局放双端定位仪所给出的局放定位图谱高度吻合，充分验证了该设备对架空线路潜伏性绝缘缺陷的双端定位准确性与检测有效性，为线路状态检修及供电可靠性提升提供了可靠的技术依据。", {
        size: 11, indent: { firstLine: 0 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 60, color: "475569"
      }));
    }

    return out;
  }

  // =====================================================
  // Pre-fetch images from defect data
  // =====================================================
  async function prefetchDefectImages(defect) {
    if (!defect) return;
    var frontend  = KPDA.g(defect["前端实时图谱图片路径"], "");
    var backend   = KPDA.g(defect["后端实时图谱图片路径"], "");
    var position  = KPDA.g(defect["双端定位谱图图片路径"], "");
    var defectImg = KPDA.g(defect["现场疑似缺陷图片路径"], "");

    var results = await Promise.all([
      fetchImgBuf(frontend),
      fetchImgBuf(backend),
      fetchImgBuf(position),
      fetchImgBuf(defectImg)
    ]);

    defect._frontendImgBuf = results[0] || null;
    defect._backendImgBuf  = results[1] || null;
    defect._positionImgBuf = results[2] || null;
    defect._defectImgBuf   = results[3] || null;
    defect._defectImgBuf2  = results[3] || null; // same image for now
  }

  // =====================================================
  // Main entry
  // =====================================================
  async function generateDocxBlob(info, defects) {
    var ri = Object.assign(KPDA.infoDefault(), info || {});
    var list = (defects && defects.length) ? defects : [];

    // Pre-fetch defect images
    for (var i = 0; i < list.length; i++) {
      await prefetchDefectImages(list[i]);
    }

    // Fetch fixed images
    var principle = await fetchImgBuf("assets/img/principle.png");
    var topology = await fetchImgBuf("assets/img/topology.png");
    var product = await fetchImgBuf("assets/img/product.jpeg");
    var xianhengLogo = await fetchImgBuf("assets/img/logo_xianheng.png");

    var sectionProps = {
      page: {
        size: { width: tw(21.01), height: tw(29.70) },
        margin: { top: tw(2.01), bottom: tw(0.49), left: tw(1.06), right: tw(1.20) }
      }
    };

    // Cover + TOC
    var coverElements = buildCover(ri);
    var tocElements = buildTOC(list);

    // Content
    var contentElements = [];

    // === Section 1: 检测目的 ===
    contentElements.push(new Paragraph({
      spacing: { before: 360, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2980B9", space: 1 } },
      children: [r("一、检测目的", { size: 16, bold: true, color: "1A5276", font: CN })]
    }));
    contentElements.push(para("针对本次 " + KPDA.g(ri["标题"], "检测任务") + "，采用 KPD-800T 架空线双端局放定位系统开展现场带电检测，识别并定位线路潜在局放缺陷，评估缺陷类别与严重等级，为线路检修与运行维护提供技术依据。", {
      size: 11, indent: { firstLine: 360 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 80
    }));

    // === Section 2: 定位原理 ===
    contentElements.push(new Paragraph({
      spacing: { before: 360, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2980B9", space: 1 } },
      children: [r("二、架空线双端局放定位检测原理", { size: 16, bold: true, color: "1A5276", font: CN })]
    }));
    if (principle) {
      contentElements.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: principle, transformation: { width: 600, height: 340 } })]
      }));
      contentElements.push(new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 20, after: 80 },
        children: [r("图  双端局放检测原理示意图", { size: 10, italics: true, color: "666666", font: CN })]
      }));
    }
    contentElements.push(para("KPD-800T 架空线双端局放定位系统采用双端行波时差法，在线路两端各布置一台检测装置，利用 GPS/北斗秒脉冲同步与高速采样，捕获同一局放信号在前后端的到达时间差，结合传播速度与线路长度计算放电源位置；可识别表面、电晕、内部等多种放电类型。", {
      size: 11, indent: { firstLine: 360 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 80
    }));

    // === Section 3: 本次使用设备 ===
    contentElements.push(new Paragraph({
      spacing: { before: 360, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2980B9", space: 1 } },
      children: [r("三、本次使用设备", { size: 16, bold: true, color: "1A5276", font: CN })]
    }));
    if (product || topology) {
      var imgCells = [];
      if (product) imgCells.push(new TableCell({
        shading: fill(CELL_LABEL), margins: { top: 30, bottom: 30, left: 50, right: 50 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: product, transformation: { width: 370, height: 280 } })] })]
      }));
      if (topology) imgCells.push(new TableCell({
        shading: fill(CELL_LABEL), margins: { top: 30, bottom: 30, left: 50, right: 50 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: topology, transformation: { width: 400, height: 280 } })] })]
      }));
      if (imgCells.length > 0) {
        contentElements.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: BORD, rows: [new TableRow({ children: imgCells })] }));
      }
    }
    var bullets = [
      ["支持中压架空线局部放电带电检测", "设备专为中压架空线路设计，可在不停电状态下实时监测局部放电情况，提前发现绝缘缺陷，避免停电检测带来的供电中断。"],
      ["支持中压架空线局部放电精准定位", "采用双端检测原理，通过捕捉放电信号的传播时间差，实现对放电点的精准定位，精度可达0.5米或0.5%线缆长度。"],
      ["自动生成立位谱图，直接给出放电位置及强度", "检测过程中自动生成立位谱图，直观显示放电位置和强度，便于快速分析和决策。"],
      ["采用电池供电，使用方便", "设备采用电池供电，无需外部电源，适合户外或无电源环境。内置大容量电池可支持长时间连续工作，续航能力达3小时以上。"],
    ];
    bullets.forEach(function(b) {
      contentElements.push(para("■ " + b[0], { size: 10.5, bold: true, spb: 0, spa: 20 }));
      contentElements.push(para(b[1], { size: 10.5, indent: { firstLine: 360 }, align: AlignmentType.JUSTIFIED, spb: 0, spa: 80 }));
    });

    // === Section 4: Defects ===
    if (!list.length) {
      contentElements.push(para("(暂无缺陷记录，请在 Excel 中填写后重新生成。)", {
        size: 11, color: "888888", align: AlignmentType.CENTER, spb: 200, spa: 200
      }));
    } else {
      list.forEach(function(d, i) {
        if (i > 0) contentElements.push(new Paragraph({ children: [r("")], pageBreakBefore: true }));
        var block = defectBlock(i + 1, d, i === 0);
        block.forEach(function(p) { contentElements.push(p); });
      });
    }

    // End note
    contentElements.push(new Paragraph({
      spacing: { before: 240, after: 80 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "AEB6BF", space: 1 } },
      children: [r("注：因无近距离及解体分析，结论为初步判断，请以最终解体或近距离确诊为准，合理检修。", { size: 10, italics: true, color: "666666", font: CN })]
    }));

    // Header with logo
    var headerChildren;
    if (xianhengLogo) {
      headerChildren = [new ImageRun({ data: xianhengLogo, transformation: { width: 180, height: 45 } })];
    } else {
      headerChildren = [r("咸亨国际", { size: 10, font: HT, color: "475569" })];
    }

    // Footer
    var footerPara = new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        r("第", { size: 9, color: "64748b", font: CN }),
        new TextRun({ children: ["1"], font: CN, size: 18, color: "64748b" }),
        r(" 页 / 共 ", { size: 9, color: "64748b", font: CN }),
        new TextRun({ children: ["1"], font: CN, size: 18, color: "64748b" }),
        r(" 页", { size: 9, color: "64748b", font: CN }),
      ]
    });

    var doc = new Document({
      creator: "KPD-800T Report System",
      title: "架空线双端局放检测报告",
      sections: [
        { properties: sectionProps, children: coverElements },
        { properties: sectionProps, children: tocElements },
        {
          properties: sectionProps,
          headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: headerChildren })] }) },
          footers: { default: new Footer({ children: [footerPara] }) },
          children: contentElements
        }
      ]
    });

    return await Packer.toBlob(doc);
  }

  W.KPDADocx = { generateDocxBlob: generateDocxBlob };
})();
