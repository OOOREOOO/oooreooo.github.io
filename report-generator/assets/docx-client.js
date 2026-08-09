/* ==========================================================================
   docx-client.js — 客户端 Word 生成 (GitHub Pages 静态部署)
   使用 docx.umd.js 库，无需后端服务
   v19: 目录加页码 / 一、二、三内容完整 / 图片插入 / 字段名同步
   ========================================================================== */
(function () {
  "use strict";
  var W = window;
  if (!W.KPDA) { console.warn("[docx-client] KPDA not loaded"); return; }
  var KPDA = W.KPDA;
  var D = W.docx;
  if (!D) { console.warn("[docx-client] docx.umd.js not loaded"); return; }

  var Document = D.Document, Packer = D.Packer, Paragraph = D.Paragraph,
      TextRun = D.TextRun, Table = D.Table, TableRow = D.TableRow,
      TableCell = D.TableCell, ImageRun = D.ImageRun, Header = D.Header,
      Footer = D.Footer, AlignmentType = D.AlignmentType,
      BorderStyle = D.BorderStyle, WidthType = D.WidthType,
      VerticalAlign = D.VerticalAlign, ShadingType = D.ShadingType,
      PageBreak = D.PageBreak, PageNumber = D.PageNumber,
      TabStopType = D.TabStopType, TabStopPosition = D.TabStopPosition,
      Tab = D.Tab;

  var CN = "Microsoft YaHei";
  function pts(n) { return Math.round(n * 20); }
  var MARGIN = pts(61.9), CONTENT_W = pts(471.6);

  var DARK = "0E3A5F", LIGHT = "D4E6F1", CELL_BG = "EBF5FB";
  var NO_BORD = { top:{style:BorderStyle.NONE,size:0,color:"auto",space:0}, bottom:{style:BorderStyle.NONE,size:0,color:"auto",space:0},
                  left:{style:BorderStyle.NONE,size:0,color:"auto",space:0}, right:{style:BorderStyle.NONE,size:0,color:"auto",space:0} };
  var NO_BORD_IN = Object.assign({}, NO_BORD,
    { insideHorizontal:{style:BorderStyle.NONE,size:0,color:"auto",space:0}, insideVertical:{style:BorderStyle.NONE,size:0,color:"auto",space:0} });

  function r(txt, o) {
    o = o || {};
    return new TextRun({ text: String(txt == null ? "" : txt),
      font: o.font || CN, size: o.size || 22, bold: !!o.bold,
      color: o.color || "1B1B1B", italics: !!o.italic });
  }

  function shd(hex) {
    return { type: ShadingType.CLEAR, fill: hex, color: "auto" };
  }

  function cell(txt, o) {
    o = o || {};
    var al = o.align === "center" ? AlignmentType.CENTER
           : o.align === "right"  ? AlignmentType.RIGHT : AlignmentType.LEFT;
    return new TableCell({
      borders: NO_BORD,
      shading: o.fill ? shd(o.fill) : undefined,
      width: o.width != null ? { size: o.width, type: WidthType.DXA } : undefined,
      margins: { top: 120, bottom: 120, left: 140, right: 140 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: al, spacing: { before: 40, after: 40, line: 320 },
        children: [r(txt, { size: o.size || 22, bold: !!o.bold, color: o.color || "1B1B1B", font: o.font || CN })]
      })]
    });
  }

  function bar(text, isDark) {
    // 用 1×1 表格包裹，让底纹落在单元格级（tcPr/shd），与KV行对齐
    var color = isDark ? DARK : LIGHT;
    var textColor = isDark ? "FFFFFF" : "1B1B1B";
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],  // 显式列宽=总宽，确保单 cell 表撑满 CONTENT_W（与 KV 表右对齐）
      borders: NO_BORD_IN,
      rows: [new TableRow({
        children: [new TableCell({
          borders: NO_BORD,
          shading: shd(color),
          width: { size: CONTENT_W, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 140, right: 140 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 20, after: 20 },
            children: [r(text, { size: 26, bold: true, color: textColor, font: CN })]
          })]
        })]
      })]
    });
  }

  function hdg(text) {
    return new Paragraph({
      spacing: { before: 300, after: 160 },
      children: [r(text, { size: 32, bold: true, color: "0E3A5F", font: CN })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "0E3A5F", space: 4 } }
    });
  }

  function kvt(rows, widths) {
    var trs = rows.map(function(kv) {
      var cells = [];
      for (var ci = 0; ci < kv.length; ci += 2) {
        var w = widths[ci/2];
        cells.push(cell(kv[ci] || "", { fill: CELL_BG, bold: true, color: "0E3A5F", size: 20, width: w[0] }));
        cells.push(cell(kv[ci+1] || "", { width: w[1] }));
      }
      return new TableRow({ children: cells });
    });
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: widths.reduce(function(a, w) { return a.concat(w); }, []),
      borders: NO_BORD_IN, rows: trs
    });
  }

  /* KV 表格列宽：精确撑满 CONTENT_W（471.6pt）。
     标签 85pt / 值 150.8pt 严格镜像对称——WPS 加粗中文渲染稍宽，留 ~15pt 余量；标签字号 10pt（size 20） */
  var KV_W = [[pts(85), pts(150.8)], [pts(85), pts(150.8)]];

  /* 读取图片自然宽高（dataURL / http 均适用），计算 aspect */
  function getImageNaturalWH(src) {
    return new Promise(function(resolve) {
      var img = new Image();
      img.onload = function() { resolve({ w: img.naturalWidth, h: img.naturalHeight }); };
      img.onerror = function() { resolve({ w: 3, h: 2 }); };
      img.src = src;
    });
  }

  /* 图片处理：支持 dataURL（base64）。直接 atob 解码，避免 fetch(dataURL) 的浏览器兼容性问题。
     size: number → 按宽度 auto 高度保持比例；{width, height} → 强制指定尺寸 */
  async function imgRun(src, size) {
    var s = String(src || "").trim();
    if (!s || s.indexOf("=DISPIMG(") === 0) return null;
    // 计算目标宽高
    var hasExplicit = (typeof size === "object" && size !== null && size.height != null);
    var targetW, targetH;
    if (typeof size === "object" && size !== null && size.width != null) {
      targetW = size.width;
      targetH = hasExplicit ? size.height : null;
    } else {
      targetW = size || 560;
      targetH = null;
    }
    // 1) dataURL 直接 base64 解码
    var m = /^data:([^;,]+)?(;base64)?,(.*)$/i.exec(s);
    if (m && m[2]) {
      try {
        var b64 = m[3];
        var binStr = atob(b64);
        var len = binStr.length;
        var buf = new Uint8Array(len);
        for (var i = 0; i < len; i++) buf[i] = binStr.charCodeAt(i);
        var mime = (m[1] || "image/png").toLowerCase();
        var ext = mime === "image/jpeg" ? "jpeg" : mime === "image/gif" ? "gif" : "png";
        // 若未指定高度，按自然比例计算
        if (targetH == null) {
          var nat = await getImageNaturalWH(s);
          targetH = Math.round(targetW * nat.h / nat.w);
        }
        return new ImageRun({ data: buf, transformation: { width: targetW, height: targetH },
          type: mime, extension: ext });
      } catch (e) { return null; }
    }
    // 2) http(s)/blob URL 用 fetch
    if (/^(https?:\/\/|blob:)/i.test(s)) {
      try {
        var resp = await fetch(s);
        if (!resp.ok) return null;
        var buf2 = await resp.arrayBuffer();
        var ext2 = (s.split("?")[0].split(".").pop() || "png").toLowerCase();
        var ct = ext2 === "jpg" || ext2 === "jpeg" ? "image/jpeg"
               : ext2 === "gif" ? "image/gif" : "image/png";
        if (targetH == null) {
          var nat2 = await getImageNaturalWH(s);
          targetH = Math.round(targetW * nat2.h / nat2.w);
        }
        return new ImageRun({ data: buf2, transformation: { width: targetW, height: targetH },
          type: ct, extension: ext2 });
      } catch (e) { return null; }
    }
    return null;
  }

  function imgPh(text) {
    return new Paragraph({ alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 40 },
      children: [r(text, { size: 20, italic: true, color: "64748B", font: CN })] });
  }

  /* 后处理器：修复 docx.umd.js 生成的 docx 中 Content-Type/媒体命名 bug。
     症状：[Content_Types].xml 缺少 Default/Override，媒体路径为
       word/media/<hash>.image/<ext>，导致 Word/WPS/python-docx 无法识别图片。
     修复：①媒体文件重命名为 word/media/image{n}.{ext}
          ②更新 word/_rels/*.rels 中的 Target 引用
          ③注入 [Content_Types].xml 的 Default 与 Override */
  async function fixDocxBlob(blob) {
    if (!W.JSZip) return blob; // 无 JSZip 则跳过（理论上不会发生）
    var zip = await W.JSZip.loadAsync(blob);

    // 1. 扫描 word/media/，构建重命名映射
    var renames = {};      // oldPath -> newPath
    var mediaFiles = [];   // [{newPath, ext}]
    var dirEntries = [];   // 待删除的目录条目
    var counter = {};
    Object.keys(zip.files).forEach(function(relPath) {
      if (relPath.indexOf("word/media/") !== 0) return;
      var entry = zip.files[relPath];
      if (entry.dir) { dirEntries.push(relPath); return; }
      var base = relPath.substring("word/media/".length);
      var m = /^(.*?)\.image\/(\w+)$/.exec(base);
      var ext;
      if (m) ext = m[2].toLowerCase();
      else { var dot = base.lastIndexOf("."); ext = dot >= 0 ? base.substring(dot + 1).toLowerCase() : "png"; }
      counter[ext] = (counter[ext] || 0) + 1;
      var newPath = "word/media/image" + counter[ext] + "." + ext;
      renames[relPath] = newPath;
      mediaFiles.push({ newPath: newPath, ext: ext });
    });

    if (Object.keys(renames).length === 0) return blob; // 无媒体文件，无需修复

    // 2. 读取旧文件内容 -> 删除旧条目 -> 写入新文件
    var contents = {};
    for (var oldPath in renames) {
      contents[renames[oldPath]] = await zip.file(oldPath).async("uint8array");
    }
    for (var oldPath in renames) { zip.remove(oldPath); }
    dirEntries.forEach(function(d) { zip.remove(d); });
    for (var np in contents) { zip.file(np, contents[np]); }

    // 3. 更新 word/_rels/*.rels 中的 Target 引用（相对 word/ 目录）
    var relsRegex = /^word\/_rels\/.+\.rels$/;
    var relsFiles = zip.file(relsRegex);
    for (var i = 0; i < relsFiles.length; i++) {
      var rf = relsFiles[i];
      var xml = await rf.async("string");
      var changed = false;
      for (var oldPath in renames) {
        var oldRef = renames[oldPath]; // full path used in Target for image rels? Target in document.xml.rels is relative to word/
        // Target 形如 "media/hash.image/png"
        var oldTarget = oldPath.substring("word/".length);
        var newTarget = renames[oldPath].substring("word/".length);
        if (xml.indexOf(oldTarget) >= 0) {
          xml = xml.split(oldTarget).join(newTarget);
          changed = true;
        }
      }
      if (changed) zip.file(rf.name, xml);
    }

    // 4. 改写 [Content_Types].xml：移除旧 Override，注入 Default 与新 Override
    //    关键：docx.umd.js 在无任何 default/override 时会输出自闭合 <Types .../>，
    //    必须先转为开放形式再注入，否则正则匹配不到。
    var ct = await zip.file("[Content_Types].xml").async("string");
    for (var oldPath in renames) {
      var esc = oldPath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      var re = new RegExp('<Override\\s+PartName="\\/' + esc + '"[^>]*\\/>', "g");
      ct = ct.replace(re, "");
    }
    var mimeOf = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
                   gif: "image/gif", bmp: "image/bmp" };
    var defaultsXml = "";
    Object.keys(counter).forEach(function(ext) {
      if (mimeOf[ext] && ct.indexOf('Extension="' + ext + '"') < 0) {
        defaultsXml += '<Default Extension="' + ext + '" ContentType="' + mimeOf[ext] + '"/>';
      }
    });
    var overridesXml = "";
    mediaFiles.forEach(function(m) {
      overridesXml += '<Override PartName="/' + m.newPath + '" ContentType="' + (mimeOf[m.ext] || "image/png") + '"/>';
    });
    var selfCloseRe = /<Types([^>]*)\/>/;          // <Types .../>
    var openTagRe   = /<Types([^>]*)>/;            // <Types ...>
    var closeTagRe  = /<\/Types>/;                 // </Types>
    if (selfCloseRe.test(ct)) {
      // 自闭合 -> 转为开放并注入
      ct = ct.replace(selfCloseRe, '<Types$1>' + defaultsXml + overridesXml + '</Types>');
    } else if (openTagRe.test(ct)) {
      if (defaultsXml) ct = ct.replace(openTagRe, '<Types$1>' + defaultsXml);
      ct = ct.replace(closeTagRe, overridesXml + "</Types>");
    } else {
      // 兜底：在 <?xml?> 后新建 Types 块
      ct = ct.replace(/(<\?xml[^>]*\?>\s*)/, '$1<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + defaultsXml + overridesXml + '</Types>');
    }
    zip.file("[Content_Types].xml", ct);

    return await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
  }

  /* 拆分"温湿度"组合字段（如 "26℃ / 未知"）为温度/湿度两项。
     容错：只有一项时填温度，湿度留空；无数据时返回 —— */
  function parseTempHumi(s) {
    if (!s) return { temp: "——", humi: "——" };
    var parts = String(s).split(/\s*\/\s*/);
    return {
      temp: parts[0] || "——",
      humi: parts[1] || ""
    };
  }

  async function buildDefectBlock(i, d, isFirst) {
    var el = [];
    if (isFirst) {
      el.push(new Paragraph({ children: [new PageBreak()] }));
      el.push(hdg("四、缺陷综合检测报告"));
    }

    var line = KPDA.g(d["线路名称"], "——");
    el.push(new Paragraph({ spacing: { before: 120, after: 80 },
      children: [r((i+1)+".", { size: 26, bold: true, color: "0E3A5F", font: CN }),
                 r(" "+KPDA.g(d["结论"],""), { size: 26, bold: true, color: "0E3A5F", font: CN })] }));

    // 基本信息/现场条件从 KPDA.state.info 读取（整份报告共享）
    var info = KPDA.state.info || {};
    // 线路名称已从报告基本信息移至缺陷表，整份报告共用首行缺陷的线路名称
    var sharedLine = KPDA.g(d["线路名称"], "——");
    el.push(bar("1）基本信息", true));
    el.push(kvt([
      ["客户单位", KPDA.g(info["客户单位"],"——"), "检测单位", KPDA.g(info["检测单位"],"——")],
      ["线路名称", sharedLine, "检测仪器", KPDA.g(info["检测仪器"],"——")],
      ["检测时间", KPDA.g(info["检测时间"],"——"), "检测人员", KPDA.g(info["检测人员"],"——")],
    ], KV_W));

    el.push(bar("2）现场条件", true));
    el.push(kvt([["天气", KPDA.g(info["天气"],"——"), "温湿度", KPDA.g(info["温湿度"],"——")]], KV_W));

    el.push(bar("3）双端局放现场试验条件", true));

    // 前端设备悬挂杆号行（浅蓝底+加粗+左对齐，与3）双端局放现场试验条件同宽）
    el.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      borders: NO_BORD_IN,
      rows: [new TableRow({
        children: [new TableCell({
          borders: NO_BORD,
          shading: shd(LIGHT),
          width: { size: CONTENT_W, type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 40, after: 40 },
            children: [r("前端设备悬挂杆号：" + KPDA.g(d["前端设备悬挂杆号"],"——"),
              { size: 22, bold: true, color: "1B1B1B", font: CN })]
          })]
        })]
      })]
    }));

    // 前端实时图谱图片
    var imgFt = await imgRun(d["前端实时图谱图片"]);
    if (imgFt) {
      el.push(new Paragraph({ alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [imgFt] }));
    } else {
      el.push(imgPh("[前端实时图谱 - 未检测到图片]"));
    }

    // 前端图谱标题（斜体+灰色）
    el.push(new Paragraph({ alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 80 },
      children: [r("前端设备实时图谱（三相幅值最大的 C 相数值为"+KPDA.g(d["前端C相幅值(mV)"],"")+"mV，其余两相分别为"+KPDA.g(d["前端A相幅值(mV)"],"")+"mV、"+KPDA.g(d["前端B相幅值(mV)"],"")+"mV）",
        { size: 20, italic: true, color: "64748B", font: CN })] }));

    // 后端设备悬挂杆号行（浅蓝底+加粗+左对齐）
    el.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: NO_BORD_IN,
      rows: [new TableRow({
        children: [new TableCell({
          borders: NO_BORD,
          shading: shd(LIGHT),
          width: { size: CONTENT_W, type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 120, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 40, after: 40 },
            children: [r("后端设备悬挂杆号：" + KPDA.g(d["后端设备悬挂杆号"],"——"),
              { size: 22, bold: true, color: "1B1B1B", font: CN })]
          })]
        })]
      })]
    }));

    // 后端实时图谱图片
    var imgBk = await imgRun(d["后端实时图谱图片"]);
    if (imgBk) {
      el.push(new Paragraph({ alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [imgBk] }));
    } else {
      el.push(imgPh("[后端实时图谱 - 未检测到图片]"));
    }

    // 后端图谱标题（斜体+灰色）
    el.push(new Paragraph({ alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 80 },
      children: [r("后端设备实时图谱（三相幅值最大的 C 相数值为"+KPDA.g(d["后端C相幅值(mV)"],"")+"mV，其余两相分别为"+KPDA.g(d["后端A相幅值(mV)"],"")+"mV、"+KPDA.g(d["后端B相幅值(mV)"],"")+"mV）",
        { size: 20, italic: true, color: "64748B", font: CN })] }));

    // 定位图谱（仅在有图时生成整段，无图则全部跳过）
    var hasLocImg = KPDA.hasVal(d["双端定位谱图图片"]);
    if (hasLocImg) {
      var imgLoc = await imgRun(d["双端定位谱图图片"]);
      if (imgLoc) {
        el.push(bar("定位图谱", false));
        el.push(new Paragraph({ alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
          children: [imgLoc] }));
        el.push(imgPh("双端定位谱图"));
      }
    }

    if (KPDA.hasVal(d["定位分析描述"])) {
      el.push(bar("定位分析", false));
      el.push(new Paragraph({ spacing: { before: 40, after: 80 }, indent: { firstLine: pts(24) },
        children: [r(KPDA.g(d["定位分析描述"],""), { size: 22, font: CN })] }));
    }

    // 疑似杆现场图（仅在有图时生成整段）
    var hasSceneImg = KPDA.hasVal(d["现场图片"]);
    if (hasSceneImg) {
      var imgScene = await imgRun(d["现场图片"]);
      if (imgScene) {
        el.push(bar("疑似杆现场图", false));
        el.push(new Paragraph({ alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
          children: [imgScene] }));
      }
    }

    if (KPDA.hasVal(d["结论"]) || KPDA.hasVal(d["检测结果分析"]) || KPDA.hasVal(d["疑似杆情况说明"]) || KPDA.hasVal(d["处理意见"]))
      el.push(bar("2）检测结果分析", true));

    if (KPDA.hasVal(d["结论"])) {
      el.push(bar("结论", false));
      el.push(new Paragraph({ spacing: { before: 40, after: 40 }, indent: { firstLine: pts(24) },
        children: [r(KPDA.g(d["结论"],""), { size: 22, bold: true, font: CN })] }));
    }
    if (KPDA.hasVal(d["检测结果分析"])) {
      el.push(bar("检测结果分析", false));
      KPDA.splitLines(d["检测结果分析"]).forEach(function(s) {
        if (s) el.push(new Paragraph({ spacing: { before: 0, after: 60 }, indent: { firstLine: pts(24) },
          children: [r(s, { size: 22, font: CN })] }));
      });
    }
    if (KPDA.hasVal(d["疑似杆情况说明"])) {
      el.push(bar("疑似杆情况说明", false));
      el.push(new Paragraph({ spacing: { before: 40, after: 60 }, indent: { firstLine: pts(24) },
        children: [r(KPDA.g(d["疑似杆情况说明"],""), { size: 22, font: CN })] }));
    }
    if (KPDA.hasVal(d["处理意见"])) {
      el.push(bar("处理意见", false));
      el.push(new Paragraph({ spacing: { before: 40, after: 40 }, indent: { firstLine: pts(24) },
        children: [r(KPDA.g(d["处理意见"],""), { size: 22, font: CN })] }));
    }
    if (KPDA.hasVal(d["验证描述"])) {
      el.push(bar("验证", false));
      KPDA.splitLines(d["验证描述"]).forEach(function(s) {
        if (s) el.push(new Paragraph({ spacing: { before: 0, after: 60 }, indent: { firstLine: pts(24) },
          children: [r(s, { size: 22, font: CN })] }));
      });
    }
    return el;
  }

  /* 插入隐形的 TC（Table of Contents）标记，供 TOC 域自动收集页码 */
  /* 估算每个缺陷占据的页数：基础 0.6 页 + 文本字数/800（封顶 +2.5）+ 图片 0.3/张 */
  function estimateDefectPages(d) {
    var pages = 0.6;  // 标题+基本KV+现场条件+试验条件
    var totalChars = 0;
    ["结论","检测结果分析","疑似杆情况说明","处理意见","验证描述","定位分析描述"]
      .forEach(function(k) { totalChars += String(KPDA.g(d[k],"")).replace(/\s+/g,"").length; });
    if (totalChars > 0) pages += Math.min(2.5, totalChars / 800);
    if (KPDA.hasVal(d["前端实时图谱图片"])) pages += 0.32;
    if (KPDA.hasVal(d["后端实时图谱图片"])) pages += 0.22;
    if (KPDA.hasVal(d["双端定位谱图图片"])) pages += 0.32;
    if (KPDA.hasVal(d["现场图片"])) pages += 0.22;
    return Math.max(1, Math.round(pages));
  }

  /* 目录行：2 列表格——标题(420pt 可换行) | 页码(51.6pt 右对齐)，彻底解决页码被截断 */
  function tocLine(title, page) {
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [pts(420), pts(51.6)],
      borders: NO_BORD_IN,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: NO_BORD,
            width: { size: pts(420), type: WidthType.DXA },
            margins: { top: 20, bottom: 20, left: 0, right: 20 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              spacing: { line: 320 },
              children: [r(title, { size: 20, font: CN })]
            })]
          }),
          new TableCell({
            borders: NO_BORD,
            width: { size: pts(51.6), type: WidthType.DXA },
            margins: { top: 20, bottom: 20, left: 0, right: 0 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { line: 320 },
              children: [r(String(page), { size: 20, font: CN })]
            })]
          })
        ]
      })]
    });
  }

  async function generateBlob(info, defects) {
    var ri = Object.assign({}, KPDA.DEFAULT_INFO, info || {});
    var list = defects || [];
    var sec = { page: { size: { width: pts(595.3), height: pts(841.9) },
              margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } };

    // 优先从内嵌资源（assets/img_assets.js）取 dataURL；无则尝试 fetch（兜底）
    function getAsset(name) {
      if (KPDA.IMG_ASSETS && KPDA.IMG_ASSETS[name]) return KPDA.IMG_ASSETS[name];
      return null;
    }
    var logoImg     = await imgRun(getAsset("LOGO_DATA_URL"),      { width: 76,  height: 19 });
    var principleImg = await imgRun(getAsset("PRINCIPLE_DATA_URL"), 560);
    var productImg  = await imgRun(getAsset("PRODUCT_DATA_URL"),   209);
    var topologyImg = await imgRun(getAsset("TOPOLOGY_DATA_URL"),  380);

    // 页眉：右上角 logo + 底部横线（每页）
    var hdrParaChildren = logoImg ? [logoImg]
                                 : [r("咸亨国际", { size: 22, bold: true, color: "C00000", font: CN })];
    var hdr = new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 1 } },
      children: hdrParaChildren
    })] });

    var ftr = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      r("第 ", { size: 18, color: "888888", font: CN }),
      new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "888888" }),
      r(" 页 / 共 ", { size: 18, color: "888888", font: CN }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: "888888" }),
      r(" 页", { size: 18, color: "888888", font: CN }),
    ]})] });

    // 封面
    var cover = [];
    // 顶部：标题 + 副标题
    cover.push(new Paragraph({ spacing: { before: 240, after: 0 }, indent: { firstLine: pts(8) },
      children: [r(KPDA.g(ri["标题"],"架空线双端局放检测报告"), { size: 30, bold: true, font: CN })] }));
    cover.push(new Paragraph({ spacing: { before: 200, after: 0 }, indent: { firstLine: pts(8) },
      children: [r(KPDA.g(ri["副标题"],""), { size: 30, bold: true, font: CN })] }));

    // 上方留白推 "检 测 报 告" 到页面中部
    for (var i=0;i<10;i++) cover.push(new Paragraph({ spacing: { before: 0, after: 60 }, children: [r("")] }));

    // "检 测 报 告" 满版蓝条（左右覆盖至页面边缘）+ 文字右对齐
    cover.push(new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { left: -MARGIN, right: -MARGIN },
      spacing: { before: 80, after: 80 },
      shading: shd("10407C"),
      children: [r("检 测 报 告  ", { size: 120, bold: true, color: "FFFFFF", font: CN })]
    }));

    // 下方留白推公司/日期到页面底部
    for (var j=0;j<10;j++) cover.push(new Paragraph({ children: [r("")] }));

    // 公司 + 日期（底部）
    cover.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      children: [r(KPDA.g(ri["出具单位"],"咸亨国际科技股份有限公司"), { size: 30, font: CN })] }));
    cover.push(new Paragraph({ alignment: AlignmentType.CENTER,
      children: [r(KPDA.g(ri["报告日期"],""), { size: 30, font: CN })] }));

    // 目录（内容高度估算自动计算页码）
    var toc = [];
    toc.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 320 },
      children: [r("目 录", { size: 36, bold: true, color: "2E75B6", font: CN })] }));
    toc.push(tocLine("一、检测目的", ""));
    toc.push(tocLine("二、架空线双端局放定位检测原理", ""));
    toc.push(tocLine("三、本次使用设备", ""));
    if (list.length) {
      toc.push(tocLine("四、缺陷综合检测报告", ""));
      list.forEach(function(d, i) {
        var raw = KPDA.g(d["结论"],"");
        var clean = String(raw).replace(/&(#?[a-z0-9]+);/gi, function(m, e) {
          return e[0]==='#' ? String.fromCharCode(parseInt(e.slice(1),10)||32) : ' ';
        }).replace(/\s+/g, ' ').trim();
        toc.push(tocLine((i+1)+". "+clean, ""));
      });
    }

    // 内容（一、二、三完整文本）
    var content = [];
    content.push(hdg("一、检测目的"));
    content.push(new Paragraph({ spacing: { before: 0, after: 80 }, indent: { firstLine: pts(24) },
      children: [r("根据《电力设备带电检测技术规范》（试行）以及《国家电网公司变电检测管理规定（试行）国网（运检3）829-2017》等相关规程和规定，电力行业通过巡检巡视，对变电站、线路等电气设备运行状态进行不停电监控、检测和故障诊断，以保障电气设备安全稳定运行和及时合理地安排检修工作。", { size: 22, font: CN })] }));
    content.push(new Paragraph({ spacing: { before: 0, after: 80 }, indent: { firstLine: pts(24) },
      children: [r("由于缺陷发生时，可能伴随发热、表面电晕放电、机械振动、电磁波、声波等不同介质，因此，在检测过程中，可利用不同传感器，对不同形式的介质进行多维度、全方位的检测，以判断缺陷问题，确定缺陷位置。", { size: 22, font: CN })] }));

    content.push(hdg("二、架空线双端局放定位检测原理"));
    if (principleImg) {
      content.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 },
        children: [principleImg] }));
    }
    content.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20, after: 40 },
      children: [r("架空线双端局放定位检测原理图", { size: 20, italic: true, color: "64748B", font: CN })] }));
    content.push(new Paragraph({ spacing: { before: 0, after: 80 }, indent: { firstLine: pts(24) },
      children: [r("如图所示，当架空线中间发生局部放电时，地线将产生电荷迁移，高频电流沿架空线以20cm/ns的速度向两侧传播，系统采用GPS授时方式与高速集成电路芯片同步，确保所测时间差值的误差在10纳秒以内，通过两端的时间差计算放电的具体位置。", { size: 22, font: CN })] }));

    // 三、本次使用设备：产品外观 + 系统拓扑图 双列
    content.push(hdg("三、本次使用设备"));
    var prodKids = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [r("产品外观", { size: 20, color: "64748B", font: CN })] })
    ];
    if (productImg) prodKids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 }, children: [productImg] }));
    var topoKids = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [r("系统拓扑图", { size: 20, color: "64748B", font: CN })] })
    ];
    if (topologyImg) topoKids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 }, children: [topologyImg] }));
    content.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: NO_BORD_IN,
      rows: [new TableRow({
        children: [
          new TableCell({ borders: NO_BORD, shading: shd(CELL_BG),
            width: { size: pts(180), type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER, children: prodKids }),
          new TableCell({ borders: NO_BORD, shading: shd(CELL_BG),
            width: { size: pts(290), type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER, children: topoKids })
        ]
      })]
    }));
    // 主要特性（直接列出，无 bar 标题）
    var bullets = [
      ["■支持中压架空线局部放电带电检测", "设备专为中压架空线路设计，可在不停电状态下实时监测局部放电情况，提前发现绝缘缺陷，避免停电检测带来的供电中断。"],
      ["■支持中压架空线局部放电带电精确定位", "采用双端检测原理，通过捕捉放电信号的传播时间差，实现对放电点的精确定位，精度可达0.5米或0.5%线缆长度。"],
      ["■自动生成立位谱图，直接给出放电位置及强度", "检测过程中自动生成立位谱图，直观显示放电位置和强度，便于快速分析和决策。"],
      ["■采用电池供电，使用方便", "设备采用电池供电，无需外部电源，适合户外或无电源环境。内置大容量电池可支持长时间连续工作，续航能力达3小时以上。"]
    ];
    bullets.forEach(function(b) {
      content.push(new Paragraph({ spacing: { before: 24, after: 2 }, indent: { firstLine: pts(24) },
        children: [r(b[0], { size: 20, bold: true, color: "0E3A5F", font: CN })] }));
      content.push(new Paragraph({ spacing: { before: 0, after: 14, line: 280 }, indent: { firstLine: pts(24) },
        children: [r(b[1], { size: 19, font: CN })] }));
    });

    if (!list.length) {
      content.push(new Paragraph({ children: [new PageBreak()] }));
      content.push(hdg("四、缺陷综合检测报告"));
      content.push(new Paragraph({ spacing: { before: 160, after: 80 },
        children: [r("(暂无缺陷记录)", { size: 22, color: "888888", font: CN })] }));
    } else {
      for (var k = 0; k < list.length; k++) {
        if (k > 0) content.push(new Paragraph({ children: [new PageBreak()] }));
        var block = await buildDefectBlock(k, list[k], k === 0);
        block.forEach(function(p) { content.push(p); });
      }
    }

    content.push(new Paragraph({ spacing: { before: 240, after: 80 },
      border: { top: { style: BorderStyle.SINGLE, size: 2, color: "AEB6BF", space: 1 } },
      children: [r("注：因无近距离及解体分析，结论为初步判断，请以最终解体或近距离确诊为准，合理检修。", { size: 18, italics: true, color: "888888", font: CN })] }));

    var doc = new Document({
      creator: "KPD-800T", title: "架空线双端局放检测报告",
      styles: { default: { document: { run: { font: CN } } } },
      sections: [
        { properties: sec, headers: { default: hdr }, children: cover },
        { properties: sec, headers: { default: hdr }, children: toc },
        { properties: sec, headers: { default: hdr }, footers: { default: ftr }, children: content }
      ]
    });
    var docBlob;
    try {
      docBlob = await Packer.toBlob(doc);
      docBlob = await fixDocxBlob(docBlob);
    } catch (e) {
      console.warn("[docx-client] post-process failed, fallback to raw blob:", e);
      if (!docBlob) docBlob = await Packer.toBlob(doc);
    }
    return docBlob;
  }

  W.KPDADocx = { generateBlob: generateBlob };
})();
