/* ==========================================================================
   app.js - 核心逻辑
   Excel解析 / 数据管理 / 报告HTML渲染 / 模板下载
   ========================================================================== */
(function () {
  "use strict";

  // === 常量 ===
  var HEADERS = [
    "序号", "线路名称", "结论",
    "故障相", "前端设备悬挂杆号", "前端实时图谱图片",
    "前端A相幅值(mV)", "前端B相幅值(mV)", "前端C相幅值(mV)",
    "后端设备悬挂杆号", "后端实时图谱图片",
    "后端A相幅值(mV)", "后端B相幅值(mV)", "后端C相幅值(mV)",
    "双端定位谱图图片",
    "定位分析描述", "检测结果分析", "疑似杆情况说明", "处理意见",
    "现场图片1", "现场图片2", "验证描述"
  ];

  var BASIC_INFO = [
    "标题", "副标题", "出具单位", "报告日期",
    "客户单位", "检测单位", "检测仪器",
    "检测时间", "检测人员", "天气", "温湿度"
  ];

  var DEFAULT_INFO = {
    "标题": "南网云南会泽待补供电所架空线双端局放检测报告",
    "副标题": "10kV架空线双端局放定位仪巡检",
    "出具单位": "咸亨国际科技股份有限公司",
    "报告日期": "二〇二六年八月",
    "客户单位": "南网云南会泽待补供电所",
    "线路名称": "10kV野马线",
    "检测单位": "咸亨国际科技股份有限公司",
    "检测仪器": "架空线双端局放定位仪",
    "检测时间": "2026.8.4",
    "检测人员": "刘奥",
    "天气": "晴",
    "温湿度": "26℃ / 未知"
  };

  // === 工具函数 ===
  function esc(s) {
    if (s == null) return "";
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function hasVal(v) { return v != null && String(v).trim() !== ""; }
  function num(v) { var x = Number(v); return Number.isFinite(x) ? x : null; }
  function g(v, fb) { return hasVal(v) ? v : (fb || ""); }
  function splitSent(s) { return s ? String(s).split(/[。！？；.!?;]/).filter(Boolean).map(function(s){return s.trim();}) : []; }
  // 按 Excel 单元格换行符 \n 拆分（保留用户原始排版）
  function splitLines(s) {
    if (s == null) return [];
    return String(s).split(/\r?\n/).map(function(t){return t.trim();}).filter(Boolean);
  }

  // === 状态 ===
  var state = { info: {}, defects: [], file: null };

  // === 示例数据 ===
  function getExampleData() {
    return {
      info: Object.assign({}, DEFAULT_INFO),
      defects: [
        {
          "序号":1, "线路名称":"10kV野马线", "结论":"10kV野马线T车路箐支线#16.9 C相绝缘子存在重度局部放电隐患",
          "故障相":"C", "前端设备悬挂杆号":"10kV野马线T车路箐支线#16.5",
          "前端A相幅值(mV)":7148,"前端B相幅值(mV)":6198,"前端C相幅值(mV)":7460,
          "后端设备悬挂杆号":"10kV野马线T车路箐支线#16.10",
          "后端A相幅值(mV)":6148,"后端B相幅值(mV)":6198,"后端C相幅值(mV)":7510,
          "定位分析描述":"根据定位谱图，故障位置距离#2设备（10号杆）约56m，初步判断为#9号杆。",
          "检测结果分析":"根据现场测试数据可知，C相数值相比其他两相显著偏高且放电特征一致（其余两相信号疑似为感应耦合的高频信号），且图谱符合局部放电特性；10kV野马线#16.5断开后端供电后，发现疑似局放信号骤降95%左右，因此判断疑似局放来源后端杆塔；后续在10kV野马线#16.5与#16.10进行双端定位时，发现疑似局放信号来源于 #16.9，信号表明此处位置处数值幅值最大且图谱特征极为明显。",
          "疑似杆情况说明":"到现场杆塔下人耳可听见明显放电声，肉眼观察无明显异常的现象，后续用其他远程拍摄设备发现绝缘子内部存在疑似穿孔、绝缘子捆扎线存在明显放电痕迹。",
          "处理意见":"放电强度较强，人耳可听明显放电声，现场已确定疑似故障位置，需及时符合验证，并进行缺陷故障处理（停电更换或带电左右更换）。",
          "验证描述": "经带电作业（或停电检修）流程，对设备定位的故障缺陷位置实施了精准消缺处理。现场拆解疑似缺陷绝缘子，可见明显穿孔击穿痕迹。采用兆欧表对该绝缘子进行绝缘电阻检测，实测绝缘电阻值为5MΩ。",
          "前端实时图谱图片": "", "后端实时图谱图片": "", "双端定位谱图图片": "", "现场图片1": "", "现场图片2": ""
        },
        {
          "序号":2, "线路名称":"10kV野马线", "结论":"10kV野马线T车路箐支线#16.29 C相绝缘子存在局部放电隐患",
          "故障相":"C", "前端设备悬挂杆号":"10kV野马线#16.29",
          "前端A相幅值(mV)":748,"前端B相幅值(mV)":698,"前端C相幅值(mV)":803,
          "后端设备悬挂杆号":"10kV野马线#16.25",
          "后端A相幅值(mV)":1148,"后端B相幅值(mV)":1020,"后端C相幅值(mV)":1510,
          "定位分析描述":"根据定位谱图，故障位置距离#1设备（3号杆）约8m，初步判断为#16.29。",
          "检测结果分析":"根据现场测试数据可知，C相数值相比其他两相显著偏高且放电特征一致（其余两相信号疑似为感应耦合的高频信号），且图谱符合局部放电特性；10kV野马线#16.29，在电源侧与负载侧分别挂载时，其电源侧信号略大于负载侧，因而在 #16.25（小号侧，电缆供电杆塔）进行定位；后续在10kV野马线#16.25与#16.29进行双端定位时，发现疑似局放信号来源于 #16.29。",
          "疑似杆情况说明":"人耳未听见放电声，肉眼观察无明显异常的现象，待进一步跟进分析，初步判断为绝缘子内部缺陷，后续用其他远程拍摄设备发现绝缘子内部存在疑似放电产生的烧蚀痕迹。",
          "处理意见":"放电强度轻中度，人耳不可听见明显放电声，针对疑似故障位置进行多方位确认，条件允许时可考虑进行更换。",
          "验证描述":"经带电作业（或停电检修）流程，对设备定位的故障缺陷位置实施了精准消缺处理。现场拆解疑似缺陷绝缘子，可见明显穿孔击穿痕迹。采用兆欧表对该绝缘子进行绝缘电阻检测，实测绝缘电阻值为10MΩ。",
          "前端实时图谱图片":"","后端实时图谱图片":"","双端定位谱图图片":"","现场图片":""
        }
      ]
    };
  }

  // === Excel 解析（客户端 SheetJS + DISPIMG 提取）===
  // 从 xlsx 的 cellimages.xml + media/ 中提取 WPS DISPIMG 内嵌图片，转 base64 dataURL
  // XML 命名空间: etc=cellImages, xdr=drawing, a=drawingml, r=relationships
  function extractDispimgImages(uint8) {
    return new Promise(function(resolve) {
      try {
        if (typeof JSZip === "undefined") return resolve({});
        var zip = new JSZip();
        zip.loadAsync(uint8).then(function(z) {
          if (!z.file("xl/cellimages.xml")) return resolve({});
          z.file("xl/cellimages.xml").async("string").then(function(ciXml) {
            var id2rid = {};
            var blockRe = /<etc:cellImage\b[^>]*>([\s\S]*?)<\/etc:cellImage>/g;
            var m;
            while ((m = blockRe.exec(ciXml))) {
              var body = m[1];
              var nameM = /<xdr:cNvPr\b[^>]*\sname\s*=\s*"([^"]+)"/.exec(body);
              var blipM = /<a:blip\b[^>]*\sr:embed\s*=\s*"([^"]+)"/.exec(body);
              if (nameM && blipM) id2rid[nameM[1]] = blipM[1];
            }
            return z.file("xl/_rels/cellimages.xml.rels").async("string").then(function(relsXml) {
              var rid2file = {};
              var relRe = /<Relationship\b[^>]*Id\s*=\s*"([^"]+)"[^>]*Target\s*=\s*"([^"]+)"/g;
              var r;
              while ((r = relRe.exec(relsXml))) rid2file[r[1]] = r[2];
              var sheetFiles = [];
              Object.keys(z.files).filter(function(n){ return /^xl\/worksheets\/sheet\d+\.xml$/i.test(n); }).forEach(function(n){
                if (z.file(n)) sheetFiles.push(n);
              });
              if (!sheetFiles.length) return resolve({});

              // 修复（2026-08-12）：只扫描「缺陷检测记录」工作表，避免残留旧数据 sheet
              // （如从旧文件复制的 Sheet1）中的 DISPIMG 图片按相同行列覆盖正确图片。
              // 找不到该工作表名时保持旧行为（扫描全部），向后兼容。
              var wbXmlFile = z.file("xl/workbook.xml");
              var filterPromise = wbXmlFile ? wbXmlFile.async("string").then(function(wbXml) {
                var sm = /<sheet\b[^>]*name="缺陷检测记录"[^>]*r:id="([^"]+)"/.exec(wbXml);
                if (!sm) return sheetFiles;
                var relsFile = z.file("xl/_rels/workbook.xml.rels");
                if (!relsFile) return sheetFiles;
                return relsFile.async("string").then(function(wbRels) {
                  var rm = new RegExp('Id="' + sm[1] + '"[^>]*Target="([^"]+)"').exec(wbRels);
                  if (!rm) return sheetFiles;
                  var t = rm[1].replace(/\\/g, "/");
                  if (t.indexOf("worksheets/") === -1) return sheetFiles;
                  var full = t.indexOf("xl/") === 0 ? t : "xl/" + t;
                  var hit = sheetFiles.filter(function(n){ return n.toLowerCase() === full.toLowerCase(); });
                  return hit.length ? hit : sheetFiles;
                });
              }) : Promise.resolve(sheetFiles);
              return filterPromise.then(function(sf) {
                sheetFiles = sf;
                if (!sheetFiles.length) return resolve({});

                // 修复竞态：先读完所有 sheet XML，再扫描 DISPIMG 收集图片任务
                var sheetReadPromises = sheetFiles.map(function(sn) {
                  return z.file(sn).async("string");
                });
                Promise.all(sheetReadPromises).then(function(sheets) {
                  var imgMap = {};
                  var pending = [];
                  var dispRe = /<c\s+r="([A-Z]+)(\d+)"[^>]*>(?:[^<]*<[^>]+>)?_xlfn\.DISPIMG\(&quot;([^&]+)&quot;\s*,\s*\d+\s*\)/g;
                  sheets.forEach(function(sxml) {
                    var dm;
                    dispRe.lastIndex = 0; // 加固：防止上一个 sheet 匹配残留 lastIndex 导致漏匹配
                    while ((dm = dispRe.exec(sxml))) {
                      // 修复闭包：用 let 块级作用域，每次迭代独立捕获
                      var col = colLetterToNum(dm[1]);
                      var row = parseInt(dm[2]);
                      var did = dm[3];
                      if (id2rid[did] && rid2file[id2rid[did]]) {
                        var target = rid2file[id2rid[did]];
                        var mediaPath = target.indexOf("media/") === 0 ? "xl/" + target : "xl/" + target;
                        if (z.file(mediaPath)) {
                          // IIFE 捕获当前迭代的 col/row/target/mediaPath，避免闭包共享
                          (function(c, rw, tgt, mp) {
                            pending.push(z.file(mp).async("base64").then(function(b64) {
                              var ext = (tgt.split(".").pop() || "png").toLowerCase();
                              var mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg"
                                       : ext === "gif" ? "image/gif" : "image/png";
                              var key = "row_" + rw;
                              if (!imgMap[key]) imgMap[key] = {};
                              imgMap[key][c] = "data:" + mime + ";base64," + b64;
                            }));
                          })(col, row, target, mediaPath);
                        }
                      }
                    }
                  });
                  // pending 此时已填满，安全等待
                  Promise.all(pending).then(function(){ resolve(imgMap); }).catch(function(){ resolve(imgMap); });
                }).catch(function(){ resolve({}); });
              }).catch(function(){ resolve({}); });
            });
          });
        }).catch(function(){ resolve({}); });
      } catch (e) { resolve({}); }
    });
  }
  function colLetterToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  function parseExcel(file) {
    return new Promise(function(resolve, reject) {
      var reader = new FileReader();
      reader.onload = function(e) {
        try {
          var data = new Uint8Array(e.target.result);
          var wb = XLSX.read(data, { type: "array" });
          var info = Object.assign({}, DEFAULT_INFO);
          var defects = [];

          // 读取报告基本信息
          var wsi = wb.Sheets["报告基本信息"];
          if (!wsi) { reject(new Error("未找到「报告基本信息」工作表，请使用下载的模板填写")); return; }
          if (wsi) {
            var arr = XLSX.utils.sheet_to_json(wsi, { header: 1 });
            arr.forEach(function(r) {
              var k = (r[0] || "").toString().trim();
              if (BASIC_INFO.indexOf(k) >= 0 && hasVal(r[1])) info[k] = String(r[1]).trim();
            });
          }

          // 读取缺陷检测记录
          var wsd = wb.Sheets["缺陷检测记录"];
          if (!wsd) { reject(new Error("未找到「缺陷检测记录」工作表，请使用下载的模板填写")); return; }
          if (wsd) {
            var arr2 = XLSX.utils.sheet_to_json(wsd, { header: 1 });
            if (arr2.length < 3) { reject(new Error("「缺陷检测记录」未识别到数据行：表头需在第 2 行、数据从第 3 行开始")); return; }
            if (arr2.length >= 3) {
              var headers = arr2[1];
              for (var r = 2; r < arr2.length; r++) {
                var row = arr2[r];
                if (!row || row.every(function(c){return !hasVal(c);})) continue;
                var obj = {};
                headers.forEach(function(h, i) {
                  if (h && i < row.length) {
                    var v = row[i];
                    obj[String(h).trim()] = v;
                  }
                });
                if (Object.keys(obj).length > 0) defects.push(obj);
              }
            }
          }

          // 异步提取 DISPIMG 内嵌图片（前端静态版无需后端）
          extractDispimgImages(data).then(function(imgMap) {
            // XML 行号与 SheetJS 行列对应：第3行=第1条数据 (i=0)
            defects.forEach(function(defect, i) {
              var rowKey = "row_" + (i + 3);
              if (imgMap[rowKey]) {
                var cellArr = imgMap[rowKey];
                headers.forEach(function(h, j) {
                  if (h && h.indexOf("图片") >= 0) {
                    var colNum = j + 1;
                    if (cellArr[colNum]) defect[h] = cellArr[colNum];
                  }
                });
              }
            });
                        // 向后兼容：旧模板的"现场图片"字段映射为"现场图片1"
            defects.forEach(function(def) {
              if (!KPDA.hasVal(def["现场图片1"]) && KPDA.hasVal(def["现场图片"])) {
                def["现场图片1"] = def["现场图片"]; def["现场图片"] = "";
              }
            });
            resolve({ info: info, defects: defects });
          }).catch(function(){             // 向后兼容：旧模板的"现场图片"字段映射为"现场图片1"
            defects.forEach(function(def) {
              if (!KPDA.hasVal(def["现场图片1"]) && KPDA.hasVal(def["现场图片"])) {
                def["现场图片1"] = def["现场图片"]; def["现场图片"] = "";
              }
            });
            resolve({ info: info, defects: defects }); });
        } catch(err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // === 报告 HTML 渲染 ===
  function buildReportHtml(info, defects) {
    var d = info || DEFAULT_INFO;
    var list = defects || [];
    var html = "";

    // 封面
    html += '<div class="report-cover">';
    html += '<div style="font-size:14px;font-weight:700;">' + esc(g(d["标题"], "")) + '</div>';
    html += '<div style="font-size:14px;font-weight:700;margin-top:8px;">' + esc(g(d["副标题"], "")) + '</div>';
    html += '<div class="banner">检 测 报 告</div>';
    html += '<div class="org"><span>' + esc(g(d["出具单位"], "")) + '</span><span>' + esc(g(d["报告日期"], "")) + '</span></div>';
    html += '</div>';
    html += '<h2 class="report-h">检测基本信息</h2>';
    html += '<table class="report-table"><tbody>';
    html += kvRow("客户单位", d["客户单位"], "检测单位", d["检测单位"]);
    html += kvRow("检测时间", d["检测时间"], "检测人员", d["检测人员"]);
    html += kvRow("天气", d["天气"], "温湿度", d["温湿度"]);
    html += '</tbody></table>';

    // 目录
    html += '<div class="report-toc"><h2>目 录</h2>';
    var entries = [["一、检测目的","3"],["二、架空线双端局放定位检测原理","3"],["三、本次使用设备","4"]];
    if (list.length) entries.push(["四、缺陷综合检测报告","5"]);
    list.forEach(function(_, i) { entries.push([(i+1) + "、" + esc(g(list[i]["结论"], "")).slice(0,30), String(5+i*5)]); });
    entries.forEach(function(e) {
      html += '<div class="toc-line"><span>' + esc(e[0]) + '</span><span></span><span>' + esc(e[1]) + '</span></div>';
    });
    html += '</div>';

    // 一、检测目的
    html += '<h2 class="report-h">一、检测目的</h2>';
    html += '<p class="report-para">根据《电力设备带电检测技术规范》（试行）以及《国家电网公司变电检测管理规定（试行）国网（运检3）829-2017》等相关规程和规定，电力行业通过巡检巡视，对变电站、线路等电气设备运行状态进行不停电监控、检测和故障诊断，以保障电气设备安全稳定运行和及时合理地安排检修工作。</p>';
    html += '<p class="report-para">由于缺陷发生时，可能伴随发热、表面电晕放电、机械振动、电磁波、声波等不同介质，因此，在检测过程中，可利用不同传感器，对不同形式的介质进行多维度、全方位的检测，以判断缺陷问题，确定缺陷位置。</p>';

    // 二、检测原理
    html += '<h2 class="report-h">二、架空线双端局放定位检测原理</h2>';
    html += '<div class="report-img-ph">架空线双端局放定位检测原理图</div>';
    html += '<p class="report-para">如图所示，当架空线中间发生局部放电时，地线将产生电荷迁移，高频电流沿架空线以20cm/ns的速度向两侧传播，系统采用GPS授时方式与高速集成电路芯片同步，确保所测时间差值的误差在10纳秒以内，通过两端的时间差计算放电的具体位置。</p>';

    // 三、使用设备
    html += '<h2 class="report-h">三、本次使用设备</h2>';
    html += '<div class="report-blue-bar">KPD-800T 双端局放定位仪</div>';
    html += '<div class="report-blue-bar sub">主要特性</div>';
    var bullets = [
      ["■支持中压架空线局部放电带电检测","设备专为中压架空线路设计，可在不停电状态下实时监测局部放电情况，提前发现绝缘缺陷，避免停电检测带来的供电中断。"],
      ["■支持中压架空线局部放电带电精确定位","采用双端检测原理，通过捕捉放电信号的传播时间差，实现对放电点的精确定位，精度可达0.5米或0.5%线缆长度。"],
      ["■自动生成立位谱图，直接给出放电位置及强度","检测过程中自动生成立位谱图，直观显示放电位置和强度，便于快速分析和决策。"],
      ["■采用电池供电，使用方便","设备采用电池供电，无需外部电源，适合户外或无电源环境。内置大容量电池可支持长时间连续工作，续航能力达3小时以上。"]
    ];
    bullets.forEach(function(b) {
      html += '<p class="report-para" style="font-weight:700;">' + esc(b[0]) + '</p>';
      html += '<p class="report-para" style="font-size:10.5px;">' + esc(b[1]) + '</p>';
    });

    // 四、缺陷综合检测报告
    html += '<h2 class="report-h">四、缺陷综合检测报告</h2>';
    list.forEach(function(defect, i) { html += buildDefectHtml(i+1, defect); });

    // 结尾
    html += '<div class="report-end">注：因无近距离及解体分析，结论为初步判断，请以最终解体或近距离确诊为准，合理检修。</div>';

    return html;
  }

  function imgHtml(src, ph) {
    var s = g(src, "");
    if (hasVal(s) && s.indexOf("data:image") === 0) {
      return '<div class="report-img"><img src="' + s + '" alt="' + esc(ph) + '"/></div>';
    }
    return '<div class="report-img-ph">' + esc(ph) + '</div>';
  }

  function buildDefectHtml(idx, d) {
    d = d || {};
    var h = "";
    h += '<div class="report-blue-bar sub">' + idx + '、' + esc(g(d["线路名称"], "——")) + ' / ' + esc(g(d["结论"], "")) + '</div>';

    // 双端试验
    h += '<h3 class="report-h2">1）双端局放现场试验条件</h3>';
    h += '<p class="report-para" style="text-indent:0;font-weight:600;">前端设备悬挂杆号：' + esc(g(d["前端设备悬挂杆号"], "——")) + '</p>';
    h += imgHtml(d["前端实时图谱图片"], "前端设备实时图谱");
    h += '<p class="report-para" style="text-indent:0;font-weight:600;">后端设备悬挂杆号：' + esc(g(d["后端设备悬挂杆号"], "——")) + '</p>';
    h += imgHtml(d["后端实时图谱图片"], "后端设备实时图谱");
    h += '<div class="report-blue-bar">定位图谱</div>';
    h += imgHtml(d["双端定位谱图图片"], "双端定位谱图");

    if (hasVal(d["定位分析描述"])) {
      h += '<div class="report-blue-bar sub">定位分析</div>';
      h += '<p class="report-para">' + esc(g(d["定位分析描述"], "")) + '</p>';
    }

      var hasS1 = hasVal(d["现场图片1"]);
      var hasS2 = hasVal(d["现场图片2"]);
      if (hasS1 || hasS2) {
        h += '<div class="report-blue-bar">现场疑似缺陷图片</div>';
        if (hasS1 && hasS2) {
          h += '<div class="report-img-row">';
          h += '<div class="report-img-half">' + imgHtml(d["现场图片1"], "绝缘子穿孔") + '<div class="report-img-caption">绝缘子穿孔</div></div>';
          h += '<div class="report-img-half">' + imgHtml(d["现场图片2"], "捆扎线破损") + '<div class="report-img-caption">捆扎线破损</div></div>';
          h += '</div>';
        } else if (hasS1) {
          h += imgHtml(d["现场图片1"], "现场图片（绝缘子穿孔）");
        } else if (hasS2) {
          h += imgHtml(d["现场图片2"], "现场图片（捆扎线破损）");
        }
      }

    if (hasVal(d["结论"]) || hasVal(d["检测结果分析"]) || hasVal(d["疑似杆情况说明"]) || hasVal(d["处理意见"])) {
      h += '<div class="report-blue-bar">2）检测结果分析</div>';
    }
    if (hasVal(d["结论"])) { h += '<div class="report-blue-bar sub">结论</div><p class="report-para" style="font-weight:700;">' + esc(d["结论"]) + '</p>'; }
    if (hasVal(d["检测结果分析"])) {
      h += '<div class="report-blue-bar sub">检测结果分析</div>';
      splitLines(d["检测结果分析"]).forEach(function(s) { h += '<p class="report-para" style="white-space:pre-wrap;">' + esc(s) + '</p>'; });
    }
    if (hasVal(d["疑似杆情况说明"])) { h += '<div class="report-blue-bar sub">疑似杆情况说明</div><p class="report-para">' + esc(d["疑似杆情况说明"]) + '</p>'; }
    if (hasVal(d["处理意见"])) {
      h += '<div class="report-blue-bar sub">处理意见</div><p class="report-para" style="font-weight:700;">' + esc(d["处理意见"]) + '</p>';
    }
    if (hasVal(d["验证描述"])) {
      h += '<div class="report-blue-bar sub">验证</div>';
      splitLines(d["验证描述"]).forEach(function(s) { h += '<p class="report-para" style="white-space:pre-wrap;">' + esc(s) + '</p>'; });
    }

    return h;
  }

  function kvRow(l1, v1, l2, v2) {
    return '<tr><th>' + esc(l1) + '</th><td>' + esc(g(v1, "——")) + '</td><th>' + esc(l2) + '</th><td>' + esc(g(v2, "——")) + '</td></tr>';
  }

  // === 缺陷卡片渲染 ===
  function renderDefectCards(defects, container) {
    if (!defects || !defects.length) {
      container.innerHTML = '<div class="empty-state">暂无数据，请先导入 Excel 或加载示例</div>';
      return;
    }
    var html = "";
    defects.forEach(function(d, i) {
      var idx = i + 1;
      var line = g(d["线路名称"], "——");
      var conclusion = g(d["结论"], "——");
      var ftTower = g(d["前端设备悬挂杆号"], "——");
      var bkTower = g(d["后端设备悬挂杆号"], "——");

      html += '<div class="defect-card">';
      html += '<div class="dc-header" role="button" tabindex="0" aria-expanded="false" onclick="var b=this.nextElementSibling;var c=this.querySelector(\'.dc-chevron\');b.classList.toggle(\'expanded\');c.classList.toggle(\'open\');this.setAttribute(\'aria-expanded\', b.classList.contains(\'expanded\'))" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();this.click()}">';
      html += '<div class="dc-index">' + idx + '</div>';
      html += '<div class="dc-info"><div class="dc-title">' + esc(conclusion).slice(0,60) + '</div>';
      html += '<div class="dc-meta">';
      html += '<span>' + esc(line).slice(0,20) + '</span>';
      html += '<span>前端: ' + esc(ftTower).slice(0,25) + '</span>';
      html += '<span>后端: ' + esc(bkTower).slice(0,25) + '</span>';
      html += '</div></div>';
      html += '<svg class="dc-chevron" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>';
      html += '</div>';
      html += '<div class="dc-body">';

      // 三相幅值
      html += '<div class="dc-section"><div class="dc-section-title">三相幅值 (mV)</div>';
      html += '<div class="dc-kv">';
      html += dcKV("故障相", g(d["故障相"],"—"));
      html += dcKV("前端A/B/C", (g(d["前端A相幅值(mV)"],"—")+" / "+g(d["前端B相幅值(mV)"],"—")+" / "+g(d["前端C相幅值(mV)"],"—")));
      html += dcKV("后端A/B/C", (g(d["后端A相幅值(mV)"],"—")+" / "+g(d["后端B相幅值(mV)"],"—")+" / "+g(d["后端C相幅值(mV)"],"—")));
      html += '</div></div>';

      // 文本内容
      if (hasVal(d["定位分析描述"])) html += dcSection("定位分析", d["定位分析描述"]);
      if (hasVal(d["检测结果分析"])) html += dcSection("检测结果分析", d["检测结果分析"]);
      if (hasVal(d["疑似杆情况说明"])) html += dcSection("疑似杆情况说明", d["疑似杆情况说明"]);
      if (hasVal(d["处理意见"])) html += dcSection("处理意见", d["处理意见"]);
      if (hasVal(d["验证描述"])) html += dcSection("验证", d["验证描述"]);

      html += '</div></div>';
    });
    container.innerHTML = html;
  }

  function dcKV(k, v) { return '<div class="dc-kv-item"><span class="dc-kv-key">' + esc(k) + '</span><span class="dc-kv-val">' + esc(g(v, "—")) + '</span></div>'; }
  function dcSection(title, content) { return '<div class="dc-section"><div class="dc-section-title">' + esc(title) + '</div><div class="dc-text-block">' + esc(content) + '</div></div>'; }

  // === Excel 模板下载（客户端生成）===
  function downloadTemplate() {
    if (!window.XLSX) throw new Error("XLSX 库未加载");

    // ── 模板视觉令牌（与报告主视觉一致）──
    var BORDER = "D0D7E5";   // 细边框浅灰
    var NAVY = "1A5276";     // 深蓝：标题/表头
    var NAVY_LIGHT = "EBF5FB"; // 浅蓝：标签/分节
    var ORANGE = "F0531C";   // 品牌橙：故障相列表头
    var ZEBRA = "F7FAFD";    // 斑马纹浅蓝
    var IMG_HINT = "FFF6E5"; // 图片列浅黄提示
    var WHITE = "FFFFFF";

    // 单元格设值 + 统一样式（边框默认全加）
    function styleCell(ws, ref, value, o) {
      o = o || {};
      var addr = XLSX.utils.encode_cell(XLSX.utils.decode_cell(ref));
      ws[addr] = { t: typeof value === "number" ? "n" : "s", v: value };
      var s = {};
      if (o.font) s.font = o.font;
      if (o.fill) s.fill = { patternType: "solid", fgColor: { rgb: o.fill } };
      if (o.align) s.alignment = o.align;
      if (o.border !== false) {
        s.border = {
          top: { style: "thin", color: { rgb: BORDER } },
          bottom: { style: "thin", color: { rgb: BORDER } },
          left: { style: "thin", color: { rgb: BORDER } },
          right: { style: "thin", color: { rgb: BORDER } }
        };
      }
      if (Object.keys(s).length) ws[addr].s = s;
      return ws[addr];
    }
    function bFont(bold, color, sz) {
      var f = { name: "微软雅黑" };
      if (bold) f.bold = true;
      if (color) f.color = { rgb: color };
      if (sz) f.sz = sz;
      return f;
    }
    function cAlign() { return { horizontal: "center", vertical: "center" }; }
    function lAlign() { return { horizontal: "left", vertical: "center" }; }

    var wb = XLSX.utils.book_new();

    // ═══ Sheet 1: 报告基本信息 ═══
    var ws1 = XLSX.utils.aoa_to_sheet([
      ["报告基本信息（每份报告填写一次）"],
      ["", ""],
      ["标题", DEFAULT_INFO["标题"]],
      ["副标题", DEFAULT_INFO["副标题"]],
      ["出具单位", DEFAULT_INFO["出具单位"]],
      ["报告日期", DEFAULT_INFO["报告日期"]],
      ["客户单位", DEFAULT_INFO["客户单位"]],
      ["检测单位", DEFAULT_INFO["检测单位"]],
      ["线路名称", DEFAULT_INFO["线路名称"]],
      ["检测仪器", DEFAULT_INFO["检测仪器"]],
      ["检测时间", DEFAULT_INFO["检测时间"]],
      ["检测人员", DEFAULT_INFO["检测人员"]],
      ["天气", DEFAULT_INFO["天气"]],
      ["温湿度", DEFAULT_INFO["温湿度"]]
    ]);
    ws1["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    ws1["!cols"] = [{ wch: 18 }, { wch: 56 }];
    // 标题行：深蓝底白字居中加粗
    styleCell(ws1, "A1", "报告基本信息（每份报告填写一次）", { font: bFont(true, WHITE, 13), fill: NAVY, align: cAlign(), border: false });
    // 标签列：浅蓝底深蓝加粗；值列：白底深灰
    ["A3","A4","A5","A6","A7","A8","A9","A10","A11","A12","A13","A14"].forEach(function(r) {
      if (!ws1[r]) return;
      styleCell(ws1, r, ws1[r].v, { font: bFont(true, "1A5276", 11), fill: NAVY_LIGHT, align: lAlign() });
    });
    ["B3","B4","B5","B6","B7","B8","B9","B10","B11","B12","B13","B14"].forEach(function(r) {
      if (!ws1[r]) return;
      styleCell(ws1, r, ws1[r].v, { font: bFont(false, "333333", 11), fill: WHITE, align: lAlign() });
    });
    ws1["!rows"] = [{ hpt: 30 }];
    for (var r1 = 1; r1 < 14; r1++) ws1["!rows"].push({ hpt: 25 });
    XLSX.utils.book_append_sheet(wb, ws1, "报告基本信息");

    // ═══ Sheet 2: 缺陷检测记录 ═══
    var headerRow = HEADERS.slice();
    var aoa2 = [
      ["缺陷检测记录（每行 = 一个缺陷点位，删除示例后填实数据）"],
      headerRow
    ];
    var ws2 = XLSX.utils.aoa_to_sheet(aoa2);
    ws2["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headerRow.length - 1 } }];
    var sample = getExampleData().defects;
    sample.forEach(function(d) {
      var row = headerRow.map(function(h) { return g(d[h], ""); });
      XLSX.utils.sheet_add_aoa(ws2, [row], { origin: -1 });
    });

    // 列宽（按列类型定制：序号窄 / 文本宽 / 幅值中 / 图片中）
    var colW = { 0: 6, 1: 14, 2: 36, 3: 10, 4: 18, 5: 18, 6: 13, 7: 13, 8: 13, 9: 18, 10: 18, 11: 13, 12: 13, 13: 13, 14: 18, 15: 30, 16: 42, 17: 28, 18: 28, 19: 14, 20: 14, 21: 36 };
    ws2["!cols"] = headerRow.map(function(h, i) { return { wch: colW[i] || 14 }; });

    // 标题行（第1行）：深蓝白字居中
    styleCell(ws2, "A1", ws2["A1"].v, { font: bFont(true, WHITE, 12), fill: NAVY, align: cAlign(), border: false });
    // 表头（第2行）：深蓝白字加粗居中；「故障相」列品牌橙高亮
    for (var c = 0; c < headerRow.length; c++) {
      var addr = XLSX.utils.encode_cell({ r: 1, c: c });
      var isFP = headerRow[c] === "故障相";
      styleCell(ws2, addr, ws2[addr].v, {
        font: bFont(true, WHITE, 11), fill: isFP ? ORANGE : NAVY, align: cAlign()
      });
    }
    // 示例数据行：斑马纹交替 + 图片列浅黄提示 + 数值列居中
    for (var r = 2; r < 2 + sample.length; r++) {
      for (var c2 = 0; c2 < headerRow.length; c2++) {
        var addr2 = XLSX.utils.encode_cell({ r: r, c: c2 });
        if (!ws2[addr2]) continue;
        var isImg = String(headerRow[c2]).indexOf("图片") >= 0;
        var isNum = /幅值/.test(headerRow[c2]);
        styleCell(ws2, addr2, ws2[addr2].v, {
          font: bFont(false, "333333", 10.5),
          fill: isImg ? IMG_HINT : ((r - 2) % 2 === 0 ? ZEBRA : WHITE),
          align: isNum ? cAlign() : lAlign()
        });
      }
    }
    // 行高 + 冻结窗格（冻结前 2 行 + A 列，滚动时表头始终可见）
    ws2["!rows"] = [{ hpt: 30 }, { hpt: 32 }];
    for (var r2 = 2; r2 < 2 + sample.length; r2++) ws2["!rows"].push({ hpt: 26 });
    ws2["!freeze"] = { x: 1, y: 2 };
    XLSX.utils.book_append_sheet(wb, ws2, "缺陷检测记录");

    // ═══ Sheet 3: 填写说明 ═══
    var notes = [
      ["【填写说明】"],
      [""],
      ["1. 报告基本信息：每份报告填写一次，决定报告封面和页脚的信息。"],
      [""],
      ["2. 缺陷检测记录："],
      ["   - 每行代表一个检测到的缺陷点位"],
      ["   - 示例数据来源于 2026.8.4 会泽待补供电所现场检测，供参考"],
      ["   - 删除示例行后填入实际数据（在示例下方继续添加新行即可）"],
      [""],
      ["3. 关键字段说明："],
      ["   - 检测结果分析：完整描述文字，系统会按段落拆分"],
      ["   - 验证描述（含绝缘电阻值MΩ）：验证环节的完整描述，需包含实测绝缘电阻值"],
      ["   - 定位分析描述：填入定位谱图分析文字"],
      ["   - 前端/后端 A、B、C 相幅值(mV)：现场实测三相幅值"],
      ["   - 故障相：缺陷对应的相别（A/B/C 等），自动填入图谱说明文字"],
      [""],
      ["4. 图片路径列（可选）："],
      ["   - 填入本地图片绝对路径，留空 = 显示占位符"],
      ["   - 支持 jpg/jpeg/png/bmp/gif 或 WPS 内嵌图片（DISPIMG）"],
      [""],
      ["5. 填写完成后，上传此 Excel 即可自动生成 DOCX 报告。"],
      [""],
      ["6. 列顺序严格固定，请勿增删或调换列，以保证自动填充准确。"]
    ];
    var ws3 = XLSX.utils.aoa_to_sheet(notes);
    ws3["!cols"] = [{ wch: 100 }];
    // 主标题：深蓝底白字
    styleCell(ws3, "A1", ws3["A1"].v, { font: bFont(true, WHITE, 13), fill: NAVY, align: lAlign(), border: false });
    // 分节标题（数字.开头）：浅蓝底深蓝加粗；正文：白底深灰
    var rows3 = [{ hpt: 30 }];
    for (var i = 1; i < notes.length; i++) {
      var addr3 = "A" + (i + 1);
      if (!ws3[addr3]) { rows3.push({ hpt: 10 }); continue; }
      var t3 = String(ws3[addr3].v);
      if (/^\d+\./.test(t3)) {
        styleCell(ws3, addr3, t3, { font: bFont(true, "1A5276", 11), fill: NAVY_LIGHT, align: lAlign() });
        rows3.push({ hpt: 24 });
      } else {
        styleCell(ws3, addr3, t3, { font: bFont(false, "555555", 10.5), fill: WHITE, align: lAlign() });
        rows3.push({ hpt: 21 });
      }
    }
    ws3["!rows"] = rows3;
    XLSX.utils.book_append_sheet(wb, ws3, "填写说明");

    // 写入并下载：SheetJS 社区版 write 不输出样式 → 用 JSZip 注入 styles.xml 与单元格样式索引
    var wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    // v78 修复：JSZip.generateAsync({type:"array"}) 返回的是普通数字数组，
    // new Blob([普通数组]) 会被 toString() 成 "80,75,3,4,..." 逗号文本，导致下载的模板损坏。
    // 这里统一把输入归一化成 Uint8Array，保证二进制内容原样写入 Blob。
    var deliver = function(arr) {
      var bin;
      if (arr instanceof ArrayBuffer) bin = new Uint8Array(arr);
      else if (ArrayBuffer.isView(arr)) bin = new Uint8Array(arr.buffer, arr.byteOffset, arr.byteLength);
      else if (Array.isArray(arr)) bin = Uint8Array.from(arr);
      else bin = new Uint8Array(arr || []);
      var blob = new Blob([bin], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = "架空线双端局放检测_数据填写模板.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
    };
    if (typeof JSZip === "undefined") { deliver(wbout); return; }
    injectTemplateStyles(wbout, headerRow, sample.length, notes).then(deliver)["catch"](function() { deliver(wbout); });
  }

  // === 模板样式注入（SheetJS CE 的 write 不输出样式，用 JSZip 后处理补样式）===
  function buildTemplateStylesXml() {
    var X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    X += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    // fonts: 0 默认 | 1 雅黑13粗白 | 2 雅黑11粗深蓝 | 3 雅黑11深灰 | 4 雅黑12粗白 | 5 雅黑11粗白 | 6 雅黑10.5深灰 | 7 雅黑10.5浅灰
    X += '<fonts count="8">';
    X += '<font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
    X += '<font><sz val="13"/><b/><color rgb="FFFFFFFF"/><name val="微软雅黑"/></font>';
    X += '<font><sz val="11"/><b/><color rgb="FF1A5276"/><name val="微软雅黑"/></font>';
    X += '<font><sz val="11"/><color rgb="FF333333"/><name val="微软雅黑"/></font>';
    X += '<font><sz val="12"/><b/><color rgb="FFFFFFFF"/><name val="微软雅黑"/></font>';
    X += '<font><sz val="11"/><b/><color rgb="FFFFFFFF"/><name val="微软雅黑"/></font>';
    X += '<font><sz val="10.5"/><color rgb="FF333333"/><name val="微软雅黑"/></font>';
    X += '<font><sz val="10.5"/><color rgb="FF555555"/><name val="微软雅黑"/></font>';
    X += '</fonts>';
    // fills: 0 none | 1 gray125(惯例) | 2 深蓝 | 3 浅蓝 | 4 白 | 5 橙 | 6 斑马 | 7 浅黄
    X += '<fills count="8">';
    X += '<fill><patternFill patternType="none"/></fill>';
    X += '<fill><patternFill patternType="gray125"/></fill>';
    X += '<fill><patternFill patternType="solid"><fgColor rgb="FF1A5276"/><bgColor indexed="64"/></patternFill></fill>';
    X += '<fill><patternFill patternType="solid"><fgColor rgb="FFEBF5FB"/><bgColor indexed="64"/></patternFill></fill>';
    X += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>';
    X += '<fill><patternFill patternType="solid"><fgColor rgb="FFF0531C"/><bgColor indexed="64"/></patternFill></fill>';
    X += '<fill><patternFill patternType="solid"><fgColor rgb="FFF7FAFD"/><bgColor indexed="64"/></patternFill></fill>';
    X += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF6E5"/><bgColor indexed="64"/></patternFill></fill>';
    X += '</fills>';
    // borders: 0 无 | 1 thin 灰
    X += '<borders count="2">';
    X += '<border><left/><right/><top/><bottom/><diagonal/></border>';
    X += '<border><left style="thin"><color rgb="FFD0D7E5"/></left><right style="thin"><color rgb="FFD0D7E5"/></right><top style="thin"><color rgb="FFD0D7E5"/></top><bottom style="thin"><color rgb="FFD0D7E5"/></bottom><diagonal/></border>';
    X += '</borders>';
    X += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    // cellXfs 索引：0 默认 | 1 深蓝白粗13居中(Sheet1标题) | 2 浅蓝深蓝粗左+边框(标签/分节) | 3 白底深灰左+边框(Sheet1值)
    //            4 深蓝白粗12居中无边框(Sheet2/3标题) | 5 深蓝白粗11居中+边框(表头) | 6 橙白粗11居中+边框(故障相表头)
    //            7 斑马深灰左+边框 | 8 白深灰左+边框 | 9 浅黄深灰左+边框(图片列) | 10 斑马深灰中+边框 | 11 白深灰中+边框 | 12 浅灰正文左+边框
    X += '<cellXfs count="13">';
    X += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    X += '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="4" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="5" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
    X += '<xf numFmtId="0" fontId="7" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>';
    X += '</cellXfs>';
    X += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    X += '</styleSheet>';
    return X;
  }

  function injectTemplateStyles(wbout, headerRow, sampleLen, notes) {
    return JSZip.loadAsync(wbout).then(function(z) {
      z.file("xl/styles.xml", buildTemplateStylesXml());
      var setS = function(sxml, ref, idx) {
        return sxml.replace(new RegExp('(<c r="' + ref + '")(\s+s="\d+")?'), '$1 s="' + idx + '"');
      };
      var jobs = [];
      // Sheet1：标题深蓝居中；标签浅蓝加粗；值白底
      jobs.push(z.file("xl/worksheets/sheet1.xml").async("string").then(function(sxml) {
        var out = sxml;
        out = setS(out, "A1", 1);
        for (var r = 3; r <= 14; r++) { out = setS(out, "A" + r, 2); out = setS(out, "B" + r, 3); }
        z.file("xl/worksheets/sheet1.xml", out);
      }));
      // Sheet2：标题行；表头深蓝（故障相橙）；数据斑马纹/白/图片浅黄/数值居中；冻结前2行+A列
      jobs.push(z.file("xl/worksheets/sheet2.xml").async("string").then(function(sxml) {
        var out = sxml;
        out = setS(out, "A1", 4);
        for (var c = 0; c < headerRow.length; c++) {
          var ref = XLSX.utils.encode_cell({ r: 1, c: c });
          out = setS(out, ref, headerRow[c] === "故障相" ? 6 : 5);
        }
        for (var r = 2; r < 2 + sampleLen; r++) {
          for (var c2 = 0; c2 < headerRow.length; c2++) {
            var ref2 = XLSX.utils.encode_cell({ r: r, c: c2 });
            var h = headerRow[c2];
            var isImg = String(h).indexOf("图片") >= 0;
            var isNum = /幅值/.test(h);
            var zebra = (r - 2) % 2 === 0;
            var idx = isImg ? 9 : (isNum ? (zebra ? 10 : 11) : (zebra ? 7 : 8));
            out = setS(out, ref2, idx);
          }
        }
        // 冻结窗格：冻结前 2 行 + A 列（表头滚动时始终可见）
        out = out.replace(/(<sheetView[^>]*)\/>/, '$1><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"\/><\/sheetView>');
        z.file("xl/worksheets/sheet2.xml", out);
      }));
      // Sheet3：标题深蓝；分节浅蓝加粗；正文浅灰
      jobs.push(z.file("xl/worksheets/sheet3.xml").async("string").then(function(sxml) {
        var out = sxml;
        out = setS(out, "A1", 4);
        for (var i = 0; i < notes.length; i++) {
          var r = i + 1;
          var t = String((notes[i] || [""])[0] || "");
          if (/^\d+\./.test(t)) out = setS(out, "A" + r, 2);
          else if (t.trim() !== "" && r !== 1) out = setS(out, "A" + r, 12);
        }
        z.file("xl/worksheets/sheet3.xml", out);
      }));
      return Promise.all(jobs).then(function() { return z.generateAsync({ type: "arraybuffer" }); });
    });
  }


  // === 暴露 ===
  window.KPDA = {
    HEADERS: HEADERS, BASIC_INFO: BASIC_INFO, DEFAULT_INFO: DEFAULT_INFO,
    state: state, getExampleData: getExampleData,
    parseExcel: parseExcel, buildReportHtml: buildReportHtml, renderDefectCards: renderDefectCards,
    downloadTemplate: downloadTemplate,
    esc: esc, hasVal: hasVal, num: num, g: g, splitSent: splitSent, splitLines: splitLines
  };
})();


