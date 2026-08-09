/* ==========================================================================
   app.js - 核心逻辑
   Excel解析 / 数据管理 / 报告HTML渲染 / 模板下载
   ========================================================================== */
(function () {
  "use strict";

  // === 常量 ===
  var HEADERS = [
    "序号", "线路名称", "结论",
    "前端设备悬挂杆号", "前端实时图谱图片",
    "前端A相幅值(mV)", "前端B相幅值(mV)", "前端C相幅值(mV)",
    "后端设备悬挂杆号", "后端实时图谱图片",
    "后端A相幅值(mV)", "后端B相幅值(mV)", "后端C相幅值(mV)",
    "双端定位谱图图片",
    "定位分析描述", "检测结果分析", "疑似杆情况说明", "处理意见",
    "现场图片", "验证描述"
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
          "前端设备悬挂杆号":"10kV野马线T车路箐支线#16.5",
          "前端A相幅值(mV)":7148,"前端B相幅值(mV)":6198,"前端C相幅值(mV)":7460,
          "后端设备悬挂杆号":"10kV野马线T车路箐支线#16.10",
          "后端A相幅值(mV)":6148,"后端B相幅值(mV)":6198,"后端C相幅值(mV)":7510,
          "定位分析描述":"根据定位谱图，故障位置距离#2设备（10号杆）约56m，初步判断为#9号杆。",
          "检测结果分析":"根据现场测试数据可知，C相数值相比其他两相显著偏高且放电特征一致（其余两相信号疑似为感应耦合的高频信号），且图谱符合局部放电特性；10kV野马线#16.5断开后端供电后，发现疑似局放信号骤降95%左右，因此判断疑似局放来源后端杆塔；后续在10kV野马线#16.5与#16.10进行双端定位时，发现疑似局放信号来源于 #16.9，信号表明此处位置处数值幅值最大且图谱特征极为明显。",
          "疑似杆情况说明":"到现场杆塔下人耳可听见明显放电声，肉眼观察无明显异常的现象，后续用其他远程拍摄设备发现绝缘子内部存在疑似穿孔、绝缘子捆扎线存在明显放电痕迹。",
          "处理意见":"放电强度较强，人耳可听明显放电声，现场已确定疑似故障位置，需及时符合验证，并进行缺陷故障处理（停电更换或带电左右更换）。",
          "验证描述": "经带电作业（或停电检修）流程，对设备定位的故障缺陷位置实施了精准消缺处理。现场拆解疑似缺陷绝缘子，可见明显穿孔击穿痕迹。采用兆欧表对该绝缘子进行绝缘电阻检测，实测绝缘电阻值为5MΩ。",
          "前端实时图谱图片": "", "后端实时图谱图片": "", "双端定位谱图图片": "", "现场图片": ""
        },
        {
          "序号":2, "线路名称":"10kV野马线", "结论":"10kV野马线T车路箐支线#16.29 C相绝缘子存在局部放电隐患",
          "前端设备悬挂杆号":"10kV野马线#16.29",
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
              ["xl/worksheets/sheet1.xml","xl/worksheets/sheet2.xml","xl/worksheets/sheet3.xml"].forEach(function(n){
                if (z.file(n)) sheetFiles.push(n);
              });
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
          if (wsi) {
            var arr = XLSX.utils.sheet_to_json(wsi, { header: 1 });
            arr.forEach(function(r) {
              var k = (r[0] || "").toString().trim();
              if (BASIC_INFO.indexOf(k) >= 0 && hasVal(r[1])) info[k] = String(r[1]).trim();
            });
          }

          // 读取缺陷检测记录
          var wsd = wb.Sheets["缺陷检测记录"];
          if (wsd) {
            var arr2 = XLSX.utils.sheet_to_json(wsd, { header: 1 });
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
            resolve({ info: info, defects: defects });
          }).catch(function(){ resolve({ info: info, defects: defects }); });
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

    h += '<div class="report-blue-bar">现场图片</div>';
    h += imgHtml(d["现场图片"], "现场图片");

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

    // 辅助：cell 设值 + 样式
    function setCell(ws, ref, value, opts) {
      opts = opts || {};
      var cell = ws[ref];
      if (!cell) cell = XLSX.utils.allocate_named_range ? null : null;
      cell = XLSX.utils.decode_cell(ref);
      var addr = XLSX.utils.encode_cell(cell);
      ws[addr] = { t: typeof value === "number" ? "n" : "s", v: value };
      if (opts.font) ws[addr].s = ws[addr].s || {};
      if (opts.bold || opts.fill || opts.font) {
        ws[addr].s = ws[addr].s || {};
        if (opts.bold) ws[addr].s.font = Object.assign({}, ws[addr].s.font || {}, { bold: true });
        if (opts.fill) {
          ws[addr].s.fill = { patternType: "solid", fgColor: { rgb: opts.fill } };
        }
        if (opts.align) ws[addr].s.alignment = { horizontal: opts.align };
      }
    }

    var wb = XLSX.utils.book_new();

    // Sheet 1: 报告基本信息
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
    // 合并标题行
    ws1["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    // 列宽
    ws1["!cols"] = [{ wch: 16 }, { wch: 60 }];
    // 标题样式
    setCell(ws1, "A1", "报告基本信息（每份报告填写一次）", { fill: "1A5276", bold: true, align: "center" });
    setCell(ws1, "A1", "报告基本信息（每份报告填写一次）", { fill: "1A5276", bold: true, align: "center" });
    // 标签样式（列A）
    ["A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14"].forEach(function(r) {
      if (ws1[r]) ws1[r].s = Object.assign({}, ws1[r].s || {}, { font: { bold: true }, fill: { patternType: "solid", fgColor: { rgb: "EBF5FB" } } });
    });
    // 标题字体白色（合并单元格）
    ws1["A1"].s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { patternType: "solid", fgColor: { rgb: "1A5276" } }, alignment: { horizontal: "center" } };

    XLSX.utils.book_append_sheet(wb, ws1, "报告基本信息");

    // Sheet 2: 缺陷检测记录
    var headerRow = HEADERS.slice();
    var aoa2 = [
      ["缺陷检测记录（每行 = 一个缺陷点位，删除示例后填实数据）"],
      headerRow
    ];
    // 合并标题行
    var ws2 = XLSX.utils.aoa_to_sheet(aoa2);
    ws2["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headerRow.length - 1 } }];
    // 示例数据
    var sample = getExampleData().defects;
    sample.forEach(function(d) {
      var row = headerRow.map(function(h) { return g(d[h], ""); });
      XLSX.utils.sheet_add_aoa(ws2, [row], { origin: -1 });
    });
    // 表头样式（第二行）
    for (var c = 0; c < headerRow.length; c++) {
      var addr = XLSX.utils.encode_cell({ r: 1, c: c });
      if (ws2[addr]) {
        ws2[addr].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { patternType: "solid", fgColor: { rgb: "1A5276" } },
          alignment: { horizontal: "center", vertical: "center", wrapText: true }
        };
      }
    }
    // 标题行样式
    ws2["A1"].s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { patternType: "solid", fgColor: { rgb: "1A5276" } }, alignment: { horizontal: "center" } };
    // 列宽
    ws2["!cols"] = headerRow.map(function() { return { wch: 14 }; });
    XLSX.utils.book_append_sheet(wb, ws2, "缺陷检测记录");

    // Sheet 3: 填写说明
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
    XLSX.utils.book_append_sheet(wb, ws3, "填写说明");

    // 写入并下载
    var wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    var blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "架空线双端局放检测_数据填写模板.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
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
