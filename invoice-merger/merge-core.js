/*
 * merge-core.js — 发票整合核心逻辑
 * 同时可在浏览器（配合 pdf-lib UMD）与 Node（配合 pdf-lib npm）中使用。
 *
 * 用法：
 *   mergeInvoices(PDFDocument, files, opts)
 *     PDFDocument : pdf-lib 的 PDFDocument 类
 *     files       : [{ name: string, bytes: Uint8Array,
 *                      train?: boolean,            // 是否火车票（由票面内容识别）
 *                      content?: [{                // 每页「实际内容」信息（去掉空白）
 *                        hCm: number,              // 内容高度 cm
 *                        wCm: number,              // 内容宽度 cm
 *                        empty?: boolean,          // 整页空白
 *                        bbox?: {x,y,w,h}          // 内容包围盒（PDF 源坐标，原点左下，单位 pt）
 *                      }] }]
 *     opts        : {
 *        thresholdCm : number  (内容高度超过该值 cm 的发票独占一页，默认 14)
 *        marginMm   : number  (A4 页边距，单位 mm，默认 8)
 *        isTrain    : (name:string)=>boolean  (兜底：无法提取文字时按文件名关键字识别)
 *     }
 *   说明：
 *     - “小发票”的判定依据是【实际内容高度】（去掉四周空白），而非整页 PDF 高度。
 *       这样“扫描进 A4 里的小票/火车票”也能被正确识别为小发票、两联拼版。
 *     - 两联拼版时会按内容包围盒裁剪绘制，避免空白被一起缩小。
 *   返回：合并后的 PDF Uint8Array；无文件时返回 null。
 */
(function (global) {
  const PT_PER_CM = 28.3464567;
  const PT_PER_MM = 2.83464567;
  const A4 = [595.28, 841.89]; // A4 单位 pt
  const PDFLibRef = global.PDFLib; // 浏览器为 window.PDFLib，Node 下需在 globalThis 上挂 PDFLib

  const DEFAULT_TRAIN_KEYWORDS = ['火车票', '铁路', '高铁', '客票', '车票', 'train', '行程单'];

  function defaultIsTrain(name) {
    const lower = (name || '').toLowerCase();
    return DEFAULT_TRAIN_KEYWORDS.some((k) => lower.includes(k.toLowerCase()));
  }

  // v15：行程单预旋转——把源 PDF 旋转 90°（MediaBox 互换 + 内容流包裹旋转矩阵），
  // 绕开 pdf-lib 旧版 drawPage+rotate 的内容流序列化 bug（Do 丢失导致空页）。
  async function preRotate90(PDFDocument, bytes) {
    const src = await PDFDocument.load(bytes);
    const context = src.context;
    const pages = src.getPages();
    for (const pg of pages) {
      // 用 CropBox（回退 MediaBox）：真实发票常用 CropBox 定义可见区（如 [0,433,595,842]），
      // MediaBox 缺失时 pdf-lib 会返回默认 A4 导致隐藏内容（如 didi）无法裁剪
      const mb = pg.node.CropBox() || pg.node.MediaBox();
      const mbX = parseFloat(String(mb.get(0)));
      const mbY = parseFloat(String(mb.get(1)));
      const mbW = parseFloat(String(mb.get(2))) - mbX;
      const mbH = parseFloat(String(mb.get(3))) - mbY;
      // MediaBox 完整重设：旋转 90° 后内容占 [0, mbH]×[0, mbW]
      mb.set(0, context.obj(0));
      mb.set(1, context.obj(0));
      mb.set(2, context.obj(mbH));
      mb.set(3, context.obj(mbW));
      // head：先裁剪到原 MediaBox（剔除 MediaBox 外的隐藏内容，如某些发票的水印/装饰字），再旋转
      // 平移量 = mbY+mbH：原 y=mbY 的内容旋转后 x'=0，y=mbY+mbH → x'=mbH
      const csHead = context.flateStream(
        "q\n" + mbX + " " + mbY + " " + mbW + " " + mbH + " re\nW\nn\n" +
        "0 1 -1 0 " + (mbY + mbH) + " 0 cm\n"
      );
      const csTail = context.flateStream("Q\n");
      // 三态处理：Contents 可能是 PDFArray / 单流 PDFStream / 不存在（扫描件常见单流）
      const headRef = context.register(csHead);
      const tailRef = context.register(csTail);
      const contents = pg.node.Contents();
      if (!contents) {
        // 无 Contents：新建数组
        pg.node.set(PDFLibRef.PDFName.of('Contents'), context.obj([headRef, tailRef]));
      } else if (typeof contents.asArray === 'function') {
        // PDFArray：内部数组操作
        const arr = contents.asArray();
        arr.unshift(headRef);
        arr.push(tailRef);
      } else {
        // 单流 PDFStream：包成 [head, 原流, tail] 数组替换
        pg.node.set(PDFLibRef.PDFName.of('Contents'), context.obj([headRef, contents, tailRef]));
      }
    }
    return await src.save();
  }

  // v19：按内容包围盒物理裁剪页面（MediaBox+CropBox 同步设为 bbox）——
  // pdf-lib embedPdf 只取 MediaBox 作为 Form XObject 的 BBox，仅设 CropBox 会被忽略；
  // MediaBox 缩小后，嵌入页 BBox=bbox，绘制时只渲染正文区域（剔除下方图片/广告区，如滴滴发票的 didi 区）
  // v23 修复：bbox 来自画布坐标系（基于 CropBox，如滴滴发票 CropBox=[0,383,610,864]），
  // 需平移到 MediaBox 坐标系（加 CropBox 左下原点偏移），否则裁剪框错位、正文被裁掉只剩底部装饰区。
  // v27 修复：仅改 MediaBox/CropBox 不改内容流会导致内容坐标错位（内容流坐标基于原 MediaBox 原点，
  // 落在新 BBox 之外 → embedPdf 绘制后整页空白）。正确做法与 preRotate90 一致：
  // MediaBox/CropBox 归一化到 [0,0,w,h]，同时用内容流包裹矩阵把原内容平移到新原点。
  async function applyCropBoxes(PDFDocument, bytes, content) {
    const src = await PDFDocument.load(bytes);
    const context = src.context;
    const pages = src.getPages();
    pages.forEach(function(pg, idx) {
      const c = content && content[idx];
      if (!c || !c.cropByText || !c.bbox || c.bbox.w <= 0 || c.bbox.h <= 0) return;
      const media = pg.node.MediaBox();
      const crop = pg.node.CropBox() || media;
      const cx = parseFloat(String(crop.get(0)));
      const cy = parseFloat(String(crop.get(1)));
      const bb = c.bbox;
      // 裁剪区(原 PDF 坐标系,左下原点)
      const bx = cx + bb.x;
      const by = cy + bb.y;
      const bw = bb.w;
      const bh = bb.h;
      if (bw <= 0 || bh <= 0) return;
      // MediaBox/CropBox 归一化到 [0,0,w,h]
      const box = context.obj([0, 0, bw, bh]);
      pg.node.set(PDFLibRef.PDFName.of('MediaBox'), box);
      pg.node.set(PDFLibRef.PDFName.of('CropBox'), box);
      // 内容流包裹：平移 (-bx,-by) 把原内容移到新原点（裁剪区之外的内容自然落到新 MediaBox 外被裁掉）
      const csHead = context.flateStream(
        "q\n1 0 0 1 " + (-bx) + " " + (-by) + " cm\n"
      );
      const csTail = context.flateStream("Q\n");
      const headRef = context.register(csHead);
      const tailRef = context.register(csTail);
      const contents = pg.node.Contents();
      if (!contents) {
        pg.node.set(PDFLibRef.PDFName.of('Contents'), context.obj([headRef, tailRef]));
      } else if (typeof contents.asArray === 'function') {
        const arr = contents.asArray();
        arr.unshift(headRef);
        arr.push(tailRef);
      } else {
        pg.node.set(PDFLibRef.PDFName.of('Contents'), context.obj([headRef, contents, tailRef]));
      }
    });
    return await src.save();
  }

  // 把单张内容（带 bbox）裁剪绘制到 A4 的某个槽位（v10：支持水平并排槽位，slotX=槽左边界）
  function drawInSlot(p, it, slotW, slotH, slotX, slotBottomY, A4, topAlign) {
    // v15：统一整页缩放绘制（行程单内容已预旋转，绘制阶段不再旋转；无 bbox 裁剪防截断）
    // v19：topAlign=true 时内容顶部对齐槽顶（单张票居上显示）
    const w = it.w, h = it.h;
    const s = Math.min(slotW / w, slotH / h);
    const dw = w * s;
    const dh = h * s;
    const tx = slotX + (slotW - dw) / 2;
    const ty = topAlign ? (slotBottomY + slotH - dh) : (slotBottomY + (slotH - dh) / 2);
    p.drawPage(it.ep, { x: tx, y: ty, width: dw, height: dh });
  }

  async function mergeInvoices(PDFDocument, files, opts) {
    opts = opts || {};
    const thresholdCm = opts.thresholdCm != null ? opts.thresholdCm : 14;
    const forceTwoUp = !!opts.forceTwoUp;   // v7：强制两联拼版（忽略 14cm 阈值）
    const margin = (opts.marginMm != null ? opts.marginMm : 8) * PT_PER_MM;
    const isTrain = opts.isTrain || defaultIsTrain;
    const trainDouble = opts.trainDouble !== false;  // v24：火车票加印默认开启（默认双份打印）

    if (!files || files.length === 0) return null;

    const out = await PDFDocument.create();

    // 1) 嵌入每个 PDF，逐页记录「内容高度 / 内容包围盒」与是否火车票
    const items = [];
    for (const f of files) {
      // v18：任一页实际内容高 >14cm 的发票/行程单旋转 90°（不适宜正常两联拼版 → 横放缩放适配 A4）；
      // 火车票豁免旋转（按原版方向上下排布）
      let fBytes = f.bytes;
      const train = f.train != null ? !!f.train : isTrain(f.name);
      const maxContentCm = Math.max(0, ...(f.content || []).map(function(c){ return c ? (c.hCm || 0) : 0; }));
      // v19：命中 cropByText 的页面先按内容包围盒裁剪（CropBox），只保留正文区域（如滴滴发票去掉 didi 区）
      // v21 修复：火车票跳过内容裁剪——“购买方名称”版式客票的 bbox 计算会把票面上半部裁掉（发票号/车站/票价丢失）
      const needCrop = (f.content || []).some(function(c){ return c && c.cropByText && c.bbox && c.bbox.w > 0 && c.bbox.h > 0; });
      if (needCrop && !train) fBytes = await applyCropBoxes(PDFDocument, f.bytes, f.content);
      // v22：仅竖版高票（内容宽 < 内容高）才旋转 90°——横版高票（如滴滴发票 566×422）旋转后会进一步压扁，
      // 按原方向两联拼版 + bbox 裁剪后字号接近原大，打印清晰；无 content 时保持旧行为（旋转）。
      const c0 = (f.content && f.content[0]) || {};
      const cw = c0.wCm || 0, ch = c0.hCm || 0;
      const isPortrait = cw > 0 && ch > 0 && cw < ch;
      if (maxContentCm > 14 && !train && (!cw || isPortrait)) {
        fBytes = await preRotate90(PDFDocument, f.bytes);
      }
      // v25：修复多页 PDF 静默丢页——pdf-lib embedPdf 默认只嵌第 1 页（indices=[0]），
      // 先 load 源文件取页数，显式嵌入全部页（多页扫描件/多联票据不再丢页）
      const srcDoc = await PDFDocument.load(fBytes, { ignoreEncryption: true });
      const srcPages = srcDoc.getPageCount();
      const indices = [];
      for (let pi = 0; pi < srcPages; pi++) indices.push(pi);
      const embedded = await out.embedPdf(fBytes, indices); // PDFEmbeddedPage[]
      embedded.forEach((ep, idx) => {
        const c = f.content && f.content[idx];
        const contentHcm = c ? (c.hCm || ep.height / PT_PER_CM) : (ep.height / PT_PER_CM);
        const contentBbox = c && !c.empty && c.bbox ? c.bbox : null;
        items.push({
          ep,
          w: ep.width,
          h: ep.height,
          contentHcm,        // 实际内容高度（cm）——小发票判定的依据
          contentBbox,       // 内容包围盒（用于裁剪绘制）
          train,
          name: f.name,
          page: idx + 1,
          type: f.type,                                    // v7：票据类型（itinerary=行程单）
          rotate90: f.type === 'itinerary',                // v7：行程单旋转 90° 后拼版
        });
      });
    }

    // 2) 火车票加印（默认开启双份；opts.trainDouble=false 可关闭）
    const seq = [];
    for (const it of items) {
      seq.push(it);
      if (it.train && trainDouble) seq.push(it);
    }

    // 3) 布局：内容高度 >= 阈值 → 独占一页；否则两联拼版（裁剪到内容）
    //    火车票固定一页两张，不受阈值影响。
    // v13：长于 14cm 的发票缩放至 ≤14cm 两联拼版，不再独占一页（避免单票单页浪费）
    const isTall = () => false;

    let i = 0;
    while (i < seq.length) {
      const it = seq[i];
      if (isTall(it)) {
        // 独占整页 A4（整页缩放居中）
        const p = out.addPage(A4);
        const scale = Math.min((A4[0] - 2 * margin) / it.w, (A4[1] - 2 * margin) / it.h);
        const dw = it.w * scale;
        const dh = it.h * scale;
        p.drawPage(it.ep, { x: (A4[0] - dw) / 2, y: (A4[1] - dh) / 2, width: dw, height: dh });
        i++;
      } else {
        // 2-up：尝试与下一张（同样为小发票）拼在一起，按内容裁剪绘制
        const pair = [it];
        if (i + 1 < seq.length && !isTall(seq[i + 1])) {
          pair.push(seq[i + 1]);
        }
        const p = out.addPage(A4);
        const n = pair.length;
        const gap = margin;
        // v16：统一上下堆叠排布（发票/行程单一视同仁）；单张时居上（与两张时第一张同位置）
        const avail = A4[1] - 2 * margin;
        const slotH = Math.min(14 * PT_PER_CM, (avail - (n - 1) * gap) / n);
        const slotW = A4[0] - 2 * margin;
        for (let k = 0; k < n; k++) {
          const it2 = pair[k];
          const slotTopY = A4[1] - margin - k * (slotH + gap);
          const slotBottomY = slotTopY - slotH;
          drawInSlot(p, it2, slotW, slotH, margin, slotBottomY, A4, n === 1);
        }
        i += n;
      }
    }

    return await out.save();
  }

  global.mergeInvoices = mergeInvoices;
})(typeof window !== 'undefined' ? window : globalThis);
