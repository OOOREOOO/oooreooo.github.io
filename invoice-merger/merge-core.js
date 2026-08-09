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

  // 把单张内容（带 bbox）裁剪绘制到 A4 的某个槽位
  function drawCropped(p, it, slotW, slotH, slotBottomY, A4) {
    const bb = it.contentBbox;
    if (bb && bb.w > 0 && bb.h > 0) {
      const s = Math.min(slotW / bb.w, slotH / bb.h);
      const dw = bb.w * s;
      const dh = bb.h * s;
      const tx = (A4[0] - dw) / 2;
      const ty = slotBottomY + (slotH - dh) / 2;
      const pad = 2; // 留 2pt 余量，避免裁掉反锯齿边缘
      p.pushOperators(
        PDFLibRef.pushGraphicsState(),
        PDFLibRef.rectangle(tx - pad, ty - pad, dw + 2 * pad, dh + 2 * pad),
        PDFLibRef.clip(),
        PDFLibRef.endPath()
      );
      // 让内容包围盒的左下角 (bb.x, bb.y) 落在 (tx, ty)
      const xoff = tx - bb.x * s;
      const yoff = ty - bb.y * s;
      p.drawPage(it.ep, { x: xoff, y: yoff, width: it.w * s, height: it.h * s });
      p.pushOperators(PDFLibRef.popGraphicsState());
    } else {
      // 无内容包围盒：退化为整页缩放绘制
      const s = Math.min(slotW / it.w, slotH / it.h);
      const dw = it.w * s;
      const dh = it.h * s;
      const tx = (A4[0] - dw) / 2;
      const ty = slotBottomY + (slotH - dh) / 2;
      p.drawPage(it.ep, { x: tx, y: ty, width: dw, height: dh });
    }
  }

  async function mergeInvoices(PDFDocument, files, opts) {
    opts = opts || {};
    const thresholdCm = opts.thresholdCm != null ? opts.thresholdCm : 14;
    const margin = (opts.marginMm != null ? opts.marginMm : 8) * PT_PER_MM;
    const isTrain = opts.isTrain || defaultIsTrain;

    if (!files || files.length === 0) return null;

    const out = await PDFDocument.create();

    // 1) 嵌入每个 PDF，逐页记录「内容高度 / 内容包围盒」与是否火车票
    const items = [];
    for (const f of files) {
      const embedded = await out.embedPdf(f.bytes); // PDFEmbeddedPage[]
      const train = f.train != null ? !!f.train : isTrain(f.name);
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
        });
      });
    }

    // 2) 火车票额外打印一次（序列中紧接着再放一份）
    const seq = [];
    for (const it of items) {
      seq.push(it);
      if (it.train) seq.push(it);
    }

    // 3) 布局：内容高度 >= 阈值 → 独占一页；否则两联拼版（裁剪到内容）
    //    火车票固定一页两张，不受阈值影响。
    const isTall = (it) => !it.train && it.contentHcm >= thresholdCm;

    let i = 0;
    while (i < seq.length) {
      const it = seq[i];
      if (isTall(it)) {
        // 独占整页 A4（整页缩放居中）
        const p = out.addPage(A4);
        const scale = Math.min((A4[0] - 2 * margin) / it.w, (A4[1] - 2 * margin) / it.h);
        const dw = it.w * scale;
        const dh = it.h * scale;
        p.drawPage(it.ep, {
          x: (A4[0] - dw) / 2,
          y: (A4[1] - dh) / 2,
          width: dw,
          height: dh,
        });
        i++;
      } else {
        // 2-up：尝试与下一张（同样为小发票）拼在一起，按内容裁剪绘制
        const pair = [it];
        if (i + 1 < seq.length && !isTall(seq[i + 1])) {
          pair.push(seq[i + 1]);
        }
        const p = out.addPage(A4);
        const n = pair.length;
        const avail = A4[1] - 2 * margin;
        const gap = margin;
        const slotH = (avail - (n - 1) * gap) / n;
        const slotW = A4[0] - 2 * margin;
        for (let k = 0; k < n; k++) {
          const it2 = pair[k];
          const slotTopY = A4[1] - margin - k * (slotH + gap);
          const slotBottomY = slotTopY - slotH;
          drawCropped(p, it2, slotW, slotH, slotBottomY, A4);
        }
        i += n;
      }
    }

    return await out.save();
  }

  global.mergeInvoices = mergeInvoices;
})(typeof window !== 'undefined' ? window : globalThis);
