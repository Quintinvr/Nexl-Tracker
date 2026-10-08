/*
 * Nexl Check — container photo verification.
 *
 * Reads the CONTAINER photo the driver uploaded in Nexl (Upload Viewer), runs OCR on it in the
 * browser (tesseract.js), and confirms it shows the same container number the driver entered in the app.
 * Only a confirmed match unlocks the "Apply" button that writes the number into the sheet.
 *
 * Matching rules (tuned on real Nexl photos):
 *  - the 6-digit serial must be found in order, allowing common OCR mix-ups (O/0, I/1, S/5, Z/2/7, B/8, G/6)
 *  - the 4-letter owner code must be found just before it (≤1 character wrong, small gaps allowed)
 *  - the app's own number must pass the ISO 6346 check digit, so the boxed digit the OCR often misses
 *    is still confirmed mathematically.
 */
(function (root) {
  "use strict";

  // ---------- ISO 6346 ----------
  const LETTER_VALUES = (() => {
    const v = {}; let n = 10;
    for (const ch of "ABCDEFGHIJKLMNOPQRSTUVWXYZ") { if (n % 11 === 0) n++; v[ch] = n++; }
    return v;
  })();
  function iso6346Valid(no) {
    const s = String(no || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!/^[A-Z]{4}\d{7}$/.test(s)) return false;
    let sum = 0;
    for (let i = 0; i < 10; i++) {
      const c = s[i];
      sum += (/\d/.test(c) ? +c : LETTER_VALUES[c]) * Math.pow(2, i);
    }
    return (sum % 11) % 10 === +s[10];
  }

  // Characters OCR confuses. Keys are what the photo-reader may output; values what it may have meant.
  const AS_DIGIT = { O: "0", Q: "0", D: "0", U: "0", I: "1", L: "1", J: "1", T: "7", Z: "2", S: "5", B: "8", G: "6", A: "4" };
  const AS_LETTER = { "0": "O", "1": "I", "2": "Z", "5": "S", "7": "Z", "8": "B", "6": "G", "4": "A" };
  const digitOk = (got, want) => got === want || AS_DIGIT[got] === want;
  const letterOk = (got, want) => got === want || AS_LETTER[got] === want || (want === "O" && got === "D") || (want === "U" && got === "V");

  /**
   * @param ocrText raw OCR text of the CONTAINER photo
   * @param appNo   container number the driver entered in the Nexl app
   * @returns {status:"match"|"mismatch"|"unreadable"|"invalid", seen?:string, detail:string}
   */
  function verifyContainer(ocrText, appNo) {
    const want = String(appNo || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!/^[A-Z]{4}\d{7}$/.test(want)) return { status: "invalid", detail: `"${appNo}" isn't a standard container number` };
    if (!iso6346Valid(want)) return { status: "invalid", detail: `${want} fails the ISO check digit, so the app entry looks mistyped` };
    const owner = want.slice(0, 4), serial = want.slice(4, 10), check = want[10];
    const lines = String(ocrText || "").toUpperCase().split(/\n/).map((l) => l.replace(/[^A-Z0-9]/g, ""));
    const text = lines.join("|");

    // 1) find the serial (6 chars) allowing digit mix-ups
    let best = null;
    for (let i = 0; i + 6 <= text.length; i++) {
      const chunk = text.slice(i, i + 6);
      if (chunk.includes("|")) continue;
      let ok = true;
      for (let k = 0; k < 6; k++) if (!digitOk(chunk[k], serial[k])) { ok = false; break; }
      if (!ok) continue;
      // 2) owner code just before it (allow 0-2 junk chars between, ≤1 wrong letter)
      for (let gap = 0; gap <= 2; gap++) {
        const s = i - gap - 4;
        if (s < 0) continue;
        const own = text.slice(s, s + 4);
        if (own.includes("|")) continue;
        let wrong = 0;
        for (let k = 0; k < 4; k++) if (!letterOk(own[k], owner[k])) wrong++;
        if (wrong <= 1) {
          const after = text.slice(i + 6, i + 9);
          const sawCheck = after.includes(check);
          const cand = { pos: i, wrong, gap, own, sawCheck };
          if (!best || cand.wrong < best.wrong || (cand.wrong === best.wrong && cand.gap < best.gap)) best = cand;
        }
      }
    }
    if (best) {
      return { status: "match", seen: `${best.own} ${text.slice(best.pos, best.pos + 6)}`,
        detail: `Photo shows ${owner} ${serial}${best.sawCheck ? " " + check : ""}; check digit ${check} confirmed (ISO 6346)` };
    }

    // 3) a different container number visible? -> mismatch (only when we can read one clearly)
    const other = /([A-Z]{3}U)(\d{6})(\d?)/.exec(text.replace(/\|/g, " "));
    if (other && (other[1] !== owner || other[2] !== serial)) {
      const seenNo = other[1] + other[2] + (other[3] || "");
      // Only say "photo shows another number" when the full number was read and its ISO check digit is right;
      // a partial read (e.g. "HLBU 914015" for HLBU 914918 1) is just an unclear photo.
      if (seenNo.length === 11 && iso6346Valid(seenNo)) {
        return { status: "mismatch", seen: `${other[1]} ${other[2]}${other[3] ? " " + other[3] : ""}`, detail: `Photo shows ${other[1]} ${other[2]}, the app says ${owner} ${serial} ${check}` };
      }
    }
    return { status: "unreadable", detail: "Couldn't read the container number on the photo; check it by eye" };
  }

  /** Pull photo URLs out of Nexl's Upload Viewer HTML. */
  function parseUploads(html) {
    const urls = [...new Set(String(html || "").match(/\/php\/Uploads\/[^"'<> )]+\.(?:jpe?g|png)/gi) || [])];
    const pick = (kind) => urls.find((u) => u.split("/").includes(kind)) || null;
    return { container: pick("CONTAINER"), seal: pick("SEAL"), psli: pick("PSLI"), codn: pick("CODN"), damages: pick("DAMAGES"), all: urls };
  }

  // ---------- runtime (browser only) ----------
  let workerP = null;
  function ocrWorker() {
    if (!workerP) {
      if (!root.Tesseract) return Promise.reject(new Error("Photo reader didn't load (no internet?)"));
      workerP = (async () => {
        const w = await root.Tesseract.createWorker("eng");
        await w.setParameters({ tessedit_char_whitelist: "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 " });
        return w;
      })();
    }
    return workerP;
  }
  const cache = new Map(); // key -> result (per session)

  // Upload Viewer per Nexl container row. A port slip, once there, stays; otherwise re-check after 4 min.
  const upCache = new Map();
  function uploads(rowId) {
    const hit = upCache.get(rowId);
    if (hit && (hit.final || Date.now() - hit.at < 4 * 60000)) return hit.p;
    const p = root.NexlClient.getUploads(rowId).then(parseUploads);
    const ent = { p, at: Date.now(), final: false };
    upCache.set(rowId, ent);
    p.then((u) => { ent.final = !!u.psli; }, () => upCache.delete(rowId));
    return p;
  }
  const imgCache = new Map();
  function image(path) {
    if (!imgCache.has(path)) {
      const p = root.NexlClient.getImage(path);
      imgCache.set(path, p);
      p.catch(() => imgCache.delete(path));
    }
    return imgCache.get(path);
  }
  /** Seal photos can't be read reliably by OCR (small embossed bolts), so they are shown for a quick visual check. */
  async function sealPhoto(rowId) {
    const up = await uploads(rowId);
    if (!up.seal) return { status: "nophoto", detail: "The driver hasn't uploaded a seal photo yet" };
    return { status: "manual", detail: "Compare the seal photo with the number, then confirm", photo: await image(up.seal), photoPath: up.seal };
  }
  let queue = Promise.resolve();

  /**
   * Verifies one container fill. fill = {value (app container), nexlRowId}.
   * Resolves {status, detail, seen, photo (data URL), photoPath}.
   */
  function check(fill) {
    const key = `${fill.nexlRowId}|${fill.value}`;
    if (cache.has(key)) return cache.get(key);
    const p = (queue = queue.then(async () => {
      const up = await uploads(fill.nexlRowId);
      if (!up.container) return { status: "nophoto", detail: "The driver hasn't uploaded a container photo yet" };
      const img = await image(up.container);
      const w = await ocrWorker();
      const r = await w.recognize(img);
      return Object.assign(verifyContainer(r.data.text, fill.value), { photo: img, photoPath: up.container, text: r.data.text });
    }).catch((e) => ({ status: "error", detail: "Photo check failed: " + (e.message || e) })));
    cache.set(key, p);
    p.then((r) => { if (r.status === "error") cache.delete(key); });
    return p;
  }

  root.NexlPhotoCheck = { iso6346Valid, verifyContainer, parseUploads, check, uploads, image, sealPhoto };
})(typeof window !== "undefined" ? window : globalThis);
