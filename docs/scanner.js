'use strict';
/* סריקה במצלמה – גיבוי לסורק החיצוני.
 * BarcodeDetector (אנדרואיד/Chrome) כשיש; אחרת ZXing (נטען מקומית – עובד גם בלי קליטה, גם באייפון).
 * אם בתמונה יש כמה ברקודים: עדיפות לברקוד שמתאים לתבנית של השדה. ברקוד אחר מתקבל רק אם נקרא כמה פעמים ברצף.
 */
const Camera = (() => {
  let stream = null, timer = null, zx = null, done = null;

  function stop() {
    clearTimeout(timer); timer = null;
    if (zx) { try { zx.stop(); } catch (_) { /* */ } zx = null; }
    if (stream) { stream.getTracks().forEach(t => t.stop()); stream = null; }
    if (done) { const d = done; done = null; d(null); }
  }

  function loadZxing() {
    if (window.ZXingBrowser) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'vendor/zxing-browser.min.js';
      s.onload = () => res();
      s.onerror = () => rej(new Error('ספריית הסריקה לא נטענה'));
      document.head.appendChild(s);
    });
  }

  /** מחזיר את הערך הנסרק, או null אם בוטל */
  async function scan(video, pattern, onMsg) {
    stop();
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('הדפדפן לא תומך במצלמה (צריך https)');
    const result = new Promise(r => { done = r; });
    const seen = {};
    const pick = values => {
      const good = values.find(v => pattern.test(v.trim()));
      if (good) return good;
      for (const v of values) { seen[v] = (seen[v] || 0) + 1; if (seen[v] >= 4) return v; }
      if (values.length) onMsg('נקרא ברקוד שלא מתאים לשדה – מכוון לברקוד הנכון…');
      return null;
    };
    const finish = v => { const d = done; done = null; stop(); if (d) d(v); };

    let useNative = false;
    if ('BarcodeDetector' in window) {
      try {
        const wanted = ['code_128', 'code_39', 'code_93', 'codabar', 'itf', 'ean_13', 'ean_8', 'upc_a', 'upc_e', 'qr_code', 'data_matrix', 'pdf417'];
        const sup = await window.BarcodeDetector.getSupportedFormats();
        const formats = wanted.filter(f => sup.includes(f));
        if (formats.length) {
          const det = new window.BarcodeDetector({ formats });
          useNative = true;
          try {
            stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment', width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false });
          } catch (e) { throw camError(e); }
          video.srcObject = stream;
          await video.play();
          const loop = async () => {
            if (!stream) return;
            try {
              const codes = await det.detect(video);
              const v = pick(codes.map(c => c.rawValue).filter(Boolean));
              if (v) return finish(v);
            } catch (_) { /* פריים לא מוכן */ }
            timer = setTimeout(loop, 120);
          };
          loop();
        }
      } catch (e) {
        if (e.isCam) throw e;
        useNative = false;
      }
    }

    if (!useNative) {
      onMsg('טוען סורק…');
      await loadZxing();
      onMsg('מכוון את המצלמה לברקוד…');
      const reader = new window.ZXingBrowser.BrowserMultiFormatReader();
      try {
        zx = await reader.decodeFromConstraints(
          { video: { facingMode: 'environment', width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false },
          video,
          r => { if (r) { const v = pick([r.getText()]); if (v) finish(v); } }
        );
      } catch (e) { throw camError(e); }
    }
    return result;
  }

  function camError(e) {
    const err = new Error(e && e.name === 'NotAllowedError' ? 'אין הרשאה למצלמה – אשר גישה בהגדרות הדפדפן' : 'המצלמה לא זמינה');
    err.isCam = true;
    return err;
  }

  return { scan, stop };
})();
