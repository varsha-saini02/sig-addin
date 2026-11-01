/* taskpane.js — Word Signature Add-in (Test Version) */
/* NCCS CADCAM                                        */

(function () {
  'use strict';

  // ── State ──────────────────────────────────────────────────────────────────
  let isAssessorMode = false;
  let isDrawing      = false;
  let hasSignature   = false;
  let lastX = 0, lastY = 0;
  let userEmail = '';

  // ── Office Ready ───────────────────────────────────────────────────────────
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
      initUI();
      detectEmail();
    }
  });

  // ── Detect M365 Login Email ────────────────────────────────────────────────
  function detectEmail() {
    const display = document.getElementById('emailDisplay');

    // Try Office userProfile (works in Word Online / M365)
    try {
      const profile = Office.context.userProfile;
      const email   = (profile && (profile.email || profile.emailAddress)) || '';
      if (email) {
        userEmail = email;
        display.textContent = email;
        return;
      }
    } catch (e) { /* continue to fallback */ }

    // Fallback — manual input
    userEmail = '';
    display.innerHTML =
      '<input type="email" id="emailInput" class="email-input" ' +
      'placeholder="Enter your M365 email address">';
    document.getElementById('emailInput').addEventListener('input', function () {
      userEmail = this.value.trim();
    });
  }

  // ── UI Setup ───────────────────────────────────────────────────────────────
  function initUI() {
    const toggle    = document.getElementById('modeToggle');
    const badge     = document.getElementById('modeBadge');
    const insertBtn = document.getElementById('insertBtn');
    const clearBtn  = document.getElementById('clearBtn');

    toggle.addEventListener('change', function () {
      isAssessorMode = this.checked;
      if (isAssessorMode) {
        badge.textContent  = 'Assessor';
        badge.className    = 'mode-badge assessor';
        insertBtn.className = 'btn btn-insert assessor';
        insertBtn.textContent = 'Insert Assessor Signature';
      } else {
        badge.textContent  = 'Student';
        badge.className    = 'mode-badge student';
        insertBtn.className = 'btn btn-insert';
        insertBtn.textContent = 'Insert Signature';
      }
      clearStatus();
    });

    insertBtn.addEventListener('click', doInsertSignature);
    clearBtn.addEventListener('click',  clearCanvas);

    initCanvas();
  }

  // ── Canvas Drawing ─────────────────────────────────────────────────────────
  function initCanvas() {
    const canvas = document.getElementById('sigCanvas');
    const ctx    = canvas.getContext('2d');

    // Match physical pixel size to display size
    function resize() {
      const rect = canvas.getBoundingClientRect();
      const dpr  = window.devicePixelRatio || 1;
      canvas.width  = rect.width  * dpr;
      canvas.height = rect.height * dpr;
      ctx.scale(dpr, dpr);
      ctx.strokeStyle = '#1a1a1a';
      ctx.lineWidth   = 2.2;
      ctx.lineCap     = 'round';
      ctx.lineJoin    = 'round';
    }
    resize();

    function getPos(e) {
      const rect   = canvas.getBoundingClientRect();
      const source = e.touches ? e.touches[0] : e;
      return {
        x: source.clientX - rect.left,
        y: source.clientY - rect.top
      };
    }

    function onStart(e) {
      e.preventDefault();
      isDrawing = true;
      const p = getPos(e);
      lastX = p.x; lastY = p.y;
    }

    function onMove(e) {
      if (!isDrawing) return;
      e.preventDefault();
      const p = getPos(e);
      ctx.beginPath();
      ctx.moveTo(lastX, lastY);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      lastX = p.x; lastY = p.y;

      if (!hasSignature) {
        hasSignature = true;
        canvas.classList.add('active');
        document.getElementById('insertBtn').disabled = false;
      }
    }

    function onEnd() { isDrawing = false; }

    canvas.addEventListener('mousedown',  onStart);
    canvas.addEventListener('mousemove',  onMove);
    canvas.addEventListener('mouseup',    onEnd);
    canvas.addEventListener('mouseleave', onEnd);
    canvas.addEventListener('touchstart', onStart, { passive: false });
    canvas.addEventListener('touchmove',  onMove,  { passive: false });
    canvas.addEventListener('touchend',   onEnd);
  }

  function clearCanvas() {
    const canvas = document.getElementById('sigCanvas');
    const ctx    = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    canvas.classList.remove('active');
    hasSignature = false;
    document.getElementById('insertBtn').disabled = true;
    clearStatus();
  }

  // ── Insert Into Word Document ──────────────────────────────────────────────
  async function doInsertSignature() {
    // Collect email (handles both auto-detected and manual input)
    const emailInput = document.getElementById('emailInput');
    if (emailInput && emailInput.value.trim()) {
      userEmail = emailInput.value.trim();
    }

    if (!userEmail) {
      showStatus('Please enter your email address first.', 'error');
      return;
    }

    // Export canvas to base64 PNG
    const canvas = document.getElementById('sigCanvas');
    const base64 = canvas.toDataURL('image/png').split(',')[1];

    const mode       = isAssessorMode ? 'Assessor' : 'Student';
    const searchText = 'Signature and Date, ' + mode;
    const dateStr    = new Date().toLocaleDateString('en-AU', {
      day: '2-digit', month: '2-digit', year: 'numeric'
    });

    showStatus('Finding ' + mode.toLowerCase() + ' signature fields…', 'info');

    try {
      await Word.run(async function (context) {

        // Search the document for matching signature labels
        const searchResults = context.document.body.search(searchText, {
          matchCase: false,
          matchWholeWord: false
        });
        context.load(searchResults, 'items');
        await context.sync();

        const count = searchResults.items.length;

        if (count === 0) {
          showStatus(
            'No ' + mode + ' signature fields found in this document. ' +
            'Make sure the document contains the text "' + searchText + '".',
            'error'
          );
          return;
        }

        // Insert signature image + email into each matching field
        for (let i = 0; i < count; i++) {
          const labelRange = searchResults.items[i];

          // 1. Insert the drawn signature image right after the label text
          labelRange.insertInlinePictureFromBase64(base64, 'End');

          // 2. Append the email and date on the same line after the image
          //    (Word queues these in order, so text follows the image)
          labelRange.insertText(
            '  ' + userEmail + '  |  ' + dateStr,
            'End'
          );
        }

        await context.sync();

        showStatus(
          '✓ Signature inserted into ' + count + ' ' +
          mode.toLowerCase() + ' field' + (count !== 1 ? 's' : '') + '.',
          'success'
        );
      });

    } catch (err) {
      console.error(err);
      showStatus('Error: ' + (err.message || 'Something went wrong.'), 'error');
    }
  }

  // ── Helpers ────────────────────────────────────────────────────────────────
  function showStatus(msg, type) {
    const el = document.getElementById('statusMsg');
    el.textContent = msg;
    el.className   = 'status ' + type;
  }

  function clearStatus() {
    const el = document.getElementById('statusMsg');
    el.className = 'status';
  }

})();
