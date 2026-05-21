// ─── Border Rail Scene Animation ──────────────────────────────────
(function () {
  const canvas = document.getElementById('rail-canvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const PAD  = 36;
  const GAP  = 8;
  const TL   = 68;
  const TW   = 22;
  const MAX_V      = 3.2;
  const BRAKE_DIST = 190;
  const STOP_GAP   = 55;

  const signals = [
    { frac: 0.035, state: 'red',   timer: rnd(),       rot:  Math.PI / 2, poleDirY:  1 },
    { frac: 0.35,  state: 'lunar', timer: rnd() + 70,  rot:  Math.PI / 2, poleDirY:  1 },
    { frac: 0.52,  state: 'lunar', timer: rnd() + 140, rot: -Math.PI / 2, poleDirY: -1 },
    { frac: 0.85,  state: 'red',   timer: rnd() + 30,  rot: -Math.PI / 2, poleDirY: -1 },
  ];

  function rnd() { return 220 + Math.random() * 430; }

  let trainFrac = 0, trainV = MAX_V;

  // ── Geometry ──────────────────────────────────────────────────
  function perim() {
    return 2 * (canvas.width - 2 * PAD + canvas.height - 2 * PAD);
  }
  function wrap(t) { return ((t % 1) + 1) % 1; }
  function fwdDist(a, b) { let d = b - a; if (d < 0) d += 1; return d * perim(); }

  function trackPoint(t) {
    const W = canvas.width, H = canvas.height;
    const w = W - 2 * PAD, h = H - 2 * PAD;
    const d = wrap(t) * perim();
    if (d < w)       return { x: PAD + d,           y: PAD,             angle: 0 };
    if (d < w + h)   return { x: PAD + w,            y: PAD + (d-w),     angle: Math.PI / 2 };
    if (d < 2*w + h) return { x: PAD + w - (d-w-h), y: PAD + h,         angle: Math.PI };
                     return { x: PAD,                y: PAD+h-(d-2*w-h), angle: -Math.PI / 2 };
  }

  // ── Update ────────────────────────────────────────────────────
  function update() {
    signals.forEach(s => {
      if (--s.timer <= 0) {
        s.state = s.state === 'red' ? 'lunar' : 'red';
        s.timer = rnd();
      }
    });

    const nextRed = signals
      .filter(s => s.state === 'red')
      .map(s => ({ s, d: fwdDist(trainFrac, s.frac) }))
      .filter(({ d }) => d > 8 && d < BRAKE_DIST + 100)
      .sort((a, b) => a.d - b.d)[0];

    if (nextRed) {
      const d = nextRed.d;
      if (d < BRAKE_DIST) trainV = Math.max(0, trainV - 0.09);
      if (d <= STOP_GAP)  trainV = 0;
    } else {
      trainV = Math.min(MAX_V, trainV + 0.055);
    }
    trainFrac = wrap(trainFrac + trainV / perim());
  }

  // ── Rounded-rect path ─────────────────────────────────────────
  function rr(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y);
    ctx.arcTo(x+w, y,   x+w, y+r,   r); ctx.lineTo(x+w, y+h-r);
    ctx.arcTo(x+w, y+h, x+w-r, y+h, r); ctx.lineTo(x+r, y+h);
    ctx.arcTo(x, y+h,   x, y+h-r,   r); ctx.lineTo(x, y+r);
    ctx.arcTo(x, y,     x+r, y,      r); ctx.closePath();
  }

  // ── Draw track ────────────────────────────────────────────────
  function drawTrack() {
    const W = canvas.width, H = canvas.height;
    const p = perim();

    ctx.fillStyle = 'rgba(20,35,55,0.8)';
    for (let d = 0; d < p; d += 22) {
      const pt = trackPoint(d / p);
      ctx.save();
      ctx.translate(pt.x, pt.y); ctx.rotate(pt.angle);
      ctx.fillRect(-GAP * 1.6, -2, GAP * 3.2, 4);
      ctx.restore();
    }

    for (const offset of [-GAP / 2, GAP / 2]) {
      const r = PAD + offset;
      ctx.strokeStyle = offset < 0 ? 'rgba(155,180,200,0.75)' : 'rgba(100,130,155,0.65)';
      ctx.lineWidth = 2.2;
      ctx.strokeRect(r, r, W - 2*r, H - 2*r);
    }
  }

  // ── Draw signal (portrait head + pole/base, rotated per signal) ────
  function drawSignal(s) {
    const pt = trackPoint(s.frac);
    const lunarOn = Math.floor(Date.now() / 545) % 2 === 0;

    ctx.save();
    ctx.translate(pt.x, pt.y);
    ctx.rotate(pt.angle);
    // +x = direction of travel, +y = inward toward canvas centre

    const hW = 12, hD = 28, bR = 4.0;
    const poleW = 3, poleH = 18, baseW = 10, baseH = 4;
    const pD = s.poleDirY;   // +1 or -1: which portrait y-end the pole attaches to

    // Translate to signal centre (inner side, clear of train wheels ~21.5px)
    ctx.translate(0, 44);
    ctx.rotate(s.rot);

    // Pole and base drawn first so head sits on top
    ctx.fillStyle = '#3a4a5a';
    const poleY = pD > 0 ? hD / 2 : -hD / 2 - poleH;
    rr(-poleW / 2, poleY, poleW, poleH, 1); ctx.fill();
    const baseY = pD > 0 ? hD / 2 + poleH : -hD / 2 - poleH - baseH;
    rr(-baseW / 2, baseY, baseW, baseH, 2); ctx.fill();

    // Head centred at origin
    ctx.fillStyle = '#1e2d3d'; ctx.strokeStyle = '#253d5a'; ctx.lineWidth = 1;
    rr(-hW / 2, -hD / 2, hW, hD, 3); ctx.fill(); ctx.stroke();

    // Lights: lunar always at top (away from pole), red at bottom (pole side)
    const lx      = 0;
    const lyLunar = -pD * (hD / 2 - bR - 2);
    const lyMid   =  0;
    const lyRed   =  pD * (hD / 2 - bR - 2);

    [lyRed, lyMid, lyLunar].forEach(ly => {
      ctx.beginPath(); ctx.arc(lx, ly, bR, 0, Math.PI * 2);
      ctx.fillStyle = '#050b12'; ctx.fill();
    });

    if (s.state === 'lunar' && lunarOn) {
      ctx.beginPath(); ctx.arc(lx, lyLunar, bR * 2.2, 0, Math.PI * 2);
      ctx.fillStyle = 'rgba(180,215,255,0.15)'; ctx.fill();
      ctx.beginPath(); ctx.arc(lx, lyLunar, bR, 0, Math.PI * 2);
      ctx.fillStyle = 'rgba(220,240,255,0.95)'; ctx.fill();
      ctx.beginPath(); ctx.arc(lx, lyLunar, bR * 0.5, 0, Math.PI * 2);
      ctx.fillStyle = '#f4faff'; ctx.fill();
    }

    if (s.state === 'red') {
      ctx.beginPath(); ctx.arc(lx, lyRed, bR * 2.8, 0, Math.PI * 2);
      ctx.fillStyle = 'rgba(255,30,30,0.12)'; ctx.fill();
      ctx.beginPath(); ctx.arc(lx, lyRed, bR, 0, Math.PI * 2);
      ctx.fillStyle = '#ff2020'; ctx.fill();
      ctx.beginPath(); ctx.arc(lx, lyRed, bR * 0.5, 0, Math.PI * 2);
      ctx.fillStyle = '#ff8888'; ctx.fill();
    }

    ctx.restore();
  }

  // ── Draw train ────────────────────────────────────────────────
  function drawTrain() {
    const pt = trackPoint(trainFrac);

    ctx.save();
    ctx.translate(pt.x, pt.y);
    ctx.rotate(pt.angle);
    // Roof toward inside on all edges except top-going-right (angle≈0)
    if (Math.abs(pt.angle) > 0.01) ctx.scale(1, -1);

    const bx = -TL / 2, by = -TW / 2;

    // Shadow
    ctx.fillStyle = 'rgba(0,0,0,0.22)';
    rr(bx + 2, by + 2, TL, TW, 3); ctx.fill();

    // Body
    const g = ctx.createLinearGradient(bx, by, bx, by + TW);
    g.addColorStop(0,   '#ccd8e4');
    g.addColorStop(0.4, '#b0c2d0');
    g.addColorStop(1,   '#808fa0');
    ctx.fillStyle = g; rr(bx, by, TL, TW, 3); ctx.fill();

    // RTA red stripe
    ctx.fillStyle = '#c8102e';
    ctx.fillRect(bx, by + Math.round(TW * 0.60), TL, 4);

    // Roof shine
    const shine = ctx.createLinearGradient(bx, by, bx, by + TW * 0.35);
    shine.addColorStop(0, 'rgba(255,255,255,0.35)'); shine.addColorStop(1, 'rgba(255,255,255,0)');
    ctx.fillStyle = shine; ctx.fillRect(bx + 3, by, TL - 6, TW * 0.35);

    // Pantograph (extends from roof outward toward overhead wire)
    const panSpread = 5, panRise = 6, barHW = 7;
    ctx.strokeStyle = '#888'; ctx.lineWidth = 1.2; ctx.lineCap = 'round';
    ctx.beginPath(); ctx.moveTo(0, by); ctx.lineTo(-panSpread, by - panRise); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(0, by); ctx.lineTo(+panSpread, by - panRise); ctx.stroke();
    ctx.lineWidth = 1.6;
    ctx.beginPath(); ctx.moveTo(-barHW, by - panRise); ctx.lineTo(+barHW, by - panRise); ctx.stroke();

    // Windows
    ctx.fillStyle = 'rgba(10,28,58,0.9)';
    for (let i = 0; 5 + i * 11 + 8 <= TL - 5; i++) {
      rr(bx + 5 + i * 11, by + 3, 8, 6, 1); ctx.fill();
    }

    // Bogies
    ctx.fillStyle = '#111a24';
    rr(bx + 5,       by + TW, 22, 4, 1); ctx.fill();
    rr(bx + TL - 27, by + TW, 22, 4, 1); ctx.fill();

    // Wheels
    [[bx + 12], [bx + 22], [bx + TL - 22], [bx + TL - 12]].forEach(([wx]) => {
      const wy = by + TW + 6;
      ctx.beginPath(); ctx.arc(wx, wy, 4.5, 0, Math.PI * 2);
      ctx.fillStyle = '#111a24'; ctx.fill();
      ctx.strokeStyle = '#334455'; ctx.lineWidth = 1; ctx.stroke();
      ctx.beginPath(); ctx.arc(wx, wy, 1.5, 0, Math.PI * 2);
      ctx.fillStyle = '#445566'; ctx.fill();
    });

    ctx.restore();
  }

  // ── Draw frame ────────────────────────────────────────────────
  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    drawTrack();
    signals.forEach(drawSignal);
    drawTrain();
  }

  function resize() { canvas.width = window.innerWidth; canvas.height = window.innerHeight; }
  window.addEventListener('resize', resize);
  resize();

  function frame() { update(); draw(); requestAnimationFrame(frame); }
  requestAnimationFrame(frame);
}());
