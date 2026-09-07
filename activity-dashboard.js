import * as THREE from './vendor/three.module.min.js';

const dashboard = document.getElementById('activityDashboard');
const monthSelect = document.getElementById('activityMonth');
const numberFormat = new Intl.NumberFormat();
const monthFormat = new Intl.DateTimeFormat(undefined, { month: 'long', year: 'numeric', timeZone: 'UTC' });
const shortMonthFormat = new Intl.DateTimeFormat(undefined, { month: 'short', timeZone: 'UTC' });
const countryNames = typeof Intl.DisplayNames === 'function'
  ? new Intl.DisplayNames([navigator.language || 'en'], { type: 'region' })
  : null;

const countryCentroids = {
  SG: [1.35, 103.82], MY: [4.21, 101.98], ID: [-2.55, 118.01], PH: [12.88, 121.77],
  TH: [15.87, 100.99], VN: [14.06, 108.28], KH: [12.57, 104.99], MM: [21.91, 95.96],
  BN: [4.54, 114.73], LA: [19.86, 102.5], IN: [20.59, 78.96], BD: [23.68, 90.36],
  PK: [30.38, 69.35], LK: [7.87, 80.77], NP: [28.39, 84.12], CN: [35.86, 104.2],
  HK: [22.32, 114.17], TW: [23.7, 120.96], JP: [36.2, 138.25], KR: [35.91, 127.77],
  AU: [-25.27, 133.78], NZ: [-40.9, 174.89], US: [39.83, -98.58], CA: [56.13, -106.35],
  MX: [23.63, -102.55], BR: [-14.24, -51.93], AR: [-38.42, -63.62], CL: [-35.68, -71.54],
  CO: [4.57, -74.3], PE: [-9.19, -75.02], GB: [55.38, -3.44], IE: [53.14, -7.69],
  FR: [46.23, 2.21], DE: [51.17, 10.45], ES: [40.46, -3.75], PT: [39.4, -8.22],
  IT: [41.87, 12.57], NL: [52.13, 5.29], BE: [50.5, 4.47], CH: [46.82, 8.23],
  AT: [47.52, 14.55], DK: [56.26, 9.5], SE: [60.13, 18.64], NO: [60.47, 8.47],
  FI: [61.92, 25.75], PL: [51.92, 19.15], CZ: [49.82, 15.47], GR: [39.07, 21.82],
  RO: [45.94, 24.97], UA: [48.38, 31.17], RU: [61.52, 105.32], TR: [38.96, 35.24],
  IL: [31.05, 34.85], AE: [23.42, 53.85], SA: [23.89, 45.08], EG: [26.82, 30.8],
  ZA: [-30.56, 22.94], NG: [9.08, 8.68], KE: [-0.02, 37.91], GH: [7.95, -1.02]
};

let latestData = null;
let selectedMonth = '';

function monthDate(key) {
  const [year, month] = key.split('-').map(Number);
  return new Date(Date.UTC(year, month - 1, 1));
}

function countryLabel(code) {
  if (!code || code === 'UN') return 'Location unspecified';
  try { return countryNames?.of(code) || code; }
  catch (_) { return code; }
}

function monthStats(data, key) {
  return data?.months?.[key] || { visits: 0, docxSelected: 0, reviewEdits: 0, exports: 0, countries: {} };
}

function setMetric(id, value) {
  document.getElementById(id).textContent = numberFormat.format(Number(value) || 0);
}

function updateMonthOptions(data) {
  const keys = Object.keys(data.months || {}).sort().reverse();
  const current = new Date().toISOString().slice(0, 7);
  if (!keys.includes(current)) keys.unshift(current);
  if (!selectedMonth || !keys.includes(selectedMonth)) selectedMonth = keys[0] || current;
  monthSelect.innerHTML = keys.map(key =>
    `<option value="${key}"${key === selectedMonth ? ' selected' : ''}>${monthFormat.format(monthDate(key))}</option>`
  ).join('');
}

function updateCountries(countries) {
  const list = document.getElementById('activityCountries');
  const empty = document.getElementById('activityCountriesEmpty');
  const entries = Object.entries(countries || {}).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const max = Math.max(1, ...entries.map(([, count]) => count));
  list.innerHTML = entries.map(([code, count]) => `
    <li class="activity-country-row">
      <span>${countryLabel(code)}</span>
      <span class="activity-country-bar" aria-hidden="true"><span style="width:${Math.max(4, count / max * 100)}%"></span></span>
      <span class="activity-country-count">${numberFormat.format(count)}</span>
    </li>`).join('');
  empty.hidden = entries.length > 0;
}

function updateTrend(data) {
  const trend = document.getElementById('activityTrend');
  const entries = Object.entries(data.months || {}).sort(([a], [b]) => a.localeCompare(b)).slice(-12);
  const max = Math.max(1, ...entries.map(([, stats]) => stats.visits || 0));
  trend.innerHTML = entries.map(([key, stats]) => {
    const height = Math.max(4, (stats.visits || 0) / max * 56);
    return `<button class="activity-trend-button" type="button" data-month="${key}" aria-current="${key === selectedMonth}" aria-label="${monthFormat.format(monthDate(key))}: ${numberFormat.format(stats.visits || 0)} visits">
      <span class="activity-trend-bar" style="height:${height}px"></span>
      <span class="activity-trend-label">${shortMonthFormat.format(monthDate(key))}</span>
    </button>`;
  }).join('');
}

function relativeTime(epochSeconds) {
  const seconds = Math.max(0, Math.round(Date.now() / 1000 - epochSeconds));
  if (seconds < 45) return 'Just now';
  if (seconds < 3600) return `${Math.floor(seconds / 60)}m ago`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)}h ago`;
  return `${Math.floor(seconds / 86400)}d ago`;
}

function eventMessage(entry) {
  const place = entry.country && entry.country !== 'UN' ? ` from ${countryLabel(entry.country)}` : '';
  return {
    visit: `A visitor${place} opened the converter`,
    docx_selected: `A visitor${place} selected a DOCX locally`,
    review_edit: `A visitor${place} refined the review`,
    export: `A visitor${place} exported a QTI package`
  }[entry.event] || `Anonymous activity${place}`;
}

function updateFeed(data) {
  const feed = document.getElementById('activityFeed');
  const empty = document.getElementById('activityFeedEmpty');
  const entries = (data.recent || []).filter(entry =>
    new Date(entry.time * 1000).toISOString().slice(0, 7) === selectedMonth
  ).slice(0, 6);
  feed.innerHTML = entries.map(entry => `<li>
    <span class="activity-feed-dot" aria-hidden="true"></span>
    <span class="activity-feed-time">${relativeTime(entry.time)}</span>
    <span>${eventMessage(entry)}</span>
  </li>`).join('');
  empty.hidden = entries.length > 0;
}

function render(data) {
  latestData = data;
  const liveLabel = document.getElementById('activityLiveLabel');
  const status = document.getElementById('activityStatus');

  if (!data?.ok) {
    liveLabel.textContent = 'Activity service temporarily unavailable';
    status.textContent = 'The converter still works normally; only aggregate analytics are unavailable.';
    ['metricActive', 'metricVisits', 'metricDocx', 'metricEdits', 'metricExports'].forEach(id => {
      document.getElementById(id).textContent = '—';
    });
    return;
  }

  updateMonthOptions(data);
  const stats = monthStats(data, selectedMonth);
  setMetric('metricActive', data.activeNow);
  setMetric('metricVisits', stats.visits);
  setMetric('metricDocx', stats.docxSelected);
  setMetric('metricEdits', stats.reviewEdits);
  setMetric('metricExports', stats.exports);
  updateCountries(stats.countries);
  updateTrend(data);
  updateFeed(data);
  globe?.setCountries(stats.countries);

  const activeText = data.activeNow === 1 ? '1 active visitor' : `${numberFormat.format(data.activeNow)} active visitors`;
  liveLabel.textContent = `${activeText} in the last ${data.activeWindowMinutes || 5} minutes`;
  const updated = data.updated ? new Date(data.updated) : new Date();
  status.textContent = `Aggregate dashboard updated ${updated.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}.`;
}

monthSelect.addEventListener('change', () => {
  selectedMonth = monthSelect.value;
  if (latestData) render(latestData);
});

document.getElementById('activityTrend').addEventListener('click', event => {
  const button = event.target.closest('[data-month]');
  if (!button) return;
  selectedMonth = button.dataset.month;
  monthSelect.value = selectedMonth;
  if (latestData) render(latestData);
});

function latLngToVector3(lat, lng, radius = 1.035) {
  const phi = (90 - lat) * Math.PI / 180;
  const theta = (lng + 180) * Math.PI / 180;
  return new THREE.Vector3(
    -radius * Math.sin(phi) * Math.cos(theta),
    radius * Math.cos(phi),
    radius * Math.sin(phi) * Math.sin(theta)
  );
}

function initGlobe(container) {
  const fallback = document.getElementById('activityGlobeFallback');
  try {
    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(42, 1, 0.1, 100);
    camera.position.z = 3.5;
    const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, powerPreference: 'low-power' });
    renderer.setPixelRatio(Math.min(devicePixelRatio || 1, 1.8));
    renderer.setClearColor(0x000000, 0);
    container.appendChild(renderer.domElement);

    const globeGroup = new THREE.Group();
    globeGroup.rotation.set(0.18, -1.1, -0.06);
    scene.add(globeGroup);

    const earthMaterial = new THREE.MeshPhongMaterial({
      color: 0xdceeff,
      emissive: 0x294c78,
      emissiveIntensity: .55,
      transparent: true,
      opacity: .92,
      shininess: 38
    });
    const sphere = new THREE.Mesh(
      new THREE.SphereGeometry(1, 64, 36),
      earthMaterial
    );
    globeGroup.add(sphere);
    new THREE.TextureLoader().load('./vendor/earth_atmos_2048.jpg', texture => {
      texture.colorSpace = THREE.SRGBColorSpace;
      texture.anisotropy = Math.min(8, renderer.capabilities.getMaxAnisotropy());
      earthMaterial.map = texture;
      earthMaterial.needsUpdate = true;
    });

    const grid = new THREE.Mesh(
      new THREE.SphereGeometry(1.012, 32, 20),
      new THREE.MeshBasicMaterial({ color: 0xdbeafe, transparent: true, opacity: .24, wireframe: true })
    );
    globeGroup.add(grid);

    const atmosphere = new THREE.Mesh(
      new THREE.SphereGeometry(1.08, 48, 28),
      new THREE.MeshBasicMaterial({ color: 0x60a5fa, transparent: true, opacity: .09, side: THREE.BackSide })
    );
    globeGroup.add(atmosphere);

    const starGeometry = new THREE.BufferGeometry();
    const starPositions = [];
    for (let i = 0; i < 150; i++) {
      const v = new THREE.Vector3().randomDirection().multiplyScalar(1.32 + Math.random() * .45);
      starPositions.push(v.x, v.y, v.z);
    }
    starGeometry.setAttribute('position', new THREE.Float32BufferAttribute(starPositions, 3));
    globeGroup.add(new THREE.Points(starGeometry, new THREE.PointsMaterial({ color: 0x60a5fa, size: .018, transparent: true, opacity: .45 })));

    scene.add(new THREE.HemisphereLight(0xffffff, 0x2563eb, 2.4));
    const keyLight = new THREE.DirectionalLight(0xffffff, 2.8);
    keyLight.position.set(-3, 4, 5);
    scene.add(keyLight);

    let markerGroup = new THREE.Group();
    globeGroup.add(markerGroup);

    function setCountries(countries = {}) {
      globeGroup.remove(markerGroup);
      markerGroup.traverse(object => {
        object.geometry?.dispose?.();
        object.material?.dispose?.();
      });
      markerGroup = new THREE.Group();
      const entries = Object.entries(countries).filter(([code]) => countryCentroids[code]);
      const max = Math.max(1, ...entries.map(([, count]) => count));
      entries.forEach(([code, count], index) => {
        const [lat, lng] = countryCentroids[code];
        const size = .018 + Math.sqrt(count / max) * .03;
        const marker = new THREE.Mesh(
          new THREE.SphereGeometry(size, 18, 12),
          new THREE.MeshStandardMaterial({ color: 0x0ea5e9, emissive: 0x2563eb, emissiveIntensity: 1.6, roughness: .25 })
        );
        marker.position.copy(latLngToVector3(lat, lng));
        marker.userData = { base: 1, phase: index * .8 };
        markerGroup.add(marker);
      });
      globeGroup.add(markerGroup);
    }

    let dragging = false;
    let lastX = 0;
    let lastY = 0;
    container.addEventListener('pointerdown', event => {
      dragging = true;
      lastX = event.clientX;
      lastY = event.clientY;
      container.setPointerCapture?.(event.pointerId);
    });
    container.addEventListener('pointermove', event => {
      if (!dragging) return;
      globeGroup.rotation.y += (event.clientX - lastX) * .008;
      globeGroup.rotation.x = THREE.MathUtils.clamp(globeGroup.rotation.x + (event.clientY - lastY) * .006, -.85, .85);
      lastX = event.clientX;
      lastY = event.clientY;
    });
    const stopDrag = () => { dragging = false; };
    container.addEventListener('pointerup', stopDrag);
    container.addEventListener('pointercancel', stopDrag);
    container.addEventListener('wheel', event => {
      event.preventDefault();
      camera.position.z = THREE.MathUtils.clamp(camera.position.z + event.deltaY * .002, 2.8, 5);
    }, { passive: false });
    container.addEventListener('keydown', event => {
      const step = .12;
      if (event.key === 'ArrowLeft') globeGroup.rotation.y -= step;
      else if (event.key === 'ArrowRight') globeGroup.rotation.y += step;
      else if (event.key === 'ArrowUp') globeGroup.rotation.x -= step;
      else if (event.key === 'ArrowDown') globeGroup.rotation.x += step;
      else return;
      event.preventDefault();
    });

    function resize() {
      const width = Math.max(1, container.clientWidth);
      const height = Math.max(1, container.clientHeight);
      renderer.setSize(width, height, false);
      camera.aspect = width / height;
      camera.updateProjectionMatrix();
    }
    new ResizeObserver(resize).observe(container);
    resize();

    const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
    let visible = true;
    new IntersectionObserver(entries => { visible = entries[0]?.isIntersecting ?? true; }, { rootMargin: '120px' }).observe(container);

    renderer.setAnimationLoop(time => {
      if (!visible) return;
      if (!dragging && !reducedMotion) globeGroup.rotation.y += .0012;
      markerGroup.children.forEach(marker => {
        const pulse = reducedMotion ? 1 : 1 + Math.sin(time * .003 + marker.userData.phase) * .18;
        marker.scale.setScalar(pulse);
      });
      renderer.render(scene, camera);
    });

    fallback.hidden = true;
    return { setCountries };
  } catch (_) {
    fallback.hidden = false;
    container.setAttribute('aria-label', 'Globe fallback; WebGL is unavailable');
    return { setCountries() {} };
  }
}

const globe = initGlobe(document.getElementById('activityGlobe'));
window.qtiActivity.subscribe(render);
window.qtiActivity.start().then(data => { if (data) render(data); });

setInterval(() => {
  if (document.visibilityState === 'visible') window.qtiActivity.track('heartbeat');
}, 60000);

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') window.qtiActivity.track('heartbeat');
});

if (!dashboard) throw new Error('Activity dashboard markup is missing.');
