import * as THREE from './vendor/three.module.min.js';

const mount = document.getElementById('conversionHero3d');
const stateLabel = document.getElementById('conversionHeroState');
const replayButton = document.getElementById('heroReplayBtn');
const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

const stateCopy = {
  idle: 'A private, browser-only transformation. No document is uploaded.',
  ready: 'Document ready — choose the paper type, then convert.',
  converting: 'Parsing questions, answer choices, tables and images…',
  review: 'Questions structured — human review completes the final 25%.',
  exported: 'QTI 2.1 package built and ready for SLS import.',
  error: 'The document needs attention. Nothing has left your browser.'
};

if (mount) {
  let renderer;
  let frameId = 0;
  let isVisible = true;
  let currentState = 'idle';
  let stateChangedAt = performance.now();
  let pointerX = 0;
  let pointerY = 0;
  let keyboardYaw = 0;
  let replayTimers = [];

  try {
    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(38, 1, 0.1, 100);
    camera.position.set(0, 0.25, 11.2);

    renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, powerPreference: 'high-performance' });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.5));
    renderer.outputColorSpace = THREE.SRGBColorSpace;
    renderer.setClearColor(0x071a2b, 0);
    renderer.domElement.setAttribute('aria-hidden', 'true');
    mount.prepend(renderer.domElement);

    const world = new THREE.Group();
    world.rotation.x = -0.035;
    scene.add(world);

    scene.add(new THREE.HemisphereLight(0xc9edff, 0x07111f, 2.25));
    const keyLight = new THREE.DirectionalLight(0xffffff, 3.1);
    keyLight.position.set(-3, 5, 7);
    scene.add(keyLight);
    const blueLight = new THREE.PointLight(0x38bdf8, 30, 9);
    blueLight.position.set(0, 0, 2.5);
    scene.add(blueLight);
    const mintLight = new THREE.PointLight(0x34d399, 22, 7);
    mintLight.position.set(3.5, 0, 2);
    scene.add(mintLight);

    const grid = new THREE.GridHelper(18, 24, 0x1d4ed8, 0x173451);
    grid.position.y = -2.3;
    grid.rotation.x = 0.015;
    grid.material.transparent = true;
    grid.material.opacity = 0.3;
    world.add(grid);

    const documentGroup = new THREE.Group();
    const paperMaterial = new THREE.MeshPhysicalMaterial({ color: 0xf8fafc, roughness: 0.42, metalness: 0.05 });
    const inkMaterial = new THREE.MeshBasicMaterial({ color: 0x94a3b8 });
    const wordMaterial = new THREE.MeshBasicMaterial({ color: 0x2563eb });
    for (let pageIndex = 0; pageIndex < 3; pageIndex += 1) {
      const page = new THREE.Group();
      page.add(new THREE.Mesh(new THREE.BoxGeometry(1.25, 1.72, 0.09), paperMaterial));
      const wordTab = new THREE.Mesh(new THREE.BoxGeometry(0.3, 0.34, 0.04), wordMaterial);
      wordTab.position.set(-0.37, 0.5, 0.07);
      page.add(wordTab);
      for (let lineIndex = 0; lineIndex < 5; lineIndex += 1) {
        const width = lineIndex === 0 ? 0.58 : (0.82 - lineIndex * 0.045);
        const line = new THREE.Mesh(new THREE.BoxGeometry(width, 0.045, 0.035), inkMaterial);
        line.position.set(0.07, 0.38 - lineIndex * 0.23, 0.07);
        page.add(line);
      }
      page.position.set(-3.65 + pageIndex * 0.42, (pageIndex - 1) * 0.08, -pageIndex * 0.24);
      page.rotation.set(0.02, -0.22 + pageIndex * 0.08, -0.06 + pageIndex * 0.055);
      documentGroup.add(page);
    }
    world.add(documentGroup);

    const core = new THREE.Group();
    const ringMaterial = new THREE.MeshPhysicalMaterial({ color: 0x2563eb, emissive: 0x0c4ad7, emissiveIntensity: 1.35, roughness: 0.25, metalness: 0.55 });
    const innerRing = new THREE.Mesh(new THREE.TorusGeometry(1.27, 0.16, 18, 72), ringMaterial);
    const outerRing = new THREE.Mesh(new THREE.TorusGeometry(1.6, 0.055, 12, 72), new THREE.MeshBasicMaterial({ color: 0x38bdf8, transparent: true, opacity: 0.75 }));
    core.add(innerRing, outerRing);
    const hub = new THREE.Mesh(new THREE.IcosahedronGeometry(0.46, 1), new THREE.MeshPhysicalMaterial({ color: 0x7dd3fc, emissive: 0x2563eb, emissiveIntensity: 1.4, roughness: 0.25, metalness: 0.3, transparent: true, opacity: 0.82 }));
    core.add(hub);
    world.add(core);

    const qtiGroup = new THREE.Group();
    const qtiMaterial = new THREE.MeshPhysicalMaterial({ color: 0xffffff, roughness: 0.38, metalness: 0.04 });
    const mintMaterial = new THREE.MeshBasicMaterial({ color: 0x34d399 });
    for (let cardIndex = 0; cardIndex < 3; cardIndex += 1) {
      const card = new THREE.Group();
      card.add(new THREE.Mesh(new THREE.BoxGeometry(1.52, 0.92, 0.1), qtiMaterial));
      const tag = new THREE.Mesh(new THREE.BoxGeometry(0.42, 0.17, 0.04), mintMaterial);
      tag.position.set(-0.43, 0.27, 0.075);
      card.add(tag);
      for (let lineIndex = 0; lineIndex < 3; lineIndex += 1) {
        const line = new THREE.Mesh(new THREE.BoxGeometry(0.67 - lineIndex * 0.07, 0.035, 0.03), inkMaterial);
        line.position.set(0.05, 0.12 - lineIndex * 0.19, 0.075);
        card.add(line);
      }
      card.position.set(3.15 + cardIndex * 0.24, 0.98 - cardIndex * 0.92, -cardIndex * 0.18);
      card.rotation.y = 0.16 - cardIndex * 0.05;
      qtiGroup.add(card);
    }
    const packageCube = new THREE.Mesh(new THREE.BoxGeometry(0.78, 0.78, 0.78), new THREE.MeshPhysicalMaterial({ color: 0x10b981, emissive: 0x065f46, emissiveIntensity: 0.45, roughness: 0.32, metalness: 0.15 }));
    packageCube.position.set(4.5, -1.25, -0.2);
    qtiGroup.add(packageCube);
    world.add(qtiGroup);

    const particleCount = 70;
    const particles = new THREE.InstancedMesh(new THREE.BoxGeometry(0.07, 0.07, 0.07), new THREE.MeshBasicMaterial({ color: 0x7dd3fc }), particleCount);
    particles.instanceMatrix.setUsage(THREE.DynamicDrawUsage);
    world.add(particles);
    const particleDummy = new THREE.Object3D();
    const particleSeeds = Array.from({ length: particleCount }, (_, index) => ({ offset: index / particleCount, lane: ((index % 9) - 4) / 4, depth: ((index * 7) % 13 - 6) / 10, spin: 0.45 + (index % 5) * 0.16 }));

    function resize() {
      const rect = mount.getBoundingClientRect();
      if (!rect.width || !rect.height) return;
      renderer.setSize(rect.width, rect.height, false);
      camera.aspect = rect.width / rect.height;
      camera.position.z = rect.width < 700 ? 13.8 : 11.2;
      camera.updateProjectionMatrix();
    }

    function stateSpeed() {
      if (currentState === 'converting') return 0.32;
      if (currentState === 'ready') return 0.15;
      if (currentState === 'review' || currentState === 'exported') return 0.11;
      return 0.075;
    }

    function updateParticles(time) {
      const speed = stateSpeed();
      particleSeeds.forEach((seed, index) => {
        const progress = (seed.offset + time * speed) % 1;
        const focus = 1 - Math.abs(progress - 0.5) * 2;
        particleDummy.position.set(-2.85 + progress * 5.9, seed.lane * (0.84 - focus * 0.58) + Math.sin(time * 2 + index) * 0.035, seed.depth * (0.68 - focus * 0.42));
        particleDummy.rotation.set(time * seed.spin, time * seed.spin * 0.8, time * seed.spin * 0.55);
        particleDummy.scale.setScalar(0.65 + focus * 0.65);
        particleDummy.updateMatrix();
        particles.setMatrixAt(index, particleDummy.matrix);
      });
      particles.instanceMatrix.needsUpdate = true;
    }

    function render(timeMs = 0) {
      const time = timeMs / 1000;
      const stateAge = Math.max(0, (timeMs - stateChangedAt) / 1000);
      updateParticles(time);
      core.rotation.z = time * (currentState === 'converting' ? 1.25 : 0.34);
      innerRing.rotation.x = Math.sin(time * 0.45) * 0.13;
      outerRing.rotation.y = time * -0.22;
      hub.rotation.set(time * 0.42, time * 0.58, time * 0.28);
      documentGroup.position.y = Math.sin(time * 0.8) * 0.055;
      qtiGroup.position.y = Math.sin(time * 0.72 + 1.4) * 0.055;
      packageCube.rotation.set(time * 0.36, time * 0.55, 0.12);
      const resolve = currentState === 'review' || currentState === 'exported';
      qtiGroup.scale.setScalar(resolve ? Math.min(1.07, 1 + stateAge * 0.08) : 1);
      ringMaterial.emissiveIntensity = currentState === 'converting' ? 2.45 : 1.35;
      packageCube.material.emissiveIntensity = currentState === 'exported' ? 1.7 : 0.45;
      world.rotation.y += ((pointerX * 0.075 + keyboardYaw) - world.rotation.y) * 0.045;
      world.rotation.x += ((-0.035 + pointerY * 0.045) - world.rotation.x) * 0.045;
      renderer.render(scene, camera);
      if (isVisible && !reducedMotion) frameId = requestAnimationFrame(render);
    }

    function showState(state, detail = {}) {
      currentState = stateCopy[state] ? state : 'idle';
      stateChangedAt = performance.now();
      mount.dataset.state = currentState;
      const countText = detail.questionCount ? ` ${detail.questionCount} questions are ready.` : '';
      if (stateLabel) stateLabel.textContent = stateCopy[currentState] + countText;
      if (reducedMotion) render(performance.now());
    }

    function clearReplayTimers() { replayTimers.forEach(window.clearTimeout); replayTimers = []; }
    function replayFlow() {
      clearReplayTimers();
      showState('ready');
      replayTimers.push(window.setTimeout(() => showState('converting'), 700));
      replayTimers.push(window.setTimeout(() => showState('review', { questionCount: 12 }), 2600));
      replayTimers.push(window.setTimeout(() => showState('exported'), 3900));
      replayTimers.push(window.setTimeout(() => showState('idle'), 5700));
    }

    window.addEventListener('qti-conversion-state', (event) => { clearReplayTimers(); showState(event.detail?.state || 'idle', event.detail || {}); });
    replayButton?.addEventListener('click', replayFlow);
    mount.addEventListener('pointermove', (event) => {
      const rect = mount.getBoundingClientRect();
      pointerX = ((event.clientX - rect.left) / rect.width - 0.5) * 2;
      pointerY = ((event.clientY - rect.top) / rect.height - 0.5) * 2;
    });
    mount.addEventListener('pointerleave', () => { pointerX = 0; pointerY = 0; });
    mount.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      keyboardYaw = THREE.MathUtils.clamp(keyboardYaw + (event.key === 'ArrowLeft' ? -0.08 : 0.08), -0.28, 0.28);
      if (reducedMotion) render(performance.now());
    });

    new IntersectionObserver((entries) => {
      isVisible = entries[0]?.isIntersecting ?? true;
      if (isVisible && !reducedMotion && !frameId) frameId = requestAnimationFrame(render);
      if (!isVisible && frameId) { cancelAnimationFrame(frameId); frameId = 0; }
    }, { threshold: 0.08 }).observe(mount);
    new ResizeObserver(resize).observe(mount);
    resize();
    showState('idle');
    render(performance.now());
  } catch (error) {
    console.warn('3D conversion preview unavailable:', error);
    mount.classList.add('is-fallback');
    if (stateLabel) stateLabel.textContent = stateCopy.idle;
  }
}
