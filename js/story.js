/* 스크롤 스토리텔링: 히어로 패럴랙스, 문장 하이라이트, 기능 챕터 단계 전환, 상단 탭 위치 표시 */
(() => {
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const clamp = (v, a = 0, b = 1) => Math.min(b, Math.max(a, v));

  /* ---------- 소개 문장을 어절 단위로 나눔 ---------- */
  const statement = document.getElementById('statement');
  let words = [];
  if (statement) {
    const text = statement.textContent.trim().replace(/\s+/g, ' ');
    statement.setAttribute('aria-label', text);
    statement.innerHTML = text.split(' ').map((w) => `<span class="w" aria-hidden="true">${w}</span>`).join(' ');
    words = [...statement.querySelectorAll('.w')];
  }

  /* ---------- 요소 ---------- */
  const heroWord = document.querySelector('.hero__word');
  const chapters = [...document.querySelectorAll('.chapter')].map((el) => ({
    el,
    steps: [...el.querySelectorAll('.steps li')],
    frames: [...el.querySelectorAll('.frame')],
    current: -1,
  }));
  const seg = document.querySelector('.seg');
  const thumb = document.getElementById('segThumb');
  const tabs = [...document.querySelectorAll('.seg .tab')];
  const zones = [...document.querySelectorAll('[data-tab]:not(.tab)')];
  let activeTab = null;

  function setTab(t) {
    if (t === activeTab) return;
    activeTab = t;
    tabs.forEach((a) => {
      const on = a.dataset.tab === t;
      a.classList.toggle('on', on);
      if (on) a.setAttribute('aria-current', 'true'); else a.removeAttribute('aria-current');
    });
    seg.classList.toggle('has-active', !!t);
    placeThumb();
  }
  function placeThumb() {
    const on = tabs.find((a) => a.dataset.tab === activeTab);
    if (!on) return;
    thumb.style.width = on.offsetWidth + 'px';
    thumb.style.transform = `translateX(${on.offsetLeft - 4}px)`;
  }

  function setStep(ch, i) {
    if (ch.current === i) return;
    ch.current = i;
    ch.steps.forEach((s, k) => s.classList.toggle('is-active', k === i));
    ch.frames.forEach((f, k) => f.classList.toggle('is-active', k === i));
  }

  /* ---------- 매 프레임 ---------- */
  function frame() {
    const vh = window.innerHeight;

    // 히어로 단어: 스크롤할수록 살짝 올라가며 작아짐
    if (heroWord && !reduce) {
      const p = clamp(window.scrollY / vh);
      heroWord.style.transform = `translateY(${-60 * p}px) scale(${1 - 0.06 * p})`;
      heroWord.style.opacity = String(1 - 0.5 * p);
    }

    // 소개 문장: 화면 가운데를 지나는 만큼 어절이 밝아짐
    if (words.length) {
      const r = statement.getBoundingClientRect();
      const p = reduce ? 1 : clamp((vh * 0.8 - r.top) / (r.height + vh * 0.35));
      const lit = Math.round(p * words.length);
      words.forEach((w, i) => w.classList.toggle('lit', i < lit));
    }

    // 챕터: 고정 구간 진행률을 3단계로
    chapters.forEach((ch) => {
      const r = ch.el.getBoundingClientRect();
      const p = clamp(-r.top / Math.max(1, ch.el.offsetHeight - vh));
      setStep(ch, Math.min(ch.steps.length - 1, Math.floor(p * ch.steps.length * 0.999)));
    });

    // 상단 탭: 화면 40% 지점에 걸린 기능
    const probe = vh * 0.4;
    const z = zones.find((el) => { const r = el.getBoundingClientRect(); return r.top <= probe && r.bottom > probe; });
    setTab(z ? z.dataset.tab : null);
  }

  let ticking = false;
  const onScroll = () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => { frame(); ticking = false; });
  };
  window.addEventListener('scroll', onScroll, { passive: true });
  window.addEventListener('resize', () => { placeThumb(); onScroll(); });
  if (document.fonts) document.fonts.ready.then(placeThumb);

  // 탭 클릭 시 해당 기능으로 부드럽게 이동 (기본 앵커 이동 대신)
  tabs.forEach((a) => a.addEventListener('click', (e) => {
    const el = document.querySelector(a.getAttribute('href'));
    if (!el) return;
    e.preventDefault();
    el.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'start' });
    history.replaceState(null, '', a.getAttribute('href'));
  }));

  frame();
})();
