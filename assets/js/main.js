'use strict';

(function () {
  document.documentElement.classList.add('js-on');

  const nav = document.getElementById('nav');
  const menuToggle = document.getElementById('menuToggle');
  const navLinks = document.getElementById('navLinks');
  const motionAllowed = !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const finePointer = window.matchMedia('(pointer: fine)').matches;
  const richMotionAllowed = motionAllowed && finePointer;

  function initNavigation() {
    if (nav) {
      const updateNavState = () => {
        nav.classList.toggle('scrolled', window.scrollY > 24);
      };

      window.addEventListener('scroll', updateNavState, { passive: true });
      updateNavState();
    }

    if (!menuToggle || !navLinks) return;

    const backdrop = document.createElement('div');
    backdrop.className = 'nav-backdrop';
    document.body.appendChild(backdrop);

    const serviceDropdowns = navLinks.querySelectorAll('.nav-dropdown');

    const closeServiceDropdowns = () => {
      serviceDropdowns.forEach((dropdown) => {
        dropdown.classList.remove('open');
        const toggle = dropdown.querySelector('.nav-dropdown-toggle');
        if (toggle) toggle.setAttribute('aria-expanded', 'false');
      });
    };

    const setMenuOpen = (open) => {
      navLinks.classList.toggle('open', open);
      menuToggle.classList.toggle('active', open);
      menuToggle.setAttribute('aria-expanded', String(open));
      backdrop.classList.toggle('active', open);
      document.body.style.overflow = open ? 'hidden' : '';
      if (!open) closeServiceDropdowns();
    };

    menuToggle.addEventListener('click', () => {
      setMenuOpen(!navLinks.classList.contains('open'));
    });

    serviceDropdowns.forEach((dropdown) => {
      const toggle = dropdown.querySelector('.nav-dropdown-toggle');
      if (!toggle) return;

      toggle.addEventListener('click', () => {
        const nextOpen = !dropdown.classList.contains('open');
        closeServiceDropdowns();
        dropdown.classList.toggle('open', nextOpen);
        toggle.setAttribute('aria-expanded', String(nextOpen));
      });
    });

    backdrop.addEventListener('click', () => setMenuOpen(false));

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && navLinks.classList.contains('open')) {
        setMenuOpen(false);
        menuToggle.focus();
      }
    });

    navLinks.querySelectorAll('a').forEach((link) => {
      link.addEventListener('click', () => setMenuOpen(false));
    });
  }

  function initCurrentYear() {
    const yearEl = document.getElementById('year');
    if (yearEl) yearEl.textContent = new Date().getFullYear();
  }

  function initReveals() {
    const revealTargets = document.querySelectorAll('.reveal, .reveal-stagger');
    if (!revealTargets.length) return;

    document.documentElement.classList.add('stagger-ready');

    if (!('IntersectionObserver' in window)) {
      revealTargets.forEach((target) => target.classList.add('in'));
      return;
    }

    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        entry.target.classList.add('in');
        observer.unobserve(entry.target);
      });
    }, {
      threshold: 0.12,
      rootMargin: '0px 0px -60px 0px',
    });

    revealTargets.forEach((target) => observer.observe(target));
  }

  function initCardTilt() {
    if (!richMotionAllowed) return;

    document.querySelectorAll('.service-pillar, .info-card').forEach((card) => {
      card.addEventListener('mousemove', (event) => {
        const rect = card.getBoundingClientRect();
        const x = (event.clientX - rect.left) / rect.width - 0.5;
        const y = (event.clientY - rect.top) / rect.height - 0.5;

        card.style.transform =
          `perspective(900px) rotateX(${-y * 3.5}deg) rotateY(${x * 4.5}deg) translateY(-5px)`;
      });

      card.addEventListener('mouseleave', () => {
        card.style.transform = '';
      });
    });
  }

  function animateCounter(counter) {
    const original = counter.textContent.trim();
    const match = original.match(/^([\d,]+)(.*)/);
    if (!match) return;

    const target = parseInt(match[1].replace(/,/g, ''), 10);
    const suffix = match[2];
    if (!target) return;

    const duration = 1400;
    const startTime = performance.now();
    const easeOut = (value) => 1 - Math.pow(1 - value, 3);

    const step = (now) => {
      const progress = Math.min((now - startTime) / duration, 1);
      counter.textContent = Math.round(easeOut(progress) * target) + suffix;

      if (progress < 1) {
        requestAnimationFrame(step);
      } else {
        counter.textContent = original;
        counter.classList.add('count-done');
      }
    };

    requestAnimationFrame(step);
  }

  function initCounters() {
    const counters = document.querySelectorAll('.stat strong');
    if (!motionAllowed || !counters.length || !('IntersectionObserver' in window)) return;

    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        animateCounter(entry.target);
        observer.unobserve(entry.target);
      });
    }, { threshold: 0.6 });

    counters.forEach((counter) => observer.observe(counter));
  }

  function initHeroPhotoScrollParallax() {
    if (!motionAllowed) return;

    const heroImages = document.querySelectorAll('.hero-visual--photo .hero-real-img');
    if (!heroImages.length) return;

    let ticking = false;

    window.addEventListener('scroll', () => {
      if (ticking) return;
      ticking = true;

      requestAnimationFrame(() => {
        heroImages.forEach((image) => {
          const hero = image.closest('.page-hero');
          if (!hero) return;

          const rect = hero.getBoundingClientRect();
          if (rect.bottom < 0 || rect.top > window.innerHeight) return;

          const offset = (rect.top / window.innerHeight) * 0.14 * 60;
          image.style.transform = `scale(1.08) translateY(${offset}px)`;
        });

        ticking = false;
      });
    }, { passive: true });
  }

  function initHeroMouseParallax() {
    if (!richMotionAllowed) return;

    document.querySelectorAll('.page-hero:not(.page-hero--centered)').forEach((section) => {
      const visual = section.querySelector('.hero-visual--photo');
      if (!visual) return;

      section.addEventListener('mousemove', (event) => {
        const rect = section.getBoundingClientRect();
        const x = (event.clientX - rect.left - rect.width / 2) / rect.width;
        const y = (event.clientY - rect.top - rect.height / 2) / rect.height;

        visual.style.transform = `translateX(${x * 12}px) translateY(${y * 8}px)`;
      });

      section.addEventListener('mouseleave', () => {
        visual.style.transform = '';
      });
    });
  }

  function initFaqAccordion() {
    document.querySelectorAll('.faq-question').forEach((btn) => {
      btn.addEventListener('click', () => {
        const item = btn.closest('.faq-item');
        const isOpen = item.classList.contains('open');
        document.querySelectorAll('.faq-item.open').forEach((el) => {
          el.classList.remove('open');
          el.querySelector('.faq-question').setAttribute('aria-expanded', 'false');
        });
        if (!isOpen) {
          item.classList.add('open');
          btn.setAttribute('aria-expanded', 'true');
        }
      });
    });
  }

  initNavigation();
  initCurrentYear();
  initReveals();
  initCardTilt();
  initCounters();
  initHeroPhotoScrollParallax();
  initHeroMouseParallax();
  initFaqAccordion();
})();
