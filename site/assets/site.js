/* officelens — progressive enhancement only.
   Everything here is optional: the site is fully usable with JS disabled.
   This file adds a persisted light/dark theme toggle, copy-to-clipboard
   buttons, a nav shadow on scroll, scroll reveal, and the mobile nav.
   No network access. */
(function () {
  "use strict";

  var THEME_KEY = "officelens-theme";

  function prefersReducedMotion() {
    return (
      window.matchMedia &&
      window.matchMedia("(prefers-reduced-motion: reduce)").matches
    );
  }

  /* ----- Theme -----------------------------------------------------------
     The stylesheet already honours prefers-color-scheme. This adds an
     explicit [data-theme] override on <html>, remembered in localStorage. */
  function readStoredTheme() {
    try {
      var value = window.localStorage.getItem(THEME_KEY);
      return value === "dark" || value === "light" ? value : null;
    } catch (error) {
      return null;
    }
  }

  function storeTheme(theme) {
    try {
      window.localStorage.setItem(THEME_KEY, theme);
    } catch (error) {
      /* storage may be unavailable — the toggle still works for this page */
    }
  }

  function systemTheme() {
    if (
      window.matchMedia &&
      window.matchMedia("(prefers-color-scheme: light)").matches
    ) {
      return "light";
    }
    return "dark";
  }

  function currentTheme() {
    var explicit = document.documentElement.getAttribute("data-theme");
    return explicit === "dark" || explicit === "light"
      ? explicit
      : systemTheme();
  }

  function updateThemeButton(theme) {
    var button = document.getElementById("theme-toggle");
    if (!button) return;
    var next = theme === "dark" ? "light" : "dark";
    button.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
    button.setAttribute("aria-label", "Switch to " + next + " theme");
    button.setAttribute("title", "Switch to " + next + " theme");
  }

  function applyTheme(theme, persist) {
    document.documentElement.setAttribute("data-theme", theme);
    if (persist) storeTheme(theme);
    updateThemeButton(theme);
  }

  function initTheme() {
    var stored = readStoredTheme();
    if (stored) applyTheme(stored, false);
    updateThemeButton(stored || systemTheme());

    var button = document.getElementById("theme-toggle");
    if (!button) return;
    button.addEventListener("click", function () {
      applyTheme(currentTheme() === "dark" ? "light" : "dark", true);
    });
  }

  /* ----- Clipboard ------------------------------------------------------- */
  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text);
    }

    return new Promise(function (resolve, reject) {
      var area = document.createElement("textarea");
      area.value = text;
      area.setAttribute("readonly", "");
      area.style.position = "fixed";
      area.style.top = "-1000px";
      document.body.appendChild(area);
      area.select();

      try {
        var ok = document.execCommand("copy");
        document.body.removeChild(area);
        ok ? resolve() : reject(new Error("copy failed"));
      } catch (err) {
        document.body.removeChild(area);
        reject(err);
      }
    });
  }

  function canCopy() {
    if (navigator.clipboard && window.isSecureContext) {
      return true;
    }
    return (
      typeof document.queryCommandSupported === "function" &&
      document.queryCommandSupported("copy")
    );
  }

  function flashCopied(button, ok) {
    var original = button.getAttribute("data-label") || button.textContent;
    button.setAttribute("data-label", original);
    button.textContent = ok ? "Copied" : "Copy failed";
    button.setAttribute("data-copied", ok ? "true" : "false");

    window.clearTimeout(button._resetTimer);
    button._resetTimer = window.setTimeout(function () {
      button.textContent = original;
      button.removeAttribute("data-copied");
    }, 1600);
  }

  function wireCopy(button, getText) {
    button.addEventListener("click", function () {
      copyText(getText()).then(
        function () {
          flashCopied(button, true);
        },
        function () {
          flashCopied(button, false);
        }
      );
    });
  }

  function enhanceCodeBlocks() {
    var blocks = document.querySelectorAll("pre.code");
    Array.prototype.forEach.call(blocks, function (pre) {
      var code = pre.querySelector("code") || pre;
      var text = code.textContent;

      var button = document.createElement("button");
      button.type = "button";
      button.className = "code-copy";
      button.textContent = "Copy";
      button.setAttribute("aria-label", "Copy code to clipboard");
      wireCopy(button, function () {
        return text;
      });

      pre.appendChild(button);
    });
  }

  function enhanceChips() {
    var buttons = document.querySelectorAll("[data-copy]");
    Array.prototype.forEach.call(buttons, function (button) {
      var target = button.getAttribute("data-copy");
      wireCopy(button, function () {
        if (target && target.charAt(0) === "#") {
          var node = document.querySelector(target);
          return node ? node.textContent.trim() : "";
        }
        return target || "";
      });
    });
  }

  /* ----- Nav shadow on scroll ------------------------------------------- */
  function initScrollShadow() {
    var header = document.querySelector(".site-header");
    if (!header) return;

    function update() {
      if (window.scrollY > 8) {
        header.classList.add("is-scrolled");
      } else {
        header.classList.remove("is-scrolled");
      }
    }

    update();
    window.addEventListener("scroll", update, { passive: true });
  }

  /* ----- Mobile nav ------------------------------------------------------ */
  function initMobileNav() {
    var toggle = document.getElementById("nav-toggle");
    var nav = document.getElementById("primary-nav");
    if (!toggle || !nav) return;

    function setOpen(open) {
      nav.classList.toggle("is-open", open);
      toggle.setAttribute("aria-expanded", open ? "true" : "false");
      toggle.setAttribute("aria-label", open ? "Close navigation" : "Open navigation");
    }

    toggle.addEventListener("click", function () {
      setOpen(!nav.classList.contains("is-open"));
    });

    nav.addEventListener("click", function (event) {
      if (event.target && event.target.tagName === "A") setOpen(false);
    });

    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape") setOpen(false);
    });

    document.addEventListener("click", function (event) {
      if (!nav.classList.contains("is-open")) return;
      if (nav.contains(event.target) || toggle.contains(event.target)) return;
      setOpen(false);
    });
  }

  /* ----- Scroll reveal ---------------------------------------------------
     Content is visible by default, so it can never be stranded hidden when
     JS is off or IntersectionObserver is missing. Only when we can animate do
     we hide an element, then reveal it as it enters the viewport. */
  function initReveal() {
    var nodes = document.querySelectorAll(".reveal");
    if (!nodes.length) return;
    if (prefersReducedMotion() || !("IntersectionObserver" in window)) return;

    var observer = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) return;
          entry.target.classList.remove("reveal--hidden");
          observer.unobserve(entry.target);
        });
      },
      { rootMargin: "0px 0px -8% 0px", threshold: 0.05 }
    );

    Array.prototype.forEach.call(nodes, function (node) {
      node.classList.add("reveal--hidden");
      observer.observe(node);
    });
  }

  function init() {
    initTheme();
    if (canCopy()) {
      enhanceCodeBlocks();
      enhanceChips();
    }
    initScrollShadow();
    initMobileNav();
    initReveal();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
