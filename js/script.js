(function () {
  "use strict";

  // === Dark Mode ===
  const themeToggle = document.getElementById("theme-toggle");
  const html = document.documentElement;

  function getTheme() {
    const stored = localStorage.getItem("theme");
    if (stored) return stored;
    return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }

  function setTheme(theme) {
    if (theme === "dark") {
      html.setAttribute("data-theme", "dark");
      if (themeToggle) themeToggle.textContent = "☀";
    } else {
      html.removeAttribute("data-theme");
      if (themeToggle) themeToggle.textContent = "☾";
    }
    localStorage.setItem("theme", theme);
  }

  setTheme(getTheme());

  if (themeToggle) {
    themeToggle.addEventListener("click", function () {
      const next = html.hasAttribute("data-theme") ? "light" : "dark";
      setTheme(next);
    });
  }

  // === Stats Counter ===
  function animateCount(el, target, duration) {
    var start = 0;
    var step = Math.max(1, Math.floor(target / (duration / 16)));
    var timer = setInterval(function () {
      start += step;
      if (start >= target) {
        el.textContent = target;
        clearInterval(timer);
      } else {
        el.textContent = start;
      }
    }, 16);
  }

  var statsPanel = document.getElementById("stats-panel");
  if (statsPanel) {
    var observer = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            var numbers = statsPanel.querySelectorAll("[data-count]");
            numbers.forEach(function (n) {
              var target = parseInt(n.getAttribute("data-count"), 10);
              if (target) animateCount(n, target, 1200);
            });
            observer.unobserve(statsPanel);
          }
        });
      },
      { threshold: 0.5 }
    );
    observer.observe(statsPanel);
  }

  // === Filter + Search ===
  var filterBtns = document.querySelectorAll(".filter-btn");
  var searchInput = document.getElementById("search-input");
  var bookGrid = document.getElementById("book-grid");
  var emptyState = document.getElementById("empty-state");
  var cards = bookGrid ? bookGrid.querySelectorAll(".book-card") : [];

  var activeFilter = "all";
  var searchQuery = "";

  function updateCards() {
    var visibleCount = 0;
    cards.forEach(function (card) {
      var cat = card.getAttribute("data-category") || "";
      var text = (card.getAttribute("data-search") || "").toLowerCase();
      var matchFilter = activeFilter === "all" || cat === activeFilter;
      var matchSearch = !searchQuery || text.indexOf(searchQuery.toLowerCase()) !== -1;

      if (matchFilter && matchSearch) {
        card.classList.remove("hidden");
        visibleCount++;
      } else {
        card.classList.add("hidden");
      }
    });

    if (emptyState) {
      if (visibleCount === 0) {
        emptyState.classList.add("visible");
      } else {
        emptyState.classList.remove("visible");
      }
    }
  }

  if (filterBtns.length) {
    filterBtns.forEach(function (btn) {
      btn.addEventListener("click", function () {
        filterBtns.forEach(function (b) { b.classList.remove("active"); });
        btn.classList.add("active");
        activeFilter = btn.getAttribute("data-filter") || "all";
        updateCards();
      });
    });
  }

  if (searchInput) {
    var searchTimer;
    searchInput.addEventListener("input", function () {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(function () {
        searchQuery = searchInput.value.trim();
        updateCards();
      }, 200);
    });
  }

  // === Scroll Reveal ===
  var revealEls = document.querySelectorAll(".reveal");
  if (revealEls.length) {
    var revealObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("visible");
            revealObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.1, rootMargin: "0px 0px -40px 0px" }
    );
    revealEls.forEach(function (el) { revealObserver.observe(el); });
  }

  // === Reading Progress Bar ===
  var progressBar = document.getElementById("progress-bar");
  if (progressBar) {
    window.addEventListener("scroll", function () {
      var scrollTop = window.scrollY || document.documentElement.scrollTop;
      var docHeight = document.documentElement.scrollHeight - window.innerHeight;
      var pct = docHeight > 0 ? (scrollTop / docHeight) * 100 : 0;
      progressBar.style.width = Math.min(100, Math.max(0, pct)) + "%";
    });
  }

  // === Back to Top ===
  var backBtn = document.getElementById("back-to-top");
  if (backBtn) {
    window.addEventListener("scroll", function () {
      if ((window.scrollY || document.documentElement.scrollTop) > 500) {
        backBtn.classList.add("visible");
      } else {
        backBtn.classList.remove("visible");
      }
    });

    backBtn.addEventListener("click", function () {
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  }

})();
