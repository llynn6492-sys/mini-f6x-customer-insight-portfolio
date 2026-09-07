(() => {
  const MIN_WIDTH = 176;
  const MAX_WIDTH = 280;
  const DEFAULT_WIDTH = 208;
  const STORAGE_KEY = "mini-dashboard-sidebar-width";
  const DESKTOP_QUERY = window.matchMedia("(min-width: 821px)");

  const root = document.documentElement;
  const handle = document.querySelector(".sidebar-resizer");

  if (!handle) {
    return;
  }

  const clamp = (value) => Math.min(MAX_WIDTH, Math.max(MIN_WIDTH, value));

  const applyWidth = (width, persist = true) => {
    const nextWidth = clamp(Math.round(width));
    root.style.setProperty("--sidebar-width", `${nextWidth}px`);
    handle.setAttribute("aria-valuenow", String(nextWidth));

    if (persist) {
      localStorage.setItem(STORAGE_KEY, String(nextWidth));
    }
  };

  const storedWidth = Number(localStorage.getItem(STORAGE_KEY));
  if (Number.isFinite(storedWidth) && DESKTOP_QUERY.matches) {
    applyWidth(storedWidth, false);
  }

  const finishResize = (event) => {
    document.body.classList.remove("is-resizing-sidebar");
    handle.releasePointerCapture?.(event.pointerId);
    window.removeEventListener("pointermove", resizeFromPointer);
    window.removeEventListener("pointerup", finishResize);
  };

  const resizeFromPointer = (event) => {
    applyWidth(event.clientX);
  };

  handle.addEventListener("pointerdown", (event) => {
    if (!DESKTOP_QUERY.matches) {
      return;
    }

    event.preventDefault();
    handle.setPointerCapture?.(event.pointerId);
    document.body.classList.add("is-resizing-sidebar");
    window.addEventListener("pointermove", resizeFromPointer);
    window.addEventListener("pointerup", finishResize);
  });

  handle.addEventListener("dblclick", () => {
    applyWidth(DEFAULT_WIDTH);
  });

  handle.addEventListener("keydown", (event) => {
    const currentWidth =
      Number.parseInt(
        getComputedStyle(root).getPropertyValue("--sidebar-width"),
        10,
      ) || DEFAULT_WIDTH;

    const keyActions = {
      ArrowLeft: () => applyWidth(currentWidth - 8),
      ArrowRight: () => applyWidth(currentWidth + 8),
      Home: () => applyWidth(MIN_WIDTH),
      End: () => applyWidth(MAX_WIDTH),
    };

    if (keyActions[event.key]) {
      event.preventDefault();
      keyActions[event.key]();
    }
  });
})();

(() => {
  const links = [...document.querySelectorAll("[data-page-target]")];
  const pages = [...document.querySelectorAll("[data-page]")];
  const defaultPage = document.body.dataset.defaultPage || "home";
  const validPages = new Set(pages.map((page) => page.dataset.page));

  if (links.length === 0 || pages.length === 0) {
    return;
  }

  const activatePage = (requestedPage, updateHistory = false) => {
    const normalizedPage =
      requestedPage === "conversion" ? "purchase-journey" : requestedPage;
    const pageName = validPages.has(normalizedPage)
      ? normalizedPage
      : defaultPage;

    pages.forEach((page) => {
      const isActive = page.dataset.page === pageName;
      page.hidden = !isActive;
      page.classList.toggle("dashboard-page--active", isActive);
    });

    links.forEach((link) => {
      const isActive = link.dataset.pageTarget === pageName;
      link.classList.toggle("sidebar__link--active", isActive);
      if (isActive) {
        link.setAttribute("aria-current", "page");
      } else {
        link.removeAttribute("aria-current");
      }
    });

    if (updateHistory) {
      window.history.pushState(
        { page: pageName },
        "",
        `#${pageName}`,
      );
    }

    window.scrollTo({ top: 0, behavior: "auto" });
    window.dispatchEvent(
      new CustomEvent("dashboard:pagechange", {
        detail: { page: pageName },
      }),
    );
  };

  links.forEach((link) => {
    link.addEventListener("click", (event) => {
      event.preventDefault();
      activatePage(link.dataset.pageTarget, true);
    });
  });

  window.addEventListener("popstate", () => {
    activatePage(window.location.hash.slice(1), false);
  });

  activatePage(window.location.hash.slice(1) || defaultPage, false);
})();

(() => {
  let initialized = false;
  let initializing = false;

  const initializePurchaseSection = async () => {
    if (
      initialized ||
      initializing ||
      !window.PurchaseCharts
    ) {
      if (initialized) {
        requestAnimationFrame(() => window.PurchaseCharts?.resizeAll());
        window.dispatchEvent(
          new CustomEvent("dashboard:sectionready", {
            detail: { page: "purchase-behavior" },
          }),
        );
      }
      return;
    }

    initializing = true;
    try {
      await Promise.all([
        window.PurchaseCharts.createPurchaseTypePie(),
        window.PurchaseCharts.createFemaleUserChart(),
        window.PurchaseCharts.createPurchaseProfileDrilldown(
          document.getElementById("purchase-type-filter")?.value ?? "首购",
          document.getElementById("purchase-model-filter")?.value ?? "F65",
        ),
      ]);
      initialized = true;
      window.dispatchEvent(
        new CustomEvent("dashboard:sectionready", {
          detail: { page: "purchase-behavior" },
        }),
      );
    } catch (error) {
      console.error("购车行为分析初始化失败：", error);
    } finally {
      initializing = false;
    }
  };

  const purchaseTypeFilter = document.getElementById("purchase-type-filter");
  const purchaseModelFilter = document.getElementById("purchase-model-filter");
  const refreshPurchaseProfile = async () => {
    try {
      await window.PurchaseCharts?.createPurchaseProfileDrilldown(
        purchaseTypeFilter?.value ?? "首购",
        purchaseModelFilter?.value ?? "F65",
      );
    } catch (error) {
      console.error("购车类型画像切换失败：", error);
    }
  };
  purchaseTypeFilter?.addEventListener("change", refreshPurchaseProfile);
  purchaseModelFilter?.addEventListener("change", refreshPurchaseProfile);

  window.addEventListener("dashboard:pagechange", (event) => {
    if (event.detail?.page === "purchase-behavior") {
      initializePurchaseSection();
    }
  });

  const purchasePage = document.querySelector(
    '[data-page="purchase-behavior"]',
  );
  if (purchasePage && !purchasePage.hidden) {
    initializePurchaseSection();
  }
})();

(() => {
  let initialized = false;
  let initializing = false;

  const initializeOverview = async () => {
    if (initialized || initializing || !window.OverviewDashboard) {
      if (initialized) {
        requestAnimationFrame(() => window.OverviewDashboard?.resizeAll());
      }
      return;
    }
    initializing = true;
    try {
      await window.OverviewDashboard.initialize();
      initialized = true;
    } catch (error) {
      console.error("首页初始化失败：", error);
    } finally {
      initializing = false;
    }
  };

  window.addEventListener("dashboard:pagechange", (event) => {
    if (event.detail?.page === "home") initializeOverview();
  });

  const homePage = document.querySelector('[data-page="home"]');
  if (homePage && !homePage.hidden) initializeOverview();
})();

(() => {
  let initialized = false;
  let initializing = false;

  const initializeMotivationSection = async () => {
    if (initialized || initializing || !window.MotivationCharts) {
      if (initialized) {
        requestAnimationFrame(() => window.MotivationCharts?.resizeAll());
        window.dispatchEvent(
          new CustomEvent("dashboard:sectionready", {
            detail: { page: "motivation-product" },
          }),
        );
      }
      return;
    }

    initializing = true;
    try {
      await Promise.all([
        window.MotivationCharts.createMotivationChart(),
        window.MotivationCharts.createReasonChart(),
        window.MotivationCharts.createQuoteSection(),
      ]);
      initialized = true;
      window.dispatchEvent(
        new CustomEvent("dashboard:sectionready", {
          detail: { page: "motivation-product" },
        }),
      );
    } catch (error) {
      console.error("购买动机与产品认知初始化失败：", error);
    } finally {
      initializing = false;
    }
  };

  window.addEventListener("dashboard:pagechange", (event) => {
    if (event.detail?.page === "motivation-product") {
      initializeMotivationSection();
    }
  });

  const motivationPage = document.querySelector(
    '[data-page="motivation-product"]',
  );
  if (motivationPage && !motivationPage.hidden) {
    initializeMotivationSection();
  }
})();

(() => {
  let initialized = false;
  let initializing = false;

  const initializePurchaseJourney = async () => {
    if (initialized || initializing || !window.PurchaseJourney) {
      if (initialized) {
        requestAnimationFrame(() => window.PurchaseJourney?.resizeAll());
        window.dispatchEvent(
          new CustomEvent("dashboard:sectionready", {
            detail: { page: "purchase-journey" },
          }),
        );
      }
      return;
    }

    initializing = true;
    try {
      await window.PurchaseJourney.PurchaseJourneySection();
      initialized = true;
      window.dispatchEvent(
        new CustomEvent("dashboard:sectionready", {
          detail: { page: "purchase-journey" },
        }),
      );
    } catch (error) {
      console.error("购买决策路径分析初始化失败：", error);
    } finally {
      initializing = false;
    }
  };

  window.addEventListener("dashboard:pagechange", (event) => {
    if (event.detail?.page === "purchase-journey") {
      initializePurchaseJourney();
    }
  });

  const journeyPage = document.querySelector(
    '[data-page="purchase-journey"]',
  );
  if (journeyPage && !journeyPage.hidden) {
    initializePurchaseJourney();
  }
})();

(() => {
  let initialized = false;

  const initializeLeadsIntent = () => {
    if (!window.LeadsIntent) return;
    if (!initialized) {
      try {
        window.LeadsIntent.initialize();
        initialized = true;
      } catch (error) {
        console.error("潜客购车意向分析初始化失败：", error);
      }
      return;
    }
    requestAnimationFrame(() => window.LeadsIntent.resizeAll());
  };

  window.addEventListener("dashboard:pagechange", (event) => {
    if (event.detail?.page === "leads-intent") {
      initializeLeadsIntent();
    }
  });

  const page = document.querySelector('[data-page="leads-intent"]');
  if (page && !page.hidden) initializeLeadsIntent();
})();

(() => {
  let initialized = false;
  let initializing = false;

  const initializeLeadsLoss = async () => {
    if (!window.LeadsLoss || initializing) return;
    if (!initialized) {
      initializing = true;
      try {
        await window.LeadsLoss.initialize();
        initialized = true;
      } catch (error) {
        console.error("潜客流失原因分析初始化失败：", error);
      } finally {
        initializing = false;
      }
      return;
    }
    requestAnimationFrame(() => window.LeadsLoss.resizeAll());
  };

  window.addEventListener("dashboard:pagechange", (event) => {
    if (event.detail?.page === "leads-loss") {
      initializeLeadsLoss();
    }
  });

  const page = document.querySelector('[data-page="leads-loss"]');
  if (page && !page.hidden) initializeLeadsLoss();
})();

(() => {
  let activeAudience = "owner";
  let renderVersion = 0;

  const CustomerProfileSection = async ({
    audience = "owner",
  } = {}) => {
    if (!window.DashboardCharts) return;
    const currentVersion = ++renderVersion;
    activeAudience = audience;

    document.querySelectorAll("[data-audience]").forEach((button) => {
      const isActive = button.dataset.audience === audience;
      button.classList.toggle("audience-switcher__button--active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
      button.disabled = true;
    });

    const sourceLabel = document.getElementById(
      "profile-data-source-label",
    );
    if (sourceLabel) {
      sourceLabel.textContent =
        audience === "prospect" ? "潜客样本" : "车主样本";
    }
    window.DashboardUI?.hydrateInsightBoxes({ audience });

    const options = { audience };
    const results = await Promise.allSettled([
      window.DashboardCharts.hydrateCustomerSnapshot(options),
      window.DashboardCharts.hydrateChartSampleSizes(options),
      window.DashboardCharts.createGenderChart(options),
      window.DashboardCharts.createAgeChart(options),
      window.DashboardCharts.createMaritalStatusChart(options),
      window.DashboardCharts.createEducationChart(options),
      window.DashboardCharts.createIncomeChart(options),
      window.DashboardCharts.createIndustryChart(options),
    ]);

    if (currentVersion === renderVersion) {
      document.querySelectorAll("[data-audience]").forEach((button) => {
        button.disabled = false;
      });
    }

    results
      .filter((result) => result.status === "rejected")
      .forEach((result) =>
        console.error("客户基础画像刷新失败：", result.reason),
      );
  };

  window.CustomerProfileSection = CustomerProfileSection;

  document.querySelectorAll("[data-audience]").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.audience !== activeAudience) {
        CustomerProfileSection({
          audience: button.dataset.audience,
        });
      }
    });
  });

  if (document.body.dataset.dashboardMode !== "quarterly") {
    CustomerProfileSection({ audience: activeAudience });
  }
})();
