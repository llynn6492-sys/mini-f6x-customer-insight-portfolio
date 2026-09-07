(() => {
  const closestCard = (element) =>
    element.matches?.(".card") ? element : element.closest?.(".card");

  const InsightBanner = ({ text, type = "default" }) => {
    const banner = document.createElement("div");
    banner.className = `insight-banner insight-banner--${type}`;
    banner.setAttribute("role", "note");
    banner.innerHTML = `<span class="insight-banner__accent" aria-hidden="true"></span><p></p>`;
    banner.querySelector("p").textContent = text;
    return banner;
  };

  const mountAll = (root = document) => {
    const config = window.DASHBOARD_INSIGHTS ?? [];
    let mounted = false;
    config.forEach(({ selector, text, type }) => {
      root.querySelectorAll(selector).forEach((target) => {
        const card = closestCard(target);
        if (!card || card.querySelector(":scope > .insight-banner")) return;
        card.classList.add("has-insight-banner");
        card.append(InsightBanner({ text, type }));
        mounted = true;
      });
    });
    if (mounted) {
      requestAnimationFrame(() => {
        window.dispatchEvent(new Event("resize"));
      });
    }
  };

  let scheduled = false;
  const scheduleMount = () => {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      mountAll();
    });
  };

  const observer = new MutationObserver(scheduleMount);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  document.addEventListener("DOMContentLoaded", scheduleMount, { once: true });

  window.InsightBanner = Object.freeze({ create: InsightBanner, mountAll });
})();
