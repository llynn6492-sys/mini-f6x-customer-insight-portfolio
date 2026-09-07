(() => {
  const createElement = (tag, className) => {
    const element = document.createElement(tag);
    if (className) element.className = className;
    return element;
  };

  const InsightBox = ({ title, items = [], marker = "" }) => {
    const article = createElement("article", "insight-card");
    const markerElement = createElement("div", "insight-card__marker");
    markerElement.setAttribute("aria-hidden", "true");
    markerElement.textContent = marker;
    const content = createElement("div", "insight-card__content");
    const label = createElement("p", "insight-card__label");
    label.textContent = title;
    const list = createElement("ul", "insight-list");
    items.forEach((item) => {
      const row = document.createElement("li");
      row.textContent = item;
      list.append(row);
    });
    content.append(label, list);
    article.append(markerElement, content);
    return article;
  };

  const ChartCard = ({ title, sampleText = "", className = "" }) => {
    const card = createElement("article", `card chart-card ${className}`.trim());
    const header = createElement("header", "chart-card__header");
    const heading = document.createElement("h3");
    heading.textContent = title;
    header.append(heading);
    if (sampleText) {
      const sample = createElement("span", "chart-card__sample-caption");
      sample.textContent = sampleText;
      header.append(sample);
    }
    const canvas = createElement("div", "chart-canvas");
    card.append(header, canvas);
    return { card, canvas };
  };

  const SnapshotCard = ({ title = "客户概览" } = {}) => {
    const card = createElement("article", "card snapshot-card");
    card.dataset.component = "SnapshotCard";
    card.setAttribute("aria-label", title);
    return card;
  };

  const TableCard = ({ title, sampleText = "" }) => {
    const card = createElement("article", "card table-card");
    const header = createElement("header", "table-card__header");
    const heading = document.createElement("h3");
    heading.textContent = title;
    const sample = createElement("span", "table-card__sample");
    sample.textContent = sampleText;
    header.append(heading, sample);
    const body = createElement("div", "table-card__body");
    card.append(header, body);
    return { card, body };
  };

  const SectionLayout = ({ title, question, number }) => {
    const section = createElement("section", "dashboard-section");
    const heading = createElement("div", "section-heading");
    const left = document.createElement("div");
    const numberElement = createElement("span", "section-heading__number");
    numberElement.textContent = number;
    const titleElement = document.createElement("h2");
    titleElement.textContent = title;
    const questionElement = document.createElement("p");
    questionElement.textContent = question;
    left.append(numberElement, titleElement);
    heading.append(left, questionElement);
    section.append(heading);
    return section;
  };

  const SidebarNavigation = ({ items = [] } = {}) => {
    const nav = createElement("nav", "sidebar__nav");
    nav.setAttribute("aria-label", "主要导航");
    items.forEach(({ page, label }) => {
      const link = createElement("a", "sidebar__link");
      link.href = `#${page}`;
      link.dataset.pageTarget = page;
      link.textContent = label;
      nav.append(link);
    });
    return nav;
  };

  const hydrateInsightBoxes = ({ audience = "owner" } = {}) => {
    const config = window.DASHBOARD_SECTION_CONFIG ?? {};
    document.querySelectorAll("[data-insight-section]").forEach((box) => {
      const section = config[box.dataset.insightSection];
      if (!section) return;
      const audienceContent = section.insightsByAudience?.[audience];
      const title = audienceContent?.title ?? section.insightTitle;
      const insights = audienceContent?.insights ?? section.insights ?? [];
      const label = box.querySelector(".insight-card__label");
      const list = box.querySelector(".insight-list");
      if (label) label.textContent = title;
      if (list) {
        list.innerHTML = "";
        insights.forEach((item) => {
          const row = document.createElement("li");
          row.textContent = item;
          list.append(row);
        });
      }
    });
  };

  hydrateInsightBoxes();

  window.DashboardUI = Object.freeze({
    ChartCard,
    InsightBox,
    SectionLayout,
    SidebarNavigation,
    SnapshotCard,
    TableCard,
    hydrateInsightBoxes,
  });
})();
