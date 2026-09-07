(() => {
  const WORKBOOK_URL =
    window.DASHBOARD_WORKBOOK_URL ??
    "data/mock_dashboard_data.xlsx?v=20260727-3";
  const CHART_SHEET = "购买动机与购车原因";
  const QUOTE_SHEET = "客户原话摘录";
  const SAMPLE_SHEET = "sample_overview";
  const MODELS = ["F65", "F66"];
  const MODEL_COLORS = { F65: "#0066B3", F66: "#259AC5" };
  const QUOTE_REASONS = [
    "外观设计好",
    "操控性好",
    "动力性能好",
    "销售服务好",
    "内部设计好",
  ];
  const charts = new Map();
  let workbookPromise;
  let quoteRowsPromise;

  const escapeHtml = (value) =>
    String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");

  const formatPercentage = (value) => {
    const number = Number(value);
    return Number.isFinite(number)
      ? `${(number * 100).toFixed(1)}%`
      : "N/A";
  };

  const loadEmbeddedWorkbook = () => {
    const encoded = window.DASHBOARD_WORKBOOK_BASE64;
    if (!encoded) throw new Error("内置数据副本不可用。");
    const binary = window.atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    return window.XLSX.read(bytes, { type: "array" });
  };

  const loadWorkbook = async () => {
    if (!window.XLSX) throw new Error("Excel 读取组件不可用。");
    if (!workbookPromise) {
      workbookPromise =
        window.location.protocol === "file:"
          ? Promise.resolve(loadEmbeddedWorkbook())
          : fetch(WORKBOOK_URL)
              .then((response) => {
                if (!response.ok) {
                  throw new Error(`数据文件加载失败：${response.status}`);
                }
                return response.arrayBuffer();
              })
              .then((buffer) => window.XLSX.read(buffer))
              .catch((error) => {
                if (window.DASHBOARD_WORKBOOK_BASE64) {
                  return loadEmbeddedWorkbook();
                }
                throw error;
              });
    }
    return workbookPromise;
  };

  const loadSheetRows = async (sheetName) => {
    let workbook = await loadWorkbook();
    let worksheet = workbook.Sheets[sheetName];
    if (!worksheet && window.DASHBOARD_WORKBOOK_BASE64) {
      workbook = loadEmbeddedWorkbook();
      workbookPromise = Promise.resolve(workbook);
      worksheet = workbook.Sheets[sheetName];
    }
    if (!worksheet) throw new Error(`未找到工作表：${sheetName}`);
    return window.XLSX.utils.sheet_to_json(worksheet, {
      defval: null,
      raw: true,
    });
  };

  const normalizeChartRows = (rows) =>
    rows
      .map((row, rowIndex) => ({
        metric: row["指标"],
        category: row["分类"],
        model: row["车型"],
        count: Number.isFinite(Number(row.N)) ? Number(row.N) : null,
        percentage: Number.isFinite(Number(row["占比"]))
          ? Number(row["占比"])
          : null,
        sortOrder: Number.isFinite(Number(row["排序"]))
          ? Number(row["排序"])
          : rowIndex,
      }))
      .filter(
        (row) =>
          MODELS.includes(row.model) &&
          typeof row.category === "string",
      );

  const compactTooltip = (detailLabel) => ({
    trigger: "item",
    confine: true,
    transitionDuration: 0,
    backgroundColor: "rgba(255,255,255,.96)",
    borderColor: "rgba(156,181,192,.82)",
    borderWidth: 1,
    padding: [7, 9],
    textStyle: { color: "#203741", fontSize: 10 },
    extraCssText:
      "box-shadow:0 8px 20px rgba(7,31,45,.14);border-radius:4px;backdrop-filter:blur(4px);",
    position: (point) => [point[0] + 10, point[1] + 10],
    formatter: ({ data }) =>
      [
        '<div style="line-height:1.55;white-space:nowrap">',
        `<div>车型：<b>${escapeHtml(data.model)}</b></div>`,
        `<div>${escapeHtml(detailLabel)}：<b>${escapeHtml(data.category)}</b></div>`,
        `<div>N：<b>${escapeHtml(data.count ?? "N/A")}</b></div>`,
        `<div>占比：<b>${escapeHtml(formatPercentage(data.percentage))}</b></div>`,
        "</div>",
      ].join(""),
  });

  const disposePanelCharts = (container) => {
    container
      .querySelectorAll(".section3-model-panel__canvas")
      .forEach((canvas) => window.echarts?.getInstanceByDom(canvas)?.dispose());
  };

  const renderComparisonPanels = (
    containerId,
    rows,
    detailLabel,
    sampleSizes,
  ) => {
    const container = document.getElementById(containerId);
    if (!container) throw new Error(`未找到图表容器：${containerId}`);
    if (!window.echarts) throw new Error("ECharts 图表组件不可用。");

    disposePanelCharts(container);
    container.innerHTML = "";

    const maxPercentage = Math.max(
      0.2,
      ...rows.map((row) => row.percentage ?? 0),
    );
    const axisMax = Math.min(
      1,
      Math.max(0.5, Math.ceil((maxPercentage + 0.08) * 10) / 10),
    );

    const renderedCharts = MODELS.map((model) => {
      const modelRows = rows
        .filter((row) => row.model === model)
        .sort((a, b) => a.sortOrder - b.sortOrder);
      const panel = document.createElement("section");
      panel.className = "section3-model-panel";
      panel.style.setProperty("--model-color", MODEL_COLORS[model]);

      const header = document.createElement("div");
      header.className = "section3-model-panel__header";
      header.innerHTML = `
        <strong>${model}</strong>
        <span class="section3-model-panel__sample">N=${escapeHtml(sampleSizes[model] ?? "N/A")}</span>`;

      const canvas = document.createElement("div");
      canvas.className = "section3-model-panel__canvas";
      panel.append(header, canvas);
      container.append(panel);

      const chart = window.echarts.init(canvas, null, { renderer: "canvas" });
      chart.setOption(
        {
          animation: false,
          backgroundColor: "#FFFFFF",
          aria: { enabled: true, decal: { show: false } },
          tooltip: compactTooltip(detailLabel),
          grid: {
            left: 150,
            right: 52,
            top: 18,
            bottom: 28,
            containLabel: false,
          },
          xAxis: {
            type: "value",
            min: 0,
            max: axisMax,
            axisLabel: {
              color: "#899BA4",
              fontSize: 9,
              formatter: (value) => `${Math.round(value * 100)}%`,
            },
            axisLine: { show: false },
            axisTick: { show: false },
            splitLine: { lineStyle: { color: "#E5ECEF" } },
          },
          yAxis: {
            type: "category",
            inverse: true,
            data: modelRows.map((row) => row.category),
            axisLabel: {
              color: "#4D626D",
              fontSize: 10,
              width: 136,
              overflow: "truncate",
              margin: 10,
            },
            axisLine: { show: false },
            axisTick: { show: false },
          },
          series: [
            {
              name: model,
              type: "bar",
              barMaxWidth: 19,
              showBackground: true,
              backgroundStyle: {
                color: "#F0F4F6",
                borderRadius: 2,
              },
              data: modelRows.map((row) => ({
                value: row.percentage ?? 0,
                ...row,
              })),
              itemStyle: {
                color: MODEL_COLORS[model],
                borderRadius: [0, 2, 2, 0],
              },
              label: {
                show: true,
                position: "right",
                distance: 7,
                color: "#324A56",
                fontSize: 9,
                fontWeight: 700,
                formatter: ({ data }) =>
                  formatPercentage(data.percentage),
              },
              emphasis: { itemStyle: { opacity: 0.82 } },
            },
          ],
        },
        true,
      );
      charts.set(`${containerId}-${model}`, chart);
      return chart;
    });

    const observer = new ResizeObserver(() =>
      renderedCharts.forEach((chart) => chart.resize()),
    );
    container
      .querySelectorAll(".section3-model-panel__canvas")
      .forEach((canvas) => observer.observe(canvas));
    return renderedCharts;
  };

  const createMetricChart = async (metric, containerId, detailLabel) => {
    const container = document.getElementById(containerId);
    if (!container) return [];
    container.innerHTML = `<span class="chart-status">正在加载${escapeHtml(metric)}数据…</span>`;
    try {
      const [chartRows, sampleRows] = await Promise.all([
        loadSheetRows(CHART_SHEET),
        loadSheetRows(SAMPLE_SHEET),
      ]);
      const rows = normalizeChartRows(chartRows).filter(
        (row) => row.metric === metric,
      );
      const sampleSizes = Object.fromEntries(
        MODELS.map((model) => {
          const row = sampleRows.find(
            (item) =>
              item["车型"] === model && item["用户类型"] === "车主",
          );
          const size = Number(row?.["样本量"]);
          return [model, Number.isFinite(size) ? size : "N/A"];
        }),
      );
      return renderComparisonPanels(
        containerId,
        rows,
        detailLabel,
        sampleSizes,
      );
    } catch (error) {
      container.innerHTML = `<span class="chart-status chart-status--error">${escapeHtml(metric)}数据加载失败</span>`;
      throw error;
    }
  };

  const normalizeQuoteRows = (rows) =>
    rows
      .map((row, rowIndex) => ({
        reason: row["原因分类"],
        model: row["车型"],
        quote: row["原话"],
        words: String(row["高亮词"] ?? "")
          .split("|")
          .map((word) => word.trim())
          .filter(Boolean),
        types: String(row["高亮类型"] ?? "")
          .split("|")
          .map((type) => type.trim()),
        sortOrder: Number.isFinite(Number(row["排序"]))
          ? Number(row["排序"])
          : rowIndex,
      }))
      .filter(
        (row) =>
          QUOTE_REASONS.includes(row.reason) &&
          MODELS.includes(row.model) &&
          typeof row.quote === "string",
      );

  const highlightQuote = ({ quote, words, types }) => {
    const highlights = words
      .map((word, index) => ({
        word,
        type: types[index] ?? types[0] ?? "",
      }))
      .filter(({ word }) => quote.includes(word));
    let cursor = 0;
    const fragments = [];

    while (cursor < quote.length) {
      const candidates = highlights
        .map((highlight) => ({
          ...highlight,
          index: quote.indexOf(highlight.word, cursor),
        }))
        .filter(({ index }) => index >= 0)
        .sort(
          (a, b) =>
            a.index - b.index || b.word.length - a.word.length,
        );
      const next = candidates[0];
      if (!next) {
        fragments.push(escapeHtml(quote.slice(cursor)));
        break;
      }
      fragments.push(escapeHtml(quote.slice(cursor, next.index)));
      const className = next.type.includes("实际体验")
        ? "quote-highlight--experience"
        : "quote-highlight--vehicle";
      fragments.push(
        `<mark class="quote-highlight ${className}">${escapeHtml(next.word)}</mark>`,
      );
      cursor = next.index + next.word.length;
    }
    return fragments.join("");
  };

  const renderQuotes = async (reason) => {
    const rows =
      (await quoteRowsPromise) ??
      normalizeQuoteRows(await loadSheetRows(QUOTE_SHEET));
    quoteRowsPromise = Promise.resolve(rows);

    document.querySelectorAll(".quote-tab").forEach((tab) => {
      tab.setAttribute(
        "aria-selected",
        String(tab.dataset.reason === reason),
      );
    });

    MODELS.forEach((model) => {
      const container = document.getElementById(
        `quote-list-${model.toLowerCase()}`,
      );
      if (!container) return;
      const modelRows = rows
        .filter((row) => row.reason === reason && row.model === model)
        .sort((a, b) => a.sortOrder - b.sortOrder);
      container.innerHTML =
        modelRows.length > 0
          ? modelRows
              .map(
                (row) => {
                  const tagTypes = [...new Set(row.types ?? [])];
                  return `
                    <blockquote class="customer-quote">
                      <div class="customer-quote__meta">
                        <span class="customer-quote__model">${escapeHtml(model)}</span>
                        <span>${escapeHtml(reason)}</span>
                      </div>
                      <p>${highlightQuote(row)}</p>
                      <div class="customer-quote__tags">
                        ${tagTypes
                          .map((type) => {
                            const experience = String(type).includes("体验");
                            return `<span class="customer-quote__tag customer-quote__tag--${experience ? "experience" : "vehicle"}">${experience ? "用户体验相关" : "产品/服务相关"}</span>`;
                          })
                          .join("")}
                      </div>
                    </blockquote>`;
                },
              )
              .join("")
          : '<div class="quote-empty">暂无原话摘录</div>';
    });
  };

  const createQuoteSection = async () => {
    const tabs = document.getElementById("quote-reason-tabs");
    if (!tabs) return;
    tabs.innerHTML = QUOTE_REASONS.map(
      (reason, index) =>
        `<button class="quote-tab" type="button" role="tab" data-reason="${escapeHtml(reason)}" aria-selected="${index === 0}">${escapeHtml(reason)}</button>`,
    ).join("");
    tabs.addEventListener("click", (event) => {
      const tab = event.target.closest(".quote-tab");
      if (tab) renderQuotes(tab.dataset.reason);
    });
    await renderQuotes(QUOTE_REASONS[0]);
  };

  const createMotivationChart = () =>
    createMetricChart("购车动机", "motivation-chart", "购车动机");

  const createReasonChart = () =>
    createMetricChart("购车原因", "reason-chart", "购车原因");

  const resizeAll = () => charts.forEach((chart) => chart.resize());
  window.addEventListener("resize", resizeAll);

  window.MotivationCharts = Object.freeze({
    createMotivationChart,
    createQuoteSection,
    createReasonChart,
    renderQuotes,
    resizeAll,
  });
})();
