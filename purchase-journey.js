(() => {
  const WORKBOOK_URL =
    window.DASHBOARD_WORKBOOK_URL ??
    "data/mock_dashboard_data.xlsx?v=20260728-1";
  const SHEET_NAME = "Section4_Purchase_Journey";
  const MODELS = ["F65", "F66"];
  const MODEL_COLORS = { F65: "#0066B3", F66: "#259AC5" };
  const charts = new Map();
  let workbookPromise;
  let rowsPromise;

  const escapeHtml = (value) =>
    String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");

  const formatPercentage = (value) =>
    Number.isFinite(Number(value))
      ? `${(Number(value) * 100).toFixed(1)}%`
      : "N/A";

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

  const loadRows = async () => {
    if (!rowsPromise) {
      rowsPromise = loadWorkbook().then((workbook) => {
        let worksheet = workbook.Sheets[SHEET_NAME];
        if (!worksheet && window.DASHBOARD_WORKBOOK_BASE64) {
          workbook = loadEmbeddedWorkbook();
          workbookPromise = Promise.resolve(workbook);
          worksheet = workbook.Sheets[SHEET_NAME];
        }
        if (!worksheet) throw new Error(`未找到工作表：${SHEET_NAME}`);
        return window.XLSX.utils
          .sheet_to_json(worksheet, { defval: null, raw: true })
          .map((row, index) => ({
            module: row.Module,
            metric: row.Metric,
            model: row.Model,
            category: row.Category,
            value: Number(row.Value),
            sampleSize: Number(row.N),
            extra: String(row.Extra ?? ""),
            sourceOrder: index,
          }))
          .filter(
            (row) =>
              MODELS.includes(row.model) &&
              typeof row.module === "string" &&
              typeof row.category === "string" &&
              Number.isFinite(row.value),
          );
      });
    }
    return rowsPromise;
  };

  const compactTooltip = (formatter) => ({
    trigger: "item",
    confine: true,
    transitionDuration: 0,
    backgroundColor: "rgba(255,255,255,.97)",
    borderColor: "rgba(156,181,192,.82)",
    borderWidth: 1,
    padding: [7, 9],
    textStyle: { color: "#203741", fontSize: 10 },
    extraCssText:
      "box-shadow:0 8px 20px rgba(7,31,45,.14);border-radius:4px;",
    position: (point) => [point[0] + 10, point[1] + 10],
    formatter,
  });

  const disposeChartGroup = (prefix) => {
    [...charts.entries()]
      .filter(([key]) => key.startsWith(prefix))
      .forEach(([key, chart]) => {
        chart.dispose();
        charts.delete(key);
      });
  };

  const cardHeader = (model, sampleSize, subtitle) => `
    <header class="journey-card__header">
      <div>
        <span class="journey-card__model-dot" style="--model-color:${MODEL_COLORS[model]}"></span>
        <strong>${model}</strong>
        ${subtitle ? `<small>${escapeHtml(subtitle)}</small>` : ""}
      </div>
      <span class="journey-card__sample">N=${Number.isFinite(sampleSize) ? sampleSize : "N/A"}</span>
    </header>`;

  const createBarComparison = async ({
    module,
    metric,
    containerId,
    detailLabel,
    chartKey,
    highlightTopThree = false,
  }) => {
    const container = document.getElementById(containerId);
    if (!container) return [];
    try {
      const allRows = await loadRows();
      disposeChartGroup(chartKey);
      container.innerHTML = "";
      const rendered = MODELS.map((model) => {
        const modelRows = allRows
          .filter(
            (row) =>
              row.module === module &&
              row.metric === metric &&
              row.model === model,
          )
          .sort((a, b) => b.value - a.value);
        const sampleSize = modelRows[0]?.sampleSize;
        const card = document.createElement("article");
        card.className = "card journey-model-card journey-model-card--bar";
        card.dataset.insightKey = `${chartKey}-${model.toLowerCase()}`;
        card.innerHTML = `
          ${cardHeader(model, sampleSize, module)}
          <div class="journey-bar-chart"></div>`;
        container.append(card);
        const canvas = card.querySelector(".journey-bar-chart");
        const axisMax = Math.max(
          0.25,
          Math.ceil((Math.max(...modelRows.map((row) => row.value), 0.1) + 0.05) * 20) /
            20,
        );
        const chart = window.echarts.init(canvas, null, {
          renderer: "canvas",
        });
        chart.setOption(
          {
            animation: true,
            animationDuration: 420,
            animationEasing: "cubicOut",
            backgroundColor: "#FFFFFF",
            tooltip: compactTooltip(({ data }) =>
              [
                '<div style="line-height:1.55;white-space:nowrap">',
                `<div>车型：<b>${data.model}</b></div>`,
                `<div>${escapeHtml(detailLabel)}：<b>${escapeHtml(data.category)}</b></div>`,
                `<div>${module === "竞品考虑" ? "提及率" : "占比"}：<b>${formatPercentage(data.value)}</b></div>`,
                module === "竞品考虑"
                  ? `<div>考虑车型：<b>${escapeHtml(data.extra || "N/A")}</b></div>`
                  : `<div>Sample size：<b>${data.sampleSize ?? "N/A"}</b></div>`,
                "</div>",
              ].join(""),
            ),
            grid: {
              left: 108,
              right: 48,
              top: 22,
              bottom: 25,
              containLabel: false,
            },
            xAxis: {
              type: "value",
              min: 0,
              max: axisMax,
              axisLabel: {
                color: "#8A9CA5",
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
                color: "#405660",
                fontSize: 10,
                width: 94,
                overflow: "truncate",
                margin: 10,
              },
              axisLine: { show: false },
              axisTick: { show: false },
            },
            series: [
              {
                type: "bar",
                barMaxWidth: 19,
                showBackground: true,
                backgroundStyle: { color: "#F0F4F6", borderRadius: 2 },
                data: modelRows.map((row, index) => ({
                  ...row,
                  itemStyle: {
                    color:
                      highlightTopThree && index >= 3
                        ? "#A7D2E1"
                        : MODEL_COLORS[model],
                    borderRadius: [0, 2, 2, 0],
                  },
                })),
                label: {
                  show: true,
                  position: "right",
                  distance: 6,
                  color: "#324A56",
                  fontSize: 9,
                  fontWeight: 700,
                  formatter: ({ data }) => formatPercentage(data.value),
                },
                emphasis: { itemStyle: { opacity: 0.78 } },
              },
            ],
          },
          true,
        );
        charts.set(`${chartKey}-${model}`, chart);
        return chart;
      });
      return rendered;
    } catch (error) {
      container.innerHTML =
        '<span class="chart-status chart-status--error">数据加载失败</span>';
      throw error;
    }
  };

  const CompetitorChart = () =>
    createBarComparison({
      module: "竞品考虑",
      metric: "品牌",
      containerId: "journey-competitor-charts",
      detailLabel: "竞品",
      chartKey: "journey-competitor",
      highlightTopThree: true,
    });

  const ChannelChart = () =>
    createBarComparison({
      module: "信息渠道",
      metric: "渠道",
      containerId: "journey-channel-charts",
      detailLabel: "渠道",
      chartKey: "journey-channel",
    });

  const OwnerOriginPie = async () => {
    const container = document.getElementById("journey-origin-charts");
    if (!container) return [];
    try {
      const allRows = await loadRows();
      disposeChartGroup("journey-origin");
      container.innerHTML = "";
      const colors = ["#0066B3", "#42A6C4", "#D4E3E8"];
      return MODELS.map((model) => {
        const modelRows = allRows.filter(
          (row) => row.module === "销售构成" && row.model === model,
        );
        const sampleSize = modelRows[0]?.sampleSize;
        const bmw = modelRows.find((row) => row.category === "BMW车主");
        const mini = modelRows.find((row) => row.category === "MINI车主");
        const card = document.createElement("article");
        card.className = "card journey-model-card journey-model-card--pie";
        card.dataset.insightKey = `journey-origin-${model.toLowerCase()}`;
        card.innerHTML = `
          ${cardHeader(model, sampleSize, "来源品牌")}
          <div class="journey-origin-chart"></div>
          <div class="journey-origin-metrics">
            <span>BMW 车主 <b>${formatPercentage(bmw?.value)}</b></span>
            <span>MINI 车主 <b>${formatPercentage(mini?.value)}</b></span>
          </div>`;
        container.append(card);
        const canvas = card.querySelector(".journey-origin-chart");
        const chart = window.echarts.init(canvas, null, {
          renderer: "canvas",
        });
        chart.setOption(
          {
            animation: true,
            animationDuration: 450,
            color: colors,
            tooltip: compactTooltip(({ data }) =>
              [
                '<div style="line-height:1.55;white-space:nowrap">',
                `<div>车型：<b>${data.model}</b></div>`,
                `<div>来源品牌：<b>${escapeHtml(data.category)}</b></div>`,
                `<div>人数：<b>${Math.round(data.value * data.sampleSize)}</b></div>`,
                `<div>占比：<b>${formatPercentage(data.value)}</b></div>`,
                "</div>",
              ].join(""),
            ),
            graphic: [
              {
                type: "text",
                left: "center",
                top: "42%",
                style: {
                  text: `N=${sampleSize ?? "N/A"}`,
                  fill: "#17313D",
                  font: '700 13px Inter, "PingFang SC", sans-serif',
                  textAlign: "center",
                },
              },
            ],
            series: [
              {
                type: "pie",
                center: ["50%", "46%"],
                radius: ["39%", "65%"],
                avoidLabelOverlap: true,
                itemStyle: { borderColor: "#FFFFFF", borderWidth: 2 },
                label: {
                  color: "#405660",
                  fontSize: 10,
                  formatter: ({ data }) =>
                    `${data.category}\n${formatPercentage(data.value)}`,
                },
                labelLine: {
                  length: 10,
                  length2: 7,
                  lineStyle: { color: "#A5B8C1" },
                },
                emphasis: { scaleSize: 4 },
                data: modelRows.map((row) => ({ ...row, name: row.category })),
              },
            ],
          },
          true,
        );
        charts.set(`journey-origin-${model}`, chart);
        return chart;
      });
    } catch (error) {
      container.innerHTML =
        '<span class="chart-status chart-status--error">车主来源数据加载失败</span>';
      throw error;
    }
  };

  const parseRank = (extra, fallback) => {
    const value = Number(extra.match(/Rank(\d+)/i)?.[1]);
    return Number.isFinite(value) ? value : fallback;
  };

  const CityRankingTable = async () => {
    const container = document.getElementById("journey-city-rankings");
    if (!container) return;
    try {
      const allRows = await loadRows();
      container.innerHTML = MODELS.map((model) => {
        const modelRows = allRows
          .filter((row) => row.module === "城市分布" && row.model === model)
          .map((row, index) => ({
            ...row,
            rank: parseRank(row.extra, index + 1),
          }))
          .sort((a, b) => a.rank - b.rank)
          .slice(0, 10);
        const maxValue = Math.max(...modelRows.map((row) => row.value), 0.01);
        const totalContribution = modelRows.reduce(
          (total, row) => total + row.value,
          0,
        );
        return `
          <article class="card journey-ranking-card" data-insight-key="journey-city-${model.toLowerCase()}">
            ${cardHeader(model, modelRows[0]?.sampleSize, "Top 10 Cities")}
            <div class="journey-ranking-list">
              ${modelRows
                .map(
                  (row) => `
                    <div class="journey-ranking-row">
                      <span class="journey-rank-badge">${String(row.rank).padStart(2, "0")}</span>
                      <div class="journey-ranking-row__main">
                        <div><strong>${escapeHtml(row.category)}</strong><b>${formatPercentage(row.value)}</b></div>
                        <span class="journey-progress"><i style="--progress:${(row.value / maxValue) * 100}%;--model-color:${MODEL_COLORS[model]}"></i></span>
                      </div>
                    </div>`,
                )
                .join("")}
            </div>
            <div class="journey-ranking-total">
              <span>Top 10 城市贡献</span>
              <strong>${formatPercentage(totalContribution)}</strong>
            </div>
          </article>`;
      }).join("");
    } catch (error) {
      container.innerHTML =
        '<span class="chart-status chart-status--error">城市分布数据加载失败</span>';
      throw error;
    }
  };

  const parseAverageAge = (extra) => {
    const value = Number(extra.match(/Avg Age\s*([\d.]+)/i)?.[1]);
    return Number.isFinite(value) ? value : null;
  };

  const PreviousVehicleTable = async () => {
    const container = document.getElementById("journey-previous-vehicles");
    if (!container) return;
    try {
      const allRows = await loadRows();
      container.innerHTML = MODELS.map((model) => {
        const prepareRows = (module, limit) =>
          allRows
            .filter((row) => row.module === module && row.model === model)
            .map((row, index) => ({
              ...row,
              rank: parseRank(row.extra, index + 1),
              averageAge: parseAverageAge(row.extra),
            }))
            .sort((a, b) => a.rank - b.rank)
            .slice(0, limit);
        const bmwRows = prepareRows("前车车系-BMW", 10);
        const miniRows = prepareRows("前车车系-MINI", 4);
        const renderRows = (rows) =>
          rows
            .map((row) => `
              <div class="previous-profile-row${row.rank <= 3 ? " previous-profile-row--top" : ""}" tabindex="0">
                <span class="journey-rank-badge">${String(row.rank).padStart(2, "0")}</span>
                <strong>${escapeHtml(row.category)}</strong>
                <b>${formatPercentage(row.value)}</b>
                <span>${row.averageAge == null ? "N/A" : `${row.averageAge.toFixed(1)}年`}</span>
                <div class="journey-row-tooltip" role="tooltip">
                  <span>Vehicle：<b>${escapeHtml(row.category)}</b></span>
                  <span>Contribution：<b>${formatPercentage(row.value)}</b></span>
                  <span>Average Vehicle Age：<b>${row.averageAge == null ? "N/A" : `${row.averageAge.toFixed(1)} years`}</b></span>
                </div>
              </div>`)
            .join("");
        const sampleSize = bmwRows[0]?.sampleSize ?? miniRows[0]?.sampleSize;
        return `
          <article class="card previous-profile-card" data-insight-key="previous-profile-${model.toLowerCase()}">
            ${cardHeader(model, sampleSize, "前车车型画像")}
            <div class="previous-profile-card__body">
              <section class="previous-profile-table">
                <header>
                  <h3>BMW 前车车型前 10 名</h3>
                  <span>BMW 迁移来源</span>
                </header>
                <div class="previous-profile-head" aria-hidden="true">
                  <span>排名</span><span>BMW车型</span><span>占比</span><span>平均车龄</span>
                </div>
                <div class="previous-profile-list">${renderRows(bmwRows)}</div>
              </section>
              <section class="previous-profile-table previous-profile-table--mini">
                <header>
                  <h3>MINI 前车车型前 4 名</h3>
                  <span>MINI 存量来源</span>
                </header>
                <div class="previous-profile-head" aria-hidden="true">
                  <span>排名</span><span>MINI车型</span><span>占比</span><span>平均车龄</span>
                </div>
                <div class="previous-profile-list">${renderRows(miniRows)}</div>
              </section>
            </div>
          </article>`;
      }).join("");
    } catch (error) {
      container.innerHTML =
        '<span class="chart-status chart-status--error">前车车系数据加载失败</span>';
      throw error;
    }
  };

  const PurchaseJourneySection = async () =>
    Promise.all([
      CompetitorChart(),
      ChannelChart(),
      OwnerOriginPie(),
      CityRankingTable(),
      PreviousVehicleTable(),
    ]);

  const resizeAll = () => charts.forEach((chart) => chart.resize());
  window.addEventListener("resize", resizeAll);

  window.PurchaseJourney = Object.freeze({
    ChannelChart,
    CityRankingTable,
    CompetitorChart,
    OwnerOriginPie,
    PreviousVehicleTable,
    PurchaseJourneySection,
    resizeAll,
  });
})();
