(() => {
  const MODELS = ["F65", "F66"];
  const MODEL_COLORS = { F65: "#0066B3", F66: "#259AC5" };
  const PIE_COLORS = ["#62B5CD", "#2499BC", "#087AA5"];
  const charts = new Map();

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

  const getData = () => {
    if (!window.LEADS_INTENT_DATA) {
      throw new Error("潜客购车意向模拟数据不可用。");
    }
    return window.LEADS_INTENT_DATA;
  };

  const compactTooltip = (detailLabel) => ({
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
    formatter: ({ data }) =>
      [
        '<div style="line-height:1.55;white-space:nowrap">',
        `<div>车型：<b>${escapeHtml(data.model)} Leads</b></div>`,
        `<div>${escapeHtml(detailLabel)}：<b>${escapeHtml(data.category)}</b></div>`,
        `<div>数量：<b>${data.count ?? "N/A"}</b></div>`,
        `<div>占比：<b>${formatPercentage(data.percentage)}</b></div>`,
        "</div>",
      ].join(""),
  });

  const disposeGroup = (prefix) => {
    [...charts.entries()]
      .filter(([key]) => key.startsWith(prefix))
      .forEach(([key, chart]) => {
        chart.dispose();
        charts.delete(key);
      });
  };

  const modelHeader = (model, sampleSize, subtitle) => `
    <header class="leads-card__header">
      <div>
        <i style="--model-color:${MODEL_COLORS[model]}"></i>
        <strong>${model} Leads</strong>
        ${subtitle ? `<small>${escapeHtml(subtitle)}</small>` : ""}
      </div>
      <span>N=${sampleSize}</span>
    </header>`;

  const createPurchaseTypeChart = () => {
    const container = document.getElementById(
      "leads-purchase-type-charts",
    );
    if (!container) return [];
    const data = getData();
    disposeGroup("leads-purchase-type");
    container.innerHTML = "";

    return MODELS.map((model) => {
      const card = document.createElement("article");
      card.className = "card leads-chart-card leads-chart-card--pie";
      card.innerHTML = `
        ${modelHeader(model, data.sampleSizes[model], "潜客购车类型")}
        <div class="leads-chart-canvas"></div>`;
      container.append(card);
      const canvas = card.querySelector(".leads-chart-canvas");
      const chart = window.echarts.init(canvas, null, {
        renderer: "canvas",
      });
      chart.setOption({
        animation: true,
        animationDuration: 360,
        backgroundColor: "#FFFFFF",
        color: PIE_COLORS,
        tooltip: compactTooltip("购车类型"),
        legend: {
          bottom: 10,
          left: "center",
          itemWidth: 9,
          itemHeight: 9,
          itemGap: 22,
          textStyle: { color: "#607682", fontSize: 10 },
        },
        series: [
          {
            type: "pie",
            center: ["50%", "47%"],
            radius: ["34%", "64%"],
            minShowLabelAngle: 4,
            avoidLabelOverlap: true,
            itemStyle: { borderColor: "#FFFFFF", borderWidth: 2 },
            label: {
              color: "#405660",
              fontSize: 10,
              lineHeight: 14,
              formatter: ({ data: row }) =>
                `${row.category}\n${formatPercentage(row.percentage)}`,
            },
            labelLine: {
              length: 12,
              length2: 8,
              lineStyle: { color: "#A6B9C1" },
            },
            emphasis: { scaleSize: 4 },
            data: data.purchaseType[model].map((row) => ({
              ...row,
              model,
              name: row.category,
              value: row.count,
            })),
          },
        ],
      });
      charts.set(`leads-purchase-type-${model}`, chart);
      return chart;
    });
  };

  const renderVehicleTables = ({
    containerId,
    dataKey,
    subtitle,
  }) => {
    const container = document.getElementById(containerId);
    if (!container) return;
    const data = getData();
    container.innerHTML = MODELS.map((model) => `
      <article class="card leads-table-card">
        ${modelHeader(model, data.sampleSizes[model], subtitle)}
        <div class="leads-table">
          <div class="leads-table__head">
            <span>排名</span><span>品牌</span><span>提及率</span><span>主要车系</span>
          </div>
          ${data[dataKey][model]
            .map(
              (row) => `
                <div class="leads-table__row" tabindex="0">
                  <span class="leads-rank">${String(row.rank).padStart(2, "0")}</span>
                  <strong>${escapeHtml(row.brand)}</strong>
                  <b>${formatPercentage(row.mentionRate)}</b>
                  <small>${escapeHtml(row.vehicleSeries)}</small>
                  <span class="leads-row-tooltip">
                    品牌：${escapeHtml(row.brand)}<br>
                    提及率：${formatPercentage(row.mentionRate)}<br>
                    主要车系：${escapeHtml(row.vehicleSeries)}
                  </span>
                </div>`,
            )
            .join("")}
        </div>
      </article>`).join("");
  };

  const createBarOption = ({
    model,
    rows,
    detailLabel,
  }) => {
    const maxValue = Math.max(...rows.map((row) => row.percentage), 0.1);
    const axisMax =
      Math.ceil(Math.min(0.5, maxValue + 0.06) * 20) / 20;
    return {
      animation: true,
      animationDuration: 360,
      backgroundColor: "#FFFFFF",
      tooltip: compactTooltip(detailLabel),
      grid: {
        left: 148,
        right: 48,
        top: 22,
        bottom: 28,
        containLabel: false,
      },
      xAxis: {
        type: "value",
        min: 0,
        max: Math.max(axisMax, 0.25),
        axisLabel: {
          color: "#8A9CA5",
          fontSize: 9,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        axisLine: { show: true, lineStyle: { color: "#CBD8DE" } },
        axisTick: { show: false },
        splitLine: { lineStyle: { color: "#E5ECEF" } },
      },
      yAxis: {
        type: "category",
        inverse: true,
        data: rows.map((row) => row.category),
        axisLabel: {
          width: 132,
          overflow: "truncate",
          color: "#4D636E",
          fontSize: 10,
        },
        axisLine: { show: false },
        axisTick: { show: false },
      },
      series: [
        {
          type: "bar",
          barWidth: 17,
          showBackground: true,
          backgroundStyle: { color: "#F0F3F4" },
          itemStyle: {
            color: MODEL_COLORS[model],
            borderRadius: [0, 2, 2, 0],
          },
          label: {
            show: true,
            position: "right",
            color: "#405660",
            fontSize: 9,
            fontWeight: 700,
            formatter: ({ data: row }) =>
              formatPercentage(row.percentage),
          },
          data: rows.map((row) => ({
            ...row,
            model,
            value: row.percentage,
          })),
        },
      ],
    };
  };

  const createBarCards = ({
    containerId,
    dataKey,
    detailLabel,
    chartKey,
    nested = false,
  }) => {
    const container = document.getElementById(containerId);
    if (!container) return [];
    const data = getData();
    disposeGroup(chartKey);
    container.innerHTML = "";

    return MODELS.map((model) => {
      const card = document.createElement(nested ? "section" : "article");
      card.className = nested
        ? "leads-model-panel"
        : "card leads-chart-card leads-chart-card--bar";
      card.innerHTML = `
        ${modelHeader(model, data.sampleSizes[model], detailLabel)}
        <div class="leads-chart-canvas"></div>`;
      container.append(card);
      const canvas = card.querySelector(".leads-chart-canvas");
      const chart = window.echarts.init(canvas, null, {
        renderer: "canvas",
      });
      chart.setOption(
        createBarOption({
          model,
          rows: data[dataKey][model],
          detailLabel,
        }),
      );
      charts.set(`${chartKey}-${model}`, chart);
      return chart;
    });
  };

  const initialize = () => {
    createPurchaseTypeChart();
    renderVehicleTables({
      containerId: "leads-current-vehicle-tables",
      dataKey: "currentVehicle",
      subtitle: "当前家庭用车 Top 9",
    });
    renderVehicleTables({
      containerId: "leads-previous-vehicle-tables",
      dataKey: "previousVehicle",
      subtitle: "计划替换车型 Top 9",
    });
    createBarCards({
      containerId: "leads-channel-charts",
      dataKey: "channel",
      detailLabel: "信息来源",
      chartKey: "leads-channel",
    });
    createBarCards({
      containerId: "leads-motivation-charts",
      dataKey: "motivation",
      detailLabel: "购车动机",
      chartKey: "leads-motivation",
      nested: true,
    });
    createBarCards({
      containerId: "leads-consideration-charts",
      dataKey: "consideration",
      detailLabel: "关注原因",
      chartKey: "leads-consideration",
      nested: true,
    });
  };

  const resizeAll = () => charts.forEach((chart) => chart.resize());

  window.LeadsIntent = Object.freeze({
    createPurchaseTypeChart,
    renderVehicleTables,
    createBarCards,
    initialize,
    resizeAll,
  });
})();
