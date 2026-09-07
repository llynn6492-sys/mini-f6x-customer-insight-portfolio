(() => {
  const config = window.QUARTERLY_DASHBOARD_CONFIG;
  const charts = new Map();
  const selections = [
    { model: "F65", series: "Cooper C" },
    { model: "F65", series: "Cooper S" },
    { model: "F66", series: "Cooper C" },
    { model: "F66", series: "Cooper S" },
  ];
  let profileSelection = { model: "F65", series: "Cooper C" };
  let purchaseSelection = { model: "F65", series: "Cooper C" };
  let motivationSelection = { model: "F65", series: "Cooper C" };
  const modelColors = { F65: "#0066B3", F66: "#259AC5" };
  const categoryColors = [
    "#CDE7EF",
    "#9FD3E1",
    "#68BDD3",
    "#2EA5C3",
    "#0788B0",
    "#006D9E",
    "#07527C",
    "#083A59",
  ];
  const ageMidpoints = {
    "20岁以下": 19,
    "20-24岁": 22,
    "25-29岁": 27,
    "30-34岁": 32,
    "35-39岁": 37,
    "40-44岁": 42,
    "45-49岁": 47,
    "50岁及以上": 53,
  };

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

  const selectionKey = ({ model, series }) =>
    `${model}-${series}`.toLowerCase().replaceAll(/\s+/g, "-");

  const tooltip = (detailLabel, selection) => ({
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
        `<div>车型：<b>${escapeHtml(selection.model)}</b></div>`,
        `<div>车系：<b>${escapeHtml(selection.series)}</b></div>`,
        `<div>${escapeHtml(detailLabel)}：<b>${escapeHtml(data.category)}</b></div>`,
        `<div>N：<b>${data.count ?? "N/A"}</b></div>`,
        `<div>占比：<b>${formatPercentage(data.percentage)}</b></div>`,
        detailLabel === "性别" && Number.isFinite(data.averageAge)
          ? `<div>平均年龄：<b>${data.averageAge.toFixed(1)}岁</b></div>`
          : "",
        "</div>",
      ].join(""),
  });

  const showEmpty = (container, message = "暂无可用数据") => {
    if (!container) return;
    window.echarts?.getInstanceByDom(container)?.dispose();
    container.classList.remove("chart-canvas--rendered");
    container.innerHTML = `<span class="chart-status">${message}</span>`;
  };

  const renderChart = (containerId, rows, optionFactory, selection) => {
    const container = document.getElementById(containerId);
    if (!container) return null;
    charts.get(containerId)?.dispose();
    charts.delete(containerId);
    if (!rows.length) {
      showEmpty(container);
      return null;
    }
    container.innerHTML = "";
    container.classList.add("chart-canvas--rendered");
    const chart = window.echarts.init(container, null, { renderer: "canvas" });
    chart.setOption(optionFactory(rows, selection), true);
    charts.set(containerId, chart);
    return chart;
  };

  const horizontalOption = (detailLabel) => (rows, selection) => {
    const maxValue = Math.max(...rows.map((row) => row.percentage), 0.1);
    return {
      animationDuration: 320,
      backgroundColor: "#FFFFFF",
      tooltip: tooltip(detailLabel, selection),
      grid: { left: 126, right: 46, top: 18, bottom: 28 },
      xAxis: {
        type: "value",
        min: 0,
        max: Math.ceil((maxValue + 0.08) * 20) / 20,
        axisLabel: {
          color: "#8A9CA5",
          fontSize: 9,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        axisTick: { show: false },
        axisLine: { lineStyle: { color: "#CBD8DE" } },
        splitLine: { lineStyle: { color: "#E5ECEF" } },
      },
      yAxis: {
        type: "category",
        inverse: true,
        data: rows.map((row) => row.category),
        axisTick: { show: false },
        axisLine: { show: false },
        axisLabel: {
          color: "#405660",
          fontSize: 9,
          width: 112,
          overflow: "truncate",
        },
      },
      series: [
        {
          type: "bar",
          barMaxWidth: 18,
          showBackground: true,
          backgroundStyle: { color: "#F0F4F6", borderRadius: 2 },
          data: rows.map((row) => ({
            ...row,
            value: row.percentage,
            itemStyle: {
              color: modelColors[selection.model],
              borderRadius: [0, 2, 2, 0],
            },
          })),
          label: {
            show: true,
            position: "right",
            color: "#324A56",
            fontSize: 9,
            fontWeight: 700,
            formatter: ({ data }) => formatPercentage(data.percentage),
          },
          emphasis: { itemStyle: { opacity: 0.8 } },
        },
      ],
    };
  };

  const stackedOption = (detailLabel) => (rows, selection) => ({
    animationDuration: 340,
    backgroundColor: "#FFFFFF",
    color: categoryColors,
    tooltip: tooltip(detailLabel, selection),
    legend: {
      top: 4,
      left: 8,
      itemWidth: 7,
      itemHeight: 7,
      textStyle: { color: "#607781", fontSize: 8 },
      data: rows.map((row) => row.category),
    },
    grid: { left: 44, right: 20, top: 54, bottom: 34 },
    xAxis: {
      type: "category",
      data: [selection.series],
      axisTick: { show: false },
      axisLine: { lineStyle: { color: "#CBD8DE" } },
      axisLabel: { color: "#405660", fontSize: 9, fontWeight: 700 },
    },
    yAxis: {
      type: "value",
      min: 0,
      max: 1,
      axisTick: { show: false },
      axisLine: { show: false },
      axisLabel: {
        color: "#8A9CA5",
        fontSize: 8,
        formatter: (value) => `${Math.round(value * 100)}%`,
      },
      splitLine: { lineStyle: { color: "#E5ECEF" } },
    },
    series: rows.map((row, index) => ({
      name: row.category,
      type: "bar",
      stack: "total",
      barWidth: "44%",
      data: [{ ...row, value: row.percentage }],
      itemStyle: {
        color: categoryColors[index % categoryColors.length],
        borderColor: "#FFFFFF",
        borderWidth: 1,
      },
      label: {
        show: row.percentage >= 0.09,
        position: "inside",
        color: "#FFFFFF",
        fontSize: 8,
        fontWeight: 700,
        formatter: () => formatPercentage(row.percentage),
      },
    })),
  });

  const practicalUserOption = (detailLabel) => (rows, selection) => {
    const option = stackedOption(detailLabel)(rows, selection);
    const practicalUserColors = ["#0066B3", "#5BA9D6"];
    option.series.forEach((series, index) => {
      series.itemStyle.color =
        practicalUserColors[index % practicalUserColors.length];
    });
    return option;
  };

  const pieOption = (detailLabel) => (rows, selection) => ({
    animationDuration: 350,
    backgroundColor: "#FFFFFF",
    color: categoryColors.slice(2),
    tooltip: tooltip(detailLabel, selection),
    legend: {
      bottom: 3,
      left: "center",
      itemWidth: 7,
      itemHeight: 7,
      textStyle: { color: "#607781", fontSize: 8 },
    },
    graphic: [
      {
        type: "text",
        left: "center",
        top: "42%",
        style: {
          text: `N=${rows[0].sampleSize}`,
          fill: "#17313D",
          font: '700 10px "PingFang SC", sans-serif',
          textAlign: "center",
        },
      },
    ],
    series: [
      {
        type: "pie",
        radius: ["34%", "60%"],
        center: ["50%", "44%"],
        avoidLabelOverlap: true,
        itemStyle: { borderColor: "#FFFFFF", borderWidth: 2 },
        label: {
          color: "#405660",
          fontSize: 8,
          formatter: ({ data }) =>
            `${data.category}\n${formatPercentage(data.percentage)}`,
        },
        labelLine: {
          length: 7,
          length2: 5,
          lineStyle: { color: "#A5B8C1" },
        },
        data: rows.map((row) => ({
          ...row,
          name: row.category,
          value: row.count,
        })),
      },
    ],
  });

  const weightedAverageAge = (rows) => {
    const valid = rows.filter((row) =>
      Number.isFinite(ageMidpoints[row.category]),
    );
    const total = valid.reduce((sum, row) => sum + row.count, 0);
    if (!total) return "N/A";
    return (
      valid.reduce(
        (sum, row) => sum + row.count * ageMidpoints[row.category],
        0,
      ) / total
    ).toFixed(1);
  };

  const topCategory = (rows) =>
    rows.reduce(
      (top, row) => (!top || row.percentage > top.percentage ? row : top),
      null,
    )?.category ?? "N/A";

  const sampleLabel = (rows) =>
    Number.isFinite(rows[0]?.sampleSize) ? `N=${rows[0].sampleSize}` : "N=N/A";

  const getRows = (sheet, metric, selection) =>
    window.QuarterlyData.getMetric({ sheet, metric, ...selection });

  const profileChartMarkup = (id, title, modifier = "") => `
    <article class="quarterly-mini-chart ${modifier}">
      <header><h4>${title}</h4><span id="${id}-n">N=N/A</span></header>
      <div class="quarterly-mini-chart__canvas" id="${id}"></div>
    </article>`;

  const renderProfileSeries = async () => {
    const host = document.getElementById("quarterly-profile-series");
    if (!host) return;
    const groups = await Promise.all(
      [profileSelection].map(async (selection) => {
        const get = (metric) => getRows(config.sheets.profile, metric, selection);
        const [gender, age, marriage, education, income, industry] =
          await Promise.all([
            get("Gender"),
            get("Age"),
            get("Marital Status"),
            get("Education"),
            get("Household Income"),
            get("Industry"),
          ]);
        return {
          selection,
          metrics: { gender, age, marriage, education, income, industry },
        };
      }),
    );

    host.innerHTML = groups
      .map(({ selection, metrics }) => {
        const key = selectionKey(selection);
        const female = metrics.gender.find((row) => row.category === "女");
        return `
          <article class="card quarterly-profile-series-card">
            <header class="quarterly-series-card__header">
              <div><i style="background:${modelColors[selection.model]}"></i>
                <strong>${selection.model}</strong><span>${selection.series}</span>
              </div>
              <b>${sampleLabel(metrics.gender)}</b>
            </header>
            <div class="quarterly-series-snapshot">
              <span><small>平均年龄</small><strong>${weightedAverageAge(metrics.age)}</strong></span>
              <span><small>女性占比</small><strong>${formatPercentage(female?.percentage)}</strong></span>
              <span><small>主要婚姻状态</small><strong>${escapeHtml(topCategory(metrics.marriage))}</strong></span>
              <span><small>主要行业</small><strong>${escapeHtml(topCategory(metrics.industry))}</strong></span>
            </div>
            <div class="quarterly-profile-mini-grid">
              ${profileChartMarkup(`q-profile-gender-${key}`, "性别分布")}
              ${profileChartMarkup(`q-profile-age-${key}`, "年龄分布")}
              ${profileChartMarkup(`q-profile-marriage-${key}`, "婚姻状态")}
              ${profileChartMarkup(`q-profile-education-${key}`, "学历分布")}
              ${profileChartMarkup(`q-profile-income-${key}`, "税后家庭收入", "quarterly-mini-chart--income")}
              ${profileChartMarkup(`q-profile-industry-${key}`, "行业信息", "quarterly-mini-chart--industry")}
            </div>
          </article>`;
      })
      .join("");

    groups.forEach(({ selection, metrics }) => {
      const key = selectionKey(selection);
      [
        ["gender", metrics.gender, horizontalOption("性别")],
        ["age", metrics.age, stackedOption("年龄段")],
        ["marriage", metrics.marriage, pieOption("婚姻状态")],
        ["education", metrics.education, stackedOption("学历")],
        ["income", metrics.income, stackedOption("家庭年收入")],
        ["industry", metrics.industry, horizontalOption("行业")],
      ].forEach(([metric, rows, option]) => {
        const id = `q-profile-${metric}-${key}`;
        const n = document.getElementById(`${id}-n`);
        if (n) n.textContent = sampleLabel(rows);
        renderChart(id, rows, option, selection);
      });
    });
  };

  const renderSeriesGrid = async ({
    hostId,
    sheet,
    metric,
    title,
    detailLabel,
    optionFactory,
    selectionList = selections,
  }) => {
    const host = document.getElementById(hostId);
    if (!host) return;
    const groups = await Promise.all(
      selectionList.map(async (selection) => ({
        selection,
        rows: await getRows(sheet, metric, selection),
      })),
    );
    host.innerHTML = groups
      .map(({ selection, rows }) => {
        const id = `${hostId}-${selectionKey(selection)}`;
        return `
          <article class="card quarterly-series-chart-card">
            <header class="quarterly-series-card__header">
              <div><i style="background:${modelColors[selection.model]}"></i>
                <strong>${selection.model}</strong><span>${selection.series} · ${title}</span>
              </div>
              <b>${sampleLabel(rows)}</b>
            </header>
            <div class="quarterly-series-chart-canvas" id="${id}"></div>
          </article>`;
      })
      .join("");
    groups.forEach(({ selection, rows }) =>
      renderChart(
        `${hostId}-${selectionKey(selection)}`,
        rows,
        optionFactory(detailLabel),
        selection,
      ),
    );
  };

  const renderPurchaseSeries = () => {
    const host = document.getElementById("quarterly-purchase-series");
    if (!host) return;
    host.classList.add("quarterly-purchase-series-layout");
    host.innerHTML = `
      <section class="quarterly-series-metric-group">
        <h3>购车类型分布</h3>
        <div class="quarterly-series-chart-grid" id="quarterly-purchase-type-series"></div>
      </section>
      <section class="quarterly-series-metric-group">
        <h3>实际用车人比例</h3>
        <div class="quarterly-series-chart-grid" id="quarterly-user-series"></div>
      </section>`;
    return Promise.all([
      renderSeriesGrid({
        hostId: "quarterly-purchase-type-series",
        sheet: config.sheets.purchaseBehavior,
        metric: "购车类型",
        title: "购车类型",
        detailLabel: "购车类型",
        optionFactory: pieOption,
        selectionList: [purchaseSelection],
      }),
      renderSeriesGrid({
        hostId: "quarterly-user-series",
        sheet: config.sheets.purchaseBehavior,
        metric: "实际用车人",
        title: "实际用车人",
        detailLabel: "实际用车人",
        optionFactory: practicalUserOption,
        selectionList: [purchaseSelection],
      }),
    ]);
  };

  const renderMotivationSeries = () =>
    Promise.all([
      renderSeriesGrid({
        hostId: "quarterly-motivation-series",
        sheet: config.sheets.purchaseMotivation,
        metric: "购车动机",
        title: "购车动机",
        detailLabel: "购车动机",
        optionFactory: horizontalOption,
        selectionList: [motivationSelection],
      }),
      renderSeriesGrid({
        hostId: "quarterly-reason-series",
        sheet: config.sheets.purchaseMotivation,
        metric: "购车原因",
        title: "购车原因",
        detailLabel: "购车原因",
        optionFactory: horizontalOption,
        selectionList: [motivationSelection],
      }),
    ]);

  const renderJourneySeries = () =>
    renderSeriesGrid({
      hostId: "quarterly-channel-series",
      sheet: config.sheets.decisionJourney,
      metric: "渠道",
      title: "信息渠道",
      detailLabel: "信息渠道",
      optionFactory: horizontalOption,
    });

  const pageRenderers = {
    profile: renderProfileSeries,
    "purchase-behavior": renderPurchaseSeries,
    "motivation-product": renderMotivationSeries,
    "purchase-journey": renderJourneySeries,
  };

  const refresh = async (page) => {
    const renderer = pageRenderers[page];
    if (!renderer) return;
    try {
      await renderer();
    } catch (error) {
      console.error("季度车系下钻渲染失败：", error);
    }
  };

  const initialize = () => {
    const profileModelFilter = document.getElementById(
      "quarterly-profile-model-filter",
    );
    const profileSeriesFilter = document.getElementById(
      "quarterly-profile-series-filter",
    );
    const purchaseModelFilter = document.getElementById(
      "quarterly-purchase-model-filter",
    );
    const purchaseSeriesFilter = document.getElementById(
      "quarterly-purchase-series-filter",
    );
    const motivationModelFilters = [
      document.getElementById("quarterly-motivation-model-filter"),
      document.getElementById("quarterly-reason-model-filter"),
    ].filter(Boolean);
    const motivationSeriesFilters = [
      document.getElementById("quarterly-motivation-series-filter"),
      document.getElementById("quarterly-reason-series-filter"),
    ].filter(Boolean);
    const refreshProfileSelection = () => {
      profileSelection = {
        model: profileModelFilter?.value ?? "F65",
        series: profileSeriesFilter?.value ?? "Cooper C",
      };
      renderProfileSeries();
    };
    profileModelFilter?.addEventListener("change", refreshProfileSelection);
    profileSeriesFilter?.addEventListener("change", refreshProfileSelection);
    const refreshPurchaseSelection = () => {
      purchaseSelection = {
        model: purchaseModelFilter?.value ?? "F65",
        series: purchaseSeriesFilter?.value ?? "Cooper C",
      };
      renderPurchaseSeries();
    };
    purchaseModelFilter?.addEventListener(
      "change",
      refreshPurchaseSelection,
    );
    purchaseSeriesFilter?.addEventListener(
      "change",
      refreshPurchaseSelection,
    );
    const refreshMotivationSelection = (model, series) => {
      motivationSelection = { model, series };
      motivationModelFilters.forEach((filter) => {
        filter.value = model;
      });
      motivationSeriesFilters.forEach((filter) => {
        filter.value = series;
      });
      renderMotivationSeries();
    };
    motivationModelFilters.forEach((filter) => {
      filter.addEventListener("change", () =>
        refreshMotivationSelection(
          filter.value,
          motivationSeriesFilters[0]?.value ?? "Cooper C",
        ),
      );
    });
    motivationSeriesFilters.forEach((filter) => {
      filter.addEventListener("change", () =>
        refreshMotivationSelection(
          motivationModelFilters[0]?.value ?? "F65",
          filter.value,
        ),
      );
    });
    window.CustomerProfileSection?.({ audience: "owner" });
    window.addEventListener("dashboard:pagechange", (event) => {
      if (pageRenderers[event.detail?.page]) {
        window.setTimeout(() => refresh(event.detail.page), 60);
      }
    });
    window.addEventListener("dashboard:sectionready", (event) => {
      if (pageRenderers[event.detail?.page]) {
        refresh(event.detail.page);
      }
    });
    window.addEventListener("resize", () =>
      charts.forEach((chart) => chart.resize()),
    );
    window.setTimeout(() => refresh("profile"), 120);
  };

  window.QuarterlyDashboard = Object.freeze({
    initialize,
    refresh,
  });

  initialize();
})();
