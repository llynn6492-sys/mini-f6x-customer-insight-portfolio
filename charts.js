(() => {
  const WORKBOOK_URL =
    window.DASHBOARD_WORKBOOK_URL ??
    "data/mock_dashboard_data.xlsx?v=20260727-2";
  const DEMOGRAPHIC_SHEET = "demographic_data";
  const SAMPLE_SIZE_SHEET = "sample_size";
  const MODEL_ORDER = ["F65", "F66"];
  const DEFAULT_GENDER_CATEGORIES = ["男", "女"];
  const DEFAULT_AGE_CATEGORIES = [
    "20岁以下",
    "20-24岁",
    "25-29岁",
    "30-34岁",
    "35-39岁",
    "40-44岁",
    "45-49岁",
    "50岁及以上",
  ];
  const AGE_COLORS = [
    "#CDE7EF",
    "#9FD3E1",
    "#68BDD3",
    "#2EA5C3",
    "#0788B0",
    "#006D9E",
    "#07527C",
    "#083A59",
  ];
  const CHART_METRICS = {
    gender: "Gender",
    age: "Age",
    "marital-status": "Marital Status",
    education: "Education",
    "household-income": "Household Income",
    industry: "Industry",
  };
  const MODEL_COLORS = {
    F65: "#0066B3",
    F66: "#259AC5",
  };
  const METRIC_ALIASES = {
    Gender: "性别",
    Age: "年龄",
    "Marital Status": "婚姻状态",
    Education: "学历",
    "Household Income": "税后家庭收入",
    Industry: "行业",
  };
  const workbookCache = new Map();
  const AUDIENCE_ALIASES = {
    owner: "owner",
    Owner: "owner",
    车主: "owner",
    prospect: "prospect",
    Prospect: "prospect",
    潜客: "prospect",
  };
  const OWNER_SNAPSHOT = {
    F65: {
      sample: "222",
      averageAge: "35.6",
      femaleShare: "55.9%",
      familyStatus: "已婚有孩子",
      industry: "IT / 互联网",
    },
    F66: {
      sample: "245",
      averageAge: "34.6",
      femaleShare: "58.0%",
      familyStatus: "已婚有孩子",
      industry: "IT / 互联网",
    },
  };

  const loadEmbeddedWorkbook = () => {
    const encoded = window.DASHBOARD_WORKBOOK_BASE64;
    if (!encoded) {
      throw new Error("Embedded workbook fallback is unavailable.");
    }
    const binary = window.atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    return window.XLSX.read(bytes, { type: "array" });
  };

  const formatPercentage = (value) => {
    if (value === null || value === undefined || value === "") {
      return "N/A";
    }

    const percentage = Number(value);

    if (!Number.isFinite(percentage)) {
      return "N/A";
    }

    return `${(percentage * 100).toFixed(1)}%`;
  };

  const formatValue = (value) =>
    value === null || value === undefined || value === "" ? "N/A" : value;

  const formatSampleSize = (value) => {
    const sampleSize = Number(value);
    return Number.isFinite(sampleSize)
      ? sampleSize.toLocaleString("en-US")
      : "N/A";
  };

  const escapeHtml = (value) =>
    String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");

  const showChartStatus = (container, message, isError = false) => {
    window.echarts?.getInstanceByDom(container)?.dispose();
    container.querySelectorAll("*").forEach((element) => {
      window.echarts?.getInstanceByDom(element)?.dispose();
    });
    container.classList.remove("chart-canvas--rendered");
    container.innerHTML = "";

    const status = document.createElement("span");
    status.className = `chart-status${isError ? " chart-status--error" : ""}`;
    status.textContent = message;
    container.append(status);
  };

  const loadWorkbook = async (workbookUrl) => {
    if (!window.XLSX) {
      throw new Error("The Excel reader could not be loaded.");
    }

    if (!workbookCache.has(workbookUrl)) {
      workbookCache.set(
        workbookUrl,
        (async () => {
          if (window.location.protocol === "file:") {
            return loadEmbeddedWorkbook();
          }
          try {
            const response = await fetch(workbookUrl);
            if (!response.ok) {
              throw new Error(
                `Workbook request failed with status ${response.status}.`,
              );
            }
            const workbookBuffer = await response.arrayBuffer();
            return window.XLSX.read(workbookBuffer);
          } catch (error) {
            if (window.DASHBOARD_WORKBOOK_BASE64) {
              return loadEmbeddedWorkbook();
            }
            throw error;
          }
        })(),
      );
    }

    return workbookCache.get(workbookUrl);
  };

  const getSheetRows = (workbook, sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    return worksheet
      ? window.XLSX.utils.sheet_to_json(worksheet, {
          defval: null,
          raw: true,
        })
      : [];
  };

  const getCategoryOrder = (rows, fallbackCategories) => {
    const categorySortOrders = new Map();

    rows.forEach((row) => {
      const currentOrder = categorySortOrders.get(row.category);
      if (
        currentOrder === undefined ||
        row.sortOrder < currentOrder
      ) {
        categorySortOrders.set(row.category, row.sortOrder);
      }
    });

    if (categorySortOrders.size === 0) {
      return fallbackCategories;
    }

    return [...categorySortOrders.entries()]
      .sort(
        ([categoryA, orderA], [categoryB, orderB]) =>
          orderA - orderB || categoryA.localeCompare(categoryB, "zh-CN"),
      )
      .map(([category]) => category);
  };

  const getMetricSampleSize = (
    sampleRows,
    metric,
    model,
    audience = "owner",
  ) => {
    const matchesAudience = (row) =>
      normalizeAudience(row.Audience ?? row["受众"] ?? row["用户类型"]) ===
      audience;
    const exactMatch = sampleRows.find(
      (row) =>
        row.Metric === metric &&
        row.Model === model &&
        matchesAudience(row),
    );
    const legacyMatch = sampleRows.find(
      (row) =>
        (row.Metric === null ||
          row.Metric === undefined ||
          row.Metric === "") &&
        row.Model === model &&
        matchesAudience(row),
    );
    const sampleSize = Number(
      (exactMatch ?? legacyMatch)?.Sample_Size,
    );

    return Number.isFinite(sampleSize) ? sampleSize : null;
  };

  const normalizeAudience = (value) =>
    AUDIENCE_ALIASES[value] ?? (value ? String(value).toLowerCase() : "owner");

  const loadCustomerProfileMetric = async (
    metric,
    workbookUrl,
    fallbackCategories = [],
    {
      sheetName = DEMOGRAPHIC_SHEET,
      purchaseType = null,
      audience = "owner",
    } = {},
  ) => {
    const workbook = await loadWorkbook(workbookUrl);
    const demographicRows = getSheetRows(workbook, sheetName);
    const sampleRows = getSheetRows(workbook, SAMPLE_SIZE_SHEET);
    const hasProspectRows = demographicRows.some(
      (row) =>
        normalizeAudience(
          row.Audience ?? row["受众"] ?? row["用户类型"],
        ) === "prospect",
    );
    if (
      audience === "prospect" &&
      sheetName === DEMOGRAPHIC_SHEET &&
      !hasProspectRows &&
      window.PROSPECT_PROFILE_MOCK
    ) {
      demographicRows.push(...window.PROSPECT_PROFILE_MOCK.rows);
      sampleRows.push(...window.PROSPECT_PROFILE_MOCK.sampleRows);
    }
    const metricAlias = METRIC_ALIASES[metric] ?? metric;

    const rows = demographicRows
      .filter((row) => {
        const rowMetric = row.Metric ?? row["指标"];
        const rowPurchaseType = row["购车类型"];
        const rowAudience = normalizeAudience(
          row.Audience ?? row["受众"] ?? row["用户类型"],
        );
        return (
          (rowMetric === metric || rowMetric === metricAlias) &&
          rowAudience === audience &&
          (!purchaseType || rowPurchaseType === purchaseType)
        );
      })
      .map((row, rowIndex) => ({
        model: row.Model ?? row["车型"],
        category: row.Category ?? row["分类"],
        count:
          (row.Count ?? row.N) === null ||
          (row.Count ?? row.N) === undefined
            ? null
            : Number(row.Count ?? row.N),
        percentage:
          (row.Percentage ?? row["占比"]) === null ||
          (row.Percentage ?? row["占比"]) === undefined
            ? null
            : Number(row.Percentage ?? row["占比"]),
        averageAge:
          (row.Average_Age ?? row["平均年龄"]) === null ||
          (row.Average_Age ?? row["平均年龄"]) === undefined
            ? null
            : Number(row.Average_Age ?? row["平均年龄"]),
        sortOrder: Number.isFinite(
          Number(row.Sort_Order ?? row["排序"]),
        )
          ? Number(row.Sort_Order ?? row["排序"])
          : rowIndex,
      }))
      .filter(
        (row) =>
          MODEL_ORDER.includes(row.model) &&
          typeof row.category === "string" &&
          row.category.length > 0,
      );

    const derivedSampleSizes = Object.fromEntries(
      MODEL_ORDER.map((model) => {
        const total = rows
          .filter((row) => row.model === model)
          .reduce((sum, row) => sum + (row.count ?? 0), 0);
        return [model, total || null];
      }),
    );
    const sampleSizes =
      sheetName === DEMOGRAPHIC_SHEET
        ? Object.fromEntries(
            MODEL_ORDER.map((model) => [
              model,
              getMetricSampleSize(sampleRows, metric, model, audience),
            ]),
          )
        : derivedSampleSizes;

    return {
      rows,
      categories: getCategoryOrder(rows, fallbackCategories),
      sampleSizes,
    };
  };

  const updateChartLegend = (container, sampleSizes) => {
    const legend = container
      .closest(".chart-card")
      ?.querySelector(".chart-legend");
    if (!legend) {
      return;
    }
    legend.querySelectorAll("span").forEach((label) => {
      const model = MODEL_ORDER.find((item) =>
        label.textContent.includes(item),
      );
      if (!model) {
        return;
      }
      const dot = label.querySelector(".legend-dot");
      label.replaceChildren();
      if (dot) {
        label.append(dot);
      }
      label.append(
        document.createTextNode(
          `${model} N=${formatSampleSize(sampleSizes[model])}`,
        ),
      );
    });
  };

  const hydrateChartSampleSizes = async ({
    workbookUrl = WORKBOOK_URL,
    audience = "owner",
  } = {}) => {
    const workbook = await loadWorkbook(workbookUrl);
    const sampleRows = getSheetRows(workbook, SAMPLE_SIZE_SHEET);
    if (
      audience === "prospect" &&
      window.PROSPECT_PROFILE_MOCK &&
      !sampleRows.some(
        (row) =>
          normalizeAudience(
            row.Audience ?? row["受众"] ?? row["用户类型"],
          ) === "prospect",
      )
    ) {
      sampleRows.push(...window.PROSPECT_PROFILE_MOCK.sampleRows);
    }

    document.querySelectorAll("[data-chart]").forEach((canvas) => {
      const metric = CHART_METRICS[canvas.dataset.chart];
      const legend = canvas.closest(".chart-card")?.querySelector(".chart-legend");

      if (!metric || !legend) {
        return;
      }

      legend.querySelectorAll("span").forEach((label) => {
        const model = MODEL_ORDER.find((item) =>
          label.textContent.includes(item),
        );

        if (!model) {
          return;
        }

        const sampleSize = getMetricSampleSize(
          sampleRows,
          metric,
          model,
          audience,
        );
        const dot = label.querySelector(".legend-dot");
        label.replaceChildren();
        if (dot) {
          label.append(dot);
        }
        label.append(
          document.createTextNode(
            `${model} N=${formatSampleSize(sampleSize)}`,
          ),
        );
      });
    });
  };

  const rowsForModel = (rows, model, categories) =>
    categories.map((category) => {
      const row = rows.find(
        (item) => item.model === model && item.category === category,
      );

      return {
        value: Number.isFinite(row?.percentage) ? row.percentage : 0,
        model,
        gender: category,
        count: Number.isFinite(row?.count) ? row.count : null,
        percentage: Number.isFinite(row?.percentage) ? row.percentage : null,
        averageAge: Number.isFinite(row?.averageAge) ? row.averageAge : null,
      };
    });

  const createModelSeries = (model, rows, categories) => ({
    name: model,
    type: "bar",
    data: rowsForModel(rows, model, categories),
    barWidth: 17,
    showBackground: true,
    backgroundStyle: {
      color: "#F1F4F5",
      borderRadius: 2,
    },
    itemStyle: {
      color: MODEL_COLORS[model],
      borderRadius: [0, 2, 2, 0],
    },
    emphasis: {
      disabled: false,
      itemStyle: {
        opacity: 0.84,
      },
    },
    label: {
      show: true,
      position: "insideRight",
      distance: 7,
      color: "#FFFFFF",
      fontFamily: 'Inter, "Helvetica Neue", Arial, sans-serif',
      fontSize: 10,
      fontWeight: 700,
      formatter: (params) => formatPercentage(params.data.percentage),
    },
  });

  const createGenderOption = (model, rows, categories) => ({
    animation: false,
    backgroundColor: "#FFFFFF",
    aria: {
      enabled: true,
      decal: {
        show: false,
      },
      description:
        "Horizontal bar chart comparing female and male customer percentages for MINI F65 and F66.",
    },
    tooltip: {
      trigger: "item",
      confine: true,
      transitionDuration: 0,
      enterable: false,
      backgroundColor: "rgba(255, 255, 255, 0.96)",
      borderColor: "rgba(156, 181, 192, 0.82)",
      borderWidth: 1,
      padding: [7, 9],
      textStyle: {
        color: "#203741",
        fontSize: 10,
      },
      extraCssText:
        "box-shadow: 0 8px 20px rgba(7,31,45,.14); border-radius: 4px; backdrop-filter: blur(4px);",
      position: (point) => [point[0] + 10, point[1] + 10],
      formatter: (params) => {
        const { model, gender, count, percentage, averageAge } = params.data;

        return [
          `<div style="line-height:1.5;white-space:nowrap">`,
          `<div>车型：<b>${escapeHtml(model)}</b></div>`,
          `<div>性别：<b>${escapeHtml(gender)}</b></div>`,
          `<div>人数：<b>${escapeHtml(formatValue(count))}</b></div>`,
          `<div>占比：<b>${escapeHtml(formatPercentage(percentage))}</b></div>`,
          Number.isFinite(averageAge)
            ? `<div>平均年龄：<b>${escapeHtml(averageAge.toFixed(1))}岁</b></div>`
            : "",
          `</div>`,
        ].join("");
      },
    },
    grid: {
      left: 60,
      right: 10,
      top: 48,
      bottom: 48,
      containLabel: false,
    },
    xAxis: {
      type: "value",
      min: 0,
      max: 0.65,
      interval: 0.2,
      axisLabel: { show: false },
      axisLine: { show: false },
      axisTick: { show: false },
      splitLine: {
        show: true,
        lineStyle: {
          color: "#E9EEF1",
          width: 1,
        },
      },
    },
    yAxis: {
      type: "category",
      inverse: true,
      data: categories,
      axisLabel: {
        color: "#536872",
        fontSize: 10,
        fontWeight: 500,
        margin: 11,
      },
      axisLine: { show: false },
      axisTick: { show: false },
    },
    series: [createModelSeries(model, rows, categories)],
  });

  const createGenderPanels = (container, sampleSizes, containerId) => {
    container
      .querySelectorAll(".gender-model-chart__canvas")
      .forEach((canvas) => window.echarts.getInstanceByDom(canvas)?.dispose());
    container.innerHTML = "";
    container.classList.add("chart-canvas--rendered");

    return MODEL_ORDER.map((model) => {
      const panel = document.createElement("section");
      panel.className = "gender-model-chart";
      panel.style.setProperty("--model-color", MODEL_COLORS[model]);
      panel.setAttribute("aria-label", `${model} gender distribution`);

      const header = document.createElement("div");
      header.className = "gender-model-chart__header";
      header.innerHTML = [
        '<i class="gender-model-chart__swatch" aria-hidden="true"></i>',
        `<span>${model}</span>`,
        `<span class="gender-model-chart__sample">N=${formatSampleSize(sampleSizes[model])}</span>`,
      ].join("");

      const canvas = document.createElement("div");
      canvas.className = "gender-model-chart__canvas";
      canvas.id = `${containerId}-${model.toLowerCase()}`;

      panel.append(header, canvas);
      container.append(panel);

      return { model, canvas };
    });
  };

  const createGenderChart = async ({
    containerId = "gender-chart",
    workbookUrl = WORKBOOK_URL,
    sheetName = DEMOGRAPHIC_SHEET,
    purchaseType = null,
    audience = "owner",
  } = {}) => {
    const container = document.getElementById(containerId);

    if (!container) {
      throw new Error(`Chart container "#${containerId}" was not found.`);
    }

    if (!window.echarts) {
      showChartStatus(
        container,
        "图表组件加载失败，请检查网络连接后刷新页面。",
        true,
      );
      throw new Error("ECharts is unavailable.");
    }

    showChartStatus(container, "正在加载性别数据…");

    try {
      const { rows, categories, sampleSizes } =
        await loadCustomerProfileMetric(
          "Gender",
          workbookUrl,
          DEFAULT_GENDER_CATEGORIES,
          { sheetName, purchaseType, audience },
        );
      const panels = createGenderPanels(container, sampleSizes, containerId);
      const charts = panels.map(({ model, canvas }) => {
        const chart = window.echarts.init(canvas, null, {
          renderer: "canvas",
        });
        chart.setOption(
          createGenderOption(model, rows, categories),
          true,
        );
        return chart;
      });

      const resizeObserver = new ResizeObserver(() => {
        charts.forEach((chart) => chart.resize());
      });
      panels.forEach(({ canvas }) => resizeObserver.observe(canvas));

      return charts;
    } catch (error) {
      const isLocalFile = window.location.protocol === "file:";
      const message = isLocalFile
        ? "Excel 数据需要通过本地服务器加载，请使用 http://localhost 打开看板。"
        : "性别数据加载失败，请检查数据文件后刷新页面。";

      showChartStatus(container, message, true);
      throw error;
    }
  };

  const createAgeDatum = (rows, model, category) => {
    const row = rows.find(
      (item) => item.model === model && item.category === category,
    );

    return {
      name: category,
      value: Number.isFinite(row?.percentage) ? row.percentage : 0,
      model,
      ageGroup: category,
      count: Number.isFinite(row?.count) ? row.count : null,
      percentage: Number.isFinite(row?.percentage) ? row.percentage : null,
    };
  };

  const createAgeOption = (rows, categories) => ({
    animation: false,
    backgroundColor: "#FFFFFF",
    aria: {
      enabled: true,
      decal: {
        show: false,
      },
      description:
        "Vertical stacked bar chart comparing the age distribution of MINI F65 and F66 customers.",
    },
    color: AGE_COLORS,
    legend: {
      type: "scroll",
      top: 2,
      left: 14,
      right: 14,
      itemWidth: 9,
      itemHeight: 9,
      itemGap: 11,
      pageIconSize: 9,
      pageTextStyle: {
        color: "#70818A",
        fontSize: 9,
      },
      textStyle: {
        color: "#536872",
        fontSize: 9,
      },
      data: categories,
    },
    tooltip: {
      trigger: "item",
      confine: true,
      transitionDuration: 0,
      enterable: false,
      backgroundColor: "rgba(255, 255, 255, 0.96)",
      borderColor: "rgba(156, 181, 192, 0.82)",
      borderWidth: 1,
      padding: [7, 9],
      textStyle: {
        color: "#203741",
        fontSize: 10,
      },
      extraCssText:
        "box-shadow: 0 8px 20px rgba(7,31,45,.14); border-radius: 4px; backdrop-filter: blur(4px);",
      position: (point) => [point[0] + 10, point[1] + 10],
      formatter: (params) => {
        const { model, ageGroup, count, percentage } = params.data;

        return [
          '<div style="line-height:1.5;white-space:nowrap">',
          `<div>车型：<b>${escapeHtml(model)}</b></div>`,
          `<div>年龄段：<b>${escapeHtml(ageGroup)}</b></div>`,
          `<div>人数：<b>${escapeHtml(formatValue(count))}</b></div>`,
          `<div>占比：<b>${escapeHtml(formatPercentage(percentage))}</b></div>`,
          "</div>",
        ].join("");
      },
    },
    grid: {
      left: 44,
      right: 14,
      top: 48,
      bottom: 32,
      containLabel: false,
    },
    xAxis: {
      type: "category",
      data: MODEL_ORDER,
      axisLabel: {
        color: "#40545F",
        fontSize: 11,
        fontWeight: 700,
        margin: 12,
      },
      axisLine: {
        lineStyle: {
          color: "#C9D6DC",
        },
      },
      axisTick: { show: false },
    },
    yAxis: {
      type: "value",
      min: 0,
      max: 1,
      interval: 0.2,
      axisLabel: {
        color: "#8A9AA2",
        fontSize: 9,
        formatter: (value) => `${Math.round(value * 100)}%`,
      },
      axisLine: { show: false },
      axisTick: { show: false },
      splitLine: {
        lineStyle: {
          color: "#E9EEF1",
          width: 1,
        },
      },
    },
    series: categories.map((category, categoryIndex) => ({
      name: category,
      type: "bar",
      stack: "age-total",
      barWidth: "48%",
      data: MODEL_ORDER.map((model) =>
        createAgeDatum(rows, model, category),
      ),
      itemStyle: {
        color: AGE_COLORS[categoryIndex % AGE_COLORS.length],
        borderColor: "#FFFFFF",
        borderWidth: 0.5,
      },
      emphasis: {
        focus: "series",
      },
      label: {
        show: true,
        position: "inside",
        color: "#FFFFFF",
        fontSize: 9,
        fontWeight: 700,
        textBorderColor: "rgba(7, 31, 45, 0.22)",
        textBorderWidth: 2,
        formatter: (params) =>
          Number.isFinite(params.data.percentage) &&
          params.data.percentage >= 0.055
            ? formatPercentage(params.data.percentage)
            : "",
      },
    })),
  });

  const createAgeChart = async ({
    containerId = "age-chart",
    workbookUrl = WORKBOOK_URL,
    sheetName = DEMOGRAPHIC_SHEET,
    purchaseType = null,
    audience = "owner",
  } = {}) => {
    const container = document.getElementById(containerId);

    if (!container) {
      throw new Error(`Chart container "#${containerId}" was not found.`);
    }

    if (!window.echarts) {
      showChartStatus(
        container,
        "图表组件加载失败，请检查网络连接后刷新页面。",
        true,
      );
      throw new Error("ECharts is unavailable.");
    }

    showChartStatus(container, "正在加载年龄数据…");

    try {
      const { rows, categories, sampleSizes } = await loadCustomerProfileMetric(
        "Age",
        workbookUrl,
        DEFAULT_AGE_CATEGORIES,
        { sheetName, purchaseType, audience },
      );

      container.innerHTML = "";
      container.classList.add("chart-canvas--rendered");
      updateChartLegend(container, sampleSizes);

      const chart = window.echarts.init(container, null, {
        renderer: "canvas",
      });

      chart.setOption(createAgeOption(rows, categories), true);

      const resizeObserver = new ResizeObserver(() => chart.resize());
      resizeObserver.observe(container);

      return chart;
    } catch (error) {
      const isLocalFile = window.location.protocol === "file:";
      const message = isLocalFile
        ? "Excel 数据需要通过本地服务器加载，请使用 http://localhost 打开看板。"
        : "年龄数据加载失败，请检查数据文件后刷新页面。";

      showChartStatus(container, message, true);
      throw error;
    }
  };

  const createCompactTooltip = (detailLabel, detailKey) => ({
    trigger: "item",
    confine: true,
    transitionDuration: 0,
    enterable: false,
    backgroundColor: "rgba(255, 255, 255, 0.96)",
    borderColor: "rgba(156, 181, 192, 0.82)",
    borderWidth: 1,
    padding: [7, 9],
    textStyle: {
      color: "#203741",
      fontSize: 10,
    },
    extraCssText:
      "box-shadow: 0 8px 20px rgba(7,31,45,.14); border-radius: 4px; backdrop-filter: blur(4px);",
    position: (point) => [point[0] + 10, point[1] + 10],
    formatter: (params) => {
      const { model, count, percentage } = params.data;
      const detail = params.data[detailKey];

      return [
        '<div style="line-height:1.5;white-space:nowrap">',
        `<div>车型：<b>${escapeHtml(model)}</b></div>`,
        `<div>${escapeHtml(detailLabel)}: <b>${escapeHtml(formatValue(detail))}</b></div>`,
        `<div>人数：<b>${escapeHtml(formatValue(count))}</b></div>`,
        `<div>占比：<b>${escapeHtml(formatPercentage(percentage))}</b></div>`,
        "</div>",
      ].join("");
    },
  });

  const createDistributionDatum = (
    rows,
    model,
    category,
    detailKey,
  ) => {
    const row = rows.find(
      (item) => item.model === model && item.category === category,
    );

    return {
      value: Number.isFinite(row?.percentage) ? row.percentage : 0,
      model,
      [detailKey]: category,
      count: Number.isFinite(row?.count) ? row.count : null,
      percentage: Number.isFinite(row?.percentage) ? row.percentage : null,
    };
  };

  const createStackedDistributionOption = ({
    rows,
    categories,
    detailLabel,
    detailKey,
    stackId,
  }) => ({
    animation: false,
    backgroundColor: "#FFFFFF",
    aria: {
      enabled: true,
      decal: { show: false },
    },
    color: AGE_COLORS,
    legend: {
      type: "scroll",
      top: 2,
      left: 12,
      right: 12,
      itemWidth: 9,
      itemHeight: 9,
      itemGap: 10,
      pageIconSize: 9,
      pageTextStyle: {
        color: "#70818A",
        fontSize: 9,
      },
      textStyle: {
        color: "#536872",
        fontSize: 9,
      },
      data: categories,
    },
    tooltip: createCompactTooltip(detailLabel, detailKey),
    grid: {
      left: 44,
      right: 14,
      top: 48,
      bottom: 32,
      containLabel: false,
    },
    xAxis: {
      type: "category",
      data: MODEL_ORDER,
      axisLabel: {
        color: "#40545F",
        fontSize: 11,
        fontWeight: 700,
        margin: 12,
      },
      axisLine: {
        lineStyle: { color: "#C9D6DC" },
      },
      axisTick: { show: false },
    },
    yAxis: {
      type: "value",
      min: 0,
      max: 1,
      interval: 0.2,
      axisLabel: {
        color: "#8A9AA2",
        fontSize: 9,
        formatter: (value) => `${Math.round(value * 100)}%`,
      },
      axisLine: { show: false },
      axisTick: { show: false },
      splitLine: {
        lineStyle: {
          color: "#E9EEF1",
          width: 1,
        },
      },
    },
    series: categories.map((category, categoryIndex) => ({
      name: category,
      type: "bar",
      stack: stackId,
      barWidth: "48%",
      data: MODEL_ORDER.map((model) =>
        createDistributionDatum(rows, model, category, detailKey),
      ),
      itemStyle: {
        color: AGE_COLORS[categoryIndex % AGE_COLORS.length],
        borderColor: "#FFFFFF",
        borderWidth: 0.5,
      },
      emphasis: {
        focus: "series",
      },
      label: {
        show: true,
        position: "inside",
        color: "#FFFFFF",
        fontSize: 9,
        fontWeight: 700,
        textBorderColor: "rgba(7, 31, 45, 0.22)",
        textBorderWidth: 2,
        formatter: (params) =>
          Number.isFinite(params.data.percentage) &&
          params.data.percentage >= 0.06
            ? formatPercentage(params.data.percentage)
            : "",
      },
    })),
  });

  const createSingleChart = async ({
    containerId,
    loadingMessage,
    errorMessage,
    optionFactory,
  }) => {
    const container = document.getElementById(containerId);

    if (!container) {
      throw new Error(`Chart container "#${containerId}" was not found.`);
    }

    if (!window.echarts) {
      showChartStatus(
        container,
        "图表组件加载失败，请检查网络连接后刷新页面。",
        true,
      );
      throw new Error("ECharts is unavailable.");
    }

    showChartStatus(container, loadingMessage);

    try {
      const option = await optionFactory(container);
      container.innerHTML = "";
      container.classList.add("chart-canvas--rendered");

      const chart = window.echarts.init(container, null, {
        renderer: "canvas",
      });

      chart.setOption(option, true);

      const resizeObserver = new ResizeObserver(() => chart.resize());
      resizeObserver.observe(container);
      return chart;
    } catch (error) {
      const isLocalFile = window.location.protocol === "file:";
      showChartStatus(
        container,
        isLocalFile
          ? "Excel 数据需要通过本地服务器加载，请使用 http://localhost 打开看板。"
          : errorMessage,
        true,
      );
      throw error;
    }
  };

  const createMaritalStatusChart = ({
    containerId = "marital-status-chart",
    workbookUrl = WORKBOOK_URL,
    sheetName = DEMOGRAPHIC_SHEET,
    purchaseType = null,
    audience = "owner",
  } = {}) =>
    createSingleChart({
      containerId,
      loadingMessage: "正在加载婚姻状态数据…",
      errorMessage:
        "婚姻状态数据加载失败，请检查数据文件后刷新页面。",
      optionFactory: async (container) => {
        const { rows, categories, sampleSizes } =
          await loadCustomerProfileMetric(
            "Marital Status",
            workbookUrl,
            ["单身", "未婚", "已婚有孩子", "已婚没有孩子"],
            { sheetName, purchaseType, audience },
          );
        updateChartLegend(container, sampleSizes);

        const createPieData = (model) =>
          categories.map((category) =>
            createDistributionDatum(
              rows,
              model,
              category,
              "maritalStatus",
            ),
          );

        return {
          animation: false,
          backgroundColor: "#FFFFFF",
          aria: {
            enabled: true,
            decal: { show: false },
          },
          color: AGE_COLORS.slice(2),
          title: MODEL_ORDER.map((model, index) => ({
            text: `${model}  N=${formatSampleSize(sampleSizes[model])}`,
            left: index === 0 ? "25%" : "75%",
            top: 4,
            textAlign: "center",
            textStyle: {
              color: index === 0 ? MODEL_COLORS.F65 : "#397FA8",
              fontSize: 11,
              fontWeight: 700,
            },
          })),
          tooltip: createCompactTooltip(
            "婚姻状态",
            "maritalStatus",
          ),
          legend: {
            bottom: 2,
            left: "center",
            itemWidth: 9,
            itemHeight: 9,
            itemGap: 18,
            textStyle: {
              color: "#536872",
              fontSize: 10,
            },
            data: categories,
          },
          series: MODEL_ORDER.map((model, index) => ({
            name: model,
            type: "pie",
            center: [index === 0 ? "25%" : "75%", "50%"],
            radius: "56%",
            minShowLabelAngle: 4,
            avoidLabelOverlap: true,
            data: createPieData(model),
            itemStyle: {
              borderColor: "#FFFFFF",
              borderWidth: 2,
            },
            label: {
              show: true,
              color: "#40545F",
              fontSize: 10,
              lineHeight: 14,
              formatter: (params) =>
                Number.isFinite(params.data.percentage) &&
                params.data.percentage >= 0.04
                  ? `${params.data.maritalStatus}\n${formatPercentage(params.data.percentage)}`
                  : "",
            },
            labelLine: {
              length: 10,
              length2: 8,
              lineStyle: {
                color: "#AABAC1",
              },
            },
            emphasis: {
              scaleSize: 4,
            },
          })),
        };
      },
    });

  const createEducationChart = ({
    containerId = "education-chart",
    workbookUrl = WORKBOOK_URL,
    sheetName = DEMOGRAPHIC_SHEET,
    purchaseType = null,
    audience = "owner",
  } = {}) =>
    createSingleChart({
      containerId,
      loadingMessage: "正在加载学历数据…",
      errorMessage:
        "学历数据加载失败，请检查数据文件后刷新页面。",
      optionFactory: async (container) => {
        const { rows, categories, sampleSizes } = await loadCustomerProfileMetric(
          "Education",
          workbookUrl,
          ["高中及以下", "专科", "本科", "硕士", "博士"],
          { sheetName, purchaseType, audience },
        );
        updateChartLegend(container, sampleSizes);
        return createStackedDistributionOption({
          rows,
          categories,
          detailLabel: "学历",
          detailKey: "educationLevel",
          stackId: "education-total",
        });
      },
    });

  const createIncomeChart = ({
    containerId = "household-income-chart",
    workbookUrl = WORKBOOK_URL,
    sheetName = DEMOGRAPHIC_SHEET,
    purchaseType = null,
    audience = "owner",
  } = {}) =>
    createSingleChart({
      containerId,
      loadingMessage: "正在加载税后家庭收入数据…",
      errorMessage:
        "税后家庭收入数据加载失败，请检查数据文件后刷新页面。",
      optionFactory: async (container) => {
        const { rows, categories, sampleSizes } = await loadCustomerProfileMetric(
          "Household Income",
          workbookUrl,
          ["20万以下", "20万-30万", "30万-40万", "40万-50万", "50万以上"],
          { sheetName, purchaseType, audience },
        );
        updateChartLegend(container, sampleSizes);
        return createStackedDistributionOption({
          rows,
          categories,
          detailLabel: "收入区间",
          detailKey: "incomeGroup",
          stackId: "income-total",
        });
      },
    });

  const createIndustryOption = (model, rows) => {
    const modelRows = rows
      .filter((row) => row.model === model)
      .sort(
        (rowA, rowB) =>
          (Number.isFinite(rowB.percentage) ? rowB.percentage : -1) -
          (Number.isFinite(rowA.percentage) ? rowA.percentage : -1),
      );
    const chartRows =
      modelRows.length > 0
        ? modelRows
        : [{ category: "N/A", count: null, percentage: null }];

    return {
      animation: false,
      backgroundColor: "#FFFFFF",
      aria: {
        enabled: true,
        decal: { show: false },
      },
      tooltip: createCompactTooltip("行业", "industry"),
      grid: {
        left: 150,
        right: 42,
        top: 10,
        bottom: 18,
        containLabel: false,
      },
      xAxis: {
        type: "value",
        min: 0,
        max: (value) =>
          Math.min(
            1,
            Math.max(0.25, Math.ceil((value.max || 0.2) * 10) / 10),
          ),
        axisLabel: {
          color: "#8A9AA2",
          fontSize: 9,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        axisLine: { show: false },
        axisTick: { show: false },
        splitLine: {
          lineStyle: {
            color: "#E9EEF1",
          },
        },
      },
      yAxis: {
        type: "category",
        inverse: true,
        data: chartRows.map((row) => row.category),
        axisLabel: {
          color: "#536872",
          fontSize: 10,
          width: 138,
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
          barMaxWidth: 18,
          showBackground: true,
          backgroundStyle: {
            color: "#F1F4F5",
            borderRadius: 2,
          },
          itemStyle: {
            color: MODEL_COLORS[model],
            borderRadius: [0, 2, 2, 0],
          },
          emphasis: {
            itemStyle: {
              opacity: 0.84,
            },
          },
          label: {
            show: true,
            position: "right",
            distance: 7,
            color: "#40545F",
            fontSize: 9,
            fontWeight: 700,
            formatter: (params) =>
              formatPercentage(params.data.percentage),
          },
          data: chartRows.map((row) => ({
            value: Number.isFinite(row.percentage) ? row.percentage : 0,
            model,
            industry: row.category,
            count: Number.isFinite(row.count) ? row.count : null,
            percentage: Number.isFinite(row.percentage)
              ? row.percentage
              : null,
          })),
        },
      ],
    };
  };

  const createIndustryPanels = (container, sampleSizes, containerId) => {
    container
      .querySelectorAll(".industry-model-chart__canvas")
      .forEach((canvas) => window.echarts.getInstanceByDom(canvas)?.dispose());
    container.innerHTML = "";
    container.classList.add("chart-canvas--rendered");

    return MODEL_ORDER.map((model) => {
      const panel = document.createElement("section");
      panel.className = "industry-model-chart";
      panel.style.setProperty("--model-color", MODEL_COLORS[model]);

      const header = document.createElement("div");
      header.className = "industry-model-chart__header";
      header.innerHTML = [
        '<i class="industry-model-chart__swatch" aria-hidden="true"></i>',
        `<span>${model}</span>`,
        `<span class="industry-model-chart__sample">N=${formatSampleSize(sampleSizes[model])}</span>`,
      ].join("");

      const canvas = document.createElement("div");
      canvas.className = "industry-model-chart__canvas";
      canvas.id = `${containerId}-${model.toLowerCase()}`;

      panel.append(header, canvas);
      container.append(panel);
      return { model, canvas };
    });
  };

  const createIndustryChart = async ({
    containerId = "industry-chart",
    workbookUrl = WORKBOOK_URL,
    sheetName = DEMOGRAPHIC_SHEET,
    purchaseType = null,
    audience = "owner",
  } = {}) => {
    const container = document.getElementById(containerId);
    if (!container) {
      throw new Error(`Chart container "#${containerId}" was not found.`);
    }

    if (!window.echarts) {
      showChartStatus(
        container,
        "图表组件加载失败，请检查网络连接后刷新页面。",
        true,
      );
      throw new Error("ECharts is unavailable.");
    }

    showChartStatus(container, "正在加载行业数据…");

    try {
      const { rows, sampleSizes } = await loadCustomerProfileMetric(
        "Industry",
        workbookUrl,
        [],
        { sheetName, purchaseType, audience },
      );
      const panels = createIndustryPanels(container, sampleSizes, containerId);
      const charts = panels.map(({ model, canvas }) => {
        const chart = window.echarts.init(canvas, null, {
          renderer: "canvas",
        });
        chart.setOption(createIndustryOption(model, rows), true);
        return chart;
      });

      const resizeObserver = new ResizeObserver(() => {
        charts.forEach((chart) => chart.resize());
      });
      panels.forEach(({ canvas }) => resizeObserver.observe(canvas));
      return charts;
    } catch (error) {
      const isLocalFile = window.location.protocol === "file:";
      showChartStatus(
        container,
        isLocalFile
          ? "Excel 数据需要通过本地服务器加载，请使用 http://localhost 打开看板。"
          : "行业数据加载失败，请检查数据文件后刷新页面。",
        true,
      );
      throw error;
    }
  };

  const getTopCategory = (rows, model) => {
    const match = rows
      .filter((row) => row.model === model)
      .sort(
        (a, b) =>
          (b.percentage ?? -1) - (a.percentage ?? -1) ||
          (b.count ?? -1) - (a.count ?? -1),
      )[0];
    return match?.category ?? "N/A";
  };

  const estimateAverageAge = (rows, model) => {
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
    const modelRows = rows.filter(
      (row) =>
        row.model === model &&
        Number.isFinite(row.count) &&
        Number.isFinite(ageMidpoints[row.category]),
    );
    const total = modelRows.reduce((sum, row) => sum + row.count, 0);
    if (!total) return "N/A";
    const weightedAge = modelRows.reduce(
      (sum, row) => sum + row.count * ageMidpoints[row.category],
      0,
    );
    return (weightedAge / total).toFixed(1);
  };

  const setSnapshotValue = (metric, model, value) => {
    const cell = document.getElementById(
      `profile-snapshot-${metric}-${model.toLowerCase()}`,
    );
    if (cell) cell.textContent = value;
  };

  const hydrateCustomerSnapshot = async ({
    workbookUrl = WORKBOOK_URL,
    audience = "owner",
  } = {}) => {
    if (audience === "owner") {
      MODEL_ORDER.forEach((model) => {
        const snapshot = OWNER_SNAPSHOT[model];
        setSnapshotValue("sample", model, snapshot.sample);
        setSnapshotValue("age", model, snapshot.averageAge);
        setSnapshotValue("female", model, snapshot.femaleShare);
        setSnapshotValue("family", model, snapshot.familyStatus);
        setSnapshotValue("industry", model, snapshot.industry);
      });
      return;
    }

    const [gender, age, maritalStatus, industry] = await Promise.all([
      loadCustomerProfileMetric("Gender", workbookUrl, [], { audience }),
      loadCustomerProfileMetric("Age", workbookUrl, [], { audience }),
      loadCustomerProfileMetric("Marital Status", workbookUrl, [], {
        audience,
      }),
      loadCustomerProfileMetric("Industry", workbookUrl, [], { audience }),
    ]);

    MODEL_ORDER.forEach((model) => {
      const female = gender.rows.find(
        (row) => row.model === model && row.category === "女",
      );
      setSnapshotValue(
        "sample",
        model,
        formatSampleSize(gender.sampleSizes[model]),
      );
      setSnapshotValue("age", model, estimateAverageAge(age.rows, model));
      setSnapshotValue(
        "female",
        model,
        formatPercentage(female?.percentage),
      );
      setSnapshotValue(
        "family",
        model,
        getTopCategory(maritalStatus.rows, model),
      );
      setSnapshotValue(
        "industry",
        model,
        getTopCategory(industry.rows, model),
      );
    });
  };

  window.DashboardCharts = Object.freeze({
    createAgeChart,
    createEducationChart,
    createGenderChart,
    createIncomeChart,
    createIndustryChart,
    createMaritalStatusChart,
    hydrateCustomerSnapshot,
    hydrateChartSampleSizes,
  });
})();
