(() => {
  const WORKBOOK_URL =
    window.DASHBOARD_WORKBOOK_URL ??
    "data/mock_dashboard_data.xlsx?v=20260727-2";
  const BEHAVIOR_SHEET = "购车行为";
  const PROFILE_SHEET = "购车类型画像";
  const FIRST_PURCHASE_AVERAGE_AGE_SHEET = "first_purchase_average_age";
  const FIRST_PURCHASE_GENDER_PROFILE_SHEET = "first_purchase_gender_profile";
  const ADDITIONAL_PURCHASE_AVERAGE_AGE_SHEET = "additional_purchase_average_age";
  const ADDITIONAL_PURCHASE_GENDER_PROFILE_SHEET = "additional_purchase_gender_profile";
  const ADDITIONAL_PURCHASE_VEHICLE_SHEET = "additional_purchase_vehicle_profile";
  const REPLACEMENT_PURCHASE_AVERAGE_AGE_SHEET = "replacement_purchase_average_age";
  const REPLACEMENT_PURCHASE_GENDER_PROFILE_SHEET = "replacement_purchase_gender_profile";
  const REPLACEMENT_PURCHASE_VEHICLE_SHEET = "replacement_purchase_vehicle_profile";
  const MODELS = ["F65", "F66"];
  const MODEL_COLORS = {
    F65: "#0066B3",
    F66: "#259AC5",
  };
  const CATEGORY_COLORS = [
    "#CDE7EF",
    "#9FD3E1",
    "#68BDD3",
    "#2EA5C3",
    "#0788B0",
    "#006D9E",
    "#07527C",
    "#083A59",
  ];
  const charts = new Map();
  let workbookPromise;

  const loadEmbeddedWorkbook = () => {
    const encoded = window.DASHBOARD_WORKBOOK_BASE64;
    if (!encoded) {
      throw new Error("内置数据副本不可用。");
    }
    const binary = window.atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    return window.XLSX.read(bytes, { type: "array" });
  };

  const formatPercentage = (value) => {
    const number = Number(value);
    return Number.isFinite(number) ? `${(number * 100).toFixed(1)}%` : "N/A";
  };

  const escapeHtml = (value) =>
    String(value ?? "N/A")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");

  const loadWorkbook = async () => {
    if (!window.XLSX) {
      throw new Error("Excel 读取组件不可用。");
    }
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

  const loadRows = async (sheetName) => {
    let workbook = await loadWorkbook();
    let worksheet = workbook.Sheets[sheetName];
    if (!worksheet && window.DASHBOARD_WORKBOOK_BASE64) {
      workbook = loadEmbeddedWorkbook();
      workbookPromise = Promise.resolve(workbook);
      worksheet = workbook.Sheets[sheetName];
    }
    if (!worksheet) {
      throw new Error(`未找到工作表：${sheetName}`);
    }
    return window.XLSX.utils.sheet_to_json(worksheet, {
      defval: null,
      raw: true,
    });
  };

  const normalizeRows = (rows) =>
    rows
      .map((row, rowIndex) => ({
        metric: row["指标"] ?? row.Metric,
        category: row["分类"] ?? row.Category,
        model: row["车型"] ?? row.Model,
        count: Number.isFinite(Number(row.N ?? row.Count))
          ? Number(row.N ?? row.Count)
          : null,
        percentage: Number.isFinite(Number(row["占比"] ?? row.Percentage))
          ? Number(row["占比"] ?? row.Percentage)
          : null,
        sortOrder: Number.isFinite(Number(row["排序"] ?? row.Sort_Order))
          ? Number(row["排序"] ?? row.Sort_Order)
          : rowIndex,
        averageAge: Number.isFinite(Number(row.average_age ?? row.Average_Age))
          ? Number(row.average_age ?? row.Average_Age)
          : null,
        purchaseType: row["购车类型"],
      }))
      .filter(
        (row) =>
          MODELS.includes(row.model) &&
          typeof row.category === "string" &&
          row.category.length > 0,
      );

  const loadBehaviorMetric = async (metric) =>
    normalizeRows(await loadRows(BEHAVIOR_SHEET))
      .filter((row) => row.metric === metric)
      .sort((a, b) => a.sortOrder - b.sortOrder);

  const loadProfileMetric = async (purchaseType, metric) =>
    normalizeRows(await loadRows(PROFILE_SHEET))
      .filter(
        (row) =>
          row.purchaseType === purchaseType && row.metric === metric,
      )
      .sort((a, b) => a.sortOrder - b.sortOrder);

  const loadPurchaseProfileAverageAge = async (purchaseType, model) => {
    const sheetName = purchaseType === "增购"
      ? ADDITIONAL_PURCHASE_AVERAGE_AGE_SHEET
      : purchaseType === "换购"
        ? REPLACEMENT_PURCHASE_AVERAGE_AGE_SHEET
        : FIRST_PURCHASE_AVERAGE_AGE_SHEET;
    const rows = await loadRows(sheetName);
    const row = rows.find(
      (item) =>
        (item.model ?? item.Model) === model &&
        (item.series ?? item.Series) === "全部",
    );
    if (!row) return null;
    const averageAge = Number(row.average_age ?? row.Average_Age);
    const sampleSize = Number(row.sample_size ?? row.Sample_Size);
    return {
      averageAge: Number.isFinite(averageAge) ? averageAge : null,
      sampleSize: Number.isFinite(sampleSize) ? sampleSize : null,
    };
  };

  const loadPurchaseGenderProfile = async (purchaseType, model) => {
    const sheetName = purchaseType === "增购"
      ? ADDITIONAL_PURCHASE_GENDER_PROFILE_SHEET
      : purchaseType === "换购"
        ? REPLACEMENT_PURCHASE_GENDER_PROFILE_SHEET
        : FIRST_PURCHASE_GENDER_PROFILE_SHEET;
    const rows = await loadRows(sheetName);
    return rows
      .filter((row) =>
        (row.model ?? row.Model) === model &&
        (row.series ?? row.Series) === "全部"
      )
      .map((row, index) => ({
        metric: "性别",
        category: row.gender ?? row.Gender,
        model,
        count: Number(row.count ?? row.Count),
        percentage: Number(row.percentage ?? row.Percentage),
        averageAge: Number(row.average_age ?? row.Average_Age),
        sortOrder: Number(row.sort_order ?? row.Sort_Order ?? index + 1),
        purchaseType,
      }))
      .filter((row) =>
        ["男", "女"].includes(row.category) &&
        Number.isFinite(row.count) &&
        Number.isFinite(row.percentage)
      )
      .sort((a, b) => a.sortOrder - b.sortOrder);
  };

  const loadPurchaseVehicleProfile = async (purchaseType, model) => {
    const sheetName = purchaseType === "增购"
      ? ADDITIONAL_PURCHASE_VEHICLE_SHEET
      : REPLACEMENT_PURCHASE_VEHICLE_SHEET;
    const rows = await loadRows(sheetName);
    const selected = rows
      .filter((row) =>
        (row.model ?? row.Model) === model &&
        (row.series ?? row.Series) === "全部"
      )
      .sort((a, b) => Number(a.rank ?? a.Rank) - Number(b.rank ?? b.Rank));
    if (!selected.length) return null;
    const sampleSize = Number(selected[0].sample_size ?? selected[0].Sample_Size);
    return {
      sampleSize: Number.isFinite(sampleSize) ? sampleSize : null,
      rows: selected.map((row) => [
        row.brand ?? row.Brand,
        Number(row.mention_rate ?? row.Mention_Rate),
        row.vehicle_series ?? row.Vehicle_Series ?? "N/A",
      ]),
    };
  };

  const compactTooltip = (
    categoryLabel,
    categoryKey = "category",
    showAverageAge = false,
  ) => ({
    trigger: "item",
    confine: true,
    transitionDuration: 0,
    backgroundColor: "rgba(255,255,255,.96)",
    borderColor: "rgba(156,181,192,.82)",
    borderWidth: 1,
    padding: [7, 9],
    textStyle: {
      color: "#203741",
      fontSize: 10,
    },
    extraCssText:
      "box-shadow:0 8px 20px rgba(7,31,45,.14);border-radius:4px;backdrop-filter:blur(4px);",
    position: (point) => [point[0] + 10, point[1] + 10],
    formatter: (params) => {
      const data = params.data ?? {};
      const detailLines = [
        '<div style="line-height:1.55;white-space:nowrap">',
        `<div>车型：<b>${escapeHtml(data.model)}</b></div>`,
        `<div>${escapeHtml(categoryLabel)}：<b>${escapeHtml(data[categoryKey])}</b></div>`,
        `<div>N：<b>${escapeHtml(data.count)}</b></div>`,
        `<div>占比：<b>${escapeHtml(formatPercentage(data.percentage))}</b></div>`,
      ];
      if (showAverageAge) {
        detailLines.push(
          `<div>平均年龄：<b>${Number.isFinite(Number(data.averageAge)) ? Number(data.averageAge).toFixed(1) : "N/A"}</b></div>`,
        );
      }
      detailLines.push("</div>");
      return detailLines.join("");
    },
  });

  const sampleSizesForMetric = (rows) =>
    Object.fromEntries(
      MODELS.map((model) => {
        const total = rows
          .filter((row) => row.model === model)
          .reduce((sum, row) => sum + (row.count ?? 0), 0);
        return [model, total || null];
      }),
    );

  const sampleCaption = (sampleSizes) =>
    MODELS.map((model) => `${model} N=${sampleSizes[model] ?? "N/A"}`).join(
      "    ",
    );

  const categoriesFromRows = (rows) => {
    const categoryOrders = new Map();
    rows.forEach((row) => {
      const current = categoryOrders.get(row.category);
      if (current === undefined || row.sortOrder < current) {
        categoryOrders.set(row.category, row.sortOrder);
      }
    });
    return [...categoryOrders.entries()]
      .sort((a, b) => a[1] - b[1])
      .map(([category]) => category);
  };

  const datum = (rows, model, category, extra = {}) => {
    const row = rows.find(
      (item) => item.model === model && item.category === category,
    );
    return {
      value: row?.percentage ?? 0,
      model,
      category,
      count: row?.count ?? null,
      percentage: row?.percentage ?? null,
      averageAge: row?.averageAge ?? null,
      ...extra,
    };
  };

  const renderChart = (containerId, option) => {
    const container = document.getElementById(containerId);
    if (!container) {
      throw new Error(`未找到图表容器：${containerId}`);
    }
    if (!window.echarts) {
      throw new Error("ECharts 图表组件不可用。");
    }
    let chart = charts.get(containerId);
    if (!chart) {
      container.innerHTML = "";
      chart = window.echarts.init(container, null, { renderer: "canvas" });
    }
    container.classList.add("purchase-chart-canvas--rendered");
    charts.set(containerId, chart);
    chart.setOption(option, true);
    return chart;
  };

  const disposeChart = (containerId) => {
    const chart = charts.get(containerId);
    if (chart) {
      chart.dispose();
      charts.delete(containerId);
    }
  };

  const showError = (containerId, message) => {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = `<span class="chart-status chart-status--error">${escapeHtml(message)}</span>`;
  };

  const baseOption = {
    animation: false,
    backgroundColor: "#FFFFFF",
    aria: {
      enabled: true,
      decal: { show: false },
    },
  };

  const createPurchaseTypePie = async () => {
    const rows = await loadBehaviorMetric("购车类型");
    const categories = categoriesFromRows(rows);
    const sampleSizes = sampleSizesForMetric(rows);
    return renderChart("purchase-type-chart", {
      ...baseOption,
      color: CATEGORY_COLORS.slice(2, 5),
      title: MODELS.map((model, index) => ({
        text: `${model}  N=${sampleSizes[model] ?? "N/A"}`,
        left: index === 0 ? "27%" : "73%",
        top: 8,
        textAlign: "center",
        textStyle: {
          color: index === 0 ? MODEL_COLORS.F65 : "#167CA4",
          fontSize: 11,
          fontWeight: 700,
        },
      })),
      tooltip: compactTooltip("购车类型"),
      legend: {
        bottom: 5,
        left: "center",
        data: categories,
        itemWidth: 9,
        itemHeight: 9,
        itemGap: 18,
        textStyle: { color: "#536872", fontSize: 10 },
      },
      series: MODELS.map((model, index) => ({
        name: model,
        type: "pie",
        center: [index === 0 ? "27%" : "73%", "50%"],
        radius: "45%",
        minShowLabelAngle: 3,
        avoidLabelOverlap: true,
        data: categories.map((category) =>
          datum(rows, model, category, { name: category }),
        ),
        itemStyle: {
          borderColor: "#FFFFFF",
          borderWidth: 2,
        },
        label: {
          color: "#344D5A",
          fontSize: 10,
          lineHeight: 14,
          width: 62,
          overflow: "break",
          alignTo: "none",
          formatter: (params) =>
            `${params.data.category}\n${formatPercentage(params.data.percentage)}`,
        },
        labelLine: {
          length: 8,
          length2: 5,
          lineStyle: { color: "#AABAC1" },
        },
      })),
    });
  };

  const createFemaleUserChart = async () => {
    const rows = await loadBehaviorMetric("实际用车人");
    const categories = categoriesFromRows(rows);
    const sampleSizes = sampleSizesForMetric(rows);
    return renderChart("female-user-chart", {
      ...baseOption,
      color: [CATEGORY_COLORS[4], CATEGORY_COLORS[1]],
      title: {
        text: sampleCaption(sampleSizes),
        right: 12,
        top: 4,
        textStyle: {
          color: "#647B87",
          fontSize: 10,
          fontWeight: 500,
        },
      },
      legend: {
        top: 7,
        left: 12,
        data: categories,
        itemWidth: 9,
        itemHeight: 9,
        textStyle: { color: "#536872", fontSize: 10 },
      },
      tooltip: compactTooltip("实际用车人"),
      grid: {
        left: 48,
        right: 18,
        top: 48,
        bottom: 35,
      },
      xAxis: {
        type: "category",
        data: MODELS,
        axisLabel: {
          color: "#40545F",
          fontSize: 11,
          fontWeight: 700,
        },
        axisTick: { show: false },
        axisLine: { lineStyle: { color: "#C8D7DE" } },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: 1,
        interval: 0.2,
        axisLabel: {
          color: "#82949D",
          fontSize: 9,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        splitLine: { lineStyle: { color: "#E3EAED" } },
      },
      series: categories.map((category, index) => ({
        name: category,
        type: "bar",
        stack: "user-total",
        barWidth: "45%",
        data: MODELS.map((model) => datum(rows, model, category)),
        itemStyle: {
          color: [CATEGORY_COLORS[4], CATEGORY_COLORS[1]][index],
          borderColor: "#FFFFFF",
          borderWidth: 0.5,
        },
        label: {
          show: true,
          position: "inside",
          color: "#FFFFFF",
          fontSize: 10,
          fontWeight: 700,
          textBorderColor: "rgba(7,31,45,.22)",
          textBorderWidth: 2,
          formatter: (params) => formatPercentage(params.data.percentage),
        },
      })),
    });
  };

  const stackedProfileOption = (rows, categoryLabel) => {
    const categories = categoriesFromRows(rows);
    const sampleSizes = sampleSizesForMetric(rows);
    return {
      ...baseOption,
      color: CATEGORY_COLORS,
      title: {
        text: sampleCaption(sampleSizes),
        right: 8,
        top: 1,
        textStyle: { color: "#647B87", fontSize: 9, fontWeight: 500 },
      },
      legend: {
        type: "scroll",
        top: 1,
        left: 8,
        right: 145,
        data: categories,
        itemWidth: 8,
        itemHeight: 8,
        itemGap: 8,
        textStyle: { color: "#536872", fontSize: 8 },
      },
      tooltip: compactTooltip(categoryLabel),
      grid: { left: 40, right: 10, top: 42, bottom: 28 },
      xAxis: {
        type: "category",
        data: MODELS,
        axisLabel: {
          color: "#40545F",
          fontSize: 10,
          fontWeight: 700,
        },
        axisTick: { show: false },
        axisLine: { lineStyle: { color: "#C8D7DE" } },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: 1,
        interval: 0.25,
        axisLabel: {
          color: "#82949D",
          fontSize: 8,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        splitLine: { lineStyle: { color: "#E3EAED" } },
      },
      series: categories.map((category, index) => ({
        name: category,
        type: "bar",
        stack: "profile-total",
        barWidth: "44%",
        data: MODELS.map((model) => datum(rows, model, category)),
        itemStyle: {
          color: CATEGORY_COLORS[index % CATEGORY_COLORS.length],
          borderColor: "#FFFFFF",
          borderWidth: 0.5,
        },
        label: {
          show: true,
          position: "inside",
          color: "#FFFFFF",
          fontSize: 8,
          fontWeight: 700,
          textBorderColor: "rgba(7,31,45,.22)",
          textBorderWidth: 2,
          formatter: (params) =>
            (params.data.percentage ?? 0) >= 0.075
              ? formatPercentage(params.data.percentage)
              : "",
        },
      })),
    };
  };

  const groupedProfileOption = (rows, categoryLabel, horizontal = false) => {
    const categories = categoriesFromRows(rows);
    const sampleSizes = sampleSizesForMetric(rows);
    const common = {
      ...baseOption,
      color: [MODEL_COLORS.F65, MODEL_COLORS.F66],
      title: {
        text: sampleCaption(sampleSizes),
        right: 8,
        top: 1,
        textStyle: { color: "#647B87", fontSize: 9, fontWeight: 500 },
      },
      legend: {
        top: 1,
        left: 8,
        data: MODELS,
        itemWidth: 8,
        itemHeight: 8,
        textStyle: { color: "#536872", fontSize: 8 },
      },
      tooltip: compactTooltip(categoryLabel),
      series: MODELS.map((model) => ({
        name: model,
        type: "bar",
        barMaxWidth: 16,
        data: categories.map((category) => datum(rows, model, category)),
        itemStyle: {
          color: MODEL_COLORS[model],
          borderRadius: horizontal ? [0, 2, 2, 0] : [2, 2, 0, 0],
        },
        label: {
          show: true,
          position: horizontal ? "right" : "insideTop",
          color: horizontal ? "#344D5A" : "#FFFFFF",
          fontSize: 8,
          fontWeight: 700,
          formatter: (params) => formatPercentage(params.data.percentage),
        },
      })),
    };

    if (horizontal) {
      return {
        ...common,
        grid: { left: 145, right: 48, top: 35, bottom: 24 },
        xAxis: {
          type: "value",
          min: 0,
          axisLabel: {
            color: "#82949D",
            fontSize: 8,
            formatter: (value) => `${Math.round(value * 100)}%`,
          },
          splitLine: { lineStyle: { color: "#E3EAED" } },
        },
        yAxis: {
          type: "category",
          inverse: true,
          data: categories,
          axisLabel: {
            color: "#536872",
            fontSize: 9,
            width: 130,
            overflow: "truncate",
          },
          axisLine: { show: false },
          axisTick: { show: false },
        },
      };
    }

    return {
      ...common,
      grid: { left: 42, right: 12, top: 35, bottom: 30 },
      xAxis: {
        type: "category",
        data: categories,
        axisLabel: { color: "#536872", fontSize: 9 },
        axisTick: { show: false },
        axisLine: { lineStyle: { color: "#C8D7DE" } },
      },
      yAxis: {
        type: "value",
        min: 0,
        axisLabel: {
          color: "#82949D",
          fontSize: 8,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        splitLine: { lineStyle: { color: "#E3EAED" } },
      },
    };
  };

  const maritalProfileOption = (rows) => {
    const categories = categoriesFromRows(rows);
    const sampleSizes = sampleSizesForMetric(rows);
    return {
      ...baseOption,
      color: CATEGORY_COLORS.slice(2, 6),
      title: MODELS.map((model, index) => ({
        text: `${model} N=${sampleSizes[model] ?? "N/A"}`,
        left: index === 0 ? "25%" : "75%",
        top: 1,
        textAlign: "center",
        textStyle: {
          color: index === 0 ? MODEL_COLORS.F65 : "#167CA4",
          fontSize: 9,
          fontWeight: 700,
        },
      })),
      tooltip: compactTooltip("婚姻状态"),
      legend: {
        bottom: 0,
        left: "center",
        data: categories,
        itemWidth: 8,
        itemHeight: 8,
        itemGap: 10,
        textStyle: { color: "#536872", fontSize: 8 },
      },
      series: MODELS.map((model, index) => ({
        name: model,
        type: "pie",
        center: [index === 0 ? "25%" : "75%", "48%"],
        radius: "47%",
        data: categories.map((category) =>
          datum(rows, model, category, { name: category }),
        ),
        itemStyle: { borderColor: "#FFFFFF", borderWidth: 1.5 },
        label: {
          color: "#344D5A",
          fontSize: 8,
          formatter: (params) =>
            (params.data.percentage ?? 0) >= 0.1
              ? `${params.data.category}\n${formatPercentage(params.data.percentage)}`
              : "",
        },
        labelLine: {
          length: 7,
          length2: 5,
          lineStyle: { color: "#AABAC1" },
        },
      })),
    };
  };

  const AGE_MIDPOINTS = {
    "20岁以下": 19,
    "20-24岁": 22,
    "25-29岁": 27,
    "30-34岁": 32,
    "35-39岁": 37,
    "40-44岁": 42,
    "45-49岁": 47,
    "50岁及以上": 52,
  };

  const updateText = (id, value) => {
    const element = document.getElementById(id);
    if (element) element.textContent = value ?? "N/A";
  };

  const metricSampleSize = (rows, model) =>
    rows
      .filter((row) => row.model === model && Number.isFinite(row.count))
      .reduce((sum, row) => sum + row.count, 0);

  const sampleText = (sampleSize) =>
    `N=${sampleSize > 0 ? sampleSize : "N/A"}`;

  const selectedModelRows = (rows, model) =>
    rows.filter((row) => row.model === model);

  const singleModelStackedOption = (rows, model, categoryLabel) => {
    const categories = categoriesFromRows(rows);
    return {
      ...baseOption,
      color: CATEGORY_COLORS,
      legend: {
        type: "scroll",
        top: 4,
        left: 8,
        right: 8,
        data: categories,
        itemWidth: 8,
        itemHeight: 8,
        itemGap: 10,
        textStyle: { color: "#536872", fontSize: 9 },
      },
      tooltip: compactTooltip(categoryLabel),
      grid: { left: 48, right: 28, top: 48, bottom: 32 },
      xAxis: {
        type: "category",
        data: [model],
        axisLabel: {
          color: "#40545F",
          fontSize: 11,
          fontWeight: 700,
        },
        axisTick: { show: false },
        axisLine: { lineStyle: { color: "#C8D7DE" } },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: 1,
        interval: 0.2,
        axisLabel: {
          color: "#82949D",
          fontSize: 9,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        splitLine: { lineStyle: { color: "#E3EAED" } },
      },
      series: categories.map((category, index) => ({
        name: category,
        type: "bar",
        stack: "profile-total",
        barWidth: "34%",
        data: [datum(rows, model, category)],
        itemStyle: {
          color: CATEGORY_COLORS[index % CATEGORY_COLORS.length],
          borderColor: "#FFFFFF",
          borderWidth: 0.75,
        },
        label: {
          show: true,
          position: "inside",
          color: "#FFFFFF",
          fontSize: 10,
          fontWeight: 700,
          textBorderColor: "rgba(7,31,45,.22)",
          textBorderWidth: 2,
          formatter: (params) =>
            (params.data.percentage ?? 0) >= 0.075
              ? formatPercentage(params.data.percentage)
              : "",
        },
      })),
    };
  };

  const genderDrilldownOption = (rows, model) => {
    const categories = categoriesFromRows(rows);
    return {
      ...baseOption,
      tooltip: compactTooltip("性别", "category", true),
      grid: { left: 54, right: 30, top: 18, bottom: 28 },
      xAxis: {
        type: "value",
        min: 0,
        max: 1,
        axisLabel: {
          color: "#82949D",
          fontSize: 9,
          formatter: (value) => `${Math.round(value * 100)}%`,
        },
        splitLine: { lineStyle: { color: "#E3EAED" } },
      },
      yAxis: {
        type: "category",
        inverse: true,
        data: categories,
        axisLabel: { color: "#536872", fontSize: 11 },
        axisLine: { show: false },
        axisTick: { show: false },
      },
      series: [{
        type: "bar",
        barWidth: 24,
        showBackground: true,
        backgroundStyle: { color: "#EEF2F4" },
        data: categories.map((category) => datum(rows, model, category)),
        itemStyle: {
          color: MODEL_COLORS[model],
          borderRadius: [0, 3, 3, 0],
        },
        label: {
          show: true,
          position: "insideRight",
          color: "#FFFFFF",
          fontSize: 10,
          fontWeight: 700,
          formatter: (params) => formatPercentage(params.data.percentage),
        },
      }],
    };
  };

  const renderVehicleProfileTable = (purchaseType, model, dynamicProfile = null) => {
    disposeChart("purchase-profile-detail-content");
    const container = document.getElementById("purchase-profile-detail-content");
    const profile = dynamicProfile;
    if (!container) return;

    container.className = "purchase-profile-detail-content purchase-profile-table-wrap";
    if (!profile) {
      container.innerHTML = '<p class="purchase-profile-empty">暂无有效数据</p>';
      updateText("purchase-profile-detail-n", "N=N/A");
      return;
    }

    updateText("purchase-profile-detail-n", sampleText(profile.sampleSize));
    container.innerHTML = `
      <table class="purchase-profile-vehicle-table">
        <thead>
          <tr>
            <th scope="col">品牌</th>
            <th scope="col">提及率</th>
            <th scope="col">车系</th>
          </tr>
        </thead>
        <tbody>
          ${profile.rows.map(([brand, mentionRate, series]) => `
            <tr tabindex="0" title="${escapeHtml(`${brand}｜提及率 ${formatPercentage(mentionRate)}｜${series}`)}">
              <th scope="row">${escapeHtml(brand)}</th>
              <td>${formatPercentage(mentionRate)}</td>
              <td>${escapeHtml(series)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
      <p class="purchase-profile-table-note">样本数量较少，仅供参考</p>
    `;
  };

  const createPurchaseProfileDrilldown = async (
    purchaseType = document.getElementById("purchase-type-filter")?.value ??
      "首购",
    model = document.getElementById("purchase-model-filter")?.value ?? "F65",
  ) => {
    if (!["首购", "增购", "换购"].includes(purchaseType)) {
      purchaseType = "首购";
    }
    if (!MODELS.includes(model)) {
      model = "F65";
    }

    const [allGenderRows, allAgeRows, detailedGenderRows] = await Promise.all([
      loadProfileMetric(purchaseType, "性别"),
      loadProfileMetric(purchaseType, "年龄"),
      loadPurchaseGenderProfile(purchaseType, model),
    ]);
    let allIncomeRows = [];
    let averageAgeSummary = null;
    let vehicleProfile = null;
    if (purchaseType === "首购") {
      [allIncomeRows, averageAgeSummary] = await Promise.all([
        loadProfileMetric(purchaseType, "税后家庭收入"),
        loadPurchaseProfileAverageAge(purchaseType, model),
      ]);
    } else {
      [averageAgeSummary, vehicleProfile] = await Promise.all([
        loadPurchaseProfileAverageAge(purchaseType, model),
        loadPurchaseVehicleProfile(purchaseType, model),
      ]);
    }
    const genderRows = detailedGenderRows.length
      ? detailedGenderRows
      : selectedModelRows(allGenderRows, model);
    const ageRows = selectedModelRows(allAgeRows, model);
    const incomeRows = selectedModelRows(allIncomeRows, model);
    const genderSample = metricSampleSize(genderRows, model);
    const ageSample = metricSampleSize(ageRows, model);
    const weightedAge = ageRows.reduce(
      (sum, row) =>
        sum + (AGE_MIDPOINTS[row.category] ?? 0) * (row.count ?? 0),
      0,
    );

    updateText("purchase-profile-selected-title", `${purchaseType}用户画像`);
    updateText("purchase-profile-selected-model", model);
    updateText("purchase-profile-gender-n", sampleText(genderSample));
    const averageAgeSample = averageAgeSummary?.sampleSize ?? ageSample;
    const averageAgeValue = averageAgeSummary?.averageAge ??
      (ageSample ? weightedAge / ageSample : null);
    updateText("purchase-profile-average-age-n", sampleText(averageAgeSample));
    updateText("purchase-profile-age-n", sampleText(ageSample));
    updateText(
      "purchase-profile-average-age",
      Number.isFinite(averageAgeValue) ? averageAgeValue.toFixed(1) : "N/A",
    );

    renderChart(
      "purchase-profile-gender",
      genderDrilldownOption(genderRows, model),
    );
    renderChart(
      "purchase-profile-age",
      singleModelStackedOption(ageRows, model, "年龄段"),
    );

    if (purchaseType === "首购") {
      const detailContainer = document.getElementById(
        "purchase-profile-detail-content",
      );
      updateText("purchase-profile-detail-title", "家庭年收入分布");
      updateText(
        "purchase-profile-detail-n",
        sampleText(metricSampleSize(incomeRows, model)),
      );
      if (detailContainer) {
        detailContainer.className =
          "purchase-profile-detail-content purchase-profile-detail-content--chart chart-canvas chart-canvas--wide purchase-profile-drilldown-canvas";
      }
      renderChart(
        "purchase-profile-detail-content",
        singleModelStackedOption(incomeRows, model, "家庭年收入"),
      );
    } else {
      updateText(
        "purchase-profile-detail-title",
        purchaseType === "增购"
          ? "家庭原有车辆的品牌及车系分布"
          : "被置换车辆的品牌及车系分布",
      );
      renderVehicleProfileTable(purchaseType, model, vehicleProfile);
    }

    requestAnimationFrame(resizeAll);
  };

  const resizeAll = () => {
    charts.forEach((chart) => chart.resize());
  };

  window.addEventListener("resize", resizeAll);

  window.PurchaseCharts = Object.freeze({
    createFemaleUserChart,
    createPurchaseProfileDrilldown,
    createPurchaseTypePie,
    resizeAll,
    showError,
  });
})();
