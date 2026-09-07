(() => {
  const MODELS = ["F65", "F66"];
  const MODEL_COLORS = { F65: "#0066B3", F66: "#259AC5" };
  const charts = new Map();
  const VOICE_CATEGORIES = ["空间不足", "价格因素", "体验因素", "内饰因素"];
  const PRICE_REASON_DATA = {
    F65: {
      sampleSize: 54,
      rows: [
        { category: "价格高于期望值", count: 35, percentage: 35 / 54 },
        { category: "养护成本高", count: 10, percentage: 10 / 54 },
        { category: "金融贷款利率不满意", count: 5, percentage: 5 / 54 },
        { category: "购车礼遇不丰富", count: 4, percentage: 4 / 54 },
      ],
    },
    F66: {
      sampleSize: 60,
      rows: [
        { category: "价格高于期望值", count: 38, percentage: 38 / 60 },
        { category: "养护成本高", count: 12, percentage: 12 / 60 },
        { category: "金融贷款利率不满意", count: 5, percentage: 5 / 60 },
        { category: "购车礼遇不丰富", count: 5, percentage: 5 / 60 },
      ],
    },
  };
  const INTERIOR_VOICE_ROWS = [
    {
      model: "F65",
      category: "内饰因素",
      quote: "内饰设计很有 MINI 的个性，但部分材质和细节还可以更精致一些。",
      highlightType: "vehicle",
    },
    {
      model: "F65",
      category: "内饰因素",
      quote: "中控造型很有辨识度，不过储物空间和日常操作便利性没有达到预期。",
      highlightType: "experience",
    },
    {
      model: "F66",
      category: "内饰因素",
      quote: "座舱氛围比上一代更现代，但希望实体操作和屏幕交互可以更直观。",
      highlightType: "experience",
    },
    {
      model: "F66",
      category: "内饰因素",
      quote: "圆形中控屏很有特色，如果内饰配色和材质选择更丰富会更吸引我。",
      highlightType: "vehicle",
    },
  ];
  const WORKBOOK_URL = "data/mock_dashboard_data.xlsx?v=20260728-2";
  let dataPromise;

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
    if (window.location.protocol === "file:") {
      return loadEmbeddedWorkbook();
    }
    try {
      const response = await fetch(WORKBOOK_URL);
      if (!response.ok) throw new Error(`数据文件加载失败：${response.status}`);
      return window.XLSX.read(await response.arrayBuffer());
    } catch (error) {
      if (window.DASHBOARD_WORKBOOK_BASE64) return loadEmbeddedWorkbook();
      throw error;
    }
  };

  const getSheetRows = (workbook, sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`未找到工作表：${sheetName}`);
    return window.XLSX.utils.sheet_to_json(worksheet, {
      defval: null,
      raw: true,
    });
  };

  const groupByModel = (rows, mapper) =>
    Object.fromEntries(
      MODELS.map((model) => [
        model,
        rows.filter((row) => row.model === model).map(mapper),
      ]),
    );

  const getData = async () => {
    if (!dataPromise) {
      dataPromise = loadWorkbook()
        .then((workbook) => {
          const reasonRows = getSheetRows(workbook, "lost_reason");
          const vehicleRows = getSheetRows(workbook, "lost_vehicle");
          const factorRows = getSheetRows(workbook, "lost_factor");
          const voiceRows = getSheetRows(workbook, "customer_voice");
          const sampleSizes = Object.fromEntries(
            MODELS.map((model) => [
              model,
              Number(
                vehicleRows.find((row) => row.model === model)?.sample_size,
              ),
            ]),
          );
          return {
            userType: "Leads",
            sampleSizes,
            priceReasonSampleSizes: Object.fromEntries(
              MODELS.map((model) => [
                model,
                PRICE_REASON_DATA[model].sampleSize,
              ]),
            ),
            priceReason: Object.fromEntries(
              MODELS.map((model) => [
                model,
                PRICE_REASON_DATA[model].rows,
              ]),
            ),
            lostReason: groupByModel(reasonRows, (row) => ({
              category: row.reason,
              count: Number(row.count),
              percentage: Number(row.percentage),
            })),
            lostVehicle: groupByModel(vehicleRows, (row) => ({
              rank: Number(row.rank),
              brand: row.brand,
              mentionRate: Number(row.mention_rate),
              vehicleSeries: row.vehicle_series,
            })),
            lostFactor: groupByModel(factorRows, (row) => ({
              category: row.factor,
              percentage: Number(row.percentage),
              count: Number(row.count),
            })),
            customerVoice: [
              ...voiceRows.map((row) => ({
                model: row.model,
                category: row.category,
                quote: row.quote,
                highlightType: row.highlight_type,
              })),
              ...INTERIOR_VOICE_ROWS,
            ],
          };
        })
        .catch((error) => {
          throw error;
        });
    }
    return dataPromise;
  };

  const modelHeader = (model, sampleSize, subtitle) => `
    <header class="loss-card__header">
      <div>
        <i style="--model-color:${MODEL_COLORS[model]}"></i>
        <strong>${model} Leads</strong>
        <small>${escapeHtml(subtitle)}</small>
      </div>
      <span>N=${sampleSize}</span>
    </header>`;

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

  const createBarCharts = async ({
    containerId,
    dataKey,
    detailLabel,
    chartKey,
    sampleSizeKey = "sampleSizes",
  }) => {
    const container = document.getElementById(containerId);
    if (!container) return [];
    const data = await getData();
    disposeGroup(chartKey);
    container.innerHTML = "";

    return MODELS.map((model) => {
      const rows = [...data[dataKey][model]].sort(
        (a, b) => b.percentage - a.percentage,
      );
      const card = document.createElement("article");
      card.className = "card loss-chart-card";
      card.innerHTML = `
        ${modelHeader(model, data[sampleSizeKey][model], detailLabel)}
        <div class="loss-chart-canvas"></div>`;
      container.append(card);

      const canvas = card.querySelector(".loss-chart-canvas");
      const chart = window.echarts.init(canvas, null, {
        renderer: "canvas",
      });
      const maxValue = Math.max(...rows.map((row) => row.percentage));
      chart.setOption({
        animation: true,
        animationDuration: 360,
        backgroundColor: "#FFFFFF",
        tooltip: compactTooltip(detailLabel),
        grid: {
          left: 148,
          right: 52,
          top: 22,
          bottom: 28,
          containLabel: false,
        },
        xAxis: {
          type: "value",
          min: 0,
          max: Math.max(0.3, Math.ceil((maxValue + 0.06) * 20) / 20),
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
      });
      charts.set(`${chartKey}-${model}`, chart);
      return chart;
    });
  };

  const renderLostVehicleTables = async () => {
    const container = document.getElementById("lost-vehicle-tables");
    if (!container) return;
    const data = await getData();
    container.innerHTML = MODELS.map((model) => `
      <article class="card loss-table-card">
        ${modelHeader(model, data.sampleSizes[model], "Lost Vehicle Top 10")}
        <div class="loss-table">
          <div class="loss-table__head">
            <span>排名</span><span>品牌</span><span>提及率</span><span>主要车系</span>
          </div>
          ${data.lostVehicle[model]
            .map(
              (row) => `
                <div class="loss-table__row" tabindex="0">
                  <span class="loss-rank">${String(row.rank).padStart(2, "0")}</span>
                  <strong>${escapeHtml(row.brand)}</strong>
                  <b>${formatPercentage(row.mentionRate)}</b>
                  <small>${escapeHtml(row.vehicleSeries)}</small>
                  <span class="loss-row-tooltip">
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

  const renderCustomerVoice = async (category = VOICE_CATEGORIES[0]) => {
    const data = await getData();
    const tabs = document.getElementById("loss-voice-tabs");
    const columns = document.getElementById("loss-voice-columns");
    if (!tabs || !columns) return;

    tabs.innerHTML = VOICE_CATEGORIES.map(
      (item) =>
        `<button type="button" class="loss-voice-tab${item === category ? " loss-voice-tab--active" : ""}" data-loss-category="${escapeHtml(item)}" aria-pressed="${item === category}">${escapeHtml(item)}</button>`,
    ).join("");

    const typeLabels = {
      vehicle: "车辆 / 产品相关",
      experience: "实际体验反馈",
      service: "新增客户反馈",
    };
    columns.innerHTML = MODELS.map((model) => `
      <section class="loss-voice-column">
        <h3><i style="--model-color:${MODEL_COLORS[model]}"></i>${model} Leads</h3>
        <div class="loss-quote-list">
          ${data.customerVoice
            .filter(
              (row) => row.model === model && row.category === category,
            )
            .map(
              (row) => `
                <blockquote class="loss-quote loss-quote--${row.highlightType}" title="${escapeHtml(row.quote)}">
                  <span>${escapeHtml(typeLabels[row.highlightType])}</span>
                  <p>${escapeHtml(row.quote)}</p>
                </blockquote>`,
            )
            .join("")}
        </div>
      </section>`).join("");

    tabs.querySelectorAll("[data-loss-category]").forEach((button) => {
      button.addEventListener("click", () =>
        renderCustomerVoice(button.dataset.lossCategory),
      );
    });
  };

  const initialize = async () => {
    await Promise.all([
      createBarCharts({
      containerId: "lost-reason-charts",
      dataKey: "lostReason",
      detailLabel: "未购车原因",
      chartKey: "lost-reason",
      }),
      renderLostVehicleTables(),
      createBarCharts({
        containerId: "lost-factor-charts",
        dataKey: "lostFactor",
        detailLabel: "战败因素",
        chartKey: "lost-factor",
      }),
      createBarCharts({
        containerId: "lost-price-reason-charts",
        dataKey: "priceReason",
        detailLabel: "价格不合适原因",
        chartKey: "lost-price-reason",
        sampleSizeKey: "priceReasonSampleSizes",
      }),
      renderCustomerVoice(),
    ]);
  };

  const resizeAll = () => charts.forEach((chart) => chart.resize());

  window.LeadsLoss = Object.freeze({
    createBarCharts,
    renderLostVehicleTables,
    renderCustomerVoice,
    initialize,
    resizeAll,
  });
})();
