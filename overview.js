(() => {
  const WORKBOOK_URL = "data/mock_dashboard_data.xlsx?v=20260727-4";
  const SAMPLE_SHEET = "sample_overview";
  const TEST_DRIVE_SHEET = "test_drive";
  const BRAND_SHEET = "brand_image";
  const GROUP_ORDER = ["F65车主", "F66车主", "F65潜客", "F66潜客"];
  const STATUS_ORDER = ["已试驾", "已提供，未试驾", "未提供", "未到店"];
  const STATUS_COLORS = ["#0066B3", "#42A6C4", "#9BCEDB", "#D9E6EA"];
  const MODEL_COLORS = { F65: "#0066B3", F66: "#259AC5" };
  let workbookPromise;
  const charts = new Map();

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

  const loadRows = async (sheetName) => {
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

  const OverviewInsight = () =>
    document.querySelector(".home-insight");

  const SampleOverviewTable = async () => {
    const body = document.getElementById("sample-overview-body");
    if (!body) return;
    try {
      const rows = (await loadRows(SAMPLE_SHEET))
        .map((row) => ({
          group: row["人群分组"],
          model: row["车型"],
          userType: row["用户类型"],
          sampleSize: Number(row["样本量"]),
        }))
        .filter((row) => GROUP_ORDER.includes(row.group))
        .sort(
          (a, b) =>
            GROUP_ORDER.indexOf(a.group) - GROUP_ORDER.indexOf(b.group),
        );
      const coverage = '<span class="coverage-check" aria-label="已覆盖">✓</span>';
      body.innerHTML = rows
        .map(
          (row) => `
            <tr data-user-type="${escapeHtml(row.userType)}">
              <td>${escapeHtml(row.group)}</td>
              <td>${Number.isFinite(row.sampleSize) ? row.sampleSize.toLocaleString("en-US") : "N/A"}</td>
              <td>${coverage}</td>
              <td>${coverage}</td>
              <td>${coverage}</td>
              <td>${coverage}</td>
              <td>${coverage}</td>
            </tr>`,
        )
        .join("");
    } catch (error) {
      body.innerHTML =
        '<tr><td colspan="7">样本概览数据加载失败</td></tr>';
      throw error;
    }
  };

  const testDriveTooltip = {
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
        `<div>用户类型：<b>${escapeHtml(data.userType)}</b></div>`,
        `<div>试驾状态：<b>${escapeHtml(data.status)}</b></div>`,
        `<div>N：<b>${escapeHtml(data.count)}</b></div>`,
        `<div>占比：<b>${escapeHtml(formatPercentage(data.percentage))}</b></div>`,
        "</div>",
      ].join(""),
  };

  const TestDrivePieChart = async () => {
    const grid = document.getElementById("test-drive-chart-grid");
    if (!grid) return [];
    try {
      const rows = (await loadRows(TEST_DRIVE_SHEET))
        .map((row) => ({
          model: row["车型"],
          userType: row["用户类型"],
          status: row["状态"],
          count: Number(row.N),
          percentage: Number(row["占比"]),
        }))
        .filter(
          (row) =>
            ["F65", "F66"].includes(row.model) &&
            ["车主", "潜客"].includes(row.userType) &&
            STATUS_ORDER.includes(row.status),
        );

      [...charts.entries()]
        .filter(([key]) => key.startsWith("test-drive-"))
        .forEach(([key, chart]) => {
          chart.dispose();
          charts.delete(key);
        });
      grid.innerHTML = "";

      const rendered = GROUP_ORDER.map((group) => {
        const model = group.slice(0, 3);
        const userType = group.slice(3);
        const groupRows = rows.filter(
          (row) => row.model === model && row.userType === userType,
        );
        const sampleSize = groupRows.reduce(
          (sum, row) => sum + (Number.isFinite(row.count) ? row.count : 0),
          0,
        );

        const panel = document.createElement("section");
        panel.className = "test-drive-panel";
        panel.innerHTML = `
          <div class="test-drive-panel__header">
            <strong>${escapeHtml(group)}</strong>
            <span>N=${sampleSize || "N/A"}</span>
          </div>
          <div class="test-drive-panel__chart"></div>
        `;
        grid.append(panel);
        const canvas = panel.querySelector(".test-drive-panel__chart");
        const chart = window.echarts.init(canvas, null, {
          renderer: "canvas",
        });
        chart.setOption(
          {
            animation: false,
            backgroundColor: "#FFFFFF",
            color: STATUS_COLORS,
            tooltip: testDriveTooltip,
            legend: {
              bottom: 4,
              left: "center",
              itemWidth: 8,
              itemHeight: 8,
              itemGap: 12,
              textStyle: { color: "#627985", fontSize: 9 },
              data: STATUS_ORDER,
            },
            series: [
              {
                type: "pie",
                center: ["50%", 124],
                radius: [47, 74],
                minShowLabelAngle: 0,
                avoidLabelOverlap: false,
                data: STATUS_ORDER.map((status) => {
                  const row = groupRows.find(
                    (item) => item.status === status,
                  );
                  return {
                    name: status,
                    value: row?.percentage ?? 0,
                    model,
                    userType,
                    status,
                    count: row?.count ?? null,
                    percentage: row?.percentage ?? null,
                  };
                }),
                itemStyle: { borderColor: "#FFFFFF", borderWidth: 2 },
                label: {
                  show: true,
                  color: "#405660",
                  fontSize: 9,
                  formatter: ({ data }) =>
                    `${data.status}\n${formatPercentage(data.percentage)}`,
                },
                labelLayout: { hideOverlap: false },
                labelLine: {
                  length: 8,
                  length2: 5,
                  lineStyle: { color: "#A5B8C1" },
                },
                emphasis: { scaleSize: 4 },
              },
            ],
          },
          true,
        );
        charts.set(`test-drive-${group}`, chart);
        return chart;
      });
      return rendered;
    } catch (error) {
      grid.innerHTML =
        '<span class="chart-status chart-status--error">试驾数据加载失败</span>';
      throw error;
    }
  };

  const BrandImageScoreCard = async () => {
    const container = document.getElementById("brand-image-score-list");
    if (!container) return;
    try {
      const rows = (await loadRows(BRAND_SHEET))
        .map((row) => ({
          model: row["车型"],
          userType: row["用户类型"],
          dimension: row["品牌形象"],
          score: Number(row["评分"]),
        }))
        .filter(
          (row) =>
            ["F65", "F66"].includes(row.model) &&
            ["车主", "潜客"].includes(row.userType) &&
            typeof row.dimension === "string" &&
            Number.isFinite(row.score),
        );
      const sampleRows = (await loadRows(SAMPLE_SHEET)).map((row) => ({
        model: row["车型"],
        userType: row["用户类型"],
        sampleSize: Number(row["样本量"]),
      }));
      ["F65", "F66"].forEach((model) => {
        const existing = charts.get(`brand-image-${model}`);
        if (existing) existing.dispose();
      });
      container.innerHTML = "";
      const dimensions = [...new Set(rows.map((row) => row.dimension))];
      const modelSettings = {
        F65: { owner: "#0066B3", prospect: "#82C4DD" },
        F66: { owner: "#0B789F", prospect: "#8AC9DC" },
      };
      const rendered = ["F65", "F66"].map((model) => {
        const ownerSample = sampleRows.find(
          (row) => row.model === model && row.userType === "车主",
        )?.sampleSize;
        const prospectSample = sampleRows.find(
          (row) => row.model === model && row.userType === "潜客",
        )?.sampleSize;
        const panel = document.createElement("section");
        panel.className = "brand-image-panel";
        panel.innerHTML = `
          <div class="brand-image-panel__header">
            <span class="brand-image-panel__dot" style="--model-color:${MODEL_COLORS[model]}"></span>
            <strong>${model}</strong>
            <small>车主与潜客对比</small>
            <span class="brand-image-panel__sample" title="车主 N=${ownerSample ?? "N/A"}；潜客 N=${prospectSample ?? "N/A"}">N=${ownerSample ?? "N/A"} / ${prospectSample ?? "N/A"}</span>
          </div>
          <div class="brand-image-panel__chart"></div>`;
        container.append(panel);
        const canvas = panel.querySelector(".brand-image-panel__chart");
        const groups = [
          {
            name: "车主",
            userType: "车主",
            color: modelSettings[model].owner,
          },
          {
            name: "潜客",
            userType: "潜客",
            color: modelSettings[model].prospect,
          },
        ];
        const chart = window.echarts.init(canvas, null, { renderer: "canvas" });
        chart.setOption(
        {
          animation: false,
          backgroundColor: "#FFFFFF",
          color: groups.map((group) => group.color),
          legend: {
            top: 8,
            right: 14,
            itemWidth: 10,
            itemHeight: 10,
            itemGap: 16,
            textStyle: { color: "#627985", fontSize: 11 },
          },
          grid: {
            top: 52,
            right: 46,
            bottom: 34,
            left: 98,
            containLabel: false,
          },
          tooltip: {
            trigger: "axis",
            axisPointer: { type: "shadow", shadowStyle: { color: "rgba(0,107,158,.05)" } },
            confine: true,
            transitionDuration: 0,
            backgroundColor: "rgba(255,255,255,.96)",
            borderColor: "rgba(156,181,192,.82)",
            borderWidth: 1,
            padding: [8, 10],
            textStyle: { color: "#203741", fontSize: 10 },
            extraCssText:
              "box-shadow:0 8px 20px rgba(7,31,45,.14);border-radius:4px;",
            formatter: (params) => {
              const dimension = params[0]?.axisValue ?? "";
              return [
                `<div style="line-height:1.6;white-space:nowrap"><strong>${model} · ${escapeHtml(dimension)}</strong>`,
                ...params.map(
                  ({ marker, seriesName, value }) =>
                    `<div>${marker}${escapeHtml(seriesName)}：<b>${Number(value).toFixed(1)}</b></div>`,
                ),
                "</div>",
              ].join("");
            },
          },
          xAxis: {
            type: "value",
            min: 0,
            max: 5,
            interval: 1,
            axisLabel: { color: "#80939C", fontSize: 10 },
            axisLine: { lineStyle: { color: "#BBCBD2" } },
            axisTick: { show: false },
            splitLine: { lineStyle: { color: "#E3EAED" } },
          },
          yAxis: {
            type: "category",
            inverse: true,
            data: dimensions,
            axisTick: { show: false },
            axisLine: { show: false },
            axisLabel: {
              color: "#344D5A",
              fontSize: 12,
              fontWeight: 600,
              margin: 18,
            },
          },
          series: groups.map((group) => ({
            name: group.name,
            type: "bar",
            barMaxWidth: 20,
            barGap: "20%",
            barCategoryGap: "38%",
            itemStyle: { color: group.color, borderRadius: [0, 2, 2, 0] },
            emphasis: { focus: "series" },
            label: {
              show: true,
              position: "right",
              distance: 5,
              color: "#344D5A",
              fontSize: 10,
              fontWeight: 700,
              formatter: ({ value }) => Number(value).toFixed(1),
            },
            data: dimensions.map((dimension) => {
              const score = rows.find(
                (row) =>
                  row.model === model &&
                  row.userType === group.userType &&
                  row.dimension === dimension,
              )?.score;
              return Number.isFinite(score) ? score : 0;
            }),
          })),
        },
        true,
        );
        charts.set(`brand-image-${model}`, chart);
        return chart;
      });
      return rendered;
    } catch (error) {
      container.innerHTML =
        '<span class="chart-status chart-status--error">品牌形象数据加载失败</span>';
      throw error;
    }
  };

  const initialize = async () => {
    OverviewInsight();
    await Promise.all([
      SampleOverviewTable(),
      TestDrivePieChart(),
      BrandImageScoreCard(),
    ]);
  };

  const resizeAll = () => charts.forEach((chart) => chart.resize());
  window.addEventListener("resize", resizeAll);

  window.OverviewDashboard = Object.freeze({
    BrandImageScoreCard,
    OverviewInsight,
    SampleOverviewTable,
    TestDrivePieChart,
    initialize,
    resizeAll,
  });
})();
