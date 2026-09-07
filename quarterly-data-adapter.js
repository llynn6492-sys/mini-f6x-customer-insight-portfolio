(() => {
  const config = window.QUARTERLY_DASHBOARD_CONFIG;
  let workbookPromise;

  const loadEmbeddedWorkbook = () => {
    const encoded = window.QUARTERLY_WORKBOOK_BASE64;
    if (!encoded) throw new Error("季度内置数据副本不可用。");
    const binary = window.atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    return window.XLSX.read(bytes, { type: "array" });
  };

  const loadWorkbook = async () => {
    if (!window.XLSX) throw new Error("Excel 读取组件不可用。");
    if (!workbookPromise) {
      // The deployed static site intentionally ships the workbook as an
      // embedded JavaScript payload. Some hosts return the app HTML with a
      // successful status for a missing .xlsx URL; SheetJS can parse that HTML
      // without throwing, leaving every expected sheet unavailable. Prefer the
      // embedded payload whenever it exists so local and hosted dashboards use
      // the exact same validated workbook.
      workbookPromise = window.QUARTERLY_WORKBOOK_BASE64
        ? Promise.resolve(loadEmbeddedWorkbook())
        : fetch(window.DASHBOARD_WORKBOOK_URL)
            .then((response) => {
              if (!response.ok) {
                throw new Error(`季度数据加载失败：${response.status}`);
              }
              return response.arrayBuffer();
            })
            .then((buffer) => window.XLSX.read(buffer));
    }
    return workbookPromise;
  };

  const rowsFromSheet = async (sheetName) => {
    const workbook = await loadWorkbook();
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return [];
    return window.XLSX.utils.sheet_to_json(worksheet, {
      defval: null,
      raw: true,
    });
  };

  const normalize = (row, rowIndex = 0) => ({
    section: row.Section ?? row["模块"] ?? row.Module,
    metric: row.Metric ?? row["指标"],
    category: row.Category ?? row["分类"],
    model: row.Model ?? row["车型"],
    series: row.Series ?? row["车系"] ?? "全部",
    count: Number(row.Count ?? row.N ?? 0),
    percentage: Number(row.Percentage ?? row["占比"] ?? row.Value),
    sampleSize: Number(row.Sample_Size ?? row["Sample_Size"] ?? row.N),
    averageAge: Number(row.Average_Age ?? row["平均年龄"]),
    sortOrder: Number(row.Sort_Order ?? row["排序"] ?? rowIndex + 1),
    extra: row.Extra ?? "",
  });

  const getMetric = async ({
    sheet,
    metric,
    model,
    series,
  }) =>
    (await rowsFromSheet(sheet))
      .map(normalize)
      .filter(
        (row) =>
          row.metric === metric &&
          row.model === model &&
          row.series === series,
      )
      .sort((a, b) => a.sortOrder - b.sortOrder);

  const isValidSelection = (model, series) =>
    Boolean(config.models[model]?.includes(series));

  window.QuarterlyData = Object.freeze({
    getMetric,
    isValidSelection,
    loadWorkbook,
    rowsFromSheet,
  });
})();
