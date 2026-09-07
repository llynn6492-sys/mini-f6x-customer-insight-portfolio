(() => {
  const definitions = {
    F65: {
      sampleSize: 188,
      metrics: {
        Gender: [
          ["男", 75],
          ["女", 113],
        ],
        "Marital Status": [
          ["单身", 64],
          ["未婚", 34],
          ["已婚有孩子", 58],
          ["已婚没有孩子", 32],
        ],
        Age: [
          ["20岁以下", 5],
          ["20-24岁", 14],
          ["25-29岁", 31],
          ["30-34岁", 44],
          ["35-39岁", 39],
          ["40-44岁", 28],
          ["45-49岁", 15],
          ["50岁及以上", 12],
        ],
        Education: [
          ["高中及以下", 13],
          ["专科", 31],
          ["本科", 103],
          ["硕士", 37],
          ["博士", 4],
        ],
        "Household Income": [
          ["20万以下", 28],
          ["20万-30万", 45],
          ["30万-40万", 48],
          ["40万-50万", 38],
          ["50万以上", 29],
        ],
        Industry: [
          ["医疗/制药/生物科技", 20],
          ["教育/培训/科研", 19],
          ["金融服务", 24],
          ["IT/互联网/电子商务", 42],
          ["公共事业/社会服务", 15],
          ["制造业", 28],
          ["自由职业", 22],
          ["其他", 18],
        ],
      },
    },
    F66: {
      sampleSize: 264,
      metrics: {
        Gender: [
          ["男", 96],
          ["女", 168],
        ],
        "Marital Status": [
          ["单身", 103],
          ["未婚", 44],
          ["已婚有孩子", 74],
          ["已婚没有孩子", 43],
        ],
        Age: [
          ["20岁以下", 9],
          ["20-24岁", 27],
          ["25-29岁", 52],
          ["30-34岁", 64],
          ["35-39岁", 50],
          ["40-44岁", 33],
          ["45-49岁", 17],
          ["50岁及以上", 12],
        ],
        Education: [
          ["高中及以下", 17],
          ["专科", 39],
          ["本科", 145],
          ["硕士", 56],
          ["博士", 7],
        ],
        "Household Income": [
          ["20万以下", 34],
          ["20万-30万", 58],
          ["30万-40万", 69],
          ["40万-50万", 55],
          ["50万以上", 48],
        ],
        Industry: [
          ["医疗/制药/生物科技", 25],
          ["教育/培训/科研", 30],
          ["金融服务", 36],
          ["IT/互联网/电子商务", 66],
          ["公共事业/社会服务", 20],
          ["制造业", 35],
          ["自由职业", 29],
          ["其他", 23],
        ],
      },
    },
  };

  const rows = Object.entries(definitions).flatMap(
    ([model, { sampleSize, metrics }]) =>
      Object.entries(metrics).flatMap(([metric, categories]) =>
        categories.map(([category, count], index) => ({
          Audience: "Prospect",
          Section: "Customer Profile",
          Metric: metric,
          Category: category,
          Model: model,
          Count: count,
          Percentage: count / sampleSize,
          Sort_Order: index + 1,
          Chart_Type:
            metric === "Gender" || metric === "Industry"
              ? "bar"
              : metric === "Marital Status"
                ? "pie"
                : "stacked_bar",
        })),
      ),
  );

  const sampleRows = Object.entries(definitions).map(
    ([model, { sampleSize }]) => ({
      Audience: "Prospect",
      Model: model,
      Sample_Size: sampleSize,
    }),
  );

  window.PROSPECT_PROFILE_MOCK = Object.freeze({
    rows: Object.freeze(rows),
    sampleRows: Object.freeze(sampleRows),
  });
})();
