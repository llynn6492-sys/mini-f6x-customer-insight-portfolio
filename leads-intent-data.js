(() => {
  const sampleSizes = Object.freeze({ F65: 434, F66: 474 });

  const purchaseType = {
    F65: [
      { category: "首购", count: 143 },
      { category: "增购", count: 126 },
      { category: "换购", count: 165 },
    ],
    F66: [
      { category: "首购", count: 176 },
      { category: "增购", count: 139 },
      { category: "换购", count: 159 },
    ],
  };

  Object.entries(purchaseType).forEach(([model, rows]) => {
    rows.forEach((row) => {
      row.percentage = row.count / sampleSizes[model];
    });
  });

  const currentVehicle = {
    F65: [
      ["BMW", 0.296, "3系、5系、X1"],
      ["奔驰", 0.182, "C级、GLA、GLC"],
      ["大众", 0.139, "高尔夫、途观、Polo"],
      ["奥迪", 0.121, "A3、A4L、Q3"],
      ["丰田", 0.096, "凯美瑞、RAV4、卡罗拉"],
      ["本田", 0.081, "雅阁、CR-V、飞度"],
      ["MINI", 0.074, "三门、五门、Clubman"],
      ["特斯拉", 0.058, "Model 3、Model Y"],
      ["其他", 0.052, "沃尔沃、雷克萨斯等"],
    ],
    F66: [
      ["BMW", 0.314, "3系、X1、X3"],
      ["奔驰", 0.169, "C级、GLA、A级"],
      ["奥迪", 0.146, "A3、A4L、Q3"],
      ["大众", 0.128, "高尔夫、途岳、Polo"],
      ["丰田", 0.103, "凯美瑞、RAV4、雷凌"],
      ["MINI", 0.089, "三门、五门、Countryman"],
      ["本田", 0.076, "雅阁、CR-V、思域"],
      ["特斯拉", 0.067, "Model 3、Model Y"],
      ["其他", 0.049, "沃尔沃、领克等"],
    ],
  };

  const previousVehicle = {
    F65: [
      ["BMW", 0.224, "3系、X1、5系"],
      ["大众", 0.153, "高尔夫、Polo、途观"],
      ["奔驰", 0.137, "C级、A级、GLA"],
      ["奥迪", 0.119, "A3、A4L、Q3"],
      ["本田", 0.096, "飞度、思域、CR-V"],
      ["丰田", 0.089, "卡罗拉、凯美瑞、RAV4"],
      ["MINI", 0.078, "三门、五门、Clubman"],
      ["日产", 0.061, "轩逸、逍客、天籁"],
      ["其他", 0.043, "福特、别克等"],
    ],
    F66: [
      ["BMW", 0.246, "3系、X1、X3"],
      ["奔驰", 0.148, "A级、C级、GLA"],
      ["大众", 0.141, "高尔夫、Polo、途岳"],
      ["奥迪", 0.128, "A3、Q3、A4L"],
      ["MINI", 0.094, "三门、五门、Countryman"],
      ["丰田", 0.083, "卡罗拉、凯美瑞、RAV4"],
      ["本田", 0.076, "飞度、思域、CR-V"],
      ["特斯拉", 0.052, "Model 3、Model Y"],
      ["其他", 0.032, "沃尔沃、领克等"],
    ],
  };

  const channel = {
    F65: [
      ["MINI伙伴", 0.258],
      ["小红书", 0.226],
      ["朋友推荐", 0.174],
      ["抖音", 0.152],
      ["汽车之家", 0.128],
      ["MINI官方网站", 0.096],
      ["微信公众账号", 0.078],
      ["经销商活动", 0.064],
    ],
    F66: [
      ["小红书", 0.271],
      ["MINI伙伴", 0.249],
      ["抖音", 0.188],
      ["朋友推荐", 0.161],
      ["汽车之家", 0.116],
      ["MINI官方网站", 0.104],
      ["经销商活动", 0.082],
      ["微信公众账号", 0.071],
    ],
  };

  const motivation = {
    F65: [
      ["已有多功能大车，期待一台小车", 0.286],
      ["上下班通勤刚需", 0.244],
      ["追求驾驶乐趣", 0.218],
      ["家庭需求增长", 0.176],
      ["品牌忠诚度", 0.153],
      ["城市停车更方便", 0.139],
      ["改善出行品质", 0.112],
      ["新能源尝鲜需求", 0.084],
    ],
    F66: [
      ["上下班通勤刚需", 0.267],
      ["已有多功能大车，期待一台小车", 0.252],
      ["追求驾驶乐趣", 0.231],
      ["城市停车更方便", 0.186],
      ["家庭需求增长", 0.162],
      ["品牌忠诚度", 0.148],
      ["改善出行品质", 0.119],
      ["新能源尝鲜需求", 0.096],
    ],
  };

  const consideration = {
    F65: [
      ["外观设计好", 0.342],
      ["产品操控性好", 0.296],
      ["喜欢内饰设计", 0.247],
      ["MINI品牌", 0.221],
      ["车身尺寸合适", 0.184],
      ["动力性能好", 0.163],
      ["个性化配置丰富", 0.137],
      ["销售服务好", 0.092],
    ],
    F66: [
      ["外观设计好", 0.371],
      ["产品操控性好", 0.312],
      ["喜欢内饰设计", 0.268],
      ["MINI品牌", 0.236],
      ["车身尺寸合适", 0.192],
      ["动力性能好", 0.174],
      ["智能化体验好", 0.149],
      ["销售服务好", 0.101],
    ],
  };

  const addRank = (collection) =>
    Object.fromEntries(
      Object.entries(collection).map(([model, rows]) => [
        model,
        rows.map(([brand, mentionRate, vehicleSeries], index) => ({
          rank: index + 1,
          brand,
          mentionRate,
          vehicleSeries,
        })),
      ]),
    );

  const addBarCounts = (collection) =>
    Object.fromEntries(
      Object.entries(collection).map(([model, rows]) => [
        model,
        rows.map(([category, percentage]) => ({
          category,
          percentage,
          count: Math.round(percentage * sampleSizes[model]),
        })),
      ]),
    );

  window.LEADS_INTENT_DATA = Object.freeze({
    userType: "Leads",
    sampleSizes,
    purchaseType,
    currentVehicle: addRank(currentVehicle),
    previousVehicle: addRank(previousVehicle),
    channel: addBarCounts(channel),
    motivation: addBarCounts(motivation),
    consideration: addBarCounts(consideration),
  });
})();
