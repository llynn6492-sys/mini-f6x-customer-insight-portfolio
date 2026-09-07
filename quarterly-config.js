window.DASHBOARD_WORKBOOK_URL =
  "data/quarterly_mock_data.xlsx?v=20260810-2";

window.QUARTERLY_DASHBOARD_CONFIG = Object.freeze({
  models: Object.freeze({
    F65: Object.freeze(["全部", "Cooper C", "Cooper S"]),
    F66: Object.freeze(["全部", "Cooper C", "Cooper S"]),
  }),
  defaultModel: "F65",
  defaultSeries: "全部",
  sheets: Object.freeze({
    profile: "Q_Profile_By_Series",
    purchaseBehavior: "Q_Purchase_Behavior_By_Series",
    purchaseMotivation: "Q_Purchase_Motivation_By_Series",
    decisionJourney: "Q_Decision_Journey_By_Series",
  }),
});
