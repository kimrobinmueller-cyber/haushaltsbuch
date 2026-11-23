// ================================================================
// HAUSHALTSBUCH – API.gs
// Zentrale API für GitHub Pages / Web-App.
// Voraussetzung: Code.gs und MealPlanner.gs sind im selben Projekt.
// In Code.gs müssen die bisherigen Definitionen von apiFunctions_(),
// doPost() und setupApiToken() entfernt werden, damit nichts doppelt ist.
// ================================================================

function apiFunctions_() {
  return {
    // Haushaltsbuch
    getInitData, getAppData, getFormData, saveExpense, updateExpense,
    deleteExpense, saveMonthlyBudget,

    // Einkaufsoptimierung und Kassenbon-KI
    getShoppingData, analyzeReceipt, saveShoppingPurchase, getInsights,

    // Bisheriger Essensplaner – Daten bleiben erhalten
    getMealData, saveDish, deleteDish, savePlan, generateDish,
    suggestDishes, buildShoppingList,

    // Neuer Essensplaner
    getMealPlannerData, searchMealRecipes, importMealRecipe,
    saveOwnMealRecipe, saveMealPlannerWeek, buildMealShoppingList,
    getMealAiStatus, generateAiMealRecipes,
    getMealShoppingChecklist, saveMealShoppingCheck
  };
}

function doPost(e) {
  let out;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Ungültige API-Anfrage.');
    }
    const req = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!token || req.token !== token) throw new Error('UNAUTHORIZED');

    const fn = apiFunctions_()[req.fn];
    if (typeof fn !== 'function') throw new Error('Unbekannte Funktion: ' + req.fn);

    const args = Array.isArray(req.args) ? req.args : [];
    out = {ok:true, result:fn.apply(null,args)};
  } catch (err) {
    out = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
\n