/** API.gs – HTTP-API zusätzlich zur bestehenden google.script.run-App. */
function apiFunctions_() {
  return {
    getInitData, getAppData, getFormData, saveExpense, updateExpense,
    deleteExpense, saveMonthlyBudget, getShoppingData, analyzeReceipt,
    saveShoppingPurchase, getInsights,
    getMealData, searchMealRecipes, importMealRecipe,
    saveOwnMealRecipe, saveMealPlan, buildMealShoppingList,
    getMealShoppingChecklist, saveMealShoppingCheck
  };
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!token || req.token !== token) throw new Error('UNAUTHORIZED');
    const fn = apiFunctions_()[req.fn];
    if (!fn) throw new Error('Unbekannte Funktion: ' + req.fn);
    out = {ok:true,result:fn.apply(null,req.args||[])};
  } catch(err) {
    out = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function setupApiToken() {
  const props=PropertiesService.getScriptProperties();
  let token=props.getProperty('API_TOKEN');
  if(!token){token=Utilities.getUuid()+Utilities.getUuid().replace(/-/g,'');props.setProperty('API_TOKEN',token);}
  Logger.log('API_TOKEN ist eingerichtet.');
}

function getInitData(monthKey){return {form:getFormData(),app:getAppData(monthKey)};}

function getInsights(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(CONFIG.ITEMS_SHEET);
  if(!sheet || sheet.getLastRow()<2)return {forecast:[],savings:[]};
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,12).getValues();
  const DAY=86400000,today=new Date();today.setHours(12,0,0,0);
  const products={};
  rows.forEach(r=>{if(!(r[1] instanceof Date)||!r[11])return;const key=String(r[11]);const p=products[key]||(products[key]={name:'',category:'',dates:{},stores:{}});p.name=String(r[3]);p.category=String(r[9]);p.dates[formatDate_(r[1])]=true;const price=Number(r[7])||0;if(price>0){const store=String(r[2]||'?');const s=p.stores[store]||(p.stores[store]={sum:0,n:0,qty:0});s.sum+=price;s.n++;s.qty+=Number(r[5])||1;}});
  const forecast=[],savings=[];
  Object.values(products).forEach(p=>{const days=Object.keys(p.dates).sort();if(p.category==='Lebensmittel'&&days.length>=3){const ts=days.map(d=>new Date(d+'T12:00:00').getTime()),gaps=[];for(let i=1;i<ts.length;i++)gaps.push(Math.round((ts[i]-ts[i-1])/DAY));gaps.sort((a,b)=>a-b);const mid=Math.floor(gaps.length/2),median=gaps.length%2?gaps[mid]:Math.round((gaps[mid-1]+gaps[mid])/2);if(median>=1){const next=ts[ts.length-1]+median*DAY;forecast.push({name:p.name,purchases:days.length,everyDays:median,lastDate:days[days.length-1],daysLeft:Math.round((next-today.getTime())/DAY)});}}
    const list=Object.keys(p.stores).map(store=>({store,avg:p.stores[store].sum/p.stores[store].n,qty:p.stores[store].qty}));if(list.length>=2){list.sort((a,b)=>a.avg-b.avg);const cheapest=list[0],worst=list[list.length-1];let potential=0;list.forEach(x=>potential+=(x.avg-cheapest.avg)*x.qty);if(potential>0.01)savings.push({name:p.name,cheapestStore:cheapest.store,cheapestPrice:cheapest.avg,worstStore:worst.store,worstPrice:worst.avg,diffPct:(worst.avg/cheapest.avg-1)*100,potential});}});
  forecast.sort((a,b)=>a.daysLeft-b.daysLeft);savings.sort((a,b)=>b.potential-a.potential);return {forecast:forecast.slice(0,40),savings:savings.slice(0,25)};
}