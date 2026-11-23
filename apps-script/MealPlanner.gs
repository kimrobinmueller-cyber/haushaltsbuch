/**
 * ESSENSPLANER – Erweiterung für das bestehende Haushaltsbuch
 * Neue Tabellen werden nur angelegt, wenn sie fehlen. Bestehende Tabellen
 * und Buchungen werden nicht gelöscht oder geleert.
 */
const MEAL_CONFIG = {
  RECIPES: 'Rezepte',
  PLAN: 'Wochenplan',
  SOURCE: 'Rezeptquellen',
  CHECKLIST: 'Einkaufscheckliste',
  THEMEALDB: 'https://www.themealdb.com/api/json/v1/1/'
};

function setupMealPlannerTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defs = [
    [MEAL_CONFIG.RECIPES, ['Rezept-ID','Titel','Kategorie','Küche','Bild-URL','Quelle','Quell-ID','Zutaten-JSON','Anleitung','Portionen','Erstellt am','Zuletzt geändert','Tags','Eigene Notizen']],
    [MEAL_CONFIG.PLAN, ['Eintrag-ID','Wochenstart','Datum','Mahlzeit','Position','Typ','Rezept-ID','Titel','Zielgruppe','Portionen','Notiz','Aktualisiert am']],
    [MEAL_CONFIG.SOURCE, ['Quelle','Lizenz/Attribution','Hinweis']],
    [MEAL_CONFIG.CHECKLIST, ['Wochenstart','Zutatenschlüssel','Zutat','Erledigt','Aktualisiert am']]
  ];
  defs.forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  });
  const source = ss.getSheetByName(MEAL_CONFIG.SOURCE);
  if (source.getLastRow() < 2) source.appendRow(['TheMealDB','Bitte bei importierten Rezepten Quelle und Urheberhinweise am Rezept erhalten.','Datenübernahme nur im Rahmen der jeweils geltenden API-Bedingungen.']);
  return true;
}

function getMealData(weekStart) {
  setupMealPlannerTables_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recipeSheet = ss.getSheetByName(MEAL_CONFIG.RECIPES);
  const planSheet = ss.getSheetByName(MEAL_CONFIG.PLAN);
  const recipes = mealRows_(recipeSheet,14).map(r => ({
    id:String(r[0]||''), title:String(r[1]||''), category:String(r[2]||''), cuisine:String(r[3]||''),
    image:String(r[4]||''), source:String(r[5]||''), sourceId:String(r[6]||''),
    ingredients:mealJson_(r[7],[]), instructions:String(r[8]||''), servings:Number(r[9])||4,
    tags:String(r[12]||''), notes:String(r[13]||'')
  })).filter(x=>x.id);
  const start = mealIsoDate_(weekStart || mealMonday_(new Date()));
  const rows = mealRows_(planSheet,12).map(r=>({
    id:String(r[0]||''), weekStart:mealCellDate_(r[1]), date:mealCellDate_(r[2]),
    meal:String(r[3]||''), position:Number(r[4])||0, type:String(r[5]||'recipe'),
    recipeId:String(r[6]||''), title:String(r[7]||''), audience:String(r[8]||'Alle'),
    servings:Number(r[9])||2, note:String(r[10]||'')
  })).filter(x=>x.id && x.weekStart===start);
  return {weekStart:start, recipes, plan:rows};
}

function searchMealRecipes(query, category) {
  const q = String(query||'').trim();
  if (!q && !category) return {local:mealLocalRecipes_(''), external:[]};
  const local = mealLocalRecipes_(q,category);
  let external=[];
  if (q) {
    try {
      const url=MEAL_CONFIG.THEMEALDB+'search.php?s='+encodeURIComponent(q);
      const response=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
      if (response.getResponseCode()===200) {
        const data=JSON.parse(response.getContentText());
        external=(data.meals||[]).slice(0,12).map(mealNormalizeExternal_);
      }
    } catch(e) { /* lokale Treffer bleiben verfügbar */ }
  }
  return {local,external};
}

function importMealRecipe(sourceId) {
  const id=String(sourceId||'').trim();
  if (!/^\d+$/.test(id)) throw new Error('Ungültige Rezept-ID.');
  setupMealPlannerTables_();
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.RECIPES);
  const existing=mealRows_(sh,14).find(r=>String(r[6]||'')===id && String(r[5]||'')==='TheMealDB');
  if (existing) return {success:true,recipeId:String(existing[0]),alreadyImported:true};
  const response=UrlFetchApp.fetch(MEAL_CONFIG.THEMEALDB+'lookup.php?i='+encodeURIComponent(id),{muteHttpExceptions:true});
  if(response.getResponseCode()!==200) throw new Error('Rezeptquelle ist momentan nicht erreichbar.');
  const data=JSON.parse(response.getContentText());
  const m=(data.meals||[])[0];
  if(!m) throw new Error('Rezept nicht gefunden.');
  const normalized=mealNormalizeExternal_(m);
  const recipeId=Utilities.getUuid();
  sh.appendRow([recipeId,normalized.title,normalized.category,normalized.cuisine,normalized.image,'TheMealDB',id,JSON.stringify(normalized.ingredients),normalized.instructions,4,new Date(),new Date(),'Importiert','']);
  return {success:true,recipeId};
}

function saveOwnMealRecipe(payload) {
  if(!payload || !String(payload.title||'').trim()) throw new Error('Bitte einen Rezeptnamen eingeben.');
  setupMealPlannerTables_();
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.RECIPES);
  const id=String(payload.id||Utilities.getUuid());
  const rows=mealRows_(sh,14);
  const idx=rows.findIndex(r=>String(r[0]||'')===id);
  const ingredients=Array.isArray(payload.ingredients)?payload.ingredients:[];
  const row=[id,String(payload.title).trim(),String(payload.category||''),String(payload.cuisine||''),String(payload.image||''),'Eigene Rezepte','',JSON.stringify(ingredients),String(payload.instructions||''),Math.max(1,Number(payload.servings)||2),idx>=0?rows[idx][10]||new Date():new Date(),new Date(),String(payload.tags||''),String(payload.notes||'')];
  if(idx>=0) sh.getRange(idx+2,1,1,14).setValues([row]); else sh.appendRow(row);
  return {success:true,id};
}

function saveMealPlan(entries, weekStart) {
  setupMealPlannerTables_();
  if(!Array.isArray(entries)) throw new Error('Ungültiger Wochenplan.');
  const start=mealIsoDate_(weekStart);
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.PLAN);
  const last=sh.getLastRow();
  const old=last>1?sh.getRange(2,1,last-1,12).getValues():[];
  const keep=old.filter(r=>mealCellDate_(r[1])!==start);
  const clean=entries.map((e,i)=>{
    const date=mealIsoDate_(e.date);
    if(date<start || date>mealAddDays_(start,6)) throw new Error('Ein Eintrag liegt außerhalb der ausgewählten Woche.');
    const type=e.type==='placeholder'?'placeholder':'recipe';
    const title=String(e.title||'').trim();
    if(!title) throw new Error('Jeder Planeintrag benötigt einen Titel.');
    return [String(e.id||Utilities.getUuid()),start,mealParseDate_(date),String(e.meal||'Abendessen'),Number(e.position)||i,type,String(e.recipeId||''),title,String(e.audience||'Alle'),Math.max(1,Number(e.servings)||2),String(e.note||''),new Date()];
  });
  // Nur die Einträge der ausgewählten Woche werden ersetzt; andere Wochen bleiben erhalten.
  if(last>1) sh.getRange(2,1,last-1,12).clearContent();
  const combined=keep.concat(clean);
  if(combined.length) sh.getRange(2,1,combined.length,12).setValues(combined);
  return {success:true,weekStart:start,count:clean.length};
}

function buildMealShoppingList(weekStart) {
  const data=getMealData(weekStart);
  const recipeMap={}; data.recipes.forEach(r=>recipeMap[r.id]=r);
  const totals={};
  data.plan.filter(p=>p.type==='recipe' && recipeMap[p.recipeId]).forEach(p=>{
    const recipe=recipeMap[p.recipeId];
    const factor=(Number(p.servings)||2)/(Number(recipe.servings)||4);
    (recipe.ingredients||[]).forEach(i=>{
      const name=String(i.name||'').trim(); if(!name) return;
      const key=name.toLocaleLowerCase('de-DE');
      if(!totals[key]) totals[key]={name,amount:0,unit:String(i.unit||'')};
      const amount=Number(i.amount);
      if(Number.isFinite(amount)) totals[key].amount+=amount*factor;
    });
  });
  return Object.values(totals).sort((a,b)=>a.name.localeCompare(b.name,'de'));
}

function setupMealPlannerTables_(){ setupMealPlannerTables(); }
function mealRows_(sh,width){ const n=sh.getLastRow(); return n>1?sh.getRange(2,1,n-1,width).getValues():[]; }
function mealJson_(v,fallback){try{return JSON.parse(String(v||''));}catch(e){return fallback;}}
function mealIsoDate_(v){ if(v instanceof Date)return Utilities.formatDate(v,'Europe/Berlin','yyyy-MM-dd'); const s=String(v||''); if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Ungültiges Datum.'); return s; }
function mealCellDate_(v){return v instanceof Date?Utilities.formatDate(v,'Europe/Berlin','yyyy-MM-dd'):String(v||'').slice(0,10);}
function mealParseDate_(s){const a=s.split('-').map(Number);return new Date(a[0],a[1]-1,a[2],12);}
function mealAddDays_(s,n){const d=mealParseDate_(s);d.setDate(d.getDate()+n);return Utilities.formatDate(d,'Europe/Berlin','yyyy-MM-dd');}
function mealMonday_(d){const x=new Date(d);x.setHours(12,0,0,0);x.setDate(x.getDate()-((x.getDay()+6)%7));return Utilities.formatDate(x,'Europe/Berlin','yyyy-MM-dd');}
function mealLocalRecipes_(q,category){setupMealPlannerTables_();const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.RECIPES);const term=String(q||'').toLocaleLowerCase('de-DE');return mealRows_(sh,14).map(r=>({id:String(r[0]||''),title:String(r[1]||''),category:String(r[2]||''),cuisine:String(r[3]||''),image:String(r[4]||''),source:String(r[5]||''),sourceId:String(r[6]||''),ingredients:mealJson_(r[7],[]),instructions:String(r[8]||''),servings:Number(r[9])||4,tags:String(r[12]||'')})).filter(x=>x.id && (!term || (x.title+' '+x.category+' '+x.tags+' '+x.ingredients.map(i=>i.name).join(' ')).toLocaleLowerCase('de-DE').includes(term)) && (!category || x.category===category));}
function mealNormalizeExternal_(m){const ingredients=[];for(let i=1;i<=20;i++){const name=String(m['strIngredient'+i]||'').trim();const measure=String(m['strMeasure'+i]||'').trim();if(name)ingredients.push({name,amount:null,unit:measure});}return {sourceId:String(m.idMeal||''),title:String(m.strMeal||'Rezept'),category:String(m.strCategory||''),cuisine:String(m.strArea||''),image:String(m.strMealThumb||''),instructions:String(m.strInstructions||''),ingredients};}


/** Gemeinsame, geräteübergreifende Häkchen für den Wocheneinkauf. */
function getMealShoppingChecklist(weekStart) {
  setupMealPlannerTables_();
  const start = mealIsoDate_(weekStart);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.CHECKLIST);
  return mealRows_(sh,5).filter(r => mealCellDate_(r[0]) === start).reduce((out,r) => {
    out[String(r[1]||'')] = r[3] === true || String(r[3]).toLowerCase() === 'true';
    return out;
  },{});
}

function saveMealShoppingCheck(weekStart, itemName, checked) {
  setupMealPlannerTables_();
  const start = mealIsoDate_(weekStart);
  const name = String(itemName||'').trim();
  if (!name) throw new Error('Zutat fehlt.');
  const key = name.toLocaleLowerCase('de-DE').replace(/\s+/g,' ');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.CHECKLIST);
  const rows = mealRows_(sh,5);
  const idx = rows.findIndex(r => mealCellDate_(r[0]) === start && String(r[1]||'') === key);
  const row = [mealParseDate_(start),key,name,!!checked,new Date()];
  if (idx >= 0) sh.getRange(idx+2,1,1,5).setValues([row]);
  else sh.appendRow(row);
  return {success:true,weekStart:start,item:key,checked:!!checked};
}