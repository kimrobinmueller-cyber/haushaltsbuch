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

function getMealPlannerData(weekStart) {
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

function saveMealPlannerWeek(entries, weekStart) {
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
  const data=getMealPlannerData(weekStart);
  const recipeMap={}; data.recipes.forEach(r=>recipeMap[r.id]=r);
  const totals={};
  const normalizeUnit = raw => {
    const u=String(raw||'').trim().toLocaleLowerCase('de-DE').replace(/\./g,'');
    const map={kg:{unit:'g',factor:1000},g:{unit:'g',factor:1},gramm:{unit:'g',factor:1},l:{unit:'ml',factor:1000},ml:{unit:'ml',factor:1},milliliter:{unit:'ml',factor:1},stück:{unit:'Stück',factor:1},stk:{unit:'Stück',factor:1},el:{unit:'EL',factor:1},esslöffel:{unit:'EL',factor:1},tl:{unit:'TL',factor:1},teelöffel:{unit:'TL',factor:1},prise:{unit:'Prise',factor:1},prisen:{unit:'Prise',factor:1},dose:{unit:'Dose',factor:1},dosen:{unit:'Dose',factor:1},packung:{unit:'Packung',factor:1},packungen:{unit:'Packung',factor:1}};
    return map[u] || {unit:String(raw||'').trim(),factor:1};
  };
  data.plan.filter(p=>p.type==='recipe' && recipeMap[p.recipeId]).forEach(p=>{
    const recipe=recipeMap[p.recipeId];
    const factor=(Number(p.servings)||2)/(Number(recipe.servings)||4);
    (recipe.ingredients||[]).forEach(i=>{
      const name=String(i.name||'').trim(); if(!name) return;
      const normalized=normalizeUnit(i.unit);
      const key=name.toLocaleLowerCase('de-DE')+'|'+normalized.unit.toLocaleLowerCase('de-DE');
      const amount=Number(i.amount);
      if(!Number.isFinite(amount)) { const freeKey=key+'|freitext|'+String(i.unit||''); if(!totals[freeKey]) totals[freeKey]={name,amount:null,unit:String(i.unit||'')}; return; }
      if(!totals[key]) totals[key]={name,amount:0,unit:normalized.unit};
      totals[key].amount+=amount*factor*normalized.factor;
    });
  });
  return Object.values(totals).sort((a,b)=>a.name.localeCompare(b.name,'de') || String(a.unit).localeCompare(String(b.unit),'de'));
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



/**
 * KI-Rezeptvorschläge über die Gemini Interactions API.
 * Der API-Schlüssel wird ausschließlich aus den Script Properties gelesen.
 */
function getMealAiStatus() {
  return {configured:!!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')};
}

function generateAiMealRecipes(options) {
  options = options || {};
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Die bestehende Gemini-Verbindung ist nicht eingerichtet: Script Property GEMINI_API_KEY fehlt.');
  const models = [
    props.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL,
    CONFIG.GEMINI_FALLBACK
  ];
  const mode = options.mode === 'week' ? 'week' : 'single';
  const servings = Math.max(1, Math.min(10, Number(options.servings) || 4));
  const dayIndex = Math.max(0, Math.min(6, Number(options.dayIndex) || 0));
  const mealType = ['Frühstück','Mittagessen','Abendessen'].includes(String(options.mealType)) ? String(options.mealType) : 'Abendessen';
  const preferences = String(options.preferences || '').replace(/[<>]/g,' ').slice(0,500);
  const start = mealIsoDate_(options.weekStart || mealMonday_(new Date()));
  const schema = {
    type:'OBJECT',
    properties:{
      recipes:{type:'ARRAY',items:{
        type:'OBJECT',
        properties:{
          title:{type:'STRING'}, dayIndex:{type:'INTEGER'},
          mealType:{type:'STRING',enum:['Frühstück','Mittagessen','Abendessen']},
          servings:{type:'INTEGER'}, prepMinutes:{type:'INTEGER'},
          estimatedCostEur:{type:'NUMBER'}, category:{type:'STRING'},
          cuisine:{type:'STRING'}, tags:{type:'STRING'},
          ingredients:{type:'ARRAY',items:{type:'OBJECT',properties:{
            name:{type:'STRING'},amount:{type:'NUMBER'},unit:{type:'STRING'}
          },required:['name','amount','unit']}},
          instructions:{type:'STRING'}
        },
        required:['title','dayIndex','mealType','servings','prepMinutes','estimatedCostEur','category','cuisine','tags','ingredients','instructions']
      }}
    },
    required:['recipes']
  };
  let prompt;
  if (mode === 'week') {
    prompt = 'Erstelle genau 21 unterschiedliche, realistische und gut nachkochbare Familienrezepte für einen Wochenplan in Deutschland. Für jeden der 7 Tage (dayIndex 0=Montag bis 6=Sonntag) genau je ein Frühstück, Mittagessen und Abendessen. Jede Kombination aus Tag und Mahlzeit muss genau einmal vorkommen. Alle Rezepte für '+servings+' Portionen. Wünsche: '+(preferences || 'preiswert, familienfreundlich und alltagstauglich')+'. Verwende abwechslungsreiche, saisonale und möglichst preiswerte Zutaten; plane Reste sinnvoll ein und vermeide unnötig viele Spezialzutaten. Schreibe auf Deutsch. Zutaten müssen einzeln mit plausiblen numerischen Mengen und metrischen Einheiten angegeben werden (g, kg, ml, l, Stück, EL, TL oder Prise); keine Mengen als Freitext. Schätze die gesamten Zutatenkosten je Rezept in Euro für alle angegebenen Portionen. prepMinutes ist eine realistische Zubereitungszeit. Anleitungen in kurzen, klaren Schritten als Text. Gib ausschließlich das verlangte JSON zurück.';
  } else {
    prompt = 'Erstelle genau 3 unterschiedliche, realistische und gut nachkochbare Rezeptvorschläge für '+mealType+' am '+['Montag','Dienstag','Mittwoch','Donnerstag','Freitag','Samstag','Sonntag'][dayIndex]+' in Deutschland. Alle Rezepte für '+servings+' Portionen. Wünsche: '+(preferences || 'preiswert, familienfreundlich und alltagstauglich')+'. Schreibe auf Deutsch. Verwende abwechslungsreiche, möglichst preiswerte Zutaten. Zutaten müssen einzeln mit plausiblen numerischen Mengen und metrischen Einheiten angegeben werden (g, kg, ml, l, Stück, EL, TL oder Prise); keine Mengen als Freitext. Schätze die gesamten Zutatenkosten je Rezept in Euro für alle angegebenen Portionen. prepMinutes ist eine realistische Zubereitungszeit. Setze bei allen drei Rezepten dayIndex auf '+dayIndex+' und mealType auf "'+mealType+'". Anleitungen in kurzen, klaren Schritten als Text. Gib ausschließlich das verlangte JSON zurück.';
  }
  const call = callGemini_(models, apiKey, () => ({
    contents:[{role:'user',parts:[{text:prompt}]}],
    generationConfig:{
      responseMimeType:'application/json',
      responseSchema:schema,
      maxOutputTokens:12000
    }
  }));
  const result=JSON.parse(call.text);
  const parts=result.candidates&&result.candidates[0]&&result.candidates[0].content&&result.candidates[0].content.parts;
  const outputText=Array.isArray(parts)?parts.map(p=>p.text||'').join(''):'';
  if (!outputText) throw new Error('Gemini hat keine Rezeptdaten zurückgegeben. Bitte erneut versuchen.');
  let generated;
  try { generated=JSON.parse(outputText); } catch(e) { throw new Error('Die Gemini-Antwort war kein gültiges JSON. Bitte erneut versuchen.'); }
  if (!generated || !Array.isArray(generated.recipes)) throw new Error('Die Gemini-Antwort enthält keine Rezeptliste.');
  const expected=mode==='week'?21:3;
  if (generated.recipes.length!==expected) throw new Error('Gemini hat '+generated.recipes.length+' statt '+expected+' Rezepten erstellt. Bitte erneut versuchen.');
  const normalized=generated.recipes.map(r=>{
    const title=String(r.title||'').trim().slice(0,120);
    if (!title || !Array.isArray(r.ingredients) || !r.ingredients.length) throw new Error('Ein KI-Rezept ist unvollständig.');
    const day=Math.max(0,Math.min(6,Number(r.dayIndex)||0));
    const type=['Frühstück','Mittagessen','Abendessen'].includes(r.mealType)?r.mealType:mealType;
    const ingredients=r.ingredients.slice(0,40).map(x=>{
      const amount=Number(x.amount), name=String(x.name||'').trim().slice(0,100), unit=String(x.unit||'').trim().slice(0,20);
      if (!name || !Number.isFinite(amount) || amount<0) throw new Error('Eine Zutat hat eine ungültige Mengenangabe.');
      return {name,amount,unit};
    });
    return {
      id:Utilities.getUuid(),title,dayIndex:day,mealType:type,
      servings:Math.max(1,Math.min(10,Number(r.servings)||servings)),
      prepMinutes:Math.max(1,Math.min(300,Number(r.prepMinutes)||30)),
      estimatedCostEur:Math.max(0,Math.min(300,Number(r.estimatedCostEur)||0)),
      category:type,cuisine:String(r.cuisine||'').slice(0,60),
      tags:'KI, '+String(r.tags||'').slice(0,120),ingredients,
      instructions:String(r.instructions||'').slice(0,5000)
    };
  });
  if (mode==='week') {
    const slots=normalized.map(r=>r.dayIndex+'|'+r.mealType), expectedSlots=[];
    for(let d=0;d<7;d++) ['Frühstück','Mittagessen','Abendessen'].forEach(t=>expectedSlots.push(d+'|'+t));
    if(new Set(slots).size!==21 || expectedSlots.some(s=>!slots.includes(s))) throw new Error('Der KI-Wochenplan enthält doppelte oder fehlende Mahlzeiten. Bitte erneut versuchen.');
  } else if(normalized.some(r=>r.dayIndex!==dayIndex || r.mealType!==mealType)) {
    throw new Error('Gemini hat Mahlzeiten außerhalb der ausgewählten Auswahl erstellt. Bitte erneut versuchen.');
  }
  setupMealPlannerTables_();
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEAL_CONFIG.RECIPES);
  const now=new Date();
  const rows=normalized.map(r=>[
    r.id,r.title,r.category,r.cuisine,'','KI (Gemini)','',JSON.stringify(r.ingredients),
    r.instructions,r.servings,now,now,r.tags,
    'KI-Entwurf · ca. '+r.prepMinutes+' Min. · geschätzte Zutatenkosten: '+r.estimatedCostEur.toFixed(2)+' €'
  ]);
  sh.getRange(sh.getLastRow()+1,1,rows.length,14).setValues(rows);
  return {mode,weekStart:start,recipes:normalized};
}
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