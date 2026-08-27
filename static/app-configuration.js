async function configuration(id){
  let list=await api('/api/sessions'),s=list.find(x=>x.id===id);
  if(!s)return notice('Atelier introuvable.');
  let t=await api('/api/templates/'+s.template_id);
  let a=await api('/api/sessions/'+id+'/analysis').catch(()=>null);
  window.configHasData=!!(a&&a.participantCount>0);
  let aiCfg=await api('/api/ai/config').catch(()=>({enabled:false,provider:null,model:null,keyConfigured:false,providers:{}}));
  window.aiCfg=aiCfg;
  let profile=s.profile_schema_id?await api('/api/profile-schemas/'+s.profile_schema_id).catch(()=>null):null;
  window.scaleLocked=!!(a&&a.global&&a.global.responses>0);
  window.currentSessionId=id;
  let domains=t.domains.filter(d=>d.active);
  let totalInd=domains.reduce((n,d)=>n+d.indicators.filter(i=>i.active).length,0);
  let expected=s.expected_participants,sliderVal=expected||20;
  window.scaleDraft={min:+t.scale.min,max:+t.scale.max,labels:{...(t.scale.labels||{})}};
  let steps=[
    {label:'Mission',sub:'Défini',done:!!(s.name&&s.location)},
    {label:'Questionnaire',sub:'Sélectionné',done:domains.length>0&&totalInd>0},
    {label:'Participants',sub:'Paramétrés',done:!!s.expected_participants},
    {label:'Vérification',sub:'En attente',done:false},
  ];
  let doneCount=steps.filter(x=>x.done).length,pct=Math.round(doneCount/steps.length*100);
  let ready=domains.length>0&&totalInd>0;
  shell('config','config','CONFIGURATION DE LA MISSION','Paramétrez votre mission avant de lancer la collecte',{id,name:s.name,location:s.location,date:s.date,pct,statusLabel:'Préparation'});
  app.innerHTML=`${window.configHasData?`<div class="card" style="border-left:4px solid var(--orange-600);background:var(--orange-50);margin-bottom:1rem"><b>⚠ Cet atelier contient déjà des données réelles</b><p class="muted small" style="margin:.3rem 0 0">${a.participantCount} commencé${a.participantCount>1?'s':''} · ${a.completedCount} validé${a.completedCount>1?'s':''}. Modifier le nom ci-dessous renomme cet atelier existant — cela ne crée pas un nouvel atelier vierge. Pour démarrer un diagnostic entièrement neuf, utilisez « Nouveau diagnostic » depuis l’accueil.</p></div>`:''}<div class="cfg-grid-71"><div class="card"><div class="row" style="align-items:center;gap:1.6rem;flex-wrap:wrap">${gaugeCircle(pct)}<div style="flex:1;min-width:240px"><h2>Préparation de la mission</h2><p class="muted small">Vérifiez et ajustez les paramètres de votre atelier avant d’ouvrir la collecte.</p><div class="hstepper">${steps.map((st,i)=>`<div class="hstep ${st.done?'done':((i===0||steps[i-1].done)&&!st.done?'current':'')}"><span class="circle">${st.done?'✓':i+1}</span><span class="hlabel">${st.label}<br>${st.sub}</span></div>`).join('')}</div></div></div></div><div class="card"><h3>Aperçu de votre mission</h3><div class="overview-list"><div class="ov-row"><span class="ov-label">Objet de la mission</span><span class="ov-value">${esc(s.name)}</span></div><div class="ov-row"><span class="ov-label">Questionnaire</span><span class="ov-value">${esc(t.name)} v${t.version}</span></div><div class="ov-row"><span class="ov-label">Domaines</span><span class="ov-value">${domains.length}</span></div><div class="ov-row"><span class="ov-label">Indicateurs</span><span class="ov-value">${totalInd}</span></div><div class="ov-row"><span class="ov-label">Objectif (participants prévus)</span><span class="ov-value">${s.expected_participants||'—'}</span></div><div class="ov-row"><span class="ov-label">Échelle de notation</span><span class="ov-value">${t.scale.min} à ${t.scale.max}</span></div></div></div></div><div class="cfg-grid-3"><div class="card"><div class="section-header"><h3>1 · Informations de la mission</h3></div><form onsubmit="saveConfigInfo(event,'${id}')"><label>Objet de la mission<input name="name" required value="${esc(s.name)}"></label><div class="row row-fields"><label>Lieu<input name="location" value="${esc(s.location||'')}"></label><label>Date<input name="date" type="date" value="${esc(s.date||'')}"></label></div><label>Description (optionnelle)<textarea name="description">${esc(s.description||'')}</textarea></label><button class="secondary" type="submit">Enregistrer</button></form></div><div class="card" style="text-align:center"><div class="section-header" style="text-align:left"><h3>2 · Participants</h3></div><div id="participants-gauge">${gaugeSemi(sliderVal,1,500,!expected)}</div><input id="participants-range" type="range" min="1" max="500" value="${sliderVal}" style="width:100%" oninput="updateParticipantsGauge(this.value)"><input id="participants-number" type="number" min="1" placeholder="Non défini" value="${expected||''}" style="width:100%;margin-top:.5rem" onchange="document.querySelector('#participants-range').value=this.value||20;updateParticipantsGauge(this.value||20)"><p class="text-meta" style="margin-top:.6rem;text-align:left">Ce nombre sert de référence pour les analyses. Le nombre réel de participants reste calculé à partir des données collectées.</p><button class="secondary block" onclick="saveExpectedParticipants('${id}')">Enregistrer</button></div><div class="card"><div class="section-header"><h3>3 · Structure du questionnaire</h3></div><div class="pill-row"><div class="stat-pill"><div class="pill-ic">${ICON_LAYERS}</div><div><b>${domains.length}</b><span>Domaines</span></div></div><div class="stat-pill"><div class="pill-ic">${ICON_DOC}</div><div><b>${totalInd}</b><span>Indicateurs</span></div></div></div><p class="text-meta">Répartition par domaine</p>${domains.map((d,i)=>{let n=d.indicators.filter(x=>x.active).length,max=Math.max(...domains.map(x=>x.indicators.filter(y=>y.active).length),1);return `<div class="domain-bar-row"><span class="db-label">${i+1}. ${esc(d.label)}</span><span class="db-track"><span class="db-fill" style="width:${n/max*100}%;background:${DOMAIN_COLORS[i%7]}"></span></span><span class="db-count">${n}</span></div>`}).join('')}<button class="secondary block" style="margin-top:.8rem" onclick="edit('${t.id}')">Voir / Modifier le questionnaire →</button>${scaleLocked?`<p class="text-meta" style="margin-top:.6rem">Cet atelier a déjà des réponses : le questionnaire ne peut plus être changé ou recréé.</p>`:`<div class="row" style="margin-top:.5rem;gap:.5rem;flex-wrap:wrap"><button class="ghost" onclick="pickQuestionnaire('${id}')">🔁 Changer de questionnaire</button><button class="ghost" onclick="createQuestionnaireForSession('${id}')">+ Nouveau questionnaire</button></div>`}</div></div><div class="cfg-grid-71"><div class="card"><div class="section-header"><h3>4 · Échelle de notation</h3></div>${scaleLocked?`<p class="muted small">Cet atelier a déjà des réponses enregistrées. L’échelle de notation ne peut plus être modifiée : la changer invaliderait ces réponses.</p><div class="scale-cards" style="--scale-cols:${scaleDraft.max-scaleDraft.min+1}">${Array.from({length:scaleDraft.max-scaleDraft.min+1},(_,k)=>{let v=scaleDraft.min+k,tone=scaleTone(k+1,scaleDraft.max-scaleDraft.min+1);return `<div class="scale-card sc-${tone}"><div class="sc-num">${v}</div><div class="sc-label">${esc(scaleDraft.labels?.[v]||'')}</div><div class="sc-emoji">${SCALE_EMOJI[tone-1]}</div></div>`}).join('')}</div>`:`<p class="muted small">Les participants noteront chaque indicateur selon l’échelle suivante. Renommez les niveaux ou ajustez leur nombre selon vos besoins.</p><form onsubmit="saveScale(event,'${id}','${t.id}')"><div class="scale-cards" id="scale-cards" style="--scale-cols:${scaleDraft.max-scaleDraft.min+1}">${renderScaleCards(scaleDraft.min,scaleDraft.max,scaleDraft.labels)}</div><div class="row" style="margin-top:.7rem;flex-wrap:wrap;gap:.5rem"><button type="button" class="secondary" onclick="removeScaleLevel()">− Retirer un niveau</button><button type="button" class="secondary" onclick="addScaleLevel()">+ Ajouter un niveau</button><button class="secondary" type="submit" style="margin-left:auto">Enregistrer l’échelle</button></div></form>`}</div><div class="card launch-block"><div class="launch-ic">🛡️</div><h3>Prêt à ouvrir la collecte ?</h3><p class="muted small">Vérifiez que tous les paramètres sont corrects puis ouvrez la collecte pour les participants.</p><button class="primary-lg" onclick="launchCollecte('${id}',${ready})">🚀 Ouvrir la collecte</button><button class="secondary" onclick="preview('${id}')">👁 Aperçu du questionnaire</button></div></div>${aiConfigCardHtml(aiCfg)}${profileConfigCardHtml(id,profile)}<p class="text-meta" style="margin-top:.5rem;display:flex;justify-content:space-between;flex-wrap:wrap;gap:.4rem"><span>Vous pourrez ajuster certains paramètres pendant l’atelier. Toute modification du questionnaire sera versionnée.</span><span>Dernière sauvegarde : ${new Date().toLocaleString('fr-FR')}</span></p>`;
}
const PROFILE_TYPE_LABELS={text:'Texte libre',number:'Nombre',single_choice:'Choix unique',multi_choice:'Choix multiple'};
function profileConfigCardHtml(id,profile){
  return `<div class="cfg-grid-71"><div class="card"><div class="section-header"><h3>6 · Profil participant (optionnel)</h3></div><p class="muted small">Ajoutez des champs (organisation, fonction, ancienneté…) que chaque participant renseignera en plus de ses réponses au questionnaire.</p>${profile
    ?`<p class="text-meta">Profil actif : <b>${esc(profile.name)}</b> — ${profile.fields.length} champ${profile.fields.length>1?'s':''}.</p><div class="row" style="gap:.5rem;flex-wrap:wrap"><button class="secondary" onclick="profileEditor('${id}','${profile.id}')">Gérer les champs →</button><button class="ghost" onclick="detachProfile('${id}')">Retirer le profil de cet atelier</button></div>`
    :`<button class="secondary" onclick="createProfileForSession('${id}')">+ Créer un profil participant</button>`
  }</div></div>`;
}
async function createProfileForSession(id){
  let x=await api('/api/profile-schemas',{method:'POST',body:JSON.stringify({name:'Profil participant'})});
  let cur=(await api('/api/sessions')).find(s=>s.id===id);
  await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({name:cur.name,organization:cur.organization,location:cur.location,date:cur.date,description:cur.description,expectedParticipants:cur.expected_participants,profileSchemaId:x.id})});
  profileEditor(id,x.id);
}
async function detachProfile(id){
  if(!confirm('Retirer le profil participant de cet atelier ? Les champs et les réponses déjà collectées restent conservés, mais ne seront plus affichés ni proposés aux prochains participants.'))return;
  let cur=(await api('/api/sessions')).find(s=>s.id===id);
  await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({name:cur.name,organization:cur.organization,location:cur.location,date:cur.date,description:cur.description,expectedParticipants:cur.expected_participants,profileSchemaId:null})});
  notice('Profil retiré de cet atelier.');
  configuration(id);
}
async function profileEditor(sessionId,schemaId){
  let p=await api('/api/profile-schemas/'+schemaId);
  window.currentSessionId=sessionId;
  shell('config','config',p.name,'Profil participant — champs personnalisés');
  app.innerHTML=`<section class="card"><button class="secondary" onclick="configuration('${sessionId}')">← Retour à la configuration</button>
  <h2>Profil participant</h2>
  <form onsubmit="saveProfileSchemaName(event,'${sessionId}','${schemaId}')" class="row row-fields" style="align-items:flex-end"><label style="flex:1">Nom du profil<input name="name" required value="${esc(p.name)}"></label><button class="secondary" type="submit">Renommer</button></form>
  <hr>
  ${p.fields.length?p.fields.map((f,n)=>`<div class="indicator"><b>${n+1}. ${esc(f.label)}</b> <span class="text-meta">(${PROFILE_TYPE_LABELS[f.field_type]||f.field_type}${f.required?', obligatoire':''}${f.is_dimension?', dimension d’analyse':''}${f.options&&f.options.length?' — '+f.options.map(o=>esc(profileOptLabel(o))).join(', '):''})</span><div class="row" style="margin-top:.4rem"><button class="secondary" onclick="profileFieldForm('${sessionId}','${schemaId}','${f.id}')">Modifier</button><button class="danger" onclick="deleteProfileField('${sessionId}','${schemaId}','${f.id}')">Supprimer</button></div></div>`).join(''):'<p class="muted">Aucun champ pour le moment.</p>'}
  <button style="margin-top:.6rem" onclick="profileFieldForm('${sessionId}','${schemaId}')">+ Ajouter un champ</button>
  </section>`;
}
function profileFieldForm(sessionId,schemaId,fieldId){
  api('/api/profile-schemas/'+schemaId).then(p=>{
    let f=fieldId?p.fields.find(x=>x.id===fieldId):null;
    let optionsText=f&&f.options?f.options.map(o=>profileOptLabel(o)).join('\n'):'';
    shell('config','config',p.name,fieldId?'Modifier le champ':'Nouveau champ');
    app.innerHTML=`<section class="card"><button class="secondary" onclick="profileEditor('${sessionId}','${schemaId}')">← Retour au profil</button>
    <h2>${fieldId?'Modifier le champ':'Nouveau champ'}</h2>
    <form onsubmit="saveProfileField(event,'${sessionId}','${schemaId}','${fieldId||''}')">
    <label>Libellé<input name="label" required value="${esc(f?f.label:'')}"></label>
    <label>Type de champ<select name="fieldType" onchange="toggleProfileOptionsField(this.value)">${Object.entries(PROFILE_TYPE_LABELS).map(([k,l])=>`<option value="${k}" ${f&&f.field_type===k?'selected':''}>${l}</option>`).join('')}</select></label>
    <label class="row" style="align-items:center;gap:.5rem"><input type="checkbox" name="required" style="width:auto" ${f&&f.required?'checked':''}> Champ obligatoire</label>
    <label id="profile-options-field" style="display:${f&&(f.field_type==='single_choice'||f.field_type==='multi_choice')?'':'none'}">Options (une par ligne)<textarea name="options">${esc(optionsText)}</textarea></label>
    <label id="profile-dimension-field" class="row" style="align-items:center;gap:.5rem;display:${f&&(f.field_type==='single_choice'||f.field_type==='multi_choice')?'':'none'}"><input type="checkbox" name="isDimension" style="width:auto" ${f&&f.is_dimension?'checked':''}> Utiliser comme dimension d’analyse (filtrage / comparaison)</label>
    <p class="text-meta">Une dimension permet de filtrer et comparer les résultats du diagnostic par cette réponse (ex : rôle, structure). Les résultats d’un sous-groupe trop petit sont automatiquement masqués.</p>
    <button type="submit">${fieldId?'Enregistrer':'Ajouter le champ'}</button>
    </form></section>`;
  });
}
function toggleProfileOptionsField(type){
  let show=(type==='single_choice'||type==='multi_choice')?'':'none';
  let opt=document.querySelector('#profile-options-field'); if(opt)opt.style.display=show;
  let dim=document.querySelector('#profile-dimension-field'); if(dim)dim.style.display=show;
}
async function saveProfileField(e,sessionId,schemaId,fieldId){
  e.preventDefault();
  let f=Object.fromEntries(new FormData(e.target));
  let body={fieldType:f.fieldType,label:f.label,required:!!f.required,isDimension:!!f.isDimension,options:(f.options||'').split('\n').map(x=>x.trim()).filter(Boolean)};
  try{
    if(fieldId)await api('/api/profile-fields/'+fieldId,{method:'PUT',body:JSON.stringify(body)});
    else await api('/api/profile-schemas/'+schemaId+'/fields',{method:'POST',body:JSON.stringify(body)});
  }catch(err){return notice(err.message)}
  profileEditor(sessionId,schemaId);
}
async function deleteProfileField(sessionId,schemaId,fieldId){
  if(!confirm('Supprimer ce champ de profil ?'))return;
  try{await api('/api/profile-fields/'+fieldId,{method:'DELETE'})}catch(err){return notice(err.message)}
  profileEditor(sessionId,schemaId);
}
async function saveProfileSchemaName(e,sessionId,schemaId){
  e.preventDefault();
  let f=Object.fromEntries(new FormData(e.target));
  await api('/api/profile-schemas/'+schemaId,{method:'PUT',body:JSON.stringify({name:f.name})});
  notice('Profil renommé.');
  profileEditor(sessionId,schemaId);
}
async function participantsList(id){
  let s=(await api('/api/sessions')).find(x=>x.id===id);
  if(!s)return notice('Atelier introuvable.');
  window.currentSessionId=id;
  // profile_schema_id is already known from s (the sessions list), so this
  // fetch doesn't need to wait on the unrelated participants fetch - both
  // run concurrently instead of profile being serialized after it.
  let [participants,profile]=await Promise.all([
    api('/api/sessions/'+id+'/participants'),
    s.profile_schema_id?api('/api/profile-schemas/'+s.profile_schema_id).catch(()=>null):Promise.resolve(null),
  ]);
  let profileFields=profile?profile.fields.filter(f=>f.active):[];
  shell('participants','collecte',s.name,'Participants de l’atelier');
  app.innerHTML=`<section class="card"><button class="ghost" onclick="moderator('${id}')">← Retour à l’atelier</button><h2>Participants — ${esc(s.name)}</h2>
  <p class="muted small">${participants.length} participant${participants.length>1?'s':''} — ${participants.filter(p=>p.status==='completed').length} validé${participants.filter(p=>p.status==='completed').length>1?'s':''}.</p>
  <div class="row" style="margin-bottom:.8rem"><button type="button" class="secondary" onclick="downloadServerFile('/api/sessions/${id}/individual-responses.xlsx','reponses-individuelles.xlsx',this)">Exporter les réponses individuelles (Excel)</button><button type="button" class="secondary" onclick="downloadServerFile('/api/sessions/${id}/individual-responses.csv','reponses-individuelles.csv',this)">Exporter les réponses individuelles (CSV)</button></div>
  <div class="table-wrap"><table><tr><th>Identifiant</th><th>Statut</th><th>Commencé</th><th>Terminé</th>${profileFields.map(f=>`<th>${esc(f.label)}</th>`).join('')}</tr>
  ${participants.length?participants.map(p=>`<tr><td>${esc(p.display_name||p.anonymous_id)}</td><td>${p.status==='completed'?'Validé':'En cours'}</td><td>${new Date(p.started_at).toLocaleString('fr-FR')}</td><td>${p.completed_at?new Date(p.completed_at).toLocaleString('fr-FR'):'—'}</td>${profileFields.map(f=>`<td>${esc(profileValueDisplay(p.profileValues[f.field_key],f))}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${4+profileFields.length}">Aucun participant pour le moment.</td></tr>`}
  </table></div></section>`;
}
function profileValueDisplay(v,f){
  if(v==null||v==='')return '—';
  if(Array.isArray(v))return v.length?v.join(', '):'—';
  return String(v);
}
const SCALE_EMOJI=['😞','🙁','😐','🙂','😄'];
function scaleTone(i,count){return count<=1?3:Math.round(1+(i-1)/(count-1)*4)}
function renderScaleCards(min,max,labels){let n=max-min+1;return Array.from({length:n},(_,k)=>{let v=min+k,tone=scaleTone(k+1,n);return `<div class="scale-card sc-${tone}"><div class="sc-num">${v}</div><input class="sc-label-input" data-level="${v}" value="${esc(labels?.[v]??labels?.[String(v)]??'')}" placeholder="Libellé du niveau ${v}"><div class="sc-emoji">${SCALE_EMOJI[tone-1]}</div></div>`}).join('')}
function refreshScaleCards(){let el=document.querySelector('#scale-cards');el.style.setProperty('--scale-cols',scaleDraft.max-scaleDraft.min+1);el.innerHTML=renderScaleCards(scaleDraft.min,scaleDraft.max,scaleDraft.labels)}
function collectScaleLabels(){let labels={};document.querySelectorAll('.sc-label-input').forEach(inp=>{labels[inp.dataset.level]=inp.value});return labels}
function addScaleLevel(){if(scaleDraft.max-scaleDraft.min+1>=9)return notice('Le nombre maximum de niveaux est 9.');scaleDraft.labels=collectScaleLabels();scaleDraft.max+=1;refreshScaleCards()}
function removeScaleLevel(){if(scaleDraft.max-scaleDraft.min+1<=2)return notice('Il faut conserver au moins 2 niveaux.');scaleDraft.labels=collectScaleLabels();scaleDraft.max-=1;delete scaleDraft.labels[scaleDraft.max+1];refreshScaleCards()}
async function saveScale(e,id,tid){e.preventDefault();if(window.scaleLocked)return notice('Impossible de modifier l’échelle : cet atelier a déjà des réponses enregistrées.');let labels=collectScaleLabels();if(Object.values(labels).some(v=>!v.trim()))return notice('Veuillez renseigner un libellé pour chaque niveau.');let x=await api('/api/templates/'+tid,{method:'PUT',body:JSON.stringify({scale:{type:'numeric',min:scaleDraft.min,max:scaleDraft.max,labels}})});if(x.id!==tid){let cur=(await api('/api/sessions')).find(s=>s.id===id);await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({name:cur.name,organization:cur.organization,location:cur.location,date:cur.date,description:cur.description,expectedParticipants:cur.expected_participants,templateId:x.id})})}notice('Échelle de notation enregistrée.');configuration(id)}
async function pickQuestionnaire(id){let list=await api('/api/templates');shell('config','config','Changer de questionnaire','Choisissez le questionnaire à utiliser pour cet atelier');app.innerHTML=`<section class="card"><button class="secondary" onclick="configuration('${id}')">← Retour à la configuration</button><h2>Changer de questionnaire</h2><form onsubmit="applyQuestionnaire(event,'${id}')"><label>Questionnaire<select name="templateId">${list.map(t=>`<option value="${t.id}">${esc(t.name)} v${t.version}</option>`).join('')}</select></label><p class="text-meta">Changer de questionnaire remplace les domaines, indicateurs et l’échelle utilisés par cet atelier.</p><button>Utiliser ce questionnaire →</button></form></section>`}
async function applyQuestionnaire(e,id){e.preventDefault();if(window.scaleLocked)return notice('Impossible de changer de questionnaire : cet atelier a déjà des réponses enregistrées.');let f=Object.fromEntries(new FormData(e.target));let cur=(await api('/api/sessions')).find(s=>s.id===id);await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({name:cur.name,organization:cur.organization,location:cur.location,date:cur.date,description:cur.description,expectedParticipants:cur.expected_participants,templateId:f.templateId})});notice('Questionnaire mis à jour.');configuration(id)}
function createQuestionnaireForSession(id){shell('config','config','Nouveau questionnaire','Créer un questionnaire vierge pour cet atelier');app.innerHTML=`<section class="card"><button class="secondary" onclick="configuration('${id}')">← Retour à la configuration</button><h2>Nouveau questionnaire</h2><form onsubmit="makeQuestionnaireForSession(event,'${id}')"><label>Nom du questionnaire<input name="name" required autofocus placeholder="Ex : Diagnostic EPC / Atelier Dakar"></label><p class="text-meta">Vous pourrez ensuite ajouter des domaines et des indicateurs, ou importer une matrice XLSX.</p><button>Créer →</button></form></section>`}
async function makeQuestionnaireForSession(e,id){e.preventDefault();if(window.scaleLocked)return notice('Impossible de créer un nouveau questionnaire : cet atelier a déjà des réponses enregistrées.');let f=Object.fromEntries(new FormData(e.target));let x=await api('/api/templates',{method:'POST',body:JSON.stringify({name:f.name})});let cur=(await api('/api/sessions')).find(s=>s.id===id);await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({name:cur.name,organization:cur.organization,location:cur.location,date:cur.date,description:cur.description,expectedParticipants:cur.expected_participants,templateId:x.id})});notice('Nouveau questionnaire créé et associé à l’atelier. Ajoutez vos domaines et indicateurs.');edit(x.id)}
function updateParticipantsGauge(v){document.querySelector('#participants-gauge').innerHTML=gaugeSemi(+v,1,500,false);document.querySelector('#participants-number').value=v}
function aiConfigCardHtml(aiCfg){
  let providers=aiCfg.providers||{},curProvider=aiCfg.provider||Object.keys(providers)[0]||'',models=(providers[curProvider]||{}).models||[];
  return `<div class="cfg-grid-71"><div class="card"><div class="section-header"><h3>5 · Assistant IA</h3></div><p class="muted small">L’assistant IA est facultatif. EPC / SENEVAL fonctionne intégralement sans intelligence artificielle.</p><form onsubmit="saveAiConfig(event)"><label class="row" style="align-items:center;gap:.5rem"><input type="checkbox" name="enabled" style="width:auto" ${aiCfg.enabled?'checked':''} onchange="document.querySelector('#ai-fields').style.display=this.checked?'':'none'"> Activer l’assistant IA</label><div id="ai-fields" style="${aiCfg.enabled?'':'display:none'};margin-top:.6rem"><div class="row row-fields"><label>Fournisseur IA <span id="ai-provider-badge" class="ai-badge" data-pricing="${esc((providers[curProvider]||{}).pricing||'')}">${esc((providers[curProvider]||{}).pricing||'')}</span><select name="provider" id="ai-provider-select" onchange="updateAiModelOptions()">${Object.entries(providers).map(([k,v])=>`<option value="${k}" ${k===curProvider?'selected':''}>${esc(v.label)} — ${v.pricing}</option>`).join('')}</select></label><label>Modèle<select name="model" id="ai-model-select">${models.map(([mid,mlabel])=>`<option value="${esc(mid)}" ${mid===aiCfg.model?'selected':''}>${esc(mlabel)} (${esc(mid)})</option>`).join('')}</select></label></div><label>Clé API<div class="row" style="gap:.4rem"><input type="password" name="apiKey" id="ai-key-input" placeholder="${aiCfg.keyConfigured?'•••••••••••••••••••• (déjà configurée — laisser vide pour la conserver)':'Collez votre clé API ici'}" style="flex:1"><button type="button" class="secondary" onclick="toggleAiKeyVisibility()">👁</button><button type="button" class="secondary" onclick="aiKeyHelp()">?</button></div></label><div class="row" style="margin-top:.5rem;gap:.5rem;flex-wrap:wrap"><button class="secondary" type="submit">Enregistrer</button><button class="secondary" type="button" onclick="testAiConnection()">Tester l’assistant IA</button></div><p id="ai-test-status" class="text-meta" style="margin-top:.5rem"></p></div></form></div></div>`;
}
async function saveAiConfig(e){e.preventDefault();let f=Object.fromEntries(new FormData(e.target));let body={enabled:!!f.enabled,provider:f.provider,model:f.model};if(f.apiKey)body.apiKey=f.apiKey;await api('/api/ai/config',{method:'PUT',body:JSON.stringify(body)});notice('Configuration de l’assistant IA enregistrée.');configuration(window.currentSessionId)}
async function testAiConnection(){let status=document.querySelector('#ai-test-status');status.textContent='Test en cours…';status.style.color='';try{let r=await api('/api/ai/test',{method:'POST',body:'{}'});if(r.ok){status.textContent=`✓ Assistant IA opérationnel — ${r.provider} · ${r.model} · ${r.latencyMs} ms`;status.style.color='var(--green-700)'}else{status.textContent='⚠ '+r.reason;status.style.color='var(--red-600)'}}catch(e){status.textContent='⚠ '+e.message;status.style.color='var(--red-600)'}}
function toggleAiKeyVisibility(){let el=document.querySelector('#ai-key-input');el.type=el.type==='password'?'text':'password'}
function updateAiModelOptions(){let providerId=document.querySelector('#ai-provider-select').value,provider=(window.aiCfg.providers||{})[providerId]||{},models=provider.models||[];document.querySelector('#ai-model-select').innerHTML=models.map(([mid,mlabel])=>`<option value="${esc(mid)}">${esc(mlabel)} (${esc(mid)})</option>`).join('');let badge=document.querySelector('#ai-provider-badge');if(badge){badge.textContent=provider.pricing||'';badge.dataset.pricing=provider.pricing||''}}
function aiKeyHelp(){let providerId=document.querySelector('#ai-provider-select').value,p=(window.aiCfg.providers||{})[providerId];if(!p)return;let n=document.querySelector('#notice');if(!n){n=document.createElement('div');n.id='notice';document.body.appendChild(n)}n.className='notice';n.innerHTML=`<div><b>Obtenir une clé API ${esc(p.label)}</b><p>1. Ouvrez la page officielle du fournisseur.
2. Connectez-vous ou créez votre compte.
3. Créez une clé API.
4. Copiez-la.
5. Revenez ici et collez-la dans le champ Clé API.</p><div class="row" style="margin-top:.6rem;flex-wrap:wrap"><a href="${esc(p.keyUrl)}" target="_blank" rel="noopener"><button type="button" class="secondary">Ouvrir la page officielle de création de clé</button></a><button onclick="this.closest('.notice').remove()">Fermer</button></div></div>`}

// ==================================================
// Assistant IA — suggestions (facultatif, jamais de validation automatique)
// ==================================================
function aiConfirmConfidentiality(sid){
  if(localStorage.getItem('epc_ai_confirmed_'+sid))return Promise.resolve(true);
  return new Promise(resolve=>{
    let n=document.querySelector('#notice');if(!n){n=document.createElement('div');n.id='notice';document.body.appendChild(n)}
    n.className='notice';
    n.innerHTML=`<div><b>Confidentialité</b><p>L’utilisation de l’assistant IA transmet au fournisseur sélectionné les données nécessaires à l’analyse. Les réponses individuelles nominatives ne doivent pas être transmises.</p><div class="row" style="margin-top:.6rem;flex-wrap:wrap"><button class="secondary" id="ai-confirm-cancel">Annuler</button><button id="ai-confirm-ok">J’ai compris — utiliser l’assistant IA</button></div></div>`;
    document.querySelector('#ai-confirm-cancel').onclick=()=>{n.remove();resolve(false)};
    document.querySelector('#ai-confirm-ok').onclick=()=>{localStorage.setItem('epc_ai_confirmed_'+sid,'1');n.remove();resolve(true)};
  });
}
async function aiRun(sid,endpoint,body,onResult,btn){
  let ok=await aiConfirmConfidentiality(sid); if(!ok)return;
  let orig=btn?btn.textContent:null; if(btn){btn.disabled=true;btn.textContent='Génération en cours…'}
  try{ onResult(await api(endpoint,{method:'POST',body:JSON.stringify(body||{})})) }
  catch(e){ notice('Assistant IA momentanément indisponible. Vous pouvez poursuivre l’atelier normalement.\n\n('+e.message+')') }
  finally{ if(btn){btn.disabled=false;btn.textContent=orig} }
}
function aiButtonHtml(label,onclick,extraId){return `<button class="secondary ai-btn" ${extraId?`id="${extraId}"`:''} onclick="${onclick}">✦ ${esc(label)}</button>`}
function aiInfoCardHtml(boxId,text,regenFn){return `<div class="card ai-suggestion"><div class="section-header"><h3>✦ Lecture assistée par IA</h3></div><div style="white-space:pre-line">${esc(text)}</div><p class="text-meta" style="margin-top:.6rem">Analyse proposée par l’IA à partir des résultats EPC. Elle ne constitue pas une conclusion validée par le groupe.</p><div class="row" style="margin-top:.6rem;flex-wrap:wrap"><button class="secondary" onclick="${regenFn}">Régénérer</button><button class="ghost" onclick="document.querySelector('#${boxId}').innerHTML=''">Fermer</button></div></div>`}
function aiTextRetainCardHtml(boxId,text,retainFn,regenFn){return `<div class="card ai-suggestion"><div class="section-header"><h3>✦ Suggestion IA</h3></div><textarea class="ai-edit-area" id="${boxId}-text" style="min-height:120px;width:100%">${esc(text)}</textarea><p class="text-meta">Vous pouvez modifier ce texte avant de le retenir.</p><div class="row" style="margin-top:.6rem;flex-wrap:wrap"><button onclick="${retainFn}">Retenir</button><button class="secondary" onclick="${regenFn}">Régénérer</button><button class="ghost" onclick="document.querySelector('#${boxId}').innerHTML=''">Ignorer</button></div></div>`}
function aiListCardHtml(boxId,items,kindLabel,retainItemFn,regenFn){return `<div class="card ai-suggestion"><div class="section-header"><h3>✦ Hypothèses IA — ${esc(kindLabel)}</h3></div>${items.map((t,i)=>`<div class="ai-item"><textarea class="ai-edit-area" id="${boxId}-item-${i}" style="width:100%">${esc(t)}</textarea><div class="row" style="margin-top:.3rem"><button class="secondary" onclick="${retainItemFn}(${i},this)">Retenir cette hypothèse</button></div></div>`).join('')}<p class="text-meta" style="margin-top:.4rem">Hypothèses à discuter — aucune n’est présentée comme établie.</p><div class="row" style="margin-top:.6rem;flex-wrap:wrap"><button class="secondary" onclick="${regenFn}">Régénérer</button><button class="ghost" onclick="document.querySelector('#${boxId}').innerHTML=''">Ignorer tout</button></div></div>`}
async function saveConfigInfo(e,id){e.preventDefault();let f=Object.fromEntries(new FormData(e.target));let cur=(await api('/api/sessions')).find(x=>x.id===id);if(f.name!==cur.name&&window.configHasData){if(!confirm(`Attention : cet atelier contient déjà des réponses de participants.\n\nRenommer « ${cur.name} » en « ${f.name} » ne crée PAS un nouvel atelier : cela renomme celui-ci, avec toutes ses données existantes (participants, réponses, résultats).\n\nPour démarrer un atelier réellement vierge, utilisez plutôt « Nouveau diagnostic » depuis l’accueil.\n\nContinuer et renommer quand même ?`))return}await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({...f,expectedParticipants:cur.expected_participants})});notice('Informations enregistrées.');configuration(id)}
async function saveExpectedParticipants(id){let raw=document.querySelector('#participants-number').value;if(raw===''||raw==null)return notice('Veuillez indiquer un nombre de participants prévus avant d’enregistrer.');let v=+raw;if(!(v>0))return notice('Le nombre de participants prévus doit être supérieur à 0.');let cur=(await api('/api/sessions')).find(x=>x.id===id);await api('/api/sessions/'+id,{method:'PUT',body:JSON.stringify({name:cur.name,organization:cur.organization,location:cur.location,date:cur.date,description:cur.description,expectedParticipants:v})});notice('Nombre de participants prévus enregistré.');configuration(id)}
function launchCollecte(id,ready){
  if(!ready){notice('Impossible d’ouvrir la collecte\n\nLe questionnaire ne contient aucun indicateur.\n\n→ Modifiez le questionnaire avant de continuer.');return}
  moderator(id)
}
function stepperHtml(done){let labels=['Préparation','Collecte','Diagnostic','Priorités','Analyse','Recommandations','Rapport'];let cur=done.findIndex(x=>!x);return `<div class="stepper">${labels.map((l,i)=>{let cls=done[i]?'done':(i===cur?'current':'');return `<span class="step ${cls}"><span class="dot">${done[i]?'✓':i+1}</span>${l}</span>`}).join('')}</div>`}
// Standalone QR Code encoder (Byte mode, versions 1-6, EC levels L/M/Q/H) — ISO/IEC 18004.
// No external dependencies; verified by round-trip decode against a reference decoder.
