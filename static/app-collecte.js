const generateQR=(function(){
'use strict';
const GF_EXP=new Array(512),GF_LOG=new Array(256);
(function initGF(){let x=1;for(let i=0;i<255;i++){GF_EXP[i]=x;GF_LOG[x]=i;x<<=1;if(x&0x100)x^=0x11d}for(let i=255;i<512;i++)GF_EXP[i]=GF_EXP[i-255]})();
function gfMul(a,b){return(a===0||b===0)?0:GF_EXP[GF_LOG[a]+GF_LOG[b]]}
function rsGeneratorPoly(degree){let poly=[1];for(let i=0;i<degree;i++){const next=new Array(poly.length+1).fill(0);for(let j=0;j<poly.length;j++){next[j]^=gfMul(poly[j],1);next[j+1]^=gfMul(poly[j],GF_EXP[i])}poly=next}return poly}
function rsEncode(dataBytes,ecLen){const gen=rsGeneratorPoly(ecLen);const res=dataBytes.concat(new Array(ecLen).fill(0));for(let i=0;i<dataBytes.length;i++){const coef=res[i];if(coef===0)continue;for(let j=0;j<gen.length;j++)res[i+j]^=gfMul(gen[j],coef)}return res.slice(dataBytes.length)}
const EC_L=0,EC_M=1,EC_Q=2,EC_H=3;
// RS_BLOCK_TABLE[version-1][ecLevel]=[ecCodewordsPerBlock,blocks1,dataCw1,blocks2,dataCw2] (ISO/IEC 18004 Table 9, versions 1-10)
const RS_BLOCK_TABLE=[
[[7,1,19,0,0],[10,1,16,0,0],[13,1,13,0,0],[17,1,9,0,0]],
[[10,1,34,0,0],[16,1,28,0,0],[22,1,22,0,0],[28,1,16,0,0]],
[[15,1,55,0,0],[26,1,44,0,0],[18,2,17,0,0],[22,2,13,0,0]],
[[20,1,80,0,0],[18,2,32,0,0],[26,2,24,0,0],[16,4,9,0,0]],
[[26,1,108,0,0],[24,2,43,0,0],[18,2,15,2,16],[22,2,11,2,12]],
[[18,2,68,0,0],[16,4,27,0,0],[24,4,19,0,0],[28,4,15,0,0]],
];
const VERSION_SIZE=v=>17+4*v;
const ALIGNMENT_POS=[[],[6,18],[6,22],[6,26],[6,30],[6,34]];
function bchFormat(data){let d=data<<10;const g=0x537;while(bitLength(d)-bitLength(g)>=0)d^=g<<(bitLength(d)-bitLength(g));return(data<<10|d)^0x5412}
function bitLength(n){let l=0;while(n!==0){l++;n>>>=1}return l}
const EC_INDICATOR={0:0b01,1:0b00,2:0b11,3:0b10};
const MAX_VERSION=6;
function chooseVersion(byteLen,ec){for(let v=1;v<=MAX_VERSION;v++){const blocks=RS_BLOCK_TABLE[v-1][ec];const totalData=blocks[1]*blocks[2]+blocks[3]*blocks[4];if(byteLen<=totalData-2)return v}const b=RS_BLOCK_TABLE[MAX_VERSION-1][ec];throw new Error('texte trop long pour un QR code (max ~'+(b[1]*b[2]+b[3]*b[4]-2)+' octets)')}
function encodeBytesUtf8(str){const bytes=[];for(let i=0;i<str.length;i++){let c=str.codePointAt(i);if(c>0xFFFF)i++;if(c<0x80)bytes.push(c);else if(c<0x800)bytes.push(0xC0|(c>>6),0x80|(c&0x3F));else if(c<0x10000)bytes.push(0xE0|(c>>12),0x80|((c>>6)&0x3F),0x80|(c&0x3F));else bytes.push(0xF0|(c>>18),0x80|((c>>12)&0x3F),0x80|((c>>6)&0x3F),0x80|(c&0x3F))}return bytes}
function buildDataCodewords(byteData,version,ec){const blocks=RS_BLOCK_TABLE[version-1][ec];const totalDataCw=blocks[1]*blocks[2]+blocks[3]*blocks[4];const bits=[];function put(val,len){for(let i=len-1;i>=0;i--)bits.push((val>>>i)&1)}put(0b0100,4);put(byteData.length,8);for(const b of byteData)put(b,8);const totalBits=totalDataCw*8;for(let i=0;i<4&&bits.length<totalBits;i++)bits.push(0);while(bits.length%8!==0)bits.push(0);const padBytes=[0xEC,0x11];let pi=0;while(bits.length<totalBits){put(padBytes[pi%2],8);pi++}const dataCw=[];for(let i=0;i<bits.length;i+=8){let v=0;for(let j=0;j<8;j++)v=(v<<1)|bits[i+j];dataCw.push(v)}return dataCw}
function interleave(dataCw,version,ec){const[ecLen,b1,d1,b2,d2]=RS_BLOCK_TABLE[version-1][ec];const groups=[];let offset=0;for(let i=0;i<b1;i++){groups.push(dataCw.slice(offset,offset+d1));offset+=d1}for(let i=0;i<b2;i++){groups.push(dataCw.slice(offset,offset+d2));offset+=d2}const ecGroups=groups.map(g=>rsEncode(g,ecLen));const maxData=Math.max(d1,b2?d2:0);const result=[];for(let i=0;i<maxData;i++)for(const g of groups)if(i<g.length)result.push(g[i]);for(let i=0;i<ecLen;i++)for(const g of ecGroups)result.push(g[i]);return result}
function buildMatrix(version,ec,maskPattern,finalCodewords){
const N=VERSION_SIZE(version);const mat=Array.from({length:N},()=>new Array(N).fill(null));const isFn=Array.from({length:N},()=>new Array(N).fill(false));
function setFn(r,c,v){if(r>=0&&r<N&&c>=0&&c<N){mat[r][c]=v;isFn[r][c]=true}}
function placeFinder(r,c){for(let dr=-1;dr<=7;dr++)for(let dc=-1;dc<=7;dc++){const rr=r+dr,cc=c+dc;if(rr<0||rr>=N||cc<0||cc>=N)continue;const inRing=(dr>=0&&dr<=6&&dc>=0&&dc<=6)&&(dr===0||dr===6||dc===0||dc===6);const inCore=dr>=2&&dr<=4&&dc>=2&&dc<=4;const val=(dr>=0&&dr<=6&&dc>=0&&dc<=6)?(inRing||inCore):false;setFn(rr,cc,val)}}
placeFinder(0,0);placeFinder(0,N-7);placeFinder(N-7,0);
for(let i=8;i<N-8;i++){setFn(6,i,i%2===0);setFn(i,6,i%2===0)}
const pos=ALIGNMENT_POS[version-1];
for(const r of pos)for(const c of pos){if((r<=8&&c<=8)||(r<=8&&c>=N-9)||(r>=N-9&&c<=8))continue;for(let dr=-2;dr<=2;dr++)for(let dc=-2;dc<=2;dc++){const ring=Math.max(Math.abs(dr),Math.abs(dc));setFn(r+dr,c+dc,ring!==1)}}
setFn(4*version+9,8,true);
for(let i=0;i<=14;i++){if(i<6)setFn(i,8,false);else if(i<8)setFn(i+1,8,false);else setFn(N-15+i,8,false);if(i<8)setFn(8,N-i-1,false);else if(i<9)setFn(8,15-i-1+1,false);else setFn(8,15-i-1,false)}
function maskAt(r,c){switch(maskPattern){case 0:return(r+c)%2===0;case 1:return r%2===0;case 2:return c%3===0;case 3:return(r+c)%3===0;case 4:return(Math.floor(r/2)+Math.floor(c/3))%2===0;case 5:return(r*c)%2+(r*c)%3===0;case 6:return((r*c)%2+(r*c)%3)%2===0;case 7:return((r+c)%2+(r*c)%3)%2===0}return false}
const bits=[];for(const cw of finalCodewords)for(let i=7;i>=0;i--)bits.push((cw>>>i)&1);
let bitIdx=0,dir=-1,col=N-1;
while(col>0){if(col===6)col--;for(let i=0;i<N;i++){const row=dir===-1?N-1-i:i;for(const c of[col,col-1]){if(!isFn[row][c]){const bit=bitIdx<bits.length?bits[bitIdx]:0;bitIdx++;const m=maskAt(row,c);mat[row][c]=!!bit!==m}}}dir=-dir;col-=2}
const fmtData=(EC_INDICATOR[ec]<<3)|maskPattern;const fmtBits=bchFormat(fmtData);
for(let i=0;i<=14;i++){const mod=((fmtBits>>i)&1)===1;if(i<6)mat[i][8]=mod;else if(i<8)mat[i+1][8]=mod;else mat[N-15+i][8]=mod;if(i<8)mat[8][N-i-1]=mod;else if(i<9)mat[8][15-i-1+1]=mod;else mat[8][15-i-1]=mod}
mat[N-8][8]=true;
return mat}
function penaltyScore(mat){const N=mat.length;let score=0;
for(let r=0;r<N;r++){let runLen=1;for(let c=1;c<N;c++){if(mat[r][c]===mat[r][c-1])runLen++;else{if(runLen>=5)score+=3+(runLen-5);runLen=1}}if(runLen>=5)score+=3+(runLen-5)}
for(let c=0;c<N;c++){let runLen=1;for(let r=1;r<N;r++){if(mat[r][c]===mat[r-1][c])runLen++;else{if(runLen>=5)score+=3+(runLen-5);runLen=1}}if(runLen>=5)score+=3+(runLen-5)}
for(let r=0;r<N-1;r++)for(let c=0;c<N-1;c++){const v=mat[r][c];if(v===mat[r][c+1]&&v===mat[r+1][c]&&v===mat[r+1][c+1])score+=3}
const pattern=[true,false,true,true,true,false,true];
function matchAt(arr,i){for(let k=0;k<7;k++)if(arr[i+k]!==pattern[k])return false;return true}
for(let r=0;r<N;r++)for(let c=0;c<=N-7;c++){if(matchAt(mat[r].slice(c,c+7),0))score+=40}
for(let c=0;c<N;c++)for(let r=0;r<=N-7;r++){const col=[];for(let k=0;k<7;k++)col.push(mat[r+k][c]);if(matchAt(col,0))score+=40}
let dark=0;for(let r=0;r<N;r++)for(let c=0;c<N;c++)if(mat[r][c])dark++;
score+=Math.floor(Math.abs(Math.round((dark*100)/(N*N))-50)/5)*10;
return score}
return function generateQR(text,ecLevel){
const ec={L:EC_L,M:EC_M,Q:EC_Q,H:EC_H}[ecLevel||'M'];
const byteData=encodeBytesUtf8(text);
const version=chooseVersion(byteData.length,ec);
const dataCw=buildDataCodewords(byteData,version,ec);
const finalCw=interleave(dataCw,version,ec);
let best=null,bestScore=Infinity;
for(let mask=0;mask<8;mask++){const mat=buildMatrix(version,ec,mask,finalCw);const s=penaltyScore(mat);if(s<bestScore){bestScore=s;best=mat}}
return best.map(row=>row.map(v=>!!v))}
})();
function qrSvg(link){let mat;try{mat=generateQR(link,'M')}catch(e){return `<p class="error">${esc(e.message)}</p>`}let n=mat.length,quiet=4,size=n+2*quiet,rects='';for(let r=0;r<n;r++)for(let c=0;c<n;c++)if(mat[r][c])rects+=`<rect x="${c+quiet}" y="${r+quiet}" width="1" height="1"/>`;return `<svg viewBox="0 0 ${size} ${size}" shape-rendering="crispEdges">${rects}</svg>`}
function copyLink(link){navigator.clipboard.writeText(link).then(()=>notice('Lien copié.')).catch(()=>notice('Copie impossible : sélectionnez le lien.'))}
async function moderator(id){let [a,q]=await Promise.all([api('/api/sessions/'+id+'/analysis'),api('/api/sessions/'+id+'/qualitative-data')]);let link=location.origin+'/?session='+id;let done=[true,a.completedCount>0,q.priorities.length>0,q.analyses.some(x=>x.problem),q.entries.length>0,q.recommendations.length>0,false];let pct=pctOf(a.completedCount,a.session.expected_participants);window.currentSessionId=id;shell('collecte','collecte',a.session.name,'Tableau de bord de l’atelier',{id,name:a.session.name,location:a.session.location,date:a.session.date,pct:pct??0,statusLabel:a.session.status==='closed'?'Terminé':(pct==null?(a.participantCount?'Collecte en cours':'Préparation'):'Collecte '+pct+'%')});let primary=(!a.participantCount)?{l:'Ouvrir la collecte',a:`qr('${link}')`}:(a.completedCount<a.participantCount)?{l:'Afficher le QR',a:`qr('${link}')`}:(!q.priorities.length)?{l:'Voir le diagnostic',a:`diagnostic('${id}')`}:(!q.entries.length)?{l:'Analyser les priorités',a:`analysisPriorities('${id}')`}:(!q.recommendations.length)?{l:'Construire les recommandations',a:`recommendationsView('${id}')`}:{l:'Voir le rapport final',a:`finalReport('${id}')`};app.innerHTML=`<button class="ghost" onclick="load()">← Tous les ateliers</button><div class="card"><div class="section-header"><div><h2>${esc(a.session.name)}</h2><p class="text-meta">${sessionStatusBadge(a.session,a)}${a.session.date?' · '+esc(a.session.date):''}</p></div></div>${stepperHtml(done)}<div class="grid"><div class="metric"><span class="label">Commencés</span><b>${a.participantCount}</b></div><div class="metric"><span class="label">Validés</span><b>${a.completedCount}</b></div><div class="metric"><span class="label">Progression</span><b>${pctLabel(pct)}</b></div>${a.global.capacity!=null?`<div class="metric"><span class="label">Capacité globale</span><b>${fmt(a.global.capacity)}</b></div>`:''}${(a.global.consensus!=null||a.global.consensusNote)?`<div class="metric"><span class="label">Consensus global</span><b>${fmtConsensus(a.global)}</b></div>`:''}</div><div class="row" style="margin-top:1.1rem"><button onclick="${primary.a}">${primary.l}</button><button class="secondary" onclick="diagnostic('${id}')">Diagnostic</button><button class="ghost" onclick="dimensionAnalysisView('${id}')">Filtrer / Comparer</button><button class="ghost" onclick="preview('${id}')">Prévisualiser le questionnaire</button><button class="ghost" onclick="reportMetadata('${id}')">Informations de l’atelier</button><button class="ghost" onclick="finalReport('${id}')">Rapport final</button><button class="danger" onclick="removeSession('${id}','${esc(a.session.name)}')">Supprimer l’atelier</button></div></div><div class="card"><div class="section-header"><h3>Collecte</h3></div><div class="qr-panel sidebar-panel"><div class="qr">${qrSvg(link)}</div><div><p class="text-meta">Lien participant</p><div class="link-box"><span style="flex:1">${esc(link)}</span></div><div class="row" style="margin-top:.6rem"><button class="secondary" onclick="copyLink('${link}')">Copier le lien</button><button class="secondary" onclick="window.open('${link}','_blank')">Ouvrir participant</button><button class="ghost" onclick="qr('${link}')">Afficher en grand</button></div><p class="text-meta" style="margin-top:.6rem">${a.participantCount} commencé${a.participantCount>1?'s':''} · ${a.completedCount} validé${a.completedCount>1?'s':''}</p></div></div></div>`}
async function removeSession(id,name){if(!confirm('Supprimer définitivement l’atelier "'+name+'" ? Cette action est irréversible.'))return;await api('/api/sessions/'+id,{method:'DELETE'});notice('Atelier supprimé.');await load()}
const level=v=>v==null?'—':v<20?'':v<=39?'Loin en dessous de la moyenne':v<=59?'En dessous de la moyenne':v<=70?'Moyen':v<=80?'Au-dessus de la moyenne':'Bien au-dessus de la moyenne';
function radar(ds){let a=ds.filter(d=>d.capacity!=null),n=a.length;if(n<3)return '<p class="muted">Radar non disponible : au moins 3 domaines avec des données sont nécessaires.</p>';let cx=250,cy=235,r=170,p=(v,i)=>{let z=-Math.PI/2+i*2*Math.PI/n;return `${cx+Math.cos(z)*r*v/100},${cy+Math.sin(z)*r*v/100}`};return `<svg class="chart" viewBox="0 0 500 470"><text x="12" y="18" fill="#176b4b">■ Capacité</text><text x="120" y="18" fill="#536271">■ Consensus</text>${a.map((d,i)=>{let z=-Math.PI/2+i*2*Math.PI/n;return `<line x1="${cx}" y1="${cy}" x2="${cx+Math.cos(z)*r}" y2="${cy+Math.sin(z)*r}" stroke="#aab"/><text x="${cx+Math.cos(z)*(r+12)}" y="${cy+Math.sin(z)*(r+12)}" font-size="10">${esc(d.code||d.label).slice(0,9)}</text>`}).join('')}<polygon points="${a.map((d,i)=>p(d.capacity||0,i)).join(' ')}" fill="#176b4b55" stroke="#176b4b" stroke-width="3"/><polygon points="${a.map((d,i)=>p(d.consensus||0,i)).join(' ')}" fill="#53627144" stroke="#536271" stroke-width="3"/></svg>`}
async function diagnostic(id){let a=await api('/api/sessions/'+id+'/analysis'),q=await api('/api/sessions/'+id+'/qualitative-data'),g=a.global;window.currentSessionId=id;shell('resultats','traiter',a.session.name,'Diagnostic du niveau de capacité');app.innerHTML=`<section class="card"><button class="ghost" onclick="moderator('${id}')">← Retour à l’atelier</button><h2>Diagnostic terminé</h2><p class="muted">${esc(a.session.name)} · ${a.participantCount} commencés · ${a.completedCount} validés · taux de validation ${a.participantCount?Math.round(a.completedCount/a.participantCount*100):0}%</p><div class="grid"><div class="metric"><span class="label">Capacité</span><b>${fmt(g.capacity)}</b></div><div class="metric"><span class="label">Consensus</span><b>${fmtConsensus(g)}</b></div><div class="metric"><span class="label">Capacité graduée</span><b>${fmt(g.gradedCapacity)}</b></div><div class="metric"><span class="label">Consensus gradué</span><b>${fmt(g.gradedConsensus)}</b></div></div><div class="section-header" style="margin-top:1.6rem"><h3>Synthèse par domaine</h3></div><div class="table-wrap"><table><tr><th>Domaine</th><th>Capacité</th><th>Consensus</th><th>Graduées</th><th>Niveau</th><th>Réponses</th><th></th></tr>${a.domains.map(d=>`<tr><td>${esc(d.label)}</td><td>${fmt(d.capacity)}</td><td>${fmtConsensus(d)}</td><td>${d.gradedCapacity??'—'} / ${d.gradedConsensus??'—'}</td><td>${level(d.capacity)}</td><td>${d.responses}</td><td><button class="ghost" onclick="domainDiagnostic('${id}','${d.id}')">Détail</button></td></tr>`).join('')}</table></div><div class="section-header" style="margin-top:1.6rem"><h3>Priorités retenues</h3></div>${q.priorities.length?`<p>${q.priorities.length} priorité(s) sélectionnée(s).</p>`:'<p class="muted">Aucune priorité sélectionnée.</p>'}<p><button onclick="finalReport('${id}')">Synthèse finale</button></p></section>`}
async function dimensionAnalysisView(sessionId){
  window.currentSessionId=sessionId;
  let [dims,list]=await Promise.all([api('/api/sessions/'+sessionId+'/dimensions'),api('/api/sessions')]);
  let s=list.find(x=>x.id===sessionId);
  shell('resultats','traiter',s?s.name:'','Filtrer / comparer par profil');
  if(!dims.length){
    app.innerHTML=`<section class="card"><button class="ghost" onclick="moderator('${sessionId}')">← Retour à l’atelier</button><h2>Filtrer / comparer par profil</h2><p class="muted">Aucune dimension d’analyse configurée pour cet atelier.</p><p class="text-meta">Dans Configuration → Profil participant, marquez un champ à choix (unique ou multiple) comme « dimension d’analyse » pour pouvoir filtrer et comparer les résultats par ce champ (ex : rôle, structure).</p></section>`;
    return;
  }
  window.dimensionFields=dims;
  app.innerHTML=`<section class="card"><button class="ghost" onclick="moderator('${sessionId}')">← Retour à l’atelier</button><h2>Filtrer / comparer par profil</h2><p class="muted small">Compare les résultats du diagnostic entre sous-groupes de participants, selon un champ de profil marqué comme dimension. Les sous-groupes dont l’effectif est trop faible sont automatiquement masqués pour préserver l’anonymat (le seuil exact est rappelé après chaque comparaison).</p>
  <label>Dimension<select id="dim-select" onchange="renderDimensionOptions('${sessionId}')">${dims.map(d=>`<option value="${esc(d.fieldKey)}">${esc(d.label)}</option>`).join('')}</select></label>
  <div id="dim-options"></div>
  <div id="dim-result"></div>
  </section>`;
  renderDimensionOptions(sessionId);
}
function renderDimensionOptions(sessionId){
  let key=document.querySelector('#dim-select').value;
  let dim=window.dimensionFields.find(d=>d.fieldKey===key);
  document.querySelector('#dim-options').innerHTML=`<p class="text-meta" style="margin-top:.6rem">Sélectionnez une ou plusieurs valeurs à comparer :</p><form id="dim-values-form">${dim.options.map(o=>{let v=profileOptValue(o);return `<label class="checkbox-row"><input type="checkbox" name="v" value="${esc(v)}"> ${esc(profileOptLabel(o))}</label>`}).join('')}</form><button type="button" style="margin-top:.4rem" onclick="runDimensionFilter('${sessionId}')">Comparer</button>`;
  document.querySelector('#dim-result').innerHTML='';
}
async function runDimensionFilter(sessionId){
  let key=document.querySelector('#dim-select').value;
  let values=[...document.querySelectorAll('#dim-values-form input:checked')].map(x=>x.value);
  if(!values.length)return notice('Sélectionnez au moins une valeur à comparer.');
  let results;
  try{
    // One request compares every selected value (backend batches the
    // participant-matching query instead of re-running it once per value).
    let qs='dimension='+encodeURIComponent(key)+values.map(v=>'&value='+encodeURIComponent(v)).join('');
    results=(await api('/api/sessions/'+sessionId+'/analysis?'+qs)).results;
  }catch(err){return notice(err.message)}
  window.dimensionResults={key,values,results};
  let suppressedValues=results.filter(r=>r.dimension.suppressed).map(r=>r.dimension.value);
  // Résumé Catégorie|N|Capacité|Consensus|Graduations (mission de parite
  // :8810->:8820, cf. consignes_claude.txt), puis le graphique bars() sur ces
  // memes categories, puis le detail par domaine deja fourni par
  // comparisonTableHtml (partagee avec la comparaison de groupes/campagnes).
  let categoryRows=`<div class="table-wrap"><table><tr><th>Catégorie</th><th>N</th><th>Capacité</th><th>Consensus</th><th>Graduations</th></tr>${results.map(r=>`<tr><td>${esc(r.dimension.value)}</td><td>${r.completedCount}</td><td>${fmt(r.global.capacity)}</td><td>${fmtConsensus(r.global)}</td><td>${r.global.gradedCapacity??'—'} / ${r.global.gradedConsensus??'—'}</td></tr>`).join('')}</table></div>`;
  let categoryBars=results.map(r=>({label:r.dimension.value,code:r.dimension.value,capacity:r.global.capacity,consensus:r.global.consensus,gradedCapacity:r.global.gradedCapacity,gradedConsensus:r.global.gradedConsensus}));
  document.querySelector('#dim-result').innerHTML=categoryRows
    +bars(categoryBars,'standard')+bars(categoryBars,'graded')
    +`<h3 style="margin-top:1.2rem">Comparaison par domaine</h3>`
    +comparisonTableHtml(results,r=>`<th colspan="2">${esc(r.dimension.value)} (${r.completedCount} validé${r.completedCount>1?'s':''})<br><span class="text-meta">Capacité / Consensus</span></th>`)
    +(suppressedValues.length?`<p class="muted small" style="margin-top:.6rem">⚠ Effectif insuffisant (moins de ${results[0].dimension.minRequired} participants validés) pour : ${suppressedValues.map(v=>esc(v)).join(', ')}. Résultats masqués pour préserver l’anonymat.</p>`:'')
    +`<button class="secondary" style="margin-top:.6rem" onclick="exportDimensionCsv()">Exporter ce comparatif (CSV)</button>`;
}
function exportDimensionCsv(){
  let {values,results}=window.dimensionResults;
  let lines=[['Domaine',...values.flatMap(v=>[v+' - Capacité',v+' - Consensus'])].map(csvCell).join(';')];
  results[0].domains.forEach((d,di)=>{
    lines.push([d.label,...results.flatMap(r=>[r.domains[di]?.capacity??'',r.domains[di]?.consensus??''])].map(csvCell).join(';'));
  });
  let blob=new Blob(['﻿'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  let url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download='comparaison-profil.csv';a.click();
  URL.revokeObjectURL(url);
}
function csvCell(v){v=String(v);return /[;"\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v}
async function domainDiagnostic(sid,did){let a=await api('/api/sessions/'+sid+'/analysis'),q=await api('/api/sessions/'+sid+'/qualitative-data'),d=a.domains.find(x=>x.id===did),max=(await api('/api/templates/'+a.session.template_id)).priority.maxPerDomain;let selected=new Set(q.priorities.filter(p=>p.domain_id===did).map(p=>p.indicator_id));
  // Badge "À examiner en priorité" derive des constats automatiques deterministes
  // (objective_findings), jamais d'une selection arbitraire (mission de parite
  // :8810->:8820, cf. consignes_claude.txt) - le pilote reste seul decisionnaire.
  let flagged=new Set((a.findings?.fragilites?.indicators||[]).map(x=>x.id));
  app.innerHTML=`<section class="card"><button class="secondary" onclick="diagnostic('${sid}')">← Retour</button><h2>${esc(d.label)}</h2><p>Sélectionnez jusqu’à ${max} priorités : la décision reste humaine.</p><table><tr><th>Rang (plus faible en 1er) / Référence</th><th>Capacité</th><th>Consensus</th><th>Graduation</th><th>Distribution</th><th></th></tr>${[...d.indicators].sort((x,y)=>(x.capacity??999)-(y.capacity??999)).map((i,n)=>`<tr><td>${n+1}. ${esc(i.label)}${flagged.has(i.id)?' <span class="badge badge-danger" style="margin-left:.4rem">À examiner en priorité</span>':''}</td><td>${fmt(i.capacity)}</td><td>${fmtConsensus(i)}</td><td>${i.gradedCapacity??'—'} / ${i.gradedConsensus??'—'}</td><td>${Object.entries(i.distribution).map(x=>x.join(':')).join(' ')}</td><td>${selected.has(i.id)?`<button class="secondary" onclick="removePriority('${sid}','${did}','${i.id}')">Retirer</button>`:`<button onclick="addPriority('${sid}','${did}','${i.id}',${max},${selected.size})">Sélectionner</button>`}</td></tr>`).join('')}</table><p class="muted small" style="margin-top:1rem">Pour analyser les causes et construire des recommandations, sélectionnez vos priorités ci-dessus puis utilisez « Analyser les priorités » depuis l’écran Diagnostic.</p><button class="secondary" onclick="diagnostic('${sid}')">← Retour au diagnostic</button></section>`}
async function addPriority(s,d,i,max,count){if(count>=max)return notice('Le nombre maximum de priorités pour ce domaine est atteint. Retirez d’abord une priorité.');await api('/api/sessions/'+s+'/priorities',{method:'POST',body:JSON.stringify({domainId:d,indicatorId:i,votes:1})});domainDiagnostic(s,d)};async function removePriority(s,d,i){await api('/api/sessions/'+s+'/priorities/'+i,{method:'DELETE'});domainDiagnostic(s,d)};
function qr(link){document.querySelector('.app-shell').classList.add('participant-mode');app.innerHTML=`<section class="card">${back}<h2>Projection QR code</h2><p class="muted">Vue plein écran pour vidéoprojecteur. Lien local : non accessible depuis un téléphone hors réseau.</p><div class="qr" style="max-width:420px;margin:1.5rem auto">${qrSvg(link)}</div><div class="link-box" style="max-width:420px;margin:0 auto">${esc(link)}</div></section>`}
let participantName='';let participantAnonymous=false;let participantProfile=null,participantProfileValues={};
async function preview(sid){let a=await api('/api/sessions/'+sid+'/analysis'),t=await api('/api/templates/'+a.session.template_id);window.currentSessionId=sid;shell('questionnaire','config',a.session.name,'Aperçu du questionnaire');page(sid,null,t,a.session,{},0,true)};async function join(sid){
  document.querySelector('.app-shell').classList.add('participant-mode');participantName='';participantProfile=null;participantProfileValues={};
  let stored=null;try{stored=JSON.parse(localStorage.getItem(participantKey(sid))||'null')}catch(_){}
  participantAnonymous=!!(stored&&stored.anonymous);
  if(stored&&stored.completed)return participantDoneScreen(sid,stored.pid);
  let pid=stored&&stored.pid;
  if(!pid){let p=await api('/api/sessions/'+sid+'/participants',{method:'POST',body:'{}'});pid=p.id;try{localStorage.setItem(participantKey(sid),JSON.stringify({pid,completed:false,anonymous:participantAnonymous}))}catch(_){}}
  let d=await api('/api/participant?session='+sid+'&participant='+pid);
  if(!d.participant){try{localStorage.removeItem(participantKey(sid))}catch(_){}return join(sid)}
  if(d.participant.status==='completed'){try{localStorage.setItem(participantKey(sid),JSON.stringify({pid,completed:true,anonymous:participantAnonymous}))}catch(_){}return participantDoneScreen(sid,pid)}
  participantName=participantAnonymous?'':(d.participant.display_name||'');
  participantProfile=d.profile||null;participantProfileValues=d.profileValues||{};
  page(sid,pid,d.template,d.session,d.responses,resumeDomainIndex(d.template,d.responses),false);
}
function resumeDomainIndex(t,r){let ds=activeDomains(t);for(let n=0;n<ds.length;n++){if(ds[n].indicators.some(i=>r[i.id]===undefined))return n}return Math.max(0,ds.length-1)}
function scaleLegend(t){let min=+t.scale.min,max=+t.scale.max,labels=t.scale.labels||{},parts=Array.from({length:max-min+1},(_,k)=>max-k).map(v=>labels[v]?`${v} = ${esc(labels[v])}`:null).filter(Boolean);return parts.length?`<div class="scale-legend"><b>Échelle de notation :</b> ${parts.join(' | ')}</div>`:''}
async function setParticipantName(v){participantName=v;if(ctx&&ctx.pid&&!ctx.pre){try{await api('/api/participants/'+ctx.pid,{method:'PUT',body:JSON.stringify({displayName:v})})}catch(_){}}}
async function setAnonymous(v){participantAnonymous=v;if(v){await setParticipantName('')}try{localStorage.setItem(participantKey(ctx.sid),JSON.stringify({pid:ctx.pid,completed:false,anonymous:v}))}catch(_){}page(ctx.sid,ctx.pid,ctx.t,ctx.s,ctx.r,ctx.n,ctx.pre)}
function page(sid,pid,t,s,r,n,pre){let ds=t.domains.map(d=>({...d,indicators:d.indicators.filter(i=>i.active)})).filter(d=>d.indicators.length),d=ds[n],all=ds.flatMap(x=>x.indicators);if(!d)return notice('Impossible de continuer : ce questionnaire ne contient aucune question active.');window.ctx={sid,pid,t,s,r,n,pre};let answered=all.filter(i=>r[i.id]!==undefined).length,pct=all.length?Math.round(answered/all.length*100):0;app.innerHTML=`<section class="card">${pre?`<p class="preview">Mode prévisualisation</p><button class="ghost" onclick="moderator('${sid}')">← Retour au modérateur</button>`:''}<p class="participant-header">${esc(s.name)}</p>${scaleLegend(t)}${!pre&&n===0?`<label class="checkbox-row"><input type="checkbox" ${participantAnonymous?'checked':''} onchange="setAnonymous(this.checked)"> Participer anonymement</label>${participantAnonymous?'':`<label>Nom ou Organisation du Participant<input value="${esc(participantName)}" placeholder="Ex: M. Sow - Responsable Suivi-Évaluation" onchange="setParticipantName(this.value)"></label>`}`:''}${!pre&&n===0?profileFieldsHtml():''}<div class="domain-bar"><h2>${n+1}. ${esc(d.label)}</h2></div><p class="text-meta domain-meta">Domaine ${n+1} / ${ds.length} · ${answered}/${all.length} réponses</p><div class="progress" style="margin-bottom:1.2rem"><span style="width:${pct}%"></span></div>${d.indicators.map((i,qi)=>`<div class="q-row"><label>${n+1}.${qi+1} ${esc(i.label)}${i.required?' *':''}</label>${field(i,r[i.id],pre)}</div>`).join('')}<div class="row">${n?'<button class="secondary" onclick="nav(-1)">← Précédent</button>':(pre?'<button class="ghost" onclick="load()">← Retour à l’accueil</button>':'')}${n<ds.length-1?`<button class="block" onclick="${pre?'nav(1)':'nextDomain()'}">Domaine suivant →</button>`:`<button class="block" onclick="${pre?`moderator('${sid}')`:'finish()'}">${pre?'← Retour au modérateur':'Transmettre mes Réponses au Modérateur'}</button>`}</div></section>`}
function field(i,v,pre){if(i.response_type==='numeric'){let min=+ctx.t.scale.min,max=+ctx.t.scale.max,a=Array.from({length:max-min+1},(_,n)=>max-n);return `<div class="scale-options">${a.map(x=>`<label class="scale-opt"><input type="radio" name="f_${i.id}" value="${x}" ${String(x)===String(v)?'checked':''} onchange="answer('${i.id}',this.value,'numeric')"><span class="scale-num">${x}</span></label>`).join('')}</div>`}return `<input value="${esc(v||'')}" onchange="answer('${i.id}',this.value,'${i.response_type}')">`};function answer(i,v,type){ctx.r[i]=type==='numeric'?+v:v;let all=ctx.t.domains.flatMap(d=>d.indicators.filter(x=>x.active)),answered=all.filter(x=>ctx.r[x.id]!==undefined).length,pct=all.length?Math.round(answered/all.length*100):0,meta=app.querySelector('.domain-meta');if(meta)meta.textContent=meta.textContent.replace(/\d+\/\d+ réponses/,`${answered}/${all.length} réponses`);let bar=app.querySelector('.progress>span');if(bar)bar.style.width=pct+'%';if(!ctx.pre)api('/api/sessions/'+ctx.sid+'/responses',{method:'POST',body:JSON.stringify({participantId:ctx.pid,indicatorId:i,value:ctx.r[i],valueType:type})})};function nav(d){page(ctx.sid,ctx.pid,ctx.t,ctx.s,ctx.r,ctx.n+d,ctx.pre)};
function activeDomains(t){return t.domains.map(d=>({...d,indicators:d.indicators.filter(i=>i.active)})).filter(d=>d.indicators.length)}
function profileOptLabel(o){return typeof o==='object'&&o!==null?(o.label??o.value):o}
function profileOptValue(o){return typeof o==='object'&&o!==null?o.value:o}
function profileFieldInput(f,v){
  if(f.field_type==='number')return `<input type="number" value="${v==null?'':esc(v)}" onchange="answerProfile('${f.field_key}',this.value===''?null:+this.value)">`;
  if(f.field_type==='single_choice')return `<select onchange="answerProfile('${f.field_key}',this.value||null)"><option value="">—</option>${f.options.map(o=>{let ov=profileOptValue(o);return `<option value="${esc(ov)}" ${v!=null&&String(ov)===String(v)?'selected':''}>${esc(profileOptLabel(o))}</option>`}).join('')}</select>`;
  if(f.field_type==='multi_choice'){let arr=Array.isArray(v)?v:[];return `<div class="scale-options">${f.options.map(o=>{let ov=profileOptValue(o);return `<label class="checkbox-row"><input type="checkbox" ${arr.includes(ov)?'checked':''} onchange="toggleProfileMulti('${f.field_key}','${esc(ov)}',this.checked)"> ${esc(profileOptLabel(o))}</label>`}).join('')}</div>`}
  return `<input value="${v==null?'':esc(v)}" onchange="answerProfile('${f.field_key}',this.value)">`;
}
function profileFieldsHtml(){
  if(!participantProfile||!participantProfile.fields||!participantProfile.fields.filter(f=>f.active).length)return '';
  return `<div class="profile-fields">${participantProfile.fields.filter(f=>f.active).map(f=>{
    let labelText=`${esc(f.label)}${f.required?' *':''}`,input=profileFieldInput(f,participantProfileValues[f.field_key]);
    // multi_choice already wraps each option in its own <label><input>...; nesting
    // that inside an outer field-name <label> would let a click on the field name
    // itself implicitly activate the first checkbox (browsers associate a label
    // with its first labelable descendant) - keep the field-name label separate.
    if(f.field_type==='multi_choice')return `<div class="q-row"><label>${labelText}</label>${input}</div>`;
    return `<label>${labelText}${input}</label>`;
  }).join('')}</div>`;
}
async function answerProfile(key,value){participantProfileValues[key]=value;try{await api('/api/participants/'+ctx.pid+'/profile',{method:'POST',body:JSON.stringify({values:{[key]:value}})})}catch(_){}}
async function toggleProfileMulti(key,option,checked){let arr=Array.isArray(participantProfileValues[key])?participantProfileValues[key].slice():[];if(checked){if(!arr.includes(option))arr.push(option)}else{arr=arr.filter(x=>x!==option)}await answerProfile(key,arr)}
function nextDomain(){let ds=activeDomains(ctx.t),missing=ds[ctx.n].indicators.filter(i=>i.required&&ctx.r[i.id]===undefined);if(missing.length)return notice('Merci de répondre à toutes les questions de ce domaine avant de continuer.\n\nQuestion'+(missing.length>1?'s':'')+' sans réponse : '+missing.map(i=>i.label).join(', '));nav(1)}
function participantKey(sid){return 'epc_participant_'+sid}
async function finish(){
  let ds=activeDomains(ctx.t),missingIdx=ds.findIndex(d=>d.indicators.some(i=>i.required&&ctx.r[i.id]===undefined));
  if(missingIdx!==-1){notice('Votre questionnaire n’est pas encore complet. Merci de répondre à toutes les questions avant de le transmettre.');return page(ctx.sid,ctx.pid,ctx.t,ctx.s,ctx.r,missingIdx,false)}
  await api('/api/sessions/'+ctx.sid+'/complete',{method:'POST',body:JSON.stringify({participantId:ctx.pid})});
  try{localStorage.setItem(participantKey(ctx.sid),JSON.stringify({pid:ctx.pid,completed:true,anonymous:participantAnonymous}))}catch(_){}
  participantDoneScreen(ctx.sid,ctx.pid);
}
function participantDoneScreen(sid,pid){app.innerHTML=`<section class="card participant-done"><h2>Merci pour votre participation</h2><p>Votre questionnaire a bien été transmis au modérateur. Vous pouvez fermer cette page.</p>${sid&&pid?`<button class="secondary" onclick="myResponsesView('${sid}','${pid}')">Télécharger une copie de mes réponses</button>`:''}</section>`}
async function myResponsesView(sid,pid){
  let d=await api('/api/participant?session='+sid+'&participant='+pid),t=d.template,s=d.session,r=d.responses||{},labels=(t.scale&&t.scale.labels)||{};
  let ds=t.domains.map(dm=>({...dm,indicators:dm.indicators.filter(i=>i.active)})).filter(dm=>dm.indicators.length);
  app.innerHTML=`<article class="report"><div class="report-actions"><button class="secondary" onclick="participantDoneScreen('${sid}','${pid}')">← Retour</button><button onclick="window.print()">Imprimer / Enregistrer en PDF</button></div><section class="report-cover"><h1>Mes réponses</h1><h2>${esc(s.name)}</h2><p>${esc(s.date||'Date non renseignée')}</p></section>${ds.map(dm=>`<section><h2>${esc(dm.label)}</h2>${dm.indicators.map((i,n)=>{let v=r[i.id];let display=v==null||v===''?'Sans réponse':(i.response_type==='numeric'?`${esc(v)} — ${esc(labels[v]||'')}`:esc(v));return `<p><b>${n+1}. ${esc(i.label)}</b><br>${display}</p>`}).join('')}</section>`).join('')}</article>`
}
async function authGate(){
  document.querySelector('.app-shell').classList.remove('participant-mode');
  let st=await api('/api/auth/setup-status');
  if(st.needsSetup) return setupScreen();
  let me=await api('/api/auth/me');
  if(!me.user) return loginScreen();
  window.currentUser=me.user;
  load();
}
function authCard(title,subtitle,formHtml){
  document.querySelector('.app-shell').classList.add('participant-mode');
  app.innerHTML=`<section class="card" style="max-width:420px;margin:3rem auto"><h2>${esc(title)}</h2>${subtitle?`<p class="muted small">${esc(subtitle)}</p>`:''}${formHtml}</section>`;
}
function setupScreen(){
  authCard('Créer le premier compte','Ce compte administrateur récupère automatiquement les ateliers et questionnaires déjà présents.',
    `<form onsubmit="doSetup(event)"><label>Nom (optionnel)<input name="displayName"></label><label>Email<input name="email" type="email" required autofocus></label><label>Mot de passe (8 caractères minimum)<input name="password" type="password" required minlength="8"></label><button>Créer le compte</button></form>`)
}
async function doSetup(e){e.preventDefault();let f=Object.fromEntries(new FormData(e.target));try{let x=await api('/api/auth/setup',{method:'POST',body:JSON.stringify(f)});window.currentUser=x.user;authGate()}catch(err){notice(err.message)}}
function loginScreen(){
  authCard('Connexion','',
    `<form onsubmit="doLogin(event)"><label>Email<input name="email" type="email" required autofocus></label><label>Mot de passe<input name="password" type="password" required></label><button>Se connecter</button></form>`)
}
async function doLogin(e){e.preventDefault();let f=Object.fromEntries(new FormData(e.target));try{let x=await api('/api/auth/login',{method:'POST',body:JSON.stringify(f)});window.currentUser=x.user;authGate()}catch(err){notice(err.message)}}
async function doLogout(){await api('/api/auth/logout',{method:'POST',body:'{}'});location.href='/'}

