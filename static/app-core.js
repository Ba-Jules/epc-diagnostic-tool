const app=document.querySelector('#app'),esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])),fmt=v=>v==null?'—':Number(v).toFixed(1),fmtConsensus=o=>o&&o.consensusNote==='single_respondent'?'Non calculable':fmt(o?.consensus);
// Shared by runComparison() (groups) and runDimensionFilter() (profile
// dimensions) - both compare several analysis-shaped results side by side,
// domain by domain; kept as one renderer so the two screens can't quietly
// diverge on formatting (they already had before this was factored out).
function comparisonTableHtml(items,headerFn){let domainLabels=items[0].domains.map(d=>d.label);let rowsHtml=domainLabels.map((label,di)=>`<tr><td>${esc(label)}</td>${items.map(it=>`<td>${fmt(it.domains[di]?.capacity)}</td><td>${fmtConsensus(it.domains[di])}</td>`).join('')}</tr>`).join('');return `<div class="table-wrap" style="margin-top:1rem"><table><tr><th>Domaine</th>${items.map(headerFn).join('')}</tr>${rowsHtml}</table></div>`}
// Objectif est une cible de pilotage, jamais un plafond : sans objectif défini, on
// n'invente jamais un pourcentage de progression (on affiche '—').
const pctOf=(done,target)=>target?Math.round(done/target*100):null,pctLabel=p=>p==null?'—':p+'%'; function notice(message){let n=document.querySelector('#notice');if(!n){n=document.createElement('div');n.id='notice';document.body.appendChild(n)}n.className='notice';n.innerHTML=`<div><b>Information</b><p>${esc(message)}</p><button onclick="this.closest('.notice').remove()">Fermer</button></div>`} const api=(p,o={})=>fetch(p,{headers:{'Content-Type':'application/json'},...o}).then(async r=>{let x=await r.json();if(!r.ok)throw Error(x.error||'Action impossible.');return x}); window.addEventListener('unhandledrejection',e=>{e.preventDefault();notice(e.reason?.message||'Action impossible. Veuillez réessayer.');}); window.addEventListener('error',e=>{if(e.message)notice('Une action ne peut pas être réalisée : '+e.message)});let T=[],S=[];
async function load(){[T,S]=await Promise.all([api('/api/templates'),api('/api/sessions')]);shell('home','','Diagnostic EPC / SENEVAL','Préparation, collecte et restitution d’ateliers',null);home()};const back='<button class="ghost" onclick="load()">← Retour à l’accueil</button>';
window.currentSessionId=null;
const DOMAIN_COLORS=['var(--domain-1)','var(--domain-2)','var(--domain-3)','var(--domain-4)','var(--domain-5)','var(--domain-6)','var(--domain-7)'];
const SIDEBAR_ITEMS=[
  {key:'home',label:'Accueil',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 11l8-7 8 7"/><path d="M6 10v9a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-9"/></svg>',need:false,go:()=>load()},
  {key:'campaigns',label:'Campagnes',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M8 3v3M16 3v3"/></svg>',need:false,go:()=>campaignsHome()},
  {key:'config',label:'Configuration',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 13a7.7 7.7 0 0 0 0-2l2-1.5-2-3.4-2.3.9a7.6 7.6 0 0 0-1.7-1L15 3h-4l-.4 2.5a7.6 7.6 0 0 0-1.7 1l-2.3-.9-2 3.4L6.6 11a7.7 7.7 0 0 0 0 2l-2 1.5 2 3.4 2.3-.9c.5.4 1.1.8 1.7 1L10 21h4l.4-2.5c.6-.2 1.2-.6 1.7-1l2.3.9 2-3.4-2-1.5z"/></svg>',need:false,go:()=>window.currentSessionId?configuration(window.currentSessionId):notice('Créez ou ouvrez un atelier pour accéder à sa configuration.')},
  {key:'questionnaire',label:'Questionnaire',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3.5h8l4 4V20a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z"/><path d="M14 3.5V8h4"/><path d="M8.5 12.5h7M8.5 15.5h7"/></svg>',need:true,go:()=>preview(window.currentSessionId)},
  {key:'participants',label:'Participants',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.2"/><path d="M2.8 19c.6-3 3-5 6.2-5s5.6 2 6.2 5"/><circle cx="17.5" cy="8.5" r="2.6"/><path d="M15.5 13.6c2.6.4 4.4 2.2 4.9 4.9"/></svg>',need:true,go:()=>participantsList(window.currentSessionId)},
  {key:'collecte',label:'Collecte',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="7" height="7" rx="1"/><rect x="13" y="4" width="7" height="7" rx="1"/><rect x="4" y="13" width="7" height="7" rx="1"/><path d="M15 15h5M17.5 13v5"/></svg>',need:true,go:()=>moderator(window.currentSessionId)},
  {key:'resultats',label:'Résultats',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20V10M11 20V4M18 20v-7"/></svg>',need:true,go:()=>diagnostic(window.currentSessionId)},
  {key:'recommandations',label:'Recommandations',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3a6 6 0 0 0-3.5 10.9c.5.4.8 1 .8 1.6v.5h5.4v-.5c0-.6.3-1.2.8-1.6A6 6 0 0 0 12 3z"/><path d="M9.5 19h5M10.5 21.5h3"/></svg>',need:true,go:()=>recommendationsView(window.currentSessionId)},
  {key:'rapport',label:'Rapport',ic:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3.5h8l4 4V20a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z"/><path d="M8.5 11h7M8.5 14h7M8.5 17h4"/></svg>',need:true,go:()=>finalReport(window.currentSessionId)},
];
function renderSidebar(active){
  let cur=window.currentSessionSummary;
  return `<div class="sidebar-brand"><div class="logo"></div><div><b>EPC / SENEVAL</b><span>Diagnostic</span></div></div>
  <nav class="sidebar-nav">${SIDEBAR_ITEMS.map(it=>`<button class="${it.key===active?'active':''}" ${it.need&&!window.currentSessionId?'disabled':''} onclick="(${it.go.toString()})()"><span class="nav-ic">${it.ic}</span>${it.label}</button>`).join('')}</nav>
  ${cur?`<div class="sidebar-current"><div class="dot-row"><span class="dot"></span>Atelier actuel</div><div class="name">${esc(cur.name)}</div><div class="meta">${esc(cur.location||'—')}${cur.date?' · '+esc(cur.date):''}</div><div class="status-row"><span>Statut</span><span class="badge-mini">${esc(cur.statusLabel||'Préparation')}</span></div><div class="progress"><span style="width:${cur.pct??0}%"></span></div><button onclick="moderator('${cur.id}')">Voir le tableau de bord →</button></div>`:''}
  ${window.currentUser?`<div class="sidebar-user"><div class="name">${esc(window.currentUser.displayName||window.currentUser.email)}</div><button class="ghost" onclick="doLogout()">Se déconnecter</button></div>`:''}`;
}
function renderTopbar(title,subtitle,cycleStage){
  let stages=[['config','CONFIGURER'],['collecte','COLLECTER'],['traiter','TRAITER'],['analyse','ANALYSER'],['restitution','RESTITUER']];
  return `<div><h1>${esc(title||'')}</h1>${subtitle?`<p>${esc(subtitle)}</p>`:''}</div><div class="cycle-nav">${stages.map(([k,l])=>`<div class="cycle-step ${k===cycleStage?'active':''}"><span class="cycle-ic">${k===cycleStage?'●':'○'}</span>${l}</div>`).join('')}</div><button class="topbar-help" title="Aide" onclick="helpPanel()">?</button>`;
}
const HELP_SCREENS={
  home:{title:'Accueil',body:`
    <p>C’est le point de départ : la liste de tous vos ateliers de diagnostic.</p>
    <p class="help-callout"><b>Le parcours général :</b> CONFIGURER → COLLECTER → TRAITER → ANALYSER → RESTITUER.</p>
    <h3>Que faire ici ?</h3>
    <ol>
      <li><b>Nouveau diagnostic</b> pour créer un atelier.</li>
      <li><b>Reprendre</b> pour continuer votre atelier le plus récent.</li>
      <li><b>Ateliers récents</b> pour rouvrir n’importe quel atelier déjà créé.</li>
      <li><b>Modèles &amp; questionnaires</b> pour consulter ou dupliquer un questionnaire.</li>
    </ol>
    <p class="help-callout"><b>À retenir :</b> reprendre un atelier existant ne fait perdre aucune information déjà enregistrée.</p>
    <p><b>Étape suivante :</b> Configuration.</p>
    <h3>Espaces utilisateurs et campagnes</h3>
    <p>Chaque pilote dispose de son propre espace : il ne voit que ses ateliers, ses campagnes et ses questionnaires personnels. Les <b>Campagnes</b> (menu latéral) permettent d'organiser une collecte avec plusieurs groupes et relais — voir l'aide de l'écran Campagnes pour le détail.</p>`},
  config:{title:'Configuration de la mission',body:`
    <p>Cet écran sert à préparer la mission avant d’ouvrir la collecte auprès des participants.</p>
    <h3>Que faire, dans quel ordre ?</h3>
    <ol>
      <li><b>Informations de la mission</b> : nom, lieu, date, description.</li>
      <li><b>Participants prévus</b> : nombre indicatif, utilisé comme référence — différent du nombre réel de répondants, recalculé après la collecte.</li>
      <li><b>Questionnaire</b> : le modèle EPC / SENEVAL est utilisé par défaut. Vous pouvez le garder, en changer, ou modifier ses domaines/références/indicateurs.</li>
      <li><b>Échelle de notation</b> : renommez les niveaux ou ajustez leur nombre (2 à 9).</li>
      <li><b>Profil participant</b> (optionnel) : ajoutez des champs (rôle, structure, ancienneté…) que chaque participant renseigne en plus du questionnaire. Un champ à choix unique ou multiple peut être marqué comme « dimension d’analyse » pour filtrer et comparer les résultats par ce champ depuis l’écran Résultats.</li>
    </ol>
    <h3>Organisation de la collecte</h3>
    <p>À la création d'un diagnostic, choisissez <b>Collecte unique</b> (un seul lien/QR, fonctionnement habituel) ou <b>Plusieurs groupes / relais</b> (une campagne avec un QR et un suivi distincts par groupe — voir l'aide de l'écran Campagnes).</p>
    <p class="help-callout"><b>À retenir :</b> une fois que l’atelier a des réponses, le questionnaire et l’échelle sont verrouillés pour ne pas invalider les données déjà collectées — modifiez-les avant d’ouvrir la collecte.</p>
    <p><b>Étape suivante :</b> ouvrir la collecte.</p>
    <h3>Assistant IA — facultatif</h3>
    <p>L’assistant IA peut vous aider à interpréter les résultats, préparer l’analyse, explorer des hypothèses, formuler des recommandations et rédiger la synthèse du rapport. <b>Il ne modifie jamais les calculs EPC et ne décide pas à la place du groupe.</b></p>
    <ol>
      <li>Activer l’assistant IA.</li>
      <li>Choisir un fournisseur.</li>
      <li>Obtenir une clé API <span class="text-meta">(le « ? » à côté du champ Clé API explique comment, selon le fournisseur choisi)</span>.</li>
      <li>Coller la clé.</li>
      <li>Choisir le modèle recommandé.</li>
      <li>Tester la connexion.</li>
    </ol>
    <p><b>Badges</b> <span class="badge ai-badge" data-pricing="GRATUIT">GRATUIT</span> <span class="badge ai-badge" data-pricing="ESSAI">ESSAI</span> <span class="badge ai-badge" data-pricing="PAYANT">PAYANT</span> : ces indications concernent l’accès à l’API du fournisseur (pas son usage grand public) et peuvent évoluer.</p>
    <p class="help-callout"><b>À retenir :</b> la clé API est personnelle. Ne la partagez pas avec les participants.</p>
    <p>Lorsque vous utilisez une fonction IA, les données nécessaires à cette analyse sont transmises au fournisseur sélectionné. Les réponses individuelles nominatives et les informations techniques inutiles ne sont pas transmises. EPC / SENEVAL fonctionne normalement si l’IA est désactivée.</p>
    <p class="help-callout"><b>En cas de problème :</b> clé API refusée → vérifiez la clé, le fournisseur, le quota ou le modèle. Assistant IA indisponible → vous pouvez poursuivre l’atelier normalement, sans IA.</p>`},
  questionnaire:{title:'Questionnaire',body:`
    <p>Cet écran présente la structure du questionnaire utilisé pour l’atelier.</p>
    <h3>Vocabulaire</h3>
    <p><b>Domaine</b> : grande thématique (ex. Gestion des ressources humaines). <b>Référence</b> : identifiant court d’une question. <b>Indicateur qualitatif / capacité</b> : la question elle-même, notée par les participants selon l’échelle définie en Configuration.</p>
    <h3>Que faire ?</h3>
    <ul>
      <li>Modifier, ajouter ou supprimer un domaine ou un indicateur.</li>
      <li>Importer une matrice Excel déjà préparée.</li>
      <li>Télécharger la matrice pour la modifier hors ligne puis la réimporter.</li>
    </ul>
    <p class="help-callout"><b>À retenir :</b> la numérotation et l’ordre d’affichage sont gérés automatiquement.</p>
    <p class="help-callout"><b>En cas de problème :</b> questionnaire vide → revenez ici pour ajouter au moins un domaine avec une question.</p>`},
  collecte:{title:'Collecte',body:`
    <p>Cet écran sert à ouvrir la collecte des réponses et à en suivre l’avancement.</p>
    <h3>Que faire, dans quel ordre ?</h3>
    <ol>
      <li>Ouvrir la collecte.</li>
      <li>Afficher ou partager le QR code / le lien participant. Le bouton « Ouvrir participant » ouvre cette vue dans un nouvel onglet, sans fermer votre tableau de bord.</li>
      <li>Les participants répondent depuis leur téléphone.</li>
      <li>Suivre le nombre de participants et de questionnaires validés.</li>
      <li>Passer aux résultats lorsque la collecte est suffisante ou terminée.</li>
    </ol>
    <p class="help-callout"><b>À retenir :</b> les participants n’ont accès qu’au questionnaire — jamais à la configuration, ni à l’assistant IA.</p>
    <p class="help-callout"><b>En cas de problème :</b> QR inaccessible → vérifiez que l’adresse de l’application est bien accessible depuis le téléphone des participants.</p>
    <h3>Participant : anonymat et copie des réponses</h3>
    <p>À l'entrée du questionnaire, le participant peut cocher <b>Participer anonymement</b> : aucune identité n'apparaît alors dans les restitutions ni les exports, mais sa réponse reste techniquement rattachée à l'atelier (et au groupe, dans une campagne) pour permettre le calcul — anonyme ne veut pas dire non rattaché. Après validation, il peut <b>télécharger une copie de ses propres réponses</b>, sans jamais voir celles des autres participants.</p>`},
  resultats:{title:'Résultats / Diagnostic',body:`
    <p><b>Capacité</b> = niveau obtenu à partir des réponses. <b>Consensus</b> = degré de convergence entre les réponses des participants.</p>
    <p class="help-callout"><b>Avec un seul répondant, le consensus n'est pas calculable.</b> La capacité reste calculable dès la première réponse.</p>
    <h3>Comment lire les deux ensemble ?</h3>
    <p class="help-callout">Capacité faible + consensus élevé → faiblesse largement partagée.</p>
    <p class="help-callout">Capacité faible + consensus faible → faiblesse possible mais perceptions divergentes.</p>
    <p class="help-callout">Capacité élevée + consensus élevé → force largement reconnue.</p>
    <p class="help-callout">Capacité élevée + consensus faible → situation globalement favorable, mais vécue différemment selon les participants.</p>
    <p>Si l’assistant IA est activé, il peut proposer une lecture de ces résultats (bouton « ✦ Analyser avec l’IA »).</p>
    <h3>Filtrer / comparer par profil</h3>
    <p>Si un profil participant avec au moins une « dimension d’analyse » est configuré (voir l’aide de l’écran Configuration), le bouton « Filtrer / Comparer » permet de comparer les résultats entre sous-groupes (ex : par rôle, par structure).</p>
    <p class="help-callout"><b>À retenir :</b> pour préserver l’anonymat, les résultats d’un sous-groupe trop petit (moins de 5 participants validés) sont automatiquement masqués, quel que soit le champ choisi.</p>
    <h3>Priorités et analyse</h3>
    <p>Les priorités sont choisies par les participants et le modérateur selon le processus EPC. Pour chaque priorité retenue, vous documentez un constat, puis des causes, conséquences et leviers.</p>
    <p>L’IA peut aider à reformuler le constat, préparer des questions, suggérer des hypothèses de causes, de conséquences, ou identifier des leviers.</p>
    <p class="help-callout"><b>À retenir :</b> une suggestion IA n’est jamais une cause validée. Elle doit être discutée et retenue explicitement par le groupe.</p>
    <p class="help-callout"><b>En cas de problème :</b> pas de résultats → vérifiez qu’au moins un questionnaire a été validé par un participant.</p>`},
  recommandations:{title:'Recommandations',body:`
    <p>Cet écran construit la chaîne qui relie les résultats aux actions à mener :</p>
    <p class="help-callout">PRIORITÉ → CONSTAT → CAUSES → LEVIERS → RECOMMANDATIONS</p>
    <p>L’IA peut proposer des pistes de recommandations fondées sur les causes et leviers déjà retenus, mais le groupe et le modérateur conservent la validation finale.</p>
    <h3>Formations et plan d’action</h3>
    <p>Depuis cet écran, accédez aussi aux <b>besoins / thèmes de formation</b> (repris automatiquement des recommandations de catégorie Formation) et au <b>plan d’action</b> (tableau des recommandations retenues, avec responsable et échéance).</p>
    <p><b>Étape suivante :</b> Rapport final.</p>`},
  rapport:{title:'Rapport final',body:`
    <p>Le rapport rassemble automatiquement : méthodologie, participation, résultats EPC, graphiques, priorités, analyses, recommandations, formations, plan d’action et annexes.</p>
    <p>Téléchargez-le en Word, Excel, ou PDF (impression du rapport web).</p>
    <h3>✦ Assistance IA au rapport</h3>
    <p>L’IA peut préparer une proposition de rédaction à partir des résultats et analyses déjà validés. Elle ne recalcule pas les résultats et ne doit pas inventer les informations manquantes. Vous pouvez retenir, modifier, régénérer ou ignorer chaque proposition.</p>
    <p class="help-callout"><b>DONNÉES EPC</b> → calculées automatiquement. <b>TEXTE IA</b> → proposition rédactionnelle soumise à votre validation.</p>`},
  campaigns:{title:'Campagnes, groupes et relais',body:`
    <p>Une <b>campagne</b> organise une collecte avec plusieurs <b>groupes</b> (ex. un par structure ou par site), chacun suivi par un <b>relais</b>. Tous les groupes d'une campagne partagent le même questionnaire, ce qui permet de les comparer puis de les consolider.</p>
    <h3>Rôle du pilote</h3>
    <p>Le pilote crée la campagne, ajoute les groupes, distribue les liens/QR, suit l'avancement en temps réel et lance comparaison ou consolidation. Chaque groupe reçoit automatiquement son propre QR et son propre lien de suivi — le participant ne choisit jamais son groupe, il ne peut donc pas se tromper.</p>
    <h3>Groupes</h3>
    <p>Chaque groupe a un code court et une couleur qui servent uniquement de repère visuel — l'identification réelle se fait en interne par un identifiant unique. Un groupe sans réponse se supprime directement ; un groupe avec des réponses demande une confirmation renforcée avant suppression définitive.</p>
    <h3>Objectif de participants</h3>
    <p>L'objectif (généralement 20 à 30 par groupe, mais librement modifiable) est une cible de pilotage, jamais une limite de collecte : la collecte n'est jamais bloquée, même au-delà de l'objectif. Sans objectif défini, aucun pourcentage de progression n'est affiché (les compteurs Commencés/Validés restent, eux, toujours visibles).</p>
    <h3>Relais</h3>
    <p>Le relais reçoit un lien minimal (aucune configuration, aucun accès aux autres groupes, aucun accès à l'IA, aucune suppression, pas de détail des réponses) : nom de la campagne et du groupe, objectif, commencés, validés, progression, et son QR code. Le lien peut être régénéré si besoin — l'ancien cesse alors de fonctionner.</p>
    <h3>Comparer</h3>
    <p>Affiche les résultats (capacité, consensus) de plusieurs groupes côte à côte, domaine par domaine.</p>
    <h3>Consolider</h3>
    <p class="help-callout">L'outil ne fait pas la moyenne des résultats des groupes. Il regroupe leurs réponses individuelles puis recalcule le diagnostic EPC — 10 groupes de 20 participants donnent un diagnostic consolidé sur les ~200 personnes, pas une moyenne de 10 scores.</p>
    <p>La consolidation n'est possible qu'entre groupes utilisant exactement le même questionnaire (mêmes domaines, indicateurs, échelle et version) ; sinon elle est bloquée pour éviter tout résultat trompeur.</p>
    <h3>Supprimer une campagne</h3>
    <p>Un pilote peut supprimer définitivement une de ses propres campagnes, même si elle contient déjà des réponses. La suppression est irréversible : elle détruit groupes, participants, réponses, analyses et recommandations de cette campagne — jamais le questionnaire partagé, jamais les données d'une autre campagne ou d'un autre pilote. Une confirmation renforcée (saisie du mot « SUPPRIMER ») est obligatoire avant toute suppression définitive.</p>`},
  participants:{title:'Participants',body:`
    <p>Liste les participants d'un atelier avec leur statut (en cours / validé), leurs horodatages, et — si un profil participant est configuré (voir l'aide de l'écran Configuration) — leurs réponses au profil en colonnes supplémentaires.</p>
    <p class="help-callout"><b>À retenir :</b> cette liste montre les réponses de profil individuelles à l'équipe qui pilote l'atelier ; elle n'est jamais accessible au relais ni exportée avec les résultats individuels du diagnostic.</p>
    <p class="help-callout"><b>En cas de problème :</b> aucune colonne de profil affichée → vérifiez qu'un profil participant est bien rattaché à cet atelier dans Configuration.</p>`},
};
function helpPanel(){
  let key=window.currentHelpKey||'home',screen=HELP_SCREENS[key]||HELP_SCREENS.home;
  let n=document.querySelector('#notice');if(!n){n=document.createElement('div');n.id='notice';document.body.appendChild(n)}
  n.className='help-modal';
  n.innerHTML=`<div>
    <h2>Aide — ${esc(screen.title)}</h2>
    <p class="text-meta">Outil de diagnostic EPC / SENEVAL · le bouton « ? » ouvre toujours l’aide de l’écran où vous êtes.</p>
    ${screen.body}
    <button class="secondary help-close" onclick="this.closest('.help-modal').remove()">Fermer</button>
  </div>`;
}
function shell(active,cycleStage,title,subtitle,sessionSummary){
  document.querySelector('.app-shell')?.classList.remove('participant-mode');
  window.currentHelpKey=active;
  if(sessionSummary!==undefined)window.currentSessionSummary=sessionSummary;
  document.querySelector('#sidebar').innerHTML=renderSidebar(active);
  document.querySelector('#topbar').innerHTML=renderTopbar(title,subtitle,cycleStage);
}
function gaugeCircle(pct,size=150){
  let r=size*.36,c=2*Math.PI*r,cx=size/2,cy=size/2,off=c*(1-Math.max(0,Math.min(100,pct))/100);
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--gray-100)" stroke-width="12"/><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--accent-600)" stroke-width="12" stroke-linecap="round" stroke-dasharray="${c.toFixed(1)}" stroke-dashoffset="${off.toFixed(1)}" transform="rotate(-90 ${cx} ${cy})"/><text x="${cx}" y="${cy-3}" text-anchor="middle" class="gauge-value">${pct}<tspan class="gauge-suffix">%</tspan></text><text x="${cx}" y="${cy+20}" text-anchor="middle" class="gauge-label">${pct>=100?'PRÊT':'À COMPLÉTER'}</text></svg>`;
}
function gaugeSemi(value,min,max,unset){
  let v=value??min,pct=unset?0:Math.max(0,Math.min(1,(v-min)/(max-min||1))),r=72,cx=90,cy=92;
  let ang=Math.PI-Math.PI*pct,x2=cx+r*Math.cos(ang),y2=cy-r*Math.sin(ang),largeArc=pct>0.5?1:0;
  return `<svg width="180" height="112" viewBox="0 0 180 112"><path d="M ${cx-r} ${cy} A ${r} ${r} 0 0 1 ${cx+r} ${cy}" fill="none" stroke="var(--gray-100)" stroke-width="14" stroke-linecap="round"/>${pct>0?`<path d="M ${cx-r} ${cy} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2}" fill="none" stroke="var(--accent-600)" stroke-width="14" stroke-linecap="round"/>`:''}<text x="90" y="76" text-anchor="middle" font-size="${unset?18:30}" font-weight="800" fill="${unset?'var(--gray-400)':'var(--gray-900)'}">${unset?'Non défini':v}</text><text x="10" y="108" font-size="10" fill="var(--gray-400)" font-weight="700">${min}</text><text x="170" y="108" text-anchor="end" font-size="10" fill="var(--gray-400)" font-weight="700">${max}</text></svg>`;
}
function sessionStatusBadge(s,a){if(s.status==='closed')return /historique/i.test(s.name)?'<span class="badge badge-info">Historique</span>':'<span class="badge badge-success">Terminé</span>';if(/\btest\b/i.test(s.name))return '<span class="badge badge-neutral">Test</span>';if(!a||!a.participantCount)return '<span class="badge badge-neutral">Nouveau</span>';return '<span class="badge badge-progress">En cours</span>'}
function emptyState(title,text,btnLabel,btnAction){return `<div class="empty-state"><h3>${esc(title)}</h3><p>${esc(text||'')}</p>${btnLabel?`<button onclick="${btnAction}">${esc(btnLabel)}</button>`:''}</div>`}
const ICON_LAYERS='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>';
const ICON_CLOCK='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3.2 2"/></svg>';
const ICON_CHECK='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="8.5"/><path d="M8.3 12.3l2.6 2.6 5-5.4"/></svg>';
const ICON_DOC='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3.5h8l4 4V20a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z"/><path d="M14 3.5V8h4"/><path d="M8.5 12.5h7M8.5 15.5h7"/></svg>';
function statCard(tone,icon,num,label){return `<div class="stat-card"><div class="stat-icon tone-${tone}">${icon}</div><div class="stat-body"><div class="stat-num">${num}</div><div class="stat-label">${esc(label)}</div></div></div>`}
function sessionRow(s,a){let pct=a?pctOf(a.completedCount,s.expected_participants):null;return `<div class="session-row" onclick="moderator('${s.id}')"><div class="session-main"><div class="name">${esc(s.name)}</div><div class="meta">${sessionStatusBadge(s,a)}<span>${a?a.participantCount+' commencé'+(a.participantCount>1?'s':''):'—'}</span>${s.date?`<span>${esc(s.date)}</span>`:''}</div></div><div class="progress-wrap"><div class="progress"><span style="width:${pct??0}%"></span></div><div class="pct">${pctLabel(pct)}</div></div><div class="row-action" style="display:flex;gap:.4rem"><button class="ghost" onclick="event.stopPropagation();moderator('${s.id}')">Ouvrir →</button><button class="danger" onclick="event.stopPropagation();removeSession('${s.id}','${esc(s.name)}')">Supprimer</button></div></div>`}
let homeShowAll=false;
async function home(){let summaries=await Promise.all(S.map(s=>api('/api/sessions/'+s.id+'/analysis').catch(()=>null)));let enriched=S.map((s,i)=>({s,a:summaries[i]}));let total=enriched.length;let done=enriched.filter(x=>x.s.status==='closed').length;let inProgress=enriched.filter(x=>x.s.status==='open'&&x.a&&x.a.participantCount>0).length;let mostRecent=enriched[0];let visible=homeShowAll?enriched:enriched.slice(0,5);app.innerHTML=`<div class="hero-lead"><h2>Que souhaitez-vous faire ?</h2></div><div class="action-grid"><div class="action-card" onclick="newSession()"><div class="icon">＋</div><h3>Nouveau diagnostic</h3><p>Créer un nouvel atelier de diagnostic</p><button onclick="event.stopPropagation();newSession()">Démarrer</button></div>${mostRecent?`<div class="action-card" onclick="moderator('${mostRecent.s.id}')"><div class="icon">↻</div><h3>Reprendre</h3><p>${esc(mostRecent.s.name)}<br>${sessionStatusBadge(mostRecent.s,mostRecent.a)} ${mostRecent.a?`· ${mostRecent.a.completedCount} validé${mostRecent.a.completedCount>1?'s':''} sur ${mostRecent.a.participantCount} commencé${mostRecent.a.participantCount>1?'s':''}`:''}</p><button class="secondary" onclick="event.stopPropagation();moderator('${mostRecent.s.id}')">Continuer</button></div>`:''}</div><div class="stats-grid">${statCard('blue',ICON_LAYERS,total,'Ateliers au total')}${statCard('orange',ICON_CLOCK,inProgress,'En cours')}${statCard('green',ICON_CHECK,done,'Terminés')}${statCard('neutral',ICON_DOC,T.length,'Modèles disponibles')}</div><div class="card"><div class="section-header"><h2>Ateliers récents</h2>${enriched.length>5?`<button class="ghost" onclick="homeShowAll=!homeShowAll;home()">${homeShowAll?'Réduire':'Voir tous les ateliers'}</button>`:''}</div>${enriched.length?`<div class="session-list">${visible.map(({s,a})=>sessionRow(s,a)).join('')}</div>`:emptyState('Aucun diagnostic pour le moment','Créez votre premier atelier pour commencer.','Nouveau diagnostic','newSession()')}</div><div class="card"><div class="section-header"><h3>Modèles &amp; questionnaires</h3></div>${T.length?T.map(t=>`<div class="row"><div style="flex:2;min-width:200px"><b>${esc(t.name)}</b><div class="text-meta">Modèle disponible · v${t.version}</div></div><button class="ghost" onclick="newSession('${t.id}')">Utiliser</button><button class="ghost" onclick="dup('${t.id}')">Dupliquer</button><button class="ghost" onclick="edit('${t.id}')">Modifier</button></div>`).join(''):emptyState('Aucun modèle disponible','Importez un questionnaire pour commencer.')}<div class="row" style="margin-top:.6rem"><button class="secondary" onclick="importer()">Importer un questionnaire</button><a href="/api/templates/matrix.xlsx"><button type="button" class="secondary">Créer à partir d’un modèle Excel</button></a><button class="ghost" onclick="models()">Voir tous les modèles</button></div><p class="text-meta" style="margin-top:.4rem">Pour créer un questionnaire sur mesure : Télécharger la matrice → compléter → importer.</p></div>`}
