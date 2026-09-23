const DATA_FILE = 'books.xlsx';
const state = { books: [], route: 'home', filters: { status: 'Все', format: 'Все', q: '', sort: 'recent' } };

const $ = (s, root=document) => root.querySelector(s);
const $$ = (s, root=document) => [...root.querySelectorAll(s)];
const esc = (s='') => String(s).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

const statusMap = {
  'Я все прочитал!':'Прочитано','Бросил читать':'Брошено','Чтение':'Читаю','Пауза':'На паузе',
  'Прочитано':'Прочитано','Брошено':'Брошено','Читаю':'Читаю','На паузе':'На паузе','К прочтению':'К прочтению'
};

function excelDate(v){
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') {
    const p = XLSX.SSF.parse_date_code(v);
    return p ? new Date(p.y, p.m-1, p.d) : null;
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function text(v, fallback=''){
  if(v===null || v===undefined) return fallback;
  const s=String(v).trim();
  return s || fallback;
}
function num(v){ if(v===null||v===undefined||v==='') return null; const n=Number(String(v).replace(',','.')); return Number.isFinite(n)?n:null; }
function fmtDate(d){ return d ? new Intl.DateTimeFormat('ru-RU',{day:'numeric',month:'long',year:'numeric'}).format(d) : '—'; }
function fmtShortDate(d){ return d ? new Intl.DateTimeFormat('ru-RU',{day:'numeric',month:'long'}).format(d) : ''; }
function oneDec(n){ return Number.isInteger(n) ? String(n) : n.toFixed(1).replace('.',','); }
function ratingStars(n){ if(n===null) return ''; const full=Math.max(0,Math.min(5,Math.round(n))); return '★'.repeat(full)+'☆'.repeat(5-full); }
function initials(title=''){ return title.split(/\s+/).slice(0,2).map(x=>x[0]||'').join('').toUpperCase(); }
function coverHTML(book, badge=true){
  const url = book.cover && book.cover.replaceAll('&amp;','&');
  return `<div class="cover-wrap">${url?`<img src="${esc(url)}" alt="Обложка ${esc(book.title)}" loading="lazy" onerror="this.remove();this.parentElement.querySelector('.cover-fallback').style.display='grid'">`:''}<div class="cover-fallback" style="${url?'display:none':''}">${esc(initials(book.title))}</div>${badge && book.format==='Аудиокнига'?'<span class="format-badge">🎧</span>':''}</div>`;
}
function normalizeRow(r,i){
  const format=text(r['Формат']);
  const pages=num(r['Страницы']); const hours=num(r['Часы']);
  const pagesRead=num(r['Прочитано страниц']); const hoursRead=num(r['Прослушано часов']);
  const rawStatus=text(r['Статус']);
  const status=statusMap[rawStatus] || rawStatus;
  let progress=null, progressText='';
  if(status==='Прочитано') progress=100;
  else if(format==='Аудиокнига' && hours && hoursRead!==null){ progress=Math.min(100, Math.max(0,hoursRead/hours*100)); progressText=`${oneDec(hoursRead)} из ${oneDec(hours)} ч`; }
  else if(format!=='Аудиокнига' && pages && pagesRead!==null){ progress=Math.min(100, Math.max(0,pagesRead/pages*100)); progressText=`${pagesRead} из ${pages} стр.`; }
  if(status==='Прочитано' && !progressText){ progressText=format==='Аудиокнига'&&hours?`${oneDec(hours)} ч`:pages?`${pages} стр.`:''; }
  return {
    id:i, title:text(r['Название'],'Без названия'), author:text(r['Автор'],'Автор не указан'),
    series:text(r['Серия']), seriesNo:num(r['Книга в серии']), status, format,
    pages, hours, pagesRead, hoursRead,
    start:excelDate(r['Начало']), end:excelDate(r['Конец']), duration:num(r['Длительность чтения, дней'] ?? r['Общее время прочтения']),
    rating:num(r['Оценка']), cover:text(r['Обложка']), progress, progressText
  };
}
async function loadBooks(){
  const res=await fetch(`${DATA_FILE}?v=${Date.now()}`);
  if(!res.ok) throw new Error(`Не удалось загрузить ${DATA_FILE}`);
  const ab=await res.arrayBuffer();
  const wb=XLSX.read(ab,{type:'array',cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
  state.books=rows.map(normalizeRow).filter(b=>b.title && b.title!=='Без названия');
}
function currentYear(){ return new Date().getFullYear(); }
function byEndDesc(a,b){ return (b.end?.getTime()||0)-(a.end?.getTime()||0); }
function readThisYear(){ const y=currentYear(); return state.books.filter(b=>b.status==='Прочитано' && b.end?.getFullYear()===y); }
function avg(arr){ const x=arr.filter(n=>n!==null); return x.length?x.reduce((a,b)=>a+b,0)/x.length:null; }
function progressBlock(b){
  if(b.progress===null) return `<div class="start-date">В процессе${b.start?` · начато ${esc(fmtShortDate(b.start))}`:''}</div>`;
  return `<div class="progress-copy"><span>${esc(b.progressText||'Прогресс')}</span><strong>${Math.round(b.progress)}%</strong></div><div class="progress"><span style="width:${b.progress}%"></span></div>${b.start?`<div class="start-date">Начато ${esc(fmtShortDate(b.start))}</div>`:''}`;
}
function readingCard(b){ return `<article class="reading-card" data-book="${b.id}" role="button" tabindex="0">${coverHTML(b)}<div class="reading-info"><div class="book-title">${esc(b.title)}</div><div class="book-author">${esc(b.author)}</div>${progressBlock(b)}</div></article>`; }
function coverCard(b){ return `<article class="cover-card" data-book="${b.id}" role="button" tabindex="0">${coverHTML(b)}<div class="book-title">${esc(b.title)}</div><div class="book-author">${esc(b.author)}</div>${b.rating!==null?`<div class="stars">${ratingStars(b.rating)} <span class="muted">${oneDec(b.rating)}</span></div>`:''}${b.end?`<div class="recent-date">${esc(fmtShortDate(b.end))}</div>`:''}</article>`; }
function getSeriesData(){
  const map=new Map();
  state.books.filter(b=>b.series).forEach(b=>{ if(!map.has(b.series)) map.set(b.series,[]); map.get(b.series).push(b); });
  return [...map.entries()].map(([name,books])=>({name,books:books.sort((a,b)=>(a.seriesNo??999)-(b.seriesNo??999))}));
}
function nextSeriesCandidate(){
  const candidates=[];
  for(const s of getSeriesData()){
    const ordered=s.books.filter(b=>b.seriesNo!==null);
    if(ordered.length<2) continue;
    for(const b of ordered){
      if(b.status==='К прочтению'){
        const prev=ordered.filter(x=>x.seriesNo<b.seriesNo);
        if(prev.length && prev.every(x=>x.status==='Прочитано')) candidates.push({series:s,book:b,done:prev.length});
        break;
      }
    }
  }
  return candidates.sort((a,b)=>b.done-a.done)[0]||null;
}
function renderHome(){
  const active=state.books.filter(b=>b.status==='Читаю');
  const paused=state.books.filter(b=>b.status==='На паузе');
  const yearBooks=readThisYear();
  const audioHours=yearBooks.filter(b=>b.format==='Аудиокнига').reduce((s,b)=>s+(b.hours||0),0);
  const pages=yearBooks.filter(b=>b.format!=='Аудиокнига').reduce((s,b)=>s+(b.pages||0),0);
  const rating=avg(yearBooks.map(b=>b.rating));
  const recent=state.books.filter(b=>b.status==='Прочитано'&&b.end).sort(byEndDesc).slice(0,6);
  const next=nextSeriesCandidate();
  const unread=state.books.filter(b=>b.status==='К прочтению');
  const counts=Object.fromEntries(['Прочитано','Читаю','К прочтению','На паузе','Брошено'].map(s=>[s,state.books.filter(b=>b.status===s).length]));
  return `<div class="page" id="homePage">
    <section class="hero">
      <div class="hero-copy"><div class="eyebrow">Личная библиотека</div><h1>Книги, которые<br>живут со мной.</h1><p>Текущие книги, недавние открытия, серии и следующая история — всё собирается автоматически из Excel.</p></div>
      <div class="hero-card"><div class="year">Мой ${currentYear()}</div><div><div class="hero-number">${yearBooks.length}</div><div>прочитано в этом году</div></div><div class="hero-meta"><span>🎧 ${oneDec(audioHours)} ч</span><span>📖 ${pages} стр.</span>${rating!==null?`<span>★ ${oneDec(rating)}</span>`:''}</div></div>
    </section>

    <section class="section"><div class="section-head"><div><div class="eyebrow">В процессе</div><h2>Сейчас читаю</h2></div><span class="muted">${active.length} ${active.length===1?'книга':'книг'} в процессе</span></div>
      ${active.length?`<div class="reading-grid">${active.map(readingCard).join('')}</div>`:`<div class="empty-state">Сейчас нет книг со статусом «Читаю».</div>`}
      ${paused.length?`<button class="text-link" id="showPaused">На паузе · ${paused.length} ›</button><div id="pausedGrid" class="reading-grid" style="display:none;margin-top:14px">${paused.map(readingCard).join('')}</div>`:''}
    </section>

    <section class="year-strip" aria-label="Статистика за год">
      <div class="metric"><strong>${yearBooks.length}</strong><span>прочитано</span></div>
      <div class="metric"><strong>${oneDec(audioHours)}</strong><span>часов аудио</span></div>
      <div class="metric"><strong>${pages}</strong><span>страниц</span></div>
      <div class="metric"><strong>${rating!==null?oneDec(rating):'—'}</strong><span>средняя оценка</span></div>
    </section>

    <section class="section"><div class="section-head"><div><div class="eyebrow">Последние</div><h2>Недавно прочитано</h2></div><button class="text-link" data-go="library">Все прочитанные →</button></div>
      ${recent.length?`<div class="cover-row">${recent.map(coverCard).join('')}</div>`:`<div class="empty-state">Пока нет законченных книг с датой окончания.</div>`}
    </section>

    ${next?`<section class="section"><div class="section-head"><div><div class="eyebrow">Следующая часть</div><h2>Продолжить серию</h2></div><button class="text-link" data-go="series">Все серии →</button></div>
      <div class="series-feature">${coverHTML(next.book)}<div><div class="series-kicker">${esc(next.series.name)} · книга ${next.book.seriesNo}</div><h2 style="margin-bottom:5px">${esc(next.book.title)}</h2><div class="muted">${esc(next.book.author)}</div><div class="series-dots">${next.series.books.filter(b=>b.seriesNo!==null).map(b=>`<span class="series-dot ${b.status==='Прочитано'?'done':''} ${b.id===next.book.id?'next':''}">${b.seriesNo}</span>`).join('')}</div><p class="muted">${next.done} предыдущих ${next.done===1?'книга прочитана':'книги прочитаны'}.</p><button class="primary-btn alt" data-book="${next.book.id}">Открыть книгу</button></div></div>
    </section>`:''}

    <section class="section"><div class="picker"><div><div class="eyebrow" style="color:#d9a9ba">Случайный выбор</div><h2>Что почитать дальше?</h2><p>${unread.length} книг ждут своей очереди.</p></div><button class="primary-btn" id="randomBook">🎲 Выбрать книгу</button></div></section>

    <section class="section"><div class="section-head"><div><div class="eyebrow">Вся коллекция</div><h2>Моя библиотека</h2></div><button class="text-link" data-go="stats">Вся статистика →</button></div>
      <div class="library-summary"><div class="big-total"><strong>${state.books.length}</strong><span>книг всего</span></div><div class="status-list">${Object.entries(counts).map(([k,v])=>`<div class="status-item"><span>${esc(k)}</span><strong>${v}</strong></div>`).join('')}</div></div>
    </section>
  </div>`;
}
function renderLibrary(){
  const statuses=['Все','Читаю','К прочтению','Прочитано','На паузе','Брошено'];
  const formats=['Все','Аудиокнига','Книга','Комикс'];
  let books=state.books.filter(b=>(state.filters.status==='Все'||b.status===state.filters.status)&&(state.filters.format==='Все'||b.format===state.filters.format));
  const q=state.filters.q.trim().toLowerCase(); if(q) books=books.filter(b=>[b.title,b.author,b.series].some(x=>x.toLowerCase().includes(q)));
  books=[...books].sort(state.filters.sort==='title'?(a,b)=>a.title.localeCompare(b.title,'ru'):state.filters.sort==='rating'?(a,b)=>(b.rating??-1)-(a.rating??-1):byEndDesc);
  return `<div class="page"><div class="eyebrow">Каталог</div><h1 style="font-size:clamp(2.7rem,6vw,4.8rem)">Библиотека</h1><p class="muted">${books.length} из ${state.books.length} книг</p>
    <div class="toolbar"><input id="libSearch" type="search" placeholder="Поиск…" value="${esc(state.filters.q)}"><select id="statusFilter">${statuses.map(x=>`<option ${x===state.filters.status?'selected':''}>${x}</option>`).join('')}</select><select id="formatFilter">${formats.map(x=>`<option ${x===state.filters.format?'selected':''}>${x}</option>`).join('')}</select><select id="sortFilter"><option value="recent" ${state.filters.sort==='recent'?'selected':''}>Недавно прочитанные</option><option value="title" ${state.filters.sort==='title'?'selected':''}>По названию</option><option value="rating" ${state.filters.sort==='rating'?'selected':''}>По оценке</option></select></div>
    ${books.length?`<div class="library-grid">${books.map(coverCard).join('')}</div>`:`<div class="empty-state">По этим условиям ничего не найдено.</div>`}
  </div>`;
}
function renderSeries(){
  const series=getSeriesData().sort((a,b)=>a.name.localeCompare(b.name,'ru'));
  return `<div class="page"><div class="eyebrow">Порядок чтения</div><h1 style="font-size:clamp(2.7rem,6vw,4.8rem)">Серии</h1><p class="muted">${series.length} серий в библиотеке</p><div class="series-grid">${series.map(s=>{const done=s.books.filter(b=>b.status==='Прочитано').length;return `<article class="series-card"><h3>${esc(s.name)}</h3><div class="muted">${done} из ${s.books.length} прочитано</div><div class="progress" style="margin-top:12px"><span style="width:${s.books.length?done/s.books.length*100:0}%"></span></div><div class="series-books">${s.books.map(b=>`<div class="series-book-row" data-book="${b.id}" role="button"><span>${b.status==='Прочитано'?'✓':b.status==='Читаю'?'▶':'○'}</span><strong>${b.seriesNo??'—'}.</strong><span>${esc(b.title)}</span></div>`).join('')}</div></article>`}).join('')}</div></div>`;
}
function renderStats(){
  const counts=Object.fromEntries(['Прочитано','Читаю','К прочтению','На паузе','Брошено'].map(s=>[s,state.books.filter(b=>b.status===s).length]));
  const rated=state.books.filter(b=>b.rating!==null); const ar=avg(rated.map(b=>b.rating));
  const years={}; state.books.filter(b=>b.status==='Прочитано'&&b.end).forEach(b=>years[b.end.getFullYear()]=(years[b.end.getFullYear()]||0)+1);
  const maxY=Math.max(1,...Object.values(years));
  return `<div class="page"><div class="eyebrow">Цифры</div><h1 style="font-size:clamp(2.7rem,6vw,4.8rem)">Статистика</h1><div class="stats-grid"><div class="stat-card"><strong>${state.books.length}</strong><span>книг в библиотеке</span></div><div class="stat-card"><strong>${counts['Прочитано']}</strong><span>прочитано</span></div><div class="stat-card"><strong>${ar!==null?oneDec(ar):'—'}</strong><span>средняя оценка</span></div><div class="stat-card"><strong>${getSeriesData().length}</strong><span>серий</span></div></div>
    <section class="section"><div class="section-head"><div><div class="eyebrow">Динамика</div><h2>Прочитано по годам</h2></div></div><div class="chart-list">${Object.entries(years).sort((a,b)=>a[0]-b[0]).map(([y,v])=>`<div class="chart-row"><strong>${y}</strong><div class="bar"><span style="width:${v/maxY*100}%"></span></div><strong>${v}</strong></div>`).join('')||'<div class="empty-state">Недостаточно дат окончания.</div>'}</div></section>
    <section class="section"><div class="section-head"><div><div class="eyebrow">Статусы</div><h2>Состояние библиотеки</h2></div></div><div class="stats-grid">${Object.entries(counts).map(([k,v])=>`<div class="stat-card"><strong>${v}</strong><span>${esc(k)}</span></div>`).join('')}</div></section>
  </div>`;
}
function render(){
  const route=(location.hash||'#home').slice(1); state.route=['home','library','series','stats'].includes(route)?route:'home';
  $('#app').innerHTML=state.route==='home'?renderHome():state.route==='library'?renderLibrary():state.route==='series'?renderSeries():renderStats();
  $$('.nav a').forEach(a=>a.classList.toggle('active',a.dataset.route===state.route));
  bindDynamic();
}
function showBook(id){
  const b=state.books.find(x=>x.id===Number(id)); if(!b) return;
  const volume=b.format==='Аудиокнига'?(b.hours?`${oneDec(b.hours)} ч`:'—'):(b.pages?`${b.pages} стр.`:'—');
  const progress=b.progress!==null?`${Math.round(b.progress)}%${b.progressText?` · ${b.progressText}`:''}`:(b.status==='Прочитано'?'100%':'—');
  $('#bookDialog').innerHTML=`<button class="icon-btn dialog-close" onclick="document.querySelector('#bookDialog').close()">×</button><div class="dialog-inner">${coverHTML(b)}<div><div class="eyebrow">${esc(b.status||'Книга')}</div><h2 style="font-size:2.5rem;margin-top:6px">${esc(b.title)}</h2><p class="muted">${esc(b.author)}</p>${b.series?`<p><strong>${esc(b.series)}</strong>${b.seriesNo?` · книга ${b.seriesNo}`:''}</p>`:''}${b.progress!==null&&b.status!=='Прочитано'?`<div class="progress-copy"><span>${esc(b.progressText)}</span><strong>${Math.round(b.progress)}%</strong></div><div class="progress"><span style="width:${b.progress}%"></span></div>`:''}<div class="dialog-meta"><div class="meta-box"><small>Формат</small><strong>${esc(b.format||'—')}</strong></div><div class="meta-box"><small>Объём</small><strong>${esc(volume)}</strong></div><div class="meta-box"><small>Начало</small><strong>${esc(fmtDate(b.start))}</strong></div><div class="meta-box"><small>Конец</small><strong>${esc(fmtDate(b.end))}</strong></div><div class="meta-box"><small>Длительность</small><strong>${b.duration!==null?`${b.duration} дн.`:'—'}</strong></div><div class="meta-box"><small>Оценка</small><strong>${b.rating!==null?`${ratingStars(b.rating)} ${oneDec(b.rating)}`:'—'}</strong></div></div></div></div>`;
  $('#bookDialog').showModal();
}
function randomUnread(){ const arr=state.books.filter(b=>b.status==='К прочтению'); if(!arr.length)return; showBook(arr[Math.floor(Math.random()*arr.length)].id); }
function bindDynamic(){
  $$('[data-book]').forEach(el=>{el.addEventListener('click',()=>showBook(el.dataset.book));el.addEventListener('keydown',e=>{if(e.key==='Enter')showBook(el.dataset.book)})});
  $$('[data-go]').forEach(el=>el.addEventListener('click',()=>location.hash=el.dataset.go));
  $('#showPaused')?.addEventListener('click',()=>{const g=$('#pausedGrid');g.style.display=g.style.display==='none'?'grid':'none'});
  $('#randomBook')?.addEventListener('click',randomUnread);
  $('#libSearch')?.addEventListener('input',e=>{state.filters.q=e.target.value; render(); $('#libSearch')?.focus(); try{$('#libSearch').setSelectionRange(state.filters.q.length,state.filters.q.length)}catch{}});
  $('#statusFilter')?.addEventListener('change',e=>{state.filters.status=e.target.value;render()});
  $('#formatFilter')?.addEventListener('change',e=>{state.filters.format=e.target.value;render()});
  $('#sortFilter')?.addEventListener('change',e=>{state.filters.sort=e.target.value;render()});
}
function bindGlobalSearch(){
  const dlg=$('#searchDialog'), input=$('#globalSearch'), results=$('#searchResults');
  $('#openSearch').addEventListener('click',()=>{dlg.showModal(); setTimeout(()=>input.focus(),50)});
  input.addEventListener('input',()=>{const q=input.value.trim().toLowerCase(); const arr=q?state.books.filter(b=>[b.title,b.author,b.series].some(x=>x.toLowerCase().includes(q))).slice(0,12):[];results.innerHTML=arr.map(b=>`<div class="search-result" data-search-book="${b.id}" role="button">${coverHTML(b,false)}<div><strong>${esc(b.title)}</strong><div class="muted">${esc(b.author)}</div></div></div>`).join('')||(q?'<div class="empty-state">Ничего не найдено</div>':'<div class="muted" style="padding:16px">Начни вводить название, автора или серию.</div>'); $$('[data-search-book]',results).forEach(x=>x.addEventListener('click',()=>{dlg.close();showBook(x.dataset.searchBook)}));});
}
function handleRouteChange(){
  render();
  requestAnimationFrame(() => window.scrollTo({ top: 0, left: 0, behavior: 'auto' }));
}
async function init(){
  try {
    await loadBooks();
    if ('scrollRestoration' in history) history.scrollRestoration = 'manual';
    render();
    window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    bindGlobalSearch();
    window.addEventListener('hashchange', handleRouteChange);
  }
  catch(e){ $('#app').innerHTML=`<div class="empty-state"><h2>Не удалось загрузить Excel</h2><p>${esc(e.message)}</p><p>Проверь, что <strong>books.xlsx</strong> лежит рядом с index.html и сайт открыт через GitHub Pages или локальный веб‑сервер.</p></div>`; console.error(e); }
}
init();
