const DATA_FILE = 'books.xlsx';

const state = {
  books: [],
  route: 'home',

  filters: {
    status: 'Все',
    format: 'Все',
    q: '',
    sort: 'recent'
  },

  statsYear: new Date().getFullYear()
};


// ======================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ======================================================

const $ = (s, root = document) => root.querySelector(s);

const $$ = (s, root = document) =>
  [...root.querySelectorAll(s)];

const esc = (s = '') =>
  String(s).replace(
    /[&<>'"]/g,
    c => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      "'": '&#39;',
      '"': '&quot;'
    }[c])
  );


// ======================================================
// СТАТУСЫ
// ======================================================

const statusMap = {
  'Я все прочитал!': 'Прочитано',
  'Бросил читать': 'Брошено',
  'Чтение': 'Читаю',
  'Пауза': 'На паузе',

  'Прочитано': 'Прочитано',
  'Брошено': 'Брошено',
  'Читаю': 'Читаю',
  'На паузе': 'На паузе',
  'К прочтению': 'К прочтению'
};


// ======================================================
// РАБОТА С ДАТАМИ
// ======================================================

function excelDate(v) {
  if (!v) return null;

  if (v instanceof Date && !isNaN(v)) {
    return v;
  }

  if (typeof v === 'number') {
    const p = XLSX.SSF.parse_date_code(v);

    return p
      ? new Date(p.y, p.m - 1, p.d)
      : null;
  }

  const d = new Date(v);

  return isNaN(d)
    ? null
    : d;
}


function fmtDate(d) {
  return d
    ? new Intl.DateTimeFormat(
        'ru-RU',
        {
          day: 'numeric',
          month: 'long',
          year: 'numeric'
        }
      ).format(d)
    : '—';
}


function fmtShortDate(d) {
  return d
    ? new Intl.DateTimeFormat(
        'ru-RU',
        {
          day: 'numeric',
          month: 'long'
        }
      ).format(d)
    : '';
}


// ======================================================
// РАБОТА С ТЕКСТОМ И ЧИСЛАМИ
// ======================================================

function text(v, fallback = '') {
  if (v === null || v === undefined) {
    return fallback;
  }

  const s = String(v).trim();

  return s || fallback;
}


function num(v) {
  if (
    v === null ||
    v === undefined ||
    v === ''
  ) {
    return null;
  }

  const n = Number(
    String(v).replace(',', '.')
  );

  return Number.isFinite(n)
    ? n
    : null;
}


function oneDec(n) {
  if (n === null || n === undefined) {
    return '0';
  }

  return Number.isInteger(n)
    ? String(n)
    : n.toFixed(1).replace('.', ',');
}


function ratingStars(n) {
  if (n === null) return '';

  const full = Math.max(
    0,
    Math.min(
      5,
      Math.round(n)
    )
  );

  return (
    '★'.repeat(full) +
    '☆'.repeat(5 - full)
  );
}


function initials(title = '') {
  return title
    .split(/\s+/)
    .slice(0, 2)
    .map(x => x[0] || '')
    .join('')
    .toUpperCase();
}


// ======================================================
// ОБЛОЖКА
// ======================================================

function coverHTML(book, badge = true) {

  const url =
    book.cover &&
    book.cover.replaceAll('&amp;', '&');

  return `
    <div class="cover-wrap">

      ${
        url
          ? `
            <img
              src="${esc(url)}"
              alt="Обложка ${esc(book.title)}"
              loading="lazy"
              onerror="
                this.remove();
                this.parentElement
                  .querySelector('.cover-fallback')
                  .style.display='grid'
              "
            >
          `
          : ''
      }

      <div
        class="cover-fallback"
        style="${url ? 'display:none' : ''}"
      >
        ${esc(initials(book.title))}
      </div>

      ${
        badge &&
        book.format === 'Аудиокнига'
          ? `
            <span class="format-badge">
              🎧
            </span>
          `
          : ''
      }

    </div>
  `;
}


// ======================================================
// НОРМАЛИЗАЦИЯ СТРОКИ EXCEL
// ======================================================

function normalizeRow(r, i) {

  const format =
    text(r['Формат']);

  const pages =
    num(r['Страницы']);

  const hours =
    num(r['Часы']);

  const pagesRead =
    num(r['Прочитано страниц']);

  const hoursRead =
    num(r['Прослушано часов']);

  const rawStatus =
    text(r['Статус']);

  const status =
    statusMap[rawStatus] || rawStatus;


  let progress = null;
  let progressText = '';


  // Прочитанная книга всегда считается 100%
  if (status === 'Прочитано') {

    progress = 100;

  }

  // Аудиокнига
  else if (
    format === 'Аудиокнига' &&
    hours &&
    hoursRead !== null
  ) {

    progress = Math.min(
      100,
      Math.max(
        0,
        hoursRead / hours * 100
      )
    );

    progressText =
      `${oneDec(hoursRead)} из ${oneDec(hours)} ч`;

  }

  // Обычная книга / комикс
  else if (
    format !== 'Аудиокнига' &&
    pages &&
    pagesRead !== null
  ) {

    progress = Math.min(
      100,
      Math.max(
        0,
        pagesRead / pages * 100
      )
    );

    progressText =
      `${pagesRead} из ${pages} стр.`;

  }


  // Для полностью прочитанной книги
  if (
    status === 'Прочитано' &&
    !progressText
  ) {

    if (
      format === 'Аудиокнига' &&
      hours
    ) {

      progressText =
        `${oneDec(hours)} ч`;

    }

    else if (pages) {

      progressText =
        `${pages} стр.`;

    }

  }


  return {

    id: i,

    title:
      text(
        r['Название'],
        'Без названия'
      ),

    author:
      text(
        r['Автор'],
        'Автор не указан'
      ),

    series:
      text(r['Серия']),

    seriesNo:
      num(r['Книга в серии']),

    status,

    format,

    pages,

    hours,

    pagesRead,

    hoursRead,

    start:
      excelDate(r['Начало']),

    end:
      excelDate(r['Конец']),

    duration:
      num(
        r['Длительность чтения, дней'] ??
        r['Общее время прочтения']
      ),

    rating:
      num(r['Оценка']),

    cover:
      text(r['Обложка']),

    progress,

    progressText
  };
}


// ======================================================
// ЗАГРУЗКА EXCEL
// ======================================================

async function loadBooks() {

  const res =
    await fetch(
      `${DATA_FILE}?v=${Date.now()}`
    );

  if (!res.ok) {
    throw new Error(
      `Не удалось загрузить ${DATA_FILE}`
    );
  }

  const ab =
    await res.arrayBuffer();

  const wb =
    XLSX.read(
      ab,
      {
        type: 'array',
        cellDates: true
      }
    );

  const ws =
    wb.Sheets[
      wb.SheetNames[0]
    ];

  const rows =
    XLSX.utils.sheet_to_json(
      ws,
      {
        defval: ''
      }
    );


  state.books =
    rows
      .map(normalizeRow)
      .filter(
        b =>
          b.title &&
          b.title !== 'Без названия'
      );
}


// ======================================================
// СТАТИСТИКА
// ======================================================

function currentYear() {
  return new Date().getFullYear();
}


function byEndDesc(a, b) {

  return (
    (b.end?.getTime() || 0) -
    (a.end?.getTime() || 0)
  );
}


// Книги, полностью прочитанные в текущем году
function readThisYear() {

  const y =
    currentYear();

  return state.books.filter(
    b =>
      b.status === 'Прочитано' &&
      b.end?.getFullYear() === y
  );
}


// ======================================================
// ФАКТИЧЕСКИ ПРОСЛУШАННЫЕ ЧАСЫ ЗА ГОД
// ======================================================

function audioHoursThisYear() {

  const y =
    currentYear();

  return state.books
    .filter(
      b =>
        b.format === 'Аудиокнига'
    )
    .reduce(
      (sum, b) => {

        // ------------------------------------------
        // К прочтению = 0 часов
        // ------------------------------------------

        if (
          b.status === 'К прочтению'
        ) {
          return sum;
        }


        // ------------------------------------------
        // Определяем дату,
        // по которой книга относится к году
        // ------------------------------------------

        let date = null;


        // Прочитано / Брошено
        // сначала смотрим дату окончания
        if (
          b.status === 'Прочитано' ||
          b.status === 'Брошено'
        ) {

          date =
            b.end ||
            b.start;

        }


        // Читаю / На паузе
        // ориентируемся на дату начала
        else if (
          b.status === 'Читаю' ||
          b.status === 'На паузе'
        ) {

          date =
            b.start;

        }


        // Неизвестный статус
        else {

          return sum;

        }


        // Нет даты
        if (!date) {
          return sum;
        }


        // Не текущий год
        if (
          date.getFullYear() !== y
        ) {
          return sum;
        }


        // ------------------------------------------
        // ПРОЧИТАНО
        //
        // Если заполнено
        // "Прослушано часов",
        // используем его.
        //
        // Если нет —
        // считаем, что книга
        // прослушана полностью.
        // ------------------------------------------

        if (
          b.status === 'Прочитано'
        ) {

          return (
            sum +
            (
              b.hoursRead ??
              b.hours ??
              0
            )
          );

        }


        // ------------------------------------------
        // БРОШЕНО / НА ПАУЗЕ / ЧИТАЮ
        //
        // Только реально
        // прослушанные часы
        // ------------------------------------------

        if (
          b.status === 'Брошено' ||
          b.status === 'На паузе' ||
          b.status === 'Читаю'
        ) {

          return (
            sum +
            (
              b.hoursRead ??
              0
            )
          );

        }


        return sum;

      },
      0
    );
}


// Среднее значение
function avg(arr) {

  const x =
    arr.filter(
      n => n !== null
    );

  return x.length
    ? x.reduce(
        (a, b) => a + b,
        0
      ) / x.length
    : null;
}


// ======================================================
// ПРОГРЕСС
// ======================================================

function progressBlock(b) {

  if (
    b.progress === null
  ) {

    return `
      <div class="start-date">

        В процессе

        ${
          b.start
            ? ` · начато ${esc(
                fmtShortDate(b.start)
              )}`
            : ''
        }

      </div>
    `;
  }


  return `
    <div class="progress-copy">

      <span>
        ${esc(
          b.progressText ||
          'Прогресс'
        )}
      </span>

      <strong>
        ${Math.round(
          b.progress
        )}%
      </strong>

    </div>

    <div class="progress">

      <span
        style="width:${b.progress}%"
      ></span>

    </div>

    ${
      b.start
        ? `
          <div class="start-date">
            Начато ${esc(
              fmtShortDate(b.start)
            )}
          </div>
        `
        : ''
    }
  `;
}


// ======================================================
// КАРТОЧКА ТЕКУЩЕЙ КНИГИ
// ======================================================

function readingCard(b) {

  return `
    <article
      class="reading-card"
      data-book="${b.id}"
      role="button"
      tabindex="0"
    >

      ${coverHTML(b)}

      <div class="reading-info">

        <div class="book-title">
          ${esc(b.title)}
        </div>

        <div class="book-author">
          ${esc(b.author)}
        </div>

        ${progressBlock(b)}

      </div>

    </article>
  `;
}


// ======================================================
// ОБЫЧНАЯ КАРТОЧКА КНИГИ
// ======================================================

function coverCard(b) {

  return `
    <article
      class="cover-card"
      data-book="${b.id}"
      role="button"
      tabindex="0"
    >

      ${coverHTML(b)}

      <div class="book-title">
        ${esc(b.title)}
      </div>

      <div class="book-author">
        ${esc(b.author)}
      </div>

      ${
        b.rating !== null
          ? `
            <div class="stars">

              ${ratingStars(
                b.rating
              )}

              <span class="muted">
                ${oneDec(
                  b.rating
                )}
              </span>

            </div>
          `
          : ''
      }

      ${
        b.end
          ? `
            <div class="recent-date">
              ${esc(
                fmtShortDate(b.end)
              )}
            </div>
          `
          : ''
      }

    </article>
  `;
}


// ======================================================
// СЕРИИ
// ======================================================

function getSeriesData() {

  const map =
    new Map();


  state.books
    .filter(
      b => b.series
    )
    .forEach(
      b => {

        if (
          !map.has(b.series)
        ) {
          map.set(
            b.series,
            []
          );
        }

        map
          .get(b.series)
          .push(b);

      }
    );


  return [
    ...map.entries()
  ].map(
    ([name, books]) => ({

      name,

      books:
        books.sort(
          (a, b) =>
            (a.seriesNo ?? 999) -
            (b.seriesNo ?? 999)
        )

    })
  );
}


// ======================================================
// СЛЕДУЮЩАЯ КНИГА СЕРИИ
// ======================================================

function nextSeriesCandidate() {

  const candidates = [];


  for (
    const s of getSeriesData()
  ) {

    const ordered =
      s.books.filter(
        b =>
          b.seriesNo !== null
      );


    if (
      ordered.length < 2
    ) {
      continue;
    }


    for (
      const b of ordered
    ) {

      if (
        b.status === 'К прочтению'
      ) {

        const prev =
          ordered.filter(
            x =>
              x.seriesNo <
              b.seriesNo
          );


        if (
          prev.length &&
          prev.every(
            x =>
              x.status ===
              'Прочитано'
          )
        ) {

          candidates.push({
            series: s,
            book: b,
            done: prev.length
          });

        }

        break;
      }

    }

  }


  return (
    candidates.sort(
      (a, b) =>
        b.done -
        a.done
    )[0] ||
    null
  );
}


// ======================================================
// ГЛАВНАЯ
// ======================================================

function renderHome() {

  const active =
    state.books.filter(
      b =>
        b.status === 'Читаю'
    );


  const paused =
    state.books.filter(
      b =>
        b.status === 'На паузе'
    );


  const yearBooks =
    readThisYear();


  // ==================================================
  // ИСПРАВЛЕНО:
  // фактически прослушанные часы
  // ==================================================

  const audioHours =
    audioHoursThisYear();


  // Страницы пока считаются
  // только для полностью прочитанных книг
  const pages =
    yearBooks
      .filter(
        b =>
          b.format !==
          'Аудиокнига'
      )
      .reduce(
        (s, b) =>
          s +
          (b.pages || 0),
        0
      );


  const rating =
    avg(
      yearBooks.map(
        b => b.rating
      )
    );


  const recent =
    state.books
      .filter(
        b =>
          b.status ===
            'Прочитано' &&
          b.end
      )
      .sort(byEndDesc)
      .slice(0, 6);


  const next =
    nextSeriesCandidate();


  const unread =
    state.books.filter(
      b =>
        b.status ===
        'К прочтению'
    );


  const counts =
    Object.fromEntries(
      [
        'Прочитано',
        'Читаю',
        'К прочтению',
        'На паузе',
        'Брошено'
      ].map(
        s => [
          s,
          state.books.filter(
            b =>
              b.status === s
          ).length
        ]
      )
    );


  return `
    <div
      class="page"
      id="homePage"
    >

      <section class="hero">

        <div class="hero-copy">

          <div class="eyebrow">
            Личная библиотека
          </div>

          <h1>
            Книги, которые<br>
            живут со мной.
          </h1>

          <p>
            Текущие книги,
            недавние открытия,
            серии и следующая история.
          </p>

        </div>


        <div class="hero-card">

          <div class="year">
            Мой ${currentYear()}
          </div>

          <div>

            <div class="hero-number">
              ${yearBooks.length}
            </div>

            <div>
              прочитано в этом году
            </div>

          </div>


          <div class="hero-meta">

            <span>
              🎧 ${oneDec(audioHours)} ч
            </span>

            <span>
              📖 ${pages} стр.
            </span>

            ${
              rating !== null
                ? `
                  <span>
                    ★ ${oneDec(rating)}
                  </span>
                `
                : ''
            }

          </div>

        </div>

      </section>


      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              В процессе
            </div>

            <h2>
              Сейчас читаю
            </h2>

          </div>

          <span class="muted">
            ${active.length}
            ${
              active.length === 1
                ? 'книга'
                : 'книг'
            }
            в процессе
          </span>

        </div>


        ${
          active.length
            ? `
              <div class="reading-grid">
                ${active
                  .map(readingCard)
                  .join('')}
              </div>
            `
            : `
              <div class="empty-state">
                Сейчас нет книг
                со статусом «Читаю».
              </div>
            `
        }


        ${
          paused.length
            ? `
              <button
                class="text-link"
                id="showPaused"
              >
                На паузе ·
                ${paused.length} ›
              </button>

              <div
                id="pausedGrid"
                class="reading-grid"
                style="
                  display:none;
                  margin-top:14px
                "
              >
                ${paused
                  .map(readingCard)
                  .join('')}
              </div>
            `
            : ''
        }

      </section>


      <section
        class="year-strip"
        aria-label="Статистика за год"
      >

        <div class="metric">

          <strong>
            ${yearBooks.length}
          </strong>

          <span>
            прочитано
          </span>

        </div>


        <div class="metric">

          <strong>
            ${oneDec(audioHours)}
          </strong>

          <span>
            прослушано часов
          </span>

        </div>


        <div class="metric">

          <strong>
            ${pages}
          </strong>

          <span>
            страниц
          </span>

        </div>


        <div class="metric">

          <strong>
            ${
              rating !== null
                ? oneDec(rating)
                : '—'
            }
          </strong>

          <span>
            средняя оценка
          </span>

        </div>

      </section>


      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Последние
            </div>

            <h2>
              Недавно прочитано
            </h2>

          </div>

          <button
            class="text-link"
            data-go="library"
          >
            Все прочитанные →
          </button>

        </div>


        ${
          recent.length
            ? `
              <div class="cover-row">

                ${recent
                  .map(coverCard)
                  .join('')}

              </div>
            `
            : `
              <div class="empty-state">
                Пока нет законченных
                книг с датой окончания.
              </div>
            `
        }

      </section>


      ${
        next
          ? `
            <section class="section">

              <div class="section-head">

                <div>

                  <div class="eyebrow">
                    Следующая часть
                  </div>

                  <h2>
                    Продолжить серию
                  </h2>

                </div>

                <button
                  class="text-link"
                  data-go="series"
                >
                  Все серии →
                </button>

              </div>


              <div class="series-feature">

                ${coverHTML(
                  next.book
                )}

                <div>

                  <div class="series-kicker">
                    ${esc(
                      next.series.name
                    )}
                    · книга
                    ${next.book.seriesNo}
                  </div>

                  <h2
                    style="
                      margin-bottom:5px
                    "
                  >
                    ${esc(
                      next.book.title
                    )}
                  </h2>

                  <div class="muted">
                    ${esc(
                      next.book.author
                    )}
                  </div>


                  <div class="series-dots">

                    ${next.series.books
                      .filter(
                        b =>
                          b.seriesNo !== null
                      )
                      .map(
                        b => `
                          <span
                            class="
                              series-dot
                              ${
                                b.status ===
                                'Прочитано'
                                  ? 'done'
                                  : ''
                              }
                              ${
                                b.id ===
                                next.book.id
                                  ? 'next'
                                  : ''
                              }
                            "
                          >
                            ${b.seriesNo}
                          </span>
                        `
                      )
                      .join('')}

                  </div>


                  <p class="muted">

                    ${next.done}
                    предыдущих

                    ${
                      next.done === 1
                        ? 'книга прочитана'
                        : 'книги прочитаны'
                    }.

                  </p>


                  <button
                    class="primary-btn alt"
                    data-book="${next.book.id}"
                  >
                    Открыть книгу
                  </button>

                </div>

              </div>

            </section>
          `
          : ''
      }


      <section class="section">

        <div class="picker">

          <div>

            <div
              class="eyebrow"
              style="color:#d9a9ba"
            >
              Случайный выбор
            </div>

            <h2>
              Что почитать дальше?
            </h2>

            <p>
              ${unread.length}
              книг ждут своей очереди.
            </p>

          </div>


          <button
            class="primary-btn"
            id="randomBook"
          >
            🎲 Выбрать книгу
          </button>

        </div>

      </section>


      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Вся коллекция
            </div>

            <h2>
              Моя библиотека
            </h2>

          </div>


          <button
            class="text-link"
            data-go="stats"
          >
            Вся статистика →
          </button>

        </div>


        <div class="library-summary">

          <div class="big-total">

            <strong>
              ${state.books.length}
            </strong>

            <span>
              книг всего
            </span>

          </div>


          <div class="status-list">

            ${Object.entries(
              counts
            )
              .map(
                ([k, v]) => `
                  <div class="status-item">

                    <span>
                      ${esc(k)}
                    </span>

                    <strong>
                      ${v}
                    </strong>

                  </div>
                `
              )
              .join('')}

          </div>

        </div>

      </section>

    </div>
  `;
}


// ======================================================
// БИБЛИОТЕКА
// ======================================================

function renderLibrary() {

  const statuses = [
    'Все',
    'Читаю',
    'К прочтению',
    'Прочитано',
    'На паузе',
    'Брошено'
  ];


  const formats = [
    'Все',
    'Аудиокнига',
    'Книга',
    'Комикс'
  ];


  let books =
    state.books.filter(
      b =>
        (
          state.filters.status ===
            'Все' ||
          b.status ===
            state.filters.status
        ) &&
        (
          state.filters.format ===
            'Все' ||
          b.format ===
            state.filters.format
        )
    );


  const q =
    state.filters.q
      .trim()
      .toLowerCase();


  if (q) {

    books =
      books.filter(
        b =>
          [
            b.title,
            b.author,
            b.series
          ].some(
            x =>
              x
                .toLowerCase()
                .includes(q)
          )
      );

  }


  books =
    [...books].sort(

      state.filters.sort ===
      'title'

        ? (
            (a, b) =>
              a.title.localeCompare(
                b.title,
                'ru'
              )
          )

        : state.filters.sort ===
          'rating'

        ? (
            (a, b) =>
              (b.rating ?? -1) -
              (a.rating ?? -1)
          )

        : byEndDesc
    );


  return `
    <div class="page">

      <div class="eyebrow">
        Каталог
      </div>

      <h1
        style="
          font-size:
            clamp(
              2.7rem,
              6vw,
              4.8rem
            )
        "
      >
        Библиотека
      </h1>

      <p class="muted">
        ${books.length}
        из
        ${state.books.length}
        книг
      </p>


      <div class="toolbar">

        <input
          id="libSearch"
          type="search"
          placeholder="Поиск…"
          value="${esc(
            state.filters.q
          )}"
        >


        <select id="statusFilter">

          ${statuses
            .map(
              x => `
                <option
                  value="${x}"
                  ${
                    x ===
                    state.filters.status
                      ? 'selected'
                      : ''
                  }
                >
                  ${
                    x === 'Все'
                      ? 'Статус'
                      : x
                  }
                </option>
              `
            )
            .join('')}

        </select>


        <select id="formatFilter">

          ${formats
            .map(
              x => `
                <option
                  value="${x}"
                  ${
                    x ===
                    state.filters.format
                      ? 'selected'
                      : ''
                  }
                >
                  ${
                    x === 'Все'
                      ? 'Все форматы'
                      : x
                  }
                </option>
              `
            )
            .join('')}

        </select>


        <select id="sortFilter">

          <option
            value="recent"
            ${
              state.filters.sort ===
              'recent'
                ? 'selected'
                : ''
            }
          >
            Недавно прочитанные
          </option>

          <option
            value="title"
            ${
              state.filters.sort ===
              'title'
                ? 'selected'
                : ''
            }
          >
            По названию
          </option>

          <option
            value="rating"
            ${
              state.filters.sort ===
              'rating'
                ? 'selected'
                : ''
            }
          >
            По оценке
          </option>

        </select>

      </div>


      ${
        books.length
          ? `
            <div class="library-grid">

              ${books
                .map(coverCard)
                .join('')}

            </div>
          `
          : `
            <div class="empty-state">
              По этим условиям
              ничего не найдено.
            </div>
          `
      }

    </div>
  `;
}


// ======================================================
// СЕРИИ
// ======================================================

function renderSeries() {

  const series =
    getSeriesData()
      .sort(
        (a, b) =>
          a.name.localeCompare(
            b.name,
            'ru'
          )
      );


  return `
    <div class="page">

      <div class="eyebrow">
        Порядок чтения
      </div>

      <h1
        style="
          font-size:
            clamp(
              2.7rem,
              6vw,
              4.8rem
            )
        "
      >
        Серии
      </h1>

      <p class="muted">
        ${series.length}
        серий в библиотеке
      </p>


      <div class="series-grid">

        ${series
          .map(
            s => {

              const done =
                s.books.filter(
                  b =>
                    b.status ===
                    'Прочитано'
                ).length;


              return `
                <article
                  class="series-card"
                >

                  <h3>
                    ${esc(s.name)}
                  </h3>

                  <div class="muted">
                    ${done}
                    из
                    ${s.books.length}
                    прочитано
                  </div>


                  <div
                    class="progress"
                    style="margin-top:12px"
                  >

                    <span
                      style="
                        width:${
                          s.books.length
                            ? done /
                              s.books.length *
                              100
                            : 0
                        }%
                      "
                    ></span>

                  </div>


                  <div class="series-books">

                    ${s.books
                      .map(
                        b => `
                          <div
                            class="series-book-row"
                            data-book="${b.id}"
                            role="button"
                          >

                            <span>
                              ${
                                b.status ===
                                'Прочитано'
                                  ? '✓'
                                  : b.status ===
                                    'Читаю'
                                  ? '▶'
                                  : '○'
                              }
                            </span>

                            <strong>
                              ${
                                b.seriesNo ??
                                '—'
                              }.
                            </strong>

                            <span>
                              ${esc(
                                b.title
                              )}
                            </span>

                          </div>
                        `
                      )
                      .join('')}

                  </div>

                </article>
              `;
            }
          )
          .join('')}

      </div>

    </div>
  `;
}


// ======================================================
// СТАТИСТИКА
// ======================================================

// ======================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ СТАТИСТИКИ
// ======================================================


// Все годы, которые встречаются в данных
function getStatsYears() {

  const years = new Set();

  state.books.forEach(b => {

    if (b.start) {
      years.add(
        b.start.getFullYear()
      );
    }

    if (b.end) {
      years.add(
        b.end.getFullYear()
      );
    }

  });

  // Текущий год показываем всегда
  years.add(
    currentYear()
  );

  return [...years]
    .sort(
      (a, b) => b - a
    );
}


// ------------------------------------------------------
// Дата, по которой книга относится к году статистики
// ------------------------------------------------------

function statsDateForBook(b) {

  // Для прочитанных и брошенных
  // приоритет у даты окончания
  if (
    b.status === 'Прочитано' ||
    b.status === 'Брошено'
  ) {

    return (
      b.end ||
      b.start ||
      null
    );

  }


  // Для книг в процессе и на паузе
  // используем дату начала
  if (
    b.status === 'Читаю' ||
    b.status === 'На паузе'
  ) {

    return (
      b.start ||
      null
    );

  }


  return null;
}


// ------------------------------------------------------
// Относится ли книга к выбранному периоду
// ------------------------------------------------------

function bookInStatsPeriod(
  b,
  period = state.statsYear
) {

  // За всё время
  if (
    period === 'all'
  ) {

    return true;

  }


  const date =
    statsDateForBook(b);


  return (
    date &&
    date.getFullYear() ===
      Number(period)
  );
}


// ------------------------------------------------------
// Прочитанные книги за период
// ------------------------------------------------------

function completedBooksForStats(
  period = state.statsYear
) {

  return state.books.filter(
    b => {

      if (
        b.status !== 'Прочитано'
      ) {
        return false;
      }


      if (
        period === 'all'
      ) {
        return true;
      }


      return (
        b.end &&
        b.end.getFullYear() ===
          Number(period)
      );

    }
  );
}


// ------------------------------------------------------
// ФАКТИЧЕСКИ ПРОСЛУШАННЫЕ ЧАСЫ
//
// Прочитано:
// Прослушано часов,
// а если пусто → Часы.
//
// Читаю / На паузе / Брошено:
// только Прослушано часов.
//
// К прочтению:
// 0.
// ------------------------------------------------------

function listenedHoursForStats(
  period = state.statsYear
) {

  return state.books
    .filter(
      b =>
        b.format ===
        'Аудиокнига'
    )
    .reduce(
      (sum, b) => {

        if (
          b.status ===
          'К прочтению'
        ) {
          return sum;
        }


        if (
          !bookInStatsPeriod(
            b,
            period
          )
        ) {
          return sum;
        }


        if (
          b.status ===
          'Прочитано'
        ) {

          return (
            sum +
            (
              b.hoursRead ??
              b.hours ??
              0
            )
          );

        }


        if (
          b.status === 'Читаю' ||
          b.status === 'На паузе' ||
          b.status === 'Брошено'
        ) {

          return (
            sum +
            (
              b.hoursRead ??
              0
            )
          );

        }


        return sum;

      },
      0
    );
}


// ------------------------------------------------------
// ФАКТИЧЕСКИ ПРОЧИТАННЫЕ СТРАНИЦЫ
//
// Аналогичная логика.
// ------------------------------------------------------

function readPagesForStats(
  period = state.statsYear
) {

  return state.books
    .filter(
      b =>
        b.format !==
        'Аудиокнига'
    )
    .reduce(
      (sum, b) => {

        if (
          b.status ===
          'К прочтению'
        ) {
          return sum;
        }


        if (
          !bookInStatsPeriod(
            b,
            period
          )
        ) {
          return sum;
        }


        // Полностью прочитано
        if (
          b.status ===
          'Прочитано'
        ) {

          return (
            sum +
            (
              b.pagesRead ??
              b.pages ??
              0
            )
          );

        }


        // Читаю / пауза / брошено
        if (
          b.status === 'Читаю' ||
          b.status === 'На паузе' ||
          b.status === 'Брошено'
        ) {

          return (
            sum +
            (
              b.pagesRead ??
              0
            )
          );

        }


        return sum;

      },
      0
    );
}


// ------------------------------------------------------
// Средняя оценка
// Только прочитанные книги
// ------------------------------------------------------

function averageRatingForStats(
  period = state.statsYear
) {

  const books =
    completedBooksForStats(
      period
    );


  const ratings =
    books
      .map(
        b => b.rating
      )
      .filter(
        r => r !== null
      );


  if (!ratings.length) {
    return null;
  }


  return (
    ratings.reduce(
      (a, b) => a + b,
      0
    ) /
    ratings.length
  );
}


// ------------------------------------------------------
// Прочитано по месяцам
// ------------------------------------------------------

function monthlyReadingStats(
  year
) {

  const months =
    Array(12).fill(0);


  state.books
    .filter(
      b =>
        b.status ===
          'Прочитано' &&
        b.end &&
        b.end.getFullYear() ===
          Number(year)
    )
    .forEach(
      b => {

        months[
          b.end.getMonth()
        ]++;

      }
    );


  return months;
}


// ------------------------------------------------------
// Прочитано по годам
// ------------------------------------------------------

function yearlyReadingStats() {

  const years = {};


  state.books
    .filter(
      b =>
        b.status ===
          'Прочитано' &&
        b.end
    )
    .forEach(
      b => {

        const year =
          b.end.getFullYear();


        years[year] =
          (
            years[year] ||
            0
          ) + 1;

      }
    );


  return years;
}


// ------------------------------------------------------
// Книги, с которыми пользователь взаимодействовал
// в выбранном периоде
// ------------------------------------------------------

function activeStatsBooks(
  period = state.statsYear
) {

  return state.books.filter(
    b => {

      if (
        b.status ===
        'К прочтению'
      ) {
        return false;
      }


      return bookInStatsPeriod(
        b,
        period
      );

    }
  );
}


// ------------------------------------------------------
// Распределение форматов
// ------------------------------------------------------

function formatStats(
  period = state.statsYear
) {

  const books =
    activeStatsBooks(
      period
    );


  const audio =
    books.filter(
      b =>
        b.format ===
        'Аудиокнига'
    ).length;


  const regular =
    books.filter(
      b =>
        b.format !==
        'Аудиокнига'
    ).length;


  return {
    audio,
    regular,
    total:
      audio + regular
  };
}


// ------------------------------------------------------
// Распределение оценок
// ------------------------------------------------------

function ratingDistribution(
  period = state.statsYear
) {

  const result = {
    5: 0,
    4: 0,
    3: 0,
    2: 0,
    1: 0
  };


  completedBooksForStats(
    period
  )
    .filter(
      b =>
        b.rating !== null
    )
    .forEach(
      b => {

        // Для диаграммы
        // 4.5 попадёт в 5,
        // 3.5 в 4 и т.д.
        const rounded =
          Math.max(
            1,
            Math.min(
              5,
              Math.round(
                b.rating
              )
            )
          );


        result[rounded]++;

      }
    );


  return result;
}


// ------------------------------------------------------
// Любимые книги
// ------------------------------------------------------

function favoriteBooksForStats(
  period = state.statsYear,
  limit = 5
) {

  return completedBooksForStats(
    period
  )
    .filter(
      b =>
        b.rating !== null
    )
    .sort(
      (a, b) => {

        const ratingDiff =
          b.rating -
          a.rating;


        if (
          ratingDiff !== 0
        ) {
          return ratingDiff;
        }


        return byEndDesc(
          a,
          b
        );

      }
    )
    .slice(
      0,
      limit
    );
}


// ------------------------------------------------------
// Темп чтения
// ------------------------------------------------------

function readingPaceStats(
  period = state.statsYear
) {

  const books =
    completedBooksForStats(
      period
    )
      .filter(
        b =>
          b.duration !== null &&
          b.duration > 0
      );


  if (!books.length) {

    return {
      average: null,
      fastest: null,
      slowest: null
    };

  }


  const average =
    books.reduce(
      (sum, b) =>
        sum +
        b.duration,
      0
    ) /
    books.length;


  const fastest =
    [...books]
      .sort(
        (a, b) =>
          a.duration -
          b.duration
      )[0];


  const slowest =
    [...books]
      .sort(
        (a, b) =>
          b.duration -
          a.duration
      )[0];


  return {
    average,
    fastest,
    slowest
  };
}


// ------------------------------------------------------
// Брошенные книги
// ------------------------------------------------------

function abandonedStats(
  period = state.statsYear
) {

  const books =
    state.books.filter(
      b =>
        b.status ===
          'Брошено' &&
        bookInStatsPeriod(
          b,
          period
        )
    );


  const audioHours =
    books
      .filter(
        b =>
          b.format ===
          'Аудиокнига'
      )
      .reduce(
        (sum, b) =>
          sum +
          (
            b.hoursRead ??
            0
          ),
        0
      );


  const pages =
    books
      .filter(
        b =>
          b.format !==
          'Аудиокнига'
      )
      .reduce(
        (sum, b) =>
          sum +
          (
            b.pagesRead ??
            0
          ),
        0
      );


  return {
    books,
    count:
      books.length,
    audioHours,
    pages
  };
}


// ------------------------------------------------------
// Правильное окончание "книга"
// ------------------------------------------------------

function bookWord(n) {

  const x =
    Math.abs(n) % 100;

  const y =
    x % 10;


  if (
    x >= 11 &&
    x <= 19
  ) {
    return 'книг';
  }


  if (y === 1) {
    return 'книга';
  }


  if (
    y >= 2 &&
    y <= 4
  ) {
    return 'книги';
  }


  return 'книг';
}


function renderStats() {

  const period =
    state.statsYear;


  const completed =
    completedBooksForStats(
      period
    );


  const audioHours =
    listenedHoursForStats(
      period
    );


  const pages =
    readPagesForStats(
      period
    );


  const averageRating =
    averageRatingForStats(
      period
    );


  const formats =
    formatStats(
      period
    );


  const ratings =
    ratingDistribution(
      period
    );


  const favorites =
    favoriteBooksForStats(
      period
    );


  const pace =
    readingPaceStats(
      period
    );


  const abandoned =
    abandonedStats(
      period
    );


  const availableYears =
    getStatsYears();


  const allCounts =
    Object.fromEntries(
      [
        'Прочитано',
        'Читаю',
        'К прочтению',
        'На паузе',
        'Брошено'
      ].map(
        status => [
          status,
          state.books.filter(
            b =>
              b.status ===
              status
          ).length
        ]
      )
    );


  // --------------------------------------------------
  // График
  // --------------------------------------------------

  let readingChart = '';


  // Конкретный год → 12 месяцев
  if (
    period !== 'all'
  ) {

    const months =
      monthlyReadingStats(
        period
      );


    const monthNames = [
      'Янв',
      'Фев',
      'Мар',
      'Апр',
      'Май',
      'Июн',
      'Июл',
      'Авг',
      'Сен',
      'Окт',
      'Ноя',
      'Дек'
    ];


    const maxMonth =
      Math.max(
        1,
        ...months
      );


    readingChart = `
      <div class="chart-list">

        ${months
          .map(
            (value, index) => `
              <div class="chart-row">

                <strong>
                  ${monthNames[index]}
                </strong>

                <div class="bar">

                  <span
                    style="
                      width:${
                        value /
                        maxMonth *
                        100
                      }%
                    "
                  ></span>

                </div>

                <strong>
                  ${value}
                </strong>

              </div>
            `
          )
          .join('')}

      </div>
    `;

  }

  // За всё время → по годам
  else {

    const years =
      yearlyReadingStats();


    const values =
      Object.values(
        years
      );


    const maxYear =
      Math.max(
        1,
        ...values
      );


    readingChart = `
      <div class="chart-list">

        ${
          Object.entries(
            years
          )
            .sort(
              (a, b) =>
                Number(a[0]) -
                Number(b[0])
            )
            .map(
              ([year, value]) => `
                <div class="chart-row">

                  <strong>
                    ${year}
                  </strong>

                  <div class="bar">

                    <span
                      style="
                        width:${
                          value /
                          maxYear *
                          100
                        }%
                      "
                    ></span>

                  </div>

                  <strong>
                    ${value}
                  </strong>

                </div>
              `
            )
            .join('')

          ||

          `
            <div class="empty-state">
              Пока недостаточно
              данных.
            </div>
          `
        }

      </div>
    `;
  }


  // --------------------------------------------------
  // Процент форматов
  // --------------------------------------------------

  const audioPercent =
    formats.total
      ? formats.audio /
        formats.total *
        100
      : 0;


  const bookPercent =
    formats.total
      ? formats.regular /
        formats.total *
        100
      : 0;


  // --------------------------------------------------
  // Максимальное количество оценок
  // --------------------------------------------------

  const maxRatingCount =
    Math.max(
      1,
      ...Object.values(
        ratings
      )
    );


  return `
    <div class="page">

      <!-- ========================================== -->
      <!-- ЗАГОЛОВОК -->
      <!-- ========================================== -->

      <div
        class="section-head"
        style="
          align-items:flex-end;
          margin-bottom:38px;
        "
      >

        <div>

          <div class="eyebrow">
            Моя история чтения
          </div>

          <h1
            style="
              font-size:
                clamp(
                  2.7rem,
                  6vw,
                  4.8rem
                );
              margin-bottom:8px;
            "
          >
            Статистика
          </h1>

          <p class="muted">
            ${
              period === 'all'
                ? 'За всё время'
                : `Итоги ${period} года`
            }
          </p>

        </div>


        <div
          class="toolbar"
          style="
            margin:0;
            width:auto;
          "
        >

          <select
            id="statsYearFilter"
            aria-label="Год статистики"
          >

            ${availableYears
              .map(
                year => `
                  <option
                    value="${year}"
                    ${
                      Number(period) ===
                      year
                        ? 'selected'
                        : ''
                    }
                  >
                    ${year}
                  </option>
                `
              )
              .join('')}

            <option
              value="all"
              ${
                period === 'all'
                  ? 'selected'
                  : ''
              }
            >
              За всё время
            </option>

          </select>

        </div>

      </div>


      <!-- ========================================== -->
      <!-- ГЛАВНЫЕ ПОКАЗАТЕЛИ -->
      <!-- ========================================== -->

      <div class="stats-grid">

        <div class="stat-card">

          <strong>
            ${completed.length}
          </strong>

          <span>
            прочитано книг
          </span>

        </div>


        <div class="stat-card">

          <strong>
            ${oneDec(
              audioHours
            )}
          </strong>

          <span>
            прослушано часов
          </span>

        </div>


        <div class="stat-card">

          <strong>
            ${Math.round(
              pages
            )}
          </strong>

          <span>
            прочитано страниц
          </span>

        </div>


        <div class="stat-card">

          <strong>
            ${
              averageRating !== null
                ? oneDec(
                    averageRating
                  )
                : '—'
            }
          </strong>

          <span>
            средняя оценка
          </span>

        </div>

      </div>


      <!-- ========================================== -->
      <!-- ЧТЕНИЕ ПО МЕСЯЦАМ / ГОДАМ -->
      <!-- ========================================== -->

      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Динамика
            </div>

            <h2>
              ${
                period === 'all'
                  ? 'Прочитано по годам'
                  : 'Мой читательский год'
              }
            </h2>

          </div>

        </div>

        ${readingChart}

      </section>


      <!-- ========================================== -->
      <!-- КАК Я ЧИТАЮ -->
      <!-- ========================================== -->

      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Форматы
            </div>

            <h2>
              Как я читаю
            </h2>

          </div>

        </div>


        <div class="stats-grid">

          <div class="stat-card">

            <strong>
              🎧 ${formats.audio}
            </strong>

            <span>
              ${bookWord(
                formats.audio
              )} в аудио
            </span>

            <div
              class="progress"
              style="
                margin-top:16px
              "
            >
              <span
                style="
                  width:${audioPercent}%
                "
              ></span>
            </div>

            <div
              class="muted"
              style="
                margin-top:9px
              "
            >
              ${oneDec(
                audioHours
              )} ч прослушано
            </div>

          </div>


          <div class="stat-card">

            <strong>
              📖 ${formats.regular}
            </strong>

            <span>
              ${
                bookWord(
                  formats.regular
                )
              } в тексте
            </span>

            <div
              class="progress"
              style="
                margin-top:16px
              "
            >
              <span
                style="
                  width:${bookPercent}%
                "
              ></span>
            </div>

            <div
              class="muted"
              style="
                margin-top:9px
              "
            >
              ${Math.round(
                pages
              )} стр. прочитано
            </div>

          </div>

        </div>

      </section>


      <!-- ========================================== -->
      <!-- ОЦЕНКИ -->
      <!-- ========================================== -->

      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Впечатления
            </div>

            <h2>
              Как я оцениваю книги
            </h2>

          </div>


          ${
            averageRating !== null
              ? `
                <span class="muted">
                  Средняя оценка —
                  ${oneDec(
                    averageRating
                  )}
                </span>
              `
              : ''
          }

        </div>


        <div class="chart-list">

          ${[
            5,
            4,
            3,
            2,
            1
          ]
            .map(
              rating => `
                <div class="chart-row">

                  <strong
                    style="
                      min-width:78px
                    "
                  >
                    ${'★'.repeat(
                      rating
                    )}
                  </strong>

                  <div class="bar">

                    <span
                      style="
                        width:${
                          ratings[rating] /
                          maxRatingCount *
                          100
                        }%
                      "
                    ></span>

                  </div>

                  <strong>
                    ${ratings[rating]}
                  </strong>

                </div>
              `
            )
            .join('')}

        </div>

      </section>


      <!-- ========================================== -->
      <!-- ЛЮБИМЫЕ КНИГИ -->
      <!-- ========================================== -->

      ${
        favorites.length
          ? `
            <section class="section">

              <div class="section-head">

                <div>

                  <div class="eyebrow">
                    Лучшие оценки
                  </div>

                  <h2>
                    ${
                      period === 'all'
                        ? 'Любимые книги'
                        : 'Любимые книги года'
                    }
                  </h2>

                </div>

              </div>


              <div class="cover-row">

                ${favorites
                  .map(
                    coverCard
                  )
                  .join('')}

              </div>

            </section>
          `
          : ''
      }


      <!-- ========================================== -->
      <!-- ТЕМП ЧТЕНИЯ -->
      <!-- ========================================== -->

      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Время
            </div>

            <h2>
              Темп чтения
            </h2>

          </div>

        </div>


        ${
          pace.average !== null
            ? `
              <div class="stats-grid">

                <div class="stat-card">

                  <strong>
                    ${oneDec(
                      pace.average
                    )}
                  </strong>

                  <span>
                    дней в среднем
                    на книгу
                  </span>

                </div>


                <div
                  class="stat-card"
                  data-book="${pace.fastest.id}"
                  role="button"
                  tabindex="0"
                >

                  <strong>
                    ${pace.fastest.duration}
                  </strong>

                  <span>
                    ${
                      pace.fastest.duration === 1
                        ? 'день'
                        : 'дней'
                    }
                    — быстрее всего
                  </span>

                  <div
                    class="book-title"
                    style="
                      margin-top:10px;
                      font-size:.95rem
                    "
                  >
                    ${esc(
                      pace.fastest.title
                    )}
                  </div>

                </div>


                <div
                  class="stat-card"
                  data-book="${pace.slowest.id}"
                  role="button"
                  tabindex="0"
                >

                  <strong>
                    ${pace.slowest.duration}
                  </strong>

                  <span>
                    ${
                      pace.slowest.duration === 1
                        ? 'день'
                        : 'дней'
                    }
                    — дольше всего
                  </span>

                  <div
                    class="book-title"
                    style="
                      margin-top:10px;
                      font-size:.95rem
                    "
                  >
                    ${esc(
                      pace.slowest.title
                    )}
                  </div>

                </div>

              </div>
            `
            : `
              <div class="empty-state">
                Пока недостаточно данных
                о длительности чтения.
              </div>
            `
        }

      </section>


      <!-- ========================================== -->
      <!-- БРОШЕННЫЕ КНИГИ -->
      <!-- ========================================== -->

      ${
        abandoned.count
          ? `
            <section class="section">

              <div class="section-head">

                <div>

                  <div class="eyebrow">
                    Не дочитано
                  </div>

                  <h2>
                    Брошенные книги
                  </h2>

                </div>

              </div>


              <div class="stats-grid">

                <div class="stat-card">

                  <strong>
                    ${abandoned.count}
                  </strong>

                  <span>
                    ${bookWord(
                      abandoned.count
                    )} брошено
                  </span>

                </div>


                <div class="stat-card">

                  <strong>
                    ${oneDec(
                      abandoned.audioHours
                    )}
                  </strong>

                  <span>
                    часов всё равно
                    прослушано
                  </span>

                </div>


                <div class="stat-card">

                  <strong>
                    ${Math.round(
                      abandoned.pages
                    )}
                  </strong>

                  <span>
                    страниц всё равно
                    прочитано
                  </span>

                </div>

              </div>


              <div
                class="cover-row"
                style="
                  margin-top:30px
                "
              >

                ${abandoned.books
                  .slice(
                    0,
                    6
                  )
                  .map(
                    coverCard
                  )
                  .join('')}

              </div>

            </section>
          `
          : ''
      }


      <!-- ========================================== -->
      <!-- ЗА ВСЁ ВРЕМЯ -->
      <!-- ========================================== -->

      <section class="section">

        <div class="section-head">

          <div>

            <div class="eyebrow">
              Вся коллекция
            </div>

            <h2>
              За всё время
            </h2>

          </div>

        </div>


        <div class="library-summary">

          <div class="big-total">

            <strong>
              ${state.books.length}
            </strong>

            <span>
              книг всего
            </span>

          </div>


          <div class="status-list">

            ${Object.entries(
              allCounts
            )
              .map(
                ([status, value]) => `
                  <div class="status-item">

                    <span>
                      ${esc(status)}
                    </span>

                    <strong>
                      ${value}
                    </strong>

                  </div>
                `
              )
              .join('')}

          </div>

        </div>

      </section>

    </div>
  `;
}


// ======================================================
// РЕНДЕР
// ======================================================

function render() {

  const route =
    (
      location.hash ||
      '#home'
    ).slice(1);


  state.route =
    [
      'home',
      'library',
      'series',
      'stats'
    ].includes(route)
      ? route
      : 'home';


  $('#app').innerHTML =

    state.route === 'home'
      ? renderHome()

      : state.route === 'library'
      ? renderLibrary()

      : state.route === 'series'
      ? renderSeries()

      : renderStats();


  $$('.nav a').forEach(
    a =>
      a.classList.toggle(
        'active',
        a.dataset.route ===
        state.route
      )
  );


  bindDynamic();
}


// ======================================================
// МОДАЛЬНОЕ ОКНО КНИГИ
// ======================================================

function showBook(id) {

  const b =
    state.books.find(
      x =>
        x.id ===
        Number(id)
    );


  if (!b) return;


  const volume =
    b.format === 'Аудиокнига'

      ? (
          b.hours
            ? `${oneDec(b.hours)} ч`
            : '—'
        )

      : (
          b.pages
            ? `${b.pages} стр.`
            : '—'
        );


  const progress =
    b.progress !== null

      ? (
          `${Math.round(
            b.progress
          )}%` +
          (
            b.progressText
              ? ` · ${b.progressText}`
              : ''
          )
        )

      : (
          b.status === 'Прочитано'
            ? '100%'
            : '—'
        );


  $('#bookDialog').innerHTML = `
    <button
      class="
        icon-btn
        dialog-close
      "
      onclick="
        document
          .querySelector('#bookDialog')
          .close()
      "
    >
      ×
    </button>


    <div class="dialog-inner">

      ${coverHTML(b)}


      <div>

        <div class="eyebrow">
          ${esc(
            b.status ||
            'Книга'
          )}
        </div>


        <h2
          style="
            font-size:2.5rem;
            margin-top:6px
          "
        >
          ${esc(b.title)}
        </h2>


        <p class="muted">
          ${esc(b.author)}
        </p>


        ${
          b.series
            ? `
              <p>

                <strong>
                  ${esc(
                    b.series
                  )}
                </strong>

                ${
                  b.seriesNo
                    ? ` · книга ${b.seriesNo}`
                    : ''
                }

              </p>
            `
            : ''
        }


        ${
          b.progress !== null &&
          b.status !== 'Прочитано'

            ? `
              <div class="progress-copy">

                <span>
                  ${esc(
                    b.progressText
                  )}
                </span>

                <strong>
                  ${Math.round(
                    b.progress
                  )}%
                </strong>

              </div>


              <div class="progress">

                <span
                  style="
                    width:${b.progress}%
                  "
                ></span>

              </div>
            `

            : ''
        }


        <div class="dialog-meta">

          <div class="meta-box">

            <small>
              Формат
            </small>

            <strong>
              ${esc(
                b.format ||
                '—'
              )}
            </strong>

          </div>


          <div class="meta-box">

            <small>
              Объём
            </small>

            <strong>
              ${esc(volume)}
            </strong>

          </div>


          <div class="meta-box">

            <small>
              Начало
            </small>

            <strong>
              ${esc(
                fmtDate(
                  b.start
                )
              )}
            </strong>

          </div>


          <div class="meta-box">

            <small>
              Конец
            </small>

            <strong>
              ${esc(
                fmtDate(
                  b.end
                )
              )}
            </strong>

          </div>


          <div class="meta-box">

            <small>
              Длительность
            </small>

            <strong>
              ${
                b.duration !== null
                  ? `${b.duration} дн.`
                  : '—'
              }
            </strong>

          </div>


          <div class="meta-box">

            <small>
              Оценка
            </small>

            <strong>
              ${
                b.rating !== null
                  ? `
                    ${ratingStars(
                      b.rating
                    )}
                    ${oneDec(
                      b.rating
                    )}
                  `
                  : '—'
              }
            </strong>

          </div>

        </div>

      </div>

    </div>
  `;


  $('#bookDialog')
    .showModal();
}


// ======================================================
// СЛУЧАЙНАЯ КНИГА
// ======================================================

function randomUnread() {

  const arr =
    state.books.filter(
      b =>
        b.status ===
        'К прочтению'
    );


  if (!arr.length) {
    return;
  }


  const random =
    arr[
      Math.floor(
        Math.random() *
        arr.length
      )
    ];


  showBook(random.id);
}


// ======================================================
// СОБЫТИЯ ДИНАМИЧЕСКИХ ЭЛЕМЕНТОВ
// ======================================================

function bindDynamic() {

  $$('[data-book]')
    .forEach(
      el => {

        el.addEventListener(
          'click',
          () =>
            showBook(
              el.dataset.book
            )
        );


        el.addEventListener(
          'keydown',
          e => {

            if (
              e.key === 'Enter'
            ) {

              showBook(
                el.dataset.book
              );

            }

          }
        );

      }
    );


  $$('[data-go]')
    .forEach(
      el =>
        el.addEventListener(
          'click',
          () =>
            location.hash =
              el.dataset.go
        )
    );


  $('#showPaused')
    ?.addEventListener(
      'click',
      () => {

        const g =
          $('#pausedGrid');

        g.style.display =
          g.style.display ===
          'none'
            ? 'grid'
            : 'none';

      }
    );


  $('#randomBook')
    ?.addEventListener(
      'click',
      randomUnread
    );


  $('#libSearch')
    ?.addEventListener(
      'input',
      e => {

        state.filters.q =
          e.target.value;

        render();

        $('#libSearch')
          ?.focus();


        try {

          $('#libSearch')
            .setSelectionRange(
              state.filters.q.length,
              state.filters.q.length
            );

        }
        catch {}

      }
    );


  $('#statusFilter')
    ?.addEventListener(
      'change',
      e => {

        state.filters.status =
          e.target.value;

        render();

      }
    );


  $('#formatFilter')
    ?.addEventListener(
      'change',
      e => {

        state.filters.format =
          e.target.value;

        render();

      }
    );


  $('#sortFilter')
    ?.addEventListener(
      'change',
      e => {

        state.filters.sort =
          e.target.value;

        render();

      }
    );

  $('#statsYearFilter')
  ?.addEventListener(
    'change',
    e => {

      state.statsYear =
        e.target.value === 'all'
          ? 'all'
          : Number(
              e.target.value
            );

      render();

    }
  );
  
}


// ======================================================
// ГЛОБАЛЬНЫЙ ПОИСК
// ======================================================

function bindGlobalSearch() {

  const dlg =
    $('#searchDialog');

  const input =
    $('#globalSearch');

  const results =
    $('#searchResults');


  $('#openSearch')
    .addEventListener(
      'click',
      () => {

        dlg.showModal();

        setTimeout(
          () =>
            input.focus(),
          50
        );

      }
    );


  input.addEventListener(
    'input',
    () => {

      const q =
        input.value
          .trim()
          .toLowerCase();


      const arr =
        q
          ? state.books
              .filter(
                b =>
                  [
                    b.title,
                    b.author,
                    b.series
                  ].some(
                    x =>
                      x
                        .toLowerCase()
                        .includes(q)
                  )
              )
              .slice(0, 12)

          : [];


      results.innerHTML =

        arr
          .map(
            b => `
              <div
                class="search-result"
                data-search-book="${b.id}"
                role="button"
              >

                ${coverHTML(
                  b,
                  false
                )}

                <div>

                  <strong>
                    ${esc(
                      b.title
                    )}
                  </strong>

                  <div class="muted">
                    ${esc(
                      b.author
                    )}
                  </div>

                </div>

              </div>
            `
          )
          .join('')

        ||

        (
          q
            ? `
              <div class="empty-state">
                Ничего не найдено
              </div>
            `
            : `
              <div
                class="muted"
                style="padding:16px"
              >
                Начни вводить
                название,
                автора или серию.
              </div>
            `
        );


      $$(
        '[data-search-book]',
        results
      ).forEach(
        x =>
          x.addEventListener(
            'click',
            () => {

              dlg.close();

              showBook(
                x.dataset.searchBook
              );

            }
          )
      );

    }
  );
}


// ======================================================
// ПЕРЕХОД МЕЖДУ СТРАНИЦАМИ
// ======================================================

function handleRouteChange() {

  render();


  // При переходе между
  // Главная / Библиотека /
  // Серии / Статистика
  // всегда поднимаемся наверх

  requestAnimationFrame(
    () => {

      window.scrollTo({
        top: 0,
        left: 0,
        behavior: 'auto'
      });

    }
  );
}


// ======================================================
// ЗАПУСК
// ======================================================

async function init() {

  try {

    await loadBooks();


    // Отключаем автоматическое
    // восстановление позиции
    // прокрутки браузером

    if (
      'scrollRestoration'
      in history
    ) {

      history.scrollRestoration =
        'manual';

    }


    render();


    window.scrollTo({
      top: 0,
      left: 0,
      behavior: 'auto'
    });


    bindGlobalSearch();


    window.addEventListener(
      'hashchange',
      handleRouteChange
    );

  }

  catch (e) {

    $('#app').innerHTML = `
      <div class="empty-state">

        <h2>
          Не удалось загрузить Excel
        </h2>

        <p>
          ${esc(e.message)}
        </p>

        <p>
          Проверь, что
          <strong>
            books.xlsx
          </strong>
          лежит рядом с
          index.html и сайт
          открыт через GitHub Pages
          или локальный
          веб-сервер.
        </p>

      </div>
    `;


    console.error(e);

  }
}


init();
