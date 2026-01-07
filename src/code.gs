/***** Goal Tiles — Calendar Add-on (large display, persistent, reorder, color picker)
 * Changes in this version:
 * - Much larger title banner (taller + bigger font tuned to sidebar width)
 * - Big section header bars
 * - Body text rendered as large SVG blocks with **newline preservation**
 * - Edit/Delete hidden in reorder mode (and tap-to-edit disabled)
 */

/* ===================== Config ===================== */

const FILE_NAME = 'GoalTiles (Do Not Delete).json';
const FILE_MIME = 'application/json';
const DEFAULT_COLOR = '#6C2BD9';

// Pagination (set to 0 to disable)
const PAGE_SIZE = 0; // e.g., 6 to paginate, or 0 for "show all"

// Cache settings
function cache_() { return CacheService.getUserCache(); }
const CACHE_TILES_KEY   = 'goal_tiles_v1';
const CACHE_TTL_SECONDS = 300;   // 5 minutes
const CACHE_DIRTY_KEY   = 'goal_tiles_dirty'; // marks unflushed reorder changes

// Banner/text image cache
const CACHE_BANNER_PREFIX = 'goal_banner_';
const CACHE_BANNER_TTL    = 3600; // 1 hour

/* ===================== Entry ===================== */

function onHomepage(e) { return buildListCard_(); }

/* ===================== List / Tiles View ===================== */
/**
 * @param {string=} flash
 * @param {boolean=} reorder
 * @param {string=} offsetStr
 */
function buildListCard_(flash, reorder, offsetStr) {
  const isReorder = !!reorder;
  const offset = Math.max(0, parseInt(offsetStr || '0', 10));

  const tiles = loadTilesCached_().map(hydrateTile_);

  const card = CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle('Goal Tiles')
        .setSubtitle(isReorder ? 'Reorder mode — use arrows' : 'Tap ＋ New to add • Use Edit to modify')
    );

  if (flash) {
    card.addSection(
      CardService.newCardSection().addWidget(
        CardService.newKeyValue().setContent(flash)
      )
    );
  }

  // Controls
  const controls = CardService.newCardSection()
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText('＋ New tile')
            .setOnClickAction(CardService.newAction().setFunctionName('handleNew'))
        )
        .addButton(
          CardService.newTextButton()
            .setText(isReorder ? 'Done' : 'Reorder')
            .setOnClickAction(
              CardService.newAction().setFunctionName('handleToggleReorder')
                .setParameters({ state: String(!isReorder), o: String(offset) })
            )
        )
    );
  card.addSection(controls);

  if (!tiles.length) {
    card.addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('No tiles yet. Click “＋ New tile”.')
      )
    );
    return card.build();
  }

  // Pagination window
  const start = PAGE_SIZE > 0 ? offset : 0;
  const end   = PAGE_SIZE > 0 ? Math.min(offset + PAGE_SIZE, tiles.length) : tiles.length;
  const page  = tiles.slice(start, end);

  page.forEach((t, idxOnPage) => {
    const absoluteIdx = start + idxOnPage;
    const s = CardService.newCardSection();

    // MUCH larger banner (tuned for ~360–400px sidebar width)
    const bannerUrl = getBanner_(t.title, t.color);
    s.addWidget(CardService.newImage().setImageUrl(bannerUrl));

    // Actions & handlers
    const editAction = CardService.newAction()
      .setFunctionName('handleEditExisting')
      .setParameters({ index: String(absoluteIdx), o: String(offset) });

    const delAction = CardService.newAction()
      .setFunctionName('handleDelete')
      .setParameters({ index: String(absoluteIdx), o: String(offset) });

    // === Goal ===
    s.addWidget(CardService.newImage().setImageUrl(getSectionHeader_('Goal')));
    s.addWidget(CardService.newImage().setImageUrl(getTextBlock_(t.goal || '—')));

    // === Road Map ===
    s.addWidget(CardService.newImage().setImageUrl(getSectionHeader_('Road Map')));
    s.addWidget(CardService.newImage().setImageUrl(getTextBlock_(t.roadmap || '—')));

    // === First steps ===
    s.addWidget(CardService.newImage().setImageUrl(getSectionHeader_('First steps')));
    s.addWidget(CardService.newImage().setImageUrl(getTextBlock_(t.firststeps || '—')));

    // Edit/Delete only when NOT reordering
    if (!isReorder) {
      s.addWidget(
        CardService.newButtonSet()
          .addButton(CardService.newTextButton().setText('Edit').setOnClickAction(editAction))
          .addButton(CardService.newTextButton().setText('Delete').setOnClickAction(delAction))
      );
    }

    // Reorder controls (debounced writes; flush on Done)
    if (isReorder) {
      const up = CardService.newAction().setFunctionName('handleMoveUp').setParameters({ index: String(absoluteIdx), o: String(offset) });
      const down = CardService.newAction().setFunctionName('handleMoveDown').setParameters({ index: String(absoluteIdx), o: String(offset) });
      const top = CardService.newAction().setFunctionName('handleMoveTop').setParameters({ index: String(absoluteIdx), o: String(offset) });
      const bottom = CardService.newAction().setFunctionName('handleMoveBottom').setParameters({ index: String(absoluteIdx), o: String(offset) });

      s.addWidget(
        CardService.newButtonSet()
          .addButton(CardService.newTextButton().setText('↑ Move up').setOnClickAction(up))
          .addButton(CardService.newTextButton().setText('↓ Move down').setOnClickAction(down))
          .addButton(CardService.newTextButton().setText('Top').setOnClickAction(top))
          .addButton(CardService.newTextButton().setText('Bottom').setOnClickAction(bottom))
      );
    }

    card.addSection(s);
  });

  // Pager controls
  if (PAGE_SIZE > 0) {
    const pager = CardService.newCardSection();
    if (offset > 0) {
      pager.addWidget(
        CardService.newTextButton()
          .setText('◀︎ Newer')
          .setOnClickAction(CardService.newAction().setFunctionName('handlePage')
            .setParameters({ o: String(Math.max(0, offset - PAGE_SIZE)), r: String(isReorder) }))
      );
    }
    if (offset + PAGE_SIZE < tiles.length) {
      pager.addWidget(
        CardService.newTextButton()
          .setText('Older ▶︎')
          .setOnClickAction(CardService.newAction().setFunctionName('handlePage')
            .setParameters({ o: String(offset + PAGE_SIZE), r: String(isReorder) }))
      );
    }
    card.addSection(pager);
  }

  return card.build();
}

/* ===================== Edit Form (color picker + live preview) ===================== */

function buildEditCard_(vals, index, offset) {
  const v = hydrateTile_(vals || {});
  const isEditing = Number.isInteger(index);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(isEditing ? 'Edit tile' : 'New tile'));

  // Live banner preview (larger)
  card.addSection(
    CardService.newCardSection().addWidget(
      CardService.newImage().setImageUrl(getBanner_(v.title, v.color))
    )
  );

  const section = CardService.newCardSection();

  section.addWidget(CardService.newTextInput()
    .setFieldName('title').setTitle('Project / Title').setValue(v.title));

  section.addWidget(CardService.newTextInput()
    .setFieldName('goal').setTitle('Goal').setMultiline(true).setValue(v.goal));

  section.addWidget(CardService.newTextInput()
    .setFieldName('roadmap').setTitle('Road Map').setMultiline(true).setValue(v.roadmap));

  section.addWidget(CardService.newTextInput()
    .setFieldName('firststeps').setTitle('First steps').setMultiline(true).setValue(v.firststeps));

  // Color picker
  const palette = getPalette_();
  const colorSelect = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('colorPick')
    .setTitle('Banner color');

  palette.forEach(p => colorSelect.addItem(`${p.name} (${p.hex})`, p.hex, p.hex === v.color));

  const changeAction = CardService.newAction().setFunctionName('handleColorChange');
  if (isEditing) changeAction.setParameters({ index: String(index), o: String(offset || 0) });
  colorSelect.setOnChangeAction(changeAction);

  section.addWidget(colorSelect);

  // Optional custom hex override
  section.addWidget(
    CardService.newTextInput()
      .setFieldName('colorCustom')
      .setTitle('Custom hex (optional)')
      .setHint('#RRGGBB or #RGB')
      .setValue(v.color && !inPalette_(v.color) ? v.color : '')
  );

  // Save / Cancel
  const saveAction = CardService.newAction()
    .setFunctionName(isEditing ? 'handleSaveExisting' : 'handleSaveNew')
    .setParameters(isEditing ? { index: String(index), o: String(offset || 0) } : { o: String(offset || 0) });

  section.addWidget(
    CardService.newButtonSet()
      .addButton(CardService.newTextButton().setText('Save').setOnClickAction(saveAction))
      .addButton(CardService.newTextButton().setText('Cancel').setOnClickAction(
        CardService.newAction().setFunctionName('handleCancel').setParameters({ o: String(offset || 0) })
      ))
  );

  card.addSection(section);
  return card.build();
}

/* ===================== Actions: nav & CRUD ===================== */

function handleNew(e) {
  const o = e.parameters && e.parameters.o || '0';
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildEditCard_({}, null, parseInt(o, 10))))
    .build();
}

function handleCancel(e) {
  const o = e.parameters && e.parameters.o || '0';
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, false, o)))
    .build();
}

function handleToggleReorder(e) {
  const want = (e.parameters && e.parameters.state === 'true');
  const o = e.parameters && e.parameters.o || '0';
  // Leaving reorder? Flush debounced changes
  if (!want && isDirty_()) {
    const tiles = loadTilesCached_();
    saveTilesCached_(tiles);
    clearDirty_();
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, want, o)))
    .build();
}

function handleEditExisting(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const tiles = loadTilesCached_().map(hydrateTile_);
  const val = tiles[idx] || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildEditCard_(val, idx, parseInt(o, 10))))
    .build();
}

function handleSaveNew(e) {
  const o = e.parameters && e.parameters.o || '0';
  const f = hydrateTile_(formValues_(e));
  const tiles = loadTilesCached_();
  tiles.unshift(Object.assign(f, { updatedAt: new Date().toISOString() }));
  saveTilesCached_(tiles); // write-through
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_('Saved ✅', false, o)))
    .build();
}

function handleSaveExisting(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const f = hydrateTile_(formValues_(e));
  const tiles = loadTilesCached_();
  if (tiles[idx]) tiles[idx] = Object.assign(f, { updatedAt: new Date().toISOString() });
  saveTilesCached_(tiles); // write-through
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_('Updated ✅', false, o)))
    .build();
}

function handleDelete(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const tiles = loadTilesCached_();
  if (idx >= 0 && idx < tiles.length) tiles.splice(idx, 1);
  saveTilesCached_(tiles); // write-through
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_('Deleted 🗑️', false, o)))
    .build();
}

/* ===================== Actions: Reorder (debounced writes) ===================== */

function handleMoveUp(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const tiles = loadTilesCached_();
  if (idx > 0 && idx < tiles.length) {
    const tmp = tiles[idx - 1]; tiles[idx - 1] = tiles[idx]; tiles[idx] = tmp;
    cache_().put(CACHE_TILES_KEY, JSON.stringify(tiles), CACHE_TTL_SECONDS);
    markDirty_();
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, true, o)))
    .build();
}

function handleMoveDown(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const tiles = loadTilesCached_();
  if (idx >= 0 && idx < tiles.length - 1) {
    const tmp = tiles[idx + 1]; tiles[idx + 1] = tiles[idx]; tiles[idx] = tmp;
    cache_().put(CACHE_TILES_KEY, JSON.stringify(tiles), CACHE_TTL_SECONDS);
    markDirty_();
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, true, o)))
    .build();
}

function handleMoveTop(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const tiles = loadTilesCached_();
  if (idx > 0 && idx < tiles.length) {
    const [item] = tiles.splice(idx, 1);
    tiles.unshift(item);
    cache_().put(CACHE_TILES_KEY, JSON.stringify(tiles), CACHE_TTL_SECONDS);
    markDirty_();
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, true, o)))
    .build();
}

function handleMoveBottom(e) {
  const idx = parseInt(e.parameters.index, 10);
  const o = e.parameters && e.parameters.o || '0';
  const tiles = loadTilesCached_();
  if (idx >= 0 && idx < tiles.length - 1) {
    const [item] = tiles.splice(idx, 1);
    tiles.push(item);
    cache_().put(CACHE_TILES_KEY, JSON.stringify(tiles), CACHE_TTL_SECONDS);
    markDirty_();
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, true, o)))
    .build();
}

/* ===================== Color Picker: live preview ===================== */

function handleColorChange(e) {
  const f = (e && e.commonEventObject && e.commonEventObject.formInputs) || {};
  const idxParam = e && e.parameters && e.parameters.index;
  const o = e && e.parameters && e.parameters.o || '0';
  const idx = (typeof idxParam === 'string' && idxParam !== '') ? parseInt(idxParam, 10) : null;

  const vals = {
    title: getStr_(f, 'title'),
    goal: getStr_(f, 'goal'),
    roadmap: getStr_(f, 'roadmap'),
    firststeps: getStr_(f, 'firststeps'),
    color: normColor_(getStr_(f, 'colorPick') || DEFAULT_COLOR)
  };

  const nav = CardService.newNavigation().updateCard(
    Number.isInteger(idx) ? buildEditCard_(vals, idx, parseInt(o, 10)) : buildEditCard_(vals, null, parseInt(o, 10))
  );
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

/* ===================== Helpers: form + palette + storage ===================== */

function formValues_(e) {
  const f = (e && e.commonEventObject && e.commonEventObject.formInputs) || {};
  const pick = getStr_(f, 'colorPick');
  const custom = getStr_(f, 'colorCustom');
  const chosen = custom && normColor_(custom) ? normColor_(custom) : (pick || DEFAULT_COLOR);
  return {
    title: getStr_(f, 'title'),
    goal: getStr_(f, 'goal'),
    roadmap: getStr_(f, 'roadmap'),
    firststeps: getStr_(f, 'firststeps'),
    color: normColor_(chosen)
  };
}

function getPalette_() {
  return [
    { name: 'Purple',  hex: '#6C2BD9' },
    { name: 'Sky',     hex: '#0EA5E9' },
    { name: 'Orange',  hex: '#F97316' },
    { name: 'Emerald', hex: '#10B981' },
    { name: 'Rose',    hex: '#F43F5E' },
    { name: 'Amber',   hex: '#F59E0B' },
    { name: 'Slate',   hex: '#334155' }
  ];
}

function inPalette_(hex) {
  const h = hex ? hex.toUpperCase() : '';
  return getPalette_().some(p => p.hex === h);
}

function hydrateTile_(t) {
  let title = (t && typeof t.title === 'string' ? t.title.trim() : '') || 'Untitled';
  let color = (t && typeof t.color === 'string' ? t.color.trim() : '') || DEFAULT_COLOR;
  color = normColor_(color);
  return {
    title,
    goal: (t && t.goal) || '',
    roadmap: (t && t.roadmap) || '',
    firststeps: (t && t.firststeps) || '',
    color,
    updatedAt: t && t.updatedAt ? t.updatedAt : ''
  };
}

function getStr_(formInputs, name) {
  try {
    const o = formInputs[name];
    if (!o || !o.stringInputs) return '';
    const arr = o.stringInputs.value || [];
    return (arr.length ? arr[0] : '').toString();
  } catch (_) { return ''; }
}

function normColor_(c) {
  c = (c || '').trim();
  const ok = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(c);
  if (!ok) return DEFAULT_COLOR;
  if (c.length === 4) c = '#' + c[1]+c[1] + c[2]+c[2] + c[3]+c[3];
  return c.toUpperCase();
}

/* ----------------- Storage (Drive JSON) ----------------- */

function getFile_() {
  const it = DriveApp.getFilesByName(FILE_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFile(FILE_NAME, '[]', FILE_MIME);
}

function loadTiles_() {
  try {
    const txt = getFile_().getBlob().getDataAsString('UTF-8') || '[]';
    const arr = JSON.parse(txt);
    return Array.isArray(arr) ? arr : [];
  } catch (e) { return []; }
}

function saveTiles_(arr) {
  getFile_().setContent(JSON.stringify(arr));
}

/* ---------- Cached load/save (fast path) ---------- */

function loadTilesCached_() {
  const c = cache_().get(CACHE_TILES_KEY);
  if (c) {
    try { return JSON.parse(c); } catch (_) {}
  }
  const arr = loadTiles_();
  cache_().put(CACHE_TILES_KEY, JSON.stringify(arr), CACHE_TTL_SECONDS);
  return arr;
}

function saveTilesCached_(arr) {
  saveTiles_(arr); // write-through to Drive
  cache_().put(CACHE_TILES_KEY, JSON.stringify(arr), CACHE_TTL_SECONDS);
}

/* ---------- Debounce helpers for reorder ---------- */

function markDirty_()  { cache_().put(CACHE_DIRTY_KEY, '1', 300); }
function clearDirty_() { cache_().remove(CACHE_DIRTY_KEY); }
function isDirty_()    { return cache_().get(CACHE_DIRTY_KEY) === '1'; }

/* ----------------- SVG Generators (banners, headers, large text) ----------------- */

/**
 * We design SVGs at ~360px width (typical sidebar) to avoid auto-downscaling.
 * If your sidebar is wider, they will scale up cleanly.
 */
const SVG_BASE_WIDTH = 360;

function getBanner_(title, color) {
  const key = CACHE_BANNER_PREFIX + Utilities.base64Encode('banner|' + title + '|' + color + '|v2');
  const c = cache_().get(key);
  if (c) return c;
  const uri = makeBannerDataURI_(title, color);
  cache_().put(key, uri, CACHE_BANNER_TTL);
  return uri;
}

function makeBannerDataURI_(title, color) {
  const t = escapeHtml_(title || 'Untitled');
  const c = color || DEFAULT_COLOR;
  // Taller banner + bigger font (height 120)
  const w = SVG_BASE_WIDTH, h = 120;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
       <rect x="0" y="0" width="${w}" height="${h}" rx="14" ry="14" fill="${c}"/>
       <text x="${w/2}" y="${Math.round(h*0.60)}" font-family="Roboto,Arial,sans-serif"
             font-size="28" font-weight="800" text-anchor="middle" fill="#FFFFFF">${t}</text>
     </svg>`;
  return svgToDataUri_(svg);
}

function getSectionHeader_(label) {
  const key = CACHE_BANNER_PREFIX + Utilities.base64Encode('section|' + label + '|v2');
  const c = cache_().get(key);
  if (c) return c;
  const uri = makeSectionHeaderDataURI_(label);
  cache_().put(key, uri, CACHE_BANNER_TTL);
  return uri;
}

function makeSectionHeaderDataURI_(label) {
  const t = escapeHtml_(label || '');
  const w = SVG_BASE_WIDTH, h = 52;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
       <rect x="0" y="0" width="${w}" height="${h}" rx="10" ry="10" fill="#111827"/>
       <text x="16" y="${Math.round(h*0.66)}" font-family="Roboto,Arial,sans-serif"
             font-size="20" font-weight="700" text-anchor="start" fill="#FFFFFF">${t}</text>
     </svg>`;
  return svgToDataUri_(svg);
}

/**
 * Large body text block with **newline preservation**.
 * Each user-entered line is wrapped independently, blank lines are kept.
 */
function getTextBlock_(text) {
  const normalized = (typeof text === 'string' ? text : '').replace(/\r\n/g, '\n');
  const key = CACHE_BANNER_PREFIX + Utilities.base64Encode('txt|' + normalized + '|v4'); // bump to v4 for this logic
  const c = cache_().get(key);
  if (c) return c;
  const uri = makeTextBlockDataURI_(normalized);
  cache_().put(key, uri, CACHE_BANNER_TTL);
  return uri;
}

function makeTextBlockDataURI_(text) {
  const w = SVG_BASE_WIDTH;
  const fontSize = 18;             // large, readable
  const lineHeight = Math.round(fontSize * 1.5);
  const maxChars = 42;             // wrap width per visual line
  const padding = 14;

  // Preserve user-entered newlines (split, then wrap each line separately)
  const userLines = String(text || '—').replace(/\r\n/g, '\n').split('\n');

  // Build wrapped output lines while keeping paragraph breaks
  const outLines = [];
  userLines.forEach((ln) => {
    const trimmed = ln.trim();
    if (trimmed === '') {
      // blank line = spacer (keep one empty line)
      outLines.push('');
    } else {
      wrapText_(trimmed, maxChars).forEach(wl => outLines.push(wl));
    }
  });

  const h = padding*2 + Math.max(1, outLines.length) * lineHeight;

  const linesSvg = outLines.map((ln, i) => {
    if (ln === '') return '';
    return `<text x="${padding}" y="${padding + lineHeight*(i+0.8)}"
              font-family="Roboto,Arial,sans-serif"
              font-size="${fontSize}" font-weight="500"
              text-anchor="start" fill="#111827">${escapeXmlText_(ln)}</text>`;
  }).join('');

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
       <rect x="0" y="0" width="${w}" height="${h}" rx="10" ry="10" fill="#F3F4F6"/>
       ${linesSvg}
     </svg>`;
  return svgToDataUri_(svg);
}

// Word-wrap a single logical line to ~maxChars without losing long words
function wrapText_(txt, maxChars) {
  const words = String(txt).split(/\s+/);
  const lines = [];
  let line = '';

  words.forEach(w => {
    if (line.length === 0) {
      if (w.length <= maxChars) {
        line = w;
      } else {
        for (let i = 0; i < w.length; i += maxChars) {
          const chunk = w.slice(i, i + maxChars);
          if (i === 0) line = chunk; else lines.push(chunk);
        }
      }
    } else if ((line + ' ' + w).length <= maxChars) {
      line += ' ' + w;
    } else {
      lines.push(line);
      if (w.length <= maxChars) {
        line = w;
      } else {
        let first = true;
        for (let i = 0; i < w.length; i += maxChars) {
          const chunk = w.slice(i, i + maxChars);
          if (first) { line = chunk; first = false; }
          else lines.push(chunk);
        }
      }
    }
  });

  if (line) lines.push(line);
  return lines;
}

function svgToDataUri_(svg) {
  const bytes = Utilities.newBlob(svg, 'image/svg+xml').getBytes();
  const b64 = Utilities.base64Encode(bytes);
  return 'data:image/svg+xml;base64,' + b64;
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
function escapeXmlText_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ----------------- Pagination action ----------------- */

function handlePage(e) {
  const o = e.parameters && e.parameters.o || '0';
  const r = e.parameters && e.parameters.r === 'true';
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildListCard_(null, r, o)))
    .build();
}
