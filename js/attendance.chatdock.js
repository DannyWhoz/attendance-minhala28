(() => {
  'use strict';
  if (!document.body) return;
  const page = String(document.body.dataset.page || '').trim();
  if (!['attendance', 'admin'].includes(page)) return;
  const App = window.AttendanceApp;
  if (!App) return;

  const POLL_MS = 15000;
  const MAX_ROOMS = 6;
  const MAX_CONTACTS = 8;
  const state = {
    settings: App.createDefaultSettings(),
    selection: App.loadSelection(),
    currentUser: { name: '', email: '', login: '' },
    currentRoom: null,
    customRooms: [],
    rooms: [],
    seedMessages: [],
    contacts: [],
    activeRoomKey: '',
    messages: [],
    open: false,
    unreadCount: 0,
    sending: false,
    refreshToken: 0,
    refreshTimer: 0,
    lastSeenByRoom: new Map()
  };

  const byId = (id) => document.getElementById(id);
  const norm = (value) => String(value || '').trim().toLowerCase();
  const currentId = () => App.getPrimaryChatIdentity(state.currentUser);
  const msgUserId = (name, email, login) => App.getPrimaryChatIdentity({ name, email, login });
  const fmtTime = (value) => new Intl.DateTimeFormat('he-IL', { hour: '2-digit', minute: '2-digit' }).format(value instanceof Date ? value : new Date(value || Date.now()));
  const fmtDay = (date) => {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return 'ללא תאריך';
    const today = new Date();
    if (date.getFullYear() === today.getFullYear() && date.getMonth() === today.getMonth() && date.getDate() === today.getDate()) return 'היום';
    return new Intl.DateTimeFormat('he-IL', { day: 'numeric', month: 'long' }).format(date);
  };
  function refreshIcons(node) {
    try {
      if (window.FontAwesome && window.FontAwesome.dom && typeof window.FontAwesome.dom.i2svg === 'function') {
        window.FontAwesome.dom.i2svg({ node: node || document.body });
      }
    } catch (error) {}
  }

  function groupColumns() {
    return Array.isArray(state.settings.groupColumns) && state.settings.groupColumns.length ? state.settings.groupColumns.slice() : ['A'];
  }

  function selectedGroups(selection) {
    const raw = Array.isArray(selection && selection.groupSelections) && selection.groupSelections.length
      ? selection.groupSelections
      : (selection && selection.groupValues && Object.keys(selection.groupValues).length ? [selection.groupValues] : []);
    return raw.map((values) => App.normalizeGroupSelectionValues(values, groupColumns())).filter((values) => Object.keys(values).length);
  }

  function groupLabel(values) {
    return groupColumns().map((letter) => String(values && values[letter] ? values[letter] : '').trim()).filter(Boolean).join(' | ');
  }

  function buildCurrentRoom() {
    return App.buildSelectionChatScope(state.selection, state.settings);
  }

  function roomMetaFromScope(scopeKey, groupKey) {
    if (String(groupKey || '').trim().startsWith('custom:')) return 'חדר צוות';
    if (String(scopeKey || '').trim() === 'room-admin-all') return 'ערוץ מערכת';
    if (String(scopeKey || '').trim() === 'room-general') return 'ערוץ כללי';
    if (/^room-category-/i.test(String(scopeKey || '').trim())) return 'ערוץ קטגוריה';
    if (/^room-multi-/i.test(String(scopeKey || '').trim())) return 'צ׳אט קבוצות משותף';
    return 'צ׳אט קבוצה';
  }

  function roomNoticeFromScope(scopeKey, groupKey) {
    if (String(groupKey || '').trim().startsWith('custom:')) return 'חדר צ׳אט קבוע שיצר מנהל המערכת.';
    if (String(scopeKey || '').trim() === 'room-admin-all') return 'ערוץ משותף לכלל היחידות במערכת.';
    if (String(scopeKey || '').trim() === 'room-general') return 'ערוץ כללי לכלל המשתמשים.';
    if (/^room-category-/i.test(String(scopeKey || '').trim())) return 'ערוץ לפי קטגוריית קבוצות.';
    if (/^room-multi-/i.test(String(scopeKey || '').trim())) return 'שיחה משותפת לכמה קבוצות.';
    return 'שיחת קבוצה שנשמרה ב-SharePoint.';
  }

  function normalizeMessage(row) {
    const createdAt = row && row.Created ? new Date(row.Created) : null;
    return {
      id: Number(row && row.Id || 0),
      scopeKey: String(row && row.ScopeKey || '').trim(),
      scopeLabel: String(row && row.ScopeLabel || row && row.Title || '').trim(),
      conversationType: String(row && row.ConversationType || '').trim().toLowerCase() === 'direct' ? 'direct' : 'room',
      dateKey: App.normalizeDateKey(row && row.DateKey || ''),
      groupKey: String(row && row.GroupKey || '').trim(),
      participantA: norm(row && row.ParticipantA || ''),
      participantB: norm(row && row.ParticipantB || ''),
      participantLabel: String(row && row.ParticipantLabel || '').trim(),
      messageText: String(row && row.MessageText || '').trim(),
      authorName: String(row && row.AuthorName || '').trim(),
      authorEmail: String(row && row.AuthorEmail || '').trim(),
      authorLogin: String(row && row.AuthorLogin || '').trim(),
      recipientName: String(row && row.RecipientName || '').trim(),
      recipientEmail: String(row && row.RecipientEmail || '').trim(),
      recipientLogin: String(row && row.RecipientLogin || '').trim(),
      createdAt: createdAt instanceof Date && !Number.isNaN(createdAt.getTime()) ? createdAt : null
    };
  }

  function isOwn(message) {
    const authorId = msgUserId(message.authorName, message.authorEmail, message.authorLogin);
    const selfId = currentId();
    if (authorId && selfId) return authorId === selfId;
    return [state.currentUser.email, state.currentUser.login, state.currentUser.name].map(norm).filter(Boolean).includes(norm(message.authorEmail))
      || [state.currentUser.email, state.currentUser.login, state.currentUser.name].map(norm).filter(Boolean).includes(norm(message.authorLogin))
      || [state.currentUser.email, state.currentUser.login, state.currentUser.name].map(norm).filter(Boolean).includes(norm(message.authorName));
  }

  function latestId(list) {
    return Array.isArray(list) && list.length ? Number(list[list.length - 1].id || 0) : 0;
  }

  function peerFromDirect(message) {
    const selfId = currentId();
    const author = { name: message.authorName, email: message.authorEmail, login: message.authorLogin };
    const recipient = { name: message.recipientName, email: message.recipientEmail, login: message.recipientLogin };
    const authorId = msgUserId(author.name, author.email, author.login);
    const recipientId = msgUserId(recipient.name, recipient.email, recipient.login);
    if (selfId && authorId === selfId && recipientId) return { ...recipient, identity: recipientId };
    if (selfId && recipientId === selfId && authorId) return { ...author, identity: authorId };
    if (recipientId) return { ...recipient, identity: recipientId };
    if (authorId) return { ...author, identity: authorId };
    return null;
  }

  function buildDirectRoom(peer) {
    if (!peer || !peer.identity) return null;
    const room = App.buildDirectChatScope(state.currentUser, peer);
    if (!room) return null;
    return {
      ...room,
      isDirect: true,
      isContext: false,
      conversationType: 'direct',
      peerIdentity: peer.identity,
      recipientName: room.recipientName || App.getUserDisplayName(peer) || '',
      recipientEmail: room.recipientEmail || String(peer.email || '').trim(),
      recipientLogin: room.recipientLogin || String(peer.login || '').trim(),
      lastMessageId: 0,
      lastMessageOwn: false,
      lastMessageText: '',
      lastCreatedAt: null
    };
  }

  function roomFromMessage(message) {
    if (message.conversationType === 'direct') {
      const peer = peerFromDirect(message);
      const room = buildDirectRoom(peer);
      if (!room) return null;
      return { ...room, scopeKey: message.scopeKey || room.scopeKey, scopeLabel: message.scopeLabel || room.scopeLabel, participantA: message.participantA || room.participantA, participantB: message.participantB || room.participantB, participantLabel: message.participantLabel || room.participantLabel || '' };
    }
    if (!App.isManagedChatRoomScope(message.scopeKey, state.settings)) return null;
    const customRoom = state.customRooms.find((room) => room.scopeKey === message.scopeKey);
    if (customRoom) {
      return {
        ...customRoom,
        scopeLabel: message.scopeLabel || customRoom.scopeLabel,
        roomTitle: customRoom.roomTitle || message.scopeLabel || customRoom.scopeKey
      };
    }
    return {
      scopeKey: message.scopeKey,
      scopeLabel: message.scopeLabel || message.scopeKey,
      roomTitle: message.scopeLabel || message.scopeKey,
      roomMeta: roomMetaFromScope(message.scopeKey, message.groupKey),
      notice: roomNoticeFromScope(message.scopeKey, message.groupKey),
      dateKey: '',
      groupKey: message.groupKey,
      isDirect: false,
      isContext: false,
      isCustom: false,
      conversationType: 'room',
      lastMessageId: 0,
      lastMessageOwn: false,
      lastMessageText: '',
      lastCreatedAt: null
    };
  }

  function buildRooms(messages) {
    const map = new Map();
    if (state.currentRoom) map.set(state.currentRoom.scopeKey, { ...state.currentRoom });
    state.customRooms.forEach((room) => {
      if (!room || !room.scopeKey) return;
      map.set(room.scopeKey, { ...room });
    });
    (Array.isArray(messages) ? messages : []).forEach((message) => {
      if (!message || !message.scopeKey) return;
      const room = roomFromMessage(message);
      if (!room) return;
      const prev = map.get(room.scopeKey) || room;
      map.set(room.scopeKey, {
        ...prev,
        ...room,
        roomTitle: prev.isContext && !room.isDirect ? prev.roomTitle : (room.roomTitle || prev.roomTitle || room.scopeKey),
        roomMeta: prev.isContext && !room.isDirect ? prev.roomMeta : (room.roomMeta || prev.roomMeta || ''),
        notice: prev.isContext && !room.isDirect ? prev.notice : (room.notice || prev.notice || ''),
        dateKey: room.dateKey || prev.dateKey || '',
        groupKey: room.groupKey || prev.groupKey || '',
        lastMessageId: Number(message.id || prev.lastMessageId || 0),
        lastMessageOwn: isOwn(message),
        lastMessageText: message.messageText || prev.lastMessageText || '',
        lastCreatedAt: message.createdAt || prev.lastCreatedAt || null
      });
    });
    return Array.from(map.values()).sort((a, b) => {
      if (a.scopeKey === state.activeRoomKey) return -1;
      if (b.scopeKey === state.activeRoomKey) return 1;
      if (a.isContext && !b.isContext) return -1;
      if (b.isContext && !a.isContext) return 1;
      if (a.isCustom && !b.isCustom) return -1;
      if (b.isCustom && !a.isCustom) return 1;
      const at = a.lastCreatedAt instanceof Date ? a.lastCreatedAt.getTime() : 0;
      const bt = b.lastCreatedAt instanceof Date ? b.lastCreatedAt.getTime() : 0;
      if (at !== bt) return bt - at;
      return String(a.roomTitle || '').localeCompare(String(b.roomTitle || ''), 'he', { numeric: true, sensitivity: 'base' });
    });
  }

  function collectContacts(messages) {
    const map = new Map();
    const selfId = currentId();
    function put(user, meta) {
      const identity = App.getPrimaryChatIdentity(user);
      if (!identity || identity === selfId) return;
      const prev = map.get(identity) || { identity, name: App.getUserDisplayName(user) || 'משתמש', email: String(user.email || '').trim(), login: String(user.login || '').trim(), hint: 'משתמש פעיל בצ׳אט', lastCreatedAt: null };
      map.set(identity, { ...prev, name: App.getUserDisplayName(user) || prev.name, hint: meta && meta.hint ? meta.hint : prev.hint, lastCreatedAt: meta && meta.createdAt ? meta.createdAt : prev.lastCreatedAt });
    }
    (Array.isArray(messages) ? messages : []).forEach((message) => {
      if (!message || !message.scopeKey) return;
      if (message.conversationType === 'direct') {
        const peer = peerFromDirect(message);
        if (peer) put(peer, { createdAt: message.createdAt, hint: 'שיחה ישירה' });
      } else {
        put({ name: message.authorName, email: message.authorEmail, login: message.authorLogin }, { createdAt: message.createdAt, hint: message.scopeLabel || 'השתתף בחדר צוות' });
      }
    });
    return Array.from(map.values()).sort((a, b) => {
      const at = a.lastCreatedAt instanceof Date ? a.lastCreatedAt.getTime() : 0;
      const bt = b.lastCreatedAt instanceof Date ? b.lastCreatedAt.getTime() : 0;
      if (at !== bt) return bt - at;
      return String(a.name || '').localeCompare(String(b.name || ''), 'he', { numeric: true, sensitivity: 'base' });
    }).slice(0, MAX_CONTACTS);
  }

  function activeRoom() {
    return state.rooms.find((room) => room.scopeKey === state.activeRoomKey) || state.currentRoom;
  }

  function ensureRoom(room) {
    if (!room) return null;
    const existing = state.rooms.find((item) => item.scopeKey === room.scopeKey);
    if (existing) return existing;
    state.rooms = [room, ...state.rooms.filter((item) => item.scopeKey !== room.scopeKey)].slice(0, 18);
    return room;
  }

  function syncUnread() {
    if (state.open) {
      state.unreadCount = 0;
      return;
    }
    let count = 0;
    state.rooms.forEach((room) => {
      const seen = Number(state.lastSeenByRoom.get(room.scopeKey) || 0);
      if (Number(room.lastMessageId || 0) > seen && !room.lastMessageOwn) count += 1;
    });
    state.unreadCount = Math.min(count, 99);
  }

  function setOpen(flag) {
    state.open = !!flag;
    const root = byId('chatDockRoot');
    const launcher = byId('chatDockLauncher');
    if (root) root.dataset.open = state.open ? '1' : '0';
    if (launcher) launcher.setAttribute('aria-expanded', String(state.open));
    if (state.open) {
      markSeen();
      renderLauncher();
      window.setTimeout(() => scrollMessages(true), 0);
    } else {
      syncUnread();
      renderLauncher();
    }
  }

  function setStatus(message, kind) {
    const host = byId('chatDockStatus');
    if (!host) return;
    if (!message) {
      host.className = 'chat-dock-status hidden';
      host.textContent = '';
      return;
    }
    host.className = `chat-dock-status${kind ? ` ${kind}` : ''}`;
    host.textContent = message;
  }

  function renderLauncher() {
    const room = activeRoom();
    const title = byId('chatDockLauncherTitle');
    const meta = byId('chatDockLauncherMeta');
    const badge = byId('chatDockLauncherBadge');
    if (title) title.textContent = room ? (room.roomTitle || room.scopeLabel || 'צ׳אט') : 'צ׳אט';
    if (meta) meta.textContent = room ? (room.isDirect ? 'שיחה ישירה' : (room.roomMeta || 'שיחה פעילה')) : 'שיחה פעילה';
    if (badge) {
      badge.classList.toggle('hidden', !state.unreadCount);
      badge.textContent = state.unreadCount ? String(state.unreadCount) : '';
    }
  }

  function renderHeader() {
    const room = activeRoom();
    const me = App.getUserDisplayName(state.currentUser);
    if (byId('chatDockTitle')) byId('chatDockTitle').textContent = room ? (room.roomTitle || room.scopeLabel || 'חדר') : 'חדר';
    if (byId('chatDockMeta')) byId('chatDockMeta').textContent = room ? (room.isDirect ? `שיחה ישירה${room.recipientName ? ` עם ${room.recipientName}` : ''}` : (room.roomMeta || '')) : '';
    if (byId('chatDockComposerMeta')) byId('chatDockComposerMeta').textContent = me ? `${room && room.isDirect ? 'הודעה פרטית' : 'כותבים'} בתור ${me}` : 'הודעות נשמרות ב-SharePoint';
  }

  function renderRooms() {
    const host = byId('chatDockRoomStrip');
    if (!host) return;
    host.innerHTML = state.rooms.slice(0, MAX_ROOMS).map((room) => {
      const active = room.scopeKey === state.activeRoomKey ? ' active' : '';
      const direct = room.isDirect ? ' is-direct' : '';
      const meta = room.lastCreatedAt ? fmtTime(room.lastCreatedAt) : (room.isDirect ? 'ישיר' : (room.roomMeta || ''));
      return `<button type="button" class="chat-dock-roomchip${active}${direct}" data-room-key="${App.escapeHtml(room.scopeKey)}"><strong>${App.escapeHtml(room.roomTitle || room.scopeLabel || 'חדר')}</strong><small>${App.escapeHtml(meta)}</small></button>`;
    }).join('');
  }

  function renderContacts() {
    const host = byId('chatDockContacts');
    const label = byId('chatDockContactsLabel');
    if (!host || !label) return;
    const contacts = state.contacts.slice(0, MAX_CONTACTS);
    label.classList.toggle('hidden', !contacts.length);
    host.classList.toggle('hidden', !contacts.length);
    host.innerHTML = contacts.map((contact) => `<button type="button" class="chat-dock-contact" data-contact-key="${App.escapeHtml(contact.identity)}" title="${App.escapeHtml(contact.name)}"><span class="chat-dock-contact-avatar">${App.escapeHtml(App.getUserAvatarText(contact, 'DM'))}</span><span class="chat-dock-contact-copy"><strong>${App.escapeHtml(contact.name || 'משתמש')}</strong><small>${App.escapeHtml(contact.hint || 'שיחה ישירה')}</small></span></button>`).join('');
  }

  function scrollMessages(force) {
    const host = byId('chatDockMessages');
    if (!host) return;
    const distance = host.scrollHeight - host.clientHeight - host.scrollTop;
    if (force || distance < 120) host.scrollTop = host.scrollHeight;
  }

  function renderMessages(forceScroll) {
    const host = byId('chatDockMessages');
    const room = activeRoom();
    if (!host) return;
    if (!room) {
      host.innerHTML = '<div class="chat-dock-empty"><strong>אין שיחה פעילה</strong><p>בחרו קבוצה, פתחו חדר קבוע או התחילו שיחה ישירה.</p></div>';
      return;
    }
    if (!state.messages.length) {
      const line = room.isDirect ? `עדיין אין הודעות עם ${room.roomTitle || 'המשתמש הזה'}.` : `עדיין אין הודעות ב-${room.roomTitle || room.scopeLabel || 'החדר הזה'}.`;
      host.innerHTML = `<div class="chat-dock-empty"><strong>${room.isDirect ? 'השיחה הפרטית מוכנה' : 'החדר מוכן להתחלה'}</strong><p>${App.escapeHtml(line)}</p></div>`;
      return;
    }
    let day = '';
    host.innerHTML = state.messages.map((message) => {
      const own = isOwn(message);
      const key = message.createdAt ? `${message.createdAt.getFullYear()}-${message.createdAt.getMonth()}-${message.createdAt.getDate()}` : 'unknown';
      const divider = key !== day ? `<div class="chat-dock-day"><span>${App.escapeHtml(fmtDay(message.createdAt))}</span></div>` : '';
      day = key;
      return `${divider}<div class="chat-dock-message-row ${own ? 'sent' : 'received'}"><article class="chat-dock-bubble"><div class="chat-dock-bubble-body">${App.escapeHtml(message.messageText)}</div><div class="chat-dock-bubble-meta"><span>${App.escapeHtml(own ? 'אתם' : (message.authorName || 'משתמש'))}</span><span>${App.escapeHtml(fmtTime(message.createdAt || new Date()))}</span></div></article></div>`;
    }).join('');
    scrollMessages(forceScroll);
  }

  function updateComposer() {
    const input = byId('chatDockInput');
    const button = byId('chatDockSendBtn');
    if (!(input instanceof HTMLTextAreaElement) || !(button instanceof HTMLButtonElement)) return;
    button.disabled = state.sending || !String(input.value || '').trim() || !state.activeRoomKey;
  }

  function markSeen() {
    const room = activeRoom();
    if (!room) return;
    state.lastSeenByRoom.set(room.scopeKey, latestId(state.messages) || Number(room.lastMessageId || 0));
    syncUnread();
  }

  async function hydrateUser() {
    try {
      state.currentUser = await App.getCurrentUser();
    } catch (error) {
      state.currentUser = { name: '', email: '', login: '' };
    }
    renderHeader();
  }

  async function loadRooms(quiet) {
    const token = ++state.refreshToken;
    if (!quiet) setStatus('טוען חדרים...', '');
    const settingsResult = await App.loadManagementSettings({ preferCache: true });
    if (token !== state.refreshToken) return;
    state.settings = settingsResult.settings;
    state.selection = App.loadSelection();
    state.currentRoom = buildCurrentRoom();
    state.customRooms = App.resolveCustomChatRooms(state.settings).map((room) => App.buildCustomChatRoomScope(room)).filter(Boolean);
    const rows = await App.loadRecentChatEntries(state.settings, { top: 220 });
    if (token !== state.refreshToken) return;
    const roomRows = (Array.isArray(rows) ? rows : []).filter((row) => {
      if (String(row && row.ConversationType || '').trim().toLowerCase() === 'direct') return false;
      return App.isManagedChatRoomScope(row && row.ScopeKey || '', state.settings);
    });
    const directRows = currentId() ? await App.loadRecentChatEntries(state.settings, { top: 160, conversationType: 'direct', participantIdentity: currentId() }) : [];
    if (token !== state.refreshToken) return;
    state.seedMessages = roomRows.concat(Array.isArray(directRows) ? directRows : []).map(normalizeMessage);
    state.rooms = buildRooms(state.seedMessages);
    state.contacts = collectContacts(state.seedMessages.concat(state.messages));
    if (!state.activeRoomKey || !state.rooms.some((room) => room.scopeKey === state.activeRoomKey)) {
      state.activeRoomKey = state.currentRoom && state.currentRoom.scopeKey
        ? state.currentRoom.scopeKey
        : (state.rooms[0] && state.rooms[0].scopeKey || '');
    }
    renderRooms();
    renderContacts();
    renderHeader();
    syncUnread();
    renderLauncher();
  }

  async function loadMessages(quiet, forceScroll) {
    const room = activeRoom();
    if (!room) {
      state.messages = [];
      renderMessages(forceScroll);
      return;
    }
    if (!quiet) setStatus('טוען הודעות...', '');
    const rows = await App.loadChatMessages(state.settings, room.scopeKey, { top: 250 });
    state.messages = rows.map(normalizeMessage);
    const last = state.messages.length ? state.messages[state.messages.length - 1] : null;
    const roomRef = state.rooms.find((item) => item.scopeKey === room.scopeKey);
    if (roomRef && last) {
      roomRef.lastMessageId = latestId(state.messages);
      roomRef.lastMessageOwn = isOwn(last);
      roomRef.lastMessageText = last.messageText;
      roomRef.lastCreatedAt = last.createdAt;
    }
    if (state.open) state.lastSeenByRoom.set(room.scopeKey, latestId(state.messages));
    state.contacts = collectContacts(state.seedMessages.concat(state.messages));
    renderMessages(forceScroll);
    renderContacts();
    renderRooms();
    renderHeader();
    syncUnread();
    renderLauncher();
    setStatus('הצ׳אט מחובר ל-SharePoint.', 'ok');
  }

  async function refresh(options) {
    const quiet = !!(options && options.quiet);
    const forceScroll = !!(options && options.forceScroll);
    try {
      await loadRooms(quiet);
      await loadMessages(quiet, forceScroll);
    } catch (error) {
      setStatus(error && error.message ? String(error.message) : 'שגיאה בטעינת הצ׳אט.', 'err');
    } finally {
      updateComposer();
    }
  }

  function scheduleRefresh() {
    if (state.refreshTimer) window.clearTimeout(state.refreshTimer);
    state.refreshTimer = window.setTimeout(async () => {
      if (!document.hidden) await refresh({ quiet: true, forceScroll: false });
      scheduleRefresh();
    }, POLL_MS);
  }

  function activateRoom(room) {
    if (!room) return;
    const next = ensureRoom(room) || room;
    state.activeRoomKey = next.scopeKey;
    state.messages = [];
    renderRooms();
    renderContacts();
    renderHeader();
    renderMessages(false);
    renderLauncher();
  }

  function openDirect(contact) {
    const room = buildDirectRoom(contact);
    if (!room) return;
    activateRoom(room);
    setOpen(true);
    loadMessages(false, true).catch((error) => setStatus(error && error.message ? String(error.message) : 'שגיאה בטעינת הודעות.', 'err'));
  }

  async function sendMessage(event) {
    event.preventDefault();
    const input = byId('chatDockInput');
    if (!(input instanceof HTMLTextAreaElement)) return;
    const messageText = String(input.value || '').trim();
    if (!messageText || state.sending || !state.activeRoomKey) return;
    state.sending = true;
    updateComposer();
    try {
      if (!App.getUserDisplayName(state.currentUser)) await hydrateUser();
      const room = activeRoom();
      const payload = { scopeKey: room.scopeKey, scopeLabel: room.scopeLabel, dateKey: room.dateKey, groupKey: room.groupKey, messageText, authorName: App.getUserDisplayName(state.currentUser), authorEmail: state.currentUser.email, authorLogin: state.currentUser.login };
      if (room && room.isDirect) {
        Object.assign(payload, { conversationType: 'direct', participantA: room.participantA, participantB: room.participantB, participantLabel: room.participantLabel, recipientName: room.recipientName || room.roomTitle, recipientEmail: room.recipientEmail || '', recipientLogin: room.recipientLogin || '' });
      }
      await App.createChatMessage(state.settings, payload);
      input.value = '';
      await refresh({ quiet: true, forceScroll: true });
      setOpen(true);
      setStatus('ההודעה נשלחה.', 'ok');
    } catch (error) {
      setStatus(error && error.message ? String(error.message) : 'לא ניתן לשלוח הודעה.', 'err');
    } finally {
      state.sending = false;
      updateComposer();
    }
  }

  function createDock() {
    if (byId('chatDockRoot')) return;
    document.body.insertAdjacentHTML('beforeend', `
      <div id="chatDockRoot" class="chat-dock-root" data-open="0">
        <section class="chat-dock-window" aria-label="צ׳אט">
          <header class="chat-dock-header">
            <div class="chat-dock-header-copy">
              <span class="chat-dock-kicker">צ׳אט קבוצות</span>
              <strong id="chatDockTitle" class="chat-dock-title">טוען חדר</strong>
              <span id="chatDockMeta" class="chat-dock-meta">טוען שיחה פעילה.</span>
            </div>
            <div class="chat-dock-header-actions">
              <button id="chatDockRefreshBtn" type="button" class="chat-dock-icon-btn" aria-label="רענון צ׳אט" title="רענון"><i class="fas fa-arrows-rotate" aria-hidden="true"></i></button>
              <button id="chatDockCloseBtn" type="button" class="chat-dock-icon-btn" aria-label="מזער צ׳אט" title="מזער"><i class="fas fa-chevron-down" aria-hidden="true"></i></button>
            </div>
          </header>
          <div id="chatDockRoomStrip" class="chat-dock-roomstrip"></div>
          <div id="chatDockContactsLabel" class="chat-dock-inline-label hidden">ישיר עם</div>
          <div id="chatDockContacts" class="chat-dock-contactstrip hidden"></div>
          <div id="chatDockStatus" class="chat-dock-status hidden"></div>
          <div id="chatDockMessages" class="chat-dock-messages"></div>
          <form id="chatDockComposer" class="chat-dock-composer">
            <textarea id="chatDockInput" class="chat-dock-input" rows="3" placeholder="כתבו הודעה קצרה..."></textarea>
            <div class="chat-dock-composer-row">
              <span id="chatDockComposerMeta" class="chat-dock-composer-meta">הודעות נשמרות ב-SharePoint</span>
              <button id="chatDockSendBtn" type="submit" class="chat-dock-send">שלח</button>
            </div>
          </form>
        </section>
        <button id="chatDockLauncher" type="button" class="chat-dock-launcher" aria-expanded="false" aria-controls="chatDockRoot">
          <span class="chat-dock-launcher-icon"><i class="fas fa-comments" aria-hidden="true"></i></span>
          <span class="chat-dock-launcher-copy"><strong id="chatDockLauncherTitle">צ׳אט</strong><small id="chatDockLauncherMeta">טוען חדר...</small></span>
          <span id="chatDockLauncherBadge" class="chat-dock-launcher-badge hidden"></span>
        </button>
      </div>
    `);
    refreshIcons(byId('chatDockRoot'));
  }

  function bindEvents() {
    const launcher = byId('chatDockLauncher');
    const closeButton = byId('chatDockCloseBtn');
    const refreshButton = byId('chatDockRefreshBtn');
    const roomStrip = byId('chatDockRoomStrip');
    const contacts = byId('chatDockContacts');
    const composer = byId('chatDockComposer');
    const input = byId('chatDockInput');
    if (launcher) launcher.addEventListener('click', () => setOpen(!state.open));
    if (closeButton) closeButton.addEventListener('click', () => setOpen(false));
    if (refreshButton) refreshButton.addEventListener('click', () => { refresh({ quiet: false, forceScroll: false }).catch(() => {}); });
    if (roomStrip) roomStrip.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) return;
      const button = target.closest('button[data-room-key]');
      if (!(button instanceof HTMLButtonElement)) return;
      const nextKey = String(button.dataset.roomKey || '').trim();
      if (!nextKey || nextKey === state.activeRoomKey) return;
      state.activeRoomKey = nextKey;
      renderRooms();
      renderHeader();
      setOpen(true);
      loadMessages(false, true).catch((error) => setStatus(error && error.message ? String(error.message) : 'שגיאה בטעינת ההודעות.', 'err'));
    });
    if (contacts) contacts.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) return;
      const button = target.closest('button[data-contact-key]');
      if (!(button instanceof HTMLButtonElement)) return;
      const identity = String(button.dataset.contactKey || '').trim();
      const contact = state.contacts.find((item) => item.identity === identity);
      if (contact) openDirect(contact);
    });
    if (composer) composer.addEventListener('submit', sendMessage);
    if (input) input.addEventListener('input', updateComposer);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) refresh({ quiet: true, forceScroll: false }).catch(() => {}); });
    document.addEventListener('keydown', (event) => { if (event instanceof KeyboardEvent && event.key === 'Escape' && state.open) setOpen(false); });
  }

  async function init() {
    createDock();
    bindEvents();
    renderLauncher();
    updateComposer();
    await hydrateUser();
    await refresh({ quiet: false, forceScroll: false });
    scheduleRefresh();
  }

  init().catch((error) => {
    setStatus(error && error.message ? String(error.message) : 'שגיאה בטעינת הצ׳אט.', 'err');
  });
})();
