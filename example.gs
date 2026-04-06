const APP_TITLE = 'D-Messenger';
const PAGE_SIZE = 40;
const MAX_PAGE_SIZE = 100;
const ATTACHMENT_MAX_FILE_BYTES = 8 * 1024 * 1024;   // 8 MB
const ATTACHMENT_MAX_TOTAL_BYTES = 20 * 1024 * 1024; // 20 MB
const ROOT_FOLDER_PROP = 'CHAT_ATTACHMENTS_ROOT_FOLDER_ID';
const WORKSPACE_SHEET_NAME = 'Messenger';
const SYSTEM_SHEET_NAMES = ['rooms', 'room_members', 'messages', 'attachments', 'users'];

const SCHEMA = {
  rooms: [
    'room_id',
    'title',
    'type',
    'created_at',
    'created_by',
    'last_message_at',
    'last_message_text',
    'last_sender',
    'avatar_label',
    'avatar_color'
  ],
  room_members: [
    'room_id',
    'user_email',
    'role',
    'joined_at',
    'last_read_at',
    'is_muted'
  ],
  messages: [
    'id',
    'room_id',
    'sender',
    'message_type',
    'text',
    'created_at',
    'edited_at',
    'deleted_at',
    'reply_to_id',
    'client_id',
    'attachment_group_id'
  ],
  attachments: [
    'id',
    'message_id',
    'group_id',
    'kind',
    'drive_file_id',
    'file_name',
    'mime_type',
    'size_bytes',
    'view_url',
    'download_url',
    'thumb_url',
    'created_at',
    'sort_order'
  ],
  users: [
    'email',
    'display_name',
    'display_name_set_at',
    'avatar_label',
    'avatar_color',
    'last_seen_at'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Messenger')
    .addItem('Open messenger', 'showSidebar')
    .addSeparator()
    .addItem('Initialize messenger structure', 'initializeMessengerFromMenu')
    .addItem('Create default room', 'createDefaultRoomFromMenu')
    .addItem('Set my display name', 'setMyDisplayNameFromMenu')
    .addItem('Sync people directory', 'syncPeopleDirectoryFromMenu')
    .addItem('Backfill room metadata', 'backfillRoomStatsFromMenu')
    .addItem('Repair room_members sheet', 'repairRoomMembersFromMenu')
    .addItem('Resync file access', 'resyncFileAccessFromMenu')
    .addItem('Debug my access', 'debugCurrentAccessFromMenu')
    .addToUi();
}

function showSidebar() {
  ensureChatSchema();

  const html = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setTitle(APP_TITLE);

  SpreadsheetApp.getUi().showSidebar(html);
}

function initializeMessengerFromMenu() {
  const result = initializeMessenger();
  SpreadsheetApp.getUi().alert(result.message);
}

function createDefaultRoomFromMenu() {
  const result = createDefaultRoom();
  SpreadsheetApp.getUi().alert(result.message);
}

function setMyDisplayNameFromMenu() {
  const user = getCurrentUserOrThrow_();
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Ваше имя в мессенджере', 'Введите отображаемое имя, которое будут видеть другие участники.', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const result = saveCurrentUserDisplayName(response.getResponseText());
  ui.alert(result.message);
}

function syncPeopleDirectoryFromMenu() {
  const result = syncPeopleDirectory();
  SpreadsheetApp.getUi().alert(result.message);
}

function backfillRoomStatsFromMenu() {
  const result = backfillRoomStats();
  SpreadsheetApp.getUi().alert(result.message);
}

function repairRoomMembersFromMenu() {
  const result = repairRoomMembersSheet_();
  SpreadsheetApp.getUi().alert(result.message);
}

function resyncFileAccessFromMenu() {
  const result = resyncAllAttachmentPermissionsFromMemberships_();
  SpreadsheetApp.getUi().alert(result.message);
}

function debugCurrentAccessFromMenu() {
  const result = debugCurrentAccess();
  SpreadsheetApp.getUi().alert(JSON.stringify(result, null, 2));
}

function initializeMessenger() {
  ensureChatSchema();
  ensureWorkspaceSheet_();
  hideSystemSheets_();
  upsertCurrentUser_();
  ensureRootAttachmentsFolder_();
  syncCollaboratorsIntoUsers_();
  ensureStarterRoomForCurrentUser_();
  repairRoomMembersSheet_();
  repairRoomsFromMemberships_();
  resyncAllAttachmentPermissionsFromMemberships_();

  return {
    ok: true,
    message: 'Структура мессенджера готова: листы созданы/проверены, системные листы скрыты, папка вложений создана, стартовая комната подготовлена.'
  };
}

function createDefaultRoom() {
  ensureChatSchema();

  const me = getCurrentUserOrThrow_();
  const rooms = readSheetObjects_('rooms');
  if (rooms.some(row => normalizeRoomId_(row.room_id) === 'general')) {
    return {
      ok: true,
      message: 'Комната general уже существует.'
    };
  }

  const now = new Date().toISOString();

  appendObjectRow_('rooms', {
    room_id: 'general',
    title: 'General',
    type: 'group',
    created_at: now,
    created_by: me.email,
    last_message_at: '',
    last_message_text: '',
    last_sender: '',
    avatar_label: 'GE',
    avatar_color: pickAvatarColor_('General')
  });

  appendObjectRow_('room_members', {
    room_id: 'general',
    user_email: me.email,
    role: 'owner',
    joined_at: now,
    last_read_at: now,
    is_muted: 'false'
  });

  return {
    ok: true,
    message: 'Комната general создана. Теперь можно добавить других участников в лист room_members.'
  };
}

function getClientBoot() {
  ensureChatSchema();
  upsertCurrentUser_();
  ensureCurrentUserHasGeneralAccess_();
  const currentUser = getCurrentUserOrThrow_();

  return {
    app_title: APP_TITLE,
    current_user: currentUser,
    page_size: PAGE_SIZE,
    poll_ms: 5000,
    limits: {
      max_file_bytes: ATTACHMENT_MAX_FILE_BYTES,
      max_total_bytes: ATTACHMENT_MAX_TOTAL_BYTES
    },
    features: {
      people_directory: true,
      group_admin: true,
      direct_messages: true
    },
    rooms: getRooms()
  };
}

function getAppState() {
  ensureChatSchema();
  upsertCurrentUser_();
  ensureCurrentUserHasGeneralAccess_();
  const currentUser = getCurrentUserOrThrow_();

  return {
    current_user: currentUser,
    rooms: getRooms()
  };
}

function syncPeopleDirectory() {
  ensureChatSchema();
  upsertCurrentUser_();
  const result = syncCollaboratorsIntoUsers_();
  return {
    ok: true,
    synced_count: result.synced_count,
    message: 'Каталог людей синхронизирован. Добавлено/обновлено: ' + result.synced_count
  };
}

function getPeopleDirectory() {
  ensureChatSchema();
  upsertCurrentUser_();
  ensureCurrentUserHasGeneralAccess_();
  const currentUser = getCurrentUserOrThrow_();
  const usersMap = getUsersMap_();
  const roomMembers = readSheetObjects_('room_members');
  const collaborators = getSpreadsheetCollaborators_();
  const collaboratorMap = collaborators.reduce((acc, item) => {
    const email = normalizeEmail_(item.email);
    if (email) acc[email] = item;
    return acc;
  }, {});
  const directLookup = buildDirectRoomLookupForUser_(currentUser.email);
  const directory = {};

  Object.keys(usersMap).forEach(email => {
    const profile = usersMap[email] || {};
    directory[email] = buildPersonEntry_(email, profile, collaboratorMap[email] || null, directLookup[email] || '');
  });

  roomMembers.forEach(row => {
    const email = normalizeEmail_(row.user_email);
    if (!email) return;
    if (!directory[email]) {
      directory[email] = buildPersonEntry_(email, usersMap[email] || {}, collaboratorMap[email] || null, directLookup[email] || '');
    }
  });

  Object.keys(collaboratorMap).forEach(email => {
    if (!directory[email]) {
      directory[email] = buildPersonEntry_(email, usersMap[email] || {}, collaboratorMap[email] || null, directLookup[email] || '');
    }
  });

  directory[currentUser.email] = buildPersonEntry_(currentUser.email, usersMap[currentUser.email] || {}, collaboratorMap[currentUser.email] || null, '');
  directory[currentUser.email].is_me = true;

  return Object.keys(directory)
    .map(email => directory[email])
    .sort((a, b) => {
      if (a.is_me && !b.is_me) return -1;
      if (b.is_me && !a.is_me) return 1;
      return String(a.display_name || a.email).localeCompare(String(b.display_name || b.email), 'ru');
    });
}

function getRoomDetails(roomId) {
  ensureChatSchema();
  const user = getCurrentUserOrThrow_();
  const normalizedRoomId = normalizeRoomId_(roomId);
  if (!userHasAccessToRoom_(user.email, normalizedRoomId)) {
    throw new Error('Нет доступа к этой комнате.');
  }

  const room = getRoomById_(normalizedRoomId);
  if (!room) {
    throw new Error('Комната не найдена.');
  }

  const memberRows = getRoomMemberRows_(normalizedRoomId);
  const usersMap = getUsersMap_();
  const collaborators = getSpreadsheetCollaboratorsMap_();
  const canManage = canManageRoom_(normalizedRoomId, user.email);
  const roomSummary = buildRoomSummaryForUser_(room, user.email, usersMap, memberRows);
  const members = memberRows.map(row => buildRoomMemberView_(row, usersMap, collaborators));
  const memberEmails = new Set(members.map(item => item.email));
  const candidatePeople = getPeopleDirectory()
    .filter(person => !person.is_me)
    .filter(person => !memberEmails.has(normalizeEmail_(person.email)));

  return {
    room: roomSummary,
    can_manage: canManage,
    members: members,
    candidate_people: candidatePeople
  };
}

function createGroupRoom(title, memberEmails) {
  ensureChatSchema();
  const user = getCurrentUserOrThrow_();
  const cleanTitle = String(title || '').trim().replace(/\s+/g, ' ');
  if (!cleanTitle) throw new Error('Введите название комнаты.');

  const normalizedMembers = uniqueEmails_([user.email].concat(Array.isArray(memberEmails) ? memberEmails : []));
  const roomId = buildGroupRoomId_(cleanTitle);
  const createdAt = new Date().toISOString();

  withScriptLock_(function() {
    appendObjectRow_('rooms', {
      room_id: roomId,
      title: cleanTitle,
      type: 'group',
      created_at: createdAt,
      created_by: user.email,
      last_message_at: '',
      last_message_text: '',
      last_sender: '',
      avatar_label: buildAvatarLabel_(cleanTitle),
      avatar_color: pickAvatarColor_(cleanTitle)
    });

    appendObjectRows_('room_members', normalizedMembers.map(email => ({
      room_id: roomId,
      user_email: email,
      role: email === user.email ? 'owner' : 'member',
      joined_at: createdAt,
      last_read_at: email === user.email ? createdAt : '',
      is_muted: 'false'
    })));

    normalizedMembers.forEach(email => upsertUserProfile_(email));
    SpreadsheetApp.flush();
  });

  syncRoomFolderPermissions_(roomId);

  return {
    ok: true,
    room_id: roomId
  };
}

function renameRoom(roomId, title) {
  ensureChatSchema();
  const user = getCurrentUserOrThrow_();
  const normalizedRoomId = normalizeRoomId_(roomId);
  const cleanTitle = String(title || '').trim().replace(/\s+/g, ' ');
  if (!cleanTitle) throw new Error('Введите название комнаты.');
  if (!canManageRoom_(normalizedRoomId, user.email)) throw new Error('Нет прав на управление этой комнатой.');

  const room = getRoomById_(normalizedRoomId);
  if (!room) throw new Error('Комната не найдена.');
  if (String(room.type || '').toLowerCase() === 'direct') throw new Error('Личный чат переименовать нельзя.');

  withScriptLock_(function() {
    updateRoomFields_(normalizedRoomId, {
      title: cleanTitle,
      avatar_label: buildAvatarLabel_(cleanTitle),
      avatar_color: pickAvatarColor_(cleanTitle)
    });
    SpreadsheetApp.flush();
  });

  return getRoomDetails(normalizedRoomId);
}

function addRoomMembers(roomId, memberEmails) {
  ensureChatSchema();
  const user = getCurrentUserOrThrow_();
  const normalizedRoomId = normalizeRoomId_(roomId);
  if (!canManageRoom_(normalizedRoomId, user.email)) throw new Error('Нет прав на управление этой комнатой.');

  const room = getRoomById_(normalizedRoomId);
  if (!room) throw new Error('Комната не найдена.');
  if (String(room.type || '').toLowerCase() === 'direct') throw new Error('В личный чат нельзя добавлять участников.');

  let addedEmails = [];
  withScriptLock_(function() {
    const existing = new Set(getRoomMemberEmails_(normalizedRoomId));
    const createdAt = new Date().toISOString();
    const rowsToAppend = [];
    uniqueEmails_(memberEmails || []).forEach(email => {
      if (existing.has(email)) return;
      rowsToAppend.push({
        room_id: normalizedRoomId,
        user_email: email,
        role: 'member',
        joined_at: createdAt,
        last_read_at: '',
        is_muted: 'false'
      });
      upsertUserProfile_(email);
      addedEmails.push(email);
    });

    appendObjectRows_('room_members', rowsToAppend);
    SpreadsheetApp.flush();
  });

  if (addedEmails.length) {
    syncRoomFolderPermissions_(normalizedRoomId);
    syncExistingAttachmentPermissionsForRoom_(normalizedRoomId, addedEmails);
  }

  return getRoomDetails(normalizedRoomId);
}

function removeRoomMember(roomId, memberEmail) {
  ensureChatSchema();
  const user = getCurrentUserOrThrow_();
  const normalizedRoomId = normalizeRoomId_(roomId);
  const normalizedEmail = normalizeEmail_(memberEmail);
  if (!canManageRoom_(normalizedRoomId, user.email)) throw new Error('Нет прав на управление этой комнатой.');
  if (!normalizedEmail) throw new Error('Не указан участник.');
  if (normalizedEmail === user.email) throw new Error('Себя из комнаты через этот интерфейс удалить нельзя.');

  const room = getRoomById_(normalizedRoomId);
  if (!room) throw new Error('Комната не найдена.');
  if (String(room.type || '').toLowerCase() === 'direct') throw new Error('В личном чате нельзя удалять участников.');

  withScriptLock_(function() {
    deleteRoomMembership_(normalizedRoomId, normalizedEmail);
    SpreadsheetApp.flush();
  });
  revokeExistingAttachmentPermissionsForRoom_(normalizedRoomId, [normalizedEmail]);
  revokeRoomFolderPermissions_(normalizedRoomId, [normalizedEmail]);
  return getRoomDetails(normalizedRoomId);
}

function getOrCreateDirectRoom(otherEmail) {
  ensureChatSchema();
  const user = getCurrentUserOrThrow_();
  const other = normalizeEmail_(otherEmail);
  if (!other || !looksLikeEmail_(other)) throw new Error('Некорректный email.');
  if (other === user.email) throw new Error('Нельзя создать личный чат с самим собой.');

  const roomId = withScriptLock_(function() {
    const roomId = buildDirectRoomId_(user.email, other);
    const existingRoom = getRoomById_(roomId);
    const createdAt = new Date().toISOString();

    if (!existingRoom) {
      appendObjectRow_('rooms', {
        room_id: roomId,
        title: '',
        type: 'direct',
        created_at: createdAt,
        created_by: user.email,
        last_message_at: '',
        last_message_text: '',
        last_sender: '',
        avatar_label: '',
        avatar_color: ''
      });
    }

    ensureRoomMembership_(roomId, user.email, 'owner', createdAt);
    ensureRoomMembership_(roomId, other, 'member', createdAt);
    upsertUserProfile_(other);
    SpreadsheetApp.flush();

    return roomId;
  });

  syncRoomFolderPermissions_(roomId);

  return {
    ok: true,
    room_id: roomId
  };
}

function getCurrentUserProfile() {
  ensureChatSchema();
  upsertCurrentUser_();
  return getCurrentUserOrThrow_();
}

function getInlineImagePreviews(attachmentIds) {
  ensureChatSchema();

  const user = getCurrentUserOrThrow_();
  const ids = Array.from(new Set(
    (Array.isArray(attachmentIds) ? attachmentIds : [])
      .map(id => String(id || '').trim())
      .filter(Boolean)
  )).slice(0, 12);

  if (!ids.length) return [];

  const attachmentRows = readSheetObjects_('attachments')
    .filter(row => ids.indexOf(String(row.id || '')) >= 0 && isLikelyImageFile_(row.mime_type, row.file_name));

  if (!attachmentRows.length) return [];

  const roomIdByMessageId = {};
  const wantedMessageIds = Array.from(new Set(attachmentRows.map(row => String(row.message_id || '')).filter(Boolean)));
  readSheetObjects_('messages')
    .filter(row => wantedMessageIds.indexOf(String(row.id || '')) >= 0)
    .forEach(row => {
      roomIdByMessageId[String(row.id || '')] = normalizeRoomId_(row.room_id);
    });

  return attachmentRows.reduce((result, row) => {
    const messageId = String(row.message_id || '');
    const roomId = roomIdByMessageId[messageId];
    if (!roomId || !userHasAccessToRoom_(user.email, roomId)) {
      return result;
    }

    try {
      const file = DriveApp.getFileById(String(row.drive_file_id || ''));
      const blob = file.getBlob();
      const bytes = blob.getBytes();
      if (!bytes || !bytes.length) return result;

      result.push({
        id: String(row.id || ''),
        file_name: String(row.file_name || ''),
        mime_type: String(row.mime_type || blob.getContentType() || 'image/png'),
        data_url: 'data:' + String(row.mime_type || blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(bytes),
        view_url: String(row.view_url || '')
      });
    } catch (err) {
      // skip broken previews
    }

    return result;
  }, []);
}

function saveCurrentUserDisplayName(displayName) {
  ensureChatSchema();

  const email = getCurrentUserEmailOrThrow_();
  const nextName = String(displayName || '').trim().replace(/\s+/g, ' ');

  if (nextName.length < 2) {
    throw new Error('Имя должно содержать минимум 2 символа.');
  }

  if (nextName.length > 40) {
    throw new Error('Имя должно быть не длиннее 40 символов.');
  }

  setUserDisplayName_(email, nextName);
  const currentUser = getCurrentUserOrThrow_();

  return {
    ok: true,
    message: 'Имя сохранено: ' + currentUser.name,
    current_user: currentUser
  };
}

function debugCurrentAccess() {
  ensureChatSchema();
  upsertCurrentUser_();
  repairRoomMembersSheet_();

  const email = getCurrentUserEmailOrThrow_();
  const memberships = readSheetObjects_('room_members')
    .filter(row => normalizeEmail_(row.user_email) === email)
    .map(row => ({
      room_id: normalizeRoomId_(row.room_id),
      raw_room_id: String(row.room_id || ''),
      raw_user_email: String(row.user_email || ''),
      role: String(row.role || ''),
      joined_at: String(row.joined_at || ''),
      last_read_at: String(row.last_read_at || '')
    }));

  const roomRows = readSheetObjects_('rooms').map(row => ({
    room_id: normalizeRoomId_(row.room_id),
    raw_room_id: String(row.room_id || ''),
    title: String(row.title || ''),
    type: String(row.type || '')
  }));

  const visibleRooms = getRooms().map(room => room.room_id);

  return {
    current_email: email,
    current_user: getCurrentUserOrThrow_(),
    memberships: memberships,
    room_rows: roomRows,
    visible_rooms: visibleRooms
  };
}

function ensureCurrentUserHasGeneralAccess_() {
  const email = getCurrentUserEmailOrThrow_();
  const existingMemberships = readSheetObjects_('room_members');
  const alreadyInGeneral = existingMemberships.some(row => normalizeEmail_(row.user_email) === email && normalizeRoomId_(row.room_id) === 'general');
  if (alreadyInGeneral) return false;

  return withScriptLock_(() => {
    const memberRows = readSheetObjects_('room_members');
    const stillMissing = !memberRows.some(row => normalizeEmail_(row.user_email) === email && normalizeRoomId_(row.room_id) === 'general');
    if (!stillMissing) return false;

    const now = new Date().toISOString();
    const roomRows = readSheetObjects_('rooms');
    const hasGeneral = roomRows.some(row => normalizeRoomId_(row.room_id) === 'general');

    if (!hasGeneral) {
      appendObjectRow_('rooms', {
        room_id: 'general',
        title: 'General',
        type: 'group',
        created_at: now,
        created_by: email,
        last_message_at: '',
        last_message_text: '',
        last_sender: '',
        avatar_label: 'GE',
        avatar_color: pickAvatarColor_('General')
      });
    }

    appendObjectRow_('room_members', {
      room_id: 'general',
      user_email: email,
      role: 'member',
      joined_at: now,
      last_read_at: '',
      is_muted: 'false'
    });

    return true;
  }, 4000);
}

function ensureStarterRoomForCurrentUser_() {
  const email = getCurrentUserEmailOrThrow_();
  const memberships = readSheetObjects_('room_members')
    .filter(row => normalizeEmail_(row.user_email) === email);

  if (memberships.length) return;

  const now = new Date().toISOString();

  const roomsSheet = getSheetOrThrow_('rooms');
  const roomRows = readSheetObjects_('rooms');
  const hasGeneral = roomRows.some(row => normalizeRoomId_(row.room_id) === 'general');

  if (!hasGeneral) {
    appendObjectRow_('rooms', {
      room_id: 'general',
      title: 'General',
      type: 'group',
      created_at: now,
      created_by: email,
      last_message_at: '',
      last_message_text: '',
      last_sender: '',
      avatar_label: 'GE',
      avatar_color: pickAvatarColor_('General')
    });
  }

  appendObjectRow_('room_members', {
    room_id: 'general',
    user_email: email,
    role: 'owner',
    joined_at: now,
    last_read_at: now,
    is_muted: 'false'
  });
}


function deriveGroupTitleFromRoomId_(roomId) {
  const normalized = normalizeRoomId_(roomId);
  const match = normalized.match(/^room__([^_]+(?:-[^_]+)*)__(?:[a-f0-9]{8}|[a-f0-9-]{8,})$/i);
  if (!match) return '';
  const slug = String(match[1] || '').trim();
  if (!slug) return '';
  const spaced = slug.replace(/[-_]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!spaced) return '';
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

function isPlaceholderRoomTitle_(title, roomId) {
  const cleanTitle = String(title || '').trim();
  const normalizedRoomId = normalizeRoomId_(roomId);
  if (!cleanTitle) return true;
  return normalizeRoomId_(cleanTitle) === normalizedRoomId;
}

function pickPreferredRoomRow_(currentRow, nextRow) {
  if (!currentRow) return nextRow;
  if (!nextRow) return currentRow;

  const currentRoomId = normalizeRoomId_(currentRow.room_id);
  const nextRoomId = normalizeRoomId_(nextRow.room_id);
  const currentPlaceholder = isPlaceholderRoomTitle_(currentRow.title, currentRoomId);
  const nextPlaceholder = isPlaceholderRoomTitle_(nextRow.title, nextRoomId);

  if (currentPlaceholder && !nextPlaceholder) return nextRow;
  if (!currentPlaceholder && nextPlaceholder) return currentRow;

  const currentUpdatedAt = Math.max(toDateMs_(currentRow.last_message_at), toDateMs_(currentRow.created_at));
  const nextUpdatedAt = Math.max(toDateMs_(nextRow.last_message_at), toDateMs_(nextRow.created_at));
  if (nextUpdatedAt >= currentUpdatedAt) return nextRow;
  return currentRow;
}

function repairRoomsFromMembershipsNoLock_() {
  const roomRows = readSheetObjects_('rooms');
  const memberRows = readSheetObjects_('room_members');

  const existing = new Set(roomRows.map(row => normalizeRoomId_(row.room_id)).filter(Boolean));
  const missing = [];

  memberRows.forEach(row => {
    const roomId = normalizeRoomId_(row.room_id);
    if (!roomId || existing.has(roomId)) return;
    existing.add(roomId);
    missing.push(roomId);
  });

  if (!missing.length) return;

  const now = new Date().toISOString();
  missing.forEach(roomId => {
    appendObjectRow_('rooms', {
      room_id: roomId,
      title: roomId === 'general' ? 'General' : (deriveGroupTitleFromRoomId_(roomId) || roomId),
      type: 'group',
      created_at: now,
      created_by: '',
      last_message_at: '',
      last_message_text: '',
      last_sender: '',
      avatar_label: buildAvatarLabel_(roomId),
      avatar_color: pickAvatarColor_(roomId)
    });
  });
}

function repairRoomMembersSheetNoLock_() {
  const sheet = getSheetOrThrow_('room_members');
  const headers = SCHEMA.room_members.slice();
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { ok: true, repaired_rows: 0, message: 'Лист room_members проверен.' };
  }

  const rawHeaders = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(normalizeHeader_)
    : [];
  const values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];

  const existingRoomIds = new Set(
    readSheetObjects_('rooms').map(row => normalizeRoomId_(row.room_id)).filter(Boolean)
  );

  const normalizedRows = [];
  values.forEach(row => {
    if (!row.some(cell => String(cell || '').trim() !== '')) return;

    const mapped = {};
    rawHeaders.forEach((header, index) => {
      if (header) mapped[header] = row[index];
    });

    let roomId = normalizeRoomId_(mapped.room_id);
    let userEmail = normalizeEmail_(mapped.user_email);
    let role = String(mapped.role || '').trim().toLowerCase();
    let joinedAt = toIsoString_(mapped.joined_at);
    let lastReadAt = toIsoString_(mapped.last_read_at);
    let isMuted = normalizeBooleanString_(mapped.is_muted);

    if (!roomId || !userEmail) {
      row.forEach(cell => {
        const raw = String(cell || '').trim();
        if (!raw) return;
        const emailCandidate = normalizeEmail_(raw);
        const roomCandidate = normalizeRoomId_(raw);
        if (!userEmail && looksLikeEmail_(raw)) {
          userEmail = emailCandidate;
          return;
        }
        if (!roomId && roomCandidate && existingRoomIds.has(roomCandidate)) {
          roomId = roomCandidate;
          return;
        }
        if (!role && /^(owner|member|admin)$/i.test(raw)) {
          role = raw.toLowerCase();
          return;
        }
        if (!joinedAt && looksLikeIsoDate_(raw)) {
          joinedAt = toIsoString_(raw);
          return;
        }
        if (!lastReadAt && looksLikeIsoDate_(raw) && joinedAt) {
          lastReadAt = toIsoString_(raw);
          return;
        }
        if (!isMuted && /^(true|false)$/i.test(raw)) {
          isMuted = raw.toLowerCase();
        }
      });
    }

    if (!roomId || !userEmail) return;

    normalizedRows.push({
      room_id: roomId,
      user_email: userEmail,
      role: role || 'member',
      joined_at: joinedAt || new Date().toISOString(),
      last_read_at: lastReadAt || '',
      is_muted: isMuted || 'false'
    });
  });

  const deduped = [];
  const seen = new Set();
  normalizedRows.forEach(row => {
    const key = row.room_id + '|' + row.user_email;
    if (seen.has(key)) return;
    seen.add(key);
    deduped.push(row);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (deduped.length) {
    sheet.getRange(2, 1, deduped.length, headers.length).setValues(
      deduped.map(row => headers.map(header => row[header] || ''))
    );
  }

  return {
    ok: true,
    repaired_rows: deduped.length,
    message: 'Лист room_members пересобран. Строк: ' + deduped.length
  };
}


function repairRoomsFromMemberships_() {
  return withScriptLock_(function() {
    return repairRoomsFromMembershipsNoLock_();
  });
}

function repairRoomMembersSheet_() {
  return withScriptLock_(function() {
    return repairRoomMembersSheetNoLock_();
  });
}

function getRooms() {
  ensureChatSchema();

  const user = getCurrentUserOrThrow_();
  const roomRows = readSheetObjects_('rooms');
  const memberRows = readSheetObjects_('room_members');
  const messageRows = readSheetObjects_('messages').filter(row => !row.deleted_at);
  const usersMap = getUsersMap_();

  const allowedMemberships = memberRows.filter(row => normalizeEmail_(row.user_email) === user.email);
  const allowedRoomIds = new Set(allowedMemberships.map(row => normalizeRoomId_(row.room_id)).filter(Boolean));

  const membershipByRoomId = {};
  allowedMemberships.forEach(row => {
    const roomId = normalizeRoomId_(row.room_id);
    if (roomId) membershipByRoomId[roomId] = row;
  });

  const memberRowsByRoom = groupMembersByRoomRows_(memberRows);
  const latestByRoom = {};
  const unreadCountByRoom = {};

  messageRows.forEach(row => {
    const roomId = normalizeRoomId_(row.room_id);
    if (!allowedRoomIds.has(roomId)) return;
    const createdAtMs = toDateMs_(row.created_at);
    const previousLatestMs = latestByRoom[roomId] ? toDateMs_(latestByRoom[roomId].created_at) : 0;

    if (!latestByRoom[roomId] || createdAtMs >= previousLatestMs) {
      latestByRoom[roomId] = row;
    }

    const membership = membershipByRoomId[roomId] || {};
    const lastReadAtMs = toDateMs_(membership.last_read_at);
    const senderEmail = normalizeEmail_(row.sender);

    if (createdAtMs > lastReadAtMs && senderEmail && senderEmail !== user.email) {
      unreadCountByRoom[roomId] = (unreadCountByRoom[roomId] || 0) + 1;
    }
  });

  const roomMap = {};
  roomRows.forEach(row => {
    const roomId = normalizeRoomId_(row.room_id);
    if (!roomId) return;
    roomMap[roomId] = pickPreferredRoomRow_(roomMap[roomId], row);
  });

  allowedRoomIds.forEach(roomId => {
    if (!roomMap[roomId]) {
      roomMap[roomId] = {
        room_id: roomId,
        title: roomId === 'general' ? 'General' : (deriveGroupTitleFromRoomId_(roomId) || roomId),
        type: roomId.indexOf('dm__') === 0 ? 'direct' : 'group',
        created_at: '',
        created_by: '',
        last_message_at: '',
        last_message_text: '',
        last_sender: '',
        avatar_label: '',
        avatar_color: ''
      };
    }
  });

  return Object.keys(roomMap)
    .map(roomId => roomMap[roomId])
    .filter(row => allowedRoomIds.has(normalizeRoomId_(row.room_id)))
    .map(row => {
      const roomId = normalizeRoomId_(row.room_id);
      const latest = latestByRoom[roomId] || null;
      const presentation = buildRoomSummaryForUser_(row, user.email, usersMap, memberRowsByRoom[roomId] || []);
      const lastMessageText = row.last_message_text || (latest ? latest.text : '');
      const lastMessageAt = row.last_message_at || (latest ? latest.created_at : '');
      const lastSender = row.last_sender || (latest ? latest.sender : '');

      return {
        room_id: roomId,
        title: presentation.title,
        type: presentation.type,
        peer_email: presentation.peer_email || '',
        created_at: toIsoString_(row.created_at),
        created_by: String(row.created_by || ''),
        last_message_at: toIsoString_(lastMessageAt),
        last_sender: String(lastSender || ''),
        preview_text: buildPreviewText_(lastMessageText),
        unread_count: unreadCountByRoom[roomId] || 0,
        avatar_label: presentation.avatar_label,
        avatar_color: presentation.avatar_color,
        is_muted: String((membershipByRoomId[roomId] || {}).is_muted || '').toLowerCase() === 'true',
        member_count: (memberRowsByRoom[roomId] || []).length,
        can_manage: canManageRoom_(roomId, user.email)
      };
    })
    .sort((a, b) => {
      const aMs = toDateMs_(a.last_message_at);
      const bMs = toDateMs_(b.last_message_at);
      if (aMs !== bMs) return bMs - aMs;
      return a.title.localeCompare(b.title, 'ru');
    });
}

function getMessagesPage(roomId, limit, beforeCreatedAt) {
  ensureChatSchema();

  const user = getCurrentUserOrThrow_();
  const allRows = getRoomMessageRowsForUser_(roomId, user.email);
  const beforeMs = toDateMs_(beforeCreatedAt);
  const pageSize = clampPageSize_(limit);

  const filteredRows = beforeMs
    ? allRows.filter(row => toDateMs_(row.created_at) < beforeMs)
    : allRows;

  const startIndex = Math.max(0, filteredRows.length - pageSize);
  const pageRows = filteredRows.slice(startIndex);
  const usersMap = getUsersMap_();
  const attachmentsByMessage = getAttachmentsByMessageIds_(pageRows.map(row => String(row.id || '')));

  return {
    items: pageRows.map(row => mapMessageRow_(row, usersMap, attachmentsByMessage[String(row.id || '')] || [])),
    total_count: allRows.length,
    loaded_count: pageRows.length,
    has_more: startIndex > 0,
    next_before: pageRows.length ? toIsoString_(pageRows[0].created_at) : '',
    newest_created_at: pageRows.length ? toIsoString_(pageRows[pageRows.length - 1].created_at) : ''
  };
}

function getNewMessages(roomId, afterCreatedAt) {
  ensureChatSchema();

  const user = getCurrentUserOrThrow_();
  const allRows = getRoomMessageRowsForUser_(roomId, user.email);
  const afterMs = toDateMs_(afterCreatedAt);
  const usersMap = getUsersMap_();

  const rows = allRows.filter(row => toDateMs_(row.created_at) > afterMs);
  const attachmentsByMessage = getAttachmentsByMessageIds_(rows.map(row => String(row.id || '')));

  const items = rows.map(row => mapMessageRow_(row, usersMap, attachmentsByMessage[String(row.id || '')] || []));

  return {
    items: items,
    count: items.length,
    newest_created_at: items.length ? items[items.length - 1].created_at : ''
  };
}

function sendMessagePayload(roomId, text, filesPayload, clientId) {
  ensureChatSchema();

  const user = getCurrentUserOrThrow_();
  if (!userHasAccessToRoom_(user.email, roomId)) {
    throw new Error('Нет доступа к этой комнате.');
  }

  const cleanText = String(text || '').trim();
  const files = Array.isArray(filesPayload) ? filesPayload.filter(Boolean) : [];

  if (!cleanText && !files.length) {
    throw new Error('Сообщение пустое.');
  }

  validateFilesPayload_(files);

  const createdAt = new Date().toISOString();
  const messageId = Utilities.getUuid();
  const attachmentGroupId = files.length ? Utilities.getUuid() : '';
  const room = getRoomById_(roomId);

  let attachments = [];
  if (files.length) {
    const folder = getRoomFolder_(roomId, room ? room.title : roomId);
    const roomMemberEmails = getRoomMemberEmails_(roomId);
    syncRoomFolderPermissions_(roomId, folder);
    attachments = files.map((file, index) => saveAttachment_(folder, messageId, attachmentGroupId, file, createdAt, index + 1, roomMemberEmails));
  }

  const messageType = files.length ? deriveMessageType_(attachments, cleanText) : 'text';
  const messageRow = {
    id: messageId,
    room_id: normalizeRoomId_(roomId),
    sender: user.email,
    message_type: messageType,
    text: cleanText,
    created_at: createdAt,
    edited_at: '',
    deleted_at: '',
    reply_to_id: '',
    client_id: String(clientId || ''),
    attachment_group_id: attachmentGroupId
  };

  withScriptLock_(function() {
    attachments.forEach(function(attachment) {
      appendObjectRow_('attachments', attachment);
    });
    appendObjectRow_('messages', messageRow);

    const roomPreview = buildRoomLastMessageText_(cleanText, attachments);
    updateRoomOnNewMessage_(roomId, roomPreview, user.email, createdAt);
    setRoomMemberLastRead_(roomId, user.email, createdAt);
    upsertUserProfile_(user.email);
    SpreadsheetApp.flush();
  }, 8000);

  return {
    ok: true,
    message: {
      id: messageId,
      room_id: normalizeRoomId_(roomId),
      sender: user.email,
      sender_name: user.name,
      message_type: messageType,
      text: cleanText,
      created_at: createdAt,
      edited_at: '',
      reply_to_id: '',
      client_id: String(clientId || ''),
      attachment_group_id: attachmentGroupId,
      attachments: attachments.map(mapAttachmentRow_)
    }
  };
}

/**
 * Совместимый вход для WebApp/API-роутера.
 * Поддерживает:
 * - sendMessage(roomId, text, files, clientId)
 * - sendMessage({ roomId, text, files, filesPayload, clientId })
 * - sendMessage({ room_id, text, attachments, client_id })
 */
function sendMessage(roomIdOrPayload, text, files, clientId) {
  let roomId = roomIdOrPayload;
  let messageText = text;
  let filesPayload = files;
  let optimisticClientId = clientId;

  if (roomIdOrPayload && typeof roomIdOrPayload === 'object' && !Array.isArray(roomIdOrPayload)) {
    roomId = roomIdOrPayload.roomId || roomIdOrPayload.room_id || '';
    messageText = roomIdOrPayload.text || '';
    filesPayload = roomIdOrPayload.filesPayload || roomIdOrPayload.files || roomIdOrPayload.attachments || [];
    optimisticClientId = roomIdOrPayload.clientId || roomIdOrPayload.client_id || '';
  }

  return sendMessagePayload(roomId, messageText, filesPayload, optimisticClientId);
}

function markRoomAsRead(roomId) {
  ensureChatSchema();

  const user = getCurrentUserOrThrow_();
  if (!userHasAccessToRoom_(user.email, roomId)) {
    throw new Error('Нет доступа к этой комнате.');
  }

  const timestamp = new Date().toISOString();
  setRoomMemberLastRead_(roomId, user.email, timestamp);
  upsertUserProfile_(user.email);

  return {
    ok: true,
    room_id: normalizeRoomId_(roomId),
    last_read_at: timestamp
  };
}

function backfillRoomStats() {
  ensureChatSchema();

  const roomsSheet = getSheetOrThrow_('rooms');
  const roomHeaders = getHeaders_(roomsSheet);
  const roomHeaderMap = getHeaderIndexMap_(roomHeaders);
  const roomData = roomsSheet.getLastRow() > 1
    ? roomsSheet.getRange(2, 1, roomsSheet.getLastRow() - 1, roomsSheet.getLastColumn()).getValues()
    : [];

  const messages = readSheetObjects_('messages').filter(row => !row.deleted_at);
  const attachments = readSheetObjects_('attachments');
  const attachmentCountByMessage = {};
  attachments.forEach(row => {
    const messageId = String(row.message_id || '');
    attachmentCountByMessage[messageId] = (attachmentCountByMessage[messageId] || 0) + 1;
  });

  const latestByRoom = {};

  messages.forEach(row => {
    const roomId = normalizeRoomId_(row.room_id);
    if (!roomId) return;

    const previous = latestByRoom[roomId];
    if (!previous || toDateMs_(row.created_at) >= toDateMs_(previous.created_at)) {
      latestByRoom[roomId] = row;
    }
  });

  roomData.forEach((row, index) => {
    const roomId = normalizeRoomId_(row[roomHeaderMap.room_id]);
    const latest = latestByRoom[roomId];

    row[roomHeaderMap.type] = row[roomHeaderMap.type] || 'group';
    row[roomHeaderMap.avatar_label] = row[roomHeaderMap.avatar_label] || String(row[roomHeaderMap.title] || roomId || '#').trim().slice(0, 2).toUpperCase();
    row[roomHeaderMap.avatar_color] = row[roomHeaderMap.avatar_color] || pickAvatarColor_(String(row[roomHeaderMap.title] || roomId || '#'));

    if (latest) {
      const attachmentCount = attachmentCountByMessage[String(latest.id || '')] || 0;
      const latestPreview = attachmentCount
        ? latest.text
          ? latest.text
          : attachmentCount > 1 ? '[Вложения] ' + attachmentCount : '[Вложение]'
        : latest.text;

      row[roomHeaderMap.last_message_at] = toIsoString_(latest.created_at);
      row[roomHeaderMap.last_message_text] = buildPreviewText_(latestPreview);
      row[roomHeaderMap.last_sender] = String(latest.sender || '');
    }

    roomData[index] = row;
  });

  if (roomData.length) {
    roomsSheet.getRange(2, 1, roomData.length, roomsSheet.getLastColumn()).setValues(roomData);
  }

  return {
    ok: true,
    message: 'Готово. Метаданные комнат пересчитаны: ' + roomData.length
  };
}

/* =========================
   Helpers: identity & access
   ========================= */

function getCurrentUserEmailOrThrow_() {
  const email = normalizeEmail_(Session.getActiveUser().getEmail());

  if (!email) {
    throw new Error(
      'Apps Script не вернул email текущего пользователя. ' +
      'Для этой архитектуры нужен сценарий, в котором Session.getActiveUser().getEmail() доступен. ' +
      'Проверь, что пользователь открыл именно таблицу, авторизовал скрипт и что ваш тип аккаунтов/домен это позволяет.'
    );
  }

  return email;
}

function getCurrentUserOrThrow_() {
  const email = getCurrentUserEmailOrThrow_();
  const usersMap = getUsersMap_();
  const profile = usersMap[email] || {};
  const explicitName = String(profile.display_name || '').trim();
  const effectiveName = explicitName || email.split('@')[0] || 'Unknown';

  return {
    email: email,
    name: effectiveName,
    needs_display_name: !String(profile.display_name_set_at || '').trim(),
    avatar_label: String(profile.avatar_label || buildAvatarLabel_(effectiveName || email)),
    avatar_color: String(profile.avatar_color || pickAvatarColor_(email))
  };
}


function buildPersonEntry_(email, profile, collaborator, directRoomId) {
  const normalizedEmail = normalizeEmail_(email);
  const displayName = String(profile.display_name || (collaborator && collaborator.name) || normalizedEmail.split('@')[0] || 'Unknown').trim();
  return {
    email: normalizedEmail,
    display_name: displayName,
    explicit_display_name: String(profile.display_name || '').trim(),
    avatar_label: String(profile.avatar_label || buildAvatarLabel_(displayName || normalizedEmail)),
    avatar_color: String(profile.avatar_color || pickAvatarColor_(normalizedEmail)),
    last_seen_at: toIsoString_(profile.last_seen_at),
    is_me: false,
    is_collaborator: !!collaborator,
    direct_room_id: String(directRoomId || '')
  };
}

function buildRoomSummaryForUser_(roomRow, currentUserEmail, usersMap, roomMembers) {
  const roomId = normalizeRoomId_(roomRow.room_id);
  const type = String(roomRow.type || (roomId.indexOf('dm__') === 0 ? 'direct' : 'group')).trim().toLowerCase() || 'group';
  const members = Array.isArray(roomMembers) ? roomMembers : [];
  let title = String(roomRow.title || roomId || 'Untitled').trim();
  if (type !== 'direct' && isPlaceholderRoomTitle_(title, roomId)) {
    title = deriveGroupTitleFromRoomId_(roomId) || title;
  }
  let avatarLabel = String(roomRow.avatar_label || buildAvatarLabel_(title || roomId)).trim() || '#';
  let avatarColor = String(roomRow.avatar_color || pickAvatarColor_(title || roomId)).trim() || '#5b8def';
  let peerEmail = '';

  if (type === 'direct') {
    const otherMember = members.find(item => normalizeEmail_(item.user_email) !== normalizeEmail_(currentUserEmail)) || null;
    peerEmail = normalizeEmail_((otherMember && otherMember.user_email) || '');
    const peerProfile = usersMap[peerEmail] || {};
    const peerName = String(peerProfile.display_name || (peerEmail ? peerEmail.split('@')[0] : '') || roomRow.title || 'Личный чат').trim();
    title = peerName || 'Личный чат';
    avatarLabel = String(peerProfile.avatar_label || buildAvatarLabel_(peerName || peerEmail)).trim() || 'DM';
    avatarColor = String(peerProfile.avatar_color || pickAvatarColor_(peerEmail || roomId)).trim() || '#5b8def';
  }

  return {
    room_id: roomId,
    type: type,
    title: title,
    avatar_label: avatarLabel,
    avatar_color: avatarColor,
    peer_email: peerEmail
  };
}

function groupMembersByRoomRows_(memberRows) {
  return (memberRows || []).reduce((acc, row) => {
    const roomId = normalizeRoomId_(row.room_id);
    if (!roomId) return acc;
    if (!acc[roomId]) acc[roomId] = [];
    acc[roomId].push(row);
    return acc;
  }, {});
}

function getRoomMemberRows_(roomId) {
  const normalizedRoomId = normalizeRoomId_(roomId);
  return readSheetObjects_('room_members').filter(row => normalizeRoomId_(row.room_id) === normalizedRoomId);
}

function getRoomMemberRole_(roomId, email) {
  const normalizedRoomId = normalizeRoomId_(roomId);
  const normalizedEmail = normalizeEmail_(email);
  const member = getRoomMemberRows_(normalizedRoomId).find(row => normalizeEmail_(row.user_email) === normalizedEmail) || null;
  return member ? String(member.role || 'member').trim().toLowerCase() : '';
}

function canManageRoom_(roomId, email) {
  const room = getRoomById_(roomId);
  if (!room) return false;
  if (String(room.type || '').trim().toLowerCase() === 'direct') return false;
  const role = getRoomMemberRole_(roomId, email);
  if (role === 'owner') return true;
  return normalizeEmail_(room.created_by) === normalizeEmail_(email);
}

function buildRoomMemberView_(row, usersMap, collaboratorsMap) {
  const email = normalizeEmail_(row.user_email);
  const profile = usersMap[email] || {};
  const collaborator = collaboratorsMap[email] || {};
  const displayName = String(profile.display_name || collaborator.name || email.split('@')[0] || 'Unknown').trim();
  return {
    email: email,
    display_name: displayName,
    avatar_label: String(profile.avatar_label || buildAvatarLabel_(displayName || email)),
    avatar_color: String(profile.avatar_color || pickAvatarColor_(email)),
    role: String(row.role || 'member').trim().toLowerCase() || 'member',
    joined_at: toIsoString_(row.joined_at),
    last_read_at: toIsoString_(row.last_read_at),
    is_collaborator: !!collaborator.email
  };
}

function updateRoomFields_(roomId, fields) {
  const sheet = getSheetOrThrow_('rooms');
  const headers = getHeaders_(sheet);
  const headerMap = getHeaderIndexMap_(headers);
  if (sheet.getLastRow() < 2) throw new Error('Комната не найдена.');
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const rowIndex = values.findIndex(row => normalizeRoomId_(row[headerMap.room_id]) === normalizeRoomId_(roomId));
  if (rowIndex === -1) throw new Error('Комната не найдена.');
  Object.keys(fields || {}).forEach(key => {
    if (headerMap[key] == null) return;
    values[rowIndex][headerMap[key]] = fields[key];
  });
  sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).setValues([values[rowIndex]]);
}

function deleteRoomMembership_(roomId, email) {
  const sheet = getSheetOrThrow_('room_members');
  const headers = getHeaders_(sheet);
  const headerMap = getHeaderIndexMap_(headers);
  if (sheet.getLastRow() < 2) return false;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (normalizeRoomId_(values[index][headerMap.room_id]) === normalizeRoomId_(roomId) && normalizeEmail_(values[index][headerMap.user_email]) === normalizeEmail_(email)) {
      sheet.deleteRow(index + 2);
      return true;
    }
  }
  return false;
}

function ensureRoomMembership_(roomId, email, role, timestamp) {
  const normalizedRoomId = normalizeRoomId_(roomId);
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedRoomId || !normalizedEmail) return;
  const existing = getRoomMemberRows_(normalizedRoomId).find(row => normalizeEmail_(row.user_email) === normalizedEmail);
  if (existing) return;
  appendObjectRow_('room_members', {
    room_id: normalizedRoomId,
    user_email: normalizedEmail,
    role: role || 'member',
    joined_at: timestamp || new Date().toISOString(),
    last_read_at: '',
    is_muted: 'false'
  });
}

function uniqueEmails_(list) {
  return Array.from(new Set((Array.isArray(list) ? list : []).map(item => normalizeEmail_(item)).filter(email => email && looksLikeEmail_(email))));
}

function buildGroupRoomId_(title) {
  const slug = String(title || '').trim().toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '-').replace(/^-+|-+$/g, '').slice(0, 30) || 'room';
  return 'room__' + slug + '__' + Utilities.getUuid().slice(0, 8);
}

function buildDirectRoomId_(emailA, emailB) {
  return 'dm__' + [normalizeEmail_(emailA), normalizeEmail_(emailB)].sort().map(part => part.replace(/[^a-z0-9]+/gi, '_')).join('__');
}

function buildDirectRoomLookupForUser_(currentUserEmail) {
  const email = normalizeEmail_(currentUserEmail);
  const roomRows = readSheetObjects_('rooms');
  const memberRowsByRoom = groupMembersByRoomRows_(readSheetObjects_('room_members'));
  const lookup = {};
  roomRows.forEach(room => {
    const roomId = normalizeRoomId_(room.room_id);
    if (String(room.type || '').trim().toLowerCase() !== 'direct') return;
    const members = memberRowsByRoom[roomId] || [];
    const emails = members.map(item => normalizeEmail_(item.user_email)).filter(Boolean);
    if (!emails.includes(email)) return;
    const other = emails.find(item => item !== email);
    if (other) lookup[other] = roomId;
  });
  return lookup;
}

function getSpreadsheetCollaborators_() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const map = {};
  const collect = user => {
    try {
      const email = normalizeEmail_(user && user.getEmail && user.getEmail());
      if (!email) return;
      map[email] = {
        email: email,
        name: String(user && user.getName && user.getName() || '').trim()
      };
    } catch (err) {}
  };
  try { collect(file.getOwner()); } catch (err) {}
  try { file.getEditors().forEach(collect); } catch (err) {}
  try { file.getViewers().forEach(collect); } catch (err) {}
  return Object.keys(map).map(email => map[email]);
}

function getSpreadsheetCollaboratorsMap_() {
  return getSpreadsheetCollaborators_().reduce((acc, item) => {
    const email = normalizeEmail_(item.email);
    if (email) acc[email] = item;
    return acc;
  }, {});
}

function syncCollaboratorsIntoUsers_() {
  const collaborators = getSpreadsheetCollaborators_();
  let synced = 0;
  collaborators.forEach(item => {
    const email = normalizeEmail_(item.email);
    if (!email) return;
    upsertUserProfile_(email);
    if (String(item.name || '').trim()) {
      try {
        const usersMap = getUsersMap_();
        const existing = usersMap[email] || {};
        if (!String(existing.display_name || '').trim()) {
          setUserDisplayName_(email, item.name);
        }
      } catch (err) {}
    }
    synced += 1;
  });
  return { synced_count: synced };
}

function getAllowedRoomIdsForUser_(email) {
  ensureChatSchema_();

  if (!email) return [];
  const normalizedEmail = normalizeEmail_(email);

  return readSheetObjects_('room_members')
    .filter(row => normalizeEmail_(row.user_email) === normalizedEmail)
    .map(row => normalizeRoomId_(row.room_id))
    .filter(Boolean);
}

function userHasAccessToRoom_(email, roomId) {
  const normalizedRoomId = normalizeRoomId_(roomId);
  return getAllowedRoomIdsForUser_(email).includes(normalizedRoomId);
}

function getRoomById_(roomId) {
  const normalizedRoomId = normalizeRoomId_(roomId);
  const matches = readSheetObjects_('rooms').filter(row => normalizeRoomId_(row.room_id) === normalizedRoomId);
  if (!matches.length) return null;
  return matches.reduce((best, row) => pickPreferredRoomRow_(best, row), null);
}

/* =========================
   Helpers: schema
   ========================= */

function ensureChatSchema() {
  return ensureChatSchema_();
}

function ensureChatSchema_() {
  ensureSheetHeaders_('rooms', SCHEMA.rooms, { createIfMissing: true });
  ensureSheetHeaders_('room_members', SCHEMA.room_members, { createIfMissing: true });
  ensureSheetHeaders_('messages', SCHEMA.messages, { createIfMissing: true });
  ensureSheetHeaders_('attachments', SCHEMA.attachments, { createIfMissing: true });
  ensureSheetHeaders_('users', SCHEMA.users, { createIfMissing: true });
}

function ensureSheetHeaders_(name, requiredHeaders, options) {
  const settings = options || {};
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);

  if (!sheet && settings.createIfMissing) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  }

  if (!sheet) {
    throw new Error('Лист ' + name + ' не найден.');
  }

  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(normalizeHeader_)
    : [];

  if (!headers.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return sheet;
  }

  const missingHeaders = requiredHeaders.filter(header => headers.indexOf(header) === -1);
  if (!missingHeaders.length) return sheet;

  sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  return sheet;
}

function ensureWorkspaceSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  let visibleNonSystemSheet = sheets.find(sheet => {
    return !sheet.isSheetHidden() && SYSTEM_SHEET_NAMES.indexOf(sheet.getName()) === -1;
  });

  let workspaceSheet = spreadsheet.getSheetByName(WORKSPACE_SHEET_NAME);

  if (!visibleNonSystemSheet) {
    if (!workspaceSheet) {
      workspaceSheet = spreadsheet.insertSheet(WORKSPACE_SHEET_NAME, 0);
    }

    if (workspaceSheet.isSheetHidden()) {
      workspaceSheet.showSheet();
    }

    setupWorkspaceSheet_(workspaceSheet);
    spreadsheet.setActiveSheet(workspaceSheet);
    return workspaceSheet;
  }

  if (workspaceSheet && workspaceSheet.isSheetHidden()) {
    workspaceSheet.showSheet();
  }

  if (workspaceSheet) {
    setupWorkspaceSheet_(workspaceSheet);
  }

  return visibleNonSystemSheet;
}

function setupWorkspaceSheet_(sheet) {
  if (!sheet) return;

  if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) return;

  sheet.getRange('A1').setValue(APP_TITLE);
  sheet.getRange('A2').setValue('Открой меню Messenger → Open messenger.');
  sheet.getRange('A3').setValue('Служебные листы мессенджера скрыты и используются скриптом.');
  sheet.getRange('A5').setValue('Если нужно добавить комнаты и участников, временно открой скрытые листы rooms и room_members.');
  sheet.setColumnWidth(1, 560);
}

function hideSystemSheets_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  SYSTEM_SHEET_NAMES.forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet || sheet.isSheetHidden()) return;

    const visibleSheets = spreadsheet.getSheets().filter(item => !item.isSheetHidden());
    if (visibleSheets.length <= 1) return;

    sheet.hideSheet();
  });
}

/* =========================
   Helpers: rows
   ========================= */

function readSheetObjects_(sheetName) {
  const sheet = getSheetOrThrow_(sheetName);
  const headers = getHeaders_(sheet);

  if (!headers.length || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return values
    .filter(row => row.some(cell => String(cell || '').trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => obj[header] = row[index]);
      return obj;
    });
}

function appendObjectRow_(sheetName, data) {
  const sheet = getSheetOrThrow_(sheetName);
  const headers = getHeaders_(sheet);
  const row = headers.map(header => Object.prototype.hasOwnProperty.call(data, header) ? data[header] : '');
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function appendObjectRows_(sheetName, items) {
  const rows = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!rows.length) return 0;
  const sheet = getSheetOrThrow_(sheetName);
  const headers = getHeaders_(sheet);
  const values = rows.map(data => headers.map(header => Object.prototype.hasOwnProperty.call(data, header) ? data[header] : ''));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
  return values.length;
}

function withScriptLock_(callback, timeoutMs) {
  const lock = LockService.getDocumentLock();
  const maxWaitMs = Math.max(1000, Number(timeoutMs || 8000));
  const startedAt = Date.now();

  while (!lock.tryLock(500)) {
    if (Date.now() - startedAt >= maxWaitMs) {
      throw new Error('Мессенджер занят другой операцией. Повторите через пару секунд.');
    }
    Utilities.sleep(150);
  }

  try {
    return callback();
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // ignore release issues
    }
  }
}

function getSheetOrThrow_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Лист ' + name + ' не найден.');
  return sheet;
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(normalizeHeader_);
}

function getHeaderIndexMap_(headers) {
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}

/* =========================
   Helpers: messages
   ========================= */

function getRoomMessageRowsForUser_(roomId, email) {
  if (!userHasAccessToRoom_(email, roomId)) {
    throw new Error('Нет доступа к этой комнате.');
  }

  return readSheetObjects_('messages')
    .filter(row => !row.deleted_at)
    .filter(row => normalizeRoomId_(row.room_id) === normalizeRoomId_(roomId))
    .sort((a, b) => toDateMs_(a.created_at) - toDateMs_(b.created_at));
}

function getAttachmentsByMessageIds_(messageIds) {
  const ids = new Set((messageIds || []).filter(Boolean));
  if (!ids.size) return {};

  return readSheetObjects_('attachments')
    .filter(row => ids.has(String(row.message_id || '')))
    .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0))
    .reduce((acc, row) => {
      const key = String(row.message_id || '');
      if (!acc[key]) acc[key] = [];
      acc[key].push(mapAttachmentRow_(row));
      return acc;
    }, {});
}

function mapMessageRow_(row, usersMap, attachments) {
  const senderEmail = String(row.sender || '').trim();
  const profile = usersMap[normalizeEmail_(senderEmail)] || {};

  return {
    id: String(row.id || ''),
    room_id: normalizeRoomId_(row.room_id),
    sender: senderEmail,
    sender_name: String(profile.display_name || senderEmail.split('@')[0] || 'Unknown'),
    message_type: String(row.message_type || (attachments && attachments.length ? 'mixed' : 'text')),
    text: String(row.text || ''),
    created_at: toIsoString_(row.created_at),
    edited_at: toIsoString_(row.edited_at),
    reply_to_id: String(row.reply_to_id || ''),
    client_id: String(row.client_id || ''),
    attachment_group_id: String(row.attachment_group_id || ''),
    attachments: attachments || []
  };
}

function mapAttachmentRow_(row) {
  const fileName = String(row.file_name || '');
  const mimeType = String(row.mime_type || '');
  const driveFileId = String(row.drive_file_id || '');
  const normalizedKind = isLikelyImageFile_(mimeType, fileName)
    ? 'image'
    : String(row.kind || 'file');

  return {
    id: String(row.id || ''),
    message_id: String(row.message_id || ''),
    group_id: String(row.group_id || ''),
    kind: normalizedKind,
    drive_file_id: driveFileId,
    file_name: fileName,
    mime_type: mimeType,
    size_bytes: Number(row.size_bytes || 0),
    view_url: String(row.view_url || ''),
    download_url: String(row.download_url || ''),
    thumb_url: String(row.thumb_url || '') || (normalizedKind === 'image' && driveFileId ? buildDriveUrls_(driveFileId, true).thumb_url : ''),
    created_at: toIsoString_(row.created_at),
    sort_order: Number(row.sort_order || 0)
  };
}

function updateRoomOnNewMessage_(roomId, previewText, senderEmail, createdAt) {
  const sheet = getSheetOrThrow_('rooms');
  const headers = getHeaders_(sheet);
  const headerMap = getHeaderIndexMap_(headers);
  const normalizedRoomId = normalizeRoomId_(roomId);

  if (sheet.getLastRow() < 2) {
    appendObjectRow_('rooms', {
      room_id: normalizedRoomId,
      title: normalizedRoomId,
      type: 'group',
      created_at: createdAt,
      created_by: senderEmail,
      last_message_at: '',
      last_message_text: '',
      last_sender: '',
      avatar_label: buildAvatarLabel_(normalizedRoomId),
      avatar_color: pickAvatarColor_(normalizedRoomId)
    });
  }

  let values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  let rowIndex = values.findIndex(row => normalizeRoomId_(row[headerMap.room_id]) === normalizedRoomId);

  if (rowIndex === -1) {
    appendObjectRow_('rooms', {
      room_id: normalizedRoomId,
      title: normalizedRoomId,
      type: 'group',
      created_at: createdAt,
      created_by: senderEmail,
      last_message_at: '',
      last_message_text: '',
      last_sender: '',
      avatar_label: buildAvatarLabel_(normalizedRoomId),
      avatar_color: pickAvatarColor_(normalizedRoomId)
    });
    values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    rowIndex = values.findIndex(row => normalizeRoomId_(row[headerMap.room_id]) === normalizedRoomId);
    if (rowIndex === -1) throw new Error('Комната не найдена.');
  }

  values[rowIndex][headerMap.last_message_at] = createdAt;
  values[rowIndex][headerMap.last_message_text] = buildPreviewText_(previewText);
  values[rowIndex][headerMap.last_sender] = senderEmail;
  values[rowIndex][headerMap.type] = values[rowIndex][headerMap.type] || 'group';
  values[rowIndex][headerMap.avatar_label] = values[rowIndex][headerMap.avatar_label] || String(values[rowIndex][headerMap.title] || normalizedRoomId || '#').trim().slice(0, 2).toUpperCase();
  values[rowIndex][headerMap.avatar_color] = values[rowIndex][headerMap.avatar_color] || pickAvatarColor_(String(values[rowIndex][headerMap.title] || normalizedRoomId || '#'));

  sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).setValues([values[rowIndex]]);
}

function setRoomMemberLastRead_(roomId, email, timestamp) {
  const sheet = getSheetOrThrow_('room_members');
  const headers = getHeaders_(sheet);
  const headerMap = getHeaderIndexMap_(headers);

  if (sheet.getLastRow() < 2) return false;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const normalizedRoomId = normalizeRoomId_(roomId);
  const normalizedEmail = normalizeEmail_(email);

  const rowIndex = values.findIndex(row => {
    return normalizeRoomId_(row[headerMap.room_id]) === normalizedRoomId &&
      normalizeEmail_(row[headerMap.user_email]) === normalizedEmail;
  });

  if (rowIndex === -1) return false;

  values[rowIndex][headerMap.last_read_at] = timestamp;
  values[rowIndex][headerMap.joined_at] = values[rowIndex][headerMap.joined_at] || timestamp;
  values[rowIndex][headerMap.role] = values[rowIndex][headerMap.role] || 'member';
  values[rowIndex][headerMap.is_muted] = values[rowIndex][headerMap.is_muted] || 'false';

  sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).setValues([values[rowIndex]]);
  return true;
}

function upsertCurrentUser_() {
  const email = getCurrentUserEmailOrThrow_();
  upsertUserProfile_(email);
}

function upsertUserProfile_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return;

  const sheet = getSheetOrThrow_('users');
  const headers = getHeaders_(sheet);
  const headerMap = getHeaderIndexMap_(headers);
  const nowIso = new Date().toISOString();

  if (sheet.getLastRow() < 2) {
    appendObjectRow_('users', {
      email: normalizedEmail,
      display_name: '',
      display_name_set_at: '',
      avatar_label: buildAvatarLabel_(normalizedEmail),
      avatar_color: pickAvatarColor_(normalizedEmail),
      last_seen_at: nowIso
    });
    return;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const rowIndex = values.findIndex(row => normalizeEmail_(row[headerMap.email]) === normalizedEmail);

  if (rowIndex === -1) {
    appendObjectRow_('users', {
      email: normalizedEmail,
      display_name: '',
      display_name_set_at: '',
      avatar_label: buildAvatarLabel_(normalizedEmail),
      avatar_color: pickAvatarColor_(normalizedEmail),
      last_seen_at: nowIso
    });
    return;
  }

  values[rowIndex][headerMap.avatar_label] = values[rowIndex][headerMap.avatar_label] || buildAvatarLabel_(normalizedEmail);
  values[rowIndex][headerMap.avatar_color] = values[rowIndex][headerMap.avatar_color] || pickAvatarColor_(normalizedEmail);
  values[rowIndex][headerMap.last_seen_at] = nowIso;

  sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).setValues([values[rowIndex]]);
}

function setUserDisplayName_(email, displayName) {
  const normalizedEmail = normalizeEmail_(email);
  const cleanName = String(displayName || '').trim().replace(/\s+/g, ' ');
  if (!normalizedEmail || !cleanName) {
    throw new Error('Не удалось сохранить имя пользователя.');
  }

  upsertUserProfile_(normalizedEmail);

  const sheet = getSheetOrThrow_('users');
  const headers = getHeaders_(sheet);
  const headerMap = getHeaderIndexMap_(headers);
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    : [];
  const rowIndex = values.findIndex(row => normalizeEmail_(row[headerMap.email]) === normalizedEmail);

  if (rowIndex === -1) {
    throw new Error('Профиль пользователя не найден.');
  }

  values[rowIndex][headerMap.display_name] = cleanName;
  if (headerMap.display_name_set_at != null) {
    values[rowIndex][headerMap.display_name_set_at] = new Date().toISOString();
  }
  values[rowIndex][headerMap.avatar_label] = buildAvatarLabel_(cleanName);
  values[rowIndex][headerMap.avatar_color] = values[rowIndex][headerMap.avatar_color] || pickAvatarColor_(normalizedEmail);
  values[rowIndex][headerMap.last_seen_at] = new Date().toISOString();

  sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).setValues([values[rowIndex]]);
}

function getUsersMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  if (!sheet || sheet.getLastRow() < 2) return {};

  return readSheetObjects_('users').reduce((acc, row) => {
    const email = normalizeEmail_(row.email);
    if (email) acc[email] = row;
    return acc;
  }, {});
}

/* =========================
   Helpers: attachments
   ========================= */

function validateFilesPayload_(files) {
  if (!files.length) return;

  let total = 0;
  files.forEach(file => {
    const size = Number(file && file.size || 0);
    const name = String(file && file.name || 'file');
    if (!file || !file.base64) {
      throw new Error('Не удалось подготовить файл "' + name + '".');
    }
    if (size <= 0) {
      throw new Error('Файл "' + name + '" пустой.');
    }
    if (size > ATTACHMENT_MAX_FILE_BYTES) {
      throw new Error('Файл "' + name + '" превышает лимит ' + formatBytes_(ATTACHMENT_MAX_FILE_BYTES) + '.');
    }
    total += size;
  });

  if (total > ATTACHMENT_MAX_TOTAL_BYTES) {
    throw new Error('Суммарный размер вложений превышает ' + formatBytes_(ATTACHMENT_MAX_TOTAL_BYTES) + '.');
  }
}

function saveAttachment_(folder, messageId, groupId, payload, createdAt, sortOrder, memberEmails) {
  const name = sanitizeFileName_(payload.name || 'file');
  const mimeType = normalizeUploadMimeType_(payload.mimeType, name);
  const bytes = Utilities.base64Decode(String(payload.base64 || ''));
  const blob = Utilities.newBlob(bytes, mimeType, name);
  const file = folder.createFile(blob);

  lockFileToExplicitViewers_(file);
  syncFilePermissions_(file, memberEmails || []);

  const kind = isLikelyImageFile_(mimeType, name) ? 'image' : 'file';
  const urls = buildDriveUrls_(file.getId(), kind === 'image');

  return {
    id: Utilities.getUuid(),
    message_id: messageId,
    group_id: String(groupId || ''),
    kind: kind,
    drive_file_id: file.getId(),
    file_name: name,
    mime_type: mimeType,
    size_bytes: Number(payload.size || bytes.length || 0),
    view_url: urls.view_url,
    download_url: urls.download_url,
    thumb_url: kind === 'image' ? urls.thumb_url : '',
    created_at: createdAt,
    sort_order: Number(sortOrder || 0)
  };
}

function ensureRootAttachmentsFolder_() {
  const props = PropertiesService.getDocumentProperties();
  const savedId = props.getProperty(ROOT_FOLDER_PROP);
  if (savedId) {
    try {
      return DriveApp.getFolderById(savedId);
    } catch (err) {}
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const folderName = '[' + spreadsheet.getName() + '] Messenger attachments';
  let folder;

  try {
    const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    const parents = spreadsheetFile.getParents();
    if (parents.hasNext()) {
      folder = parents.next().createFolder(folderName);
    }
  } catch (err) {}

  if (!folder) {
    folder = DriveApp.createFolder(folderName);
  }

  props.setProperty(ROOT_FOLDER_PROP, folder.getId());
  return folder;
}

function getRoomFolder_(roomId, roomTitle) {
  const root = ensureRootAttachmentsFolder_();
  const folderName = String(roomId || '').trim() + '__' + sanitizeFileName_(roomTitle || roomId || 'room');
  const folders = root.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return root.createFolder(folderName);
}

function getRoomMemberEmails_(roomId) {
  return Array.from(new Set(
    readSheetObjects_('room_members')
      .filter(row => normalizeRoomId_(row.room_id) === normalizeRoomId_(roomId))
      .map(row => normalizeEmail_(row.user_email))
      .filter(Boolean)
  ));
}

function syncFilePermissions_(fileOrFileId, memberEmails) {
  const targetFile = typeof fileOrFileId === 'string'
    ? DriveApp.getFileById(fileOrFileId)
    : fileOrFileId;

  (memberEmails || []).forEach(email => {
    try {
      targetFile.addViewer(email);
    } catch (err) {
      // ignore permission sync issues per user
    }
  });

  return targetFile;
}

function lockFileToExplicitViewers_(fileOrFileId) {
  const targetFile = typeof fileOrFileId === 'string'
    ? DriveApp.getFileById(fileOrFileId)
    : fileOrFileId;

  try {
    targetFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  } catch (err) {
    // ignore privacy lock issues; explicit viewers may still work
  }

  return targetFile;
}

function syncRoomFolderPermissions_(roomId, folder) {
  const targetFolder = folder || getRoomFolder_(roomId, roomId);
  const memberEmails = getRoomMemberEmails_(roomId);

  memberEmails.forEach(email => {
    try {
      targetFolder.addViewer(email);
    } catch (err) {
      // ignore permission sync issues per user
    }
  });

  return {
    ok: true,
    folder_id: targetFolder.getId(),
    viewers_count: memberEmails.length
  };
}

function buildDriveUrls_(fileId, isImage) {
  const id = String(fileId || '').trim();
  return {
    view_url: 'https://drive.google.com/file/d/' + id + '/view',
    download_url: 'https://drive.google.com/uc?export=download&id=' + id,
    thumb_url: isImage
      ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600'
      : ''
  };
}

function revokeFilePermissions_(fileOrFileId, memberEmails) {
  const targetFile = typeof fileOrFileId === 'string'
    ? DriveApp.getFileById(fileOrFileId)
    : fileOrFileId;

  (memberEmails || []).forEach(email => {
    try {
      targetFile.removeViewer(email);
    } catch (err) {
      // ignore permission revoke issues per user
    }
  });

  return targetFile;
}

function getAttachmentFileIdsForRoom_(roomId) {
  const normalizedRoomId = normalizeRoomId_(roomId);
  const messageIds = new Set(
    readSheetObjects_('messages')
      .filter(row => !row.deleted_at)
      .filter(row => normalizeRoomId_(row.room_id) === normalizedRoomId)
      .map(row => String(row.id || '').trim())
      .filter(Boolean)
  );

  if (!messageIds.size) return [];

  return Array.from(new Set(
    readSheetObjects_('attachments')
      .filter(row => messageIds.has(String(row.message_id || '').trim()))
      .map(row => String(row.drive_file_id || '').trim())
      .filter(Boolean)
  ));
}

function syncExistingAttachmentPermissionsForRoom_(roomId, memberEmails) {
  const targetEmails = uniqueEmails_(memberEmails || []);
  if (!targetEmails.length) {
    return { ok: true, room_id: normalizeRoomId_(roomId), files_count: 0, viewers_count: 0 };
  }

  const fileIds = getAttachmentFileIdsForRoom_(roomId);
  fileIds.forEach(fileId => {
    lockFileToExplicitViewers_(fileId);
    syncFilePermissions_(fileId, targetEmails);
  });

  return {
    ok: true,
    room_id: normalizeRoomId_(roomId),
    files_count: fileIds.length,
    viewers_count: targetEmails.length
  };
}

function revokeExistingAttachmentPermissionsForRoom_(roomId, memberEmails) {
  const targetEmails = uniqueEmails_(memberEmails || []);
  if (!targetEmails.length) {
    return { ok: true, room_id: normalizeRoomId_(roomId), files_count: 0, viewers_count: 0 };
  }

  const fileIds = getAttachmentFileIdsForRoom_(roomId);
  fileIds.forEach(fileId => revokeFilePermissions_(fileId, targetEmails));

  return {
    ok: true,
    room_id: normalizeRoomId_(roomId),
    files_count: fileIds.length,
    viewers_count: targetEmails.length
  };
}

function revokeRoomFolderPermissions_(roomId, memberEmails, folder) {
  const targetFolder = folder || getRoomFolder_(roomId, roomId);
  const targetEmails = uniqueEmails_(memberEmails || []);

  targetEmails.forEach(email => {
    try {
      targetFolder.removeViewer(email);
    } catch (err) {
      // ignore permission revoke issues per user
    }
  });

  return {
    ok: true,
    folder_id: targetFolder.getId(),
    viewers_count: targetEmails.length
  };
}

function resyncAllAttachmentPermissionsFromMemberships_() {
  const roomIds = Array.from(new Set(readSheetObjects_('rooms')
    .map(row => normalizeRoomId_(row.room_id))
    .filter(Boolean)));

  let roomsCount = 0;
  let filesCount = 0;
  roomIds.forEach(roomId => {
    const memberEmails = getRoomMemberEmails_(roomId);
    syncRoomFolderPermissions_(roomId);
    const summary = syncExistingAttachmentPermissionsForRoom_(roomId, memberEmails);
    roomsCount += 1;
    filesCount += Number(summary.files_count || 0);
  });

  return {
    ok: true,
    rooms_count: roomsCount,
    files_count: filesCount,
    message: 'Доступ к вложениям пересинхронизирован. Комнат: ' + roomsCount + ', файлов: ' + filesCount + '.'
  };
}

function deriveMessageType_(attachments, text) {
  const hasText = !!String(text || '').trim();
  const hasImages = attachments.some(item => item.kind === 'image');
  const hasFiles = attachments.some(item => item.kind === 'file');

  if (hasText && (hasImages || hasFiles)) return 'mixed';
  if (hasImages && !hasFiles) return attachments.length > 1 ? 'gallery' : 'image';
  if (hasFiles && !hasImages) return 'file';
  return 'text';
}

function buildRoomLastMessageText_(text, attachments) {
  const cleanText = String(text || '').trim();
  const list = Array.isArray(attachments) ? attachments : [];

  if (cleanText) return cleanText;
  if (!list.length) return '';

  const imageCount = list.filter(item => item.kind === 'image').length;
  const fileCount = list.filter(item => item.kind === 'file').length;

  if (imageCount && !fileCount) {
    return imageCount === 1 ? '[Фото]' : '[Фото] ' + imageCount;
  }
  if (fileCount && !imageCount) {
    return fileCount === 1
      ? '[Файл] ' + String(list[0].file_name || '')
      : '[Файлы] ' + fileCount;
  }
  return '[Вложения] ' + list.length;
}

/* =========================
   Helpers: utils
   ========================= */

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeRoomId_(value) {
  return String(value || '').trim().toLowerCase();
}

function looksLikeEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function looksLikeIsoDate_(value) {
  const text = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(text);
}

function normalizeBooleanString_(value) {
  const text = String(value || '').trim().toLowerCase();
  if (text === 'true' || text === 'false') return text;
  if (text === '1') return 'true';
  if (text === '0') return 'false';
  return '';
}

function buildPreviewText_(value) {
  const singleLine = String(value || '').replace(/\s+/g, ' ').trim();
  if (singleLine.length <= 80) return singleLine;
  return singleLine.slice(0, 77) + '...';
}

function buildAvatarLabel_(value) {
  const source = String(value || '').trim();
  if (!source) return 'U';

  const parts = source
    .split(/\s+/)
    .map(part => part.trim())
    .filter(Boolean);

  if (parts.length >= 2) {
    return (parts[0][0] + parts[1][0]).toUpperCase();
  }

  const compact = source.replace(/[^\p{L}\p{N}]+/gu, '');
  return (compact.slice(0, 2) || source.slice(0, 2) || 'U').toUpperCase();
}

function pickAvatarColor_(seed) {
  const palette = [
    '#5b8def',
    '#60a5fa',
    '#34d399',
    '#f59e0b',
    '#f97316',
    '#a78bfa',
    '#ec4899',
    '#ef4444',
    '#14b8a6',
    '#8b5cf6'
  ];

  const source = String(seed || '#');
  let hash = 0;
  for (let i = 0; i < source.length; i += 1) {
    hash = ((hash << 5) - hash) + source.charCodeAt(i);
    hash |= 0;
  }

  return palette[Math.abs(hash) % palette.length];
}

function toDateMs_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') return value.getTime();
  const date = new Date(value);
  return isNaN(date.getTime()) ? 0 : date.getTime();
}

function toIsoString_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') return value.toISOString();
  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return date.toISOString();
}

function clampPageSize_(limit) {
  const numeric = Number(limit);
  if (!numeric || numeric < 1) return PAGE_SIZE;
  return Math.min(Math.floor(numeric), MAX_PAGE_SIZE);
}

function sanitizeFileName_(name) {
  return String(name || 'file')
    .replace(/[\\\/:*?"<>|#%]+/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 180) || 'file';
}

function slugify_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9а-яё_\-\s]/gi, '')
    .replace(/\s+/g, '-')
    .slice(0, 60);
}

function isImageMimeType_(mimeType) {
  return /^image\//i.test(String(mimeType || ''));
}

function isImageFileName_(fileName) {
  return /\.(apng|avif|bmp|gif|heic|heif|jpe?g|png|svg|tiff?|webp)$/i.test(String(fileName || ''));
}

function isLikelyImageFile_(mimeType, fileName) {
  return isImageMimeType_(mimeType) || isImageFileName_(fileName);
}

function normalizeUploadMimeType_(mimeType, fileName) {
  const raw = String(mimeType || '').trim();
  if (raw && raw !== 'application/octet-stream') {
    return raw;
  }

  const lower = String(fileName || '').toLowerCase();
  if (/\.jpe?g$/i.test(lower)) return 'image/jpeg';
  if (/\.png$/i.test(lower)) return 'image/png';
  if (/\.gif$/i.test(lower)) return 'image/gif';
  if (/\.webp$/i.test(lower)) return 'image/webp';
  if (/\.bmp$/i.test(lower)) return 'image/bmp';
  if (/\.svg$/i.test(lower)) return 'image/svg+xml';
  if (/\.avif$/i.test(lower)) return 'image/avif';
  if (/\.tiff?$/i.test(lower)) return 'image/tiff';
  if (/\.heic$/i.test(lower)) return 'image/heic';
  if (/\.heif$/i.test(lower)) return 'image/heif';
  return raw || 'application/octet-stream';
}

function formatBytes_(bytes) {
  const value = Number(bytes || 0);
  if (value < 1024) return value + ' B';
  if (value < 1024 * 1024) return (value / 1024).toFixed(1) + ' KB';
  return (value / (1024 * 1024)).toFixed(1) + ' MB';
}
