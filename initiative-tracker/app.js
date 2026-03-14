const APP_STORAGE_KEY = "initiativeTrackerState.v2";
const DAMAGE_MONITOR_COLUMN_COUNT = 16;
const DAMAGE_QM_VALUES = ["-", "-", "-", "-", "-", "-", "-", "-1", "-2", "-3", "-4", "-7", "-8", "-9", "-12", "-15"];
const DAMAGE_BEW_VALUES = ["-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-1", "-1", "-2", "-3", "-5", "-7"];
const PIXELS_MODE = {
  PC_SET_3: "pc-set-3",
  PC_SINGLE_3X: "pc-single-3x",
  SHARED_SET_3: "shared-set-3",
};

const state = {
  characters: [],
  nextId: 1,
  turnEntries: [],
  currentTurnIndex: -1,
  activeCharacterId: null,
  round: 0,
  turnHistory: [],
  turnHistoryIndex: -1,
  usedTurnEntryIds: [],
  turnOrderUndoStack: [],
  turnOrderRedoStack: [],
  characterUndoStack: [],
  characterRedoStack: [],
  combatLog: [],
  nextLogId: 1,
  activeTurnTab: "order",
  activeAbilityCharacterIds: [],
  dazedAppliedCharacterIds: [],
  dazedAppliedRound: 0,
  uiSettings: {
    fontScale: 1,
    forceOneColumn: false,
  },
};

const addForm = document.getElementById("add-form");
const saveFullSnapshotCsvBtn = document.getElementById("save-full-snapshot-csv");
const loadFullSnapshotCsvInput = document.getElementById("load-full-snapshot-csv");
const savePartyCsvBtn = document.getElementById("save-party-csv");
const loadPartyCsvInput = document.getElementById("load-party-csv");
const saveEnemyCsvBtn = document.getElementById("save-enemy-csv");
const enemyGroupCsvInput = document.getElementById("enemy-group-csv");
const rosterEl = document.getElementById("roster");
const trackerEl = document.getElementById("tracker");
const combatLogEl = document.getElementById("combat-log");
const roundStatusEl = document.getElementById("round-status");
const activateCharacterUpBtn = document.getElementById("activate-character-up");
const activateCharacterDownBtn = document.getElementById("activate-character-down");
const nextTurnBtn = document.getElementById("next-turn");
const headerUndoBtn = document.getElementById("header-undo");
const headerRedoBtn = document.getElementById("header-redo");
const endCombatBtn = document.getElementById("end-combat");
const characterSettingsMenu = document.getElementById("character-settings-menu");
const removeAllCharactersBtn = document.getElementById("remove-all-characters");
const removePcCharactersBtn = document.getElementById("remove-pc-characters");
const removeNpcCharactersBtn = document.getElementById("remove-npc-characters");
const surpriseAllPcEl = document.getElementById("surprise-all-pc");
const surpriseAllNpcEl = document.getElementById("surprise-all-npc");
const openPixelsSettingsBtn = document.getElementById("open-pixels-settings");
const pixelsSettingsDialogEl = document.getElementById("pixels-settings-dialog");
const pixelsSettingsCloseBtn = document.getElementById("pixels-settings-close");
const pixelsUseForRollsEl = document.getElementById("pixels-use-for-rolls");
const pixelsSimulatedEl = document.getElementById("pixels-simulated");
const pixelsModeSelectEl = document.getElementById("pixels-mode-select");
const pixelsModeHelpEl = document.getElementById("pixels-mode-help");
const pixelsAssignmentListEl = document.getElementById("pixels-assignment-list");
const pixelsDisconnectAllBtn = document.getElementById("pixels-disconnect-all");
const pixelsRollDialogEl = document.getElementById("pixels-roll-dialog");
const pixelsRollTitleEl = document.getElementById("pixels-roll-title");
const pixelsRollStatusEl = document.getElementById("pixels-roll-status");
const pixelsRollListEl = document.getElementById("pixels-roll-list");
const pixelsStatusEl = document.getElementById("pixels-status");
const clearLogBtn = document.getElementById("clear-log");
const characterTemplate = document.getElementById("character-template");
const editCharacterDialogEl = document.getElementById("edit-character-dialog");
const editCharacterFormEl = document.getElementById("edit-character-form");
const editCharacterTitleEl = document.getElementById("edit-character-title");
const editCharacterCancelBtn = document.getElementById("edit-character-cancel");
const editRollInitiativeBtn = document.getElementById("edit-roll-initiative");
const warningDialogEl = document.getElementById("warning-dialog");
const warningDialogTitleEl = document.getElementById("warning-dialog-title");
const warningDialogMessageEl = document.getElementById("warning-dialog-message");
const warningDialogConfirmBtn = document.getElementById("warning-dialog-confirm");
const warningDialogCancelBtn = document.getElementById("warning-dialog-cancel");
const editNameEl = document.getElementById("edit-name");
const editTypeEl = document.getElementById("edit-type");
const editIniEl = document.getElementById("edit-ini");
const editUseManualRollEl = document.getElementById("edit-use-manual-roll");
const editSpecialAbilityEl = document.getElementById("edit-special-ability");
const editAlwaysMinimizedEl = document.getElementById("edit-always-minimized");
const appVersionEl = document.getElementById("app-version");
const appShellEl = document.querySelector(".app-shell");
const fontSizeSliderEl = document.getElementById("font-size-slider");
const forceOneColumnEl = document.getElementById("force-one-column");
const resetAppStateBtn = document.getElementById("reset-app-state");
const settingsMenuEls = Array.from(document.querySelectorAll(".settings-menu"));

const pixels = {
  sdk: null,
  loading: false,
  reconnecting: false,
  simulated: false,
  useForRolls: false,
  mode: PIXELS_MODE.PC_SET_3,
  assignmentsByCharacterId: {},
  sharedSet: [null, null, null],
  rememberedAssignmentsByCharacterId: {},
  rememberedSharedSet: [null, null, null],
  lastStatus:
    "Pixels: nicht verbunden (Chromium-Browser mit Web Bluetooth erforderlich).",
};
const PIXELS_DEBUG = true;

let editingCharacterId = null;
let warningDialogResolver = null;
let draggedRosterCharacterId = null;
let draggedTurnGroupIndex = null;
const rollDialogState = {
  rowsByCharacterId: new Map(),
};

addForm.addEventListener("submit", (event) => {
  event.preventDefault();

  const formData = new FormData(addForm);
  const name = String(formData.get("name") || "").trim();
  const type = String(formData.get("type") || "PC");
  const iniValue = Number(formData.get("ini"));
  const specialAbility = String(formData.get("specialAbility") || "").trim();

  if (!name) {
    return;
  }

  pushCharacterUndoSnapshot("Charakter-Änderung zurückgesetzt.");
  state.characters.push(createCharacter({ name, type, ini: iniValue, specialAbility: specialAbility || null }));

  addForm.reset();
  addForm.ini.value = "10";
  addForm.type.value = "PC";
  addForm.specialAbility.value = "";

  if (state.round <= 0 || !state.turnEntries.length) {
    invalidateTurnState();
  } else {
    // Laufende KR beibehalten; neuer Charakter wird erst per Nachwurf integriert.
    syncCurrentSnapshotFromState();
  }
  logEvent(`Charakter hinzugefügt: ${name} (${type}), INI ${clamp(iniValue, 1, 30)}.`);
  persistAppState();
  render();
});

nextTurnBtn.addEventListener("click", async () => {
  await generateNewTurn();
});

function activateAdjacentCharacter(direction) {
  if (!state.turnEntries.length) {
    return;
  }

  const targetCharacterId = getAdjacentCharacterIdForTurnOrder(direction);
  if (targetCharacterId === null) {
    return;
  }

  pushTurnOrderUndoSnapshot("INI-Reihenfolge-Status wiederhergestellt.");
  state.activeCharacterId = targetCharacterId;
  syncActiveTurnPointer();
  const activeCharacter = state.characters.find((character) => character.id === targetCharacterId);
  if (activeCharacter) {
    logEvent(`Aktiver Charakter gewechselt: ${activeCharacter.name}.`);
  }

  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

activateCharacterUpBtn.addEventListener("click", () => {
  activateAdjacentCharacter(-1);
});

activateCharacterDownBtn.addEventListener("click", () => {
  activateAdjacentCharacter(1);
});

if (headerUndoBtn) {
  headerUndoBtn.addEventListener("click", () => {
    undoLatestPanelAction();
  });
}

if (headerRedoBtn) {
  headerRedoBtn.addEventListener("click", () => {
    redoLatestPanelAction();
  });
}

endCombatBtn.addEventListener("click", async () => {
  const confirmed = await showWarningDialog("Kampf beenden und aktuellen KR-Status löschen?", "Kampf beenden");
  if (!confirmed) {
    return;
  }

  invalidateTurnState({ clearDazed: true });
  logEvent("Kampf beendet.");
  persistAppState();
  render();
});

removeAllCharactersBtn.addEventListener("click", async () => {
  const confirmed = await showWarningDialog("Alle Charaktere/NSC entfernen?", "Alle entfernen");
  if (!confirmed) {
    return;
  }

  const beforeCount = state.characters.length;
  if (beforeCount <= 0) return;
  pushCharacterUndoSnapshot("Charaktere wiederhergestellt.");
  state.characters = [];
  const removed = beforeCount;
  closeCharacterSettingsMenu();

  invalidateTurnState({ clearDazed: true });
  state.combatLog = [];
  state.nextLogId = 1;
  persistAppState();
  render();
});

removePcCharactersBtn.addEventListener("click", async () => {
  const confirmed = await showWarningDialog("Alle SC entfernen?", "SC entfernen");
  if (!confirmed) {
    return;
  }

  const beforeCount = state.characters.length;
  const remaining = state.characters.filter((character) => character.type !== "PC");
  const removed = beforeCount - remaining.length;
  closeCharacterSettingsMenu();
  if (removed <= 0) return;
  pushCharacterUndoSnapshot("SC wiederhergestellt.");
  state.characters = remaining;

  invalidateTurnState({ clearDazed: true });
  logEvent(`SC entfernt (${removed}).`);
  persistAppState();
  render();
});

removeNpcCharactersBtn.addEventListener("click", async () => {
  const confirmed = await showWarningDialog("Alle NSC entfernen?", "NSC entfernen");
  if (!confirmed) {
    return;
  }

  const beforeCount = state.characters.length;
  const remaining = state.characters.filter((character) => character.type !== "NPC");
  const removed = beforeCount - remaining.length;
  closeCharacterSettingsMenu();
  if (removed <= 0) return;
  pushCharacterUndoSnapshot("NSC wiederhergestellt.");
  state.characters = remaining;

  invalidateTurnState({ clearDazed: true });
  logEvent(`NSC entfernt (${removed}).`);
  persistAppState();
  render();
});

if (surpriseAllPcEl) {
  surpriseAllPcEl.addEventListener("change", (event) => {
    toggleSurprisedByType("PC", event.target.checked);
  });
}

if (surpriseAllNpcEl) {
  surpriseAllNpcEl.addEventListener("change", (event) => {
    toggleSurprisedByType("NPC", event.target.checked);
  });
}

clearLogBtn.addEventListener("click", () => {
  state.combatLog = [];
  state.nextLogId = 1;
  persistAppState();
  renderCombatLog();
});

if (editCharacterCancelBtn) {
  editCharacterCancelBtn.addEventListener("click", () => {
    editingCharacterId = null;
    editCharacterDialogEl?.close();
  });
}

if (editCharacterDialogEl) {
  editCharacterDialogEl.addEventListener("close", () => {
    editingCharacterId = null;
  });
}

if (warningDialogConfirmBtn) {
  warningDialogConfirmBtn.addEventListener("click", () => {
    resolveWarningDialog(true);
  });
}

if (warningDialogCancelBtn) {
  warningDialogCancelBtn.addEventListener("click", () => {
    resolveWarningDialog(false);
  });
}

if (warningDialogEl) {
  warningDialogEl.addEventListener("close", () => {
    resolveWarningDialog(false);
  });
}

if (editCharacterFormEl) {
  editCharacterFormEl.addEventListener("submit", (event) => {
    event.preventDefault();
    saveCharacterEdits();
  });
}

if (editRollInitiativeBtn) {
  editRollInitiativeBtn.addEventListener("click", async () => {
    if (editingCharacterId === null) {
      return;
    }
    await rollCharacterIntoCurrentRound(editingCharacterId);
  });
}

if (editTypeEl) {
  editTypeEl.addEventListener("change", () => {
    updateEditDialogFieldStates();
  });
}

if (editUseManualRollEl) {
  editUseManualRollEl.addEventListener("change", () => {
    updateEditDialogFieldStates();
  });
}

if (openPixelsSettingsBtn) {
  openPixelsSettingsBtn.addEventListener("click", () => {
    openPixelsSettingsDialog();
  });
}

if (pixelsSettingsCloseBtn) {
  pixelsSettingsCloseBtn.addEventListener("click", () => {
    closePixelsSettingsDialog();
  });
}

if (pixelsRollDialogEl) {
  pixelsRollDialogEl.addEventListener("cancel", (event) => {
    event.preventDefault();
  });
}

if (pixelsUseForRollsEl) {
  pixelsUseForRollsEl.addEventListener("change", () => {
    pixels.useForRolls = Boolean(pixelsUseForRollsEl.checked);
    logEvent(`Pixels-Würfe für INI ${pixels.useForRolls ? "aktiviert" : "deaktiviert"}.`);
    persistAppState();
    renderCombatLog();
    updatePixelsControls();
  });
}

if (pixelsSimulatedEl) {
  pixelsSimulatedEl.addEventListener("change", () => {
    pixels.simulated = Boolean(pixelsSimulatedEl.checked);
    if (!pixels.simulated && !hasAnyConnectedPixels()) {
      pixels.useForRolls = false;
    }
    logEvent(`Pixels-Simulation ${pixels.simulated ? "aktiviert" : "deaktiviert"}.`);
    persistAppState();
    renderCombatLog();
    updatePixelsControls();
    renderPixelsSettingsDialog();
  });
}

if (pixelsModeSelectEl) {
  pixelsModeSelectEl.addEventListener("change", () => {
    pixels.mode = normalizePixelsMode(pixelsModeSelectEl.value);
    persistAppState();
    updatePixelsControls();
    renderPixelsSettingsDialog();
  });
}

if (pixelsDisconnectAllBtn) {
  pixelsDisconnectAllBtn.addEventListener("click", async () => {
    await disconnectAllPixels();
    renderPixelsSettingsDialog();
  });
}

if (fontSizeSliderEl) {
  fontSizeSliderEl.addEventListener("change", () => {
    const fontScale = clamp(Number(fontSizeSliderEl.value) / 100, 0.7, 1.3);
    applyUiSettings({
      fontScale,
      forceOneColumn: state.uiSettings.forceOneColumn,
    });
    persistAppState();
  });
}

if (forceOneColumnEl) {
  forceOneColumnEl.addEventListener("change", () => {
    applyUiSettings({
      fontScale: state.uiSettings.fontScale,
      forceOneColumn: Boolean(forceOneColumnEl.checked),
    });
    persistAppState();
  });
}

if (resetAppStateBtn) {
  resetAppStateBtn.addEventListener("click", async () => {
    const confirmed = await showWarningDialog(
      "Die gesamte App wird zurückgesetzt. Alle Daten, Kampfstatus und Logs gehen verloren. Fortfahren?",
      "App resetten"
    );
    if (!confirmed) {
      return;
    }

    resetEntireAppState();
  });
}

if (savePartyCsvBtn) {
  savePartyCsvBtn.addEventListener("click", async () => {
    await exportPcGroupToCsv();
  });
}

if (saveFullSnapshotCsvBtn) {
  saveFullSnapshotCsvBtn.addEventListener("click", async () => {
    await exportFullSnapshotToCsv();
  });
}

if (loadFullSnapshotCsvInput) {
  loadFullSnapshotCsvInput.addEventListener("change", async (event) => {
    const file = event.target?.files?.[0];
    if (!file) {
      return;
    }

    try {
      const content = await file.text();
      const rows = parseSimpleCsv(content);
      const restored = importFullSnapshotRows(rows);
      if (restored) {
        logEvent("Vollsnapshot geladen.");
        persistAppState();
        render();
      } else {
        logEvent("Vollsnapshot-CSV enthält keine gültigen Snapshot-Daten.");
        renderCombatLog();
      }
    } catch {
      logEvent("Vollsnapshot-CSV konnte nicht gelesen werden.");
      renderCombatLog();
    } finally {
      event.target.value = "";
    }
  });
}

if (saveEnemyCsvBtn) {
  saveEnemyCsvBtn.addEventListener("click", async () => {
    await exportNpcGroupToCsv();
  });
}

if (loadPartyCsvInput) {
  loadPartyCsvInput.addEventListener("change", async (event) => {
    const file = event.target?.files?.[0];
    if (!file) {
      return;
    }

    try {
      const content = await file.text();
      const rows = parseSimpleCsv(content);
      const imported = importPcGroupRows(rows);
      if (imported > 0) {
        logEvent(`SC-Gruppe aus CSV geladen: ${imported} SC.`);
        persistAppState();
        render();
      }
    } catch {
      logEvent("Gruppen-CSV konnte nicht gelesen werden.");
      persistAppState();
      renderCombatLog();
    } finally {
      event.target.value = "";
    }
  });
}

if (enemyGroupCsvInput) {
  enemyGroupCsvInput.addEventListener("change", async (event) => {
    const file = event.target?.files?.[0];
    if (!file) {
      return;
    }

    try {
      const content = await file.text();
      const rows = parseSimpleCsv(content);
      const imported = importEnemyGroupRows(rows);
      if (imported > 0) {
        logEvent(`Gegnergruppe aus CSV geladen: ${imported} NSC.`);
        persistAppState();
        render();
      }
    } catch {
      logEvent("Gegner-CSV konnte nicht gelesen werden.");
      persistAppState();
      renderCombatLog();
    } finally {
      event.target.value = "";
    }
  });
}

async function generateNewTurn() {
  if (!state.characters.length) {
    return;
  }

  const surprisedCharacters = state.characters.filter((character) => character.surprised);
  if (surprisedCharacters.length) {
    const keepSurprisedStatus = await showWarningDialog(
      "Einige Charaktere/NSC sind überrascht. Soll der Status für diese Kampfrunde bestehen bleiben?",
      "Überrascht-Status",
      "Ja",
      "Nein"
    );
    if (!keepSurprisedStatus) {
      state.characters = state.characters.map((character) => ({
        ...character,
        surprised: false,
      }));
      logEvent("Überraschung vor KR-Start für alle Charaktere/NSC deaktiviert.");
    }
  }

  const usePixelsForTurn = canUsePixelsForRolls();
  const rolledCharacters = [];
  const rollDetailsByCharacterId = new Map();
  const shouldShowPixelsRollDialog = preparePixelsRollDialog(state.characters, usePixelsForTurn, "INI Würfelphase");
  if (shouldShowPixelsRollDialog) {
    setPixelsRollDialogStatus("Warte auf die anstehenden INI-Würfe.");
  }

  try {
    for (const character of state.characters) {
      if (character.incapacitated) {
        if (shouldShowPixelsRollDialog) {
          updatePixelsRollDialogRow(character.id, {
            mode: "skip",
            stateClass: "done",
            status: "Übersprungen.",
            detail: "Aktionsunfähig - kein INI-Wurf in dieser KR.",
            clearControls: true,
            resultBadges: [],
          });
        }
        rolledCharacters.push({
          ...character,
          lastRoll: null,
          critBonusRoll: null,
          totalInitiative: null,
          lastPixelsFaces: null,
          unfreeDefensePenalty: 0,
          paradeClickCount: 0,
          minimized: character.alwaysMinimized ? true : character.minimized,
          manualExpandedUntilNextTurn: false,
        });
        continue;
      }

      const rollMode = getCharacterRollMode(character, usePixelsForTurn);
      if (shouldShowPixelsRollDialog) {
        updatePixelsRollDialogRow(character.id, {
          mode: rollMode,
          stateClass: rollMode === "auto" ? "pending" : "waiting",
          status:
            rollMode === "manual"
              ? "Warte auf manuellen Wurf."
              : rollMode === "pixels"
                ? "Warte auf Pixels-Wurf."
                : "Automatischer Wurf läuft.",
          detail:
            rollMode === "manual"
              ? "3w6 im Dialog eingeben."
              : rollMode === "pixels"
                ? "Warte auf den verbundenen Würfel."
                : "Dieser Wurf wird automatisch aufgelöst.",
        });
      }
      const rollData = await resolveCharacterRoll(character, usePixelsForTurn);
      const roll = rollData.total;
      rollDetailsByCharacterId.set(character.id, rollData);
      const totalInitiative = computeTotalInitiative(character, {
        lastRoll: roll,
        critBonusRoll: rollData.critBonusRoll,
        surprised: character.surprised,
      });
      if (shouldShowPixelsRollDialog) {
        updatePixelsRollDialogRowResult(character, rollData, totalInitiative, rollMode);
      }

      rolledCharacters.push({
        ...character,
        lastRoll: roll,
        critBonusRoll: rollData.critBonusRoll,
        manualRoll:
          rollData.source === "manual" && Number.isFinite(Number(rollData.manualRollValue))
            ? clamp(Number(rollData.manualRollValue), 3, 18)
            : character.manualRoll,
        totalInitiative,
        lastPixelsFaces: Array.isArray(rollData.faces) ? rollData.faces.map((face) => Math.round(Number(face) || 0)) : null,
        unfreeDefensePenalty: 0,
        paradeClickCount: 0,
        minimized: character.alwaysMinimized ? true : character.minimized,
        manualExpandedUntilNextTurn: false,
      });
    }
  } finally {
    if (shouldShowPixelsRollDialog) {
      setPixelsRollDialogStatus("INI-Würfe abgeschlossen.");
    }
    closePixelsRollDialog();
  }
  state.characters = rolledCharacters;

  const nextTurnNumber = state.round + 1;
  state.round = nextTurnNumber;
  state.activeAbilityCharacterIds = [];
  state.dazedAppliedCharacterIds = [];
  state.dazedAppliedRound = nextTurnNumber;
  state.turnOrderUndoStack = [];
  state.turnOrderRedoStack = [];
  const entries = buildSortedTrackerEntries();
  state.turnEntries = entries;
  state.usedTurnEntryIds = [];
  state.activeCharacterId = entries.length ? entries[0].characterId : null;
  syncActiveTurnPointer();

  if (state.turnHistoryIndex < state.turnHistory.length - 1) {
    state.turnHistory = state.turnHistory.slice(0, state.turnHistoryIndex + 1);
  }

  state.turnHistory.push(captureTurnSnapshot(nextTurnNumber));
  state.turnHistoryIndex = state.turnHistory.length - 1;

  const mainActions = entries.filter((entry) => entry.turn === "Main");
  logEvent(`KR ${nextTurnNumber} gewürfelt (${mainActions.length} Charaktere).`);
  for (const action of mainActions) {
    const character = state.characters.find((item) => item.id === action.characterId);
    if (!character) {
      continue;
    }

    const critText =
      action.critical === "failure"
        ? ", Krit. Fehlschlag"
        : action.critical === "success"
          ? ", Krit. Erfolg"
          : "";
    const bonusText = action.groupInitiative > 30 ? ", Bonus" : "";
    const dazedText = action.dazed ? ", Benommenheit!" : "";
    const rollDetails = rollDetailsByCharacterId.get(action.characterId) || null;
    const sourceText =
      rollDetails?.source === "manual"
        ? ", manuell"
        : rollDetails?.source === "pixels-sim"
          ? `, Pixels Sim [${rollDetails.faces.join(", ")}]`
        : rollDetails?.source === "pixels"
          ? `, Pixels [${rollDetails.faces.join(", ")}]`
          : rollDetails?.source === "pixels-fallback"
            ? `, Pixels Fallback (${rollDetails.error || "unbekannter Grund"})`
            : "";
    const critBonusRollText =
      action.critical === "success" && Number.isFinite(Number(character.critBonusRoll))
        ? `, Krit-W6=${character.critBonusRoll}`
        : "";
    logEvent(
      `${action.name}: 3w6=${character.lastRoll}, INI ${character.ini}, ges. ${action.groupInitiative}${sourceText}${critBonusRollText}${bonusText}${critText}${dazedText}.`
    );
  }

  updatePixelsControls();

  persistAppState();
  render();
}

function captureTurnSnapshot(turnNumber) {
  return {
    turn: turnNumber,
    currentTurnIndex: state.currentTurnIndex,
    activeCharacterId: state.activeCharacterId,
    usedTurnEntryIds: [...state.usedTurnEntryIds],
    activeAbilityCharacterIds: [...state.activeAbilityCharacterIds],
    dazedAppliedCharacterIds: [...state.dazedAppliedCharacterIds],
    dazedAppliedRound: state.dazedAppliedRound,
    entries: state.turnEntries.map((entry) => ({ ...entry })),
    rolls: state.characters.map((character) => ({
      id: character.id,
      lastRoll: character.lastRoll,
      critBonusRoll: character.critBonusRoll ?? null,
      totalInitiative: character.totalInitiative,
      lastPixelsFaces: Array.isArray(character.lastPixelsFaces)
        ? character.lastPixelsFaces.map((face) => Math.round(Number(face) || 0))
        : null,
      dazedUntilRound: character.dazedUntilRound ?? null,
      unfreeDefensePenalty: Math.max(0, Math.round(Number(character.unfreeDefensePenalty) || 0)),
      paradeClickCount: Math.max(0, Math.round(Number(character.paradeClickCount) || 0)),
    })),
  };
}

function applyTurnSnapshot(snapshot) {
  const byId = new Map(snapshot.rolls.map((roll) => [roll.id, roll]));

  state.characters = state.characters.map((character) => {
    const rollData = byId.get(character.id);
    if (!rollData) {
      return {
        ...character,
        lastRoll: null,
        critBonusRoll: null,
        totalInitiative: null,
        lastPixelsFaces: null,
        dazedUntilRound: rollData?.dazedUntilRound ?? character.dazedUntilRound ?? null,
        unfreeDefensePenalty: 0,
        paradeClickCount: 0,
      };
    }

    return {
      ...character,
      lastRoll: rollData.lastRoll,
      critBonusRoll: rollData.critBonusRoll ?? null,
      totalInitiative: rollData.totalInitiative,
      lastPixelsFaces: Array.isArray(rollData.lastPixelsFaces)
        ? rollData.lastPixelsFaces.map((face) => Math.round(Number(face) || 0))
        : null,
      dazedUntilRound: rollData.dazedUntilRound ?? character.dazedUntilRound ?? null,
      unfreeDefensePenalty: Math.max(0, Math.round(Number(rollData.unfreeDefensePenalty) || 0)),
      paradeClickCount: Math.max(0, Math.round(Number(rollData.paradeClickCount) || 0)),
    };
  });

  state.turnEntries = snapshot.entries.map((entry) => ({ ...entry }));
  state.usedTurnEntryIds = Array.isArray(snapshot.usedTurnEntryIds)
    ? snapshot.usedTurnEntryIds.filter((id) => typeof id === "string")
    : [];
  const validEntryIds = new Set(state.turnEntries.map((entry) => entry.id));
  state.usedTurnEntryIds = state.usedTurnEntryIds.filter((id) => validEntryIds.has(id));
  state.activeAbilityCharacterIds = Array.isArray(snapshot.activeAbilityCharacterIds)
    ? snapshot.activeAbilityCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
    : [...new Set(state.turnEntries.filter((entry) => entry.abilityActive).map((entry) => entry.characterId))];
  state.dazedAppliedCharacterIds = Array.isArray(snapshot.dazedAppliedCharacterIds)
    ? snapshot.dazedAppliedCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
    : [];
  state.dazedAppliedRound = Math.max(0, Math.round(Number(snapshot.dazedAppliedRound) || state.round || 0));
  state.activeCharacterId =
    snapshot.activeCharacterId === null || snapshot.activeCharacterId === undefined
      ? null
      : Number.isFinite(Number(snapshot.activeCharacterId))
        ? Number(snapshot.activeCharacterId)
        : null;
  syncActiveTurnPointer();
  state.round = Math.max(0, Number(snapshot.turn) || 0);
  state.turnOrderUndoStack = [];
  state.turnOrderRedoStack = [];
}

function syncCurrentSnapshotFromState() {
  if (state.turnHistoryIndex < 0 || state.turnHistoryIndex >= state.turnHistory.length) {
    return;
  }

  state.turnHistory[state.turnHistoryIndex] = captureTurnSnapshot(state.round);
}

function invalidateTurnState(options = {}) {
  const { keepCharacterRolls = false, clearDazed = false } = options;

  if (!keepCharacterRolls || clearDazed) {
    state.characters = state.characters.map((character) => ({
      ...character,
      lastRoll: keepCharacterRolls ? character.lastRoll : null,
      critBonusRoll: keepCharacterRolls ? character.critBonusRoll ?? null : null,
      totalInitiative: keepCharacterRolls ? character.totalInitiative : null,
      lastPixelsFaces: keepCharacterRolls ? character.lastPixelsFaces ?? null : null,
      dazedUntilRound: clearDazed ? null : character.dazedUntilRound,
      unfreeDefensePenalty: 0,
      paradeClickCount: 0,
    }));
  }

  state.turnEntries = [];
  state.currentTurnIndex = -1;
  state.activeCharacterId = null;
  state.round = 0;
  state.activeAbilityCharacterIds = [];
  state.dazedAppliedCharacterIds = [];
  state.dazedAppliedRound = 0;
  state.turnHistory = [];
  state.turnHistoryIndex = -1;
  state.usedTurnEntryIds = [];
  state.turnOrderUndoStack = [];
  state.turnOrderRedoStack = [];
}

function removeCharacter(id) {
  const character = state.characters.find((item) => item.id === id);
  if (!character) {
    return;
  }
  pushCharacterUndoSnapshot("Charakter wiederhergestellt.");
  state.characters = state.characters.filter((character) => character.id !== id);

  if (state.round > 0 && state.turnEntries.length) {
    rebuildTurnEntriesPreserveActive();
    state.turnOrderUndoStack = [];
    state.turnOrderRedoStack = [];
    syncCurrentSnapshotFromState();
  } else {
    invalidateTurnState();
  }

  logEvent(`Charakter entfernt: ${character.name} (${character.type}).`);
  persistAppState();
  render();
}

function toggleSurprised(id, value) {
  state.characters = state.characters.map((character) => {
    if (character.id !== id) {
      return character;
    }

    return {
      ...character,
      surprised: value,
      totalInitiative: computeTotalInitiative(character, { surprised: value }),
    };
  });

  if (state.round > 0 && state.turnEntries.length) {
    rebuildTurnEntriesPreserveActive();
    state.turnOrderUndoStack = [];
    state.turnOrderRedoStack = [];
    syncCurrentSnapshotFromState();
  }

  const character = state.characters.find((item) => item.id === id);
  if (character) {
    logEvent(`Überraschung ${value ? "aktiviert" : "deaktiviert"}: ${character.name}.`);
  }
  persistAppState();
  render();
}

function toggleSurprisedByType(type, value) {
  const targetType = type === "NPC" ? "NPC" : "PC";
  let affectedCount = 0;

  state.characters = state.characters.map((character) => {
    if (character.type !== targetType) {
      return character;
    }

    affectedCount += 1;
    return {
      ...character,
      surprised: value,
      totalInitiative: computeTotalInitiative(character, { surprised: value }),
    };
  });

  if (affectedCount <= 0) {
    renderRoster();
    return;
  }

  if (state.round > 0 && state.turnEntries.length) {
    rebuildTurnEntriesPreserveActive();
    state.turnOrderUndoStack = [];
    state.turnOrderRedoStack = [];
    syncCurrentSnapshotFromState();
  }

  logEvent(`Überraschung ${value ? "aktiviert" : "deaktiviert"} für ${targetType === "PC" ? "alle SC" : "alle NSC"} (${affectedCount}).`);
  persistAppState();
  render();
}

function syncRosterSurpriseHeaderControls() {
  if (!surpriseAllPcEl || !surpriseAllNpcEl) {
    return;
  }

  const pcCharacters = state.characters.filter((character) => character.type === "PC");
  const npcCharacters = state.characters.filter((character) => character.type === "NPC");
  const surprisedPcCount = pcCharacters.filter((character) => character.surprised).length;
  const surprisedNpcCount = npcCharacters.filter((character) => character.surprised).length;

  surpriseAllPcEl.disabled = pcCharacters.length === 0;
  surpriseAllPcEl.indeterminate = surprisedPcCount > 0 && surprisedPcCount < pcCharacters.length;
  surpriseAllPcEl.checked = pcCharacters.length > 0 && surprisedPcCount === pcCharacters.length;

  surpriseAllNpcEl.disabled = npcCharacters.length === 0;
  surpriseAllNpcEl.indeterminate = surprisedNpcCount > 0 && surprisedNpcCount < npcCharacters.length;
  surpriseAllNpcEl.checked = npcCharacters.length > 0 && surprisedNpcCount === npcCharacters.length;
}

function toggleIncapacitated(id, value) {
  pushCharacterUndoSnapshot("Aktionsunfähig-Status zurückgesetzt.");
  state.characters = state.characters.map((character) =>
    character.id === id
      ? {
          ...character,
          incapacitated: value,
          lastRoll: value ? null : character.lastRoll,
          critBonusRoll: value ? null : character.critBonusRoll ?? null,
          totalInitiative: value ? null : character.totalInitiative,
          lastPixelsFaces: value ? null : character.lastPixelsFaces ?? null,
        }
      : character
  );

  if (state.turnEntries.length) {
    rebuildTurnEntriesPreserveActive();
    state.turnOrderUndoStack = [];
    state.turnOrderRedoStack = [];
    syncCurrentSnapshotFromState();
  }

  const character = state.characters.find((item) => item.id === id);
  if (character) {
    logEvent(`Aktionsunfähig ${value ? "aktiviert" : "deaktiviert"}: ${character.name}.`);
  }
  persistAppState();
  render();
}

function triggerParade(id) {
  const character = state.characters.find((item) => item.id === id);
  if (!character) {
    return;
  }

  const bonusEntry = state.turnEntries.find(
    (entry) => entry.characterId === id && entry.turn === "Bonus" && !isTurnEntryUsed(entry.id)
  );
  if (bonusEntry) {
    const activeCharacterId = getActiveCharacterId();
    pushTurnOrderUndoSnapshot("Parade rückgängig.");
    const changed = setTurnEntryUsed(bonusEntry.id, true);
    if (!changed) {
      return;
    }
    state.characters = state.characters.map((item) =>
      item.id === id
        ? {
            ...item,
            paradeClickCount: Math.max(0, Math.round(Number(item.paradeClickCount) || 0)) + 1,
          }
        : item
    );

    if (!hasUnresolvedActionForCharacter(activeCharacterId)) {
      const nextCharacterId = getNextCharacterIdForTurnOrder();
      if (nextCharacterId !== null) {
        state.activeCharacterId = nextCharacterId;
      }
    }

    syncActiveTurnPointer();
    syncCurrentSnapshotFromState();
    logEvent(`Parade: ${character.name} (Bonusaktion verbraucht).`);
    persistAppState();
    render();
    return;
  }

  const actionEntry = state.turnEntries.find(
    (entry) => entry.characterId === id && entry.turn === "Main" && !isTurnEntryUsed(entry.id)
  );
  if (actionEntry) {
    const activeCharacterId = getActiveCharacterId();
    pushTurnOrderUndoSnapshot("Parade rückgängig.");
    const changed = setTurnEntryUsed(actionEntry.id, true);
    if (!changed) {
      return;
    }

    state.characters = state.characters.map((item) =>
      item.id === id
        ? {
            ...item,
            paradeClickCount: Math.max(0, Math.round(Number(item.paradeClickCount) || 0)) + 1,
            unfreeDefensePenalty: Math.max(0, Math.round(Number(item.unfreeDefensePenalty) || 0)) + 6,
          }
        : item
    );

    if (!hasUnresolvedActionForCharacter(activeCharacterId)) {
      const nextCharacterId = getNextCharacterIdForTurnOrder();
      if (nextCharacterId !== null) {
        state.activeCharacterId = nextCharacterId;
      }
    }

    syncActiveTurnPointer();
    syncCurrentSnapshotFromState();
    const updated = state.characters.find((item) => item.id === id);
    const penalty = Math.max(0, Math.round(Number(updated?.unfreeDefensePenalty) || 0));
    logEvent(`Parade: ${character.name} (Aktion verbraucht, Nächste Parade -${penalty}).`);
    persistAppState();
    render();
    return;
  }

  pushTurnOrderUndoSnapshot("Parade rückgängig.");
  state.characters = state.characters.map((item) =>
    item.id === id
      ? {
          ...item,
          paradeClickCount: Math.max(0, Math.round(Number(item.paradeClickCount) || 0)) + 1,
          unfreeDefensePenalty: Math.max(0, Math.round(Number(item.unfreeDefensePenalty) || 0)) + 6,
        }
      : item
  );
  syncCurrentSnapshotFromState();
  const updated = state.characters.find((item) => item.id === id);
  if (updated) {
    logEvent(`Parade: ${updated.name} hat Nächste Parade -${updated.unfreeDefensePenalty}.`);
  }
  persistAppState();
  render();
}

function setManualRoll(id, rawValue) {
  const hasValue = String(rawValue).trim() !== "";
  const parsedValue = Number(rawValue);
  const manualRoll = !hasValue || Number.isNaN(parsedValue) ? null : clamp(Math.round(parsedValue), 3, 18);

  state.characters = state.characters.map((character) =>
    character.id === id
      ? {
          ...character,
          manualRoll,
          useManualRoll: character.type === "PC" ? true : character.useManualRoll,
        }
      : character
  );

  invalidateTurnState();
  const character = state.characters.find((item) => item.id === id);
  if (character) {
    logEvent(`Manueller 3w6 ${manualRoll === null ? "gelöscht" : `gesetzt auf ${manualRoll}`}: ${character.name}.`);
  }
  persistAppState();
  render();
}

function setDamageMonitorMark(id, columnIndex, checked) {
  if (!Number.isInteger(columnIndex) || columnIndex < 1 || columnIndex >= DAMAGE_MONITOR_COLUMN_COUNT) {
    return;
  }

  state.characters = state.characters.map((character) => {
    if (character.id !== id) {
      return character;
    }

    const nextMarks = normalizeDamageMonitorMarks(character.damageMonitorMarks);
    nextMarks[columnIndex] = Boolean(checked);
    return {
      ...character,
      damageMonitorMarks: nextMarks,
    };
  });

  persistAppState();
  render();
}

function updateEditDialogFieldStates() {
  if (!editTypeEl || !editUseManualRollEl) {
    return;
  }

  const isPc = editTypeEl.value === "PC";
  if (!isPc) {
    editUseManualRollEl.checked = false;
  }

  editUseManualRollEl.disabled = !isPc;
}

function openCharacterEditDialog(id) {
  if (!editCharacterDialogEl || !editCharacterFormEl) {
    return;
  }
  const character = state.characters.find((item) => item.id === id);
  if (!character) {
    return;
  }

  editingCharacterId = id;
  if (editCharacterTitleEl) {
    editCharacterTitleEl.textContent = `Charakter bearbeiten: ${character.name}`;
  }
  if (editNameEl) editNameEl.value = character.name;
  if (editTypeEl) editTypeEl.value = character.type === "NPC" ? "NPC" : "PC";
  if (editIniEl) editIniEl.value = String(character.ini);
  if (editUseManualRollEl) {
    editUseManualRollEl.checked =
      character.type === "PC" && (Boolean(character.useManualRoll) || character.manualRoll !== null);
  }
  if (editSpecialAbilityEl) editSpecialAbilityEl.value = character.specialAbility ?? "";
  if (editAlwaysMinimizedEl) editAlwaysMinimizedEl.checked = Boolean(character.alwaysMinimized);
  if (editRollInitiativeBtn) {
    const isInTurnOrder = isCharacterInCurrentTurnOrder(character.id);
    const canRollIntoRound = state.round > 0 && state.turnEntries.length > 0 && !character.incapacitated && !isInTurnOrder;
    editRollInitiativeBtn.hidden = !canRollIntoRound;
    editRollInitiativeBtn.disabled = !canRollIntoRound;
  }

  updateEditDialogFieldStates();
  if (typeof editCharacterDialogEl.showModal === "function") {
    editCharacterDialogEl.showModal();
  } else {
    editCharacterDialogEl.setAttribute("open", "open");
  }
}

function saveCharacterEdits() {
  if (editingCharacterId === null || !editNameEl || !editTypeEl || !editIniEl) {
    return;
  }

  const existingCharacter = state.characters.find((item) => item.id === editingCharacterId);
  if (!existingCharacter) {
    editingCharacterId = null;
    editCharacterDialogEl?.close();
    return;
  }

  const name = editNameEl.value.trim();
  if (!name) {
    return;
  }

  const type = editTypeEl.value === "NPC" ? "NPC" : "PC";
  const ini = clamp(Number(editIniEl.value), 1, 30);
  const specialAbility = editSpecialAbilityEl?.value ? editSpecialAbilityEl.value : null;
  const useManualRoll = type === "PC" && Boolean(editUseManualRollEl?.checked);
  const manualRoll = useManualRoll ? existingCharacter.manualRoll : null;
  const alwaysMinimized = Boolean(editAlwaysMinimizedEl?.checked);
  const minimized = alwaysMinimized ? true : existingCharacter.minimized;

  pushCharacterUndoSnapshot("Charakter-Bearbeitung zurückgesetzt.");
  state.characters = state.characters.map((character) => {
    if (character.id !== editingCharacterId) {
      return character;
    }
    return {
      ...character,
      name,
      type,
      ini,
      manualRoll,
      useManualRoll,
      specialAbility,
      minimized,
      alwaysMinimized,
      manualExpandedUntilNextTurn: alwaysMinimized ? false : character.manualExpandedUntilNextTurn ?? false,
    };
  });

  if (!specialAbility) {
    state.activeAbilityCharacterIds = state.activeAbilityCharacterIds.filter((id) => id !== editingCharacterId);
  }

  const editedCharacter = state.characters.find((item) => item.id === editingCharacterId) || null;
  if (editedCharacter && editedCharacter.lastRoll !== null) {
    editedCharacter.totalInitiative = computeTotalInitiative(editedCharacter);
  }

  if (state.turnEntries.length) {
    rebuildTurnEntriesPreserveActive();
    state.turnOrderUndoStack = [];
    state.turnOrderRedoStack = [];
    syncCurrentSnapshotFromState();
  }

  logEvent(`Charakter bearbeitet: ${name} (${getTypeLabel(type)}).`);
  persistAppState();
  render();

  editingCharacterId = null;
  editCharacterDialogEl?.close();
}

function isCharacterInCurrentTurnOrder(id) {
  return state.turnEntries.some((entry) => entry.characterId === id);
}

async function rollCharacterIntoCurrentRound(id) {
  const character = state.characters.find((item) => item.id === id);
  if (!character) {
    return;
  }
  if (!state.turnEntries.length || state.round <= 0) {
    return;
  }
  if (character.incapacitated || isCharacterInCurrentTurnOrder(id)) {
    return;
  }

  pushTurnOrderUndoSnapshot("INI-Nachwurf rückgängig.");
  const usePixelsForTurn = canUsePixelsForRolls();
  const shouldShowPixelsRollDialog = preparePixelsRollDialog([character], usePixelsForTurn, "INI Nachwurf");
  let rollData;
  let totalInitiative = null;
  try {
    if (shouldShowPixelsRollDialog) {
      setPixelsRollDialogStatus(`Warte auf den Nachwurf für ${character.name}.`);
    }
    const rollMode = getCharacterRollMode(character, usePixelsForTurn);
    rollData = await resolveCharacterRoll(character, usePixelsForTurn);
    totalInitiative = computeTotalInitiative(character, {
      lastRoll: rollData.total,
      critBonusRoll: rollData.critBonusRoll,
    });
    if (shouldShowPixelsRollDialog) {
      updatePixelsRollDialogRowResult(character, rollData, totalInitiative, rollMode);
      setPixelsRollDialogStatus(`Nachwurf für ${character.name} abgeschlossen.`);
    }
  } finally {
    if (shouldShowPixelsRollDialog) {
      closePixelsRollDialog();
    }
  }

  state.characters = state.characters.map((item) =>
    item.id === id
      ? {
          ...item,
          lastRoll: rollData.total,
          critBonusRoll: rollData.critBonusRoll,
          manualRoll:
            rollData.source === "manual" && Number.isFinite(Number(rollData.manualRollValue))
              ? clamp(Number(rollData.manualRollValue), 3, 18)
              : item.manualRoll,
          totalInitiative,
          lastPixelsFaces: Array.isArray(rollData.faces) ? rollData.faces.map((face) => Math.round(Number(face) || 0)) : null,
          unfreeDefensePenalty: 0,
        }
      : item
  );

  rebuildTurnEntriesPreserveActive();
  syncCurrentSnapshotFromState();
  const critBonusRollText =
    Number.isFinite(Number(rollData.critBonusRoll)) && Number(rollData.critBonusRoll) > 0
      ? `, Krit-W6=${rollData.critBonusRoll}`
      : "";
  logEvent(`INI nachgewürfelt: ${character.name} (3w6=${rollData.total}${critBonusRollText}, ges. ${totalInitiative}).`);
  persistAppState();
  render();
  editCharacterDialogEl?.close();
}

function toggleSpecialAbilityForRound(id) {
  const character = state.characters.find((item) => item.id === id);
  if (!character || !character.specialAbility || !state.turnEntries.length) {
    return;
  }

  pushTurnOrderUndoSnapshot("Status Ch. Vorteil wiederhergestellt.");
  if (state.activeAbilityCharacterIds.includes(id)) {
    state.activeAbilityCharacterIds = state.activeAbilityCharacterIds.filter((characterId) => characterId !== id);
    logEvent(
      `Ch. Vorteil in dieser KR deaktiviert: ${character.name} (${getSpecialAbilityLabel(character.specialAbility)}).`
    );
  } else {
    state.activeAbilityCharacterIds.push(id);
    logEvent(
      `Ch. Vorteil in dieser KR aktiviert: ${character.name} (${getSpecialAbilityLabel(character.specialAbility)}) [+3 INI].`
    );
  }

  rebuildTurnEntriesPreserveActive();
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function toggleMinimized(id) {
  state.characters = state.characters.map((character) => {
    if (character.id !== id) {
      return character;
    }

    const nextMinimized = !character.minimized;
    if (character.alwaysMinimized) {
      return {
        ...character,
        minimized: nextMinimized,
        manualExpandedUntilNextTurn: nextMinimized ? false : true,
      };
    }

    return {
      ...character,
      minimized: nextMinimized,
      manualExpandedUntilNextTurn: false,
    };
  });

  persistAppState();
  renderRoster();
}

function applyDazedStatus(id) {
  if (state.dazedAppliedRound !== state.round) {
    state.dazedAppliedCharacterIds = [];
    state.dazedAppliedRound = state.round;
  }

  if (state.dazedAppliedCharacterIds.includes(id)) {
    return;
  }

  pushTurnOrderUndoSnapshot("Benommenheits-Änderung rückgängig.");
  const defaultUntilRound = getDefaultDazedUntilRound();

  state.characters = state.characters.map((character) => {
    if (character.id !== id) {
      return character;
    }

    const currentlyDazed = isCharacterDazed(character);
    const currentUntil = character.dazedUntilRound ?? defaultUntilRound;

    return {
      ...character,
      dazedUntilRound: currentlyDazed ? currentUntil + 1 : defaultUntilRound,
    };
  });

  state.turnEntries = buildSortedTrackerEntries();
  state.dazedAppliedCharacterIds.push(id);
  syncActiveTurnPointer();
  const character = state.characters.find((item) => item.id === id);
  if (character) {
    logEvent(`Benommenheit gesetzt für ${character.name} bis Ende KR ${character.dazedUntilRound}.`);
  }
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function rebuildTurnEntriesPreserveActive() {
  state.turnEntries = buildSortedTrackerEntries();
  const validEntryIds = new Set(state.turnEntries.map((entry) => entry.id));
  state.usedTurnEntryIds = state.usedTurnEntryIds.filter((id) => validEntryIds.has(id));

  if (!state.turnEntries.length) {
    state.currentTurnIndex = -1;
    state.activeCharacterId = null;
    return;
  }

  syncActiveTurnPointer();
}

function getSpecialAbilityBonus(characterId) {
  return state.activeAbilityCharacterIds.includes(characterId) ? 3 : 0;
}

function buildSortedTrackerEntries() {
  const entries = [];

  for (const character of state.characters) {
    if (character.totalInitiative === null) {
      continue;
    }

    const totalWithAbility = character.totalInitiative + getSpecialAbilityBonus(character.id);
    const hasBonusAction = totalWithAbility > 30;

    if (hasBonusAction) {
      entries.push({
        id: `${character.id}-bonus`,
        characterId: character.id,
        name: character.name,
        type: character.type,
        turn: "Bonus",
        initiative: totalWithAbility,
        groupInitiative: totalWithAbility,
        actionIndex: 0,
        critical: character.lastRoll === 3 ? "failure" : character.lastRoll === 18 ? "success" : null,
        dazed: isCharacterDazed(character),
        specialAbility: character.specialAbility || null,
        abilityActive: getSpecialAbilityBonus(character.id) > 0,
      });
    }

    entries.push({
      id: `${character.id}-main`,
      characterId: character.id,
      name: character.name,
      type: character.type,
      turn: "Main",
      initiative: totalWithAbility,
      groupInitiative: totalWithAbility,
      actionIndex: hasBonusAction ? 1 : 0,
      critical: character.lastRoll === 3 ? "failure" : character.lastRoll === 18 ? "success" : null,
      dazed: isCharacterDazed(character),
      specialAbility: character.specialAbility || null,
      abilityActive: getSpecialAbilityBonus(character.id) > 0,
    });

    entries.push({
      id: `${character.id}-move`,
      characterId: character.id,
      name: character.name,
      type: character.type,
      turn: "Move",
      initiative: totalWithAbility,
      groupInitiative: totalWithAbility,
      actionIndex: hasBonusAction ? 2 : 1,
      critical: character.lastRoll === 3 ? "failure" : character.lastRoll === 18 ? "success" : null,
      dazed: isCharacterDazed(character),
      specialAbility: character.specialAbility || null,
      abilityActive: getSpecialAbilityBonus(character.id) > 0,
    });
  }

  entries.sort((a, b) => {
    if (b.groupInitiative !== a.groupInitiative) {
      return b.groupInitiative - a.groupInitiative;
    }

    if (a.characterId !== b.characterId) {
      return a.name.localeCompare(b.name);
    }

    if (a.actionIndex !== b.actionIndex) {
      return a.actionIndex - b.actionIndex;
    }

    if (b.initiative !== a.initiative) {
      return b.initiative - a.initiative;
    }

    return a.characterId - b.characterId;
  });

  return entries;
}

function moveTurnGroup(groupIndex, direction) {
  const groups = buildTurnDisplayGroups();
  if (!groups.length || groupIndex < 0 || groupIndex >= groups.length) {
    return;
  }

  const targetGroupIndex = groupIndex + direction;
  if (targetGroupIndex < 0 || targetGroupIndex >= groups.length) {
    return;
  }

  pushTurnOrderUndoSnapshot("INI-Reihenfolge-Verschiebung rückgängig.");
  const currentGroup = groups[groupIndex];
  const targetGroup = groups[targetGroupIndex];
  const movingEntries = state.turnEntries.slice(currentGroup.startIndex, currentGroup.endIndex + 1);

  state.turnEntries.splice(currentGroup.startIndex, currentGroup.entries.length);

  let insertAt = targetGroup.startIndex;
  if (direction > 0) {
    insertAt = targetGroup.endIndex - currentGroup.entries.length + 1;
  }

  state.turnEntries.splice(insertAt, 0, ...movingEntries);
  syncActiveTurnPointer();

  logEvent(`INI-Reihenfolge verschoben: ${currentGroup.name} ${direction < 0 ? "hoch" : "runter"}.`);
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function renderRoster() {
  rosterEl.innerHTML = "";
  syncRosterSurpriseHeaderControls();

  if (!state.characters.length) {
    rosterEl.innerHTML = '<p class="empty">Noch keine Charaktere.</p>';
    return;
  }

  const activeCharacterId = getActiveCharacterId();

  for (const character of state.characters) {
    const characterIndex = state.characters.findIndex((item) => item.id === character.id);
    const fragment = characterTemplate.content.cloneNode(true);
    const card = fragment.querySelector(".character-card");
    const nameEl = fragment.querySelector(".character-name");
    const typeBadgeEl = fragment.querySelector(".type-badge");
    const metaEl = fragment.querySelector(".character-meta");
    const surpriseInput = fragment.querySelector(".surprise-input");
    const incapacitatedInput = fragment.querySelector(".incapacitated-input");
    const manualRollField = fragment.querySelector(".manual-roll");
    const manualRollInput = fragment.querySelector(".manual-roll-input");
    const woundInputs = Array.from(fragment.querySelectorAll(".damage-wound-input"));
    const editBtn = fragment.querySelector(".edit-btn");
    const minimizeBtn = fragment.querySelector(".minimize-btn");
    const removeBtn = fragment.querySelector(".remove-btn");

    card.classList.add(character.type === "PC" ? "type-pc" : "type-npc");
    if (activeCharacterId !== null && character.id === activeCharacterId) {
      card.classList.add("active");
    }
    if (character.minimized) {
      card.classList.add("minimized");
    }
    nameEl.textContent = character.name;
    typeBadgeEl.textContent = getTypeLabel(character.type);

    const rollText =
      character.lastRoll === null
        ? "nicht gewürfelt"
        : `3w6=${character.lastRoll}, ges.=${character.totalInitiative}`;
    const manualText =
      character.type === "PC" && character.useManualRoll && character.manualRoll !== null
        ? `, manuell=${character.manualRoll}`
        : "";
    const critBonusRollText =
      Number.isFinite(Number(character.critBonusRoll)) && Number(character.critBonusRoll) > 0
        ? `, Krit-W6=${character.critBonusRoll}`
        : "";
    const pixelsFacesText = Array.isArray(character.lastPixelsFaces) && character.lastPixelsFaces.length
      ? `, Pixels [${character.lastPixelsFaces.join(", ")}]`
      : "";
    const specialAbilityText = character.specialAbility ? `, Ch. Vorteil=${getSpecialAbilityLabel(character.specialAbility)}` : "";
    const dazedText = isCharacterDazed(character) ? `, Ben. bis Ende KR ${character.dazedUntilRound}` : "";
    const incapacitatedText = character.incapacitated ? ", Aktionsunfähig" : "";
    const baseMetaText = `INI ${character.ini} | ${rollText}${manualText}${critBonusRollText}${pixelsFacesText}${specialAbilityText}${dazedText}${incapacitatedText}`;
    const damagePenalty = computeDamagePenalty(character);
    const dazedPenalty = isCharacterDazed(character) ? 3 : 0;
    const effectiveQmPenalty = damagePenalty.qm + dazedPenalty;
    const shouldShowDamagePenalty =
      dazedPenalty > 0 ||
      normalizeDamageMonitorMarks(character.damageMonitorMarks).some((active, index) => (index > 0 ? active : false));
    const unfreeDefensePenalty = Math.max(0, Math.round(Number(character.unfreeDefensePenalty) || 0));
    const penaltyTexts = [];
    if (shouldShowDamagePenalty) {
      penaltyTexts.push(`QM/BEW: -${effectiveQmPenalty}/-${damagePenalty.bew}`);
    }
    if (unfreeDefensePenalty > 0) {
      penaltyTexts.push(`Nächste Parade -${unfreeDefensePenalty}`);
    }

    if (penaltyTexts.length) {
      metaEl.innerHTML = `${escapeHtml(baseMetaText)} | <span class="character-meta-penalty">${escapeHtml(
        penaltyTexts.join(" | ")
      )}</span>`;
    } else {
      metaEl.textContent = baseMetaText;
    }

    surpriseInput.checked = character.surprised;
    surpriseInput.addEventListener("change", (event) => {
      toggleSurprised(character.id, event.target.checked);
    });
    incapacitatedInput.checked = Boolean(character.incapacitated);
    incapacitatedInput.addEventListener("change", (event) => {
      toggleIncapacitated(character.id, event.target.checked);
    });

    manualRollField.hidden = character.type !== "PC" || !character.useManualRoll;
    manualRollInput.disabled = character.type !== "PC" || !character.useManualRoll;

    if (character.type === "PC" && character.useManualRoll) {
      manualRollField.hidden = false;
      manualRollInput.value = character.manualRoll ?? "";
      manualRollInput.addEventListener("change", (event) => {
        setManualRoll(character.id, event.target.value);
      });
    } else {
      manualRollInput.value = "";
    }

    const damageMarks = normalizeDamageMonitorMarks(character.damageMonitorMarks);
    for (const woundInput of woundInputs) {
      const columnIndex = Number(woundInput.dataset.damageIndex);
      if (!Number.isInteger(columnIndex) || columnIndex < 1 || columnIndex >= DAMAGE_MONITOR_COLUMN_COUNT) {
        woundInput.disabled = true;
        continue;
      }
      woundInput.checked = Boolean(damageMarks[columnIndex]);
      woundInput.addEventListener("change", (event) => {
        setDamageMonitorMark(character.id, columnIndex, event.target.checked);
      });
    }

    editBtn.addEventListener("click", () => openCharacterEditDialog(character.id));

    minimizeBtn.textContent = character.minimized ? "+" : "-";
    minimizeBtn.title = character.minimized ? "Erweitern" : "Minimieren";
    minimizeBtn.addEventListener("click", () => toggleMinimized(character.id));

    removeBtn.addEventListener("click", () => removeCharacter(character.id));

    card.dataset.characterId = String(character.id);
    card.draggable = true;
    card.addEventListener("dragstart", (event) => {
      draggedRosterCharacterId = character.id;
      card.classList.add("dragging");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", String(character.id));
      }
    });
    card.addEventListener("dragend", () => {
      draggedRosterCharacterId = null;
      card.classList.remove("dragging");
    });
    card.addEventListener("dragover", (event) => {
      if (draggedRosterCharacterId === null || draggedRosterCharacterId === character.id) {
        return;
      }
      event.preventDefault();
      card.classList.add("drag-target");
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "move";
      }
    });
    card.addEventListener("dragleave", () => {
      card.classList.remove("drag-target");
    });
    card.addEventListener("drop", (event) => {
      card.classList.remove("drag-target");
      if (draggedRosterCharacterId === null || draggedRosterCharacterId === character.id) {
        return;
      }
      event.preventDefault();
      const sourceIndex = state.characters.findIndex((item) => item.id === draggedRosterCharacterId);
      const targetIndex = characterIndex;
      moveRosterCharacter(sourceIndex, targetIndex);
    });
    rosterEl.appendChild(fragment);
  }
}

function renderTracker() {
  trackerEl.innerHTML = "";
  nextTurnBtn.textContent = state.round > 0 ? "Nächste KR" : "Starte Kampf";

  if (!state.turnEntries.length) {
    activateCharacterUpBtn.disabled = true;
    activateCharacterDownBtn.disabled = true;
    roundStatusEl.textContent = "Kampfrunde: -";
    trackerEl.innerHTML = "";
    return;
  }

  syncActiveTurnPointer();
  const activeCharacterId = getActiveCharacterId();
  const orderIndexByCharacter = getTurnCharacterOrderIndexMap();
  const activeOrderIndex =
    activeCharacterId !== null && orderIndexByCharacter.has(activeCharacterId)
      ? orderIndexByCharacter.get(activeCharacterId)
      : -1;
  activateCharacterUpBtn.disabled = activeOrderIndex <= 0;
  activateCharacterDownBtn.disabled =
    activeOrderIndex < 0 || activeOrderIndex >= orderIndexByCharacter.size - 1;
  roundStatusEl.textContent = `Kampfrunde: ${state.round}`;

  const displayGroups = buildTurnDisplayGroups();
  displayGroups.forEach((group, groupIndex) => {
    const item = document.createElement("li");
    item.className = `tracker-item ${group.type === "PC" ? "type-pc" : "type-npc"}`;

    const allUsed = group.entries.every((entry) => isTurnEntryUsed(entry.id));
    const hasActive = activeCharacterId !== null && group.characterId === activeCharacterId;

    if (allUsed) {
      item.classList.add("used");
    }

    if (hasActive) {
      item.classList.add("active");
    }
    item.draggable = true;
    item.addEventListener("dragstart", (event) => {
      draggedTurnGroupIndex = groupIndex;
      item.classList.add("dragging");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", String(groupIndex));
      }
    });
    item.addEventListener("dragend", () => {
      draggedTurnGroupIndex = null;
      item.classList.remove("dragging");
    });
    item.addEventListener("dragover", (event) => {
      if (draggedTurnGroupIndex === null || draggedTurnGroupIndex === groupIndex) {
        return;
      }
      event.preventDefault();
      item.classList.add("drag-target");
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "move";
      }
    });
    item.addEventListener("dragleave", () => {
      item.classList.remove("drag-target");
    });
    item.addEventListener("drop", (event) => {
      item.classList.remove("drag-target");
      if (draggedTurnGroupIndex === null || draggedTurnGroupIndex === groupIndex) {
        return;
      }
      event.preventDefault();
      moveTurnGroupToIndex(draggedTurnGroupIndex, groupIndex);
    });

    const topRow = document.createElement("div");
    topRow.className = "tracker-row tracker-row-top";
    const topMain = document.createElement("div");
    topMain.className = "tracker-top-main";
    const initiativeBadge = document.createElement("span");
    initiativeBadge.className = "type-badge";
    initiativeBadge.textContent = String(group.groupInitiative);
    const typeBadge = document.createElement("span");
    typeBadge.className = "type-badge";
    typeBadge.textContent = getTypeLabel(group.type);
    const nameStrong = document.createElement("strong");
    nameStrong.textContent = group.name;
    nameStrong.className = "tracker-name";
    const bottomRow = document.createElement("div");
    bottomRow.className = "tracker-row tracker-row-bottom";
    const actionsLeft = document.createElement("div");
    actionsLeft.className = "tracker-actions-left";
    const actionChipWrap = document.createElement("span");
    actionChipWrap.className = "action-chip-wrap";

    const turnOrder = ["Bonus", "Main", "Move"];
    for (const turnType of turnOrder) {
      const entry = group.entries.find((candidate) => candidate.turn === turnType);
      if (!entry) {
        continue;
      }
      const chipClass = isTurnEntryUsed(entry.id)
        ? "used"
        : state.currentTurnIndex === entry.entryIndex
          ? "active"
          : "pending";
      const chipLabel =
        entry.turn === "Main"
          ? "Aktion"
          : entry.turn === "Move"
            ? "Bew."
            : "Bonus";
      const chipBtn = document.createElement("button");
      chipBtn.type = "button";
      chipBtn.className = `action-chip ${chipClass}`;
      chipBtn.textContent = chipLabel;
      chipBtn.disabled = !canInteractWithTurnEntry(entry, orderIndexByCharacter);
      chipBtn.addEventListener("click", () => toggleTurnEntryUsed(entry.id));
      actionChipWrap.appendChild(chipBtn);

      if (entry.turn === "Main") {
        const paradeClickCount = Math.max(
          0,
          Math.round(Number(state.characters.find((item) => item.id === group.characterId)?.paradeClickCount) || 0)
        );
        const paradeBtn = document.createElement("button");
        paradeBtn.type = "button";
        paradeBtn.className = "action-chip parade-action-btn";
        paradeBtn.textContent = `Parade (${paradeClickCount})`;
        paradeBtn.disabled = false;
        paradeBtn.addEventListener("click", () => triggerParade(group.characterId));
        actionChipWrap.appendChild(paradeBtn);
      }
    }

    topMain.appendChild(initiativeBadge);
    topMain.appendChild(typeBadge);
    topMain.appendChild(nameStrong);
    topRow.appendChild(topMain);
    actionsLeft.appendChild(actionChipWrap);
    bottomRow.appendChild(actionsLeft);

    const hintCharacter = state.characters.find((character) => character.id === group.characterId) || null;
    if (hintCharacter) {
      const damagePenalty = computeDamagePenalty(hintCharacter);
      const dazedPenalty = isCharacterDazed(hintCharacter) ? 3 : 0;
      const effectiveQmPenalty = damagePenalty.qm + dazedPenalty;
      const shouldShowDamagePenalty =
        dazedPenalty > 0 ||
        normalizeDamageMonitorMarks(hintCharacter.damageMonitorMarks).some((active, index) => (index > 0 ? active : false));
      const unfreeDefensePenalty = Math.max(0, Math.round(Number(hintCharacter.unfreeDefensePenalty) || 0));
      const hintTexts = [];
      if (shouldShowDamagePenalty) {
        hintTexts.push(`QM/BEW -${effectiveQmPenalty}/-${damagePenalty.bew}`);
      }
      if (unfreeDefensePenalty > 0) {
        hintTexts.push(`Nächste Parade -${unfreeDefensePenalty}`);
      }
      if (hintTexts.length) {
        const topHints = document.createElement("span");
        topHints.className = "tracker-top-hints";
        topHints.textContent = hintTexts.join(" | ");
        topRow.appendChild(topHints);
      }
    }

    if (group.critical === "failure" || group.critical === "success") {
      const criticalLabel = document.createElement("span");
      criticalLabel.className = `crit-tag ${group.critical === "failure" ? "crit-fail" : "crit-success"}`;
      criticalLabel.textContent = group.critical === "failure" ? "Krit. Fehlschlag" : "Krit. Erfolg";
      topRow.appendChild(criticalLabel);
    }

    const controls = document.createElement("div");
    controls.className = "shift-controls tracker-actions-right";

    const dazedBtn = document.createElement("button");
    dazedBtn.type = "button";
    dazedBtn.className = "shift-btn status-action-btn";
    dazedBtn.textContent = "Ben.";
    dazedBtn.disabled = state.dazedAppliedRound === state.round && state.dazedAppliedCharacterIds.includes(group.characterId);
    dazedBtn.addEventListener("click", () => applyDazedStatus(group.characterId));

    if (group.specialAbility) {
      const abilityBtn = document.createElement("button");
      abilityBtn.type = "button";
      abilityBtn.className = `ghost shift-btn ability-action-btn${group.abilityActive ? " active" : ""}`;
      abilityBtn.textContent = group.abilityActive ? `${getSpecialAbilityLabel(group.specialAbility)} (+3)` : getSpecialAbilityLabel(group.specialAbility);
      abilityBtn.addEventListener("click", () => toggleSpecialAbilityForRound(group.characterId));
      controls.appendChild(abilityBtn);
    }

    controls.appendChild(dazedBtn);

    bottomRow.appendChild(controls);
    item.appendChild(topRow);
    item.appendChild(bottomRow);
    trackerEl.appendChild(item);
  });
}

function moveRosterCharacter(sourceIndex, targetIndex) {
  if (sourceIndex < 0 || targetIndex < 0 || sourceIndex === targetIndex) {
    return;
  }
  if (sourceIndex >= state.characters.length || targetIndex >= state.characters.length) {
    return;
  }

  pushCharacterUndoSnapshot("Charakterreihenfolge zurückgesetzt.");
  const nextCharacters = [...state.characters];
  const [movedCharacter] = nextCharacters.splice(sourceIndex, 1);
  nextCharacters.splice(targetIndex, 0, movedCharacter);
  state.characters = nextCharacters;
  logEvent(`Charakterreihenfolge geändert: ${movedCharacter.name}.`);
  persistAppState();
  renderRoster();
}

function moveTurnGroupToIndex(sourceGroupIndex, targetGroupIndex) {
  if (!state.turnEntries.length || sourceGroupIndex === targetGroupIndex) {
    return;
  }
  const groups = buildTurnDisplayGroups();
  if (
    sourceGroupIndex < 0 ||
    targetGroupIndex < 0 ||
    sourceGroupIndex >= groups.length ||
    targetGroupIndex >= groups.length
  ) {
    return;
  }

  pushTurnOrderUndoSnapshot("INI-Reihenfolge-Verschiebung rückgängig.");
  const sourceGroup = groups[sourceGroupIndex];
  const targetGroup = groups[targetGroupIndex];
  const movingEntries = state.turnEntries.slice(sourceGroup.startIndex, sourceGroup.endIndex + 1);
  state.turnEntries.splice(sourceGroup.startIndex, sourceGroup.entries.length);

  const insertionBase = sourceGroupIndex < targetGroupIndex
    ? targetGroup.endIndex - sourceGroup.entries.length + 1
    : targetGroup.startIndex;
  const insertionIndex = Math.max(0, Math.min(insertionBase, state.turnEntries.length));
  state.turnEntries.splice(insertionIndex, 0, ...movingEntries);

  syncActiveTurnPointer();
  logEvent(`INI-Reihenfolge verschoben: ${sourceGroup.name}.`);
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function buildTurnDisplayGroups() {
  const groups = [];

  for (let index = 0; index < state.turnEntries.length; index += 1) {
    const entry = state.turnEntries[index];
    const previous = groups[groups.length - 1];

    if (
      previous &&
      previous.characterId === entry.characterId &&
      previous.groupInitiative === entry.groupInitiative &&
      previous.endIndex === index - 1
    ) {
      previous.entries.push({ ...entry, entryIndex: index });
      previous.endIndex = index;
      continue;
    }

    groups.push({
      characterId: entry.characterId,
      name: entry.name,
      type: entry.type,
      critical: entry.critical,
      dazed: entry.dazed,
      specialAbility: entry.specialAbility || null,
      abilityActive: Boolean(entry.abilityActive),
      groupInitiative: entry.groupInitiative,
      entries: [{ ...entry, entryIndex: index }],
      startIndex: index,
      endIndex: index,
    });
  }

  return groups;
}

function renderCombatLog() {
  combatLogEl.innerHTML = "";

  if (!state.combatLog.length) {
    combatLogEl.innerHTML = '<p class="empty">Noch keine Kampfeinträge.</p>';
    return;
  }

  const groupedEntries = groupCombatLogEntries(state.combatLog);
  for (const entry of groupedEntries) {
    const item = document.createElement("div");
    item.className = "log-entry";
    item.classList.add(isInitiativeLogMessage(entry.displayMessage) ? "log-entry-initiative" : "log-entry-muted");

    const lineNo = document.createElement("span");
    lineNo.className = "log-line-no";
    lineNo.textContent = entry.lineLabel;

    const turnCol = document.createElement("span");
    turnCol.className = "log-turn";
    const turnText = entry.turn > 0 ? `T${entry.turn}` : "T-";
    turnCol.textContent = turnText;

    const timeCol = document.createElement("span");
    timeCol.className = "log-time";
    timeCol.textContent = formatLogTimestamp(entry.timestamp);

    const message = document.createElement("span");
    message.className = "log-message";
    message.textContent = entry.displayMessage;

    item.appendChild(lineNo);
    item.appendChild(turnCol);
    item.appendChild(timeCol);
    item.appendChild(message);
    combatLogEl.appendChild(item);
  }

  combatLogEl.scrollTop = combatLogEl.scrollHeight;
}

function isUndoLogEntry(entry) {
  return typeof entry?.message === "string" && entry.message.startsWith("Rückg.:");
}

function groupCombatLogEntries(logEntries) {
  const grouped = [];

  for (let index = 0; index < logEntries.length; index += 1) {
    const current = logEntries[index];
    if (!isUndoLogEntry(current)) {
      grouped.push({
        id: current.id,
        lineLabel: `#${current.id}`,
        turn: current.turn,
        timestamp: current.timestamp,
        displayMessage: current.message,
      });
      continue;
    }

    const undoEntries = [current];
    while (index + 1 < logEntries.length && isUndoLogEntry(logEntries[index + 1])) {
      index += 1;
      undoEntries.push(logEntries[index]);
    }

    const first = undoEntries[0];
    const last = undoEntries[undoEntries.length - 1];
    const firstDetail = first.message.replace(/^Rückg\.\:\s*/i, "").trim();
    const displayMessage =
      undoEntries.length === 1
        ? first.message
        : `Rückg. x${undoEntries.length}: ${firstDetail} (+${undoEntries.length - 1} mehr).`;

    grouped.push({
      id: last.id,
      lineLabel: undoEntries.length === 1 ? `#${first.id}` : `#${first.id}-${last.id}`,
      turn: last.turn,
      timestamp: last.timestamp,
      displayMessage,
    });
  }

  return grouped;
}

function isInitiativeLogMessage(message) {
  return /^[^:]+:\s*3w6=\d+,\s*INI\s+\d+,\s*(ges\.|insgesamt)\s+\d+/i.test(message);
}

function logEvent(message) {
  state.combatLog.push({
    id: state.nextLogId++,
    turn: state.round,
    timestamp: Date.now(),
    message,
  });
}

function createCharacter(data) {
  const alwaysMinimized = Boolean(data.alwaysMinimized);
  const manualExpandedUntilNextTurn = Boolean(data.manualExpandedUntilNextTurn);
  const useManualRoll = Boolean(data.useManualRoll ?? (data.manualRoll !== null && data.manualRoll !== undefined));
  const providedId = Number.isFinite(Number(data.id)) ? Math.max(1, Math.round(Number(data.id))) : null;
  const characterId = providedId ?? state.nextId;
  state.nextId = Math.max(state.nextId, characterId + 1);
  return {
    id: characterId,
    name: String(data.name || "Unbenannt").trim() || "Unbenannt",
    type: data.type === "NPC" ? "NPC" : "PC",
    ini: clamp(Number(data.ini), 1, 30),
    surprised: Boolean(data.surprised),
    dazedUntilRound:
      data.dazedUntilRound === null || data.dazedUntilRound === undefined
        ? null
        : Math.max(1, Math.round(Number(data.dazedUntilRound))),
    minimized: alwaysMinimized && !manualExpandedUntilNextTurn ? true : Boolean(data.minimized),
    alwaysMinimized,
    manualExpandedUntilNextTurn,
    useManualRoll,
    manualRoll:
      data.manualRoll === null || data.manualRoll === undefined ? null : clamp(Number(data.manualRoll), 3, 18),
    lastRoll: data.lastRoll === null || data.lastRoll === undefined ? null : clamp(Number(data.lastRoll), 3, 18),
    critBonusRoll:
      data.critBonusRoll === null || data.critBonusRoll === undefined ? null : clamp(Number(data.critBonusRoll), 1, 6),
    lastPixelsFaces: Array.isArray(data.lastPixelsFaces)
      ? data.lastPixelsFaces.map((face) => Math.round(Number(face) || 0)).filter((face) => face >= 1 && face <= 6)
      : null,
    totalInitiative:
      data.totalInitiative === null || data.totalInitiative === undefined
        ? null
        : Math.round(Number(data.totalInitiative)),
    specialAbility: typeof data.specialAbility === "string" && data.specialAbility ? data.specialAbility : null,
    damageMonitorMarks: normalizeDamageMonitorMarks(data.damageMonitorMarks),
    incapacitated: Boolean(data.incapacitated),
    unfreeDefensePenalty: Math.max(0, Math.round(Number(data.unfreeDefensePenalty) || 0)),
    paradeClickCount: Math.max(0, Math.round(Number(data.paradeClickCount) || 0)),
  };
}

function computeTotalInitiative(character, options = {}) {
  const ini = clamp(Number(options.ini ?? character.ini), 1, 30);
  const lastRoll = options.lastRoll ?? character.lastRoll;
  if (lastRoll === null || lastRoll === undefined) {
    return null;
  }

  const critBonusRoll = options.critBonusRoll ?? character.critBonusRoll ?? 0;
  const surprised = options.surprised ?? character.surprised;
  const surprisePenalty = surprised ? -10 : 0;
  return lastRoll + ini + critBonusRoll + surprisePenalty;
}

function normalizeRememberedPixelRef(value) {
  if (!value || typeof value !== "object") {
    return null;
  }

  const deviceId = String(value.systemId || value.deviceId || "").trim();
  if (!deviceId) {
    return null;
  }

  const label = String(value.label || "").trim();
  return {
    deviceId,
    systemId: deviceId,
    label: label || "Verbunden",
  };
}

function normalizeRememberedAssignment(value) {
  const assignment = value && typeof value === "object" ? value : {};
  return {
    single: normalizeRememberedPixelRef(assignment.single),
    set3: Array.from({ length: 3 }, (_, index) => normalizeRememberedPixelRef(assignment.set3?.[index])),
  };
}

function createRememberedPixelRef(pixel) {
  const deviceId = String(pixel?.systemId || pixel?.device?.id || pixel?.id || "").trim();
  if (!deviceId) {
    return null;
  }

  return {
    deviceId,
    systemId: deviceId,
    label: getConnectedPixelLabel(pixel),
  };
}

function captureRememberedPixelsSettings() {
  const rememberedAssignmentsByCharacterId = {};
  for (const [characterId, assignment] of Object.entries(pixels.rememberedAssignmentsByCharacterId || {})) {
    const normalized = normalizeRememberedAssignment(assignment);
    if (normalized.single || normalized.set3.some((entry) => Boolean(entry))) {
      rememberedAssignmentsByCharacterId[characterId] = normalized;
    }
  }

  const rememberedSharedSet = Array.from({ length: 3 }, (_, index) =>
    normalizeRememberedPixelRef(pixels.rememberedSharedSet?.[index])
  );

  return {
    rememberedAssignmentsByCharacterId,
    rememberedSharedSet,
  };
}

function persistAppState() {
  const payload = {
    characters: state.characters,
    nextId: state.nextId,
    turnEntries: state.turnEntries,
    currentTurnIndex: state.currentTurnIndex,
    activeCharacterId: state.activeCharacterId,
    round: state.round,
    turnHistory: state.turnHistory,
    turnHistoryIndex: state.turnHistoryIndex,
    usedTurnEntryIds: state.usedTurnEntryIds,
    turnOrderUndoStack: state.turnOrderUndoStack,
    turnOrderRedoStack: state.turnOrderRedoStack,
    characterUndoStack: state.characterUndoStack,
    characterRedoStack: state.characterRedoStack,
    combatLog: state.combatLog,
    nextLogId: state.nextLogId,
    activeTurnTab: state.activeTurnTab,
    activeAbilityCharacterIds: state.activeAbilityCharacterIds,
    dazedAppliedCharacterIds: state.dazedAppliedCharacterIds,
    dazedAppliedRound: state.dazedAppliedRound,
    uiSettings: state.uiSettings,
    pixelsSettings: {
      simulated: pixels.simulated,
      useForRolls: pixels.useForRolls,
      mode: normalizePixelsMode(pixels.mode),
      ...captureRememberedPixelsSettings(),
    },
  };

  localStorage.setItem(APP_STORAGE_KEY, JSON.stringify(payload));
}

function hydrateStateFromStoragePayload(parsedApp) {
  if (!parsedApp || !Array.isArray(parsedApp.characters)) {
    return false;
  }

  state.nextId = 1;
  state.characters = parsedApp.characters.map((character) => createCharacter(character));
  state.nextId = Math.max(Number(parsedApp.nextId) || 1, state.nextId);

  const parsedEntries = Array.isArray(parsedApp.turnEntries) ? parsedApp.turnEntries : [];
  state.turnEntries = parsedEntries.filter((entry) => entry && typeof entry.id === "string").map((entry) => ({
    id: entry.id,
    characterId: Number(entry.characterId),
    name: String(entry.name || "Unbenannt"),
    type: entry.type === "NPC" ? "NPC" : "PC",
    turn: entry.turn === "Bonus" ? "Bonus" : entry.turn === "Move" ? "Move" : "Main",
    initiative: Math.round(Number(entry.initiative) || 0),
    groupInitiative: Math.round(Number(entry.groupInitiative) || Number(entry.initiative) || 0),
    actionIndex: Math.round(Number(entry.actionIndex) || 0),
    critical: entry.critical === "failure" ? "failure" : entry.critical === "success" ? "success" : null,
    dazed: Boolean(entry.dazed),
    specialAbility: typeof entry.specialAbility === "string" && entry.specialAbility ? entry.specialAbility : null,
    abilityActive: Boolean(entry.abilityActive),
  }));

  state.currentTurnIndex = clampTurnPointer(Number(parsedApp.currentTurnIndex), state.turnEntries.length);
  state.activeCharacterId =
    parsedApp.activeCharacterId === null || parsedApp.activeCharacterId === undefined
      ? null
      : Number.isFinite(Number(parsedApp.activeCharacterId))
        ? Number(parsedApp.activeCharacterId)
        : null;
  state.round = Math.max(0, Math.round(Number(parsedApp.round) || 0));
  state.usedTurnEntryIds = Array.isArray(parsedApp.usedTurnEntryIds)
    ? parsedApp.usedTurnEntryIds
        .filter((entryId) => typeof entryId === "string")
        .filter((entryId, index, array) => array.indexOf(entryId) === index)
    : [];
  const validTurnEntryIds = new Set(state.turnEntries.map((entry) => entry.id));
  state.usedTurnEntryIds = state.usedTurnEntryIds.filter((entryId) => validTurnEntryIds.has(entryId));
  syncActiveTurnPointer();

  const parsedHistory = Array.isArray(parsedApp.turnHistory) ? parsedApp.turnHistory : [];
  state.turnHistory = parsedHistory
    .map((snapshot) => ({
      turn: Math.max(0, Math.round(Number(snapshot.turn) || 0)),
      currentTurnIndex: Math.round(Number(snapshot.currentTurnIndex) || 0),
      activeCharacterId:
        snapshot.activeCharacterId === null || snapshot.activeCharacterId === undefined
          ? null
          : Number.isFinite(Number(snapshot.activeCharacterId))
            ? Number(snapshot.activeCharacterId)
            : null,
      usedTurnEntryIds: Array.isArray(snapshot.usedTurnEntryIds)
        ? snapshot.usedTurnEntryIds.filter((id) => typeof id === "string")
        : [],
      activeAbilityCharacterIds: Array.isArray(snapshot.activeAbilityCharacterIds)
        ? snapshot.activeAbilityCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
        : [],
      dazedAppliedCharacterIds: Array.isArray(snapshot.dazedAppliedCharacterIds)
        ? snapshot.dazedAppliedCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
        : [],
      dazedAppliedRound: Math.max(0, Math.round(Number(snapshot.dazedAppliedRound) || 0)),
      entries: Array.isArray(snapshot.entries)
        ? snapshot.entries
            .filter((entry) => entry && typeof entry.id === "string")
            .map((entry) => ({
              id: entry.id,
              characterId: Number(entry.characterId),
              name: String(entry.name || "Unbenannt"),
              type: entry.type === "NPC" ? "NPC" : "PC",
              turn: entry.turn === "Bonus" ? "Bonus" : entry.turn === "Move" ? "Move" : "Main",
              initiative: Math.round(Number(entry.initiative) || 0),
              groupInitiative: Math.round(Number(entry.groupInitiative) || Number(entry.initiative) || 0),
              actionIndex: Math.round(Number(entry.actionIndex) || 0),
              critical: entry.critical === "failure" ? "failure" : entry.critical === "success" ? "success" : null,
              dazed: Boolean(entry.dazed),
              specialAbility:
                typeof entry.specialAbility === "string" && entry.specialAbility ? entry.specialAbility : null,
              abilityActive: Boolean(entry.abilityActive),
            }))
        : [],
      rolls: Array.isArray(snapshot.rolls)
        ? snapshot.rolls
            .filter((roll) => roll && Number.isFinite(Number(roll.id)))
            .map((roll) => ({
              id: Number(roll.id),
              lastRoll: roll.lastRoll === null || roll.lastRoll === undefined ? null : Math.round(Number(roll.lastRoll)),
              critBonusRoll:
                roll.critBonusRoll === null || roll.critBonusRoll === undefined
                  ? null
                  : clamp(Number(roll.critBonusRoll), 1, 6),
              lastPixelsFaces: Array.isArray(roll.lastPixelsFaces)
                ? roll.lastPixelsFaces
                    .map((face) => Math.round(Number(face) || 0))
                    .filter((face) => face >= 1 && face <= 6)
                : null,
              totalInitiative:
                roll.totalInitiative === null || roll.totalInitiative === undefined
                  ? null
                  : Math.round(Number(roll.totalInitiative)),
              dazedUntilRound:
                roll.dazedUntilRound === null || roll.dazedUntilRound === undefined
                  ? null
                  : Math.max(1, Math.round(Number(roll.dazedUntilRound))),
              unfreeDefensePenalty: Math.max(0, Math.round(Number(roll.unfreeDefensePenalty) || 0)),
              paradeClickCount: Math.max(0, Math.round(Number(roll.paradeClickCount) || 0)),
            }))
        : [],
    }))
    .filter((snapshot) => snapshot.turn > 0 && snapshot.entries.length);

  state.turnHistoryIndex = clamp(Number(parsedApp.turnHistoryIndex), -1, state.turnHistory.length - 1);
  state.activeAbilityCharacterIds = Array.isArray(parsedApp.activeAbilityCharacterIds)
    ? parsedApp.activeAbilityCharacterIds
        .filter((id) => Number.isFinite(Number(id)))
        .map((id) => Number(id))
    : [];
  state.dazedAppliedCharacterIds = Array.isArray(parsedApp.dazedAppliedCharacterIds)
    ? parsedApp.dazedAppliedCharacterIds
        .filter((id) => Number.isFinite(Number(id)))
        .map((id) => Number(id))
    : [];
  state.dazedAppliedRound = Math.max(0, Math.round(Number(parsedApp.dazedAppliedRound) || 0));

  const pxSettings = parsedApp.pixelsSettings || {};
  pixels.simulated = Boolean(pxSettings.simulated);
  pixels.useForRolls = Boolean(pxSettings.useForRolls);
  pixels.mode = normalizePixelsMode(pxSettings.mode);
  pixels.assignmentsByCharacterId = {};
  pixels.sharedSet = [null, null, null];
  pixels.rememberedAssignmentsByCharacterId = Object.fromEntries(
    Object.entries(pxSettings.rememberedAssignmentsByCharacterId || {}).map(([characterId, assignment]) => [
      String(characterId),
      normalizeRememberedAssignment(assignment),
    ])
  );
  pixels.rememberedSharedSet = Array.from({ length: 3 }, (_, index) =>
    normalizeRememberedPixelRef(pxSettings.rememberedSharedSet?.[index])
  );

  state.combatLog = Array.isArray(parsedApp.combatLog)
    ? parsedApp.combatLog
        .filter((entry) => entry && typeof entry.message === "string")
        .map((entry) => ({
          id: Math.max(1, Math.round(Number(entry.id) || 1)),
          turn: Math.max(0, Math.round(Number(entry.turn) || 0)),
          timestamp: Number.isFinite(Number(entry.timestamp)) ? Math.round(Number(entry.timestamp)) : null,
          message: entry.message,
        }))
    : [];
  state.nextLogId = Math.max(
    Number(parsedApp.nextLogId) || 1,
    (state.combatLog[state.combatLog.length - 1]?.id || 0) + 1
  );
  state.activeTurnTab = parsedApp.activeTurnTab === "log" ? "log" : "order";
  state.uiSettings = {
    fontScale: clamp(Number(parsedApp?.uiSettings?.fontScale) || 1, 0.7, 1.3),
    forceOneColumn: Boolean(parsedApp?.uiSettings?.forceOneColumn),
  };
  state.turnOrderUndoStack = Array.isArray(parsedApp.turnOrderUndoStack) ? parsedApp.turnOrderUndoStack.map((entry) => ({ ...entry })) : [];
  state.turnOrderRedoStack = Array.isArray(parsedApp.turnOrderRedoStack) ? parsedApp.turnOrderRedoStack.map((entry) => ({ ...entry })) : [];
  state.characterUndoStack = Array.isArray(parsedApp.characterUndoStack) ? parsedApp.characterUndoStack.map((entry) => ({ ...entry })) : [];
  state.characterRedoStack = Array.isArray(parsedApp.characterRedoStack) ? parsedApp.characterRedoStack.map((entry) => ({ ...entry })) : [];

  if (!state.turnEntries.length && state.turnHistoryIndex >= 0) {
    applyTurnSnapshot(state.turnHistory[state.turnHistoryIndex]);
  }

  return true;
}

function loadStateFromStorage() {
  const appRaw = localStorage.getItem(APP_STORAGE_KEY);

  if (!appRaw) {
    return;
  }

  try {
    const parsedApp = JSON.parse(appRaw);
    if (!hydrateStateFromStoragePayload(parsedApp)) {
      return;
    }
  } catch {
    state.characters = [];
    invalidateTurnState({ keepCharacterRolls: true });
    state.combatLog = [];
    state.nextLogId = 1;
    state.activeTurnTab = "order";
    state.uiSettings = {
      fontScale: 1,
      forceOneColumn: false,
    };
    state.turnOrderUndoStack = [];
    state.turnOrderRedoStack = [];
    state.characterUndoStack = [];
    state.characterRedoStack = [];
  }
}

function render() {
  renderRoster();
  renderTracker();
  renderCombatLog();
  updatePixelsControls();
  if (pixelsSettingsDialogEl?.open) {
    renderPixelsSettingsDialog();
  }
  updateHeaderUndoRedoButtons();
}

function parseSimpleCsv(input) {
  const text = String(input ?? "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        field += '"';
        i += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        field += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      continue;
    }
    if (char === ",") {
      row.push(field.trim());
      field = "";
      continue;
    }
    if (char === "\n" || char === "\r") {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(field.trim());
      rows.push(row);
      row = [];
      field = "";
      continue;
    }

    field += char;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field.trim());
    rows.push(row);
  }
  return rows;
}

async function exportPcGroupToCsv() {
  const pcCharacters = state.characters.filter((character) => character.type === "PC");
  if (!pcCharacters.length) {
    return;
  }

  const lines = [];
  lines.push("Name,INI,Charaktervorteil,ManuellerWurf");
  for (const character of pcCharacters) {
    lines.push(
      toCsvLine([
        character.name,
        character.ini,
        character.specialAbility ? getSpecialAbilityLabel(character.specialAbility) : "",
        character.manualRoll ?? "",
      ])
    );
  }

  const payload = `\uFEFF${lines.join("\n")}`;
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  const saved = await saveCsvWithLocationPrompt(payload, `gruppe_sc_${stamp}.csv`);
  if (!saved) {
    return;
  }

  logEvent(`SC-Gruppe als CSV gespeichert (${pcCharacters.length} SC).`);
  persistAppState();
  renderCombatLog();
}

async function exportNpcGroupToCsv() {
  const npcCharacters = state.characters.filter((character) => character.type === "NPC");
  if (!npcCharacters.length) {
    return;
  }

  const grouped = new Map();
  for (const character of npcCharacters) {
    const ability = character.specialAbility ? getSpecialAbilityLabel(character.specialAbility) : "";
    const key = `${character.name}::${character.ini}::${ability}`;
    const entry = grouped.get(key) || {
      name: character.name,
      ini: character.ini,
      ability,
      count: 0,
    };
    entry.count += 1;
    grouped.set(key, entry);
  }

  const lines = [];
  lines.push("Name,INI,Charaktervorteil,Anzahl");
  for (const entry of grouped.values()) {
    lines.push(toCsvLine([entry.name, entry.ini, entry.ability, entry.count]));
  }

  const payload = `\uFEFF${lines.join("\n")}`;
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  const saved = await saveCsvWithLocationPrompt(payload, `gruppe_gegner_${stamp}.csv`);
  if (!saved) {
    return;
  }

  logEvent(`Gegner als CSV gespeichert (${npcCharacters.length} NSC).`);
  persistAppState();
  renderCombatLog();
}

async function exportFullSnapshotToCsv() {
  const snapshot = {
    exportedAt: new Date().toISOString(),
    state: {
      characters: state.characters,
      nextId: state.nextId,
      turnEntries: state.turnEntries,
      currentTurnIndex: state.currentTurnIndex,
      activeCharacterId: state.activeCharacterId,
      round: state.round,
      turnHistory: state.turnHistory,
      turnHistoryIndex: state.turnHistoryIndex,
      usedTurnEntryIds: state.usedTurnEntryIds,
      turnOrderUndoStack: state.turnOrderUndoStack,
      turnOrderRedoStack: state.turnOrderRedoStack,
      characterUndoStack: state.characterUndoStack,
      characterRedoStack: state.characterRedoStack,
      combatLog: state.combatLog,
      nextLogId: state.nextLogId,
      activeTurnTab: state.activeTurnTab,
      activeAbilityCharacterIds: state.activeAbilityCharacterIds,
      dazedAppliedCharacterIds: state.dazedAppliedCharacterIds,
      dazedAppliedRound: state.dazedAppliedRound,
      uiSettings: state.uiSettings,
    },
    pixels: {
      simulated: pixels.simulated,
      useForRolls: pixels.useForRolls,
      mode: normalizePixelsMode(pixels.mode),
      lastStatus: pixels.lastStatus,
      ...captureRememberedPixelsSettings(),
    },
  };

  const lines = [];
  lines.push("Typ,Daten");
  lines.push(toCsvLine(["FULL_SNAPSHOT", JSON.stringify(snapshot)]));

  const payload = `\uFEFF${lines.join("\n")}`;
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  const saved = await saveCsvWithLocationPrompt(payload, `snapshot_voll_${stamp}.csv`);
  if (!saved) {
    return;
  }

  logEvent("Vollsnapshot als CSV gespeichert.");
  persistAppState();
  renderCombatLog();
}

function importFullSnapshotRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return false;
  }

  const payloadRow = rows.find((row) => {
    const type = String(row?.[0] || "").trim().toUpperCase();
    return type === "FULL_SNAPSHOT";
  });
  if (!payloadRow || !payloadRow[1]) {
    return false;
  }

  let parsedPayload;
  try {
    parsedPayload = JSON.parse(payloadRow[1]);
  } catch {
    return false;
  }
  if (!parsedPayload || typeof parsedPayload !== "object" || !parsedPayload.state) {
    return false;
  }

  try {
    if (!hydrateStateFromStoragePayload(parsedPayload.state)) {
      return false;
    }

    const px = parsedPayload.pixels || {};
    pixels.simulated = Boolean(px.simulated);
    pixels.useForRolls = Boolean(px.useForRolls);
    pixels.mode = normalizePixelsMode(px.mode);
    pixels.assignmentsByCharacterId = {};
    pixels.sharedSet = [null, null, null];
    pixels.rememberedAssignmentsByCharacterId = Object.fromEntries(
      Object.entries(px.rememberedAssignmentsByCharacterId || {}).map(([characterId, assignment]) => [
        String(characterId),
        normalizeRememberedAssignment(assignment),
      ])
    );
    pixels.rememberedSharedSet = Array.from({ length: 3 }, (_, index) =>
      normalizeRememberedPixelRef(px.rememberedSharedSet?.[index])
    );
    pixels.lastStatus =
      typeof px.lastStatus === "string" && px.lastStatus.trim()
        ? px.lastStatus
        : "Pixels: nicht verbunden (Chromium-Browser mit Web Bluetooth erforderlich).";

    applyUiSettings(state.uiSettings);
    updatePixelsControls();
    persistAppState();
    return true;
  } catch {
    return false;
  }
}

function importPcGroupRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return 0;
  }

  const hasHeader = String(rows[0]?.[0] || "")
    .trim()
    .toLowerCase() === "name";
  const startIndex = hasHeader ? 1 : 0;
  let importedCount = 0;

  pushCharacterUndoSnapshot("SC-Gruppenimport zurückgesetzt.");
  for (let i = startIndex; i < rows.length; i += 1) {
    const [nameRaw, iniRaw, abilityRaw, manualRollRaw] = rows[i] || [];
    const name = String(nameRaw || "").trim();
    if (!name) {
      continue;
    }

    const parsedAbility = parseImportedAbility(abilityRaw);
    const manualRollValue = String(manualRollRaw || "").trim();
    const manualRoll = manualRollValue ? clamp(Number(manualRollValue), 3, 18) : null;
    const useManualRoll = manualRoll !== null;

    state.characters.push(
      createCharacter({
        name,
        type: "PC",
        ini: clamp(Number(iniRaw), 1, 30),
        specialAbility: parsedAbility,
        manualRoll,
        useManualRoll,
      })
    );
    importedCount += 1;
  }

  return importedCount;
}

function parseImportedAbility(rawAbility) {
  const value = String(rawAbility || "").trim();
  if (!value) {
    return null;
  }
  if (value === "INI Schützen" || value === "INI Archer") return "INI Archer";
  if (value === "INI Flink" || value === "INI Nimble") return "INI Nimble";
  if (value === "INI Straßenschläger" || value === "INI Boxer") return "INI Boxer";
  if (value === "INI Pazifist" || value === "INI Pacifist") return "INI Pacifist";
  if (value === "INI Reiter" || value === "INI Horsemen") return "INI Horsemen";
  return value;
}

function toCsvLine(values) {
  return values.map((value) => escapeCsvField(value)).join(",");
}

function escapeCsvField(value) {
  const text = String(value ?? "");
  const escaped = text.replaceAll('"', '""');
  if (/[",\n\r]/.test(text)) {
    return `"${escaped}"`;
  }
  return escaped;
}

async function saveCsvWithLocationPrompt(content, suggestedName) {
  if (typeof window !== "undefined" && typeof window.showSaveFilePicker === "function") {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName,
        types: [
          {
            description: "CSV-Datei",
            accept: { "text/csv": [".csv"] },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(content);
      await writable.close();
      return true;
    } catch {
      return false;
    }
  }

  const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = suggestedName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  return true;
}

function importEnemyGroupRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return 0;
  }

  const hasHeader = String(rows[0]?.[0] || "")
    .trim()
    .toLowerCase() === "name";
  const startIndex = hasHeader ? 1 : 0;
  let importedCount = 0;

  pushCharacterUndoSnapshot("Gegnergruppe-Import zurückgesetzt.");
  for (let i = startIndex; i < rows.length; i += 1) {
    const [nameRaw, iniRaw, abilityRaw, countRaw] = rows[i] || [];
    const name = String(nameRaw || "").trim();
    if (!name) {
      continue;
    }

    const ini = clamp(Number(iniRaw), 1, 30);
    const specialAbility = parseImportedAbility(abilityRaw);
    const count = Math.max(1, Math.min(99, Math.round(Number(countRaw) || 1)));

    for (let index = 0; index < count; index += 1) {
      state.characters.push(
        createCharacter({
          name: count > 1 ? `${name} ${index + 1}` : name,
          type: "NPC",
          ini,
          specialAbility,
        })
      );
      importedCount += 1;
    }
  }

  return importedCount;
}

function applyUiSettings(uiSettings) {
  const fontScale = clamp(Number(uiSettings?.fontScale) || 1, 0.7, 1.3);
  const forceOneColumn = Boolean(uiSettings?.forceOneColumn);

  state.uiSettings = { fontScale, forceOneColumn };
  if (appShellEl) {
    appShellEl.style.zoom = String(fontScale);
    const supportsZoom = typeof CSS !== "undefined" && typeof CSS.supports === "function" && CSS.supports("zoom", "1");
    if (supportsZoom) {
      appShellEl.style.transform = "";
      appShellEl.style.transformOrigin = "";
      appShellEl.style.width = "";
    } else {
      appShellEl.style.transform = `scale(${fontScale})`;
      appShellEl.style.transformOrigin = "top center";
      appShellEl.style.width = `${100 / fontScale}%`;
    }
    appShellEl.classList.toggle("one-column", forceOneColumn);
  }
  if (fontSizeSliderEl) {
    fontSizeSliderEl.value = String(Math.round(fontScale * 100));
  }
  if (forceOneColumnEl) {
    forceOneColumnEl.checked = forceOneColumn;
  }
}

function resetEntireAppState() {
  state.characters = [];
  state.nextId = 1;
  state.turnEntries = [];
  state.currentTurnIndex = -1;
  state.activeCharacterId = null;
  state.round = 0;
  state.turnHistory = [];
  state.turnHistoryIndex = -1;
  state.usedTurnEntryIds = [];
  state.turnOrderUndoStack = [];
  state.turnOrderRedoStack = [];
  state.characterUndoStack = [];
  state.characterRedoStack = [];
  state.combatLog = [];
  state.nextLogId = 1;
  state.activeTurnTab = "order";
  state.activeAbilityCharacterIds = [];
  state.dazedAppliedCharacterIds = [];
  state.dazedAppliedRound = 0;
  state.uiSettings = {
    fontScale: 1,
    forceOneColumn: false,
  };

  pixels.useForRolls = false;
  pixels.simulated = false;
  pixels.mode = PIXELS_MODE.PC_SET_3;
  pixels.assignmentsByCharacterId = {};
  pixels.sharedSet = [null, null, null];
  pixels.rememberedAssignmentsByCharacterId = {};
  pixels.rememberedSharedSet = [null, null, null];
  pixels.lastStatus = "Pixels: nicht verbunden (Chromium-Browser mit Web Bluetooth erforderlich).";

  localStorage.removeItem(APP_STORAGE_KEY);
  applyUiSettings(state.uiSettings);
  updatePixelsControls();
  persistAppState();
  render();
}

function showWarningDialog(message, title = "Warnung", confirmLabel = "Bestätigen", cancelLabel = "Abbrechen") {
  if (!warningDialogEl || !warningDialogMessageEl || !warningDialogTitleEl) {
    return Promise.resolve(window.confirm(message));
  }

  warningDialogTitleEl.textContent = title;
  warningDialogMessageEl.textContent = message;
  if (warningDialogConfirmBtn) {
    warningDialogConfirmBtn.textContent = confirmLabel;
  }
  if (warningDialogCancelBtn) {
    warningDialogCancelBtn.textContent = cancelLabel;
  }

  if (warningDialogResolver) {
    warningDialogResolver(false);
    warningDialogResolver = null;
  }

  return new Promise((resolve) => {
    warningDialogResolver = resolve;
    if (typeof warningDialogEl.showModal === "function") {
      warningDialogEl.showModal();
    } else {
      warningDialogEl.setAttribute("open", "open");
    }
  });
}

function resolveWarningDialog(value) {
  if (!warningDialogResolver) {
    return;
  }
  const resolve = warningDialogResolver;
  warningDialogResolver = null;
  if (warningDialogEl?.open) {
    warningDialogEl.close();
  }
  resolve(Boolean(value));
}

function roll3d6() {
  return d6() + d6() + d6();
}

function d6() {
  return Math.floor(Math.random() * 6) + 1;
}

function clamp(number, min, max) {
  if (Number.isNaN(number)) {
    return min;
  }

  return Math.max(min, Math.min(max, number));
}

function clampTurnPointer(value, entryCount) {
  if (Number.isNaN(value)) {
    return -1;
  }

  return Math.max(-1, Math.min(Math.round(value), entryCount));
}

function normalizeDamageMonitorMarks(value) {
  const marks = Array.from({ length: DAMAGE_MONITOR_COLUMN_COUNT }, () => false);
  if (!Array.isArray(value)) {
    return marks;
  }

  for (let index = 0; index < DAMAGE_MONITOR_COLUMN_COUNT; index += 1) {
    marks[index] = Boolean(value[index]);
  }
  marks[0] = false;
  return marks;
}

function parseDamagePenalty(rawValue) {
  if (typeof rawValue !== "string") {
    return 0;
  }
  const trimmed = rawValue.trim();
  if (!trimmed || trimmed === "-") {
    return 0;
  }
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? Math.abs(parsed) : 0;
}

function computeDamagePenalty(character) {
  const marks = normalizeDamageMonitorMarks(character.damageMonitorMarks);
  let rightmostIndex = -1;
  for (let index = DAMAGE_MONITOR_COLUMN_COUNT - 1; index >= 1; index -= 1) {
    if (marks[index]) {
      rightmostIndex = index;
      break;
    }
  }

  if (rightmostIndex === -1) {
    return { qm: 0, bew: 0 };
  }

  return {
    qm: parseDamagePenalty(DAMAGE_QM_VALUES[rightmostIndex] ?? "-"),
    bew: parseDamagePenalty(DAMAGE_BEW_VALUES[rightmostIndex] ?? "-"),
  };
}

function getTypeLabel(type) {
  return type === "NPC" ? "NSC" : "SC";
}

function getSpecialAbilityLabel(ability) {
  if (ability === "INI Archer") return "INI Schützen";
  if (ability === "INI Nimble") return "INI Flink";
  if (ability === "INI Boxer") return "INI Straßenschläger";
  if (ability === "INI Pacifist") return "INI Pazifist";
  if (ability === "INI Horsemen") return "INI Reiter";
  return ability;
}

function getDefaultDazedUntilRound() {
  const currentRound = state.round > 0 ? state.round : 1;
  return currentRound + 1;
}

function isCharacterDazed(character) {
  if (character.dazedUntilRound === null || character.dazedUntilRound === undefined) {
    return false;
  }

  if (state.round <= 0) {
    return false;
  }

  return state.round <= character.dazedUntilRound;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function supportsPixels() {
  return typeof navigator !== "undefined" && Boolean(navigator.bluetooth);
}

function closeCharacterSettingsMenu() {
  if (characterSettingsMenu) {
    characterSettingsMenu.open = false;
  }
}

function initSettingsMenusHoverBehavior() {
  for (const menu of settingsMenuEls) {
    let closeTimerId = null;

    menu.addEventListener("mouseenter", () => {
      if (closeTimerId !== null) {
        clearTimeout(closeTimerId);
        closeTimerId = null;
      }
    });

    menu.addEventListener("mouseleave", () => {
      if (!menu.open) {
        return;
      }
      if (closeTimerId !== null) {
        clearTimeout(closeTimerId);
      }
      closeTimerId = setTimeout(() => {
        menu.open = false;
        closeTimerId = null;
      }, 220);
    });
  }
}

function canUsePixelsForRolls() {
  return pixels.useForRolls && (hasAnyConnectedPixels() || pixels.simulated);
}

function formatVersionNumber(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `Version: ${year}${month}${day}:${hours}${minutes}`;
}

function updateVersionBadge() {
  if (!appVersionEl) {
    return;
  }
  appVersionEl.textContent = formatVersionNumber(new Date());
}

function formatLogTimestamp(timestamp) {
  if (!Number.isFinite(Number(timestamp))) {
    return "--:--:--";
  }

  const date = new Date(Number(timestamp));
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  return `${hours}:${minutes}:${seconds}`;
}

function isTurnEntryUsed(entryId) {
  return state.usedTurnEntryIds.includes(entryId);
}

function setTurnEntryUsed(entryId, used) {
  const wasUsed = isTurnEntryUsed(entryId);
  if (wasUsed === used) {
    return false;
  }

  if (used) {
    if (!state.usedTurnEntryIds.includes(entryId)) {
      state.usedTurnEntryIds.push(entryId);
    }
    return true;
  }

  state.usedTurnEntryIds = state.usedTurnEntryIds.filter((id) => id !== entryId);
  return true;
}

function getTurnCharacterOrder() {
  const seen = new Set();
  const order = [];

  for (const entry of state.turnEntries) {
    if (seen.has(entry.characterId)) {
      continue;
    }
    seen.add(entry.characterId);
    order.push(entry.characterId);
  }

  return order;
}

function getTurnCharacterOrderIndexMap() {
  const orderMap = new Map();
  const order = getTurnCharacterOrder();
  for (let index = 0; index < order.length; index += 1) {
    orderMap.set(order[index], index);
  }
  return orderMap;
}

function hasUnresolvedActionForCharacter(characterId) {
  return state.turnEntries.some((entry) => entry.characterId === characterId && !isTurnEntryUsed(entry.id));
}

function getFirstTurnIndexForCharacter(characterId) {
  return state.turnEntries.findIndex((entry) => entry.characterId === characterId);
}

function getFirstUnresolvedTurnIndexForCharacter(characterId) {
  return state.turnEntries.findIndex((entry) => entry.characterId === characterId && !isTurnEntryUsed(entry.id));
}

function getActiveCharacterId() {
  if (state.activeCharacterId !== null) {
    const hasCharacter = state.turnEntries.some((entry) => entry.characterId === state.activeCharacterId);
    if (hasCharacter) {
      return state.activeCharacterId;
    }
  }

  const nextIndex = getNextUnresolvedTurnIndex();
  if (nextIndex >= 0 && nextIndex < state.turnEntries.length) {
    return state.turnEntries[nextIndex].characterId;
  }

  return state.turnEntries[0]?.characterId ?? null;
}

function syncActiveTurnPointer() {
  if (!state.turnEntries.length) {
    state.activeCharacterId = null;
    state.currentTurnIndex = -1;
    return;
  }

  const activeCharacterId = getActiveCharacterId();
  state.activeCharacterId = activeCharacterId;
  if (activeCharacterId === null) {
    state.currentTurnIndex = getNextUnresolvedTurnIndex();
    return;
  }

  const unresolvedIndex = getFirstUnresolvedTurnIndexForCharacter(activeCharacterId);
  if (unresolvedIndex >= 0) {
    state.currentTurnIndex = unresolvedIndex;
    return;
  }

  const firstIndex = getFirstTurnIndexForCharacter(activeCharacterId);
  state.currentTurnIndex = firstIndex >= 0 ? firstIndex : getNextUnresolvedTurnIndex();
}

function getNextCharacterIdForTurnOrder() {
  const order = getTurnCharacterOrder();
  if (!order.length) {
    return null;
  }

  const activeCharacterId = getActiveCharacterId();
  if (activeCharacterId === null) {
    return order[0];
  }

  const startIndex = order.indexOf(activeCharacterId);
  if (startIndex < 0) {
    return order[0];
  }

  for (let step = 1; step < order.length; step += 1) {
    const candidateId = order[(startIndex + step) % order.length];
    if (hasUnresolvedActionForCharacter(candidateId)) {
      return candidateId;
    }
  }

  return null;
}

function getAdjacentCharacterIdForTurnOrder(direction) {
  const order = getTurnCharacterOrder();
  if (!order.length) {
    return null;
  }

  const activeCharacterId = getActiveCharacterId();
  if (activeCharacterId === null) {
    return order[0];
  }

  const currentIndex = order.indexOf(activeCharacterId);
  if (currentIndex < 0) {
    return order[0];
  }

  const targetIndex = currentIndex + direction;
  if (targetIndex < 0 || targetIndex >= order.length) {
    return null;
  }

  return order[targetIndex];
}

function canInteractWithTurnEntry(entry, orderIndexByCharacter = null) {
  return Boolean(entry);
}

function getLatestStackEntry(stacks) {
  let latest = null;
  for (const candidate of stacks) {
    if (!candidate) {
      continue;
    }
    const candidateTime = Number(candidate.createdAt) || 0;
    const latestTime = Number(latest?.createdAt) || 0;
    if (!latest || candidateTime >= latestTime) {
      latest = candidate;
    }
  }
  return latest;
}

function updateHeaderUndoRedoButtons() {
  if (headerUndoBtn) {
    headerUndoBtn.disabled = !state.characterUndoStack.length && !state.turnOrderUndoStack.length;
  }
  if (headerRedoBtn) {
    headerRedoBtn.disabled = !state.characterRedoStack.length && !state.turnOrderRedoStack.length;
  }
}

function undoLatestPanelAction() {
  const latestCharacterUndo = state.characterUndoStack[state.characterUndoStack.length - 1] || null;
  const latestTurnUndo = state.turnOrderUndoStack[state.turnOrderUndoStack.length - 1] || null;
  const latest = getLatestStackEntry([latestCharacterUndo, latestTurnUndo]);
  if (!latest) {
    return;
  }

  if (latest === latestCharacterUndo) {
    const snapshot = state.characterUndoStack.pop();
    if (!snapshot) {
      return;
    }
    state.characterRedoStack.push(captureCharacterUndoSnapshot(snapshot.reason));
    restoreCharacterUndoSnapshot(snapshot);
    logEvent(`Rückg.: ${snapshot.reason}`);
    persistAppState();
    render();
    return;
  }

  const snapshot = state.turnOrderUndoStack.pop();
  if (!snapshot) {
    return;
  }
  state.turnOrderRedoStack.push(captureTurnOrderUndoSnapshot(snapshot.reason));
  restoreTurnOrderUndoSnapshot(snapshot);
  logEvent(`Rückg.: ${snapshot.reason}`);
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function redoLatestPanelAction() {
  const latestCharacterRedo = state.characterRedoStack[state.characterRedoStack.length - 1] || null;
  const latestTurnRedo = state.turnOrderRedoStack[state.turnOrderRedoStack.length - 1] || null;
  const latest = getLatestStackEntry([latestCharacterRedo, latestTurnRedo]);
  if (!latest) {
    return;
  }

  if (latest === latestCharacterRedo) {
    const snapshot = state.characterRedoStack.pop();
    if (!snapshot) {
      return;
    }
    state.characterUndoStack.push(captureCharacterUndoSnapshot(snapshot.reason));
    restoreCharacterUndoSnapshot(snapshot);
    logEvent(`Wdh.: ${snapshot.reason}`);
    persistAppState();
    render();
    return;
  }

  const snapshot = state.turnOrderRedoStack.pop();
  if (!snapshot) {
    return;
  }
  state.turnOrderUndoStack.push(captureTurnOrderUndoSnapshot(snapshot.reason));
  restoreTurnOrderUndoSnapshot(snapshot);
  logEvent(`Wdh.: ${snapshot.reason}`);
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function captureTurnOrderUndoSnapshot(reason) {
  return {
    reason,
    createdAt: Date.now(),
    turnEntries: state.turnEntries.map((entry) => ({ ...entry })),
    usedTurnEntryIds: [...state.usedTurnEntryIds],
    currentTurnIndex: state.currentTurnIndex,
    activeCharacterId: state.activeCharacterId,
    activeAbilityCharacterIds: [...state.activeAbilityCharacterIds],
    dazedAppliedCharacterIds: [...state.dazedAppliedCharacterIds],
    dazedAppliedRound: state.dazedAppliedRound,
    characterDazedById: state.characters.map((character) => ({
      id: character.id,
      dazedUntilRound: character.dazedUntilRound ?? null,
    })),
    characterUnfreeDefenseById: state.characters.map((character) => ({
      id: character.id,
      unfreeDefensePenalty: Math.max(0, Math.round(Number(character.unfreeDefensePenalty) || 0)),
    })),
    characterRollStateById: state.characters.map((character) => ({
      id: character.id,
      lastRoll: character.lastRoll,
      critBonusRoll: character.critBonusRoll ?? null,
      lastPixelsFaces: Array.isArray(character.lastPixelsFaces) ? [...character.lastPixelsFaces] : null,
      totalInitiative: character.totalInitiative,
      paradeClickCount: Math.max(0, Math.round(Number(character.paradeClickCount) || 0)),
    })),
  };
}

function pushTurnOrderUndoSnapshot(reason) {
  if (!state.turnEntries.length) {
    return;
  }
  state.turnOrderUndoStack.push(captureTurnOrderUndoSnapshot(reason));
  state.turnOrderRedoStack = [];
}

function restoreTurnOrderUndoSnapshot(snapshot) {
  if (!snapshot) {
    return;
  }

  state.turnEntries = Array.isArray(snapshot.turnEntries) ? snapshot.turnEntries.map((entry) => ({ ...entry })) : [];
  state.usedTurnEntryIds = Array.isArray(snapshot.usedTurnEntryIds)
    ? snapshot.usedTurnEntryIds.filter((id) => typeof id === "string")
    : [];
  const validEntryIds = new Set(state.turnEntries.map((entry) => entry.id));
  state.usedTurnEntryIds = state.usedTurnEntryIds.filter((id) => validEntryIds.has(id));
  state.activeCharacterId =
    snapshot.activeCharacterId === null || snapshot.activeCharacterId === undefined
      ? null
      : Number.isFinite(Number(snapshot.activeCharacterId))
        ? Number(snapshot.activeCharacterId)
        : null;
  state.currentTurnIndex = clampTurnPointer(Number(snapshot.currentTurnIndex), state.turnEntries.length);
  state.activeAbilityCharacterIds = Array.isArray(snapshot.activeAbilityCharacterIds)
    ? snapshot.activeAbilityCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
    : [];
  state.dazedAppliedCharacterIds = Array.isArray(snapshot.dazedAppliedCharacterIds)
    ? snapshot.dazedAppliedCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
    : [];
  state.dazedAppliedRound = Math.max(0, Math.round(Number(snapshot.dazedAppliedRound) || state.round || 0));

  const dazedById = new Map(
    Array.isArray(snapshot.characterDazedById)
      ? snapshot.characterDazedById
          .filter((entry) => entry && Number.isFinite(Number(entry.id)))
          .map((entry) => [Number(entry.id), entry.dazedUntilRound ?? null])
      : []
  );
  state.characters = state.characters.map((character) => ({
    ...character,
    dazedUntilRound: dazedById.has(character.id) ? dazedById.get(character.id) : character.dazedUntilRound ?? null,
  }));

  const unfreeDefenseById = new Map(
    Array.isArray(snapshot.characterUnfreeDefenseById)
      ? snapshot.characterUnfreeDefenseById
          .filter((entry) => entry && Number.isFinite(Number(entry.id)))
          .map((entry) => [Number(entry.id), Math.max(0, Math.round(Number(entry.unfreeDefensePenalty) || 0))])
      : []
  );
  state.characters = state.characters.map((character) => ({
    ...character,
    unfreeDefensePenalty: unfreeDefenseById.has(character.id)
      ? unfreeDefenseById.get(character.id)
      : Math.max(0, Math.round(Number(character.unfreeDefensePenalty) || 0)),
  }));

  const rollStateById = new Map(
    Array.isArray(snapshot.characterRollStateById)
      ? snapshot.characterRollStateById
          .filter((entry) => entry && Number.isFinite(Number(entry.id)))
          .map((entry) => [
            Number(entry.id),
            {
              lastRoll: entry.lastRoll === null || entry.lastRoll === undefined ? null : Number(entry.lastRoll),
              critBonusRoll:
                entry.critBonusRoll === null || entry.critBonusRoll === undefined
                  ? null
                  : clamp(Number(entry.critBonusRoll), 1, 6),
              lastPixelsFaces: Array.isArray(entry.lastPixelsFaces)
                ? entry.lastPixelsFaces
                    .map((face) => Math.round(Number(face) || 0))
                    .filter((face) => face >= 1 && face <= 6)
                : null,
              totalInitiative:
                entry.totalInitiative === null || entry.totalInitiative === undefined
                  ? null
                  : Number(entry.totalInitiative),
              paradeClickCount: Math.max(0, Math.round(Number(entry.paradeClickCount) || 0)),
            },
          ])
      : []
  );
  state.characters = state.characters.map((character) => {
    const rollState = rollStateById.get(character.id);
    if (!rollState) {
      return character;
    }
    return {
      ...character,
      lastRoll: rollState.lastRoll,
      critBonusRoll: rollState.critBonusRoll ?? null,
      lastPixelsFaces: Array.isArray(rollState.lastPixelsFaces) ? [...rollState.lastPixelsFaces] : null,
      totalInitiative: rollState.totalInitiative,
      paradeClickCount: Math.max(0, Math.round(Number(rollState.paradeClickCount) || 0)),
    };
  });

  syncActiveTurnPointer();
}

function captureCharacterUndoSnapshot(reason) {
  return {
    reason,
    createdAt: Date.now(),
    snapshot: {
      characters: state.characters.map((character) => ({
        ...character,
        damageMonitorMarks: normalizeDamageMonitorMarks(character.damageMonitorMarks),
      })),
      nextId: state.nextId,
      turnEntries: state.turnEntries.map((entry) => ({ ...entry })),
      currentTurnIndex: state.currentTurnIndex,
      activeCharacterId: state.activeCharacterId,
      round: state.round,
      turnHistory: state.turnHistory.map((entry) => ({ ...entry })),
      turnHistoryIndex: state.turnHistoryIndex,
      usedTurnEntryIds: [...state.usedTurnEntryIds],
      turnOrderUndoStack: state.turnOrderUndoStack.map((entry) => ({ ...entry })),
      activeAbilityCharacterIds: [...state.activeAbilityCharacterIds],
      dazedAppliedCharacterIds: [...state.dazedAppliedCharacterIds],
      dazedAppliedRound: state.dazedAppliedRound,
    },
  };
}

function pushCharacterUndoSnapshot(reason) {
  state.characterUndoStack.push(captureCharacterUndoSnapshot(reason));
  state.characterRedoStack = [];
  if (state.characterUndoStack.length > 50) {
    state.characterUndoStack.shift();
  }
}

function restoreCharacterUndoSnapshot(entry) {
  if (!entry || !entry.snapshot) {
    return;
  }

  const snapshot = entry.snapshot;
  state.characters = Array.isArray(snapshot.characters)
    ? snapshot.characters.map((character) => ({
        ...character,
        damageMonitorMarks: normalizeDamageMonitorMarks(character.damageMonitorMarks),
      }))
    : [];
  state.nextId = Math.max(1, Number(snapshot.nextId) || 1);
  state.turnEntries = Array.isArray(snapshot.turnEntries) ? snapshot.turnEntries.map((item) => ({ ...item })) : [];
  state.currentTurnIndex = clampTurnPointer(Number(snapshot.currentTurnIndex), state.turnEntries.length);
  state.activeCharacterId =
    snapshot.activeCharacterId === null || snapshot.activeCharacterId === undefined
      ? null
      : Number.isFinite(Number(snapshot.activeCharacterId))
        ? Number(snapshot.activeCharacterId)
        : null;
  state.round = Math.max(0, Math.round(Number(snapshot.round) || 0));
  state.turnHistory = Array.isArray(snapshot.turnHistory) ? snapshot.turnHistory.map((item) => ({ ...item })) : [];
  state.turnHistoryIndex = clamp(Number(snapshot.turnHistoryIndex), -1, state.turnHistory.length - 1);
  state.usedTurnEntryIds = Array.isArray(snapshot.usedTurnEntryIds)
    ? snapshot.usedTurnEntryIds.filter((id) => typeof id === "string")
    : [];
  state.turnOrderUndoStack = Array.isArray(snapshot.turnOrderUndoStack)
    ? snapshot.turnOrderUndoStack.map((item) => ({ ...item }))
    : [];
  state.activeAbilityCharacterIds = Array.isArray(snapshot.activeAbilityCharacterIds)
    ? snapshot.activeAbilityCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
    : [];
  state.dazedAppliedCharacterIds = Array.isArray(snapshot.dazedAppliedCharacterIds)
    ? snapshot.dazedAppliedCharacterIds.filter((id) => Number.isFinite(Number(id))).map((id) => Number(id))
    : [];
  state.dazedAppliedRound = Math.max(0, Math.round(Number(snapshot.dazedAppliedRound) || state.round || 0));

  syncActiveTurnPointer();
}

function getNextUnresolvedTurnIndex() {
  for (let index = 0; index < state.turnEntries.length; index += 1) {
    if (!isTurnEntryUsed(state.turnEntries[index].id)) {
      return index;
    }
  }

  return state.turnEntries.length;
}

function toggleTurnEntryUsed(entryId) {
  const entry = state.turnEntries.find((item) => item.id === entryId);
  if (!entry) {
    return;
  }

  if (!canInteractWithTurnEntry(entry)) {
    return;
  }

  const activeCharacterId = getActiveCharacterId();
  pushTurnOrderUndoSnapshot("INI-Reihenfolge-Aktionen wiederhergestellt.");
  const nextUsed = !isTurnEntryUsed(entryId);
  const changed = setTurnEntryUsed(entryId, nextUsed);
  if (!changed) {
    return;
  }

  if (!hasUnresolvedActionForCharacter(activeCharacterId)) {
    const nextCharacterId = getNextCharacterIdForTurnOrder();
    if (nextCharacterId !== null) {
      state.activeCharacterId = nextCharacterId;
    }
  }
  syncActiveTurnPointer();
  logEvent(`Aktion ${nextUsed ? "genutzt" : "ungenutzt"}: ${entry.name} (${entry.turn}).`);
  syncCurrentSnapshotFromState();
  persistAppState();
  render();
}

function updatePixelsStatus(message) {
  pixels.lastStatus = message;
  if (pixelsStatusEl) {
    pixelsStatusEl.textContent = message;
  }
}

function setPixelsRollDialogStatus(message) {
  if (pixelsRollStatusEl) {
    pixelsRollStatusEl.textContent = message;
  }
}

function setPixelsRollDialogTitle(message) {
  if (pixelsRollTitleEl) {
    pixelsRollTitleEl.textContent = message;
  }
}

function clearPixelsRollDialogRows() {
  rollDialogState.rowsByCharacterId.clear();
  if (pixelsRollListEl) {
    pixelsRollListEl.innerHTML = "";
  }
}

function getRollModeLabel(mode) {
  if (mode === "manual") return "Manuell";
  if (mode === "pixels") return "Pixels";
  if (mode === "auto") return "Auto";
  return "Übersprungen";
}

function getCharacterRollMode(character, usePixelsForTurn) {
  if (!character || character.incapacitated) {
    return "skip";
  }
  if (character.type === "PC" && character.useManualRoll) {
    return "manual";
  }
  if (usePixelsForTurn && character.type === "PC") {
    return "pixels";
  }
  return "auto";
}

function shouldUseRollDialogForCharacters(characters, usePixelsForTurn) {
  return characters.some((character) => {
    const mode = getCharacterRollMode(character, usePixelsForTurn);
    return mode === "manual" || mode === "pixels";
  });
}

function createPixelsRollDialogBadge(text, className = "") {
  const badge = document.createElement("span");
  badge.className = `pixels-roll-badge${className ? ` ${className}` : ""}`;
  badge.textContent = text;
  return badge;
}

function createPixelsRollDialogRow(character, mode) {
  const row = document.createElement("div");
  row.className = "pixels-roll-row pending";

  const head = document.createElement("div");
  head.className = "pixels-roll-row-head";

  const nameEl = document.createElement("strong");
  nameEl.textContent = character.name;
  head.appendChild(nameEl);

  const badgesEl = document.createElement("div");
  badgesEl.className = "pixels-roll-row-badges";
  badgesEl.appendChild(createPixelsRollDialogBadge(getRollModeLabel(mode)));
  head.appendChild(badgesEl);

  const statusEl = document.createElement("p");
  statusEl.className = "pixels-roll-row-status";
  statusEl.textContent =
    mode === "skip" ? "Wird übersprungen." : mode === "auto" ? "Automatischer Wurf vorbereitet." : "Warte auf Wurf.";

  const detailEl = document.createElement("p");
  detailEl.className = "pixels-roll-row-detail";
  detailEl.textContent =
    mode === "manual"
      ? "3w6 im Feld eintragen und übernehmen."
      : mode === "pixels"
        ? "Warte auf verbundenen Pixels-Wurf."
        : mode === "auto"
          ? "Wird automatisch gewürfelt."
          : "Aktionsunfähig.";

  const controlsEl = document.createElement("div");
  controlsEl.className = "pixels-roll-row-controls";

  row.appendChild(head);
  row.appendChild(statusEl);
  row.appendChild(detailEl);
  row.appendChild(controlsEl);
  pixelsRollListEl?.appendChild(row);

  rollDialogState.rowsByCharacterId.set(character.id, {
    row,
    badgesEl,
    statusEl,
    detailEl,
    controlsEl,
  });
}

function updatePixelsRollDialogRow(characterId, options = {}) {
  const refs = rollDialogState.rowsByCharacterId.get(characterId);
  if (!refs) {
    return;
  }

  if (options.stateClass) {
    refs.row.className = `pixels-roll-row ${options.stateClass}`;
  }
  if (options.status !== undefined) {
    refs.statusEl.textContent = options.status;
  }
  if (options.detail !== undefined) {
    refs.detailEl.textContent = options.detail;
  }
  if (options.clearControls) {
    refs.controlsEl.innerHTML = "";
  }
  if (Array.isArray(options.resultBadges)) {
    refs.badgesEl.innerHTML = "";
    refs.badgesEl.appendChild(createPixelsRollDialogBadge(getRollModeLabel(options.mode || "auto")));
    for (const badge of options.resultBadges) {
      refs.badgesEl.appendChild(createPixelsRollDialogBadge(badge.text, badge.className));
    }
  }
}

function updatePixelsRollDialogRowResult(character, rollData, totalInitiative, mode) {
  const resultBadges = [];
  if (rollData.total === 18) {
    resultBadges.push({ text: "Krit. Erfolg", className: "crit-success" });
  } else if (rollData.total === 3) {
    resultBadges.push({ text: "Krit. Fehlschlag", className: "crit-failure" });
  }
  resultBadges.push({ text: `Ges. ${totalInitiative}`, className: "result" });

  const facesText =
    Array.isArray(rollData.faces) && rollData.faces.length ? ` [${rollData.faces.join(", ")}]` : "";
  const critBonusText =
    Number.isFinite(Number(rollData.critBonusRoll)) && Number(rollData.critBonusRoll) > 0
      ? `, Krit-W6=${rollData.critBonusRoll}`
      : "";
  const sourceText = rollData.source === "pixels-fallback" ? " Pixels-Fallback." : "";
  updatePixelsRollDialogRow(character.id, {
    mode,
    stateClass: "done",
    status: `Gewürfelt: 3w6=${rollData.total}${facesText}.`,
    detail: `INI ${character.ini}${character.surprised ? ", Überr. -10" : ""}${critBonusText}, gesamt ${totalInitiative}.${sourceText}`,
    clearControls: true,
    resultBadges,
  });
}

function describeRollDialogMix(characters, usePixelsForTurn) {
  let manualCount = 0;
  let pixelsCount = 0;
  let autoCount = 0;
  for (const character of characters) {
    const mode = getCharacterRollMode(character, usePixelsForTurn);
    if (mode === "manual") manualCount += 1;
    else if (mode === "pixels") pixelsCount += 1;
    else if (mode === "auto") autoCount += 1;
  }

  const parts = [];
  if (manualCount > 0) parts.push(`${manualCount} manuell`);
  if (pixelsCount > 0) parts.push(`${pixelsCount} mit Pixels`);
  if (autoCount > 0) parts.push(`${autoCount} automatisch`);
  return parts.length ? parts.join(", ") : "keine Würfe";
}

function preparePixelsRollDialog(characters, usePixelsForTurn, title = "INI Würfelphase") {
  if (!shouldUseRollDialogForCharacters(characters, usePixelsForTurn)) {
    return false;
  }

  clearPixelsRollDialogRows();
  setPixelsRollDialogTitle(title);
  setPixelsRollDialogStatus(`Würfelquellen: ${describeRollDialogMix(characters, usePixelsForTurn)}.`);
  for (const character of characters) {
    createPixelsRollDialogRow(character, getCharacterRollMode(character, usePixelsForTurn));
  }
  openPixelsRollDialog();
  return true;
}

async function requestRollDialogNumberInput(character, options) {
  const refs = rollDialogState.rowsByCharacterId.get(character.id);
  if (!refs) {
    return null;
  }

  updatePixelsRollDialogRow(character.id, {
    stateClass: "waiting",
    status: options.status,
    detail: options.detail,
  });

  refs.controlsEl.innerHTML = "";
  const fieldWrap = document.createElement("label");
  fieldWrap.className = "pixels-roll-input-wrap";
  fieldWrap.textContent = options.label;

  const input = document.createElement("input");
  input.type = "number";
  input.min = String(options.min);
  input.max = String(options.max);
  input.step = "1";
  input.value = options.initialValue === null || options.initialValue === undefined ? "" : String(options.initialValue);
  input.placeholder = `${options.min}-${options.max}`;
  fieldWrap.appendChild(input);

  const submitBtn = document.createElement("button");
  submitBtn.type = "button";
  submitBtn.textContent = options.buttonLabel || "Übernehmen";

  refs.controlsEl.appendChild(fieldWrap);
  refs.controlsEl.appendChild(submitBtn);

  return new Promise((resolve) => {
    const commit = () => {
      const value = Math.round(Number(input.value));
      if (!Number.isFinite(value) || value < options.min || value > options.max) {
        input.focus();
        input.select();
        return;
      }

      refs.controlsEl.innerHTML = "";
      resolve(value);
    };

    submitBtn.addEventListener("click", commit, { once: true });
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        commit();
      }
    });
    requestAnimationFrame(() => {
      input.focus();
      input.select();
    });
  });
}

function openPixelsRollDialog() {
  if (!pixelsRollDialogEl || pixelsRollDialogEl.open) {
    return;
  }
  if (typeof pixelsRollDialogEl.showModal === "function") {
    pixelsRollDialogEl.showModal();
  } else {
    pixelsRollDialogEl.setAttribute("open", "open");
  }
}

function closePixelsRollDialog() {
  if (!pixelsRollDialogEl || !pixelsRollDialogEl.open) {
    return;
  }
  clearPixelsRollDialogRows();
  pixelsRollDialogEl.close();
}

function normalizePixelsMode(value) {
  if (value === PIXELS_MODE.PC_SET_3 || value === PIXELS_MODE.PC_SINGLE_3X || value === PIXELS_MODE.SHARED_SET_3) {
    return value;
  }
  return PIXELS_MODE.PC_SET_3;
}

function getPcCharactersInRosterOrder() {
  return state.characters.filter((character) => character.type === "PC");
}

function ensureCharacterPixelsAssignment(characterId) {
  const key = String(characterId);
  const existing = pixels.assignmentsByCharacterId[key] || {};
  if (!Array.isArray(existing.set3) || existing.set3.length !== 3) {
    existing.set3 = [existing.set3?.[0] ?? null, existing.set3?.[1] ?? null, existing.set3?.[2] ?? null];
  }
  existing.single = existing.single ?? null;
  pixels.assignmentsByCharacterId[key] = existing;
  return existing;
}

function ensureRememberedCharacterPixelsAssignment(characterId) {
  const key = String(characterId);
  const existing = normalizeRememberedAssignment(pixels.rememberedAssignmentsByCharacterId[key]);
  pixels.rememberedAssignmentsByCharacterId[key] = existing;
  return existing;
}

function getConnectedPixelLabel(pixel) {
  if (!pixel) {
    return "nicht verbunden";
  }
  const labelCandidates = [pixel.name, pixel.deviceName, pixel.device?.name, pixel.id];
  for (const candidate of labelCandidates) {
    if (typeof candidate === "string" && candidate.trim()) {
      return candidate.trim();
    }
  }
  return "Verbunden";
}

function getRememberedPixelLabel(rememberedRef) {
  return rememberedRef?.label ? `${rememberedRef.label} (gemerkt)` : "nicht verbunden";
}

function rememberPixelsAssignment(characterId, assignment) {
  const remembered = ensureRememberedCharacterPixelsAssignment(characterId);
  remembered.single = createRememberedPixelRef(assignment.single);
  remembered.set3 = Array.from({ length: 3 }, (_, index) => createRememberedPixelRef(assignment.set3?.[index]));
}

function debugPixels(message, details = null) {
  if (!PIXELS_DEBUG) {
    return;
  }
  if (details === null || details === undefined) {
    console.debug("[PixelsDebug]", message);
    return;
  }
  console.debug("[PixelsDebug]", message, details);
}

function describePixelObject(pixel) {
  if (!pixel || typeof pixel !== "object") {
    return { type: typeof pixel };
  }

  const keys = Object.keys(pixel).sort();
  const proto = Object.getPrototypeOf(pixel);
  const protoKeys = proto ? Object.getOwnPropertyNames(proto).sort() : [];
  return {
    label: getConnectedPixelLabel(pixel),
    keys,
    protoKeys,
    face: pixel.face ?? null,
    value: pixel.value ?? null,
    currentFace: pixel.currentFace ?? null,
    rollState: pixel.rollState ?? null,
    hasAddEventListener: typeof pixel.addEventListener === "function",
    hasOn: typeof pixel.on === "function",
    hasOnRoll: typeof pixel.onRoll === "function",
  };
}

function getFirstMissingPixelIndex(pixelList) {
  if (!Array.isArray(pixelList)) {
    return 0;
  }
  for (let index = 0; index < 3; index += 1) {
    if (!pixelList[index]) {
      return index;
    }
  }
  return -1;
}

function getPixelsModeDescription(mode) {
  if (mode === PIXELS_MODE.PC_SINGLE_3X) {
    return "Jeder SC würfelt mit einem eigenen Pixel dreimal nacheinander.";
  }
  if (mode === PIXELS_MODE.SHARED_SET_3) {
    return "Alle SC würfeln mit demselben 3er-Set in Reihenfolge des Charaktere/NSC-Panels.";
  }
  return "Jeder SC nutzt ein eigenes Set aus drei spezifischen Pixels (ein Würfel pro W6).";
}

function hasAnyConnectedPixels() {
  if (pixels.sharedSet.some((pixel) => Boolean(pixel))) {
    return true;
  }
  return Object.values(pixels.assignmentsByCharacterId).some((assignment) => {
    if (!assignment || typeof assignment !== "object") {
      return false;
    }
    if (assignment.single) {
      return true;
    }
    return Array.isArray(assignment.set3) && assignment.set3.some((pixel) => Boolean(pixel));
  });
}

function getRollPixelsForCharacter(characterId) {
  const assignment = ensureCharacterPixelsAssignment(characterId);
  if (pixels.mode === PIXELS_MODE.PC_SINGLE_3X) {
    return assignment.single ? [assignment.single, assignment.single, assignment.single] : [];
  }
  if (pixels.mode === PIXELS_MODE.SHARED_SET_3) {
    return pixels.sharedSet.filter((pixel) => Boolean(pixel));
  }
  return Array.isArray(assignment.set3) ? assignment.set3.filter((pixel) => Boolean(pixel)) : [];
}

function updatePixelsControls() {
  const supported = supportsPixels();
  const connected = hasAnyConnectedPixels();
  const available = connected || pixels.simulated;

  if (!available && pixels.useForRolls) {
    pixels.useForRolls = false;
  }

  if (!available) {
    if (!supported) {
      updatePixelsStatus("Pixels: Web Bluetooth ist in diesem Browser nicht verfügbar.");
    } else if (!pixels.lastStatus.startsWith("Pixels: nicht verbunden")) {
      updatePixelsStatus("Pixels: nicht verbunden (Chromium-Browser mit Web Bluetooth erforderlich).");
    }
  } else if (pixels.useForRolls) {
    updatePixelsStatus(`Pixels: aktiv (${getPixelsModeDescription(pixels.mode)})`);
  } else {
    updatePixelsStatus("Pixels: verbunden (manueller Modus).");
  }

  if (pixelsUseForRollsEl) {
    pixelsUseForRollsEl.checked = pixels.useForRolls;
    pixelsUseForRollsEl.disabled = !available;
  }
  if (pixelsSimulatedEl) {
    pixelsSimulatedEl.checked = pixels.simulated;
    pixelsSimulatedEl.disabled = pixels.loading;
  }
  if (pixelsModeSelectEl) {
    pixelsModeSelectEl.value = normalizePixelsMode(pixels.mode);
  }
  if (openPixelsSettingsBtn) {
    openPixelsSettingsBtn.disabled = pixels.loading;
  }
}

function openPixelsSettingsDialog() {
  if (!pixelsSettingsDialogEl) {
    return;
  }
  renderPixelsSettingsDialog();
  if (typeof pixelsSettingsDialogEl.showModal === "function") {
    pixelsSettingsDialogEl.showModal();
  } else {
    pixelsSettingsDialogEl.setAttribute("open", "open");
  }
}

function closePixelsSettingsDialog() {
  if (!pixelsSettingsDialogEl) {
    return;
  }
  pixelsSettingsDialogEl.close();
}

async function ensurePixelsSdkLoaded() {
  if (pixels.sdk) {
    return pixels.sdk;
  }
  pixels.sdk = await import("https://cdn.jsdelivr.net/npm/@systemic-games/pixels-web-connect/+esm");
  return pixels.sdk;
}

function supportsPixelsAutoReconnect() {
  return supportsPixels() && typeof navigator?.bluetooth?.getDevices === "function";
}

function isUsablePixelsObject(pixel) {
  return Boolean(
    pixel &&
      typeof pixel === "object" &&
      (typeof pixel.addEventListener === "function" || typeof pixel.on === "function" || typeof pixel.onRoll === "function")
  );
}

async function reconnectPixelFromRememberedRef(rememberedRef) {
  const normalizedRef = normalizeRememberedPixelRef(rememberedRef);
  if (!normalizedRef) {
    return null;
  }

  const sdk = await ensurePixelsSdkLoaded();
  const pixel = typeof sdk?.getPixel === "function" ? await sdk.getPixel(normalizedRef.deviceId) : null;
  if (!pixel) {
    return null;
  }

  if (typeof sdk?.repeatConnect === "function") {
    await sdk.repeatConnect(pixel);
  } else if (typeof pixel.connect === "function") {
    await pixel.connect();
  }

  return isUsablePixelsObject(pixel) ? pixel : null;
}

async function reconnectRememberedPixels() {
  if (pixels.reconnecting || !supportsPixelsAutoReconnect()) {
    return;
  }

  const rememberedAssignments = captureRememberedPixelsSettings();
  const hasRememberedPixels =
    Object.keys(rememberedAssignments.rememberedAssignmentsByCharacterId).length > 0 ||
    rememberedAssignments.rememberedSharedSet.some((entry) => Boolean(entry));
  if (!hasRememberedPixels) {
    return;
  }

  pixels.reconnecting = true;
  pixels.loading = true;
  updatePixelsStatus("Pixels: versuche gemerkte Würfel automatisch wieder zu verbinden...");
  updatePixelsControls();

  try {
    let reconnectedCount = 0;

    for (const [characterId, rememberedAssignment] of Object.entries(rememberedAssignments.rememberedAssignmentsByCharacterId)) {
      const assignment = ensureCharacterPixelsAssignment(characterId);
      const remembered = ensureRememberedCharacterPixelsAssignment(characterId);

      if (rememberedAssignment.single) {
        const pixel = await reconnectPixelFromRememberedRef(rememberedAssignment.single);
        assignment.single = pixel;
        remembered.single = normalizeRememberedPixelRef(rememberedAssignment.single);
        if (pixel) {
          reconnectedCount += 1;
        }
      }

      assignment.set3 = await Promise.all(
        Array.from({ length: 3 }, async (_, index) => reconnectPixelFromRememberedRef(rememberedAssignment.set3[index]))
      );
      remembered.set3 = Array.from({ length: 3 }, (_, index) => normalizeRememberedPixelRef(rememberedAssignment.set3[index]));
      reconnectedCount += assignment.set3.filter((pixel) => Boolean(pixel)).length;
    }

    pixels.sharedSet = await Promise.all(
      Array.from({ length: 3 }, async (_, index) => reconnectPixelFromRememberedRef(rememberedAssignments.rememberedSharedSet[index]))
    );
    pixels.rememberedSharedSet = Array.from({ length: 3 }, (_, index) =>
      normalizeRememberedPixelRef(rememberedAssignments.rememberedSharedSet[index])
    );
    reconnectedCount += pixels.sharedSet.filter((pixel) => Boolean(pixel)).length;

    if (reconnectedCount > 0) {
      updatePixelsStatus(`Pixels: ${reconnectedCount} gemerkte Würfel automatisch wieder verbunden.`);
    } else {
      updatePixelsStatus("Pixels: keine gemerkten Würfel automatisch wieder verbunden.");
    }
  } catch (error) {
    const message = error && typeof error.message === "string" ? error.message : String(error);
    updatePixelsStatus(`Pixels: Auto-Reconnect fehlgeschlagen (${message}).`);
  } finally {
    pixels.loading = false;
    pixels.reconnecting = false;
    persistAppState();
    render();
  }
}

async function requestAndConnectSinglePixel(waitMessage) {
  if (!supportsPixels()) {
    throw new Error("Web Bluetooth ist in diesem Browser nicht verfügbar");
  }

  const sdk = await ensurePixelsSdkLoaded();
  updatePixelsStatus(waitMessage);
  const selectedPixel = await sdk.requestPixel();
  if (!selectedPixel) {
    throw new Error("kein Würfel ausgewählt");
  }
  if (typeof sdk.repeatConnect === "function") {
    await sdk.repeatConnect(selectedPixel);
  } else if (typeof selectedPixel.connect === "function") {
    await selectedPixel.connect();
  }
  debugPixels("Verbundenes Pixel-Objekt", describePixelObject(selectedPixel));
  return selectedPixel;
}

function getPixelsBlinkColor(sdk) {
  if (sdk?.Color?.red) {
    return sdk.Color.red;
  }
  if (sdk?.Color?.Red) {
    return sdk.Color.Red;
  }
  throw new Error("keine unterstützte Pixels-Farbe gefunden");
}

async function blinkPixelsDice(targetPixels, targetLabel) {
  const connectedPixels = Array.from(
    new Set((Array.isArray(targetPixels) ? targetPixels : []).filter((pixel) => Boolean(pixel)))
  );
  if (!connectedPixels.length) {
    throw new Error("keine verbundenen Pixels vorhanden");
  }

  const sdk = await ensurePixelsSdkLoaded();
  const blinkColor = getPixelsBlinkColor(sdk);
  let blinkCount = 0;

  for (const pixel of connectedPixels) {
    if (typeof pixel.blink !== "function") {
      throw new Error(`${getConnectedPixelLabel(pixel)} unterstützt keine LED-Steuerung`);
    }
    await pixel.blink(blinkColor);
    blinkCount += 1;
  }

  const suffix = blinkCount === 1 ? "" : "n";
  updatePixelsStatus(`Pixels: LED-Test für ${targetLabel} auf ${blinkCount} Würfel${suffix} gestartet.`);
  logEvent(`Pixels: LED-Test für ${targetLabel} auf ${blinkCount} Würfel${suffix} gestartet.`);
}

async function runPixelsLedTest(targetPixels, targetLabel) {
  pixels.loading = true;
  updatePixelsControls();
  renderPixelsSettingsDialog();
  try {
    await blinkPixelsDice(targetPixels, targetLabel);
  } catch (error) {
    const message = error && typeof error.message === "string" ? error.message : String(error);
    updatePixelsStatus(`Pixels: LED-Test fehlgeschlagen (${message}).`);
    logEvent(`Pixels-LED-Test fehlgeschlagen: ${message}.`);
  } finally {
    pixels.loading = false;
    updatePixelsControls();
    renderPixelsSettingsDialog();
  }
}

function appendPixelsLedTestButton(actionsEl, targetPixels, targetLabel) {
  const ledBtn = document.createElement("button");
  ledBtn.type = "button";
  ledBtn.className = "ghost";
  ledBtn.textContent = "LED testen";
  ledBtn.disabled = pixels.loading || !targetPixels.some((pixel) => Boolean(pixel));
  ledBtn.addEventListener("click", async () => {
    await runPixelsLedTest(targetPixels, targetLabel);
  });
  actionsEl.appendChild(ledBtn);
}

function uniquePixelsFromAssignments() {
  const unique = new Set();
  for (const assignment of Object.values(pixels.assignmentsByCharacterId)) {
    if (!assignment || typeof assignment !== "object") {
      continue;
    }
    if (assignment.single) {
      unique.add(assignment.single);
    }
    if (Array.isArray(assignment.set3)) {
      for (const pixel of assignment.set3) {
        if (pixel) {
          unique.add(pixel);
        }
      }
    }
  }
  for (const pixel of pixels.sharedSet) {
    if (pixel) {
      unique.add(pixel);
    }
  }
  return Array.from(unique);
}

async function disconnectAllPixels() {
  const allPixels = uniquePixelsFromAssignments();
  for (const pixel of allPixels) {
    try {
      if (typeof pixel.disconnect === "function") {
        await pixel.disconnect();
      }
    } catch {
      // Trennungsfehler ignorieren.
    }
  }
  pixels.assignmentsByCharacterId = {};
  pixels.sharedSet = [null, null, null];
  pixels.rememberedAssignmentsByCharacterId = {};
  pixels.rememberedSharedSet = [null, null, null];
  if (!pixels.simulated) {
    pixels.useForRolls = false;
  }
  updatePixelsStatus("Pixels: getrennt.");
  logEvent("Alle Pixels-Würfel getrennt.");
  persistAppState();
  renderCombatLog();
  updatePixelsControls();
}

async function connectSet3ForCharacter(character) {
  if (!character) {
    return;
  }
  const assignment = ensureCharacterPixelsAssignment(character.id);
  const nextIndex = getFirstMissingPixelIndex(assignment.set3);
  if (nextIndex < 0) {
    updatePixelsStatus(`Pixels: Set für ${character.name} ist bereits vollständig verbunden.`);
    renderPixelsSettingsDialog();
    return;
  }
  pixels.loading = true;
  updatePixelsControls();
  try {
    const pixel = await requestAndConnectSinglePixel(
      `Pixels: SC ${character.name}, Set-Würfel ${nextIndex + 1}/3 auswählen...`
    );
    assignment.single = null;
    assignment.set3[nextIndex] = pixel;
    rememberPixelsAssignment(character.id, assignment);
    const remainingIndex = getFirstMissingPixelIndex(assignment.set3);
    if (remainingIndex < 0) {
      logEvent(`Pixels: Set (3) für ${character.name} vollständig verbunden.`);
      updatePixelsStatus(`Pixels: Set für ${character.name} vollständig verbunden.`);
    } else {
      logEvent(`Pixels: Set-Würfel ${nextIndex + 1}/3 für ${character.name} verbunden.`);
      updatePixelsStatus(`Pixels: ${character.name}, bitte nun Würfel ${remainingIndex + 1}/3 verbinden.`);
    }
    persistAppState();
    renderCombatLog();
  } catch (error) {
    const message = error && typeof error.message === "string" ? error.message : String(error);
    updatePixelsStatus(`Pixels: Verbindung fehlgeschlagen (${message}).`);
    logEvent(`Pixels-Verbindung fehlgeschlagen: ${message}.`);
    persistAppState();
    renderCombatLog();
  } finally {
    pixels.loading = false;
    updatePixelsControls();
    renderPixelsSettingsDialog();
  }
}

async function connectSingleForCharacter(character) {
  if (!character) {
    return;
  }
  const assignment = ensureCharacterPixelsAssignment(character.id);
  pixels.loading = true;
  updatePixelsControls();
  try {
    assignment.set3 = [null, null, null];
    assignment.single = await requestAndConnectSinglePixel(`Pixels: SC ${character.name}, Einzelwürfel auswählen...`);
    rememberPixelsAssignment(character.id, assignment);
    logEvent(`Pixels: Einzelwürfel für ${character.name} verbunden.`);
    updatePixelsStatus(`Pixels: Einzelwürfel für ${character.name} verbunden.`);
    persistAppState();
    renderCombatLog();
  } catch (error) {
    const message = error && typeof error.message === "string" ? error.message : String(error);
    updatePixelsStatus(`Pixels: Verbindung fehlgeschlagen (${message}).`);
    logEvent(`Pixels-Verbindung fehlgeschlagen: ${message}.`);
    persistAppState();
    renderCombatLog();
  } finally {
    pixels.loading = false;
    updatePixelsControls();
    renderPixelsSettingsDialog();
  }
}

async function connectSharedSet3() {
  const nextIndex = getFirstMissingPixelIndex(pixels.sharedSet);
  if (nextIndex < 0) {
    updatePixelsStatus("Pixels: gemeinsames 3er-Set ist bereits vollständig verbunden.");
    renderPixelsSettingsDialog();
    return;
  }
  pixels.loading = true;
  updatePixelsControls();
  try {
    const pixel = await requestAndConnectSinglePixel(`Pixels: Gemeinsames Set, Würfel ${nextIndex + 1}/3 auswählen...`);
    pixels.sharedSet[nextIndex] = pixel;
    pixels.rememberedSharedSet[nextIndex] = createRememberedPixelRef(pixel);
    const remainingIndex = getFirstMissingPixelIndex(pixels.sharedSet);
    if (remainingIndex < 0) {
      logEvent("Pixels: Gemeinsames 3er-Set vollständig verbunden.");
      updatePixelsStatus("Pixels: gemeinsames 3er-Set vollständig verbunden.");
    } else {
      logEvent(`Pixels: gemeinsamer Set-Würfel ${nextIndex + 1}/3 verbunden.`);
      updatePixelsStatus(`Pixels: bitte nun Würfel ${remainingIndex + 1}/3 des gemeinsamen Sets verbinden.`);
    }
    persistAppState();
    renderCombatLog();
  } catch (error) {
    const message = error && typeof error.message === "string" ? error.message : String(error);
    updatePixelsStatus(`Pixels: Verbindung fehlgeschlagen (${message}).`);
    logEvent(`Pixels-Verbindung fehlgeschlagen: ${message}.`);
    persistAppState();
    renderCombatLog();
  } finally {
    pixels.loading = false;
    updatePixelsControls();
    renderPixelsSettingsDialog();
  }
}

function renderPixelsAssignmentRow(character) {
  if (!pixelsAssignmentListEl) {
    return;
  }
  const assignment = ensureCharacterPixelsAssignment(character.id);
  const rememberedAssignment = ensureRememberedCharacterPixelsAssignment(character.id);
  const row = document.createElement("div");
  row.className = "pixels-assignment-row";

  const head = document.createElement("div");
  head.className = "pixels-assignment-row-head";
  const name = document.createElement("strong");
  name.textContent = character.name;
  head.appendChild(name);
  row.appendChild(head);

  const actions = document.createElement("div");
  actions.className = "actions";

  if (pixels.mode === PIXELS_MODE.PC_SINGLE_3X) {
    const status = document.createElement("span");
    status.className = "hint";
    status.textContent = `Pixel: ${
      assignment.single ? getConnectedPixelLabel(assignment.single) : getRememberedPixelLabel(rememberedAssignment.single)
    }`;
    row.appendChild(status);

    const connectBtn = document.createElement("button");
    connectBtn.type = "button";
    connectBtn.className = "ghost";
    connectBtn.textContent = "Pixel verbinden";
    connectBtn.disabled = pixels.loading || !supportsPixels();
    connectBtn.addEventListener("click", async () => {
      await connectSingleForCharacter(character);
    });
    actions.appendChild(connectBtn);

    appendPixelsLedTestButton(actions, assignment.single ? [assignment.single] : [], character.name);

    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "ghost";
    clearBtn.textContent = "Zuweisung löschen";
    clearBtn.disabled = pixels.loading || (!assignment.single && !rememberedAssignment.single);
    clearBtn.addEventListener("click", () => {
      assignment.single = null;
      rememberedAssignment.single = null;
      persistAppState();
      updatePixelsControls();
      renderPixelsSettingsDialog();
    });
    actions.appendChild(clearBtn);
    row.appendChild(actions);
    pixelsAssignmentListEl.appendChild(row);
    return;
  }

  const labels = assignment.set3
    .map((pixel, index) => {
      const rememberedRef = rememberedAssignment.set3[index];
      return `W${index + 1}: ${pixel ? getConnectedPixelLabel(pixel) : getRememberedPixelLabel(rememberedRef)}`;
    })
    .join(" | ");
  const status = document.createElement("span");
  status.className = "hint";
  status.textContent = labels;
  row.appendChild(status);

  const connectBtn = document.createElement("button");
  connectBtn.type = "button";
  connectBtn.className = "ghost";
  const nextSetIndex = getFirstMissingPixelIndex(assignment.set3);
  connectBtn.textContent = nextSetIndex < 0 ? "Set vollständig" : `W${nextSetIndex + 1} verbinden`;
  connectBtn.disabled = pixels.loading || !supportsPixels();
  connectBtn.addEventListener("click", async () => {
    await connectSet3ForCharacter(character);
  });
  actions.appendChild(connectBtn);

  appendPixelsLedTestButton(actions, assignment.set3, character.name);

  const clearBtn = document.createElement("button");
  clearBtn.type = "button";
  clearBtn.className = "ghost";
  clearBtn.textContent = "Zuweisung löschen";
  clearBtn.disabled =
    pixels.loading || (!assignment.set3.some((pixel) => Boolean(pixel)) && !rememberedAssignment.set3.some((entry) => Boolean(entry)));
  clearBtn.addEventListener("click", () => {
    assignment.set3 = [null, null, null];
    rememberedAssignment.set3 = [null, null, null];
    persistAppState();
    updatePixelsControls();
    renderPixelsSettingsDialog();
  });
  actions.appendChild(clearBtn);
  row.appendChild(actions);
  pixelsAssignmentListEl.appendChild(row);
}

function renderPixelsSettingsDialog() {
  if (!pixelsAssignmentListEl) {
    return;
  }
  const validPcIds = new Set(getPcCharactersInRosterOrder().map((character) => String(character.id)));
  for (const key of Object.keys(pixels.assignmentsByCharacterId)) {
    if (!validPcIds.has(key)) {
      delete pixels.assignmentsByCharacterId[key];
    }
  }
  for (const key of Object.keys(pixels.rememberedAssignmentsByCharacterId)) {
    if (!validPcIds.has(key)) {
      delete pixels.rememberedAssignmentsByCharacterId[key];
    }
  }
  if (!Array.isArray(pixels.sharedSet) || pixels.sharedSet.length !== 3) {
    pixels.sharedSet = [pixels.sharedSet?.[0] ?? null, pixels.sharedSet?.[1] ?? null, pixels.sharedSet?.[2] ?? null];
  }
  if (!Array.isArray(pixels.rememberedSharedSet) || pixels.rememberedSharedSet.length !== 3) {
    pixels.rememberedSharedSet = [
      normalizeRememberedPixelRef(pixels.rememberedSharedSet?.[0]),
      normalizeRememberedPixelRef(pixels.rememberedSharedSet?.[1]),
      normalizeRememberedPixelRef(pixels.rememberedSharedSet?.[2]),
    ];
  }
  if (pixelsModeSelectEl) {
    pixelsModeSelectEl.value = normalizePixelsMode(pixels.mode);
  }
  if (pixelsModeHelpEl) {
    pixelsModeHelpEl.textContent = getPixelsModeDescription(pixels.mode);
  }
  if (pixelsUseForRollsEl) {
    pixelsUseForRollsEl.checked = pixels.useForRolls;
  }
  if (pixelsSimulatedEl) {
    pixelsSimulatedEl.checked = pixels.simulated;
  }

  pixelsAssignmentListEl.innerHTML = "";
  const pcCharacters = getPcCharactersInRosterOrder();
  if (pixels.mode === PIXELS_MODE.SHARED_SET_3) {
    const row = document.createElement("div");
    row.className = "pixels-assignment-row";

    const head = document.createElement("div");
    head.className = "pixels-assignment-row-head";
    const title = document.createElement("strong");
    title.textContent = "Gemeinsames Set";
    head.appendChild(title);
    row.appendChild(head);

    const status = document.createElement("span");
    status.className = "hint";
    status.textContent = pixels.sharedSet
      .map((pixel, index) =>
        `W${index + 1}: ${pixel ? getConnectedPixelLabel(pixel) : getRememberedPixelLabel(pixels.rememberedSharedSet[index])}`
      )
      .join(" | ");
    row.appendChild(status);

    const orderHint = document.createElement("span");
    orderHint.className = "hint";
    orderHint.textContent = `Reihenfolge: ${pcCharacters.length ? pcCharacters.map((character) => character.name).join(" → ") : "Keine SC vorhanden."}`;
    row.appendChild(orderHint);

    const actions = document.createElement("div");
    actions.className = "actions";
    const connectBtn = document.createElement("button");
    connectBtn.type = "button";
    connectBtn.className = "ghost";
    const nextSharedIndex = getFirstMissingPixelIndex(pixels.sharedSet);
    connectBtn.textContent = nextSharedIndex < 0 ? "Set vollständig" : `W${nextSharedIndex + 1} verbinden`;
    connectBtn.disabled = pixels.loading || !supportsPixels();
    connectBtn.addEventListener("click", async () => {
      await connectSharedSet3();
    });
    actions.appendChild(connectBtn);

    appendPixelsLedTestButton(actions, pixels.sharedSet, "Gemeinsames Set");

    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "ghost";
    clearBtn.textContent = "Set löschen";
    clearBtn.disabled =
      pixels.loading || (!pixels.sharedSet.some((pixel) => Boolean(pixel)) && !pixels.rememberedSharedSet.some((entry) => Boolean(entry)));
    clearBtn.addEventListener("click", () => {
      pixels.sharedSet = [null, null, null];
      pixels.rememberedSharedSet = [null, null, null];
      persistAppState();
      updatePixelsControls();
      renderPixelsSettingsDialog();
    });
    actions.appendChild(clearBtn);
    row.appendChild(actions);
    pixelsAssignmentListEl.appendChild(row);
    return;
  }

  if (!pcCharacters.length) {
    const hint = document.createElement("p");
    hint.className = "hint";
    hint.textContent = "Keine SC vorhanden.";
    pixelsAssignmentListEl.appendChild(hint);
    return;
  }
  for (const character of pcCharacters) {
    renderPixelsAssignmentRow(character);
  }
}

function delay(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function readFaceFromPixelsEvent(payload) {
  if (Number.isFinite(Number(payload))) {
    return clamp(Math.round(Number(payload)), 1, 6);
  }

  const candidates = [
    payload?.face,
    payload?.value,
    payload?.currentFace,
    payload?.faceIndex !== undefined ? Number(payload.faceIndex) + 1 : null,
    payload?.detail?.face,
    payload?.detail?.value,
    payload?.detail?.currentFace,
    payload?.detail?.faceIndex !== undefined ? Number(payload.detail.faceIndex) + 1 : null,
    payload?.rollState?.face,
    payload?.rollState?.currentFace,
    payload?.data?.face,
    payload?.data?.value,
    payload?.data?.currentFace,
    payload?.target?.face,
    payload?.target?.value,
    payload?.target?.rollState?.face,
    payload?.target?.rollState?.currentFace,
    payload?.target?.currentFace,
  ];
  for (const candidate of candidates) {
    if (Number.isFinite(Number(candidate))) {
      return clamp(Math.round(Number(candidate)), 1, 6);
    }
  }

  return null;
}

function waitForSinglePixelsD6(pixel, timeoutMs = 30000) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const detachFns = [];
    let sampleIntervalId = null;
    let sampleTimeoutId = null;
    let sawRollEvent = false;
    let sawResolvableFace = false;
    const pixelLabel = getConnectedPixelLabel(pixel);

    const timeoutId = setTimeout(() => {
      cleanup();
      const reason = sawRollEvent
        ? `Zeitüberschreitung: ${pixelLabel} hat ein Roll-Event geliefert, aber kein Gesicht erkannt`
        : `Zeitüberschreitung: ${pixelLabel} hat kein Roll-Event geliefert`;
      debugPixels("Timeout beim Warten auf Wurf", {
        pixel: describePixelObject(pixel),
        sawRollEvent,
        sawResolvableFace,
      });
      reject(new Error(reason));
    }, timeoutMs);

    const stopSampling = () => {
      if (sampleIntervalId !== null) {
        clearInterval(sampleIntervalId);
        sampleIntervalId = null;
      }
      if (sampleTimeoutId !== null) {
        clearTimeout(sampleTimeoutId);
        sampleTimeoutId = null;
      }
    };

    const resolveIfFaceAvailable = (payload) => {
      const face = readFaceFromPixelsEvent(payload);
      debugPixels("Pruefe Payload auf Face", {
        pixel: pixelLabel,
        face,
        payload,
        currentState: describePixelObject(pixel),
      });
      if (!face || face < 1 || face > 6) {
        return false;
      }
      sawResolvableFace = true;
      cleanup();
      resolve(face);
      return true;
    };

    const startSampling = (eventOrFace) => {
      stopSampling();
      const buildPayload = () => ({
        ...((eventOrFace && typeof eventOrFace === "object") ? eventOrFace : {}),
        face: pixel?.face,
        value: pixel?.value,
        currentFace: pixel?.currentFace,
        rollState: pixel?.rollState,
        data: pixel?.data,
        target: eventOrFace?.target ?? pixel ?? null,
      });

      if (resolveIfFaceAvailable(buildPayload())) {
        return;
      }

      sampleIntervalId = setInterval(() => {
        resolveIfFaceAvailable(buildPayload());
      }, 120);
      sampleTimeoutId = setTimeout(() => {
        stopSampling();
      }, 2000);
    };

    const onRoll = (eventOrFace) => {
      sawRollEvent = true;
      debugPixels("Roll-Event empfangen", {
        pixel: pixelLabel,
        event: eventOrFace,
        currentState: describePixelObject(pixel),
      });
      if (resolveIfFaceAvailable(eventOrFace)) {
        return;
      }
      startSampling(eventOrFace);
    };

    function cleanup() {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeoutId);
      stopSampling();
      for (const detach of detachFns) {
        if (typeof detach === "function") {
          detach();
        }
      }
    }

    if (typeof pixel.addEventListener === "function" && typeof pixel.removeEventListener === "function") {
      pixel.addEventListener("roll", onRoll);
      detachFns.push(() => pixel.removeEventListener("roll", onRoll));
      for (const eventName of ["status", "message", "messageReceived", "messageRollState"]) {
        const debugListener = (event) => {
          debugPixels(`Zusatz-Event ${eventName}`, {
            pixel: pixelLabel,
            event,
            currentState: describePixelObject(pixel),
          });
        };
        pixel.addEventListener(eventName, debugListener);
        detachFns.push(() => pixel.removeEventListener(eventName, debugListener));
      }
    }

    if (typeof pixel.on === "function" && typeof pixel.off === "function") {
      pixel.on("roll", onRoll);
      detachFns.push(() => pixel.off("roll", onRoll));
      for (const eventName of ["status", "message", "messageReceived", "messageRollState"]) {
        const debugListener = (event) => {
          debugPixels(`Zusatz-on-Event ${eventName}`, {
            pixel: pixelLabel,
            event,
            currentState: describePixelObject(pixel),
          });
        };
        pixel.on(eventName, debugListener);
        detachFns.push(() => pixel.off(eventName, debugListener));
      }
    }

    if (typeof pixel.onRoll === "function") {
      const maybeDetach = pixel.onRoll(onRoll);
      if (typeof maybeDetach === "function") {
        detachFns.push(maybeDetach);
      }
    }

    if (!detachFns.length) {
      cleanup();
      reject(new Error(`${pixelLabel} hat keinen Roll-Listener`));
      return;
    }
  });
}

function hasThreeDistinctPixels(rollPixels) {
  if (!Array.isArray(rollPixels) || rollPixels.length < 3) {
    return false;
  }
  return new Set(rollPixels.slice(0, 3)).size === 3;
}

function promptForRollValue(message, initialValue, min, max, fallbackValue) {
  if (typeof window === "undefined" || typeof window.prompt !== "function") {
    return fallbackValue;
  }

  while (true) {
    const rawValue = window.prompt(message, initialValue === null || initialValue === undefined ? "" : String(initialValue));
    if (rawValue === null) {
      return fallbackValue;
    }

    const parsedValue = Math.round(Number(rawValue));
    if (Number.isFinite(parsedValue) && parsedValue >= min && parsedValue <= max) {
      return parsedValue;
    }
  }
}

async function requestManualRollTotal(character) {
  if (rollDialogState.rowsByCharacterId.has(character.id)) {
    return requestRollDialogNumberInput(character, {
      status: "Manueller INI-Wurf ausstehend.",
      detail: "3w6 manuell eintragen und übernehmen.",
      label: "3w6",
      min: 3,
      max: 18,
      initialValue: character.manualRoll ?? "",
      buttonLabel: "Wurf übernehmen",
    });
  }

  return promptForRollValue(`Manueller 3w6-Wurf für ${character.name} eingeben (3-18).`, character.manualRoll ?? "", 3, 18, 10);
}

async function requestManualCriticalBonusRoll(character) {
  if (rollDialogState.rowsByCharacterId.has(character.id)) {
    return requestRollDialogNumberInput(character, {
      status: "Kritischer Erfolg: zusätzlicher W6 ausstehend.",
      detail: "Zusätzlichen W6 für die Initiative eintragen.",
      label: "Krit-W6",
      min: 1,
      max: 6,
      initialValue: 1,
      buttonLabel: "Krit-W6 übernehmen",
    });
  }

  return promptForRollValue(
    `Kritischer Erfolg für ${character.name}: zusätzlichen W6 für Initiative eingeben (1-6).`,
    1,
    1,
    6,
    d6()
  );
}

async function roll1d6WithPixels(characterName, rollPixels) {
  if (pixels.simulated) {
    const statusMessage = `Warte auf Krit-W6 von ${characterName} (simuliert).`;
    setPixelsRollDialogStatus(statusMessage);
    updatePixelsStatus(`Pixels (Sim): ${statusMessage}`);
    await delay(220);
    return {
      face: d6(),
      simulated: true,
    };
  }

  const pixel = Array.isArray(rollPixels) ? rollPixels.find((item) => Boolean(item)) : null;
  if (!pixel) {
    throw new Error("kein Pixel für Krit-W6 zugewiesen");
  }

  const dieLabel = getConnectedPixelLabel(pixel);
  const statusMessage = `Warte auf Krit-W6 von ${characterName} (${dieLabel}).`;
  setPixelsRollDialogStatus(statusMessage);
  updatePixelsStatus(`Pixels: ${statusMessage}`);
  return {
    face: await waitForSinglePixelsD6(pixel),
    simulated: false,
  };
}

async function roll3d6WithPixels(characterName, rollPixels) {
  if (pixels.simulated) {
    const faces = [];
    for (let i = 0; i < 3; i += 1) {
      const statusMessage = `Warte auf Wurf von ${characterName}: w6 ${i + 1}/3 (simuliert).`;
      setPixelsRollDialogStatus(statusMessage);
      updatePixelsStatus(`Pixels (Sim): ${statusMessage}`);
      await delay(220);
      faces.push(d6());
    }

    return {
      faces,
      total: faces[0] + faces[1] + faces[2],
      simulated: true,
    };
  }

  if (!Array.isArray(rollPixels) || rollPixels.length < 3) {
    throw new Error("nicht genügend Pixels zugewiesen");
  }

  if (hasThreeDistinctPixels(rollPixels)) {
    const dieLabels = rollPixels.slice(0, 3).map((pixel, index) => `W${index + 1}: ${getConnectedPixelLabel(pixel)}`).join(" | ");
    const statusMessage = `Warte auf drei Würfe von ${characterName} (${dieLabels}).`;
    setPixelsRollDialogStatus(statusMessage);
    updatePixelsStatus(`Pixels: ${statusMessage}`);
    const faces = await Promise.all(
      rollPixels.slice(0, 3).map((pixel, index) =>
        waitForSinglePixelsD6(pixel).catch((error) => {
          const message = error && typeof error.message === "string" ? error.message : String(error);
          throw new Error(`W${index + 1} ${getConnectedPixelLabel(pixel)}: ${message}`);
        })
      )
    );
    return {
      faces,
      total: faces[0] + faces[1] + faces[2],
      simulated: false,
    };
  }

  const faces = [];
  for (let i = 0; i < 3; i += 1) {
    const pixel = rollPixels[i];
    if (!pixel) {
      throw new Error(`fehlender Würfel ${i + 1}/3`);
    }
    const dieLabel = getConnectedPixelLabel(pixel);
    const statusMessage = `Warte auf Wurf von ${characterName}: w6 ${i + 1}/3 (${dieLabel}).`;
    setPixelsRollDialogStatus(statusMessage);
    updatePixelsStatus(`Pixels: ${statusMessage}`);
    const face = await waitForSinglePixelsD6(pixel);
    faces.push(face);
  }

  return {
    faces,
    total: faces[0] + faces[1] + faces[2],
    simulated: false,
  };
}

async function resolveCharacterRoll(character, usePixelsForTurn) {
  const usesManualRoll = character.type === "PC" && character.useManualRoll;
  if (usesManualRoll) {
    const total = clamp(await requestManualRollTotal(character), 3, 18);
    return {
      total,
      source: "manual",
      faces: null,
      critBonusRoll: total === 18 ? await requestManualCriticalBonusRoll(character) : null,
      manualRollValue: total,
    };
  }

  const canUsePixelsForCharacter = usePixelsForTurn && character.type === "PC";
  if (canUsePixelsForCharacter) {
    try {
      const assignedPixels = getRollPixelsForCharacter(character.id);
      const pixelsRoll = await roll3d6WithPixels(character.name, assignedPixels);
      const critBonusPixelsRoll = pixelsRoll.total === 18 ? await roll1d6WithPixels(character.name, assignedPixels) : null;
      return {
        total: pixelsRoll.total,
        source: pixelsRoll.simulated ? "pixels-sim" : "pixels",
        faces: pixelsRoll.faces,
        critBonusRoll: critBonusPixelsRoll?.face ?? null,
        error: null,
      };
    } catch (error) {
      const message = error && typeof error.message === "string" ? error.message : String(error);
      updatePixelsStatus(`Pixels: Wurf fehlgeschlagen (${message}), nutze Zufallswürfe.`);
      logEvent(`Pixels-Fallback für ${character.name}: ${message}.`);
      const total = roll3d6();
      return {
        total,
        source: "pixels-fallback",
        faces: null,
        critBonusRoll: total === 18 ? d6() : null,
        error: message,
      };
    }
  }

  const total = roll3d6();
  return { total, source: "random", faces: null, critBonusRoll: total === 18 ? d6() : null, error: null };
}

function registerServiceWorker() {
  if (typeof window === "undefined" || !("serviceWorker" in navigator)) {
    return;
  }

  const isLocalhost =
    window.location.hostname === "localhost" ||
    window.location.hostname === "127.0.0.1" ||
    window.location.hostname === "[::1]";

  if (isLocalhost) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .getRegistrations()
        .then((registrations) => Promise.all(registrations.map((registration) => registration.unregister())))
        .catch(() => {});
    });
    return;
  }

  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch((error) => {
      console.warn("Service Worker konnte nicht registriert werden:", error);
    });
  });
}

registerServiceWorker();
loadStateFromStorage();
applyUiSettings(state.uiSettings);
updateVersionBadge();
initSettingsMenusHoverBehavior();
render();
void reconnectRememberedPixels();
