// "nothanks_gs", v1.0, by Eran Hirsch, 2020 (Protected under the GPL3)
// Visit the project at: https://github.com/eranhirsch/nothanks_gs

/** @OnlyCurrentDoc */

/**
 * Game option modifiers (TODO: create settings page)
 */
const OPTIONS = {
  SETUP: {
    /**
     * Remove the X0 cards from the game, out of the initial cards removed?
     * Options: true/false
     * Default: false
     */
    REMOVE_TENS: false,

    /**
     * Should the player order be shuffled when creating a new table?
     * Options: yes/no/ask
     * Default: yes
     */
    SHUFFLE_PLAYERS: "yes",

    /**
     * Pick a specific start player?
     * Options: true/false
     * Default: false
     */
    PICK_START_PLAYER: false,
  },

  /**
   * After taking a card, should the next card be auto-revealed?
   * Options: true/false
   * Default: false
   */
  AUTO_REVEAL_NEXT_CARD: false,

  /**
   * Play with hidden tokens (regular) or exposed?
   * Options: true/false
   * Default: true
   */
  HIDDEN_TOKENS: true,
};

/**
 * Game's rules consts
 */
const MAX_PLAYER_COUNT = 7;
const MIN_CARD = 3;
const MAX_CARD = 35;
const SETUP_CARDS_REMOVED = 9;

/**
 * UI consts
 */
const TOKEN_REPR = "🌑";
const ACTIVE_PLAYER_MARKER = "➡️";
const RANK_MARKERS = ["🏆", "🥈", "🥉", "", "", "", ""];

const BG_COLOR = "#fff3dc";
const DECK_BACK_COLOR = "#1c4587";

const PLAYER1_A1 = "Z6";
const LOCATION_A1 = { DECK: "B2", CARD: "N2", TOKENS: "C13" };

const CELL_DIMENSION = { WIDTH: 21, HEIGHT: 30 };

const FONT_FAMILY = "Francois One";

const TABLE_SHEET_NAME = "Table";

const CARD_SIZE = 10;
const PLAYER_NAME_LENGTH = 4;
const TABLE_DIMENSIONS = { HEIGHT: 17, WIDTH: 56 };

const MSG_REVEAL = "Click DECK to reveal next card";
const MSG_TURN = `Click the CARD to take it and add it to your table
OR
Click HERE to pay 1${TOKEN_REPR} to skip your turn`;
const MSG_END_GAME = `Game Over!
Click HERE to reveal the final score`;

// Our calculation of hotspot width seems to miss by a bit so we can just expand
// it a little by a constant factor to try and fix it.
const HOTSPOT_WIDTH_CORRECTION_FACTOR = 1.05;

/**
 * Mechanism consts
 */
const MUTEX_LOCKOUT_PERIOD_MS = 5000;
const TRANSPARENT_PIXEL_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAsSAAALEgHS3X78AAAADUlEQVQImWP4//8/AwAI/AL+hc2rNAAAAABJRU5ErkJggg==";

////// API HOOKS ///////////////////////////////////////////////////////////////

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("No Thanks")
    .addItem("New Table", "newTable")
    .addToUi();
}

////// USER ACTIONS ////////////////////////////////////////////////////////////

function newTable(players) {
  const file = SpreadsheetApp.getActive();
  const newSheet = file.insertSheet("Creating new table...");
  file.setActiveSheet(newSheet);

  singleEntry(() => {
    try {
      let players = getPlayersForNewTable();

      const ui = SpreadsheetApp.getUi();
      if (
        OPTIONS.SETUP.SHUFFLE_PLAYERS === "yes" ||
        (OPTIONS.SETUP.SHUFFLE_PLAYERS === "ask" &&
          ui.alert(
            "Randomize player order?",
            `Click "Yes" to randomize the play order
                   
                   Or "No" to use the current one:
                   ${players.join(", ")}`,
            ui.ButtonSet.YES_NO,
          ) === ui.Button.YES)
      ) {
        players = shuffle(players);
      }

      renderNewTable(file, newSheet, players);

      newGame(players);

      // Replace the previous table once we are done setting everything up
      const previousSheet = file.getSheetByName(TABLE_SHEET_NAME);
      if (previousSheet != null) {
        file.deleteSheet(previousSheet);
      }
      newSheet.setName(TABLE_SHEET_NAME);
    } catch (err) {
      file.deleteSheet(newSheet);
      throw err;
    }
  });
}

function revealTopCard() {
  singleEntry(() => {
    revealTopCardImpl();
  });
}

function takeCard() {
  singleEntry(() => {
    const card = getCurrentCard();
    if (card == null) {
      throw new Error("No card revealed yet to take");
    }

    resetHotspot("CARD");
    setCurrentCard(null);
    resetHotspot("TOKENS");

    player = getActivePlayer();
    addCardToPlayer(player, card);

    takeTokensFromPool(player);

    const deck = getDeck();
    if (deck != null && deck.length > 0) {
      enableDeck();
    } else {
      setInstructionsMessage(MSG_END_GAME);
      enableHotspot("TOKENS", "revealFinalScore");
    }
  });
}

function revealFinalScore() {
  singleEntry(() => {
    resetHotspot("TOKENS");
    setInstructionsMessage("");
    if (OPTIONS.HIDDEN_TOKENS) {
      revealTokens();
    }
    renderFinalScore();
  });
}

function noThanks() {
  singleEntry(() => {
    const currentCard = getCurrentCard();
    if (currentCard == null) {
      throw new Error("No card revealed yet!");
    }

    addTokenToPool(getActivePlayer());

    advanceActivePlayer();
  });
}

////// LOGICAL ACTIONS /////////////////////////////////////////////////////////

function revealTopCardImpl() {
  const currentCard = getCurrentCard();
  if (currentCard != null) {
    throw new Error(
      "The card '" +
        currentCard +
        "' is still out, someone needs to take it first!",
    );
  }

  const deck = getDeck();
  if (deck == null || deck.length === 0) {
    throw new Error("No more cards in the deck!");
  }

  const { remainingDeck, card } = drawCard(deck);

  enableHotspot("CARD", "takeCard");
  setCurrentCard(card);

  resetHotspot("DECK");
  setDeck(remainingDeck);

  enableHotspot("TOKENS", "noThanks");
  setTokensInPool(0);
}

function getPlayersForNewTable() {
  const ui = SpreadsheetApp.getUi();
  const players = getPlayersFromPreviousTable();
  if (
    players != null &&
    ui.alert(
      "Same Players?",
      `Do you want to use the same list of players as in the previous round?
(${players.join(", ")})`,
      ui.ButtonSet.YES_NO,
    ) === ui.Button.YES
  ) {
    return players;
  }

  return getNewPlayersFromUser();
}

function newGame(players) {
  // Randomize a new deck
  setDeck(newDeck());

  // deal tokens to each player
  dealTokens(players.length);

  setActivePlayer(setStartingPlayer(players));

  enableDeck();
}

function setStartingPlayer(players) {
  if (OPTIONS.SETUP.PICK_START_PLAYER) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      "Choose starting player?",
      `Pick who goes first by entering their number:
    ${players.map((name, i) => i + 1 + " - " + name).join("\n")}
    
    Or select "No" for a random pick`,
      ui.ButtonSet.YES_NO,
    );

    if (response.getSelectedButton() === ui.Button.YES) {
      const startPlayer = parseInt(response.getResponseText(), 10) - 1;

      if (!(startPlayer >= 0 && startPlayer < players.length)) {
        throw new Error(`Invalid start player value: ${startPlayer}`);
      }

      return startPlayer;
    }
  }

  return randInt(players.length - 1);
}

function enableDeck() {
  if (OPTIONS.AUTO_REVEAL_NEXT_CARD) {
    revealTopCardImpl();
  } else {
    enableHotspot("DECK", "revealTopCard");
    setInstructionsMessage(MSG_REVEAL);
  }
}

function drawCard(deck) {
  const randIndex = randInt(deck.length - 1);
  const card = deck[randIndex];

  // Remove the card in-place
  deck.splice(randIndex, 1);

  return {
    card,
    remainingDeck: deck,
  };
}

function newDeck() {
  // Create a new deck (which is just a range of numbers
  let deck = Array.of(...xrange(MIN_CARD, MAX_CARD + 1));

  let cardsRemoved = 0;
  if (OPTIONS.SETUP.REMOVE_TENS) {
    deck = deck.filter((card) => card % 10 !== 0);
    cardsRemoved = MAX_CARD - MIN_CARD + 1 - deck.length;
  }

  // Remove cards from the deck
  for (; cardsRemoved < SETUP_CARDS_REMOVED; cardsRemoved++) {
    deck = drawCard(deck).remainingDeck;
  }

  return deck;
}

function advanceActivePlayer() {
  const activePlayer = getActivePlayer();
  const playerCount = getPlayerCount();
  const nextPlayer = (activePlayer + 1) % playerCount;
  setActivePlayer(nextPlayer);
}

function takeTokensFromPool(player) {
  const pool = getTokensInPool();

  if (pool > 0) {
    if (OPTIONS.HIDDEN_TOKENS) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        getPlayerName(player),
        `Add ${pool}${TOKEN_REPR} to your personal pool`,
        ui.ButtonSet.OK,
      );
    }

    const playerTokensCell = SpreadsheetApp.getActiveSheet()
      .getRange(PLAYER1_A1)
      .offset(player, PLAYER_NAME_LENGTH);
    const current = parseInt(playerTokensCell.getValue(), 10);
    playerTokensCell.setValue(current + pool);
  }

  setTokensInPool(0);
}

function addTokenToPool(player) {
  const playerTokensCell = SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(player, PLAYER_NAME_LENGTH);
  const current = parseInt(playerTokensCell.getValue(), 10);
  if (current === 0) {
    throw new Error(`${getPlayerName(player)} doesn't have any tokens left!`);
  }

  playerTokensCell.setValue(current - 1);

  const pool = getTokensInPool();
  setTokensInPool(pool + 1);
}

function dealTokens(playerCount) {
  const tokens = tokensPerPlayer(playerCount);

  if (OPTIONS.HIDDEN_TOKENS) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "All Players",
      `Add ${tokens}${TOKEN_REPR} to your personal pool`,
      ui.ButtonSet.OK,
    );
  }

  SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(0, PLAYER_NAME_LENGTH, playerCount, 1)
    .setValues(Array(playerCount).fill([tokens]));
}

function revealTokens() {
  SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(0, PLAYER_NAME_LENGTH, getPlayerCount(), 1)
    .setNumberFormat(getPlayerTokensNumberFormat(false /* isHidden */));
}

function tokensPerPlayer(playerCount) {
  switch (playerCount) {
    case 3:
    case 4:
    case 5:
      return 11;
    case 6:
      return 9;
    case 7:
      return 7;
    default:
      throw new Error(`Unsupported player count ${playerCount}`);
  }
}

////// STATE MANAGEMENT ////////////////////////////////////////////////////////

function getDeck() {
  const { deck } = PropertiesService.getDocumentProperties().getProperties();
  if (deck == null || deck === "") {
    return null;
  }

  return JSON.parse(deck);
}

function setDeck(deck) {
  const cardRange = SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.DECK)
    .offset(0, 0, CARD_SIZE, CARD_SIZE);

  if (deck == null || deck.length === 0) {
    PropertiesService.getDocumentProperties().deleteProperty("deck");

    cardRange
      .breakApart()
      .clear()
      .setBackground(BG_COLOR)
      .setBorder(false, false, false, false, false, false);
    return;
  }

  const serialized = JSON.stringify(deck);
  PropertiesService.getDocumentProperties().setProperty("deck", serialized);

  renderDeck(cardRange);
}

function getCurrentCard() {
  const currentCardA1 = SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.CARD)
    .offset(1, 1)
    .getA1Notation();
  const cardStr = getCellValue(currentCardA1);
  return cardStr != null ? parseInt(cardStr, 10) : null;
}

function setCurrentCard(cardVal) {
  const cardRange = SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.CARD)
    .offset(0, 0, CARD_SIZE, CARD_SIZE);

  if (cardVal == null) {
    cardRange
      .breakApart()
      .clear()
      .setBackground(BG_COLOR)
      .setBorder(false, false, false, false, false, false);
    return;
  }

  renderCurrentCard(cardRange, cardVal);
}

function getTokensInPool() {
  const tokenStr = getCellValue(LOCATION_A1.TOKENS);

  if (tokenStr === MSG_TURN) {
    // We show this message when there are 0 tokens
    return 0;
  }

  if (
    tokenStr != null &&
    tokenStr.match(new RegExp(`^${TOKEN_REPR}+$`, "gu"))
  ) {
    return tokenStr.length / TOKEN_REPR.length;
  }

  // We use the current token pool for messaging too, but in those cases the
  // actual tokens are always null (and could be coalesced to 0).
  return null;
}

function setTokensInPool(tokens) {
  if (tokens === 0) {
    setInstructionsMessage(MSG_TURN);
    return;
  }

  SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.TOKENS)
    // We want to fit as many tokens as possible in the box for any token count
    .setFontSize(
      tokens <= 24
        ? 39 - 3 * Math.ceil(Math.max(0, tokens - 16) / 2)
        : tokens <= 39
        ? 25
        : 21,
    )
    // If the value is a number we create a string made up of the token icons
    .setValue(TOKEN_REPR.repeat(tokens));
}

function setInstructionsMessage(message) {
  SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.TOKENS)
    .setFontSize(12)
    .setValue(message);
}

function getActivePlayer() {
  const player1Cell = SpreadsheetApp.getActiveSheet().getRange(PLAYER1_A1);
  const activeMarkerCell = player1Cell
    .offset(0, -1, MAX_PLAYER_COUNT, 1)
    .createTextFinder(ACTIVE_PLAYER_MARKER)
    .findNext();
  return activeMarkerCell.getRow() - player1Cell.getRow();
}

function setActivePlayer(player) {
  renderActivePlayerMarker(
    SpreadsheetApp.getActiveSheet()
      .getRange(PLAYER1_A1)
      .offset(0, -1, MAX_PLAYER_COUNT, 1)
      .clearContent()
      .offset(player, 0, 1, 1),
  );
}

function getPlayerName(player) {
  return SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(player, 0)
    .getValue();
}

function getPlayerCount() {
  const playerNames = SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(0, 0, MAX_PLAYER_COUNT, 1)
    .getValues()
    .map((row) => row[0])
    .filter((name) => name !== "");
  return playerNames.length;
}

function addCardToPlayer(player, card) {
  const playerCardsRange = SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(player, 2 + PLAYER_NAME_LENGTH, 1, 24);

  let currentCards = playerCardsRange
    .getValues()[0]
    .filter((card) => card !== "")
    .map((card) =>
      // We need to remove the single-quote from cells in runs so we can get
      // their actual numerical value
      typeof card === "number"
        ? card
        : parseInt(card.match(/^'?([1-3]?\d)$/)[1], 10),
    );
  currentCards.push(card);

  let nextCell = playerCardsRange.offset(0, 0, 1, 1);
  for (let run of groupConsecutiveRuns(currentCards)) {
    const firstCard = run.shift();
    // We add a single-quote to the rest of the cards in the run so that they
    // aren't counted as numbers when summing over the whole range.
    run = run.map((card) => `'${card}`);
    renderPlayerCardRun(nextCell, firstCard, run);
    nextCell = nextCell.offset(0, 1 + run.length, 1, 1);
  }
}

function* groupConsecutiveRuns(numbers) {
  const sorted = numbers.slice().sort(intComparator);

  let currentRun = [];
  let totalRuns = 0;
  for (let num of sorted) {
    if (
      currentRun.length > 0 &&
      // Check if the number matches at the end of the current run
      currentRun[currentRun.length - 1] !== num - 1
    ) {
      totalRuns++;
      yield currentRun;
      currentRun = [];
    }

    currentRun.push(num);
  }

  if (currentRun.length > 0) {
    totalRuns++;
    yield currentRun;
  }

  return totalRuns;
}

function enableHotspot(location, script) {
  const dimensions =
    location !== "TOKENS"
      ? { width: CARD_SIZE, height: CARD_SIZE }
      : { width: CARD_SIZE * 2, height: 4 };
  const image = getHotspotImage(location);
  image
    .setAnchorCell(image.getSheet().getRange(LOCATION_A1[location]))
    .setAnchorCellXOffset(0)
    .setAnchorCellYOffset(0)
    .setHeight(CELL_DIMENSION.HEIGHT * dimensions.height)
    .setWidth(
      CELL_DIMENSION.WIDTH * HOTSPOT_WIDTH_CORRECTION_FACTOR * dimensions.width,
    )
    .assignScript(script);
}

function resetHotspot(location) {
  getHotspotImage(location).setHeight(0).setWidth(0).assignScript("");
}

function getPlayersFromPreviousTable() {
  const previousSheet = SpreadsheetApp.getActive().getSheetByName(
    TABLE_SHEET_NAME,
  );
  if (previousSheet == null) {
    return null;
  }

  return previousSheet
    .getRange(PLAYER1_A1)
    .offset(0, 0, MAX_PLAYER_COUNT, 1)
    .getValues()
    .map((row) => row[0])
    .filter((name) => name != "");
}

////// RENDER //////////////////////////////////////////////////////////////////

function renderNewTable(file, sheet, players) {
  renderTable(file, sheet);

  renderPlayerArea(sheet, players);
  renderTokensBox(sheet);
  renderHotspots(sheet);
}

function renderHotspots(sheet) {
  sheet.getImages().forEach((image) => image.remove());

  // Insert hotspot images for the hotspot locations
  Object.keys(LOCATION_A1).forEach((_) =>
    sheet.insertImage(TRANSPARENT_PIXEL_URL, 1, 1),
  );
}

function renderDeck(cardRange) {
  const noTextStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontFamily(FONT_FAMILY)
    .setFontSize(65)
    .setForegroundColor("red")
    .build();
  const thanksTextStyle = noTextStyle.copy().setFontSize(20).build();
  const noThanksRichTextValue = SpreadsheetApp.newRichTextValue()
    .setText("NO\nTHANKS!")
    .setTextStyle(0, 2, noTextStyle)
    .setTextStyle(3, 10, thanksTextStyle)
    .build();

  cardRange
    .setBackground(DECK_BACK_COLOR)
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      "white",
      SpreadsheetApp.BorderStyle.SOLID_THICK,
    )
    .offset(2, 2, 4, CARD_SIZE - 4)
    .merge()
    .setBackground("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("bottom")
    .setRichTextValue(noThanksRichTextValue);
}

function renderCurrentCard(cardRange, cardVal) {
  cardRange
    .setBackground(cardBorderColor(cardVal))
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      "white",
      SpreadsheetApp.BorderStyle.SOLID_THICK,
    )
    .offset(1, 1, CARD_SIZE - 2, CARD_SIZE - 2)
    .merge()
    .setBackground("white")
    .setFontColor(cardNumberColor(cardVal))
    .setFontFamily(FONT_FAMILY)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(96)
    .setFontWeight("bold")
    .setValue(cardVal);
}

function renderPlayerCardRun(firstCell, firstCard, restOfRun) {
  const runRange = firstCell.offset(0, 0, 1, restOfRun.length + 1);
  runRange
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      cardBorderColor(firstCard),
      SpreadsheetApp.BorderStyle.SOLID_THICK,
    )
    .setBackground("white")
    .setFontColor(cardNumberColor(firstCard))
    .setFontFamily(FONT_FAMILY)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);
  runRange.setValues([[firstCard].concat(restOfRun)]);
}

function renderPlayerArea(sheet, players) {
  sheet
    .getRange(PLAYER1_A1)
    .offset(0, 0, players.length, PLAYER_NAME_LENGTH)
    .mergeAcross()
    .offset(0, 0, players.length, 1)
    .setValues(players.map((player) => [player]))
    .offset(0, -1, players.length, 1 + PLAYER_NAME_LENGTH + 2 + 24)
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_THICK,
    )
    .offset(0, 1 + PLAYER_NAME_LENGTH, players.length, 2)
    .mergeAcross()
    .offset(0, 0, players.length, 1)
    .setNumberFormat(getPlayerTokensNumberFormat(OPTIONS.HIDDEN_TOKENS))
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
}

function getPlayerTokensNumberFormat(isHidden) {
  return isHidden
    ? `"??" ${TOKEN_REPR}`
    : `# ${TOKEN_REPR};-# ${TOKEN_REPR};"None";"??" ${TOKEN_REPR}`;
}

function renderFinalScore() {
  const numPlayers = getPlayerCount();
  const player1Cell = SpreadsheetApp.getActiveSheet().getRange(PLAYER1_A1);
  const cardsRange = player1Cell.offset(
    0,
    PLAYER_NAME_LENGTH + 2,
    numPlayers,
    24,
  );
  const maxCards = Math.max(
    ...cardsRange.getValues().map((row) => row.indexOf("")),
  );

  const scores = cardsRange
    .offset(0, maxCards, numPlayers, 24 - maxCards)
    .mergeAcross()
    .offset(0, 0, numPlayers, 1)
    .setFormulaR1C1(
      `=SUM(R[0]C[${-maxCards}]:R[0]C[-1])-IF(ISNUMBER(R[0]C[${
        -maxCards - 2
      }]), R[0]C[${-maxCards - 2}], 0)`,
    )
    .setFontSize(16)
    .setFontWeight("bold")
    .setNumberFormat("0")
    .getValues();

  const scoresRanked = scores.slice().sort(intComparator);
  const activePlayerRange = player1Cell
    .offset(0, -1, numPlayers, 1)
    .setValues(
      scores.map((score) => [RANK_MARKERS[scoresRanked.indexOf(score)]]),
    );
}

function renderActivePlayerMarker(range) {
  return range
    .setFontSize(14)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setValue(ACTIVE_PLAYER_MARKER);
}

function renderTable(file, sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < TABLE_DIMENSIONS.HEIGHT) {
    sheet.insertRows(1, TABLE_DIMENSIONS.HEIGHT - maxRows);
  } else if (maxRows > TABLE_DIMENSIONS.HEIGHT) {
    sheet.deleteRows(
      TABLE_DIMENSIONS.HEIGHT + 1,
      maxRows - TABLE_DIMENSIONS.HEIGHT,
    );
  }

  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < TABLE_DIMENSIONS.WIDTH) {
    sheet.insertColumns(1, TABLE_DIMENSIONS.WIDTH - maxColumns);
  } else if (maxColumns > TABLE_DIMENSIONS.WIDTH) {
    sheet.deleteColumns(
      TABLE_DIMENSIONS.WIDTH + 1,
      maxColumns - TABLE_DIMENSIONS.WIDTH,
    );
  }

  for (let i = 1; i <= TABLE_DIMENSIONS.HEIGHT; i++) {
    file.setRowHeight(i, CELL_DIMENSION.HEIGHT);
  }
  for (let i = 1; i <= TABLE_DIMENSIONS.WIDTH; i++) {
    file.setColumnWidth(i, CELL_DIMENSION.WIDTH);
  }

  sheet
    .getRange(1, 1, TABLE_DIMENSIONS.HEIGHT, TABLE_DIMENSIONS.WIDTH)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setBackground(BG_COLOR);
}

function renderTokensBox(sheet) {
  const tokensRange = SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.TOKENS)
    .offset(0, 0, 4, CARD_SIZE * 2);

  tokensRange
    .merge()
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      "white",
      SpreadsheetApp.BorderStyle.SOLID_THICK,
    )
    .setWrap(true);
}

function getNewPlayersFromUser() {
  const ui = SpreadsheetApp.getUi();

  const players = [];
  for (const i of xrange(1, MAX_PLAYER_COUNT + 1)) {
    const response = ui.prompt(
      `${i <= 3 ? "" : "Add "}Player ${i}${i <= 3 ? "" : "?"}`,
      "Name:",
      i <= 3 ? ui.ButtonSet.OK : ui.ButtonSet.YES_NO,
    );

    if (response.getSelectedButton() === ui.Button.NO) {
      break;
    }

    players.push(response.getResponseText());
  }

  if (players.length < 3 || players.length > 7) {
    throw new Error("We only support between 3 to 7 players");
  }

  return players;
}

function cardNumberColor(cardVal) {
  if (cardVal >= 3 && cardVal <= 13) {
    return "#0586ff";
  } else if (cardVal >= 14 && cardVal <= 24) {
    return "#ff9d14";
  } else if (cardVal >= 15 && cardVal <= 35) {
    return "#d11500";
  }

  throw new Error("No color defined for card value " + cardVal);
}

function cardBorderColor(cardVal) {
  return (
    "#" +
    [
      "FFE541",
      "FCF93D",
      "E5F938",
      "CBF734",
      "B1F42F",
      "96F12B",
      "7BEF27",
      "5FEC22",
      "43EA1E",
      "27E71A",
      "16E422",
      "13E237",
      "0FDF4C",
      "0BDC61",
      "07DA77",
      "04D78D",
      "00D5A3",
      "00D0B0",
      "00CCBD",
      "00C8C8",
      "00B5C4",
      "00A2C0",
      "0090BC",
      "007FB8",
      "006EB4",
      "005EB0",
      "004FAC",
      "0040A8",
      "0031A4",
      "0024A0",
      "00179C",
      "000A98",
      "010094",
    ][cardVal - 3]
  );
}

////// GENERIC SHEET HELPERS ///////////////////////////////////////////////////

function getCellValue(a1Notation) {
  const cell = SpreadsheetApp.getActiveSheet().getRange(a1Notation);
  return cell.isBlank() ? null : cell.getValue();
}

function singleEntry(func) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const fullTableRange = sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns(),
  );

  const lock = LockService.getUserLock();
  lock.waitLock(MUTEX_LOCKOUT_PERIOD_MS);

  // We activate the whole table as a signal that the script is running
  fullTableRange.activate();

  try {
    func();
  } finally {
    fullTableRange.offset(0, 0, 1, 1).activate();

    lock.releaseLock();
  }
}

function getHotspotImage(location) {
  const imageIndex = Object.keys(LOCATION_A1).sort().indexOf(location);
  if (imageIndex === -1) {
    throw new Error(`Unknown hotspot location ${location}`);
  }

  const images = SpreadsheetApp.getActiveSheet().getImages();
  if (images.length !== Object.keys(LOCATION_A1).length) {
    throw new Error(
      `Expecting exactly ${
        Object.keys(LOCATION_A1).length
      } images on the sheet, found ${images.length} instead!`,
    );
  }

  return images[imageIndex];
}

////// GENERIC JS HELPERS //////////////////////////////////////////////////////

function* xrange(a, b) {
  const start = b == null ? 0 : a;
  const end = b == null ? a : b;

  for (let i = start; i < end; i++) {
    yield i;
  }

  return start - end;
}

function randInt(a, b) {
  const min = b == null ? 0 : a;
  const max = b == null ? a : b;
  return min + Math.floor(Math.random() * (max - min + 1));
}

function shuffle(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const randomIndex = randInt(i);
    const temporaryValue = array[i];
    array[i] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}

function intComparator(a, b) {
  return a - b;
}
