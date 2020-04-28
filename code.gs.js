/** @OnlyCurrentDoc */

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
Click HERE to reveal everyone's personal ${TOKEN_REPR}`;

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
  // Remove any previous locks
  resetSheetMetadataObject("lock");

  const ui = SpreadsheetApp.getUi();
  ui.createMenu("No Thanks")
    .addSubMenu(
      ui
        .createMenu("New Table")
        .addItem("Same Players", "newTableSamePlayers")
        .addItem("New Players", "newTableNewPlayers"),
    )
    .addToUi();
}

////// USER ACTIONS ////////////////////////////////////////////////////////////

function newTableNewPlayers() {
  singleEntry(() => {
    const players = [];
    for (let i = 1; i <= MAX_PLAYER_COUNT; i++) {
      const name = Browser.inputBox(
        "Add Player" + (i <= 3 ? "" : "?"),
        "Enter name" +
          (i <= 3 ? "" : " (or leave blank to finish entering names)"),
        Browser.Buttons.OK,
      );
      if (name == null || name === "") {
        break;
      }
      players.push(name);
    }

    if (players.length < 3 || players.length > 7) {
      throw new Error("We only support between 3 to 7 players");
    }

    newTable(players);
  });
}

function newTableSamePlayers() {
  singleEntry(() => {
    const previousNames = SpreadsheetApp.getActive()
      .getSheetByName(TABLE_SHEET_NAME)
      .getRange(PLAYER1_A1)
      .offset(0, 0, MAX_PLAYER_COUNT, 1)
      .getValues()
      .map((row) => row[0])
      .filter((name) => name != "");
    newTable(previousNames);
  });
}

function revealTopCard() {
  singleEntry(() => {
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
    setInstructionsMessage(MSG_TURN);
  });
}

function takeCard() {
  singleEntry(() => {
    const card = getCurrentCard();
    if (card == null) {
      throw new Error("No card revealed yet to take");
    }

    const tokens = getCurrentTokens();

    resetHotspot("CARD");
    setCurrentCard(null);
    resetHotspot("TOKENS");

    player = getActivePlayer();
    addCardToPlayer(player, card);

    if (tokens > 0) {
      Browser.msgBox(
        getPlayerName(player),
        `Add ${tokens}${TOKEN_REPR} to your personal pool`,
        Browser.Buttons.OK,
      );
      addTokensToPlayer(player, tokens);
    }

    const deck = getDeck();
    if (deck != null && deck.length > 0) {
      setInstructionsMessage(MSG_REVEAL);
      enableHotspot("DECK", "revealTopCard");
    } else {
      setInstructionsMessage(MSG_END_GAME);
      enableHotspot("TOKENS", "revealTokens");
    }
  });
}

function revealTokens() {
  singleEntry(() => {
    renderPlayerTokens();
    resetHotspot("TOKENS");
    setInstructionsMessage("");
  });
}

function noThanks() {
  singleEntry(() => {
    const currentCard = getCurrentCard();
    if (currentCard == null) {
      Browser.msgBox("No card revealed yet!");
      return;
    }

    // Take token from player
    addTokensToPlayer(getActivePlayer(), -1);

    // And add it to the pool
    setCurrentTokens(getCurrentTokens() + 1);
    advanceActivePlayer();
  });
}

////// LOGICAL ACTIONS /////////////////////////////////////////////////////////

function newTable(players) {
  const file = SpreadsheetApp.getActive();
  const newSheet = file.insertSheet("Creating new table...");

  try {
    file.setActiveSheet(newSheet);

    renderTable(file, newSheet);

    // We also shuffle the players to randomize seating each time
    renderPlayerArea(newSheet, shuffle(players));
    renderTokensBox(newSheet);

    // Insert hotspot images for the hotspot locations
    range(Object.keys(LOCATION_A1).length - 1).forEach((_) =>
      newSheet.insertImage(TRANSPARENT_PIXEL_URL, 1, 1),
    );

    newGame(players.length);

    // Enable the hotspot for the first action
    enableHotspot("DECK", "revealTopCard");
    // ... and update the tokens pool with instructions
    setInstructionsMessage(MSG_REVEAL);

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
}

function newGame(numPlayers) {
  // Randomize a new deck
  setDeck(newDeck());

  // deal tokens to each player
  dealTokens(numPlayers);

  // Pick a random starting player
  setActivePlayer(randInt(numPlayers - 1));
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
  let deck = range(MIN_CARD, MAX_CARD);

  // Remove cards from the deck
  for (let i = 0; i < SETUP_CARDS_REMOVED; i++) {
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

function addTokensToPlayer(player, tokens) {
  const playerTokens = getPlayerTokens();

  const playersTokens = player in playerTokens ? playerTokens[player] : 0;
  const newPlayerTokensCount = playersTokens + tokens;
  if (newPlayerTokensCount < 0) {
    throw new Error(getPlayerName(player) + " doesn't have any tokens left!");
  }
  playerTokens[player] = newPlayerTokensCount;

  setPlayerTokens(playerTokens);
}

function dealTokens(playerCount) {
  let tokens;
  if (playerCount <= 5) {
    tokens = 11;
  } else if (playerCount == 6) {
    tokens = 9;
  } else if (playerCount == 7) {
    tokens = 7;
  }

  Browser.msgBox(
    "All Players",
    `Add ${tokens}${TOKEN_REPR} to your personal pool`,
    Browser.Buttons.OK,
  );
  range(playerCount - 1).forEach((player) => addTokensToPlayer(player, tokens));
}

////// STATE MANAGEMENT ////////////////////////////////////////////////////////

function getDeck() {
  const value = getSheetMetadataObject("deck").getValue();
  if (value == null || value === "") {
    return null;
  }

  return JSON.parse(value);
}

function setDeck(deck) {
  const cardRange = SpreadsheetApp.getActiveSheet()
    .getRange(LOCATION_A1.DECK)
    .offset(0, 0, CARD_SIZE, CARD_SIZE);

  if (deck == null || deck.length === 0) {
    resetSheetMetadataObject("deck");

    cardRange
      .breakApart()
      .clear()
      .setBackground(BG_COLOR)
      .setBorder(false, false, false, false, false, false);
    return;
  }

  const serialized = JSON.stringify(deck);
  getSheetMetadataObject("deck").setValue(serialized);

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

function getCurrentTokens() {
  const tokenStr = getCellValue(LOCATION_A1.TOKENS);

  if (
    tokenStr != null &&
    tokenStr.match(new RegExp(`^${TOKEN_REPR}+$`, "gu"))
  ) {
    return tokenStr.length / TOKEN_REPR.length;
  }

  // We use the current token pool for messaging too, but in those cases the actual
  // tokens are always null (and could be coalesced to 0).
  return null;
}

function setCurrentTokens(tokens) {
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

  let currentCards = playerCardsRange.getValues()[0];
  currentCards = currentCards.filter((card) => card !== "");
  currentCards.push(card);
  currentCards
    .sort((a, b) => a - b)
    .reduce(
      (currentCell, currentCard) =>
        renderPlayerCard(currentCell, currentCard).offset(0, 1),
      playerCardsRange.offset(0, 0, 1, 1),
    );
}

function getPlayerTokens() {
  const serialized = getSheetMetadataObject("playerTokens").getValue();
  return serialized !== "" ? JSON.parse(serialized) : {};
}

function setPlayerTokens(playerTokens) {
  if (playerTokens == null) {
    resetSheetMetadataObject("playerTokens");
    return;
  }
  const serialized = JSON.stringify(playerTokens);
  getSheetMetadataObject("playerTokens").setValue(serialized);
}

function enableHotspot(location, script, title = "", description = "") {
  const image = getHotspotImage(location);
  if (location === "TOKENS") {
    image
      .setHeight(CELL_DIMENSION.HEIGHT * 4)
      .setWidth(
        CELL_DIMENSION.WIDTH * CARD_SIZE * 2 * HOTSPOT_WIDTH_CORRECTION_FACTOR,
      );
  } else {
    image
      .setHeight(CELL_DIMENSION.HEIGHT * CARD_SIZE)
      .setWidth(
        CELL_DIMENSION.WIDTH * CARD_SIZE * HOTSPOT_WIDTH_CORRECTION_FACTOR,
      );
  }
  image
    .setAltTextTitle(title !== "" ? title : script)
    .setAltTextDescription(description)
    .assignScript(script);
}

function resetHotspot(location) {
  getHotspotImage(location)
    .setHeight(0)
    .setWidth(0)
    .setAltTextTitle("")
    .setAltTextDescription("")
    .assignScript("");
}

////// RENDER //////////////////////////////////////////////////////////////////

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
  return renderCard(
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
      .merge(),
    cardVal,
  )
    .setFontSize(96)
    .setFontWeight("bold");
}

function renderPlayerCard(cardRange, cardVal) {
  return renderCard(
    cardRange.setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      cardBorderColor(cardVal),
      SpreadsheetApp.BorderStyle.SOLID_THICK,
    ),
    cardVal,
  ).setFontSize(12);
}

function renderCard(cardRange, cardVal) {
  return cardRange
    .setBackground("white")
    .setFontColor(cardNumberColor(cardVal))
    .setFontFamily(FONT_FAMILY)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setValue(cardVal);
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
    .merge()
    .offset(0, 0, 1, 1)
    .setValue(TOKEN_REPR + ": Hidden")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setTextRotation(-90);
}

function renderPlayerTokens() {
  const playerCount = getPlayerCount();
  SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_A1)
    .offset(0, PLAYER_NAME_LENGTH, playerCount, 2)
    .breakApart()
    .mergeAcross()
    .offset(0, 0, playerCount, 1)
    .setTextRotation(0)
    .setValues(
      Object.values(getPlayerTokens()).map((tokens) => [
        `${tokens}${TOKEN_REPR}`,
      ]),
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

function getSheetMetadataObject(key) {
  const sheet = SpreadsheetApp.getActiveSheet();

  let arr = sheet.createDeveloperMetadataFinder().withKey(key).find();

  if (arr.length !== 1) {
    resetSheetMetadataObject(key);
    arr = [];
  }

  if (arr.length === 0) {
    sheet.addDeveloperMetadata(key);
    arr = sheet.createDeveloperMetadataFinder().withKey(key).find();
  }

  return arr[0];
}

function resetSheetMetadataObject(key) {
  SpreadsheetApp.getActive()
    .getActiveSheet()
    .createDeveloperMetadataFinder()
    .withKey(key)
    .find()
    .forEach((md) => md.remove());
}

function getCellValue(a1Notation) {
  const cell = SpreadsheetApp.getActiveSheet().getRange(a1Notation);
  return cell.isBlank() ? null : cell.getValue();
}

function singleEntry(func) {
  const metadata = getSheetMetadataObject("lock");
  const timeoutStr = metadata.getValue();
  const timestamp = new Date().getTime();

  if (timeoutStr !== "") {
    if (parseInt(timeoutStr, 10) > timestamp) {
      throw new Error("Lock active, previous operation hasn't completed yet");
    } else {
      Browser.msgBox(
        "Previous lock wasn't cleared but we are out of the lockout period." +
          "Timeout was: " +
          timeoutStr +
          ", Timestamp is: " +
          timestamp,
      );
    }
  }

  metadata.setValue(timestamp + MUTEX_LOCKOUT_PERIOD_MS);

  try {
    func();
  } finally {
    // Release the lock
    resetSheetMetadataObject("lock");
  }
}

function getHotspotImage(location) {
  const imageIndex = Object.keys(LOCATION_A1).sort().indexOf(location);
  if (imageIndex === -1) {
    throw new Error("Unknown hotspot location " + location);
  }

  const images = SpreadsheetApp.getActiveSheet().getImages();
  if (images.length !== Object.keys(LOCATION_A1).length) {
    throw new Error(
      "Expecting exactly " +
        Object.keys(LOCATION_A1).length +
        " images on the sheet, found " +
        images.length +
        " instead!",
    );
  }

  return images[imageIndex]
    .setAnchorCell(
      SpreadsheetApp.getActiveSheet().getRange(LOCATION_A1[location]),
    )
    .setAnchorCellXOffset(0)
    .setAnchorCellYOffset(0);
}

////// GENERIC JS HELPERS //////////////////////////////////////////////////////

function range(a, b) {
  const min = b == null ? 0 : a;
  const max = b == null ? a : b;
  return [...Array(max + 1).keys()].splice(min);
}

function randInt(a, b) {
  const min = b == null ? 0 : a;
  const max = b == null ? a : b;
  return min + Math.floor(Math.random() * (max - min + 1));
}

function shuffle(array) {
  for (let i = array.length; i > 0; i--) {
    const randomIndex = randInt(i - 1);
    const temporaryValue = array[i - 1];
    array[i - 1] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}
