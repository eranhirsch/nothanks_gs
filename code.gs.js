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
const CARD_SIZE = 10;
const TOKEN_REPR = "🌑";
const ACTIVE_PLAYER_MARKER = "➡️";
const BG_COLOR = "#fff3dc";
const DECK_BACK_COLOR = "#1c4587";
const PLAYER1_ROW = 18;
const CELL_DIMENSION = { WIDTH: 21, HEIGHT: 30 };
const LOCATION_A1 = { DECK: "D2", CARD: "P2", TOKENS: "E13" };

// Our calculation of hotspot width seems to miss by a bit so we can just expand
// it a little by a constant factor to try and fix it.
const HOTSPOT_WIDTH_CORRECTION_FACTOR = 1.05;

/**
 * Mechanism consts
 */
const MUTEX_LOCKOUT_PERIOD_MS = 5000;

////// API HOOKS ///////////////////////////////////////////////////////////////

function onOpen() {
  // Remove any previous locks
  resetSheetMetadataObject("lock");

  SpreadsheetApp.getUi()
    .createMenu("No Thanks")
    .addItem("New Game", "newGame")
    .addToUi();
}

////// USER ACTIONS ////////////////////////////////////////////////////////////

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
    setCurrentTokens(0);
  });
}

function takeCard() {
  singleEntry(() => {
    const card = getCurrentCard();
    const tokens = getCurrentTokens();
    if (card == null) {
      Browser.msgBox("No card revealed yet to take");
      return;
    }
    resetHotspot("CARD");
    setCurrentCard(null);
    resetHotspot("TOKENS");
    setCurrentTokens(null);

    player = getActivePlayer();
    addCardToPlayer(player, card);
    addTokensToPlayer(player, tokens);

    const deck = getDeck();
    if (deck != null && deck.length > 0) {
      enableHotspot("DECK", "revealTopCard");
    }
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

function newGame() {
  singleEntry(() => {
    // Reset all hotspots
    Object.keys(LOCATION_A1).forEach((location) => resetHotspot(location));
    setDeck(null);
    setPlayerTokens(null);
    setCurrentCard(null);
    setCurrentTokens(null);
    resetPlayerCards();

    const playerCount = getPlayerCount();

    // Pick a random starting player
    setActivePlayer(randInt(playerCount - 1));

    // deal tokens to each player
    dealTokens(playerCount);

    // Recreate the starting deck
    enableHotspot("DECK", "revealTopCard");
    setDeck(newDeck());
  });
}

////// LOGICAL ACTIONS /////////////////////////////////////////////////////////

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
  var deck = range(MAX_CARD, MIN_CARD);

  // Remove cards from the deck
  for (var i = 0; i < SETUP_CARDS_REMOVED; i++) {
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

  const newPlayerTokensCount = playerTokens[player] + tokens;
  if (newPlayerTokensCount < 0) {
    throw new Error("The player doesn't have any tokens left!");
  }
  playerTokens[player] = newPlayerTokensCount;

  setPlayerTokens(playerTokens);
}

function dealTokens(playerCount) {
  var tokens;
  if (playerCount <= 5) {
    tokens = 11;
  } else if (playerCount == 6) {
    tokens = 9;
  } else if (playerCount == 7) {
    tokens = 7;
  }

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
  return cardStr != null ? parseInt(cardStr) : null;
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

  renderCard(cardRange);
}

function getCurrentTokens() {
  const tokenStr = getCellValue(LOCATION_A1.TOKENS);
  return tokenStr != null ? tokenStr.length / TOKEN_REPR.length : null;
}

function setCurrentTokens(tokens) {
  const tokensStr = TOKEN_REPR.repeat(tokens);
  setCellValue(LOCATION_A1.TOKENS, tokensStr);
}

function getActivePlayer() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeMarkerCell = sheet
    .getRange(PLAYER1_ROW, 1, MAX_PLAYER_COUNT, 1)
    .createTextFinder(ACTIVE_PLAYER_MARKER)
    .findNext();
  return activeMarkerCell.getRow() - PLAYER1_ROW;
}

function setActivePlayer(player) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(PLAYER1_ROW, 1, MAX_PLAYER_COUNT, 1).clearContent();
  sheet.getRange(PLAYER1_ROW + player, 1).setValue(ACTIVE_PLAYER_MARKER);
}

function getPlayerCount() {
  const playerNames = SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_ROW, 2, MAX_PLAYER_COUNT, 1)
    .getValues()
    .map((row) => row[0])
    .filter((name) => name !== "");
  return playerNames.length;
}

function addCardToPlayer(player, card) {
  const playerCardsRange = SpreadsheetApp.getActiveSheet().getRange(
    PLAYER1_ROW + player,
    3,
    1,
    24,
  );

  var currentCards = playerCardsRange.getValues()[0];
  currentCards = currentCards.filter((card) => card !== "");
  currentCards.push(card);
  currentCards.sort((a, b) => a - b);

  playerCardsRange
    .clearContent()
    .offset(0, 0, 1, currentCards.length)
    .setValues([currentCards])
    .setBorder(true, true, true, true, true, true);
}

function resetPlayerCards() {
  SpreadsheetApp.getActiveSheet()
    .getRange(PLAYER1_ROW, 3, MAX_PLAYER_COUNT, 24)
    .clear();
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
      .setWidth(CELL_DIMENSION.WIDTH * 20 * HOTSPOT_WIDTH_CORRECTION_FACTOR);
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
    .setFontFamily("Francois One")
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

function renderCard(cardRange) {
  cardRange
    .setBackground("red")
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
    .setFontColor("blue")
    .setFontFamily("Impact")
    .setFontSize(96)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setValue(cardVal);
}

////// GENERIC SHEET HELPERS ///////////////////////////////////////////////////

function getSheetMetadataObject(key) {
  const sheet = SpreadsheetApp.getActiveSheet();

  var arr = sheet.createDeveloperMetadataFinder().withKey(key).find();

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

function setCellValue(a1Notation, value) {
  const cell = SpreadsheetApp.getActiveSheet().getRange(a1Notation);
  if (value == null) {
    cell.clearContent();
  }
  cell.setValue(value);
}

function singleEntry(func) {
  const metadata = getSheetMetadataObject("lock");
  const timeoutStr = metadata.getValue();
  const timestamp = new Date().getTime();

  if (timeoutStr !== "") {
    if (parseInt(timeoutStr) > timestamp) {
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

function range(max, min = 0) {
  return [...Array(max + 1).keys()].splice(min);
}

function randInt(max, min = 0) {
  return min + Math.floor(Math.random() * (max - min + 1));
}
