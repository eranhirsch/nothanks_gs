/** @OnlyCurrentDoc */

const MAX_PLAYER_COUNT = 7;
const PLAYER1_ROW = 20;
const MIN_CARD = 3;
const MAX_CARD = 35;
const SETUP_CARDS_REMOVED = 9;
const TOKEN_REPR = "🌑 ";
const ACTIVE_PLAYER_MARKER = "➡️";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("No Thanks")
    .addItem("New Game", "newGame")
    .addToUi();
}

function revealTopCard() {
  if (getCurrentCard() != null) {
    Browser.msgBox(
      "You can't reveal the next card until the current card is taken!",
    );
    return;
  }

  const deck = getDeck();
  if (deck.length === 0) {
    Browser.msgBox("Deck is empty!");
    return;
  }

  const { remainingDeck, card } = drawCard(deck);
  setCurrentCard(card);
  setDeck(remainingDeck);
  setCurrentTokens(0);
}

function takeCard() {
  const card = getCurrentCard();
  const tokens = getCurrentTokens();
  if (card == null) {
    Browser.msgBox("No card revealed yet to take");
    return;
  }
  setCurrentCard(null);
  setCurrentTokens(null);

  player = getActivePlayer();
  addCardToPlayer(player, card);
  addTokensToPlayer(player, tokens);
}

function noThanks() {
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
}

///////////////////////////////////////////////////////////////////////////////////

function newGame() {
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
  setDeck(newDeck());
}

function drawCard(deck) {
  const randIndex = randInt(deck.length);
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
    throw new Error("Not enough tokens");
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

  const playerTokens = range(playerCount - 1).reduce((obj, player) => {
    obj[player] = tokens;
    return obj;
  }, {});
  setPlayerTokens(playerTokens);
}

//////////////////////////////////////////////////////////////////////////

function getDeck() {
  const value = getSheetMetadataObject("deck").getValue();
  if (value == null || value === "") {
    return null;
  }

  return JSON.parse(value);
}

function setDeck(deck) {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (deck == null) {
    resetSheetMetadataObject("deck");
    sheet.getRange("C1").clearContent();
    return;
  }

  const serialized = JSON.stringify(deck);
  getSheetMetadataObject("deck").setValue(serialized);

  sheet.getRange("C1").setValue(deck.length);
}

function getCurrentCard() {
  const cardStr = getCellValue("N2");
  return cardStr != null ? parseInt(cardStr) : null;
}

function setCurrentCard(cardVal) {
  setCellValue("N2", cardVal);
}

function getCurrentTokens() {
  const tokenStr = getCellValue("F12");
  return tokenStr != null ? tokenStr.length / TOKEN_REPR.length : null;
}

function setCurrentTokens(tokens) {
  const tokensStr = TOKEN_REPR.repeat(tokens);
  setCellValue("F12", tokensStr);
}

function getActivePlayer() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeMarkersRange = sheet.getRange(
    PLAYER1_ROW,
    1,
    MAX_PLAYER_COUNT,
    1,
  );
  const activeMarkerCell = activeMarkersRange
    .createTextFinder(ACTIVE_PLAYER_MARKER)
    .findNext();
  const activePlayer = activeMarkerCell.getRow() - PLAYER1_ROW;
  return activePlayer;
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
  const sheet = SpreadsheetApp.getActiveSheet();
  const playerCardsRange = sheet.getRange(PLAYER1_ROW + player, 3, 1, 24);

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
  return JSON.parse(serialized);
}

function setPlayerTokens(playerTokens) {
  if (playerTokens == null) {
    resetSheetMetadataObject("playerTokens");
    return;
  }
  const serialized = JSON.stringify(playerTokens);
  getSheetMetadataObject("playerTokens").setValue(serialized);
}

/////////////////////////////////////////////////////////////////////////

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

/////////////////////////////////////////////////////////////////////////

function range(max, min = 0) {
  return [...Array(max + 1).keys()].splice(min);
}

function randInt(max, min = 0) {
  return min + Math.floor(Math.random() * (max - min + 1));
}