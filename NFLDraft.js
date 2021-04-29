// UPDATE EACH YEAR: Team needs (from highest priority to lowest priority)
let needs = {};
needs.BUF = [[`EDGE`], [`TE`, `IOL`, `EDGE`, `DI`, `LB`], [`CB`], [`IOL`, `EDGE`], [`DI`, `LB`]];
needs.MIA = [[`WR`], [`EDGE`], [`IOL`], [`QB`, `RB`, `TE`, `OT`, `LB`], [`IOL`, `EDGE`, `OT`, `LB`]];
needs.NE = [[`QB`], [`WR`, `EDGE`], [`OT`, `DI`, `LB`, `CB`, `S`], [`WR`, `EDGE`], [`DI`, `LB`]];
needs.NYJ = [[`QB`], [`EDGE`, `CB`], [`RB`, `TE`, `CB`], [`OT`, `IOL`, `LB`], [`RB`, `EDGE`]];
needs.BAL = [[`WR`], [`OT`], [`IOL`, `EDGE`], [`LB`, `S`], [`WR`, `IOL`, `EDGE`, `LB`, `S`]];
needs.CIN = [[`WR`], [`TE`, `IOL`, `EDGE`, `LB`], [`RB`, `OT`, `EDGE`], [`IOL`, `LB`], [`OT`, `EDGE`]];
needs.CLE = [[`EDGE`, `LB`, `CB`], [`WR`, `OT`, `DI`], [`EDGE`, `LB`], [`K`], [`OT`, `DI`]];
needs.PIT = [[`QB`, `RB`, `OT`, `IOL`, `CB`], [`EDGE`, `LB`], [`OT`, `IOL`, `CB`], [`EDGE`, `LB`], [`QB`]];
needs.HOU = [[`EDGE`, `DI`], [`QB`, `WR`, `TE`, `IOL`, `CB`], [`EDGE`, `DI`], [`WR`, `TE`, `IOL`], [`CB`, `S`]];
needs.IND = [[`OT`], [`WR`, `EDGE`], [`EDGE`, `CB`], [`S`], [`WR`, `EDGE`]];
needs.JAX = [[`QB`, `S`], [`OT`], [`WR`, `TE`, `EDGE`, `DI`, `LB`], [`OT`, `S`], [`RB`, `DI`, `LB`]];
needs.TEN = [[`WR`], [`EDGE`, `CB`], [`TE`, `OT`, `DI`, `S`], [`WR`, `EDGE`, `CB`], [`RB`, `OT`, `DI`, `S`]];
needs.DEN = [[`QB`], [`IOL`, `LB`], [`OT`, `EDGE`, `DI`, `CB`], [`IOL`, `LB`], [`EDGE`, `DI`]];
needs.KC = [[`EDGE`, `LB`], [`WR`, `CB`], [`EDGE`, `LB`], [`OT`, `IOL`], [`WR`, `CB`]];
needs.LAC = [[`OT`], [`WR`, `TE`, `EDGE`, `CB`, `S`], [`OT`, `CB`], [`K`], [`WR`, `IOL`]];
needs.LV = [[`OT`], [`IOL`, `DI`, `CB`, `S`], [`LB`], [`CB`, `S`], [`IOL`, `DI`]];
needs.DAL = [[`CB`], [`OT`, `IOL`, `EDGE`, `DI`, `LB`], [`TE`, `CB`, `S`], [`P`], [`OT`, `IOL`]];
needs.NYG = [[`LB`, `EDGE`], [`OT`, `IOL`], [`CB`], [`IOL`], [`OT`, `IOL`, `EDGE`, `LB`, `CB`]];
needs.PHI = [[`CB`], [`QB`, `WR`, `OT`, `IOL`, `EDGE`, `LB`], [`WR`, `LB`, `CB`], [`P`], [`IOL`, `EDGE`]];
needs.WAS = [[`QB`, `WR`, `TE`, `OT`, `LB`], [`CB`, `S`], [`EDGE`], [`QB`, `WR`, `LB`], [`TE`, `OT`]];
needs.CHI = [[`QB`, `WR`, `OT`, `CB`], [`EDGE`, `LB`], [`WR`, `CB`], [`QB`, `OT`], [`EDGE`, `LB`]];
needs.DET = [[`WR`], [`IOL`, `CB`, `S`], [`OT`], [`EDGE`, `CB`, `S`], [`QB`]];
needs.GB = [[`WR`], [`OT`, `IOL`, `LB`, `CB`], [`DI`], [`WR`, `IOL`, `LB`, `CB`], [`EDGE`, `DI`]];
needs.MIN = [[`IOL`], [`WR`, `TE`, `OT`, `IOL`, `EDGE`, `CB`, `S`], [`OT`, `EDGE`], [`K`], [`QB`,  `CB`]];
needs.ATL = [[`TE`], [`IOL`, `EDGE`], [`RB`, `OT`, `LB`, `CB`, `S`], [`EDGE`], [`QB`, `IOL`]];
needs.CAR = [[`OT`], [`TE`, `IOL`, `CB`], [`EDGE`, `S`], [`IOL`, `CB`], [`S`]];
needs.NO = [[`WR`], [`CB`], [`TE`, `IOL`, `EDGE`, `S`], [`WR`, `CB`], [`DI`, `LB`]];
needs.TB = [[`EDGE`], [`QB`, `WR`, `OT`, `IOL`, `DI`, `LB`], [`EDGE`], [`OT`, `IOL`, `DI`, `LB`], [`WR`, `TE`]];
needs.ARI = [[`WR`, `TE`, `CB`], [`IOL`], [`DI`, `LB`], [`EDGE`], [`IOL`]];
needs.LAR = [[`EDGE`, `S`], [`IOL`, `LB`], [`TE`], [`EDGE`, `S`], [`OT`]];
needs.SEA = [[`EDGE`], [`WR`, `IOL`], [`OT`, `CB`], [`EDGE`], [`WR`, `IOL`]];
needs.SF = [[`QB`], [`IOL`, `EDGE`, `CB`], [`LB`, `CB`, `S`], [`EDGE`], [`RB`, `WR`, `OT`]];

// UPDATE EACH YEAR: Number of available prospects and draft picks
const numQB = 47;
const numRB = 95;
const numWR = 163;
const numTE = 62;
const numOT = 78;
const numOG = 98;
const numOC = 29;
const numEDGE = 143;
const numDI = 85;
const numLB = 93;
const numCB = 128;
const numS = 118;
const numK = 7;
const numP = 8;
const numPicks = [0, 32, 32, 41, 39, 40, 44, 31];

// Available prospects per position
const allQB = SpreadsheetApp.getActive().getSheetByName(`QB`).getRange(2, 1, numQB, 4).getValues();
const allRB = SpreadsheetApp.getActive().getSheetByName(`RB`).getRange(2, 1, numRB, 4).getValues();
const allWR = SpreadsheetApp.getActive().getSheetByName(`WR`).getRange(2, 1, numWR, 4).getValues();
const allTE = SpreadsheetApp.getActive().getSheetByName(`TE`).getRange(2, 1, numTE, 4).getValues();
const allOT = SpreadsheetApp.getActive().getSheetByName(`OT`).getRange(2, 1, numOT, 4).getValues();
const allOG = SpreadsheetApp.getActive().getSheetByName(`OG`).getRange(2, 1, numOG, 4).getValues();
const allOC = SpreadsheetApp.getActive().getSheetByName(`OC`).getRange(2, 1, numOC, 4).getValues();
const allEDGE = SpreadsheetApp.getActive().getSheetByName(`EDGE`).getRange(2, 1, numEDGE, 4).getValues();
const allDI = SpreadsheetApp.getActive().getSheetByName(`DI`).getRange(2, 1, numDI, 4).getValues();
const allLB = SpreadsheetApp.getActive().getSheetByName(`LB`).getRange(2, 1, numLB, 4).getValues();
const allCB = SpreadsheetApp.getActive().getSheetByName(`CB`).getRange(2, 1, numCB, 4).getValues();
const allS = SpreadsheetApp.getActive().getSheetByName(`S`).getRange(2, 1, numS, 4).getValues();
const allK = SpreadsheetApp.getActive().getSheetByName(`K`).getRange(2, 1, numK, 4).getValues();
const allP = SpreadsheetApp.getActive().getSheetByName(`P`).getRange(2, 1, numP, 4).getValues();

// Indices of best available prospects per position (not chosen by simulator yet)
let nextQB = 0;
let nextRB = 0;
let nextWR = 0;
let nextTE = 0;
let nextOT = 0;
let nextOG = 0;
let nextOC = 0;
let nextEDGE = 0;
let nextDI = 0;
let nextLB = 0;
let nextCB = 0;
let nextS = 0;
let nextK = 0;
let nextP = 0;

// Stores picks by team
let picks = {};
picks.BUF = [];
picks.MIA = [];
picks.NE = [];
picks.NYJ = [];
picks.BAL = [];
picks.CIN = [];
picks.CLE = [];
picks.PIT = [];
picks.HOU = [];
picks.IND = [];
picks.JAX = [];
picks.TEN = [];
picks.DEN = [];
picks.KC = [];
picks.LAC = [];
picks.LV = [];
picks.DAL = [];
picks.NYG = [];
picks.PHI = [];
picks.WAS = [];
picks.CHI = [];
picks.DET = [];
picks.GB = [];
picks.MIN = [];
picks.ATL = [];
picks.CAR = [];
picks.NO = [];
picks.TB = [];
picks.ARI = [];
picks.LAR = [];
picks.SEA = [];
picks.SF = [];

// Whether or not the upcoming pick has been simulated yet
let madeUpcomingPick = false;

/* MAIN FUNCTION: Runs both clear & pick functions */
function main() {
  clear()
  pick()
}

/* KEY FUNCTION: Clears the sheet */
function clear() {
  
  // Clear the pick/team specific information on the main sheet
  clearMain();
  
  // Clear each round
  for (let i = 1; i <= 7; i++) {
    clearRound(i);
  }
  
  // After clearing all rounds, clear the picks by team sheet
  SpreadsheetApp.getActive().getSheetByName(`BY TEAM`).getRange(3, 1, 30, 128).clearContent();
}

/* Clears the pick/team specific information on the main sheet */
function clearMain() {
  let sheet = SpreadsheetApp.getActive().getSheetByName(`ALL`);
  sheet.getRange(1, 1, 2, 1).clearContent();
  sheet.getRange(6, 4, 1, 4).clearContent();
  sheet.getRange(9, 1, 1, 1).clearContent();
  sheet.getRange(11, 1, 1, 1).clearContent();
  sheet.getRange(13, 1, 1, 1).clearContent();
  sheet.getRange(15, 1, 1, 1).clearContent();
  sheet.getRange(17, 1, 1, 1).clearContent();
  sheet.getRange(3, 9, 15, 4).clearContent();
  sheet.getRange(18, 1, 300, 1).clearContent();
}

/* Clears the sheet of the given round (roundNum) */
function clearRound(roundNum) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(`R${roundNum}`);
  // sheet.getRange(4, 1, 1, 4 * numPicks[roundNum]).clearContent(); // only use to reset the entire draft
  // sheet.getRange(6, 1, 1, 4 * numPicks[roundNum]).clearContent(); // only use to reset the entire draft
  sheet.getRange(8, 1, 15, 4 * numPicks[roundNum]).clearContent();
}

/* KEY FUNCTION: Simulates all remaining picks in the draft */
function pick() {
  
  // Pick each round
  for (let i = 1; i <= 7; i++) {
    pickRound(i);
  }
  
  // After picking all rounds, update the picks by team sheet
  updatePicksByTeam();
}

/* Simulates all remaining picks in the given round (roundNum) */
function pickRound(roundNum) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(`R${roundNum}`);
  let teams = sheet.getRange(2, 1, 1, numPicks[roundNum] * 4).getValues();
  let actualPicks = sheet.getRange(4, 1, 1, numPicks[roundNum] * 4).getValues();
  for (let i = 1; i <= numPicks[roundNum]; i++) {
    makePick(roundNum, i, sheet, teams, actualPicks);
  }
}

/** Handles the selection of the pick in the given round (roundNum) at the given pick (pickNum)
  * Only simulates the pick if it hasn't been picked in the actual draft yet
  * @param sheet: Reference to the sheet for the given round
  * @param teams: Order of teams picking in the given round
  * @param actualPicks: List of picks already taken in the actual draft
  */
function makePick(roundNum, pickNum, sheet, teams, actualPicks) {
  const team = teams[0][(pickNum - 1) * 4];
  let pickedPOS = actualPicks[0][(pickNum - 1) * 4];
  
  // If an actual pick hasn't been made yet, simulate this pick
  if (pickedPOS == ``) {
    pickedPOS = simPick(roundNum, pickNum, sheet, team);
  }
  
  // Groups interior offensive linemen as the same need
  if (pickedPOS == `OG` || pickedPOS == `OC`) {
    pickedPOS = `IOL`;
  }
  
  // Remove the first occurrence of the selected position from the team's needs
  if (needs[team][0].indexOf(pickedPOS) != -1) {
    needs[team][0].splice(needs[team][0].indexOf(pickedPOS), 1);
  }
  else if (needs[team][1].indexOf(pickedPOS) != -1) {
    needs[team][1].splice(needs[team][1].indexOf(pickedPOS), 1);
  }
  else if (needs[team][2].indexOf(pickedPOS) != -1) {
    needs[team][2].splice(needs[team][2].indexOf(pickedPOS), 1);
  }
  else if (needs[team][3].indexOf(pickedPOS) != -1) {
    needs[team][3].splice(needs[team][3].indexOf(pickedPOS), 1);
  }
  else if (needs[team][4].indexOf(pickedPOS) != -1) {
    needs[team][4].splice(needs[team][4].indexOf(pickedPOS), 1);
  }
}

/** Simulates the pick in the given round (roundNum) at the given pick (pickNum)
  * @param sheet: Reference to the sheet for the given round
  * @param team: The team selecting this pick
  */
function simPick(roundNum, pickNum, sheet, team) {
  
  // Call the helper function to find the best available position for this pick
  let bestAvailable = getBestAvailable(pickNum, sheet, team);
  
  // Set the projected pick to the best available player for this pick
  picks[team].push([`Round ${roundNum} - Pick ${pickNum}`, ``, ``, ``]);
  let pick = [];
  switch (bestAvailable) {
    case `QB`:
      pick = allQB[nextQB];
      nextQB++;
      break;
    case `RB`:
      pick = allRB[nextRB];
      nextRB++;
      break;
    case `WR`:
      pick = allWR[nextWR];
      nextWR++;
      break;
    case `TE`:
      pick = allTE[nextTE];
      nextTE++;
      break;
    case `OT`:
      pick = allOT[nextOT];
      nextOT++;
      break;
    case `OG`:
      pick = allOG[nextOG];
      nextOG++;
      break;
    case `OC`:
      pick = allOC[nextOC];
      nextOC++;
      break;
    case `EDGE`:
      pick = allEDGE[nextEDGE];
      nextEDGE++;
      break;
    case `DI`:
      pick = allDI[nextDI];
      nextDI++;
      break;
    case `LB`:
      pick = allLB[nextLB];
      nextLB++;
      break;
    case `CB`:
      pick = allCB[nextCB];
      nextCB++;
      break;
    case `S`:
      pick = allS[nextS];
      nextS++;
      break;
    case `K`:
      pick = allK[nextK];
      nextK++;
      break;
    case `P`:
      pick = allP[nextP];
      nextP++;
      break;
  }
  
  // Append the pick to the team's list of picks and show the projected pick on the sheet
  picks[team].push(pick);
  sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([pick]);
  
  // If this is the upcoming pick, show the simulated pick on the main sheet and stop further updates to the main sheet
  if (!madeUpcomingPick) {
    simMainPick(roundNum, pickNum, team, [pick]);
    madeUpcomingPick = true;
  }
  
  // Return the best available position picked
  return bestAvailable
}

/** Finds the best available players for this pick (pickNum)
  * Updates the big board with the top 15 best available players
  * Selects the best available player to be picked and returns their position
  * @param sheet: Reference to the sheet for the given round
  * @param team: The team selecting this pick
  */
function getBestAvailable(pickNum, sheet, team) {
  
  // Gets the team's top remaining needs
  let teamNeeds = getTeamNeeds(team);
  
  // Stores the best available players
  let bigBoard = [];
  
  // Loads the best available players for each position of the team's needs
  for (let i = 0; i < teamNeeds.length; i++) {
    switch (teamNeeds[i]) {
      case `QB`:
        for (let i = nextQB; i < nextQB + 5 && i < numQB; i++) {
          bigBoard.push(allQB[i]);
        }
        break;
      case `RB`:
        for (let i = nextRB; i < nextRB + 5 && i < numRB; i++) {
          bigBoard.push(allRB[i]);
        }
        break;
      case `WR`:
        for (let i = nextWR; i < nextWR + 5 && i < numWR; i++) {
          bigBoard.push(allWR[i]);
        }
        break;
      case `TE`:
        for (let i = nextTE; i < nextTE + 5 && i < numTE; i++) {
          bigBoard.push(allTE[i]);
        }
        break;
      case `OT`:
        for (let i = nextOT; i < nextOT + 5 && i < numOT; i++) {
          bigBoard.push(allOT[i]);
        }
        break;
      case `IOL`:
        for (let i = nextOG; i < nextOG + 5 && i < numOG; i++) {
          bigBoard.push(allOG[i]);
        }
        for (let i = nextOC; i < nextOC + 5 && i < numOC; i++) {
          bigBoard.push(allOC[i]);
        }
        break;
      case `EDGE`:
        for (let i = nextEDGE; i < nextEDGE + 5 && i < numEDGE; i++) {
          bigBoard.push(allEDGE[i]);
        }
        break;
      case `DI`:
        for (let i = nextDI; i < nextDI + 5 && i < numDI; i++) {
          bigBoard.push(allDI[i]);
        }
        break;
      case `LB`:
        for (let i = nextLB; i < nextLB + 5 && i < numLB; i++) {
          bigBoard.push(allLB[i]);
        }
        break;
      case `CB`:
        for (let i = nextCB; i < nextCB + 5 && i < numCB; i++) {
          bigBoard.push(allCB[i]);
        }
        break;
      case `S`:
        for (let i = nextS; i < nextS + 5 && i < numS; i++) {
          bigBoard.push(allS[i]);
        }
        break;
      case `K`:
        for (let i = nextK; i < nextK + 5 && i < numK; i++) {
          bigBoard.push(allK[i]);
        }
        break;
      case `P`:
        for (let i = nextP; i < nextP + 5 && i < numP; i++) {
          bigBoard.push(allP[i]);
        }
        break;
    }
  }
  
  // Sorts the best available players by their grades in descending order
  bigBoard.sort(function(x, y) {
    // return y[3] - x[3];
    if (x[3] > y[3]) {
      return -1;
    }
    if (x[3] < y[3]) {
      return 1;
    }
    return 0;
  });

  // Trims the big board to only the top 15 best available players
  if (bigBoard.length > 15) {
    bigBoard.splice(15, bigBoard.length - 15);
  }
  
  // Updates the team big board
  sheet.getRange(8, pickNum * 4 - 3, bigBoard.length, 4).setValues(bigBoard);
  
  // If this is the upcoming pick, show the team big board on the main sheet
  if (!madeUpcomingPick) {
    simMainBigBoard(bigBoard)
  }
  
  // Return the best available position
  return bigBoard[0][0];
}

/* Gets the given team's (team) top remaining needs */
function getTeamNeeds(team) {
  
  // If the team has highest priority needs remaining, return them
  if (needs[team][0].length > 0) {
    return needs[team][0];
  }
  
  // If the team has high priority needs remaining, return them
  if (needs[team][1].length > 0) {
    return needs[team][1];
  }
  
  // If the team has medium priority needs remaining, return them
  if (needs[team][2].length > 0) {
    return needs[team][2];
  }
  
  // If the team has low priority needs remaining, return them
  if (needs[team][3].length > 0) {
    return needs[team][3];
  }
  
  // If the team has lowest priority needs remaining, return them
  if (needs[team][4].length > 0) {
    return needs[team][4];
  }
  
  // If the team has no needs remaining, return all positions and select best available
  return [`QB`, `RB`, `WR`, `TE`, `OT`, `IOL`, `EDGE`, `DI`, `LB`, `CB`, `S`, `K`, `P`];
}

/* Updates the picks by team sheet with each team's picks */
function updatePicksByTeam() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(`BY TEAM`);
  const teams = sheet.getRange(2, 1, 1, 128).getValues();
  for (let i = 0; i < 128; i += 4) {
    const team = teams[0][i];
    const teamPicks = picks[team];
    sheet.getRange(3, i + 1, teamPicks.length, 4).setValues(teamPicks);
  }
}

/* Updates the left side of the main sheet */
function simMainPick(roundNum, pickNum, team, pick) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(`ALL`);
  sheet.getRange(1, 1).setValue(`ROUND ${roundNum}  -  PICK ${pickNum}`);
  sheet.getRange(2, 1).setValue(team);
  sheet.getRange(6, 4, 1, 4).setValues(pick);
  sheet.getRange(6, 1, 1, 1).setValue(`--->`);
  sheet.getRange(9, 1, 1, 1).setValue(needsToStr(needs[team][0]));
  sheet.getRange(11, 1, 1, 1).setValue(needsToStr(needs[team][1]));
  sheet.getRange(13, 1, 1, 1).setValue(needsToStr(needs[team][2]));
  sheet.getRange(15, 1, 1, 1).setValue(needsToStr(needs[team][3]));
  sheet.getRange(17, 1, 1, 1).setValue(needsToStr(needs[team][4]));
  let currentPickRow = 19 + pickNum;
  for (let i = 1; i < roundNum; i++) {
    currentPickRow += numPicks[i] + 2;
  }
  sheet.getRange(currentPickRow, 1).setValue(`--->`);
}

/* Updates the right side of the main sheet */
function simMainBigBoard(bigBoard) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(`ALL`);
  sheet.getRange(3, 9, bigBoard.length, 4).setValues(bigBoard);
}

/* Converts needs array to the string to show on the main sheet */
function needsToStr(needs) {
  let str = ``;
  for (let i = 0; i < needs.length; i++) {
    str += needs[i];
    if (i < needs.length - 1) {
      str += `  `;
    }
  }
  return str;
}
