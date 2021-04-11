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

// UPDATE EACH YEAR: Team needs (highest priority, high priority, low priority, lowest priority)
let needs = {};
needs.BUF = [[], [`CB`, `EDGE`, `WR`], [`RB`, `DI`, `TE`, `CB`], [`OT`, `IOL`, `S`, `LB`, `QB`]];
needs.MIA = [[`WR`], [`EDGE`], [`RB`, `TE`, `LB`, `IOL`, `WR`, `DI`, `S`, `LB`, `P`], [`OT`, `QB`]];
needs.NE = [[`QB`], [`LB`, `WR`], [`CB`, `S`, `WR`, `RB`, `DI`], [`CB`, `EDGE`, `IOL`, `OT`]];
needs.NYJ = [[], [`QB`], [`CB`, `EDGE`, `WR`, `TE`, `RB`, `LB`, `IOL`, `OT`, `K`], [`WR`, `EDGE`, `S`, `CB`]];
needs.BAL = [[`WR`], [`EDGE`, `S`], [`OT`, `EDGE`, `TE`, `IOL`, `WR`], [`LB`, `CB`, `DI`]];
needs.CIN = [[], [`WR`, `OT`], [`TE`, `IOL`, `EDGE`, `IOL`, `LB`, `CB`], [`S`, `LB`, `QB`]];
needs.CLE = [[], [`EDGE`, `DI`, `WR`, `CB`], [`LB`, `IOL`], [`OT`, `S`, `LB`, `DI`, `CB`, `TE`, `QB`, `K`]];
needs.PIT = [[], [`CB`, `OT`, `RB`, `IOL`], [`LB`, `EDGE`, `DI`, `QB`], [`TE`, `WR`, `DI`]];
needs.HOU = [[], [`CB`, `DI`, `WR`, `IOL`, `TE`, `EDGE`], [`LB`, `S`, `K`], [`WR`, `OT`]];
needs.IND = [[`OT`], [`CB`, `WR`, `EDGE`], [`DI`, `LB`, `TE`], [`QB`, `S`]];
needs.JAX = [[`QB`], [`S`], [`OT`, `WR`, `RB`, `CB`, `TE`, `DI`, `EDGE`], [`WR`, `IOL`]];
needs.TEN = [[], [`CB`, `WR`, `OT`, `S`, `TE`, `EDGE`], [`LB`, `DI`, `WR`, `DI`], [`IOL`, `QB`]];
needs.DEN = [[`QB`], [`LB`, `CB`], [`OT`, `IOL`, `DI`, `WR`, `EDGE`, `RB`], [`LB`, `S`, `CB`, `OT`]];
needs.KC = [[`OT`], [`WR`, `LB`, `IOL`, `EDGE`], [`IOL`, `DI`, `TE`], [`CB`]];
needs.LAC = [[`OT`], [`EDGE`, `S`, `CB`], [`WR`, `IOL`, `LB`, `TE`, `IOL`, `OT`, `K`], [`CB`, `RB`, `DI`]];
needs.LV = [[`OT`], [`LB`, `IOL`, `S`, `DI`], [`CB`, `EDGE`, `TE`], [`RB`, `CB`, `WR`, `QB`]];
needs.DAL = [[`CB`, `DI`, `OT`, `EDGE`], [`TE`, `IOL`, `S`], [`CB`, `P`], [`LB`, `QB`]];
needs.NYG = [[], [`EDGE`, `LB`, `CB`, `WR`, `OT`], [`IOL`, `EDGE`, `LB`, `TE`], [`QB`, `S`, `IOL`, `RB`]];
needs.PHI = [[`WR`, `CB`, `S`], [`TE`, `LB`, `IOL`, `OT`, `LB`], [`CB`, `P`], [`DI`, `QB`, `EDGE`, `RB`]];
needs.WAS = [[], [`OT`, `LB`, `S`, `CB`], [`WR`, `QB`, `IOL`, `CB`, `TE`], [`WR`, `RB`, `DI`]];
needs.CHI = [[`OT`], [`QB`, `WR`, `CB`, `IOL`], [`WR`, `DI`, `EDGE`], [`TE`, `LB`, `RB`]];
needs.DET = [[`WR`], [`CB`, `OT`, `LB`], [`QB`, `DI`, `DI`], [`WR`, `S`, `EDGE`, `K`]];
needs.GB = [[], [`OT`, `WR`, `CB`, `LB`], [`DI`, `IOL`, `WR`], [`S`, `TE`, `IOL`, `LB`, `RB`]];
needs.MIN = [[`IOL`], [`S`, `CB`, `EDGE`, `OT`], [`WR`, `IOL`, `DI`, `QB`, `K`, `P`], [`LB`, `TE`]];
needs.ATL = [[`TE`], [`QB`, `EDGE`, `OT`], [`LB`, `RB`, `S`, `CB`, `WR`], [`DI`, `S`]];
needs.CAR = [[`OT`], [`CB`], [`TE`, `QB`, `LB`, `IOL`, `S`, `DI`, `CB`], [`IOL`, `RB`, `EDGE`]];
needs.NO = [[], [`CB`, `WR`, `LB`, `DI`, `TE`], [`EDGE`, `S`], [`QB`, `OT`, `LB`]];
needs.TB = [[], [`DI`, `RB`, `EDGE`, `WR`], [`IOL`, `OT`, `QB`], [`DI`, `LB`, `OT`, `S`]];
needs.ARI = [[`CB`], [`RB`, `IOL`], [`TE`, `LB`, `CB`, `DI`, `IOL`, `WR`], [`WR`, `EDGE`]];
needs.LAR = [[], [`OT`, `IOL`, `TE`, `LB`, `EDGE`], [`CB`, `WR`, `OT`], [`IOL`, `LB`]];
needs.SEA = [[], [`EDGE`, `OT`, `CB`], [`IOL`, `DI`, `WR`], [`WR`, `S`, `TE`]];
needs.SF = [[`QB`], [`CB`], [`IOL`, `OT`, `S`, `CB`, `EDGE`], [`WR`, `RB`, `LB`]];

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

/* MAIN FUNCTION: Clears the sheet */
function clear() {
  for (let i = 1; i <= 7; i++) {
    clearRound(i);
  }
  
  // After clearing all rounds, clear the picks by team sheet
  SpreadsheetApp.getActive().getSheetByName(`PICKS BY TEAM`).getRange(3, 1, 30, 128).clearContent();
}

/* Clears the sheet of the given round (roundNum) */
function clearRound(roundNum) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(`R${roundNum}`);
  sheet.getRange(4, 1, 1, 4 * numPicks[roundNum]).clearContent();
  sheet.getRange(6, 1, 1, 4 * numPicks[roundNum]).clearContent();
  sheet.getRange(8, 1, 15, 4 * numPicks[roundNum]).clearContent();
}

/* MAIN FUNCTION: Simulates all remaining picks in the draft */
function pick() {
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
  if (pickedPOS == "") {
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
}

/** Simulates the pick in the given round (roundNum) at the given pick (pickNum)
  * @param sheet: Reference to the sheet for the given round
  * @param team: The team selecting this pick
  */
function simPick(roundNum, pickNum, sheet, team) {
  
  // Call the helper function to find the best available player for this pick
  let bestAvailable = getBestAvailable(pickNum, sheet, team);
  
  // Set the projected pick to the best available player for this pick
  picks[team].push([`Round ${roundNum} - Pick ${pickNum}`, ``, ``, ``]);
  switch (bestAvailable) {
    case `QB`:
      picks[team].push(allQB[nextQB]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allQB[nextQB]]);
      nextQB++;
      return `QB`;
    case `RB`:
      picks[team].push(allRB[nextRB]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allRB[nextRB]]);
      nextRB++;
      return `RB`;
    case `WR`:
      picks[team].push(allWR[nextWR]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allWR[nextWR]]);
      nextWR++;
      return `WR`;
    case `TE`:
      picks[team].push(allTE[nextTE]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allTE[nextTE]]);
      nextTE++;
      return `TE`;
    case `OT`:
      picks[team].push(allOT[nextOT]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allOT[nextOT]]);
      nextOT++;
      return `OT`;
    case `OG`:
      picks[team].push(allOG[nextOG]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allOG[nextOG]]);
      nextOG++;
      return `OG`;
    case `OC`:
      picks[team].push(allOC[nextOC]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allOC[nextOC]]);
      nextOC++;
      return `OC`;
    case `EDGE`:
      picks[team].push(allEDGE[nextEDGE]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allEDGE[nextEDGE]]);
      nextEDGE++;
      return `EDGE`;
    case `DI`:
      picks[team].push(allDI[nextDI]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allDI[nextDI]]);
      nextDI++;
      return `DI`;
    case `LB`:
      picks[team].push(allLB[nextLB]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allLB[nextLB]]);
      nextLB++;
      return `LB`;
    case `CB`:
      picks[team].push(allCB[nextCB]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allCB[nextCB]]);
      nextCB++;
      return `CB`;
    case `S`:
      picks[team].push(allS[nextS]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allS[nextS]]);
      nextS++;
      return `S`;
    case `K`:
      picks[team].push(allK[nextK]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allK[nextK]]);
      nextK++;
      return `K`;
    case `P`:
      picks[team].push(allP[nextP]);
      sheet.getRange(6, pickNum * 4 - 3, 1, 4).setValues([allP[nextP]]);
      nextP++;
      return `P`;
  }
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
  
  // Updates the big board and returns the #1 player's position from it
  sheet.getRange(8, pickNum * 4 - 3, bigBoard.length, 4).setValues(bigBoard);
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
  
  // If the team has low priority needs remaining, return them
  if (needs[team][2].length > 0) {
    return needs[team][2];
  }
  
  // If the team has lowest priority needs remaining, return them
  if (needs[team][3].length > 0) {
    return needs[team][3];
  }
  
  // If the team has no needs remaining, return all positions and select best available
  return [`QB`, `RB`, `WR`, `TE`, `OT`, `IOL`, `EDGE`, `DI`, `LB`, `CB`, `S`, `K`, `P`];
}

/* Updates the picks by team sheet with each team's picks */
function updatePicksByTeam() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(`PICKS BY TEAM`);
  const teams = sheet.getRange(2, 1, 1, 128).getValues();
  for (let i = 0; i < 128; i += 4) {
    const team = teams[0][i];
    const teamPicks = picks[team];
    sheet.getRange(3, i + 1, teamPicks.length, 4).setValues(teamPicks);
  }
}
