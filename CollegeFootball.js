// UPDATE EACH WEEK:
const numWeeksDone = 0; // equivalent to the upcoming week #

// UPDATE AS NEEDED:
const homefieldAdvantage = 6.66;

// UPDATE EACH SEASON:
const confSizes = {AAC:11, ACC:14, B10:14, B12:10, CUSA:14, MAC:12, MWC:12, P12:12, SBC:10, SEC:14, IND:7};
const numWeeks = 18;
const rowGroupOf5 = 19; // on the Standings sheet
const numTeams = 130;

/* MAIN FUNCTION: Runs the simulation */
function main() {
  
  let schedule = {};
  let standings = {};
  let ratings = {};
  
  clear(numWeeks, numWeeksDone, rowGroupOf5, "AAC", confSizes.AAC);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "ACC", confSizes.ACC);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "B10", confSizes.B10);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "B12", confSizes.B12);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "CUSA", confSizes.CUSA);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "MAC", confSizes.MAC);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "MWC", confSizes.MWC);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "P12", confSizes.P12);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "SBC", confSizes.SBC);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "SEC", confSizes.SEC);
  clear(numWeeks, numWeeksDone, rowGroupOf5, "IND", confSizes.IND);
  
  loadRatings(ratings, numTeams);
  
  loadGames(schedule, standings, numWeeks, numWeeksDone, "AAC", confSizes.AAC);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "ACC", confSizes.ACC);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "B10", confSizes.B10);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "B12", confSizes.B12);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "CUSA", confSizes.CUSA);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "MAC", confSizes.MAC);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "MWC", confSizes.MWC);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "P12", confSizes.P12);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "SBC", confSizes.SBC);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "SEC", confSizes.SEC);
  loadGames(schedule, standings, numWeeks, numWeeksDone, "IND", confSizes.IND);
  
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "AAC", confSizes.AAC);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "ACC", confSizes.ACC);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "B10", confSizes.B10);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "B12", confSizes.B12);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "CUSA", confSizes.CUSA);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "MAC", confSizes.MAC);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "MWC", confSizes.MWC);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "P12", confSizes.P12);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "SBC", confSizes.SBC);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "SEC", confSizes.SEC);
  simGames(schedule, standings, ratings, numWeeks, numWeeksDone, "IND", confSizes.IND);
  
  calculateRecords(standings, numWeeks, "AAC", confSizes.AAC);
  calculateRecords(standings, numWeeks, "ACC", confSizes.ACC);
  calculateRecords(standings, numWeeks, "B10", confSizes.B10);
  calculateRecords(standings, numWeeks, "B12", confSizes.B12);
  calculateRecords(standings, numWeeks, "CUSA", confSizes.CUSA);
  calculateRecords(standings, numWeeks, "MAC", confSizes.MAC);
  calculateRecords(standings, numWeeks, "MWC", confSizes.MWC);
  calculateRecords(standings, numWeeks, "P12", confSizes.P12);
  calculateRecords(standings, numWeeks, "SBC", confSizes.SBC);
  calculateRecords(standings, numWeeks, "SEC", confSizes.SEC);
  calculateRecords(standings, numWeeks, "IND", confSizes.IND);
  
  generateStandings(standings, numWeeks, "AAC", confSizes.AAC);
  generateStandings(standings, numWeeks, "ACC", confSizes.ACC);
  generateStandings(standings, numWeeks, "B10", confSizes.B10);
  generateStandings(standings, numWeeks, "B12", confSizes.B12);
  generateStandings(standings, numWeeks, "CUSA", confSizes.CUSA);
  generateStandings(standings, numWeeks, "MAC", confSizes.MAC);
  generateStandings(standings, numWeeks, "MWC", confSizes.MWC);
  generateStandings(standings, numWeeks, "P12", confSizes.P12);
  generateStandings(standings, numWeeks, "SBC", confSizes.SBC);
  generateStandings(standings, numWeeks, "SEC", confSizes.SEC);
  generateStandings(standings, numWeeks, "IND", confSizes.IND);
}

/* UTILITY FUNCTION: Sorts team ratings in descending order automatically */
function onEdit(e) {
  const SHEET_NAME = "Ratings";
  const SORT_DATA_RANGE = "C2:D131";
  const SORT_ORDER = [{column: 4, ascending: false}];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const range = sheet.getRange(SORT_DATA_RANGE);
  range.sort(SORT_ORDER);
}

/* Clears any previous simulated material from the spreadsheet to reset it */
function clear(numWeeks, numWeeksDone, rowGroupOf5, conf, confSize) {
  
  // Clear schedule sheets
  const sheetSchedule = SpreadsheetApp.getActive().getSheetByName(conf);
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Standings");
  if (numWeeksDone < 18) {
    const start = numWeeksDone + 4;
    let clearedRecords = [];
    for (let i = 0; i < confSize; i++) {
      clearedRecords.push("--");
    }
    clearedRecords = [clearedRecords];
    sheetSchedule.getRange(3, 2, 1, confSize).setValues(clearedRecords);
    sheetSchedule.getRange(start, 2, numWeeks, confSize).setBackground("white");
  }
  
  // Clear Standings sheet
  let range0 = "";
  let range1 = "";
  switch (conf) {
    case "AAC":
      range0 = "A" + (rowGroupOf5 + 1) + ":C" + (rowGroupOf5 + confSize);
      break;
    case "ACC":
      range0 = "A3:C" + (2 + confSize / 2);
      range1 = "A" + (4 + confSize / 2) + ":C" + (3 + confSize);
      break;
    case "B10":
      range0 = "D3:F" + (2 + confSize / 2);
      range1 = "D" + (4 + confSize / 2) + ":F" + (3 + confSize);
      break;
    case "B12":
      range0 = "G2:I" + (1 + confSize);
      break;
    case "CUSA":
      range0 = "D" + (rowGroupOf5 + 2) + ":F" + (rowGroupOf5 + 1 + confSize / 2);
      range1 = "D" + (rowGroupOf5 + 3 + confSize / 2) + ":F" + (rowGroupOf5 + 2 + confSize);
      break;
    case "MAC":
      range0 = "G" + (rowGroupOf5 + 2) + ":I" + (rowGroupOf5 + 1 + confSize / 2);
      range1 = "G" + (rowGroupOf5 + 3 + confSize / 2) + ":I" + (rowGroupOf5 + 2 + confSize);
      break;
    case "MWC":
      range0 = "J" + (rowGroupOf5 + 2) + ":L" + (rowGroupOf5 + 1 + confSize / 2);
      range1 = "J" + (rowGroupOf5 + 3 + confSize / 2) + ":L" + (rowGroupOf5 + 2 + confSize);
      break;
    case "P12":
      range0 = "J3:L" + (2 + confSize / 2);
      range1 = "J" + (4 + confSize / 2) + ":L" + (3 + confSize);
      break;a
    case "SBC":
      range0 = "M" + (rowGroupOf5 + 2) + ":O" + (rowGroupOf5 + 1 + confSize / 2);
      range1 = "M" + (rowGroupOf5 + 3 + confSize / 2) + ":O" + (rowGroupOf5 + 2 + confSize);
      break;
    case "SEC":
      range0 = "M3:O" + (2 + confSize / 2);
      range1 = "M" + (4 + confSize / 2) + ":O" + (3 + confSize);
      break;
    case "IND":
      range0 = "P2:R" + (1 + confSize);
      break;
  }
  sheetStandings.getRange(range0).setValue("");
  if (range1 != "") {
    sheetStandings.getRange(range1).setValue("");
  }
}

/* Loads all ratings for each team */
function loadRatings(ratings, numTeams) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Ratings");
  const ratingsData = sheet.getRange(2, 1, numTeams, 2).getValues();
  for (let i = 0; i < numTeams; i++) {
    const team = ratingsData[i][0];
    const rating = ratingsData[i][1];
    ratings[team] = rating;
  }
}

/* Loads all remaining games into each team's schedule, and loads each team into each conference's standings */
function loadGames(schedule, standings, numWeeks, numWeeksDone, conf, confSize) {
  if (numWeeksDone >= numWeeks) {
    return;
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  let teams = sheet.getRange(2, 2, 1, confSize).getValues();
  const games = sheet.getRange(numWeeksDone + 4, 2, numWeeks - numWeeksDone, confSize).getValues();
  standings[conf] = [];
  let confStandings = standings[conf];
  for (let c = 0; c < confSize; c++) {
    const team = teams[0][c];
    schedule[team] = [];
    for (let r = 0; r < games.length; r++) {
      schedule[team].push(games[r][c]);
    }
    confStandings[team] = {};
  }
}

/* Simulates every remaining game on every team's schedule */
function simGames(schedule, standings, ratings, numWeeks, numWeeksDone, conf, confSize) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, confSize).getValues();
  for (let c = 0; c < confSize; c++) {
    const team = teams[0][c];
    const teamSchedule = schedule[team];
    standings[conf][team] = {};
    for (let r = 0; r < teamSchedule.length; r++) {
      let opponent = teamSchedule[r].toUpperCase();
      if (opponent === "--") {
        continue;
      }
      let gameType = 0; // team is home
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
        gameType = 1; // team is away
      }
      else if (opponent.substring(0, 3) === "VS.") {
        opponent = opponent.substring(4, opponent.length);
        gameType = 2; // team is neutral
      }
      
      // Team is playing an FCS team (mark as a win)
      if (!(opponent in schedule)) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("limegreen");
        continue;
      }
      
      let advantage = simGame(ratings, gameType, team, opponent);
      
      // Team lost the match
      if (advantage < 0) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("red");
      }
      
      // Team won the match
      else if (advantage > 0) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("limegreen");
      }
      
      // Match is dead even - flag as yellow to manually pick later
      else {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("yellow");
      }
    }
  }
}

/* Simulates a certain given by the parameters */
function simGame(ratings, gameType, team, opponent) {
  let advantage = ratings[team] - ratings[opponent];
  if (gameType === 0) {
    advantage += homefieldAdvantage;
  }
  if (gameType === 1) {
    advantage -= homefieldAdvantage;
  }
  if (advantage == 0) {
    if (ratings[team] != ratings[opponent]) {
      advantage = ratings[team] - ratings[opponent];
    }
    else if (gameType != 2) {
      advantage = gameType === 0 ? 1 : -1;
    }
  }
  return advantage;
}

/* Calculates each team's record from its entire schedule */
function calculateRecords(standings, numWeeks, conf, confSize) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, confSize).getValues();
  const results = sheet.getRange(4, 2, numWeeks, 16).getValues();
  let records = [];
  for (let c = 0; c < confSize; c++) {
    const team = teams[0][c];
    let numWinsOVR = 0;
    let numLossesOVR = 0;
    let numWinsCONF = 0;
    let numLossesCONF = 0;
    for (let r = 0; r < numWeeks; r++) {
      const result = sheet.getRange(r + 4, c + 2).getBackgroundColor();
      
      // Match was actually a bye week
      if (result === "#ffffff") {
        continue;
      }
      
      // Parse opposing team name
      let opponent = results[r][c].toUpperCase();
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "VS.") {
        opponent = opponent.substring(4, opponent.length);
      }
      
      // Team won the match
      if (result === "#ff0000") {
        numLossesOVR++;
        if (opponent in standings[conf]) {
          numLossesCONF++;
        }
      }
      
      // Team lost the match
      else if (result === "#32cd32") {
        numWinsOVR++;
        if (opponent in standings[conf]) {
          numWinsCONF++;
        }
      }
    }
    
    // Don't show conference record for independent teams
    if (conf === "IND") {
      numWinsCONF = 0;
      numLossesCONF = 0;
      records.push(numWinsOVR + "-" + numLossesOVR);
    }
    else {
      records.push(numWinsOVR + "-" + numLossesOVR + " (" + numWinsCONF + "-" + numLossesCONF + ")");
    }
    
    standings[conf][team]["OVR Wins"] = numWinsOVR;
    standings[conf][team]["OVR Losses"] = numLossesOVR;
    standings[conf][team]["CONF Wins"] = numWinsCONF;
    standings[conf][team]["CONF Losses"] = numLossesCONF;
    
    // Don't show divisions for independents or conferences without divisions
    if (conf === "AAC" || conf === "B12" || conf === "IND") {
      standings[conf][team]["Division"] = -1;
    }
    else {
      if (c < confSize / 2) {
        standings[conf][team]["Division"] = 0;
      }
      else {
        standings[conf][team]["Division"] = 1;
      }
    }
  }
  
  records = [records];
  sheet.getRange(3, 2, 1, confSize).setValues(records);
}

/* Generates each conference's standings based on each of its team's records */
function generateStandings(standings, numWeeks, conf, confSize) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Standings");
  let confStandings = standings[conf];
  let confTeams = Object.keys(confStandings);
  confTeams.sort(sortByRecord(confStandings));
  if (conf === "AAC" || conf === "B12" || conf === "IND") {
    let row = 2;
    let col = 1;
    if (conf === "AAC") {
      row = 20;
    }
    if (conf === "B12") {
      col = 7;
    }
    if (conf === "IND") {
      col = 16;
    }
    let confData = [];
    for (let i = 0; i < confTeams.length; i++) {
      const team = confTeams[i];
      const teamRecords = confStandings[team];
      let arr = [];
      arr.push(team);
      arr.push(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
      arr.push(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
      confData.push(arr);
    }
    sheet.getRange(row, col, confSize, 3).setValues(confData);
  }
  else {
    let divARow = 3;
    if (conf === "CUSA" || conf === "MAC" || conf === "MWC" || conf === "SBC") {
      divARow = 21;
    }
    let divBRow = divARow + confSize / 2 + 1;
    let col = 1;
    if (conf === "B10" || conf === "CUSA") {
      col = 4;
    }
    if (conf === "MAC") {
      col = 7;
    }
    if (conf === "P12" || conf === "MWC") {
      col = 10;
    }
    if (conf === "SEC" || conf === "SBC") {
      col = 13;
    }
    let divAData = [];
    let divBData = [];
    for (let i = 0; i < confTeams.length; i++) {
      const team = confTeams[i];
      const teamRecords = confStandings[team];
      if (teamRecords["Division"] === 0) {
        let arr = [];
        arr.push(team);
        arr.push(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
        arr.push(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
        divAData.push(arr);
      }
      else {
        let arr = [];
        arr.push(team);
        arr.push(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
        arr.push(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
        divBData.push(arr);
      }
    }
    sheet.getRange(divARow, col, confSize / 2, 3).setValues(divAData);
    sheet.getRange(divBRow, col, confSize / 2, 3).setValues(divBData);
  }
}

/* Sorts the given conference's standings based on conference wins (if tied, overall wins) */
function sortByRecord(confStandings) {
  return function(a, b) {
    if (confStandings[a]["CONF Wins"] === confStandings[b]["CONF Wins"]) {
      return (confStandings[b]["OVR Wins"] - confStandings[b]["OVR Losses"]) - (confStandings[a]["OVR Wins"] - confStandings[a]["OVR Losses"]);
    }
    return confStandings[b]["CONF Wins"] - confStandings[a]["CONF Wins"];
  };
}
