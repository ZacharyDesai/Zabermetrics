// UPDATE EACH WEEK:
const numWeeksDone = 18; // equivalent to the upcoming week #

// UPDATE AS NEEDED:
const homefieldAdvantage = 3.33;

// UPDATE EACH SEASON:
const confSizes = {"AAC":11, "ACC":14, "B10":14, "B12":10, "CUSA":11, "MAC":12, "MWC":12, "P12":12, "SBC":14, "SEC":14, "IND":7};
const numWeeks = 18;
const rowGroupOf5 = 19; // on the Standings sheet
const numTeams = 131;

/* MAIN FUNCTION: Runs the simulation */
function main() {
  
  let schedule = {};
  let standings = {};
  let ratings = {};
  
  clear("AAC");
  clear("ACC");
  clear("B10");
  clear("B12");
  clear("CUSA");
  clear("MAC");
  clear("MWC");
  clear("P12");
  clear("SBC");
  clear("SEC");
  clear("IND");
  
  loadRatings(ratings);
  
  loadGames(schedule, standings, "AAC");
  loadGames(schedule, standings, "ACC");
  loadGames(schedule, standings, "B10");
  loadGames(schedule, standings, "B12");
  loadGames(schedule, standings, "CUSA");
  loadGames(schedule, standings, "MAC");
  loadGames(schedule, standings, "MWC");
  loadGames(schedule, standings, "P12");
  loadGames(schedule, standings, "SBC");
  loadGames(schedule, standings, "SEC");
  loadGames(schedule, standings, "IND");
  
  simGames(schedule, standings, ratings, "AAC");
  simGames(schedule, standings, ratings, "ACC");
  simGames(schedule, standings, ratings, "B10");
  simGames(schedule, standings, ratings, "B12");
  simGames(schedule, standings, ratings, "CUSA");
  simGames(schedule, standings, ratings, "MAC");
  simGames(schedule, standings, ratings, "MWC");
  simGames(schedule, standings, ratings, "P12");
  simGames(schedule, standings, ratings, "SBC");
  simGames(schedule, standings, ratings, "SEC");
  simGames(schedule, standings, ratings, "IND");
  
  calculateRecords(standings, "AAC");
  calculateRecords(standings, "ACC");
  calculateRecords(standings, "B10");
  calculateRecords(standings, "B12");
  calculateRecords(standings, "CUSA");
  calculateRecords(standings, "MAC");
  calculateRecords(standings, "MWC");
  calculateRecords(standings, "P12");
  calculateRecords(standings, "SBC");
  calculateRecords(standings, "SEC");
  calculateRecords(standings, "IND");
  
  generateStandings(standings, "AAC");
  generateStandings(standings, "ACC");
  generateStandings(standings, "B10");
  generateStandings(standings, "B12");
  generateStandings(standings, "CUSA");
  generateStandings(standings, "MAC");
  generateStandings(standings, "MWC");
  generateStandings(standings, "P12");
  generateStandings(standings, "SBC");
  generateStandings(standings, "SEC");
  generateStandings(standings, "IND");
}

/* UTILITY FUNCTION: Use to check if the schedule was inputted correctly */
function checkSchedule() {
  
  let schedule = {};
  let standings = {};
  
  clear("AAC");
  clear("ACC");
  clear("B10");
  clear("B12");
  clear("CUSA");
  clear("MAC");
  clear("MWC");
  clear("P12");
  clear("SBC");
  clear("SEC");
  clear("IND");
  
  loadGames(schedule, standings, "AAC");
  loadGames(schedule, standings, "ACC");
  loadGames(schedule, standings, "B10");
  loadGames(schedule, standings, "B12");
  loadGames(schedule, standings, "CUSA");
  loadGames(schedule, standings, "MAC");
  loadGames(schedule, standings, "MWC");
  loadGames(schedule, standings, "P12");
  loadGames(schedule, standings, "SBC");
  loadGames(schedule, standings, "SEC");
  loadGames(schedule, standings, "IND");
  
  checkScheduleHelper(schedule, "AAC");
  checkScheduleHelper(schedule, "ACC");
  checkScheduleHelper(schedule, "B10");
  checkScheduleHelper(schedule, "B12");
  checkScheduleHelper(schedule, "CUSA");
  checkScheduleHelper(schedule, "MAC");
  checkScheduleHelper(schedule, "MWC");
  checkScheduleHelper(schedule, "P12");
  checkScheduleHelper(schedule, "SBC");
  checkScheduleHelper(schedule, "SEC");
  checkScheduleHelper(schedule, "IND");
}

// Checks each match in schedule to make sure each opponent is an actual team
function checkScheduleHelper(schedule, conf) {
  const confSize = confSizes[conf];
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, confSize).getValues();
  for (let c = 0; c < confSize; c++) {
    const team = teams[0][c];
    const teamSchedule = schedule[team];
    for (let r = 0; r < teamSchedule.length; r++) {
      let opponent = teamSchedule[r].toUpperCase();
      
      // Mark byes as yellow
      if (opponent === "--") {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("yellow");
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
      
      // Mark games against FCS teams as green
      if (!(opponent in schedule)) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#93c47d");
      }
      
      // Flag mismatching games against FBS teams as red
      else {
        const opponentSchedule = schedule[opponent];
        let opponentOpponent = opponentSchedule[r].toUpperCase();
        if (gameType === 0) {
          opponentOpponent = opponentOpponent.substring(2, opponentOpponent.length);
        }
        else if (gameType === 2) {
          opponentOpponent = opponentOpponent.substring(4, opponentOpponent.length);
        }
        if (team != opponentOpponent) {
          sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#e06666");
        }
      }
    }
  }
}

/* UTILITY FUNCTION: Sorts team ratings in descending order automatically */
function onEdit(e) {
  const SHEET_NAME = "Ratings";
  const SORT_DATA_RANGE = "M2:N132";
  const SORT_ORDER = [{column: 14, ascending: false}];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const range = sheet.getRange(SORT_DATA_RANGE);
  range.sort(SORT_ORDER);
}

/* Clears any previous simulated material from the spreadsheet to reset it */
function clear(conf) {
  const confSize = confSizes[conf];
  
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
      range0 = "D" + (rowGroupOf5 + 1) + ":F" + (rowGroupOf5 + confSize);
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
      range0 = "P2:Q" + (1 + confSize);
      break;
  }
  sheetStandings.getRange(range0).setValue("");
  if (range1 != "") {
    sheetStandings.getRange(range1).setValue("");
  }
}

/* Loads all ratings for each team */
function loadRatings(ratings) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Ratings");
  const ratingsData = sheet.getRange(2, 1, numTeams, 2).getValues();
  for (let i = 0; i < numTeams; i++) {
    const team = ratingsData[i][0];
    const rating = ratingsData[i][1];
    ratings[team] = rating;
  }
}

/* Loads all remaining games into each team's schedule, and loads each team into each conference's standings */
function loadGames(schedule, standings, conf) {
  const confSize = confSizes[conf];
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  let teams = sheet.getRange(2, 2, 1, confSize).getValues();
  let games = [];
  if (numWeeksDone < numWeeks) {
    games = sheet.getRange(numWeeksDone + 4, 2, numWeeks - numWeeksDone, confSize).getValues();
  }
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
function simGames(schedule, standings, ratings, conf) {
  const confSize = confSizes[conf];
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, confSize).getValues();
  for (let c = 0; c < confSize; c++) {
    const team = teams[0][c];
    const teamSchedule = schedule[team];
    standings[conf][team] = {};

    // UNCOMMENT TO SIM
    break;

    for (let r = 0; r < teamSchedule.length; r++) {
      let opponent = teamSchedule[r].toUpperCase();
      
      // If match is a bye week, skip it
      if (opponent === "--") {
        continue;
      }
      
      // If match has the hidden special ! character, team lost the match (manual pick)
      if (opponent.charAt(0) === "!") {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#e06666");
        // Print !
        // console.log(team + " - " + opponent);
        continue;
      }
      
      // If match has the hidden special $ character, team won the match (manual pick)
      if (opponent.charAt(0) === "$") {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#93c47d");
        // Print $
        // console.log(team + " - " + opponent);
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
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#93c47d");
        continue;
      }
      
      let advantage = simGame(ratings, gameType, team, opponent);
      
      // Print match
      let advantageVal = Math.round(advantage * -10) / 10;
      let advantageStr = advantageVal < 0 ? `${advantageVal}` : `+${advantageVal}`;
      // console.log(`${team} (${advantageStr}) - ${teamSchedule[r]}`);
      
      // Team lost the match
      if (advantage < 0) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#e06666");
      }
      
      // Team won the match
      else if (advantage > 0) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("#93c47d");
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
function calculateRecords(standings, conf) {
  const confSize = confSizes[conf];
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
    // ZD_HERE
    // for (let r = 0; r < 14; r++) {
    for (let r = 0; r < numWeeks; r++) {
      const result = sheet.getRange(r + 4, c + 2).getBackgroundColor();
      
      // Match was actually a bye week
      if (result === "#ffffff") {
        continue;
      }
      
      // Parse opposing team name
      let opponent = results[r][c].toUpperCase();
      if (opponent.charAt(0) === '!' || opponent.charAt(0) === '$') {
        opponent = opponent.substring(1, opponent.length);
      }
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "VS.") {
        opponent = opponent.substring(4, opponent.length);
      }

      // Team lost the match
      if (result === "#e06666") {
        numLossesOVR++;
        if (opponent in standings[conf]) {
          numLossesCONF++;
        }
      }
      
      // Team won the match
      else if (result === "#93c47d") {
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
    if (conf === "AAC" || conf === "B12" || conf === "CUSA" || conf === "IND") {
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
function generateStandings(standings, conf) {
  const confSize = confSizes[conf];
  const sheet = SpreadsheetApp.getActive().getSheetByName("Standings");
  let confStandings = standings[conf];
  let confTeams = Object.keys(confStandings);
  confTeams.sort(sortByRecord(confStandings));
  if (conf === "AAC" || conf === "B12" || conf === "CUSA" || conf === "IND") {
    let row;
    let col;
    if (conf === "AAC") {
      row = 20;
      col = 1;
    }
    if (conf === "B12") {
      row = 2;
      col = 7;
    }
    if (conf === "CUSA") {
      row = 20;
      col = 4;
    }
    if (conf === "IND") {
      row = 2;
      col = 16;
    }
    let confData = [];
    for (let i = 0; i < confTeams.length; i++) {
      const team = confTeams[i];
      const teamRecords = confStandings[team];
      let arr = [];
      arr.push(team);
      arr.push(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
      if (conf != "IND") {
        arr.push(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
      }
      confData.push(arr);
    }
    if (conf === "IND") {
      sheet.getRange(row, col, confSize, 2).setValues(confData);
    }
    else {
      sheet.getRange(row, col, confSize, 3).setValues(confData);
    }
  }
  else {
    let divARow = 3;
    if (conf === "MAC" || conf === "MWC" || conf === "SBC") {
      divARow = 21;
    }
    let divBRow = divARow + confSize / 2 + 1;
    let col = 1;
    if (conf === "B10") {
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
