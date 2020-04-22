function main() {
  
  // UPDATE EACH WEEK:
  const numWeeksDone = 0; // equivalent to the upcoming week #
  
  // UPDATE EACH SEASON:
  const confSizes = {AAC:11, ACC:14, B10:14, B12:10, CUSA:14, MAC:12, MWC:12, P12:12, SBC:10, SEC:14, IND:7};
  const numWeeks = 18;
  const rowGroupOf5 = 19;
  const numTeams = 130;
  
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
  
  loadRatings(ratings, numWeeks);
  
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

/* Clears any previous simulated material from the spreadsheet to reset it */
function clear(numWeeks, numWeeksDone, rowGroupOf5, conf, confSize) {
  const sheetRankings = SpreadsheetApp.getActive().getSheetByName("Rankings");
  for (let i = numWeeksDone * 3; i < (numWeeks - 1) * 3; i += 3) {
    for (let j = 3; j < 28; j++) {
      sheetRankings.getRange(j, i + 2).setValue("");
      sheetRankings.getRange(j, i + 2).setBackground("white");
      sheetRankings.getRange(j, i + 3).setValue("");
      sheetRankings.getRange(j, i + 3).setBackground("white");
    }
  }
  const sheetSchedule = SpreadsheetApp.getActive().getSheetByName(conf);
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Standings");
  if (numWeeksDone < 18) {
    const start = numWeeksDone + 4;
    for (let i = 2; i < confSize + 2; i++) {
      for (let j = start; j < start + numWeeks; j++) {
        sheetSchedule.getRange(j, i).setBackground("white");
      }
      sheetSchedule.getRange(3, i).setValue("--");
      sheetSchedule.getRange(20, i).setValue("--");
      sheetSchedule.getRange(21, i).setValue("--");
    }
  }
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

/* Loads all remaining games into each team's schedule, and loads each team into each conference's standings */
function loadGames(schedule, standings, numWeeks, numWeeksDone, conf, confSize) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  let teams = {};
  const start = numWeeksDone + 4;
  for (let i = 2; i < confSize + 2; i++) {
    const team = sheet.getRange(2, i).getValue();
    let opponents = [];
    for (let j = 0; j < numWeeks; j++) {
      opponents[j] = sheet.getRange(j + start, i).getValue();
    }
    schedule[team] = opponents;
    teams[team] = {};
  }
  standings[conf] = teams;
}

/* Loads all ratings for each team */
function loadRatings(ratings, numTeams) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Ratings");
  for (let i = 2; i < numTeams + 2; i++) {
    const team = sheet.getRange(i, 1).getValue();
    const passOFF = sheet.getRange(i, 2).getValue();
    const rushOFF = sheet.getRange(i, 3).getValue();
    const passDEF = sheet.getRange(i, 4).getValue();
    const rushDEF = sheet.getRange(i, 5).getValue();
    const overall = sheet.getRange(i, 6).getValue();
    ratings[team] = {};
    ratings[team]["Pass OFF"] = passOFF;
    ratings[team]["Rush OFF"] = rushOFF;
    ratings[team]["Pass DEF"] = passDEF;
    ratings[team]["Rush DEF"] = rushDEF;
    ratings[team]["Overall"] = overall;
  }
}

/* Simulates every remaining game on every team's schedule */
function simGames(schedule, standings, ratings, numWeeks, numWeeksDone, conf, confSize) {
  // The following array is to be used as a tiebreaker in simulating games; simulated games should only look at team ratings and home field advantage first
  const teamRatings = ["CLEMSON", "OHIO STATE", "OKLAHOMA", "ALABAMA", "PENN STATE", "WISCONSIN", "TEXAS", "TEXAS A&M", "NOTRE DAME", "GEORGIA",
                       "FLORIDA", "LSU", "USC", "OREGON", "AUBURN", "MICHIGAN", "OKLAHOMA STATE", "NORTH CAROLINA", "TENNESSEE", "MINNESOTA",
                       "UCF", "NEBRASKA", "FLORIDA STATE", "UTAH", "VIRGINIA TECH", "INDIANA", "IOWA", "STANFORD", "WASHINGTON", "CALIFORNIA",
                       "IOWA STATE", "TCU", "KENTUCKY", "SOUTH CAROLINA", "LOUISVILLE", "PURDUE", "MIAMI (FL)", "NORTHWESTERN", "MISSISSIPPI STATE", "CINCINNATI",
                       "ARIZONA STATE", "OLE MISS", "PITTSBURGH", "BAYLOR", "HOUSTON", "TEXAS TECH", "VIRGINIA", "WEST VIRGINIA", "UCLA", "KANSAS STATE",
                       "BOISE STATE", "NAVY", "MISSOURI", "WASHINGTON STATE", "GEORGIA TECH", "COLORADO", "MEMPHIS", "MICHIGAN STATE", "NC STATE", "SMU",
                       "LOUISIANA-LAFAYETTE", "BYU", "ARIZONA", "DUKE", "ILLINOIS", "ARKANSAS", "OREGON STATE", "MARYLAND", "WAKE FOREST", "TULSA",
                       "BUFFALO", "SYRACUSE", "WESTERN KENTUCKY", "TULANE", "APPALACHIAN STATE", "AIR FORCE", "WYOMING", "BOSTON COLLEGE", "SOUTHERN MISS", "MARSHALL",
                       "VANDERBILT", "WESTERN MICHIGAN", "RUTGERS", "MIAMI (OH)", "UAB", "USF", "ARKANSAS STATE", "SAN DIEGO STATE", "FAU", "OHIO",
                       "TEMPLE", "MIDDLE TENNESSEE", "FRESNO STATE", "BALL STATE", "UTAH STATE", "COLORADO STATE", "EAST CAROLINA", "NEVADA", "GEORGIA SOUTHERN", "ARMY",
                       "TOLEDO", "NORTHERN ILLINOIS", "CENTRAL MICHIGAN", "KANSAS", "LOUISIANA TECH", "KENT STATE", "GEORGIA STATE", "SAN JOSE STATE", "RICE", "TROY",
                       "CHARLOTTE", "COASTAL CAROLINA", "LOUISIANA-MONROE", "SOUTH ALABAMA", "NORTH TEXAS", "UTSA", "UNLV", "OLD DOMINION", "HAWAI'I", "FIU",
                       "EASTERN MICHIGAN", "LIBERTY", "CONNECTICUT", "NEW MEXICO", "TEXAS STATE", "NEW MEXICO STATE", "BOWLING GREEN", "AKRON", "UTEP", "MASSACHUSETTS"];
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const start = numWeeksDone + 4;
  for (let i = 2; i < confSize + 2; i++) {
    const team = sheet.getRange(2, i).getValue();
    const teamSchedule = schedule[team];
    standings[conf][team] = {};
    for (let j = 0; j < teamSchedule.length; j++) {
      let opponent = teamSchedule[j].toUpperCase();
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
      if (opponent in schedule) {
        let opponentOpponent = schedule[opponent][j].toUpperCase();
        if (gameType === 0) {
          opponentOpponent = opponentOpponent.substring(2, opponentOpponent.length);
        }
        else if (gameType === 2) {
          opponentOpponent = opponentOpponent.substring(4, opponentOpponent.length);
        }
        if (team === opponentOpponent) {
          let advantage = simGame(ratings, teamRatings, gameType, team, opponent);
          if (advantage < 0) {
            sheet.getRange(j + start, i).setBackground("red");
            continue;
          }
        }
      }
      sheet.getRange(j + start, i).setBackground("limegreen");
    }
  }
}

/* Simulates a certain given by the parameters */
function simGame(ratings, teamRatings, gameType, team, opponent) {
  // The following code is a temporary placeholder to be used before team ratings for the season are actually computed
  let advantage = teamRatings.indexOf(opponent) - teamRatings.indexOf(team);
  if (gameType === 0) {
    advantage += 12.5;
  }
  if (gameType === 1) {
    advantage -= 12.5;
  }
  return advantage;
//  let advantage = ratings[team]["Overall"] - ratings[opponent]["Overall"];
//  if (gameType === 0) {
//    advantage += 10;
//  }
//  if (gameType === 1) {
//    advantage -= 10;
//  }
//  if (advantage === 0) {
//    advantage = teamRatings.indexOf(opponent) - teamRatings.indexOf(team);
//  }
//  return advantage;
}

/* Calculates each team's record from its entire schedule */
function calculateRecords(standings, numWeeks, conf, confSize) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  for (let i = 2; i < confSize + 2; i++) {
    const team = sheet.getRange(2, i).getValue();
    let numWinsOVR = 0;
    let numLossesOVR = 0;
    let numWinsCONF = 0;
    let numLossesCONF = 0;
    for (let j = 4; j < 4 + numWeeks; j++) {
      let opponent = sheet.getRange(j, i).getValue().toUpperCase();
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "VS.") {
        opponent = opponent.substring(4, opponent.length);
      }
      const result = sheet.getRange(j, i).getBackgroundColor();
      if (result === "#ffffff") {
        continue;
      }
      if (result === "#ff0000") {
        numLossesOVR++;
        if (opponent in standings[conf]) {
          numLossesCONF++;
        }
      }
      else {
        numWinsOVR++;
        if (opponent in standings[conf]) {
          numWinsCONF++;
        }
      }
    }
    if (conf === "IND") {
      numWinsCONF = 0;
      numLossesCONF = 0;
      sheet.getRange(3, i).setValue(numWinsOVR + "-" + numLossesOVR);
    }
    else {
      sheet.getRange(3, i).setValue(numWinsOVR + "-" + numLossesOVR + " (" + numWinsCONF + "-" + numLossesCONF + ")");
    }
    standings[conf][team]["OVR Wins"] = numWinsOVR;
    standings[conf][team]["OVR Losses"] = numLossesOVR;
    standings[conf][team]["CONF Wins"] = numWinsCONF;
    standings[conf][team]["CONF Losses"] = numLossesCONF;
    if (conf === "AAC" || conf === "B12" || conf === "IND") {
      standings[conf][team]["Division"] = -1;
    }
    else {
      if (i - 2 < confSize / 2) {
        standings[conf][team]["Division"] = 0;
      }
      else {
        standings[conf][team]["Division"] = 1;
      }
      index = i - 1;
    }
  }
}

/* Generates each conference's standings based on each of its team's records */
function generateStandings(standings, numWeeks, conf, confSize) {
  const sheetSchedule = SpreadsheetApp.getActive().getSheetByName(conf);
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Standings");
  let confStandings = standings[conf];
  let confTeams = Object.keys(confStandings);
  confTeams.sort(sortByRecord(confStandings));
  if (conf === "AAC" || conf === "B12" || conf === "IND") {
    let row = 2;
    let col = 0;
    if (conf === "AAC") {
      row = 20;
    }
    if (conf === "B12") {
      col = 6;
    }
    if (conf === "IND") {
      col = 15;
    }
    for (let i = 0; i < confTeams.length; i++) {
      const team = confTeams[i];
      const teamRecords = confStandings[team];
      sheetStandings.getRange(row + i, col + 1).setValue(team);
      sheetStandings.getRange(row + i, col + 2).setValue(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
      sheetStandings.getRange(row + i, col + 3).setValue(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
    }
  }
  else {
    let divARow = 2;
    if (conf === "CUSA" || conf === "MAC" || conf === "MWC" || conf === "SBC") {
      divARow = 20;
    }
    let divBRow = divARow + confSize / 2 + 1;
    let col = 0;
    if (conf === "B10" || conf === "CUSA") {
      col = 3;
    }
    if (conf === "MAC") {
      col = 6;
    }
    if (conf === "P12" || conf === "MWC") {
      col = 9;
    }
    if (conf === "SEC" || conf === "SBC") {
      col = 12;
    }
    for (let i = 0; i < confTeams.length; i++) {
      const team = confTeams[i];
      const teamRecords = confStandings[team];
      if (teamRecords["Division"] === 0) {
        divARow++;
        sheetStandings.getRange(divARow, col + 1).setValue(team);
        sheetStandings.getRange(divARow, col + 2).setValue(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
        sheetStandings.getRange(divARow, col + 3).setValue(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
      }
      else {
        divBRow++;
        sheetStandings.getRange(divBRow, col + 1).setValue(team);
        sheetStandings.getRange(divBRow, col + 2).setValue(teamRecords["OVR Wins"] + "-" + teamRecords["OVR Losses"]);
        sheetStandings.getRange(divBRow, col + 3).setValue(teamRecords["CONF Wins"] + "-" + teamRecords["CONF Losses"]);
      }
    }
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
