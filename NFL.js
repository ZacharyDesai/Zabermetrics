function main() {
  
  // UPDATE EACH WEEK:
  const numWeeksDone = 0; // equivalent to the upcoming week #
  
  let ratings   = {"BUF":{}, "MIA":{}, "NE":{}, "NYJ":{}, "BAL":{}, "CIN":{}, "CLE":{}, "PIT":{}, "HOU":{}, "IND":{}, "JAX":{}, "TEN":{}, "DEN":{}, "KC":{}, "LAC":{}, "LAV":{},
                  "DAL":{}, "NYG":{}, "PHI":{}, "WSH":{}, "CHI":{}, "DET":{}, "GB":{}, "MIN":{}, "ATL":{}, "CAR":{}, "NO":{}, "TB":{}, "ARI":{}, "LAR":{}, "SEA":{}, "SF":{}};
  let schedule  = {"BUF":[], "MIA":[], "NE":[], "NYJ":[], "BAL":[], "CIN":[], "CLE":[], "PIT":[], "HOU":[], "IND":[], "JAX":[], "TEN":[], "DEN":[], "KC":[], "LAC":[], "LAV":[],
                  "DAL":[], "NYG":[], "PHI":[], "WSH":[], "CHI":[], "DET":[], "GB":[], "MIN":[], "ATL":[], "CAR":[], "NO":[], "TB":[], "ARI":[], "LAR":[], "SEA":[], "SF":[]};
  let standings = {"BUF":{}, "MIA":{}, "NE":{}, "NYJ":{}, "BAL":{}, "CIN":{}, "CLE":{}, "PIT":{}, "HOU":{}, "IND":{}, "JAX":{}, "TEN":{}, "DEN":{}, "KC":{}, "LAC":{}, "LAV":{},
                  "DAL":{}, "NYG":{}, "PHI":{}, "WSH":{}, "CHI":{}, "DET":{}, "GB":{}, "MIN":{}, "ATL":{}, "CAR":{}, "NO":{}, "TB":{}, "ARI":{}, "LAR":{}, "SEA":{}, "SF":{}};
  let standingsAFCEast  = {"BUF":{}, "MIA":{}, "NE":{}, "NYJ":{}};
  let standingsAFCNorth = {"BAL":{}, "CIN":{}, "CLE":{}, "PIT":{}};
  let standingsAFCSouth = {"HOU":{}, "IND":{}, "JAX":{}, "TEN":{}};
  let standingsAFCWest  = {"DEN":{}, "KC":{}, "LAC":{}, "LAV":{}};
  let standingsNFCEast  = {"DAL":{}, "NYG":{}, "PHI":{}, "WSH":{}};
  let standingsNFCNorth = {"CHI":{}, "DET":{}, "GB":{}, "MIN":{}};
  let standingsNFCSouth = {"ATL":{}, "CAR":{}, "NO":{}, "TB":{}};
  let standingsNFCWest  = {"ARI":{}, "LAR":{}, "SEA":{}, "SF":{}};

  clear(numWeeksDone);
  loadRatings(ratings);
  loadSchedule(schedule, numWeeksDone);
  simGames(ratings, schedule, numWeeksDone, "AFC");
  simGames(ratings, schedule, numWeeksDone, "NFC");
  calculateRecords("AFC", standings, standingsAFCEast, standingsAFCNorth, standingsAFCSouth, standingsAFCWest);
  calculateRecords("NFC", standings, standingsNFCEast, standingsNFCNorth, standingsNFCSouth, standingsNFCWest);
  calculateStrengthMetrics("AFC", standings, standingsAFCEast, standingsAFCNorth, standingsAFCSouth, standingsAFCWest);
  calculateStrengthMetrics("NFC", standings, standingsNFCEast, standingsNFCNorth, standingsNFCSouth, standingsNFCWest);
  fillDivisionStandings("AFC East", standingsAFCEast);
  fillDivisionStandings("AFC North", standingsAFCNorth);
  fillDivisionStandings("AFC South", standingsAFCSouth);
  fillDivisionStandings("AFC West", standingsAFCWest);
  fillDivisionStandings("NFC East", standingsNFCEast);
  fillDivisionStandings("NFC North", standingsNFCNorth);
  fillDivisionStandings("NFC South", standingsNFCSouth);
  fillDivisionStandings("NFC West", standingsNFCWest);
  addStrengthToDivisionStandings("AFC East", standings, standingsAFCEast, standingsAFCNorth, standingsAFCSouth, standingsAFCWest);
  addStrengthToDivisionStandings("AFC North", standings, standingsAFCEast, standingsAFCNorth, standingsAFCSouth, standingsAFCWest);
  addStrengthToDivisionStandings("AFC South", standings, standingsAFCEast, standingsAFCNorth, standingsAFCSouth, standingsAFCWest);
  addStrengthToDivisionStandings("AFC West", standings, standingsAFCEast, standingsAFCNorth, standingsAFCSouth, standingsAFCWest);
  addStrengthToDivisionStandings("NFC East", standings, standingsNFCEast, standingsNFCNorth, standingsNFCSouth, standingsNFCWest);
  addStrengthToDivisionStandings("NFC North", standings, standingsNFCEast, standingsNFCNorth, standingsNFCSouth, standingsNFCWest);
  addStrengthToDivisionStandings("NFC South", standings, standingsNFCEast, standingsNFCNorth, standingsNFCSouth, standingsNFCWest);
  addStrengthToDivisionStandings("NFC West", standings, standingsNFCEast, standingsNFCNorth, standingsNFCSouth, standingsNFCWest);
}

/* Clears any previous simulated material from the spreadsheet to reset it */
function clear(numWeeksDone) {
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Standings");
  for (let r = 1; r <= 20; r++) {
    if (r % 5 === 1) {
      continue;
    }
    for (let c = 1; c <= 12; c++) {
      sheetStandings.getRange(r, c).setValue("");
    }
  }
  for (let r = 4; r <= 20; r++) {
    if (r === 8) {
      continue;
    }
    for (let c = 13; c <= 16; c++) {
      sheetStandings.getRange(r, c).setValue("");
    }
  }
  if (numWeeksDone > 20) {
    return;
  }
  const sheetAFC = SpreadsheetApp.getActive().getSheetByName("AFC");
  for (let r = 4 + numWeeksDone; r <= 24; r++) {
    for (let c = 2; c <= 17; c++) {
      sheetAFC.getRange(r, c).setBackground("white");
    }
  }
  for (let c = 2; c <= 17; c++) {
    sheetAFC.getRange(3, c).setValue("--");
  }
  const sheetNFC = SpreadsheetApp.getActive().getSheetByName("NFC");
  for (let r = 4 + numWeeksDone; r <= 24; r++) {
    for (let c = 2; c <= 17; c++) {
      sheetNFC.getRange(r, c).setBackground("white");
    }
  }
  for (let c = 2; c <= 17; c++) {
    sheetNFC.getRange(3, c).setValue("--");
  }
}

/* Loads all ratings for each team */
function loadRatings(ratings) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Ratings");
  for (let r = 2; r <= 33; r++) {
    const team    = sheet.getRange(r, 1).getValue();
    const passOFF = sheet.getRange(r, 2).getValue();
    const rushOFF = sheet.getRange(r, 3).getValue();
    const passDEF = sheet.getRange(r, 4).getValue();
    const rushDEF = sheet.getRange(r, 5).getValue();
    const overall = sheet.getRange(r, 6).getValue();
    ratings[team] = {};
    ratings[team]["OVR"] = overall;
    ratings[team]["OFFpass"] = passOFF;
    ratings[team]["OFFrush"] = rushOFF;
    ratings[team]["DEFpass"] = passDEF;
    ratings[team]["DEFrush"] = rushDEF;
  }
}

/* Loads all remaining games into each team's schedule */
function loadSchedule(schedule, numWeeksDone) {
  if (numWeeksDone > 16) {
    return;
  }
  const sheetAFC = SpreadsheetApp.getActive().getSheetByName("AFC");
  for (let c = 2; c <= 17; c++) {
    const team = sheetAFC.getRange(2, c).getValue();
    for (let r = 4 + numWeeksDone; r <= 20; r++) {
      schedule[team][r - 4 - numWeeksDone] = sheetAFC.getRange(r, c).getValue();
    }
  }
  const sheetNFC = SpreadsheetApp.getActive().getSheetByName("NFC");
  for (let c = 2; c <= 17; c++) {
    const team = sheetNFC.getRange(2, c).getValue();
    for (let r = 4 + numWeeksDone; r <= 20; r++) {
      schedule[team][r - 4 - numWeeksDone] = sheetNFC.getRange(r, c).getValue();
    }
  }
}

/* Simulates every remaining game on every team's schedule */
function simGames(ratings, schedule, numWeeksDone, conf) {
  // The following array is to be used as a tiebreaker in simulating games; simulated games should only look at team ratings and home field advantage first
  const teamRatings = ["BAL", "NE", "KC", "NO", "SF", "DAL", "MIN", "SEA", "TEN", "GB", "PHI", "LAR", "BUF", "TB", "CHI", "IND",
                       "ATL", "PIT", "HOU", "ARI", "LAC", "DEN", "CLE", "LAV", "DET", "NYJ", "NYG", "JAX", "CIN", "WSH", "CAR", "MIA"];
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  for (let i = 2; i <= 17; i++) {
    const team = sheet.getRange(2, i).getValue();
    const teamSchedule = schedule[team];
    for (let j = 0; j < teamSchedule.length; j++) {
      let opponent = teamSchedule[j];
      if (opponent === "--") {
        continue;
      }
      let gameType = 0; // team is home
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
        gameType = 1; // team is away
      }
      else if (opponent.substring(0, 3) === "vs.") {
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
          let advantage = simGame(ratings, gameType, team, opponent, teamRatings);
          if (advantage < 0) {
            sheet.getRange(j + 4 + numWeeksDone, i).setBackground("red");
            continue;
          }
        }
      }
      sheet.getRange(j + 4 + numWeeksDone, i).setBackground("limegreen");
    }
  }
}

/* Simulates a certain given by the parameters */
function simGame(ratings, gameType, team, opponent, teamRatings) {
  // The following code is a temporary placeholder to be used before team ratings for the season are actually computed
  let advantage = teamRatings.indexOf(opponent) - teamRatings.indexOf(team);
  if (gameType === 0) {
    advantage += 10;
  }
  if (gameType === 1) {
    advantage -= 10;
  }
  return advantage;
//  let advantage = ratings[team]["Overall"] - ratings[opponent]["Overall"];
//  if (gameType === 0) {
//    advantage += 5;
//  }
//  if (gameType === 1) {
//    advantage -= 5;
//  }
//  if (advantage === 0) {
//    advantage = teamRatings.indexOf(opponent) - teamRatings.indexOf(team);
//  }
//  return advantage;
}

/* Calculates each team's record from its entire schedule */
function calculateRecords(conf, standings, standingsEast, standingsNorth, standingsSouth, standingsWest) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  for (let c = 2; c <= 17; c++) {
    const team = sheet.getRange(2, c).getValue();
    let numWinsOVR = 0;
    let numLossesOVR = 0;
    let numWinsDIV = 0;
    let numLossesDIV = 0;
    let numWinsCONF = 0;
    let numLossesCONF = 0;
    for (let r = 4; r <= 20; r++) {
      let opponent = sheet.getRange(r, c).getValue();
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "vs.") {
        opponent = opponent.substring(4, opponent.length);
      }
      const result = sheet.getRange(r, c).getBackgroundColor();
      if (result === "#ffffff") {
        continue;
      }
      if (result === "#ff0000") {
        numLossesOVR++;
        if (opponent in standingsEast || opponent in standingsNorth || opponent in standingsSouth || opponent in standingsWest) {
          numLossesCONF++;
          if ((team in standingsEast && opponent in standingsEast) || (team in standingsNorth && opponent in standingsNorth) ||
            (team in standingsSouth && opponent in standingsSouth) || (team in standingsWest && opponent in standingsWest)) {
              numLossesDIV++;
          }
        }
      }
      else {
        numWinsOVR++;
        if (opponent in standingsEast || opponent in standingsNorth || opponent in standingsSouth || opponent in standingsWest) {
          numWinsCONF++;
          if ((team in standingsEast && opponent in standingsEast) || (team in standingsNorth && opponent in standingsNorth) ||
              (team in standingsSouth && opponent in standingsSouth) || (team in standingsWest && opponent in standingsWest)) {
                numWinsDIV++;
          }
        }
      }
    }
    sheet.getRange(3, c).setValue(numWinsOVR + "-" + numLossesOVR);
    standings[team]["OVRwins"] = numWinsOVR;
    standings[team]["OVRlosses"] = numLossesOVR;
    standings[team]["DIVwins"] = numWinsDIV;
    standings[team]["DIVlosses"] = numLossesDIV;
    standings[team]["CONFwins"] = numWinsCONF;
    standings[team]["CONFlosses"] = numLossesCONF;
    if (team in standingsEast) {
      standingsEast[team]["OVRwins"] = numWinsOVR;
      standingsEast[team]["OVRlosses"] = numLossesOVR;
      standingsEast[team]["DIVwins"] = numWinsDIV;
      standingsEast[team]["DIVlosses"] = numLossesDIV;
      standingsEast[team]["CONFwins"] = numWinsCONF;
      standingsEast[team]["CONFlosses"] = numLossesCONF;
    }
    if (team in standingsNorth) {
      standingsNorth[team]["OVRwins"] = numWinsOVR;
      standingsNorth[team]["OVRlosses"] = numLossesOVR;
      standingsNorth[team]["DIVwins"] = numWinsDIV;
      standingsNorth[team]["DIVlosses"] = numLossesDIV;
      standingsNorth[team]["CONFwins"] = numWinsCONF;
      standingsNorth[team]["CONFlosses"] = numLossesCONF;
    }
    if (team in standingsSouth) {
      standingsSouth[team]["OVRwins"] = numWinsOVR;
      standingsSouth[team]["OVRlosses"] = numLossesOVR;
      standingsSouth[team]["DIVwins"] = numWinsDIV;
      standingsSouth[team]["DIVlosses"] = numLossesDIV;
      standingsSouth[team]["CONFwins"] = numWinsCONF;
      standingsSouth[team]["CONFlosses"] = numLossesCONF;
    }
    if (team in standingsWest) {
      standingsWest[team]["OVRwins"] = numWinsOVR;
      standingsWest[team]["OVRlosses"] = numLossesOVR;
      standingsWest[team]["DIVwins"] = numWinsDIV;
      standingsWest[team]["DIVlosses"] = numLossesDIV;
      standingsWest[team]["CONFwins"] = numWinsCONF;
      standingsWest[team]["CONFlosses"] = numLossesCONF;
    }
  }
}

/* Calculates each team's strength of victory and strength of schedule for tiebreaking purposes */
function calculateStrengthMetrics(conf, standings, standingsEast, standingsNorth, standingsSouth, standingsWest) {
  const sheetSchedule = SpreadsheetApp.getActive().getSheetByName(conf);
  for (let c = 2; c <= 17; c++) {
    const team = sheetSchedule.getRange(2, c).getValue();
    let sov = 0;
    let sos = 0;
    for (let r = 4; r <= 20; r++) {
      let opponent = sheetSchedule.getRange(r, c).getValue();
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "vs.") {
        opponent = opponent.substring(4, opponent.length);
      }
      const result = sheetSchedule.getRange(r, c).getBackgroundColor();
      if (result === "#ffffff") {
        continue;
      }
      if (result != "#ff0000") {
        sov += standings[opponent]["OVRwins"];
      }
      sos += standings[opponent]["OVRwins"];
    }
    if (standings[team]["OVRwins"] === 0) {
      sov = 0;
    }
    else {
      sov /= 16 * standings[team]["OVRwins"];
    }
    sos /= 256;
    standings[team]["sov"] = sov;
    standings[team]["sos"] = sos;
    if (team in standingsEast) {
      standingsEast[team]["sov"] = sov;
      standingsEast[team]["sos"] = sos;
    }
    if (team in standingsNorth) {
      standingsNorth[team]["sov"] = sov;
      standingsNorth[team]["sos"] = sos;
    }
    if (team in standingsSouth) {
      standingsSouth[team]["sov"] = sov;
      standingsSouth[team]["sos"] = sos;
    }
    if (team in standingsWest) {
      standingsWest[team]["sov"] = sov;
      standingsWest[team]["sos"] = sos;
    }
  }
}

/* Generates each division's standings based on each of its team's records */
function fillDivisionStandings(division, standings) {
  const conf = division.substring(0, 3);
  const div = division.substring(4, division.length);
  const sheetSchedule = SpreadsheetApp.getActive().getSheetByName(conf);
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Standings");
  let teams = Object.keys(standings);
  teams.sort(sortByRecord(standings));
  let row;
  switch (div) {
    case "East":
      row = 2;
      break;
    case "North":
      row = 7;
      break;
    case "South":
      row = 12;
      break;
    case "West":
      row = 17;
      break;
  }
  let col;
  switch (conf) {
    case "AFC":
      col = 0;
      break;
    case "NFC":
      col = 6;
      break;
  }
  for (let i = 0; i < teams.length; i++) {
    const team = teams[i];
    const teamRecords = standings[team];
    sheetStandings.getRange(row + i, col + 1).setValue(team);
    sheetStandings.getRange(row + i, col + 2).setValue(teamRecords["OVRwins"] + "-" + teamRecords["OVRlosses"]);
    sheetStandings.getRange(row + i, col + 3).setValue(teamRecords["DIVwins"] + "-" + teamRecords["DIVlosses"]);
    sheetStandings.getRange(row + i, col + 4).setValue(teamRecords["CONFwins"] + "-" + teamRecords["CONFlosses"]);
  }
}

/* Adds the computed strength of victory and strength of schedule metrics to the standings */
function addStrengthToDivisionStandings(division, standings) {
  const conf = division.substring(0, 3);
  const div = division.substring(4, division.length);
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Standings");
  let row;
  switch (div) {
    case "East":
      row = 2;
      break;
    case "North":
      row = 7;
      break;
    case "South":
      row = 12;
      break;
    case "West":
      row = 17;
      break;
  }
  let col;
  switch (conf) {
    case "AFC":
      col = 0;
      break;
    case "NFC":
      col = 6;
      break;
  }
  for (let i = 0; i < 4; i++) {
    const team = sheetStandings.getRange(row + i, col + 1).getValue();
    const teamRecords = standings[team];
    sheetStandings.getRange(row + i, col + 5).setValue(teamRecords["sov"]);
    sheetStandings.getRange(row + i, col + 6).setValue(teamRecords["sos"]);
  }
}

/* Sorts the given divison's standings based on overall team record (if tied, DIV record, CONF record, strength of victory, and then strength of schedule) */
function sortByRecord(standings) {
  return function(a, b) {
    if (standings[b]["OVRwins"] != standings[a]["OVRwins"]) {
      return standings[b]["OVRwins"] - standings[a]["OVRwins"];
    }
    if (standings[b]["DIVwins"] != standings[a]["DIVwins"]) {
      return standings[b]["DIVwins"] - standings[a]["DIVwins"];
    }
    if (standings[b]["CONFwins"] != standings[a]["CONFwins"]) {
      return standings[b]["CONFwins"] - standings[a]["CONFwins"];
    }
    if (standings[b]["sov"] != standings[a]["sov"]) {
      return standings[b]["sov"] - standings[a]["sov"];
    }
    if (standings[b]["sos"] != standings[a]["sos"]) {
      return standings[b]["sos"] - standings[a]["sos"];
    }
    return 0;
  };
}
