// UPDATE EACH WEEK:
const numWeeksDone = 0; // equivalent to the previous week #

// UPDATE AS NEEDED:
const homefieldAdvantage = 4.44;

/* MAIN FUNCTION: Runs the simulation */
function main() {
  
  let ratings = {};
  let schedule  = {"BUF":[], "MIA":[], "NE":[], "NYJ":[], "BAL":[], "CIN":[], "CLE":[], "PIT":[], "HOU":[], "IND":[], "JAX":[], "TEN":[], "DEN":[], "KC":[], "LAC":[], "LV":[],
                  "DAL":[], "NYG":[], "PHI":[], "WAS":[], "CHI":[], "DET":[], "GB":[], "MIN":[], "ATL":[], "CAR":[], "NO":[], "TB":[], "ARI":[], "LAR":[], "SEA":[], "SF":[]};
  let standings = {"BUF":{}, "MIA":{}, "NE":{}, "NYJ":{}, "BAL":{}, "CIN":{}, "CLE":{}, "PIT":{}, "HOU":{}, "IND":{}, "JAX":{}, "TEN":{}, "DEN":{}, "KC":{}, "LAC":{}, "LV":{},
                  "DAL":{}, "NYG":{}, "PHI":{}, "WAS":{}, "CHI":{}, "DET":{}, "GB":{}, "MIN":{}, "ATL":{}, "CAR":{}, "NO":{}, "TB":{}, "ARI":{}, "LAR":{}, "SEA":{}, "SF":{}};
  let standingsAFCEast  = {"BUF":{}, "MIA":{}, "NE":{}, "NYJ":{}};
  let standingsAFCNorth = {"BAL":{}, "CIN":{}, "CLE":{}, "PIT":{}};
  let standingsAFCSouth = {"HOU":{}, "IND":{}, "JAX":{}, "TEN":{}};
  let standingsAFCWest  = {"DEN":{}, "KC":{}, "LAC":{}, "LV":{}};
  let standingsNFCEast  = {"DAL":{}, "NYG":{}, "PHI":{}, "WAS":{}};
  let standingsNFCNorth = {"CHI":{}, "DET":{}, "GB":{}, "MIN":{}};
  let standingsNFCSouth = {"ATL":{}, "CAR":{}, "NO":{}, "TB":{}};
  let standingsNFCWest  = {"ARI":{}, "LAR":{}, "SEA":{}, "SF":{}};

  clear();
  loadRatings(ratings);
  loadSchedule(schedule);
  simGames(ratings, schedule, "AFC");
  simGames(ratings, schedule, "NFC");
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
}

/* UTILITY FUNCTION: Use to check if the schedule was inputted correctly */
function checkSchedule() {
  let schedule  = {"BUF":[], "MIA":[], "NE":[], "NYJ":[], "BAL":[], "CIN":[], "CLE":[], "PIT":[], "HOU":[], "IND":[], "JAX":[], "TEN":[], "DEN":[], "KC":[], "LAC":[], "LV":[],
                  "DAL":[], "NYG":[], "PHI":[], "WAS":[], "CHI":[], "DET":[], "GB":[], "MIN":[], "ATL":[], "CAR":[], "NO":[], "TB":[], "ARI":[], "LAR":[], "SEA":[], "SF":[]};
  clear();
  loadSchedule(schedule);
  checkScheduleHelper(schedule, "AFC");
  checkScheduleHelper(schedule, "NFC");
}

// Checks each match in schedule to make sure each opponent is an actual team
function checkScheduleHelper(schedule, conf) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, 16).getValues();
  for (let c = 0; c < 16; c++) {
    const team = teams[0][c];
    const teamSchedule = schedule[team];
    for (let r = 0; r < teamSchedule.length; r++) {
      let opponent = teamSchedule[r];
      if (opponent === "--") {
        sheet.getRange(r + 4, c + 2).setBackground("limegreen");
        continue;
      }
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "vs.") {
        opponent = opponent.substring(4, opponent.length);
      }
      
      if (!(opponent in schedule)) {
        sheet.getRange(r + 4, c + 2).setBackground("red");
      }
    }
  }
}

/* Clears any previous simulated material from the spreadsheet to reset it */
function clear() {
  
  // Clear Standings sheet
  const sheetStandings = SpreadsheetApp.getActive().getSheetByName("Simulated Standings");
  let clearedStandings = [];
  for (let r = 0; r < 4; r++) {
    let arr = [];
    for (let c = 0; c < 12; c++) {
      arr.push("");
    }
    clearedStandings.push(arr);
  }
  for (let r = 2; r < 21; r += 5) {
    sheetStandings.getRange(r, 1, 4, 12).setValues(clearedStandings);
  }
  
  let clearedRecords = [];
  clearedRecords.push([]);
  for (let c = 0; c < 16; c++) {
    clearedRecords[0].push("--");
  }
  
  if (numWeeksDone > 17) {
    return;
  }
  
  // Clear AFC sheet
  const sheetAFC = SpreadsheetApp.getActive().getSheetByName("AFC");
  sheetAFC.getRange(4 + numWeeksDone, 2, 18 - numWeeksDone, 16).setBackground("white");
  sheetAFC.getRange(3, 2, 1, 16).setValues(clearedRecords);
  
  // Clear NFC sheet
  const sheetNFC = SpreadsheetApp.getActive().getSheetByName("NFC");
  sheetNFC.getRange(4 + numWeeksDone, 2, 18 - numWeeksDone, 16).setBackground("white");
  sheetNFC.getRange(3, 2, 1, 16).setValues(clearedRecords);
}

/* Loads all ratings for each team */
function loadRatings(ratings) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Ratings");
  const ratingsData = sheet.getRange(2, 1, 32, 2).getValues();
  for (let i = 0; i < 32; i++) {
    const team = ratingsData[i][0];
    const rating = ratingsData[i][1];
    ratings[team] = rating;
  }
}

/* Loads all remaining games into each team's schedule */
function loadSchedule(schedule) {
  if (numWeeksDone > 17) {
    return;
  }
  
  // Load from AFC sheet
  const sheetAFC = SpreadsheetApp.getActive().getSheetByName("AFC");
  const teamsAFC = sheetAFC.getRange(2, 2, 1, 16).getValues();
  const scheduleAFC = sheetAFC.getRange(4 + numWeeksDone, 2, 18 - numWeeksDone, 16).getValues();
  for (let c = 0; c < 16; c++) {
    const team = teamsAFC[0][c];
    schedule[team] = [];
    for (let r = 0; r < scheduleAFC.length; r++) {
      schedule[team].push(scheduleAFC[r][c]);
    }
  }
  
  // Load from NFC sheet
  const sheetNFC = SpreadsheetApp.getActive().getSheetByName("NFC");
  const teamsNFC = sheetNFC.getRange(2, 2, 1, 16).getValues();
  const scheduleNFC = sheetNFC.getRange(4 + numWeeksDone, 2, 18 - numWeeksDone, 16).getValues();
  for (let c = 0; c < 16; c++) {
    const team = teamsNFC[0][c];
    schedule[team] = [];
    for (let r = 0; r < scheduleNFC.length; r++) {
      schedule[team].push(scheduleNFC[r][c]);
    }
  }
}

/* Simulates every remaining game on every team's schedule */
function simGames(ratings, schedule, conf) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, 16).getValues();
  for (let c = 0; c < 16; c++) {
    const team = teams[0][c];
    const teamSchedule = schedule[team];
    for (let r = 0; r < teamSchedule.length; r++) {
      let opponent = teamSchedule[r];
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
      let advantage = simGame(ratings, gameType, team, opponent);
      if (advantage < 0) {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("red");
      }
      else {
        sheet.getRange(r + 4 + numWeeksDone, c + 2).setBackground("limegreen");
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
  return advantage;
}

/* Calculates each team's record from its entire schedule */
function calculateRecords(conf, standings, standingsEast, standingsNorth, standingsSouth, standingsWest) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(conf);
  const teams = sheet.getRange(2, 2, 1, 16).getValues();
  const results = sheet.getRange(4, 2, 18, 16).getValues();
  let records = [];
  for (let c = 0; c < 16; c++) {
    const team = teams[0][c];
    let numWinsOVR = 0;
    let numLossesOVR = 0;
    let numWinsDIV = 0;
    let numLossesDIV = 0;
    let numWinsCONF = 0;
    let numLossesCONF = 0;
    for (let r = 0; r < 18; r++) {
      const result = sheet.getRange(r + 4, c + 2).getBackgroundColor();
      
      // Match was actually a bye week
      if (result === "#ffffff") {
        continue;
      }
      
      // Parse opposing team name
      let opponent = results[r][c];
      if (opponent.charAt(0) === '@') {
        opponent = opponent.substring(2, opponent.length);
      }
      else if (opponent.substring(0, 3) === "vs.") {
        opponent = opponent.substring(4, opponent.length);
      }
      
      // Team won the match
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
      
      // Team lost the match
      else if (result === "#32cd32") {
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
    records.push(numWinsOVR + "-" + numLossesOVR);
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
  records = [records];
  sheet.getRange(3, 2, 1, 16).setValues(records);
}

/* Calculates each team's strength of victory and strength of schedule for tiebreaking purposes */
function calculateStrengthMetrics(conf, standings, standingsEast, standingsNorth, standingsSouth, standingsWest) {
  const sheetSchedule = SpreadsheetApp.getActive().getSheetByName(conf);
  for (let c = 2; c <= 17; c++) {
    const team = sheetSchedule.getRange(2, c).getValue();
    let sovWins = 0;
    let sovLosses = 0;
    let sosWins = 0;
    let sosLosses = 0;
    for (let r = 4; r <= 21; r++) {
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
        sovWins += standings[opponent]["OVRwins"];
        sovLosses += standings[opponent]["OVRlosses"];
      }
      sosWins += standings[opponent]["OVRwins"];
      sosLosses += standings[opponent]["OVRlosses"];
    }
    let numWins = standings[team]["OVRwins"];
    let sovTies = numWins * 17 - sovWins - sovLosses;
    let sosTies = 289 - sosWins - sosLosses;
    let sov = (sovWins + sovTies * 0.5) / (numWins * 17);
    if (numWins == 0) {
      sov = 0.0;
    }
    let sos = (sosWins + sosTies * 0.5) / 289;
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
  const sheet = SpreadsheetApp.getActive().getSheetByName("Simulated Standings");
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
      col = 1;
      break;
    case "NFC":
      col = 7;
      break;
  }
  let standingsData = [];
  for (let i = 0; i < teams.length; i++) {
    const team = teams[i];
    const teamRecords = standings[team];
    let arr = [];
    arr.push(team);
    arr.push(teamRecords["OVRwins"] + "-" + teamRecords["OVRlosses"]);
    arr.push(teamRecords["DIVwins"] + "-" + teamRecords["DIVlosses"]);
    arr.push(teamRecords["CONFwins"] + "-" + teamRecords["CONFlosses"]);
    arr.push(teamRecords["sov"]);
    arr.push(teamRecords["sos"]);
    standingsData.push(arr);
  }
  sheet.getRange(row, col, 4, 6).setValues(standingsData);
}

/* Sorts the given divison's standings based on overall team record (if tied, DIV record, CONF record, strength of victory, and then strength of schedule) */
function sortByRecord(standings) {
  return function(a, b) {
    if (standings[b]["OVRwins"] - standings[b]["OVRlosses"] != standings[a]["OVRwins"] - standings[a]["OVRlosses"]) {
      return (standings[b]["OVRwins"] - standings[b]["OVRlosses"]) - (standings[a]["OVRwins"] - standings[a]["OVRlosses"]);
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
