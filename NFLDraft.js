// UPDATE EACH ROUND:
const curRoundNum = 1; // the first round to be simulated; no need to simulate completed rounds

// UPDATE EACH YEAR:
const numPicks = [32, 32, 42, 40, 33, 35, 41];

/* Simulates round 1 of the draft */
function round1() {
  simRound(1);
}

/* Simulates round 2 of the draft */
function round2() {
  simRound(2);
}

/* Simulates round 3 of the draft */
function round3() {
  simRound(3);
}

/* Simulates round 4 of the draft */
function round4() {
  simRound(4);
}

/* Simulates round 5 of the draft */
function round5() {
  simRound(5);
}

/* Simulates round 6 of the draft */
function round6() {
  simRound(6);
}

/* Simulates round 7 of the draft */
function round7() {
  simRound(7);
}

/* Simulates the given round */
function simRound(roundNum) {
  let prospects = [];
  let teamNeeds = {};
  let teamProspects = {};
  loadProspects(prospects, roundNum, numPicks[roundNum - 1] * 2);
  loadTeamNeeds(teamNeeds, roundNum);
  loadTeamProspects(prospects, teamNeeds, teamProspects);
  fillDraft(teamNeeds, teamProspects, roundNum, numPicks[roundNum - 1]);
}

/* Fills the prospects array with best prospects available */
function loadProspects(prospects, roundNum, numToPick) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Prospects");
  const range = sheet.getRange(1, 1, 750, 4);
  let numPicked = 0;
  if (roundNum === curRoundNum) {
    for (let i = 0; i < 750 && numPicked < numToPick; i++) {
      prospects[i] = {pos: range.getCell(i + 1, 1).getValue(), name: range.getCell(i + 1, 2).getValue()};
      if (prospects[i].pos != "") {
        numPicked++;
      }
    }
  }
  else {
    for (let i = 0; i < 750 && numPicked < numToPick; i++) {
      prospects[i] = {pos: range.getCell(i + 1, 3).getValue(), name: range.getCell(i + 1, 4).getValue()};
      if (prospects[i].pos != "") {
        numPicked++;
      }
    }
  }
}

/* Fills each team's needs object with values corresponding to the "need" of each position */
function loadTeamNeeds(teamNeeds, roundNum) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Needs");
  const range = sheet.getRange(2, 1, 32, 26);
  if (roundNum === curRoundNum) {
    for (let i = 1; i <= 32; i++) {
      const team = range.getCell(i, 1).getValue();
      const qb = range.getCell(i, 2).getValue();
      const rb = range.getCell(i, 3).getValue();
      const wr = range.getCell(i, 4).getValue();
      const te = range.getCell(i, 5).getValue();
      const ol = range.getCell(i, 6).getValue();
      const edge = range.getCell(i, 7).getValue();
      const dl = range.getCell(i, 8).getValue();
      const lb = range.getCell(i, 9).getValue();
      const cb = range.getCell(i, 10).getValue();
      const s = range.getCell(i, 11).getValue();
      const k = range.getCell(i, 12).getValue();
      const p = range.getCell(i, 13).getValue();
      teamNeeds[team] = {"QB": qb, "RB": rb, "WR": wr, "TE": te, "OL": ol, "EDGE": edge, "DL": dl, "LB": lb, "CB": cb, "S": s, "K": k, "P": p};
    }
  }
  else {
    for (let i = 1; i <= 32; i++) {
      const team = range.getCell(i, 14).getValue();
      const qb = range.getCell(i, 15).getValue();
      const rb = range.getCell(i, 16).getValue();
      const wr = range.getCell(i, 17).getValue();
      const te = range.getCell(i, 18).getValue();
      const ol = range.getCell(i, 19).getValue();
      const edge = range.getCell(i, 20).getValue();
      const dl = range.getCell(i, 21).getValue();
      const lb = range.getCell(i, 22).getValue();
      const cb = range.getCell(i, 23).getValue();
      const s = range.getCell(i, 24).getValue();
      const k = range.getCell(i, 25).getValue();
      const p = range.getCell(i, 26).getValue();
      teamNeeds[team] = {"QB": qb, "RB": rb, "WR": wr, "TE": te, "OL": ol, "EDGE": edge, "DL": dl, "LB": lb, "CB": cb, "S": s, "K": k, "P": p};
    }
  }
}

/* Computes each team's prospects object by evaluating each player according to their rank in the prospects array and the team's need for their position */
function loadTeamProspects(prospects, teamNeeds, teamProspects) {
  for (const team in teamNeeds) {
    teamProspects[team] = [];
  }
  for (let i = 0; i < prospects.length; i++) {
    const prospect = prospects[i];
    const pos = prospect.pos;
    const name = prospect.name;
    if (pos != "") {
      for (const team in teamNeeds) {
        const val = i + teamNeeds[team][pos];
        teamProspects[team].push({pos: pos, name: name, val: val});
      }
    }
  }
  for (const team in teamProspects) {
    let theseProspects = teamProspects[team];
    theseProspects.sort(sortTeamProspects());
  }
}

/* Simulates the remaining picks of the draft by having each team select their most desired available prospect */
function fillDraft(teamNeeds, teamProspects, roundNum, numPicks) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("R" + roundNum);
  const sheetProspects = SpreadsheetApp.getActive().getSheetByName("Prospects");
  const rangeProspects = sheetProspects.getRange(1, 3, 750, 2);
  const sheetNeeds = SpreadsheetApp.getActive().getSheetByName("Needs");
  const rangeNeeds = sheetNeeds.getRange(2, 14, 32, 13);
  for (let i = 1; i <= numPicks * 2; i += 2) {
    const team = sheet.getRange(2, i).getValue();
    let prospects = teamProspects[team];
    let pos;
    let name;
    // If this pick hasn't occurred yet, simulate it and display each team's most desired prospects
    if (sheet.getRange(4, i).getValue() === "") {
      for (let j = 0; j < prospects.length; j++) {
        const prospect = prospects[j];
        if (j === 0) {
          sheet.getRange(6, i).setValue(prospect.pos);
          sheet.getRange(6, i + 1).setValue(prospect.name);
          pos = prospect.pos;
          name = prospect.name;
        }
        sheet.getRange(j + 8, i).setValue(prospect.pos);
        sheet.getRange(j + 8, i + 1).setValue(prospect.name);
      }
    }
    // If this pick has occurred, use the actual pick that was made
    else {
      pos = sheet.getRange(4, i).getValue();
      name = sheet.getRange(4, i + 1).getValue();
    }
    // Update the "needs" for current team (decrease the need for the drafted position) using the teamNeeds object
    teamNeeds[team][pos] += 10;
    for (let j = 0; j < prospects.length; j++) {
      let prospect = prospects[j];
      if (prospect.pos === pos) {
        prospect.val += 10;
      }
    }
    // Update the available prospects for each team (remove the drafted player) using the prospects array
    prospects.sort(sortTeamProspects());
    for (const thisTeam in teamProspects) {
      let theseProspects = teamProspects[thisTeam];
      for (let j = 0; j < theseProspects.length; j++) {
        const thisProspect = theseProspects[j];
        if (thisProspect.name === name && thisProspect.pos === pos) {
          theseProspects.splice(j, 1);
          j--;
        }
      }
    }
    // Update the available prospects for all teams (remove the drafted player) using the Prospects spreadsheet
    // This is critical for allowing proceeding rounds to be simulated
    for (let i = 1; i <= 750; i++) {
      if (rangeProspects.getCell(i, 1).getValue() === pos && rangeProspects.getCell(i, 2).getValue() === name) {
        rangeProspects.getCell(i, 1).setValue("");
        rangeProspects.getCell(i, 2).setValue("");
        break;
      }
    }
  }
  // Update the "needs" for all teams (decrease the need for drafted positions) using the Needs spreadsheet
  // This is critical for allowing proceeding rounds to be simulated
  for (let i = 1; i <= 32; i++) {
    const team = rangeNeeds.getCell(i, 1).getValue();
    const theseNeeds = teamNeeds[team];
    Logger.log(team);
    Logger.log(theseNeeds);
    rangeNeeds.getCell(i, 2).setValue(theseNeeds["QB"]);
    rangeNeeds.getCell(i, 3).setValue(theseNeeds["RB"]);
    rangeNeeds.getCell(i, 4).setValue(theseNeeds["WR"]);
    rangeNeeds.getCell(i, 5).setValue(theseNeeds["TE"]);
    rangeNeeds.getCell(i, 6).setValue(theseNeeds["OL"]);
    rangeNeeds.getCell(i, 7).setValue(theseNeeds["EDGE"]);
    rangeNeeds.getCell(i, 8).setValue(theseNeeds["DL"]);
    rangeNeeds.getCell(i, 9).setValue(theseNeeds["LB"]);
    rangeNeeds.getCell(i, 10).setValue(theseNeeds["CB"]);
    rangeNeeds.getCell(i, 11).setValue(theseNeeds["S"]);
    rangeNeeds.getCell(i, 12).setValue(theseNeeds["K"]);
    rangeNeeds.getCell(i, 13).setValue(theseNeeds["P"]);
  }
}

function sortTeamProspects() {
  return function(a, b) {
    return a.val - b.val;
  }
}
