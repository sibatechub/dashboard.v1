/* =========================================================
   WORK HANDOVER ACTIVE CHECK
========================================================= */

function isBuildingHandoverActive(workRow) {

  const status =
      String(workRow[8] || "")
          .trim()
          .toUpperCase();

  // New WorkLoad status
  if (status === "ACTIVE") {
    return true;
  }

  // Backward compatibility for old YES records
  if (status === "YES") {

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const start = new Date(workRow[5]);
    const end = new Date(workRow[6]);

    if (
      isNaN(start.getTime()) ||
      isNaN(end.getTime())
    ) {
      return false;
    }

    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    return (
      today >= start &&
      today <= end
    );
  }

  return false;
}
function getBuildings(sessionToken) {
  const user = getUserSession(sessionToken);

const isAdmin =
    user.role.toUpperCase() === "ADMIN";

const userName =
    String(user.name).trim();

const today = new Date();
today.setHours(0,0,0,0);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const buildingSheet = ss.getSheetByName("Building_Technician_Map");
  const ticketSheet = ss.getSheetByName("Master");

  const buildingData = buildingSheet.getDataRange().getValues();
  buildingData.shift();

  const ticketData = ticketSheet.getDataRange().getValues();
  ticketData.shift();
  const workloadSheet = ss.getSheetByName("WorkLoad");

const workloadData = workloadSheet.getDataRange().getValues();
workloadData.shift();

  // Store ticket count for each building
  const ticketCount = {};

  ticketData.forEach(function(row){

    const building = row[8];   // Building Name (Column I)

    if (!building) return;

    ticketCount[building] = (ticketCount[building] || 0) + 1;

  });

const result = [];

buildingData.forEach(function(row){

    let technician = row[2];
    let email = row[3];
    let mobile = row[4];

    // Check override
// Check active work handover
workloadData.forEach(function(work){

    if (!isBuildingHandoverActive(work)) return;

    if(
        String(work[1]).trim() !==
        String(row[1]).trim()
    ){
        return;
    }

    technician = work[3];   // Override Technician
    mobile = work[4];       // Override Mobile
    email = work[9];        // Override Email

});

    // Technician Login
    if(!isAdmin){

if (
    String(technician).trim().toUpperCase() !==
    String(user.name).trim().toUpperCase()
){
    return;
}

    }

let override = false;

workloadData.forEach(function(work){

    if (!isBuildingHandoverActive(work)) return;

    if(
        String(work[1]).trim() !==
        String(row[1]).trim()
    ){
        return;
    }

    override = true;

});

result.push({

    slNo: row[0],
    building: row[1],
    technician: technician,
    email: email,
    mobile: mobile,
    status: row[5],
    tickets: ticketCount[row[1]] || 0,
    override: override

});

});

// Highest ticket count first
result.sort(function(a, b){
    return b.tickets - a.tickets;
});

return result;

}
function getBuildingStatistics(sessionToken) {
  const user = getUserSession(sessionToken);

const isAdmin =
    user.role.toUpperCase() === "ADMIN";

const userName =
    String(user.name).trim();
      const today = new Date();
  today.setHours(0,0,0,0);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const buildingSheet = ss.getSheetByName("Building_Technician_Map");
  const workloadSheet = ss.getSheetByName("WorkLoad");

  const buildingData = buildingSheet.getDataRange().getValues();
  buildingData.shift();
    const workloadData = workloadSheet.getDataRange().getValues();
  workloadData.shift();

let totalBuildings = 0;

let activeBuildings = 0;

let technicians = {};

buildingData.forEach(function(row){

    let technician = row[2];

    // Check override
// Check active work handover
workloadData.forEach(function(work){

    if (!isBuildingHandoverActive(work)) return;

    if(
        String(work[1]).trim() !==
        String(row[1]).trim()
    ){
        return;
    }

    technician = work[3];

});

    // Technician Login
    if(!isAdmin){

if(
    String(technician).trim().toUpperCase() !==
    String(userName).trim().toUpperCase()
){
    return;
}

    }

    totalBuildings++;

    if(String(row[5]).toUpperCase() === "YES"){

        activeBuildings++;

    }

    technicians[
    String(technician).trim().toUpperCase()
] = true;

});

let overrideActive = 0;

buildingData.forEach(function(row){

    let technician = row[2];

    // Check override for this building
// Check active work handover
workloadData.forEach(function(work){

    if (!isBuildingHandoverActive(work)) return;

    if(
        String(work[1]).trim() !==
        String(row[1]).trim()
    ){
        return;
    }

    technician = work[3];

});


    // Technician login:
    // Only count override if this building belongs
    // to the logged-in technician after override.
    if(!isAdmin){

        if(
            String(technician).trim().toUpperCase() !==
            userName.toUpperCase()
        ){
            return;
        }

    }


    // Check whether this visible building has an active override
// Check whether this visible building has an active override
workloadData.forEach(function(work){

    if (!isBuildingHandoverActive(work)) return;

    if(
        String(work[1]).trim() !==
        String(row[1]).trim()
    ){
        return;
    }

    overrideActive++;

});

});

  return {

      totalBuildings: totalBuildings,
      activeBuildings: activeBuildings,
      overrideActive: overrideActive,
      technicians: Object.keys(technicians).length

  };

}
function getBuildingForEdit(buildingName) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Building Sheet
  const buildingSheet = ss.getSheetByName("Building_Technician_Map");
  const buildingData = buildingSheet.getDataRange().getValues();

  // Users Sheet
  const usersSheet = ss.getSheetByName("Users");
  const usersData = usersSheet.getDataRange().getValues();

  let building = null;

  // Find selected building
  for (let i = 1; i < buildingData.length; i++) {

    if (String(buildingData[i][1]) == String(buildingName)) {

      building = {
        building: buildingData[i][1],
        technician: buildingData[i][2],
        email: buildingData[i][3],
        mobile: buildingData[i][4],
        status: buildingData[i][5]
      };

      break;

    }

  }

  // Active technicians only
  const technicians = [];

  for (let i = 1; i < usersData.length; i++) {

    const role = String(usersData[i][6]).trim().toUpperCase();
    const status = String(usersData[i][7]).trim().toUpperCase();

    if (role == "TECHNICIAN" && status == "YES") {

      technicians.push({

        name: usersData[i][2],
        email: usersData[i][3],
        mobile: usersData[i][4]

      });

    }

  }

  return {

    building: building,
    technicians: technicians

  };

}
function updateBuilding(building){

  const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName("Building_Technician_Map");

  const data = sheet.getDataRange().getValues();

  for(let i=1;i<data.length;i++){

    if(String(data[i][1]) == String(building.building)){

      // Technician
      sheet.getRange(i+1,3).setValue(building.technician);

      // Email
      sheet.getRange(i+1,4).setValue(building.email);

      // Mobile
      sheet.getRange(i+1,5).setValue(building.mobile);

      // Status
      sheet.getRange(i+1,6).setValue(building.status);

      return true;

    }

  }

  return false;

}
