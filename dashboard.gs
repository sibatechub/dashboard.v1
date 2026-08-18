
/*************************************************
* DASHBOARD SHEETS
*************************************************/




const MASTER_SHEET = "Master";




const USERS_SHEET = "Users";








/*************************************************
* GET MASTER SHEET
*************************************************/




function getMasterSheet(){




 return SpreadsheetApp
     .openById(SPREADSHEET_ID)
     .getSheetByName(MASTER_SHEET);




}








/*************************************************
* GET USERS SHEET
*************************************************/




function getUsersSheet(){




 return SpreadsheetApp
     .openById(SPREADSHEET_ID)
     .getSheetByName(USERS_SHEET);




}








/*************************************************
* DASHBOARD STATISTICS
*************************************************/




function getDashboardStats(sessionToken, filter) {
 filter = filter || {};




const filterDate = filter.date || "";
const filterBuilding = filter.building || "";
const filterTechnician = filter.technician || "";
const filterStatus = filter.status || "";
const filterPriority = filter.priority || "";




const data = getFilteredMasterData(sessionToken);
Logger.log("Dashboard received rows = " + data.length);



/*************************************************
* DATE FILTER
*************************************************/




const filteredData = data.filter(function(row){




const ticketDate = new Date(row[0]);
const status = String(row[2]).trim().toUpperCase();
const priority = String(row[3]).trim().toUpperCase();   // <-- ADD THIS
const building = String(row[8]).trim();
const technician = String(row[15]).trim();




 // Date Filter
 if(filterDate){




   const selectedDate = new Date(filterDate);




   if(
     ticketDate.getFullYear() != selectedDate.getFullYear() ||
     ticketDate.getMonth() != selectedDate.getMonth() ||
     ticketDate.getDate() != selectedDate.getDate()
   ){
     return false;
   }




 }




 // Building Filter
 if(filterBuilding && building != filterBuilding){
   return false;
 }




 // Technician Filter
 if(filterTechnician && technician != filterTechnician){
   return false;
 }




 // Status Filter
 if(filterStatus && status != filterStatus){
   return false;
 }
 // Priority Filter
if(filterPriority && priority != filterPriority){
   return false;
}




 return true;




});




let totalTickets = filteredData.length;




let openTickets = 0;
let acceptedTickets = 0;
let pendingTickets = 0;
let reopenTickets = 0;
let closedTickets = 0;
let immediateTickets = 0;




 let buildingSet = new Set();
 let technicianSet = new Set();




 let ratingTotal = 0;
 let ratingCount = 0;




filteredData.forEach(function(row){




   // Status (Column C)
   const status = String(row[2]).trim().toUpperCase();




   // Priority (Column D)
   const priority = String(row[3]).trim().toUpperCase();




   // Building (Column I)
   const building = String(row[8]).trim();




   // Technician (Column P)
   const technician = String(row[15]).trim();




   // Rating (Column W)
   const rating = Number(row[22]);




if(status == "OPEN"){
   openTickets++;
}




if(status == "ACCEPTED"){
   acceptedTickets++;
}




if(status == "PENDING"){
   pendingTickets++;
}




if(status == "REOPEN"){
   reopenTickets++;
}




if(status == "CLOSED"){
   closedTickets++;
}




if(




   priority == "IMMEDIATE" &&




   (
       status == "OPEN" ||
       status == "ACCEPTED" ||
       status == "PENDING" ||
       status == "REOPEN"
   )




){




   immediateTickets++;




}




   if(building != "")
     buildingSet.add(building);




   if(technician != "")
     technicianSet.add(technician);




   if(!isNaN(rating) && rating > 0){




     ratingTotal += rating;
     ratingCount++;




   }




 });




 let averageRating = 0;




 if(ratingCount > 0){




   averageRating = (ratingTotal / ratingCount).toFixed(1);




 }




return {




   totalTickets: totalTickets,




   openTickets: openTickets,




   acceptedTickets: acceptedTickets,




   pendingTickets: pendingTickets,




   reopenTickets: reopenTickets,




   closedTickets: closedTickets,




   immediateTickets: immediateTickets




};




}




/*************************************************
* RECENT TICKETS
*************************************************/




function getRecentTickets(sessionToken){




const data = getFilteredMasterData(sessionToken);
 data.reverse();




 const recent = [];
 const limit = Math.min(data.length,10);




 for(let i=0;i<limit;i++){




   const row = data[i];




   Logger.log("Recent Ticket = [" + row[1] + "] Status = " + row[2]);




   recent.push({




     ticketNo: String(row[1]).trim(),




     status: row[2],




     priority: row[3],




     building: row[8],




     faculty: row[5],




     assetType: row[11],




     problem: row[13],




     technician: row[15]




   });




 }




 return recent;




}
/*************************************************
* GET TICKET DETAILS
*************************************************/




function getTicketDetails(sessionToken, ticketNo){
  const session = getUserSession(sessionToken);

if (!session) {
  throw new Error("Session Expired");
}




 const sheet = getMasterSheet();




 const data = sheet.getDataRange().getValues();




 data.shift();




for (let i = 0; i < data.length; i++) {




 const row = data[i];




 const sheetTicket = String(row[1]).trim();
 const clickedTicket = String(ticketNo).trim();




 if (sheetTicket === clickedTicket) {




   Logger.log("MATCH FOUND at Row = " + (i + 2));




return {




 ticketNo: row[1],
 status: row[2],
 priority: row[3],
 createdTime: row[0]
 ? Utilities.formatDate(
     new Date(row[0]),
     Session.getScriptTimeZone(),
     "dd-MMM-yyyy hh:mm:ss a"
   )
 : "",
 faculty: row[5],
 mobile: row[6],
 email: row[7],
 building: row[8],
 floor: row[9],
 room: row[10],
 assetType: row[11],
 assetNo: row[12],
 problem: row[13],
 technician: row[15],
 technicianEmail: row[16],
 technicianMobile: row[17],
 acceptedTime: row[18]
 ? Utilities.formatDate(
     new Date(row[18]),
     Session.getScriptTimeZone(),
     "dd-MMM-yyyy hh:mm:ss a"
   )
 : "",




closeTime: row[19]
 ? Utilities.formatDate(
     new Date(row[19]),
     Session.getScriptTimeZone(),
     "dd-MMM-yyyy hh:mm:ss a"
   )
 : "",
duration: row[20],
comments: row[21],




floor: row[9],
room: row[10],




overallSatisfaction: row[22],
technicianBehaviour: row[23],
resolutionSatisfaction: row[24],




feedbackSubmitTime: row[30]
 ? Utilities.formatDate(
     new Date(row[30]),
     Session.getScriptTimeZone(),
     "dd-MMM-yyyy hh:mm:ss a"
   )
 : "",




itHeadRemarks: row[29]




};




 }




}




Logger.log("NO MATCH FOUND");




return null;
}
function getActiveTickets(sessionToken, filter) {




const data = getFilteredMasterData(sessionToken);
 data.reverse();




 const tickets = [];
 filter = filter || {};




const filterStatus = filter.status || "";
const filterBuilding = filter.building || "";
const filterTechnician = filter.technician || "";
const filterDate = filter.date || "";
const filterPriority = filter.priority || "";




 data.forEach(function(row){




const status = String(row[2]).trim().toUpperCase();
const priority = String(row[3]).trim().toUpperCase();   // <-- ADD
const building = String(row[8]).trim();
const technician = String(row[15]).trim();
const ticketDate = new Date(row[0]);




// Status Filter
if(filterStatus && status != filterStatus){
   return;
}
// Priority Filter
if(filterPriority && priority != filterPriority){
   return;
}




// Building Filter
if(filterBuilding && building != filterBuilding){
   return;
}




// Technician Filter
if(filterTechnician && technician != filterTechnician){
   return;
}




// Date Filter
if(filterDate){




   const d = new Date(filterDate);




   if(
       ticketDate.getFullYear() != d.getFullYear() ||
       ticketDate.getMonth() != d.getMonth() ||
       ticketDate.getDate() != d.getDate()
   ){
       return;
   }




}




     tickets.push({




       ticketNo: row[1],
       status: row[2],
       priority: row[3],
       building: row[8],
       faculty: row[5],
       assetType: row[11],
       problem: row[13],
       technician: row[15]




     });




 });




 return tickets;




}
/*************************************************
* BUILDING WISE CHART DATA
*************************************************/
function getBuildingChartData(sessionToken, filter) {




 const sheet = getMasterSheet();
 filter = filter || {};




const filterStatus = filter.status || "";
const filterBuilding = filter.building || "";
const filterTechnician = filter.technician || "";
const filterDate = filter.date || "";
const filterPriority = filter.priority || "";
const data = getFilteredMasterData(sessionToken);




 const buildingData = {};




 data.forEach(function(row) {




const building = String(row[8]).trim();
const status = String(row[2]).trim().toUpperCase();
const priority = String(row[3]).trim().toUpperCase();
const technician = String(row[15]).trim();
const ticketDate = new Date(row[0]);
// Status Filter
if(filterStatus && status != filterStatus) return;




// Priority Filter
if(filterPriority && priority != filterPriority) return;




// Building Filter
if(filterBuilding && building != filterBuilding) return;




// Technician Filter
if(filterTechnician && technician != filterTechnician) return;




// Date Filter
if(filterDate){




   const d = new Date(filterDate);




   if(
       ticketDate.getFullYear() != d.getFullYear() ||
       ticketDate.getMonth() != d.getMonth() ||
       ticketDate.getDate() != d.getDate()
   ){
       return;
   }




}




   if (!building) return;




   if (!buildingData[building]) {




     buildingData[building] = {
       building: building,
       open: 0,
       accepted: 0,
       pending: 0,
       reopen: 0,
       closed: 0
     };




   }




   switch(status){




     case "OPEN":
       buildingData[building].open++;
       break;




     case "ACCEPTED":
       buildingData[building].accepted++;
       break;




     case "PENDING":
       buildingData[building].pending++;
       break;




     case "REOPEN":
       buildingData[building].reopen++;
       break;




     case "CLOSED":
       buildingData[building].closed++;
       break;




   }




 });




 return Object.values(buildingData);




}
function testBuildingChart(){




 Logger.log(
   JSON.stringify(getBuildingChartData(), null, 2)
 );




}
/*************************************************
* TECHNICIAN WISE CHART DATA
*************************************************/
function getTechnicianChartData(sessionToken, filter) {




 const sheet = getMasterSheet();
 filter = filter || {};




const filterStatus = filter.status || "";
const filterBuilding = filter.building || "";
const filterTechnician = filter.technician || "";
const filterDate = filter.date || "";
const filterPriority = filter.priority || "";
const data = getFilteredMasterData(sessionToken);




 const technicianData = {};




 data.forEach(function(row){




const technician = String(row[15]).trim();
const status = String(row[2]).trim().toUpperCase();
const priority = String(row[3]).trim().toUpperCase();
const building = String(row[8]).trim();
const ticketDate = new Date(row[0]);
// Status Filter
if(filterStatus && status != filterStatus) return;




// Priority Filter
if(filterPriority && priority != filterPriority) return;




// Building Filter
if(filterBuilding && building != filterBuilding) return;




// Technician Filter
if(filterTechnician && technician != filterTechnician) return;




// Date Filter
if(filterDate){




   const d = new Date(filterDate);




   if(
       ticketDate.getFullYear() != d.getFullYear() ||
       ticketDate.getMonth() != d.getMonth() ||
       ticketDate.getDate() != d.getDate()
   ){
       return;
   }




}




   if (!technician) return;




   if (!technicianData[technician]) {




     technicianData[technician] = {




       technician: technician,




       open: 0,
       accepted: 0,
       pending: 0,
       reopen: 0,
       closed: 0




     };




   }




   switch(status){




     case "OPEN":
       technicianData[technician].open++;
       break;




     case "ACCEPTED":
       technicianData[technician].accepted++;
       break;




     case "PENDING":
       technicianData[technician].pending++;
       break;




     case "REOPEN":
       technicianData[technician].reopen++;
       break;




     case "CLOSED":
       technicianData[technician].closed++;
       break;




   }




 });




 return Object.values(technicianData);




}
function testTechnicianChart(){




 Logger.log(
   JSON.stringify(getTechnicianChartData(), null, 2)
 );




}
/*************************************************
* LOAD DASHBOARD FILTERS
*************************************************/




function getDashboardFilters(sessionToken){




const data = getFilteredMasterData(sessionToken);




 const buildingSet = new Set();
 const technicianSet = new Set();




 data.forEach(function(row){




   const building = String(row[8]).trim();
   const technician = String(row[15]).trim();




   if(building){
     buildingSet.add(building);
   }




   if(technician){
     technicianSet.add(technician);
   }




 });




 return {




   buildings: [...buildingSet].sort(),




   technicians: [...technicianSet].sort()




 };




}


function getUsers() {


 const ss = SpreadsheetApp.openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524");
 const sheet = ss.getSheetByName("Users");


 const data = sheet.getDataRange().getValues();


 const users = [];


 for (let i = 1; i < data.length; i++) {


   users.push({
     userId: data[i][1],
     name: data[i][2],
     email: data[i][3],
     mobile: data[i][4],
     password:data[i][5],
     role: data[i][6],
     status: data[i][7]
   });


 }


 return users;


}
function getUserDetails(userId) {


 const ss = SpreadsheetApp.openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524");
 const sheet = ss.getSheetByName("Users");


 const data = sheet.getDataRange().getValues();


 for (let i = 1; i < data.length; i++) {


   if (String(data[i][1]) === String(userId)) {


     return {
       slNo: data[i][0],
       userId: data[i][1],
       name: data[i][2],
       email: data[i][3],
       mobile: data[i][4],
       password: data[i][5],
       role: data[i][6],
       status: data[i][7]
     };


   }


 }


 return null;


}
function approveUserAccount(userId) {


 const ss = SpreadsheetApp.openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524");
 const sheet = ss.getSheetByName("Users");


 const data = sheet.getDataRange().getValues();


 for (let i = 1; i < data.length; i++) {


   if (String(data[i][1]) == String(userId)) {


     sheet.getRange(i + 1, 8).setValue("YES"); // Column H = Status
     sendAccountStatusEmail(userId, "YES");


     return true;


   }


 }


 return false;


}
function rejectUserAccount(userId) {


 const ss = SpreadsheetApp.openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524");
 const sheet = ss.getSheetByName("Users");


 const data = sheet.getDataRange().getValues();


 for (let i = 1; i < data.length; i++) {


   if (String(data[i][1]) === String(userId)) {


     sheet.getRange(i + 1, 8).setValue("Rejected");
     sendAccountStatusEmail(userId, "REJECTED");


     return true;


   }


 }


 return false;


}
function disableUserAccount(userId) {


 const ss = SpreadsheetApp.openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524");
 const sheet = ss.getSheetByName(USER_SHEET);


 const data = sheet.getDataRange().getValues();


 for (let i = 1; i < data.length; i++) {


   if (String(data[i][1]) === String(userId)) {


     sheet.getRange(i + 1, 8).setValue("DISABLED");
     sendAccountStatusEmail(userId, "DISABLED");


     return true;


   }


 }


 return false;


}
function enableUserAccount(userId) {


 const ss = SpreadsheetApp.openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524");
 const sheet = ss.getSheetByName(USER_SHEET);


 const data = sheet.getDataRange().getValues();


 for (let i = 1; i < data.length; i++) {


   if (String(data[i][1]) === String(userId)) {


     sheet.getRange(i + 1, 8).setValue("YES");
     sendAccountStatusEmail(userId, "YES");


     return true;


   }


 }


 return false;


}
function sendAccountStatusEmail(userId, status) {


 const sheet = getUserSheet();
 const data = sheet.getDataRange().getValues();


 for (let i = 1; i < data.length; i++) {


   if (String(data[i][1]) === String(userId)) {


     const name = data[i][2];
     const email = data[i][3];
     const role = data[i][6];


     let title = "";
     let message = "";
     let color = "#2E2278";


     switch (status) {


case "YES":


 title = "🎉 DRIEMS IT Dashboard Account Approved";


 message = `
  <p>Dear <b>${name}</b>,</p>


 <p>
 Congratulations! Your account has been approved by the DRIEMS IT Administrator.
 You can now access the <b>DRIEMS IT Dashboard</b>.
 </p>


 <table style="
     border-collapse:collapse;
     width:100%;
     margin-top:15px;
     font-size:15px;">


     <tr>
         <td style="padding:8px;border:1px solid #ddd;"><b>Name</b></td>
         <td style="padding:8px;border:1px solid #ddd;">${name}</td>
     </tr>


     <tr>
         <td style="padding:8px;border:1px solid #ddd;"><b>User ID</b></td>
         <td style="padding:8px;border:1px solid #ddd;">${userId}</td>
     </tr>


     <tr>
         <td style="padding:8px;border:1px solid #ddd;"><b>Email</b></td>
         <td style="padding:8px;border:1px solid #ddd;">${email}</td>
     </tr>


     <tr>
         <td style="padding:8px;border:1px solid #ddd;"><b>Role</b></td>
         <td style="padding:8px;border:1px solid #ddd;">${role}</td>
     </tr>


     <tr>
         <td style="padding:8px;border:1px solid #ddd;"><b>Password</b></td>
         <td style="
             padding:8px;
             border:1px solid #ddd;
             font-weight:bold;
             color:#d63384;">
             ${data[i][5]}
         </td>
     </tr>


 </table>


 <br>


 <div style="
     background:#eef7ff;
     border-left:5px solid #2E2278;
     padding:15px;
     border-radius:6px;">


     <b>Dashboard Login</b><br><br>


     <a href="https://script.google.com/a/macros/driems.ac.in/s/AKfycbyrpGmpIamoP6L1L5gkKurk7M-JF6fcBXr7osKK_8NI57D1glVlKRWFfY7LsRqcAIE/exec"
        style="
        background:#2E2278;
        color:white;
        padding:10px 18px;
        text-decoration:none;
        border-radius:5px;">
        Login to Dashboard
     </a>


 </div>


 <br>


 <div style="
     background:#fff8e1;
     padding:12px;
     border-left:5px solid orange;">


     <b>Security Notice</b><br>


     Please change your password after your first login to keep your account secure.


 </div>


 `;


 color = "#28A745";


 break;


       case "REJECTED":
         title = "❌ Account Rejected";
         message = `
           <p>Hello <b>${name}</b>,</p>


           <p>We regret to inform you that your account registration request has been rejected by the Administrator.</p>


           <p>If you believe this is a mistake, please contact the IT Department.</p>
         `;
         color = "#DC3545";
         break;


       case "DISABLED":
         title = "⛔ Account Disabled";
         message = `
           <p>Hello <b>${name}</b>,</p>


           <p>Your DRIEMS IT Dashboard account has been temporarily <b>disabled</b>.</p>


           <p>Please contact the IT Administrator for further assistance.</p>
         `;
         color = "#6C757D";
         break;


       default:
         return;
     }


     const html = `
     <div style="font-family:Arial;padding:25px">


         <h2 style="color:${color}">
             DRIEMS IT Dashboard
         </h2>


         <h3>${title}</h3>


         ${message}


         <hr>


         <table style="border-collapse:collapse">


             <tr>
                 <td><b>User ID</b></td>
                 <td>${userId}</td>
             </tr>


             <tr>
                 <td><b>Role</b></td>
                 <td>${role}</td>
             </tr>


             <tr>
                 <td><b>Status</b></td>
                 <td>${status}</td>
             </tr>


         </table>


         <br>


         <p style="font-size:12px;color:gray;">
         This is an automated email from DRIEMS IT Dashboard.
         Please do not reply.
         </p>


     </div>
     `;


     MailApp.sendEmail({
       to: email,
       subject: title,
       htmlBody: html
     });


     return;


   }


 }


}
/*************************************************
* UPDATE USER
*************************************************/
/*************************************************
* UPDATE USER
*************************************************/
function updateUser(user){

  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();

  for(let i = 1; i < data.length; i++){

    if(String(data[i][1]).trim() === String(user.userId).trim()){

      // Name - Column C
      sheet.getRange(i + 1, 3).setValue(user.name);

      // Email - Column D
      sheet.getRange(i + 1, 4).setValue(user.email);

      // Mobile - Column E
      sheet.getRange(i + 1, 5).setValue(user.mobile);

      // Password - Column F
      sheet.getRange(i + 1, 6).setValue(user.password);

      // Role - Column G
      sheet.getRange(i + 1, 7).setValue(user.role);

      SpreadsheetApp.flush();

      return true;
    }
  }

  return false;
}

/*************************************************
* GET ALL UNIQUE ROLES
*************************************************/
function getRoles() {


 const sheet = getUserSheet();


 const data = sheet.getDataRange().getValues();


 let roles = [];


 for (let i = 1; i < data.length; i++) {


   const role = String(data[i][6]).trim();


   if (
     role &&
     role.toUpperCase() !== "ADMIN" &&
     !roles.includes(role)
   ) {
     roles.push(role);
   }


 }


 roles.sort();


 return roles;


}

function getFilteredMasterData(sessionToken) {

  Logger.log("========== FILTER START ==========");

  Logger.log("Incoming Token = " + sessionToken);

  const session = getUserSession(sessionToken);

  Logger.log("Session = " + JSON.stringify(session));

  if (!session) {
    Logger.log("SESSION IS NULL");
    throw new Error("Session expired");
  }

  const sheet = getMasterSheet();

  Logger.log("Sheet Name = " + sheet.getName());

  const data = sheet.getDataRange().getValues();

  Logger.log("Total Rows Including Header = " + data.length);

  data.shift();

  Logger.log("Rows After Header Removed = " + data.length);

  Logger.log("Role = " + session.role);

  if (String(session.role).trim().toUpperCase() == "ADMIN") {

    Logger.log("ADMIN RETURNING " + data.length + " ROWS");

    return data;

  }

  if (String(session.role).trim().toUpperCase() == "TECHNICIAN") {

    const filtered = data.filter(function(row){

      return String(row[16]).trim().toLowerCase() ==
             String(session.email).trim().toLowerCase();

    });

    Logger.log("TECHNICIAN RETURNING " + filtered.length + " ROWS");

    return filtered;

  }

  Logger.log("UNKNOWN ROLE");

  return [];

}




