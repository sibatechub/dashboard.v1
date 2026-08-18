/*************************************************
* DRIEMS IT Dashboard
* Authentication Module
*************************************************/


const SPREADSHEET_ID = "1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524";
const USER_SHEET = "Users";




/*************************************************
* Get User Sheet
*************************************************/
function getUserSheet() {


 return SpreadsheetApp
   .openById(SPREADSHEET_ID)
   .getSheetByName(USER_SHEET);


}




/*************************************************
* LOGIN
*************************************************/
function loginUser(emailOrUserId, password) {


 const sheet = getUserSheet();
 const data = sheet.getDataRange().getValues();


 const input = emailOrUserId.toString().trim().toLowerCase();


 for (let i = 1; i < data.length; i++) {


   const row = data[i];


const user = {
 userId: row[1].toString().trim().toLowerCase(),
 name: row[2],
 email: row[3].toString().trim().toLowerCase(),
 mobile: row[4].toString(),
 password: row[5].toString(),
 role: row[6],
 status: row[7].toString().trim().toUpperCase()
};


   // Login using either UserID OR Email
   if (user.email === input || user.userId === input) {


     if (user.password !== password) {
       return {
         success: false,
         message: "Wrong Password"
       };
     }


// Check Account Status
if (user.status == "PENDING") {
 return {
   success: false,
   message: "Your account is pending approval."
 };
}


if (user.status == "REJECTED") {
 return {
   success: false,
   message: "Your account has been rejected."
 };
}


if (user.status == "DISABLED") {
 return {
   success: false,
   message: "Your account has been disabled. Please contact the Administrator."
 };
}


// SAVE SESSION
// CREATE NEW SESSION
const token = createSession({

  userId: user.userId,
  name: user.name,
  email: user.email,
  role: user.role

});

return {

  success: true,

  token: token,

  userId: user.userId,

  name: user.name,

  email: user.email,

  role: user.role

};
   }
 }


 return {
   success: false,
   message: "User Not Found"
 };
}




/*************************************************
* SIGNUP
*************************************************/
function createAccount(obj) {


 const sheet = getUserSheet();


 const data = sheet.getDataRange().getValues();


for (let i = 1; i < data.length; i++) {


 // Check Duplicate User ID
 if (
   data[i][1].toString().trim().toLowerCase() ==
   obj.userId.trim().toLowerCase()
 ) {


   return {


     success:false,


     message:"User ID Already Exists"


   };


 }


 // Check Duplicate Email
 if (
   data[i][3].toString().trim().toLowerCase() ==
   obj.email.trim().toLowerCase()
 ) {


   return {


     success:false,


     message:"Email Already Registered"


   };


 }


}




 const nextRow = sheet.getLastRow();


 const slNo = nextRow;


 const userId = obj.userId.trim();


sheet.appendRow([


 slNo,


 userId,


 obj.name,


 obj.email,


 obj.mobile,


 obj.password,


 obj.role,


 "PENDING"


]);




 return{


   success:true,


   message:"Account Created Successfully"


 };


}
/*************************************************
* SEND OTP
*************************************************/
function sendOTP(emailOrUserId) {


 const sheet = getUserSheet();
 const data = sheet.getDataRange().getValues();


 const input = emailOrUserId.toString().trim().toLowerCase();


 for (let i = 1; i < data.length; i++) {


   const row = data[i];


   const userId = row[1].toString().trim().toLowerCase();
   const name = row[2];
   const email = row[3].toString().trim().toLowerCase();


if (userId == input || email == input) {


     // Check Account Status
     const status = row[7].toString().trim().toUpperCase();


     if (status !== "YES") {


       return {


         success: false,


         message: "Your account is not active. Please contact the Administrator."


       };


     }


     // Generate 6 Digit OTP
     const otp = Math.floor(100000 + Math.random() * 900000);


     // OTP Expiry (10 Minutes)
     const expiry = new Date(Date.now() + 10 * 60 * 1000);


     // Save OTP
     sheet.getRange(i + 1, 9).setValue(otp);


     // Save Expiry
     sheet.getRange(i + 1, 10).setValue(expiry);


     // Email Body
     const html = `
     <div style="font-family:Arial;padding:25px">


       <h2 style="color:#2E2278;">
         DRIEMS IT Breakdown Ticket System
       </h2>


       <p>Hello <b>${name}</b>,</p>


       <p>Your One-Time Password (OTP) is</p>


       <h1 style="
           color:#2E2278;
           letter-spacing:5px;
           font-size:40px;">
           ${otp}
       </h1>


       <p>
         This OTP is valid for
         <b>10 Minutes</b>.
       </p>


       <hr>


       <p style="font-size:13px;color:gray;">
         Please do not share this OTP with anyone.
       </p>


     </div>
     `;


     MailApp.sendEmail({


       to: email,


       subject: "DRIEMS IT Dashboard - Password Reset OTP",


       htmlBody: html


     });


     return {


       success: true,


       message: "OTP Sent Successfully"


     };


}


 }


 return {


   success:false,


   message:"User Not Found"


 };


}
/*************************************************
* VERIFY OTP & RESET PASSWORD
*************************************************/
function resetPassword(user, otp, newPassword) {


 const sheet = getUserSheet();
 const data = sheet.getDataRange().getValues();


 const input = user.toString().trim().toLowerCase();


 for (let i = 1; i < data.length; i++) {


   const userId = data[i][1].toString().trim().toLowerCase();
   const email  = data[i][3].toString().trim().toLowerCase();


   if (userId == input || email == input) {


     const savedOTP = data[i][8].toString();
     const expiry   = new Date(data[i][9]);


     // OTP Incorrect
     if (savedOTP != otp.toString()) {


       return {


         success:false,


         message:"Invalid OTP"


       };


     }


     // OTP Expired
     if (new Date() > expiry) {


       return {


         success:false,


         message:"OTP Expired"


       };


     }


     // Update Password
     sheet.getRange(i + 1, 6).setValue(newPassword);


     // Clear OTP
     sheet.getRange(i + 1, 9).clearContent();


     // Clear Expiry
     sheet.getRange(i + 1,10).clearContent();


     return {


       success:true,


       message:"Password Reset Successfully"


     };


   }


 }


 return {


   success:false,


   message:"User Not Found"


 };


}



