function doGet(e) {


 const page = (e && e.parameter.page) ? e.parameter.page : "login";


 // Dashboard
 if (page == "dashboard") {


   return HtmlService
     .createTemplateFromFile("DashboardPage")
     .evaluate()
     .setTitle("DRIEMS IT Dashboard")
     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


 }


 // Login Page
 return HtmlService
   .createTemplateFromFile("Login")
   .evaluate()
   .setTitle("DRIEMS IT Breakdown Ticket System")
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}




// Include HTML
function include(filename){


 return HtmlService
 .createHtmlOutputFromFile(filename)
 .getContent();


}
function testDriveAPI() {
 const result = Drive.Files.list({
   pageSize: 5
 });


 Logger.log(result);
}







