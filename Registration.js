//july 2025
//credits to STEPHEN (GOAT) for building most of it last year, this year main changes I did was just how the form was filled in
//instead of pulling responses based on array, pulled directly based on question (in case things moved around / got added) 
//also changed how it filled the pdf so that the final emailed version was an editable pdf & ppl didn't have to resubmit the form for minor edits!
//note if you do use this, should probably build a trigger!

const folderId = "1KStt_HkikCK68Un8TpuL8SVe1QMgBlYo";
const templateFileId = '1mpd-u2LyzmFeD9YRocb-M3Zeg8nqodq6'; // PDF template file ID

function formHandler(e) {
  const person = buildPersonFromResponse(e.response);
  replaceAndSend(person);
}

//Generates the PDF, saves it to Drive, and emails it to the respondent.

function replaceAndSend(person) {
  const pdfBlob = DriveApp.getFileById(templateFileId).getBlob();
  const pdfForm = new PdfForm();
  const replaceArray = buildReplaceArray(person);

  pdfForm.setValues(pdfBlob, replaceArray)
    .then(updatedBlob => {
      const fileName = `${person.firstName} ${person.lastName}'s Registration 2025-26.pdf`;

      const file = DriveApp.getFolderById(folderId)
        .createFile(updatedBlob)
        .setName(fileName);

      // Set permissions: Anyone with the link can view
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;

      Logger.log("About to send email to: " + person.contact1.email);
      MailApp.sendEmail({
        to: person.contact1.email,
        subject: "Completed Payton Registration Documents",
        body: `Your registration file is ready! Please view and print it using this link: ${fileUrl}. Please note that several fields are not filled out, specifically: *insert*. Additionally, you need to manually sign and date each place that requires it. We look forward to seeing you at registration :) ` ,
        name: "Walter Payton College Prep"
      });

      Logger.log("Email sent successfully");
    })
    .catch(error => {
      Logger.log('Failed to generate or send PDF: ' + error);
      throw new Error('Failed to generate or send PDF: ' + error);
    });
}


function buildPersonFromResponse(formResponse) {
  const itemResponses = formResponse.getItemResponses();
  const responseMap = new Map();

  itemResponses.forEach(itemResponse => {
    const title = itemResponse.getItem().getTitle().trim();
    const response = itemResponse.getResponse();
    responseMap.set(title, response);
  });

  return new Person(responseMap);
}

function Person(responseMap) {
  this.firstName = responseMap.get("Student First Name");
  this.lastName = responseMap.get("Student Last Name");
  this.middleName = responseMap.get("Student Middle Name");
  this.dob = formatDateString(responseMap.get("Student Birth Date"));
  this.gender = responseMap.get("Student Gender (F / M / X / N)");
  this.address = responseMap.get("Student Street Address, Primary Residence");
  this.city = responseMap.get("City");
  this.state = responseMap.get("State");
  this.zip = responseMap.get("ZIP");
  this.county = responseMap.get("County");
  this.email = responseMap.get("Student Email");
  this.siblingsAtPayton = responseMap.get("Names of any siblings at Payton, including incoming freshmen");
  this.grade = responseMap.get("Grade student will be in for the 25-26 school year");
  this.studentId = responseMap.get("CPS Student ID");
  this.advisory = responseMap.get("Advisory Number")
  this.confidential1 = responseMap.get("Confidential Information #1");
  this.confidential2 = responseMap.get("Confidential Information #2") === "Yes";
  this.contact1 = {
    firstName: responseMap.get("Contact #1 First Name"),
    lastName: responseMap.get("Contact #1 Last Name"),
    email: responseMap.get("Contact #1 Email"),
    relationship: responseMap.get("Contact #1 Relationship to Student"),
    permissions: responseMap.get("Contact #1 Check all that apply"),
    address: responseMap.get("Contact #1 Home Street Address"),
    city: responseMap.get("Contact #1 City"),
    state : responseMap.get("Contact #1 State"),
    zip : responseMap.get("Contact #1 Zip"),
    homePhone : responseMap.get("Contact #1 Home Phone"),
    cellPhone : responseMap.get("Contact #1 Cell Phone Number"),
    workPhone : responseMap.get("Contact #1 Work Phone Number"),
    prefer : responseMap.get("Contact #1 Preferred Phone Number"),
    lang : responseMap.get("Contact #1 Communication Language"),
    trans : responseMap.get("Contact #1 Requires Translator?") === "Yes",
};

const phoneMap1 = {
  "Cell": responseMap.get("Contact #1 Cell Phone Number"),
  "Home": responseMap.get("Contact #1 Home Phone"),
  "Work": responseMap.get("Contact #1 Work Phone Number"),
};
this.contact1.primaryPhone = phoneMap1[this.contact1.prefer] || responseMap.get("Contact #1 Cell Phone Number");

  this.contact2Exists = responseMap.get("Would you like to add another parent/guardian?") === "Yes";

  this.contact2 = this.contact2Exists ? {
    firstName: responseMap.get("Contact #2 First Name"),
    lastName: responseMap.get("Contact #2 Last Name"),
    email: responseMap.get("Contact #2 Email"),
    relationship: responseMap.get("Contact #2 Relationship to Student"),
    permissions: responseMap.get("Contact #2 Check all that apply"),
    address: responseMap.get("Contact #2 Home Street Address"),
    city: responseMap.get("Contact #2 City"),
    state : responseMap.get("Contact #2 State"),
    zip : responseMap.get("Contact #2 Zip"),
    homePhone : responseMap.get("Contact #2 Home Phone"),
    cellPhone : responseMap.get("Contact #2 Cell Phone Number"),
    workPhone : responseMap.get("Contact #2 Work Phone Number"),
    prefer : responseMap.get("Contact #2 Preferred Phone Number"),
    lang : responseMap.get("Contact #2 Communication Language"),
    trans : responseMap.get("Contact #2 Requires Translator?") === "Yes",
  }: {};
  
  const phoneMap2 = {
  "Cell": responseMap.get("Contact #2 Cell Phone Number"),
  "Home": responseMap.get("Contact #2 Home Phone"),
  "Work": responseMap.get("Contact #2 Work Phone Number"),
};
this.contact2.primaryPhone = phoneMap2[this.contact2.prefer] || responseMap.get("Contact #2 Cell Phone Number");
  this.emergencyContact ={
    name : responseMap.get("Relative/Neighbor Name"),
    address : responseMap.get("Relative/Neighbor Address"),
    phone : responseMap.get("Relative/Neighbor Phone"),
    relationship : responseMap.get("Relative/Neighbor Relationship"),
  }

  this.doctor = {
    name : responseMap.get("Doctor Name"),
    phone : responseMap.get("Doctor Phone"),
    address : responseMap.get("Doctor Address"),
    city : responseMap.get("Doctor Address City"),
    state : responseMap.get("Doctor Address State"),
    zip : responseMap.get("Doctor Address Zip"),
    consent : responseMap.get("I authorize you to call my family doctor, if necessary, in an emergency") === "Yes",

  }

  this.studentHealthInsurance = responseMap.get("Student Health Insurance");
  this.medicalID = responseMap.get("Medical ID");
  this.apply  = responseMap.get("Are you interested in applying for the Illinois Medical Card/All Kids?") === "Yes";

  this.armedForces = responseMap.get("Armed Forces") === "Yes";
  this.deploy = responseMap.get("Deployment") === "Yes";

}


/**
 * Formats a date string (YYYY-MM-DD) into MM/dd/yyyy. 
 */
function formatDateString(dateString, timeZone = Session.getScriptTimeZone()) {
  if (!dateString) return "";
  const date = new Date(dateString);
  return Utilities.formatDate(date, timeZone, "MM/dd/yyyy");
}


/**
 * Builds the replace array for PdfForm.setValues(), including text and checkbox mappings.
 */
function buildReplaceArray(person) {
  const textMappings = [
    { name: "Fname", value: person.firstName },
    { name: "Lname", value: person.lastName },
    { name: "Name", value: `${person.firstName} ${person.lastName}` },
    { name: "Mname", value: person.middleName },
    { name: "gender", value: person.gender },
    { name: "dob", value: person.dob },
    { name: "ID", value: person.studentId },
    { name: "grade", value: person.grade },
    { name: "room", value: person.advisory },
    { name: "address", value: person.address },
    { name: "city", value: person.city },
    { name: "state", value: person.state },
    { name: "zip", value: person.zip },
    { name: "P1Fname", value: person.contact1.firstName },
    { name: "P1Lname", value: person.contact1.lastName },
    { name: "P1Name", value: `${person.contact1.firstName} ${person.contact1.lastName}` },
    { name: "P1P1", value: person.contact1.primaryPhone },  // Preferred goes first
    { name: "P1P2", value: [person.contact1.cellPhone, person.contact1.homePhone, person.contact1.workPhone].find(p => p !== person.contact1.primaryPhone) || "" },
    { name: "P1E", value: person.contact1.email },
    { name: "P1R", value: person.contact1.relationship },
    { name: "P1Add", value: person.contact1.address },
    { name: "P1CL", value: person.contact1.lang },
    { name: "TName", value: person.emergencyContact.name },
    { name: "Tadd", value: person.emergencyContact.address },
    { name: "Tnumber", value: person.emergencyContact.phone },
    { name: "Trel", value: person.emergencyContact.relationship },
    { name: "IMC#", value: person.medicalID },
    { name: "DName", value: person.doctor.name },
    { name: "DNumber", value: person.doctor.phone },
    { name: "dAdd", value: person.doctor.address },
    { name: "dCity", value: person.doctor.city},
    { name: "dState", value: person.doctor.state},
    { name: "dZip", value: person.doctor.zip}
  ];

  const checkboxMappings = [
    { name: "SCCI1", value: person.confidential1 === "in a car/park/other public place" },
    { name: "SCCI2", value: person.confidential1 === "doubled-up" },
    { name: "SCCI3", value: person.confidential1 === "in a hotel/motel" },
    { name: "SCCI4", value: person.confidential1 === "in a shelter" },
    { name: "SCCI5", value: person.confidential1 === "in transitional housing" },
    { name: "SCCI11", value: person.confidential2 },
    { name: "SCCI12", value: person.confidential2 },
    { name: "SHIMC", value: person.studentHealthInsurance === "Illinois Medical Card/All Kids" },
    { name: "SHNI", value: person.studentHealthInsurance === "No Insurance" },
    { name: "SHP", value: person.studentHealthInsurance === "Private/Employer Health Insurance" },
    { name: "MP1", value: person.armedForces },
    { name: "MP2", value: !person.armedForces},
    { name: "AC1", value: person.deploy },
    { name: "AC2", value: !person.deploy },
    { name: "P1P1C", value: person.contact1.prefer === "Cell"},
    { name: "P1P1H", value: person.contact1.prefer === "Home"},
    { name: "P1P1W", value: person.contact1.prefer === "Work"},
    { name: "DC1", value: person.doctor.consent},
    { name: "DC2", value: !person.doctor.consent},
    { name: "P1T1", value : person.contact1.trans},
    { name: "P1T2", value : !person.contact1.trans}
  ];

  if (person.contact1.permissions && person.contact1.permissions.length > 0) {
    checkboxMappings.push(
      { name: "P1C1", value: person.contact1.permissions.includes("Lives With") },
      { name: "P1C2", value: person.contact1.permissions.includes("Gets Mailings") },
      { name: "P1C3", value: person.contact1.permissions.includes("Emergency") },
      { name: "P1C4", value: person.contact1.permissions.includes("Permission to Pickup") }
    );
}

  if (person.contact2Exists && person.contact2.permissions) {
    textMappings.push(
      { name: "P2Name", value: `${person.contact2.firstName} ${person.contact2.lastName}` },
      { name: "P2R", value: person.contact2.relationship },
      { name: "P2P1", value: person.contact2.primaryPhone },  // Preferred goes first
      { name: "P2P2", value: [person.contact2.cellPhone, person.contact2.homePhone, person.contact2.workPhone].find(p => p !== person.contact2.primaryPhone) || "" },
      { name: "P2E", value: person.contact2.email },
      { name: "P2CL", value: person.contact2.lang },
      { name: "P2Fname", value: person.contact2.firstName },
      { name: "P2Lname", value: person.contact2.lastName },
      { name: "P2Add", value: person.contact2.address}
    );

    checkboxMappings.push(
      { name: "P2P1C", value: person.contact2.prefer === "Cell"},
    { name: "P2P1H", value: person.contact2.prefer === "Home"},
    { name: "P2P1W", value: person.contact2.prefer === "Work"},
    { name: "P2T1", value : person.contact2.trans},
    { name: "P2T2", value : !person.contact2.trans}
    );
  }

  if (person.contact2.permissions && person.contact2.permissions.length > 0) {
    checkboxMappings.push(
      { name: "P2C1", value: person.contact2.permissions.includes("Lives With") },
      { name: "P2C2", value: person.contact2.permissions.includes("Gets Mailings") },
      { name: "P2C3", value: person.contact2.permissions.includes("Emergency") },
      { name: "P2C4", value: person.contact2.permissions.includes("Permission to Pickup") }
    );
}

  // Combine all mappings
  const replaceArray = [];

  textMappings.forEach(mapping => {
    if (mapping.name) {
      replaceArray.push({ name: mapping.name, value: mapping.value });
    }
  });

  checkboxMappings.forEach(mapping => {
    if (mapping.name) {
      replaceArray.push({ name: mapping.name, value: mapping.value });
    }
  });

  Logger.log(JSON.stringify(replaceArray, null, 2));
  return replaceArray;
}
