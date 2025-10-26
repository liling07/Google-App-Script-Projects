//july 2025
//credits to STEPHEN (GOAT) for building most of it last year, this year main changes I did was just how the form was filled in
//instead of pulling responses based on array, pulled directly based on question (in case things moved around / got added) 
//also changed how it filled the pdf so that the final emailed version was an editable pdf & ppl didn't have to resubmit the form for minor edits!
//note if you do use this, should probably build a trigger!

const folderId = "{redacted}";
const templateFileId = '{redacted}'; // PDF template file ID

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
        body: `Your registration file is ready here: ${fileUrl}. Please print out the forms, complete them in their entirety, and return them on your designated registration day.` ,
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


/* Stephen's code 
const folderId = "1eEPGMzx--cHeoUOr6z3NMhN4U4tt8vPv";
const templateFileId = '16aRk8imNTq9b4UolueJd5Htr61_bZNN8'; // PDF template file ID
// const folderId = "1eEPGMzx--cHeoUOr6z3NMhN4U4tt8vPv"; const fileId = '16aRk8imNTq9b4UolueJd5Htr61_bZNN8';
/**
 * Main trigger for form submissions.
 * Builds the person object and generates & sends the personalized PDF.
 
function formHandler(e) {
  const person = buildPersonFromResponse(e.response);
  replaceAndSend(person);
}

/**
 * Generates the PDF, saves it to Drive, and emails it to the respondent.
 
function replaceAndSend(person) {
  const pdfBlob = DriveApp.getFileById(templateFileId).getBlob();
  const pdfForm = new PdfForm();
  const replaceArray = buildReplaceArray(person);

  pdfForm.setValues(pdfBlob, replaceArray)
    .then(updatedBlob => {
      const fileName = `${person.firstName} ${person.lastName}'s Registration 2025-26.pdf`;

      
      DriveApp.getFolderById(folderId)
        .createFile(updatedBlob)
        .setName(fileName);

      
      const namedBlob = updatedBlob.setName(fileName);

      
      MailApp.sendEmail({
        to: person.contact1.email,
        subject: "Completed Payton Registration Documents",
        body: "Please print out the forms, complete them in their entirety, and return them before May 9th.",
        name: "Walter Payton College Prep",
        attachments: [ namedBlob ]
      });
    })
    .catch(error => {
      Logger.log('Failed to generate or send PDF: ' + error);
      throw new Error('Failed to generate or send PDF: ' + error);
    });
}
/**
 * Parses the FormResponse into a Person object for easy data handling.
 
function buildPersonFromResponse(formResponse) {
  const form = FormApp.getActiveForm();
  const items = form.getItems();
  const responses = formResponse.getItemResponses();
  const responseMap = new Map();

  responses.forEach(itemResponse => {
    responseMap.set(itemResponse.getItem().getId(), itemResponse.getResponse());
  });

  const responseArray = items
    .map(item => {
      const type = item.getType();
      if (type === FormApp.ItemType.PAGE_BREAK || type === FormApp.ItemType.SECTION_HEADER) {
        return null;
      }
      return responseMap.get(item.getId()) || "";
    })
    .filter(value => value !== null);

  return new Person(responseArray);
}

/**
 * Represents a form respondent and their data.
 
function Person(responseArray) {
  let i = 0;
  this.firstName = responseArray[i++];
  this.lastName = responseArray[i++];
  this.middleName = responseArray[i++];
  this.dob = formatDateString(responseArray[i++]);
  this.gender = responseArray[i++];
  this.address = responseArray[i++];
  this.city = responseArray[i++];
  this.state = responseArray[i++];
  this.zip = responseArray[i++];
  this.email = responseArray[i++];
  this.almaMater = responseArray[i++];
  this.siblingsAtPayton = responseArray[i++];
  this.grade = responseArray[i++];
  this.studentId = responseArray[i++];
  this.confidential1 = responseArray[i++];
  this.confidential2 = responseArray[i++] === "Yes";

  this.contact1 = createContact(responseArray, i);
  i += 19;

  this.contact2Exists = responseArray[i++] === "Yes";
  this.contact2 = this.contact2Exists ? createContact(responseArray, i) : {};
  i += this.contact2Exists ? 19 : 0;

  this.emergencyContact = {
    name: responseArray[i++],
    address: responseArray[i++],
    phone: responseArray[i++],
    relationship: responseArray[i++]
  };

  this.doctor = {
    fullName: responseArray[i++],
    phone: responseArray[i++],
    address: responseArray[i++]
  };

  this.studentHealthInsurance = responseArray[i++];
  this.medicalId = responseArray[i++];
  this.allKidsApplicationInterest = responseArray[i++] === "Yes";
  this.parentInArmedForces = responseArray[i++] === "Yes";
  this.parentIsDeployed = responseArray[i++] === "Yes";
  this.paytonFOPAuthorization = responseArray[i++] === "Yes";
  this.paytonFOPContactListAuthorization = responseArray[i++] === "Yes";

  this.calculatorPurchase = "";
  this.studentVentraCardPurchase = "";
}

/**
 * Helper to build a contact object from the response array.
 
function createContact(array, index) {
  return {
    firstName: array[index++],
    lastName: array[index++],
    email: array[index++],
    relationship: array[index++],
    permissions: array[index++],
    address: array[index++],
    city: array[index++],
    state: array[index++],
    zip: array[index++],
    homePhone: array[index++],
    cellPhone: array[index++],
    employerName: array[index++],
    employerAddress: array[index++],
    workPhone: array[index++],
    preferredPhone: array[index++],
    language: array[index++],
    volunteerInterest: array[index++],
    volunteerNoncommittalInterest: array[index++],
    volunteerCategory: array[index++]
  };
}

/**
 * Formats a date string (YYYY-MM-DD) into MM/dd/yyyy.
 
function formatDateString(dateString) {
  const date = new Date(dateString);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

/**
 * Builds the replace array for PdfForm.setValues(), including text and checkbox mappings.
 
function buildReplaceArray(person) {
  const textMappings = [
    { keys: [15, 66, 156, 163, 170, 185, 45, 105, 114, 126, 175], value: person.firstName },
    { keys: [14, 65, 155, 162, 169, 184, 41, 104, 113, 125, 154], value: person.lastName },
    { keys: [10, 51, 2, 99, 102, 149], value: `${person.firstName} ${person.middleName} ${person.lastName}` },
    { keys: [16, 67, 157, 164, 171, 186, 46, 106, 115, 127, 176], value: person.middleName },
    { keys: [], value: person.middleName.slice(0, 1) },
    { keys: [19, 172, 116, 128], value: person.gender },
    { keys: [18, 72, 174, 187, 47, 117, 129, 148, 177], value: person.dob },
    { keys: [12, 64, 159, 166, 191, 98, 100, 108, 119, 131, 150], value: person.studentId },
    { keys: [52, 190, 97, 103, 120, 132], value: person.grade },
    { keys: [73, 161, 168, 121, 133], value: person.advisory },
    { keys: [32, 53, 68, 109, 143], value: person.address },
    { keys: [33, 54, 69, 144], value: person.city },
    { keys: [34, 55, 70, 146], value: person.state },
    { keys: [35, 56, 71, 110, 147], value: person.zip },
    { keys: [], value: "" },
    { keys: [], value: person.county },
    { keys: [111], value: person.contact1.firstName },
    { keys: [112], value: person.contact1.lastName },
    { keys: [57, 75, 188, 1, 48, 122, 134, 178], value: `${person.contact1.firstName} ${person.contact1.lastName}` },
    { keys: [58, 137, 123, 136, 180], value: person.contact1.cellPhone },
    { keys: [40, 74, 138], value: person.contact1.homePhone },
    { keys: [139], value: person.contact1.workPhone },
    { keys: [59, 81, 124, 135, 179], value: person.contact1.email },
    { keys: [76], value: person.contact1.relationship },
    { keys: [], value: person.contact1.address },
    { keys: [83], value: person.contact1.language },
    { keys: [85], value: person.emergencyContact.name },
    { keys: [88], value: person.emergencyContact.address },
    { keys: [87], value: person.emergencyContact.phone },
    { keys: [86], value: person.emergencyContact.relationship },
    { keys: [], value: formatDateString(new Date().toISOString().slice(0, 10)) },
    { keys: [11, 63, 158, 165, 173, 189, 96, 107, 118, 130, 151, 152], value: "Walter Payton College Prep" },
    { keys: [], value: "60610" },
    { keys: [153], value: "1034 N. Wells St" },
    { keys: [160, 167], value: "15" },
    { keys: [95, 145], value: person.medicalId },
    { keys: [89], value: person.doctor.fullName },
    { keys: [90], value: person.doctor.phone },
    { keys: [91], value: person.doctor.address }
  ];

  const checkboxMappings = [
    { keys: [96], value: person.confidential1 === "in a car/park/other public place" },
    { keys: [97], value: person.confidential1 === "doubled-up" },
    { keys: [98], value: person.confidential1 === "in a hotel/motel" },
    { keys: [99], value: person.confidential1 === "in a shelter" },
    { keys: [100], value: person.confidential1 === "in transitional housing" },
    { keys: [101], value: person.confidential2 },
    { keys: [102], value: !person.confidential2 },
    { keys: [107], value: person.contact1.permissions.includes("Lives With") },
    { keys: [108], value: person.contact1.permissions.includes("Gets Mailings") },
    { keys: [109], value: person.contact1.permissions.includes("Emergency") },
    { keys: [110], value: person.contact1.permissions.includes("Permission to Pickup") },
    { keys: [145], value: person.studentHealthInsurance === "Illinois Medical Card/All Kids" },
    { keys: [146], value: person.studentHealthInsurance === "No Insurance" },
    { keys: [147], value: person.studentHealthInsurance === "Private/Employer Health Insurance" },
    { keys: [150], value: person.parentInArmedForces },
    { keys: [151], value: !person.parentInArmedForces },
    { keys: [152], value: person.parentIsDeployed },
    { keys: [154], value: !person.parentIsDeployed }
  ];

  // Include contact2 mappings if a second contact exists
  if (person.contact2Exists) {
    textMappings.push(
      { keys: [60, 77], value: `${person.contact2.firstName} ${person.contact2.lastName}` },
      { keys: [78], value: person.contact2.relationship },
      { keys: [61, 140], value: person.contact2.cellPhone },
      { keys: [141], value: person.contact2.homePhone },
      { keys: [142], value: person.contact2.workPhone },
      { keys: [62, 82], value: person.contact2.email },
      { keys: [84], value: person.contact2.language }
    );
    checkboxMappings.push(
      { keys: [111], value: person.contact2.permissions.includes("Lives With") },
      { keys: [112], value: person.contact2.permissions.includes("Gets Mailings") },
      { keys: [113], value: person.contact2.permissions.includes("Emergency") },
      { keys: [114], value: person.contact2.permissions.includes("Permission to Pickup") }
    );
  }

  // Flatten mappings into replace array
  const replaceArray = [];
  textMappings.forEach(mapping => {
    mapping.keys.forEach(key => {
      replaceArray.push({ name: `Text${key}`, value: mapping.value });
    });
  });
  checkboxMappings.forEach(mapping => {
    mapping.keys.forEach(key => {
      replaceArray.push({ name: `Check Box${key}`, value: mapping.value });
    });
  });

  return replaceArray;
}
*/
