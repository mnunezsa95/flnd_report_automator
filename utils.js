function generateUniqueID() {
  const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let uniqueID = "";
  for (let i = 0; i < 8; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    uniqueID += characters.charAt(randomIndex);
  }
  return uniqueID;
}

function resetFields(sheet, row, locs) {
  for (let i = 0; i < locs.length; i++) {
    sheet.getRange(row, locs[i]).setValue("");
  }
}

const sendEmail = (recipientList, subjectLine, message) => {
  const recipients = recipientList.length > 1 ? recipientList.join(",") : "marlon.nunez@newglobe.education";
  const subject = subjectLine;
  const body = message;

  GmailApp.sendEmail(recipients, subject, body);
};

function generateProgramAbbreviation(programName) {
  let programAbbreviation;
  if (programName === "BayelsaPRIME") {
    programAbbreviation = "BA";
  } else if (programName === "Bridge Andhra Pradesh") {
    programAbbreviation = "AP";
  } else if (programName === "Bridge Kenya") {
    programAbbreviation = "KE";
  } else if (programName === "Bridge Liberia") {
    programAbbreviation = "LR";
  } else if (programName === "Bridge Nigeria") {
    programAbbreviation = "NG";
  } else if (programName === "Bridge Uganda") {
    programAbbreviation = "UG";
  } else if (programName === "EdoBEST") {
    programAbbreviation = "EDO";
  } else if (programName === "EdoBEST JSS") {
    programAbbreviation = "EDO JSS";
  } else if (programName === "EKOEXCEL") {
    programAbbreviation = "EKO";
  } else if (programName === "ESPOIR République Centrafricaine") {
    programAbbreviation = "RCA";
  } else if (programName === "KwaraLEARN") {
    programAbbreviation = "KW";
  } else if (programName === "RwandaEQUIP") {
    programAbbreviation = "RW";
  } else if (programName === "STAR Education") {
    programAbbreviation = "MN";
  }

  return programAbbreviation;
}

function convertGradeToShort(grade, mapping) {
  const regex = /(Primary|Class|Standard|Grade) (\d+)/;
  const match = grade.match(regex);

  if (match) {
    const prefix = mapping[match[1]];
    const number = match[2];
    return prefix + number;
  }

  return grade;
}
