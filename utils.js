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

function applyPercentageColor(cell, value, [LOW_COLOR, MEDIUM_COLOR, HIGH_COLOR]) {
  const percentageValue = parseFloat(value.replace("%", ""));
  const percentageSign = value.includes("%") ? "%" : ""; // Get the percentage sign

  // Set the main text and the percentage sign
  const percentageText = value.replace("%", ""); // Text without the %

  if (!isNaN(percentageValue)) {
    // Determine color based on percentage value
    let color;
    if (percentageValue <= 39) {
      color = LOW_COLOR;
    } else if (percentageValue <= 69) {
      color = MEDIUM_COLOR;
    } else if (percentageValue <= 100) {
      color = HIGH_COLOR;
    }

    // Clear existing content and set new text
    cell.setText(percentageText + percentageSign);

    // Get the text object of the cell
    const text = cell.getChild(0).asText();

    // Apply color to the percentage text
    text.setForegroundColor(0, percentageText.length - 1, color); // Color the percentage value

    // Apply black color to the % sign
    text.setForegroundColor(percentageText.length, percentageText.length, "#000000"); // Color for the %
  } else {
    // Reset to default if value is not a valid percentage
    cell.setText(value); // Set the original value if invalid
    cell.setForegroundColor("#000000"); // Default color (black)
  }
}
