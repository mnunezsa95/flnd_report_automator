function extractLevelScores() {
  const constants = getConstants();
  const scoresURL = constants.rmeScoresURL;

  // Get the sheet by URL
  const sheet = SpreadsheetApp.openByUrl(scoresURL).getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // Initialize the final object structure for literacy and numeracy
  let numeracy = { Primary: {}, Progressive: {} };
  let literacy = { Primary: {}, Progressive: {} };

  // Mapping for conversions
  const gradeMapping = {
    Primary: "P",
    Class: "C",
    Standard: "S",
    Grade: "G",
  };

  // Loop through the rows of data (starting from row 2, since row 1 is the header)
  for (let i = 1; i < data.length; i++) {
    const grade = data[i][0]; // Grade (e.g., "Primary 1")
    const type = data[i][1]; // Type (e.g., "Primary", "Progressive")
    const level = data[i][2]; // Level (e.g., A, B, C)
    const literacyScore = data[i][3]; // Literacy score
    const numeracyScore = data[i][4]; // Numeracy score

    // Convert grade name to shorthand using the provided function
    const shortGrade = convertGradeToShort(grade, gradeMapping);

    // Initialize the grade entry in Literacy and Numeracy based on type if not present
    if (!literacy[type]) {
      literacy[type] = {};
    }
    if (!numeracy[type]) {
      numeracy[type] = {};
    }
    if (!literacy[type][shortGrade]) {
      literacy[type][shortGrade] = {};
    }
    if (!numeracy[type][shortGrade]) {
      numeracy[type][shortGrade] = {};
    }

    // Add Literacy score if available
    if (literacyScore !== "") {
      literacy[type][shortGrade][level] = Number(literacyScore).toFixed(1);
    }

    // Add Numeracy score if available
    if (numeracyScore !== "") {
      numeracy[type][shortGrade][level] = Number(numeracyScore).toFixed(1);
    }
  }

  // Log the resulting objects
  console.log(JSON.stringify(numeracy, null, 2));

  // Return both literacy and numeracy objects
  return { literacy, numeracy };
}
