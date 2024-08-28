function extractLevelsInfo() {
  const courseLevelsFileID = getTableFile();
  const constants = getConstants();
  const programAbbreviation = generateProgramAbbreviation(constants.programName);

  if (!courseLevelsFileID) {
    SpreadsheetApp.getUi().alert("No Course Levels file ID found.");
    return { numeracyData: [], literacyData: [] };
  }

  const file = DriveApp.getFileById(courseLevelsFileID);
  const spreadsheet = SpreadsheetApp.open(file);
  const sheet = spreadsheet.getSheetByName(`${programAbbreviation} - Courses (Leveled)`);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('No sheet named "Courses (Leveled)" found in the Course Levels file.');
    return { numeracyData: [], literacyData: [] };
  }

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(4, 1, lastRow - 3, 3);
  const data = range.getValues();

  let numeracyData = { Primary: [], Progressive: [] };
  let literacyData = { Primary: [], Progressive: [] };
  const numeracyGrades = new Set();
  const literacyGrades = new Set();

  data.forEach((row) => {
    let grade = row[0].toString().trim();
    let subject = row[1].toString().trim();
    let level = row[2].toString().trim();
    grade = grade.replace(/\s+/g, " ");
    subject = subject.replace(/\s+/g, " ");
    level = level.replace(/\s+/g, " ");

    const formattedRow = {
      Grade: grade,
      Subject: subject,
      Level: level,
    };

    // Check if grade ends with two digits
    const endsWithTwoDigits = /\d{2}$/.test(grade);
    // Determine category based on grade length and if it ends with two digits
    const category = grade.length === 3 && endsWithTwoDigits ? "Progressive" : "Primary";

    if (
      subject === "Numeracy" ||
      subject === "Mathematics 1 & 2" ||
      subject === "Supplementary Numeracy" ||
      subject === "Accelerated Maths" ||
      subject === "Supplementary Maths"
    ) {
      numeracyGrades.add(grade);
      numeracyData[category].push(formattedRow);
    } else if (
      subject === "Literacy Revision 1 & 2" ||
      subject === "Supplemental Language" ||
      subject === "English Literacy Revision 1 & 2" ||
      subject === "Literacy 1 & 2" ||
      subject === "English Studies - Language" ||
      subject === "English Studies - Reading 1 & 2" ||
      subject === "Supplementary English 1 & 2"
    ) {
      literacyGrades.add(grade);
      literacyData[category].push(formattedRow);
    }
  });

  // Function to determine missing grades and add them to data
  function addMissingGrades(gradeSet, data, category) {
    // Extract single letter grades and sort them numerically
    const gradesArray = Array.from(gradeSet);
    const letterGrades = {};

    gradesArray.forEach((grade) => {
      const letter = grade.charAt(0);
      const number = parseInt(grade.substring(1), 10);

      if (!letterGrades[letter]) {
        letterGrades[letter] = new Set();
      }

      if (!isNaN(number) && number < 10) {
        // Only consider single-digit numbers
        letterGrades[letter].add(number);
      }
    });

    const missingGrades = [];

    // For each letter grade, find missing grades in the range of single-digit numbers
    Object.keys(letterGrades).forEach((letter) => {
      const sortedNumbers = Array.from(letterGrades[letter]).sort((a, b) => a - b);
      const lowestGrade = sortedNumbers[0];
      const highestGrade = sortedNumbers[sortedNumbers.length - 1];

      for (let i = lowestGrade; i <= highestGrade; i++) {
        const grade = `${letter}${i}`;
        if (!gradeSet.has(grade)) {
          missingGrades.push(grade);
        }
      }
    });

    missingGrades.forEach((grade) => {
      const placeholderRow = {
        Grade: grade,
        Subject: "-",
        Level: "-",
      };
      data[category].push(placeholderRow);
    });
  }

  // Add missing grades separately for Literacy and Numeracy
  addMissingGrades(numeracyGrades, numeracyData, "Primary");
  addMissingGrades(literacyGrades, literacyData, "Primary");

  const sortByGrade = (a, b) => {
    const numA = parseInt(a.Grade.substring(1), 10);
    const numB = parseInt(b.Grade.substring(1), 10);
    return numA - numB;
  };

  const sortData = (data) => {
    const primary = data.Primary.sort(sortByGrade);
    const progressive = data.Progressive.sort(sortByGrade);

    // Further sort to ensure same grades appear together regardless of subject
    return {
      Primary: primary.sort((a, b) => a.Grade.localeCompare(b.Grade)),
      Progressive: progressive.sort((a, b) => a.Grade.localeCompare(b.Grade)),
    };
  };

  numeracyData = sortData(numeracyData);
  literacyData = sortData(literacyData);

  const formatData = (data) => {
    return data.map((row) => `{ Grade: ${row.Grade}, Subject: ${row.Subject}, Level: ${row.Level} }`).join("\n");
  };

  const numeracyStrPrimary = formatData(numeracyData.Primary);
  const numeracyStrProgressive = formatData(numeracyData.Progressive);
  const literacyStrPrimary = formatData(literacyData.Primary);
  const literacyStrProgressive = formatData(literacyData.Progressive);

  SpreadsheetApp.getUi().alert(
    `Literacy Data\n\nPrimary\n${literacyStrPrimary}\n\nProgressive\n${
      literacyStrProgressive ? literacyStrProgressive : "{ No Progressive Data Found }"
    }`
  );
  SpreadsheetApp.getUi().alert(
    `Numeracy Data\n\nPrimary\n${numeracyStrPrimary}\n\nProgressive\n${
      numeracyStrProgressive ? numeracyStrProgressive : "{ No Progressive Data Found }"
    }`
  );

  return { numeracyData, literacyData };
}
