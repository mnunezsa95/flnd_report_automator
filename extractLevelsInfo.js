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

    const endsWithTwoDigits = /\d{2}$/.test(grade);
    const category = grade.length === 3 && endsWithTwoDigits ? "Progressive" : "Primary";

    if (
      subject === "Numeracy" ||
      subject === "Mathematics 1 & 2" ||
      subject === "Supplementary Numeracy" ||
      subject === "Accelerated Maths" ||
      subject === "Supplementary Maths" ||
      subject === "Preparatory Maths"
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
      subject === "Supplementary English 1 & 2" ||
      subject === "Preparatory English 1 & 2"
    ) {
      literacyGrades.add(grade);
      literacyData[category].push(formattedRow);
    }
  });

  // Modify the score extraction logic
  function getScores(levels, scoresData, typeOfProgram) {
    const scores = {};

    levels.forEach(({ Grade, Level }) => {
      const [levelStart, levelEnd] = Level.split("-").map((l) => l.replace(/[0-9]/g, "").trim());

      // Treat levels as per the specifications
      let levelKeys = [];
      levelKeys.push(levelStart);
      if (levelEnd) {
        levelKeys.push(levelEnd);
      }

      // Fetch scores for each key, prioritizing existing scores
      scores[Grade] = {
        avgOne: scoresData[typeOfProgram][Grade]?.[levelKeys[0]] || "", // First level score
        avgTwo:
          levelKeys[1] && levelKeys[1] !== levelKeys[0] ? scoresData[typeOfProgram][Grade]?.[levelKeys[1]] || "" : "",
      };

      // Replace any undefined values with empty strings
      scores[Grade].avgOne = scores[Grade].avgOne === undefined ? "" : scores[Grade].avgOne;
      scores[Grade].avgTwo = scores[Grade].avgTwo === undefined ? "" : scores[Grade].avgTwo;
    });

    return scores;
  }

  // Call to extractLevelScores() to get the scores
  const { numeracy, literacy } = extractLevelScores();

  numeracyData.Primary.forEach((row) => {
    const { Grade } = row;
    const scoreData = getScores([row], numeracy, "Primary");
    row.avgOne = scoreData[Grade].avgOne;
    row.avgTwo = scoreData[Grade].avgTwo;
  });

  literacyData.Primary.forEach((row) => {
    const { Grade } = row;
    const scoreData = getScores([row], literacy, "Primary");
    row.avgOne = scoreData[Grade].avgOne;
    row.avgTwo = scoreData[Grade].avgTwo;
  });

  numeracyData.Progressive.forEach((row) => {
    const { Grade } = row;
    const scoreData = getScores([row], numeracy, "Progressive");
    row.avgOne = scoreData[Grade].avgOne;
    row.avgTwo = scoreData[Grade].avgTwo;
  });

  literacyData.Progressive.forEach((row) => {
    const { Grade } = row;
    const scoreData = getScores([row], literacy, "Progressive");
    row.avgOne = scoreData[Grade].avgOne;
    row.avgTwo = scoreData[Grade].avgTwo;
  });

  const sortByGrade = (a, b) => {
    const numA = parseInt(a.Grade.substring(1), 10);
    const numB = parseInt(b.Grade.substring(1), 10);
    return numA - numB;
  };

  const sortData = (data) => {
    return {
      Primary: data.Primary.sort(sortByGrade),
      Progressive: data.Progressive.sort(sortByGrade),
    };
  };

  numeracyData = sortData(numeracyData);
  literacyData = sortData(literacyData);

  const formatData = (data) => {
    return data
      .map(
        (row) =>
          `[ Gr: ${row.Grade}, Sub: ${row.Subject}, Lvl: ${row.Level}, avg1: ${row.avgOne}, avg2: ${row.avgTwo} ]`
      )
      .join("\n");
  };

  const numeracyStrPrimary = formatData(numeracyData.Primary);
  const numeracyStrProgressive = formatData(numeracyData.Progressive);
  const literacyStrPrimary = formatData(literacyData.Primary);
  const literacyStrProgressive = formatData(literacyData.Progressive);

  SpreadsheetApp.getUi().alert(
    `Literacy Data\n\nPrimary\n${literacyStrPrimary}\n\nProgressive\n${
      literacyStrProgressive ? literacyStrProgressive : "[ No Progressive Data Found ]"
    }`
  );
  SpreadsheetApp.getUi().alert(
    `Numeracy Data\n\nPrimary\n${numeracyStrPrimary}\n\nProgressive\n${
      numeracyStrProgressive ? numeracyStrProgressive : "[ No Progressive Data Found ]"
    }`
  );

  return { numeracyData, literacyData };
}
