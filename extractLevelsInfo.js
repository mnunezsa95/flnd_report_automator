function extractLevelsInfo() {
  const courseLevelsFileID = getTableFile(); 

  if (!courseLevelsFileID) {
    SpreadsheetApp.getUi().alert('No Course Levels file ID found.');
    return { numeracyData: [], literacyData: [] };
  }

  const file = DriveApp.getFileById(courseLevelsFileID);
  const spreadsheet = SpreadsheetApp.open(file);
  const sheet = spreadsheet.getSheetByName('Courses (Leveled)');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('No sheet named "Courses (Leveled)" found in the Course Levels file.');
    return { numeracyData: [], literacyData: [] };
  }

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(4, 1, lastRow - 3, 3); 
  const data = range.getValues();

  const numeracyData = [];
  const literacyData = [];
  const grades = new Set(); // To track unique grades

  // Process data and track grades
  data.forEach(row => {
    const grade = row[0].toString().trim(); // Trim whitespace and convert to string
    let subject = row[1].toString().trim(); // Trim whitespace and convert to string
    subject = subject.replace(/\s+/g, ' '); // Replace multiple spaces with a single space

    const formattedRow = {
      Grade: grade,
      Subject: subject,
      Level: row[2].toString().trim()
    };

    grades.add(grade); // Add grade to the set

    if (subject === 'Numeracy') {
      numeracyData.push(formattedRow);
    } else if (
      subject === 'Literacy Revision 1 & 2' ||
      subject === 'Literacy [1 & 2]'
    ) {
      literacyData.push(formattedRow);
    }
  });

  // Convert grades set to sorted array
  const sortedGrades = Array.from(grades).sort((a, b) => {
    const numA = parseInt(a.substring(1), 10);
    const numB = parseInt(b.substring(1), 10);
    return numA - numB;
  });

  // Determine missing grades
  const missingGrades = [];
  for (let i = 0; i < sortedGrades.length - 1; i++) {
    const currentGradeNum = parseInt(sortedGrades[i].substring(1), 10);
    const nextGradeNum = parseInt(sortedGrades[i + 1].substring(1), 10);

    // Check for missing grades between current and next
    for (let j = currentGradeNum + 1; j < nextGradeNum; j++) {
      const missingGrade = `G${j}`;
      missingGrades.push(missingGrade);
    }
  }

  // Insert missing grades into literacyData or numeracyData
  missingGrades.forEach(grade => {
    // Default values can be adjusted as needed
    const placeholderRow = {
      Grade: grade,
      Subject: 'n/a',
      Level: 'n/a'
    };
    
    if (grade.startsWith('G')) { // Example condition to add to literacyData
      literacyData.push(placeholderRow);
    } else {
      numeracyData.push(placeholderRow);
    }
  });

  // Function to sort data by Grade
  const sortByGrade = (a, b) => {
    const numA = parseInt(a.Grade.substring(1), 10);
    const numB = parseInt(b.Grade.substring(1), 10);
    return numA - numB;
  };

  // Sort both arrays by Grade
  numeracyData.sort(sortByGrade);
  literacyData.sort(sortByGrade);

  // Format the results as strings
  const numeracyStr = numeracyData.map(row => `Grade: ${row.Grade}, Subject: ${row.Subject}, Level: ${row.Level}`).join('\n');
  const literacyStr = literacyData.map(row => `Grade: ${row.Grade}, Subject: ${row.Subject}, Level: ${row.Level}`).join('\n');

  // Display results in pop-up messages
  SpreadsheetApp.getUi().alert('Numeracy Data:\n' + numeracyStr);
  SpreadsheetApp.getUi().alert('Literacy Data:\n' + literacyStr);

  return { numeracyData, literacyData };
}
