// --- Global Variables & Configuration ---
const SPREADSHEET_ID = '1aVmXdoThzsfjsEtV-vys3movt5FPc2QYXAWJlntPSh4'; // Replace with your Google Sheet ID
const PROMPTS_SHEET_NAME = 'Prompts'; // Sheet name where prompts are listed
const STUDENTS_SHEET_NAME = 'Students'; // Sheet for tracking student registrations and votes
const GAME_STATE_SHEET_NAME = 'GameState'; // Sheet for storing game state (optional, can also use PropertiesService)
const VOTING_DURATION_SECONDS = 20; // Define voting duration

let gameState = {
  status: 'setup', // 'setup', 'waiting', 'voting', 'round_over', 'game_over'
  currentRound: 0,
  currentMatchupIndex: 0,
  prompts: [], // Will be loaded from sheet: [{id: 1, text: "Prompt A"}, {id: 2, text: "Prompt B"}, ...]
  bracket: [], // Structure: [[{prompt1, prompt2, winner, votes1, votes2, voters: [], votingStartTime, votingEndTime}], [{...next round...}]]
  activeMatchup: null, // { promptA, promptB, votesA, votesB, codeA, codeB, voters: [], votingStartTime, votingEndTime }
  students: [] // List of registered students: [{firstName, lastName, nickname, sessionKey}]
};

// --- Nickname Generation ---
const ADJECTIVES = ['Swift', 'Clever', 'Brave', 'Mighty', 'Sneaky', 'Bold', 'Quick', 'Wise', 'Lucky', 'Fierce', 'Gentle', 'Sharp', 'Bright', 'Cool', 'Wild', 'Silent', 'Strong', 'Fast', 'Smart', 'Calm'];
const COLORS = ['Red', 'Blue', 'Green', 'Purple', 'Orange', 'Silver', 'Golden', 'Crimson', 'Azure', 'Emerald', 'Violet', 'Amber', 'Coral', 'Indigo', 'Teal', 'Pink', 'Yellow', 'Black', 'White', 'Gray'];
const ANIMALS = ['Tiger', 'Eagle', 'Wolf', 'Bear', 'Fox', 'Lion', 'Hawk', 'Shark', 'Panther', 'Dragon', 'Falcon', 'Leopard', 'Rhino', 'Cobra', 'Raven', 'Lynx', 'Jaguar', 'Owl', 'Viper', 'Phoenix'];

function generateNickname_() {
  const adjective = ADJECTIVES[Math.floor(Math.random() * ADJECTIVES.length)];
  const color = COLORS[Math.floor(Math.random() * COLORS.length)];
  const animal = ANIMALS[Math.floor(Math.random() * ANIMALS.length)];
  return `${adjective} ${color} ${animal}`;
}

function generateUniqueNickname_(existingNicknames) {
  let nickname;
  let attempts = 0;
  do {
    nickname = generateNickname_();
    attempts++;
    if (attempts > 100) {
      // Fallback with numbers if we can't generate unique after 100 attempts
      nickname = generateNickname_() + ' ' + Math.floor(Math.random() * 1000);
      break;
    }
  } while (existingNicknames.includes(nickname));
  return nickname;
}

// --- Game State Persistence Functions ---
function loadGameStateFromProperties_() {
  try {
    const savedStateString = PropertiesService.getScriptProperties().getProperty('gameState');
    if (savedStateString) {
      const loadedState = JSON.parse(savedStateString);
      gameState = {
        status: 'setup', 
        currentRound: 0, 
        currentMatchupIndex: 0, 
        prompts: [], 
        bracket: [], 
        activeMatchup: null,
        students: [],
        ...loadedState 
      };
      Logger.log('GameState loaded from PropertiesService.');
      if ((!gameState.prompts || gameState.prompts.length === 0) && gameState.status !== 'game_over') {
        Logger.log('Prompts empty or not present in loaded state, attempting to load from sheet.');
        loadPromptsFromSheet_();
      }
      if (!gameState.students) {
        gameState.students = [];
      }
    } else {
      Logger.log('No saved game state found in PropertiesService. Initializing and loading prompts.');
      gameState = {
        status: 'setup',
        currentRound: 0,
        currentMatchupIndex: 0,
        prompts: [],
        bracket: [],
        activeMatchup: null,
        students: []
      };
      loadPromptsFromSheet_();
    }
  } catch (e) {
    Logger.log(`Error loading game state from PropertiesService: ${e.toString()}. Resetting to default and loading prompts.`);
    gameState = {
        status: 'setup',
        currentRound: 0,
        currentMatchupIndex: 0,
        prompts: [],
        bracket: [],
        activeMatchup: null,
        students: []
      };
    loadPromptsFromSheet_();
    PropertiesService.getScriptProperties().deleteProperty('gameState');
    Logger.log('Corrupted gameState in PropertiesService cleared.');
  }
}

function saveGameStateToProperties_() {
  try {
    PropertiesService.getScriptProperties().setProperty('gameState', JSON.stringify(gameState));
    Logger.log('GameState saved to PropertiesService.');
  } catch (e) {
    Logger.log(`Error saving game state to PropertiesService: ${e.toString()}`);
    Logger.log('Critical Error: Could not save game state. Please check logs for teacher action if issue persists.');
  }
}

// --- Students Sheet Management ---
function createStudentsSheet_() {
  try {
    let ss;
    if (SPREADSHEET_ID && SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID') {
        ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    if (!ss) {
        Logger.log('Could not get spreadsheet instance for creating Students sheet.');
        return null;
    }

    let sheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
    if (sheet) {
        Logger.log('Students sheet already exists.');
        return sheet;
    }

    sheet = ss.insertSheet(STUDENTS_SHEET_NAME);
    // Set up headers
    const headers = ['First Name', 'Last Name', 'Nickname'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    Logger.log('Students sheet created successfully.');
    return sheet;
  } catch (e) {
    Logger.log(`Error creating Students sheet: ${e.toString()}`);
    return null;
  }
}

function addStudentToSheet_(student) {
  try {
    const sheet = createStudentsSheet_();
    if (!sheet) {
      Logger.log('Could not access Students sheet for adding student.');
      return false;
    }

    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Add student data
    sheet.getRange(newRow, 1, 1, 3).setValues([[
      student.firstName,
      student.lastName,
      student.nickname
    ]]);
    
    Logger.log(`Added student ${student.nickname} to sheet at row ${newRow}.`);
    return true;
  } catch (e) {
    Logger.log(`Error adding student to sheet: ${e.toString()}`);
    return false;
  }
}

function addVoteColumnHeader_(roundNumber) {
  try {
    const sheet = createStudentsSheet_();
    if (!sheet) return false;

    const headerRow = 1;
    const voteColumnIndex = 4 + roundNumber - 1; // Columns D, E, F, etc.
    const headerName = `Round ${roundNumber} Vote`;
    
    // Check if header already exists
    const currentHeader = sheet.getRange(headerRow, voteColumnIndex).getValue();
    if (currentHeader === headerName) {
      return true; // Already exists
    }
    
    sheet.getRange(headerRow, voteColumnIndex).setValue(headerName);
    sheet.getRange(headerRow, voteColumnIndex).setFontWeight('bold');
    
    Logger.log(`Added vote column header for Round ${roundNumber}.`);
    return true;
  } catch (e) {
    Logger.log(`Error adding vote column header: ${e.toString()}`);
    return false;
  }
}

function recordVoteInSheet_(nickname, roundNumber, votedFor) {
  try {
    const sheet = createStudentsSheet_();
    if (!sheet) return false;

    // Find the student's row
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false; // No data rows
    
    const nicknameColumn = 3; // Column C
    const nicknameRange = sheet.getRange(2, nicknameColumn, lastRow - 1, 1);
    const nicknames = nicknameRange.getValues().flat();
    
    const studentRowIndex = nicknames.indexOf(nickname);
    if (studentRowIndex === -1) {
      Logger.log(`Student with nickname ${nickname} not found in sheet.`);
      return false;
    }
    
    const studentRow = studentRowIndex + 2; // Add 2 because we started from row 2 and arrays are 0-indexed
    const voteColumnIndex = 4 + roundNumber - 1; // Columns D, E, F, etc.
    
    // Ensure the vote column header exists
    addVoteColumnHeader_(roundNumber);
    
    // Record the vote
    sheet.getRange(studentRow, voteColumnIndex).setValue(votedFor);
    
    Logger.log(`Recorded vote for ${nickname} in Round ${roundNumber}: ${votedFor}`);
    return true;
  } catch (e) {
    Logger.log(`Error recording vote in sheet: ${e.toString()}`);
    return false;
  }
}

// --- Custom Menu ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bracket Battle Game')
    .addItem('Start Game (Show Waiting Room)', 'startGameMenuItem')
    .addItem('Launch Next Round/Prompts', 'launchNextRoundMenuItem')
    .addItem('Reset Game', 'resetGameMenuItem')
    .addToUi();
}

// --- Web App Endpoint ---
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('Prompt Bracket Battle')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// --- Menu Item Functions ---
function startGameMenuItem() {
  loadGameStateFromProperties_();

  if (!gameState.prompts || gameState.prompts.length === 0) {
    Logger.log('startGameMenuItem: Prompts are empty, attempting to load from sheet.');
    loadPromptsFromSheet_();
  }

  if (gameState.prompts.length < 2) {
    SpreadsheetApp.getUi().alert('Not enough prompts in the sheet to start a game. Please add at least 2 prompts to Column A of the "Prompts" sheet.');
    return;
  }

  gameState.status = 'waiting';
  gameState.currentRound = 0;
  gameState.currentMatchupIndex = 0;
  gameState.bracket = [];
  gameState.activeMatchup = null;
  gameState.students = []; // Reset students for new game
  setupInitialBracket_(); 
  createStudentsSheet_(); // Ensure Students sheet exists

  SpreadsheetApp.getUi().alert('Game set to "Waiting Room". Students can now register and join the web app.');
  saveGameStateToProperties_();
}

function launchNextRoundMenuItem() {
  loadGameStateFromProperties_();

  if (gameState.status === 'setup') {
      SpreadsheetApp.getUi().alert('Please "Start Game" first to initialize prompts and waiting room.');
      return;
  }

  if (gameState.status === 'voting' && gameState.activeMatchup) {
    Logger.log('Teacher manually advancing from voting state.');
    determineWinnerAndAdvance_();
  }

  if (gameState.status === 'game_over') {
    SpreadsheetApp.getUi().alert('Game is over. Please reset to start a new game.');
    return;
  }

  const success = prepareNextMatchup_(); 
  if (success) {
    gameState.status = 'voting';
    SpreadsheetApp.getUi().alert(`Round ${gameState.currentRound + 1}, Matchup ${gameState.currentMatchupIndex + 1} launched!`);
  } else {
    if (gameState.bracket[gameState.currentRound] && gameState.bracket[gameState.currentRound].every(m => m.winner)) {
        if (gameState.bracket[gameState.currentRound].length === 1 && gameState.bracket[gameState.currentRound][0].winner) {
            gameState.status = 'game_over';
            gameState.activeMatchup = null;
            SpreadsheetApp.getUi().alert('Game Over! Final winner determined.');
        } else {
            gameState.currentRound++;
            gameState.currentMatchupIndex = 0; 
            setupNextRoundBracket_(); 

            if (gameState.status === 'game_over') {
                SpreadsheetApp.getUi().alert('Game Over! Determined after setting up next round.');
            } else {
                const nextRoundSuccess = prepareNextMatchup_();
                if (nextRoundSuccess) {
                    gameState.status = 'voting';
                    SpreadsheetApp.getUi().alert(`Advanced to Round ${gameState.currentRound + 1}. Next matchup launched!`);
                } else {
                    if(gameState.status === 'game_over'){ 
                        SpreadsheetApp.getUi().alert('Game Over! No more matchups to prepare.');
                    } else {
                        SpreadsheetApp.getUi().alert('Could not prepare next matchup after advancing round. Game might be stuck or over. Check logs.');
                    }
                }
            }
        }
    } else {
      if (gameState.status !== 'game_over') { 
          SpreadsheetApp.getUi().alert('Could not prepare the next matchup. Current round may not be complete or an error occurred. Check logs.');
      }
    }
  }
  saveGameStateToProperties_();
}

function resetGameMenuItem() {
  gameState = {
    status: 'setup',
    currentRound: 0,
    currentMatchupIndex: 0,
    prompts: [],
    bracket: [],
    activeMatchup: null,
    students: []
  };
  loadPromptsFromSheet_();
  
  // Clear the Students sheet
  try {
    let ss;
    if (SPREADSHEET_ID && SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID') {
        ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    const sheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
    if (sheet) {
      sheet.clear();
      const headers = ['First Name', 'Last Name', 'Nickname'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      Logger.log('Students sheet cleared and headers reset.');
    }
  } catch (e) {
    Logger.log(`Error clearing Students sheet: ${e.toString()}`);
  }
  
  SpreadsheetApp.getUi().alert('Game has been reset. Prompts reloaded from sheet and Students sheet cleared.');
  saveGameStateToProperties_();
}

// --- Core Game Logic ---
function setupInitialBracket_() {
  if (!gameState.prompts || gameState.prompts.length === 0) {
      Logger.log('setupInitialBracket_: No prompts available. Attempting load.');
      loadPromptsFromSheet_();
      if (!gameState.prompts || gameState.prompts.length === 0) {
          Logger.log('Cannot set up bracket: No prompts found even after reload.');
          return; 
      }
  }
  if (gameState.prompts.length < 2) {
      Logger.log('Cannot set up bracket: Need at least 2 prompts.');
      return;
  }

  let shuffledPrompts = [...gameState.prompts].sort(() => 0.5 - Math.random());
  
  gameState.bracket = [];
  const firstRoundMatchups = [];
  for (let i = 0; i < shuffledPrompts.length; i += 2) {
    if (shuffledPrompts[i+1]) {
      firstRoundMatchups.push({
        promptA: shuffledPrompts[i],
        promptB: shuffledPrompts[i+1],
        votesA: 0, votesB: 0, winner: null,
        codeA: `A${Math.floor(i/2)+1}`, codeB: `B${Math.floor(i/2)+1}`,
        voters: [] 
      });
    } else { 
      firstRoundMatchups.push({
        promptA: shuffledPrompts[i],
        promptB: { id: 'BYE_ID', text: 'BYE (Auto-Win)'},
        votesA: 1, votesB: 0, winner: shuffledPrompts[i],
        codeA: `A${Math.floor(i/2)+1}`, codeB: 'BYE',
        voters: [] 
      });
    }
  }
  gameState.bracket.push(firstRoundMatchups);
  Logger.log(`Initial bracket setup with ${firstRoundMatchups.length} matchups for Round 1.`);
}

function setupNextRoundBracket_() {
    const previousRoundIndex = gameState.currentRound -1; 
    if (previousRoundIndex < 0 || !gameState.bracket[previousRoundIndex]) {
        Logger.log('Cannot setup next round: Invalid previous round index or bracket data missing.');
        gameState.status = 'error'; 
        return;
    }

    const previousRound = gameState.bracket[previousRoundIndex];
    if (!previousRound.every(m => m.winner)) {
        Logger.log('Cannot setup next round: Previous round not fully completed.');
        return;
    }

    const winners = previousRound.map(matchup => matchup.winner).filter(winner => winner && winner.id !== 'BYE_ID');
    
    if (winners.length === 0) {
        Logger.log('No valid winners from previous round. Game might have error.');
        gameState.status = 'error'; 
        return;
    }

    if (winners.length === 1) { 
        gameState.status = 'game_over';
        gameState.activeMatchup = null; 
        Logger.log('Game Over - Final winner determined in setupNextRoundBracket_ as only one winner advanced.');
        return; 
    }

    const nextRoundMatchups = [];
    for (let i = 0; i < winners.length; i += 2) {
        if (winners[i+1]) {
            nextRoundMatchups.push({
                promptA: winners[i], promptB: winners[i+1],
                votesA: 0, votesB: 0, winner: null,
                codeA: `A${Math.floor(i/2)+1}_R${gameState.currentRound+1}`, 
                codeB: `B${Math.floor(i/2)+1}_R${gameState.currentRound+1}`,
                voters: []
            });
        } else { 
             nextRoundMatchups.push({
                promptA: winners[i], promptB: { id: 'BYE_ID', text: 'BYE (Auto-Win)'},
                votesA: 1, votesB: 0, winner: winners[i],
                codeA: `A${Math.floor(i/2)+1}_R${gameState.currentRound+1}`, codeB: 'BYE',
                voters: []
            });
        }
    }

    if (nextRoundMatchups.length > 0) {
        gameState.bracket.push(nextRoundMatchups);
        Logger.log(`Setup Round ${gameState.currentRound + 1} with ${nextRoundMatchups.length} matchups.`);
    } else if (winners.length > 0) { 
        Logger.log('Had winners but could not form next round matchups. This might indicate an issue or game end.');
        if (winners.length === 1) gameState.status = 'game_over'; 
    } else {
        Logger.log('No matchups created for the next round. Previous round might have been the final.');
    }
}

function prepareNextMatchup_() {
  if (!gameState.bracket[gameState.currentRound] || gameState.bracket[gameState.currentRound].length === 0) {
    Logger.log(`No matchups defined for Round ${gameState.currentRound + 1}. Cannot prepare.`);
    return false;
  }

  const currentRoundMatchups = gameState.bracket[gameState.currentRound];
  let matchupToSet = null;
  let foundMatchupIndex = -1;

  for(let i = gameState.currentMatchupIndex; i < currentRoundMatchups.length; i++) {
      if (!currentRoundMatchups[i].winner) {
          matchupToSet = currentRoundMatchups[i];
          foundMatchupIndex = i;
          break;
      }
  }
  
  if (matchupToSet) {
    gameState.currentMatchupIndex = foundMatchupIndex;
    const now = Date.now();
    const votingEndTime = now + (VOTING_DURATION_SECONDS * 1000);

    gameState.activeMatchup = {
      promptA: matchupToSet.promptA,
      promptB: matchupToSet.promptB,
      votesA: matchupToSet.votesA || 0,
      votesB: matchupToSet.votesB || 0,
      codeA: matchupToSet.codeA,
      codeB: matchupToSet.codeB,
      voters: matchupToSet.voters || [], 
      round: gameState.currentRound,
      matchupIndexInRound: gameState.currentMatchupIndex,
      votingStartTime: now,
      votingEndTime: votingEndTime
    };
    
    if(gameState.bracket[gameState.currentRound] && gameState.bracket[gameState.currentRound][foundMatchupIndex]){
        gameState.bracket[gameState.currentRound][foundMatchupIndex].votingStartTime = now;
        gameState.bracket[gameState.currentRound][foundMatchupIndex].votingEndTime = votingEndTime;
    }

    Logger.log(`Prepared active matchup for R${gameState.currentRound + 1}, M${gameState.currentMatchupIndex + 1}: "${matchupToSet.promptA.text}" vs "${matchupToSet.promptB.text}". Voting ends: ${new Date(votingEndTime).toLocaleTimeString()}`);
    return true;
  } else {
    Logger.log(`No more pending matchups found in Round ${gameState.currentRound + 1} from index ${gameState.currentMatchupIndex}.`);
    gameState.activeMatchup = null;
    return false;
  }
}

function determineWinnerAndAdvance_() {
  if (!gameState.activeMatchup ) {
    Logger.log('determineWinnerAndAdvance_: No active matchup to determine winner.');
    return; 
  }
   if (gameState.status !== 'voting' && gameState.status !== 'round_over_pending_auto_advance') {
      Logger.log(`determineWinnerAndAdvance_ called with status ${gameState.status}. No action taken.`);
      return;
  }

  const { promptA, promptB, votesA, votesB, round, matchupIndexInRound } = gameState.activeMatchup;
  let winner = null;

  if (promptB.id === 'BYE_ID') {
      winner = promptA;
      Logger.log(`Matchup was a BYE for "${promptA.text}". Winner is "${winner.text}".`);
  } else if (votesA > votesB) {
    winner = promptA;
  } else if (votesB > votesA) {
    winner = promptB;
  } else { 
    winner = Math.random() < 0.5 ? promptA : promptB;
    Logger.log(`Tie occurred ("${promptA.text}" ${votesA} vs "${promptB.text}" ${votesB}). Winner by tie-breaker: "${winner.text}"`);
  }

  if (gameState.bracket[round] && gameState.bracket[round][matchupIndexInRound]) {
    gameState.bracket[round][matchupIndexInRound].winner = winner;
    gameState.bracket[round][matchupIndexInRound].votesA = votesA; 
    gameState.bracket[round][matchupIndexInRound].votesB = votesB;
    gameState.status = 'round_over';
    Logger.log(`Winner of R${round+1}, M${matchupIndexInRound+1} ("${promptA.text}" vs "${promptB.text}") is: "${winner.text}". Status set to round_over.`);
  } else {
    Logger.log(`Error in determineWinnerAndAdvance_: Could not find matchup in bracket at R${round}, M${matchupIndexInRound} to update winner.`);
  }
}

// --- Student Registration Functions ---
function registerStudent(registrationData) {
  loadGameStateFromProperties_();
  
  if (gameState.status !== 'waiting') {
    return { 
      success: false, 
      message: 'Student registration is only available when the game is in waiting room mode.' 
    };
  }

  const { firstName, lastName } = registrationData;
  
  if (!firstName || !lastName || firstName.trim() === '' || lastName.trim() === '') {
    return { 
      success: false, 
      message: 'First name and last name are required.' 
    };
  }

  const cleanFirstName = firstName.trim();
  const cleanLastName = lastName.trim();
  const sessionKey = Session.getTemporaryActiveUserKey();
  
  // Check if this session key is already registered
  const existingStudent = gameState.students.find(s => s.sessionKey === sessionKey);
  if (existingStudent) {
    return { 
      success: false, 
      message: 'This browser session is already registered. Please refresh the page if you need to re-register.' 
    };
  }
  
  // Generate unique nickname
  const existingNicknames = gameState.students.map(s => s.nickname);
  const nickname = generateUniqueNickname_(existingNicknames);
  
  const newStudent = {
    firstName: cleanFirstName,
    lastName: cleanLastName,
    nickname: nickname,
    sessionKey: sessionKey
  };
  
  // Add to game state
  gameState.students.push(newStudent);
  
  // Add to sheet
  const sheetSuccess = addStudentToSheet_(newStudent);
  if (!sheetSuccess) {
    Logger.log('Warning: Could not add student to sheet, but proceeding with registration.');
  }
  
  saveGameStateToProperties_();
  
  Logger.log(`Student registered: ${cleanFirstName} ${cleanLastName} as ${nickname}`);
  
  return {
    success: true,
    message: `Welcome, ${nickname}!`,
    student: {
      firstName: cleanFirstName,
      lastName: cleanLastName,
      nickname: nickname
    }
  };
}

function confirmStudentIdentity(confirmationData) {
  loadGameStateFromProperties_();
  
  if (gameState.status !== 'waiting') {
    return { 
      success: false, 
      message: 'Student registration is only available when the game is in waiting room mode.' 
    };
  }

  const { firstName, lastName, nickname } = confirmationData;
  
  if (!firstName || !lastName || !nickname || 
      firstName.trim() === '' || lastName.trim() === '' || nickname.trim() === '') {
    return { 
      success: false, 
      message: 'All fields are required for identity confirmation.' 
    };
  }

  const cleanFirstName = firstName.trim();
  const cleanLastName = lastName.trim();
  const cleanNickname = nickname.trim();
  
  // Find student with matching details
  const existingStudent = gameState.students.find(s => 
    s.firstName.toLowerCase() === cleanFirstName.toLowerCase() &&
    s.lastName.toLowerCase() === cleanLastName.toLowerCase() &&
    s.nickname === cleanNickname
  );
  
  if (!existingStudent) {
    return { 
      success: false, 
      message: 'Could not find a student with those details. Please check your information or start a new session.' 
    };
  }
  
  // Update session key for this student
  const sessionKey = Session.getTemporaryActiveUserKey();
  existingStudent.sessionKey = sessionKey;
  
  saveGameStateToProperties_();
  
  Logger.log(`Student identity confirmed: ${cleanFirstName} ${cleanLastName} as ${cleanNickname}`);
  
  return {
    success: true,
    message: `Welcome back, ${cleanNickname}!`,
    student: {
      firstName: existingStudent.firstName,
      lastName: existingStudent.lastName,
      nickname: existingStudent.nickname
    }
  };
}

// --- Web App Callable Functions ---
function getGameData() {
    loadGameStateFromProperties_(); 

    // Check for automatic advancement due to timer expiry
    if (gameState.status === 'voting' && gameState.activeMatchup && gameState.activeMatchup.votingEndTime && Date.now() >= gameState.activeMatchup.votingEndTime) {
        Logger.log(`Automatic advancement: Voting time expired for matchup ${gameState.activeMatchup.promptA.text} vs ${gameState.activeMatchup.promptB.text}.`);
        
        gameState.status = 'round_over_pending_auto_advance';
        determineWinnerAndAdvance_();

        if (gameState.status === 'game_over') { 
            Logger.log('Auto-advanced: Game is over.');
        } else {
            const success = prepareNextMatchup_(); 
            if (success) {
                gameState.status = 'voting'; 
                Logger.log(`Auto-advanced: Round ${gameState.currentRound + 1}, Matchup ${gameState.currentMatchupIndex + 1} launched after timer.`);
            } else {
                if (gameState.bracket[gameState.currentRound] && gameState.bracket[gameState.currentRound].every(m => m.winner)) {
                    if (gameState.bracket[gameState.currentRound].length === 1 && gameState.bracket[gameState.currentRound][0].winner) {
                        gameState.status = 'game_over';
                        gameState.activeMatchup = null; 
                        Logger.log('Auto-advanced: Game Over! Final winner determined.');
                    } else {
                        gameState.currentRound++;
                        gameState.currentMatchupIndex = 0;
                        setupNextRoundBracket_();

                        if (gameState.status === 'game_over') { 
                            Logger.log('Auto-advanced: Game Over! Determined after setting up next round.');
                        } else {
                            const nextRoundSuccess = prepareNextMatchup_();
                            if (nextRoundSuccess) {
                                gameState.status = 'voting'; 
                                Logger.log(`Auto-advanced: Advanced to Round ${gameState.currentRound + 1}. Next matchup launched after timer.`);
                            } else {
                                if (gameState.status === 'game_over') {
                                    Logger.log('Auto-advanced: Game Over! No more matchups to prepare.');
                                } else {
                                    Logger.log('Auto-advanced: Could not prepare next matchup after advancing round. Game might be stuck or over.');
                                    gameState.status = 'round_over';
                                }
                            }
                        }
                    }
                } else {
                     Logger.log('Auto-advanced: Could not prepare next matchup. Current round may not be complete or an error occurred. Status set to round_over.');
                     gameState.status = 'round_over';
                }
            }
        }
        saveGameStateToProperties_();
    }
  
  let activeMatchupClient = null;
  if (gameState.activeMatchup) {
      const sessionKey = Session.getTemporaryActiveUserKey();
      const student = gameState.students.find(s => s.sessionKey === sessionKey);
      const currentUserHasVoted = student ? gameState.activeMatchup.voters.includes(student.nickname) : false;
      
      activeMatchupClient = {
          ...gameState.activeMatchup,
          currentUserHasVoted: currentUserHasVoted,
          votingStartTime: gameState.activeMatchup.votingStartTime || null,
          votingEndTime: gameState.activeMatchup.votingEndTime || null 
      };
  }
  
  return {
    status: gameState.status,
    activeMatchup: activeMatchupClient, 
    bracket: gameState.bracket, 
    currentRound: gameState.currentRound,
    currentMatchupIndex: gameState.currentMatchupIndex,
    students: gameState.students.map(s => ({ 
      firstName: s.firstName, 
      lastName: s.lastName, 
      nickname: s.nickname 
    }))
  };
}

function submitVote(voteData) {
  loadGameStateFromProperties_();

  if (gameState.status !== 'voting' || !gameState.activeMatchup) {
    return { success: false, message: 'Voting is not currently active or no matchup is live.' };
  }

  if (gameState.activeMatchup.votingEndTime && Date.now() >= gameState.activeMatchup.votingEndTime) {
      Logger.log(`Vote submitted after voting period ended for ${gameState.activeMatchup.codeA} vs ${gameState.activeMatchup.codeB}.`);
      return { 
          success: false, 
          message: "Time's up! Voting for this matchup has ended.",
          updatedMatchup: {
            ...gameState.activeMatchup,
            currentUserHasVoted: false
          }
      };
  }

  const { code, studentNickname } = voteData;
  
  if (!studentNickname) {
    return { success: false, message: 'Student nickname is required to vote.' };
  }
  
  // Verify student is registered
  const student = gameState.students.find(s => s.nickname === studentNickname);
  if (!student) {
    return { success: false, message: 'Student not found. Please re-register.' };
  }
  
  // Verify session key matches
  const sessionKey = Session.getTemporaryActiveUserKey();
  if (student.sessionKey !== sessionKey) {
    return { success: false, message: 'Session mismatch. Please re-register.' };
  }
  
  let voteRegistered = false;
  
  const currentMatchupInBracket = gameState.bracket[gameState.activeMatchup.round][gameState.activeMatchup.matchupIndexInRound];

  if (!currentMatchupInBracket) {
      Logger.log(`Error in submitVote: Active matchup (R${gameState.activeMatchup.round}, M${gameState.activeMatchup.matchupIndexInRound}) not found in bracket.`);
      return { success: false, message: 'Internal error: Active matchup mismatch.'};
  }

  if (!currentMatchupInBracket.voters) currentMatchupInBracket.voters = [];
  if (!gameState.activeMatchup.voters) gameState.activeMatchup.voters = [];
  

  if (currentMatchupInBracket.voters.includes(studentNickname)) {
      Logger.log(`Student ${studentNickname} has already voted in this matchup.`);
      return { 
          success: false, message: 'You have already voted in this matchup.',
          alreadyVoted: true, 
          updatedMatchup: { 
            ...gameState.activeMatchup,
            currentUserHasVoted: true 
          }
      };
  }
  
  let votedForPrompt = '';
  
  if (code === gameState.activeMatchup.codeA) {
    gameState.activeMatchup.votesA++;
    currentMatchupInBracket.votesA = gameState.activeMatchup.votesA;
    votedForPrompt = gameState.activeMatchup.promptA.text;
    voteRegistered = true;
  } else if (code === gameState.activeMatchup.codeB) {
    gameState.activeMatchup.votesB++;
    currentMatchupInBracket.votesB = gameState.activeMatchup.votesB;
    votedForPrompt = gameState.activeMatchup.promptB.text;
    voteRegistered = true;
  }

  if (voteRegistered) {
    currentMatchupInBracket.voters.push(studentNickname);
    if (!gameState.activeMatchup.voters) gameState.activeMatchup.voters = []; 
    gameState.activeMatchup.voters.push(studentNickname);
    
    // Record vote in sheet
    const roundNumber = gameState.activeMatchup.round + 1;
    recordVoteInSheet_(studentNickname, roundNumber, votedForPrompt);
    
    Logger.log(`Vote registered for ${code} by ${studentNickname}. Voters: ${currentMatchupInBracket.voters.length}. Votes: A:${gameState.activeMatchup.votesA}, B:${gameState.activeMatchup.votesB}`);
    saveGameStateToProperties_();
    return {
      success: true, message: 'Vote registered!',
      updatedMatchup: {
          ...gameState.activeMatchup, 
          currentUserHasVoted: true 
      }
    };
  } else {
    return { 
        success: false, message: 'Invalid prompt code.',
        updatedMatchup: { 
            ...gameState.activeMatchup,
            currentUserHasVoted: gameState.activeMatchup.voters ? gameState.activeMatchup.voters.includes(studentNickname) : false
        }
    };
  }
}

// --- Google Sheet Interaction ---
function loadPromptsFromSheet_() {
  try {
    let ss;
    if (SPREADSHEET_ID && SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID') {
        ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    if (!ss) {
        Logger.log('Could not get spreadsheet instance.');
        gameState.prompts = [];
        return;
    }

    const sheet = ss.getSheetByName(PROMPTS_SHEET_NAME);
    if (!sheet) {
      Logger.log(`Sheet "${PROMPTS_SHEET_NAME}" not found in spreadsheet ID: ${ss.getId()}.`);
      gameState.prompts = []; 
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { 
        Logger.log(`No data in "Prompts" sheet (A2 onwards). Last row is ${lastRow}.`);
        gameState.prompts = [];
        return;
    }

    const range = sheet.getRange('A2:A' + lastRow); 
    const data = range.getValues()
                        .map(row => row[0]) 
                        .filter(cellValue => cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== ""); 

    if (data.length === 0) {
        Logger.log('Filtered data is empty. No valid prompts found in Column A of "Prompts" sheet (A2 onwards).');
        gameState.prompts = [];
        return;
    }
    
    gameState.prompts = data.map((text, index) => ({ 
        id: `prompt_${index + 1}_${Utilities.getUuid()}`, 
        text: text.toString().trim() 
    }));

    Logger.log(`Loaded ${gameState.prompts.length} prompts from sheet "${PROMPTS_SHEET_NAME}".`);

  } catch (e) {
    Logger.log(`Error loading prompts from sheet: ${e.toString()}\nStack: ${e.stack}`);
    gameState.prompts = []; 
  }
}