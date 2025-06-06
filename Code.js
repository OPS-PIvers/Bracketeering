// --- Global Variables & Configuration ---
const VOTING_DURATION_SECONDS = 20; // Define voting duration

let gameState = {
  status: 'setup', // 'setup', 'waiting', 'voting', 'round_over', 'game_over'
  currentRound: 0,
  currentMatchupIndex: 0,
  prompts: [], // Will be loaded from sheet: [{id: 1, text: "Prompt A"}, {id: 2, text: "Prompt B"}, ...]
  bracket: [], // Structure: [[{prompt1, prompt2, winner, votes1, votes2, voters: [], votingStartTime, votingEndTime}], [{...next round...}]]
  activeMatchup: null, // { promptA, promptB, votesA, votesB, codeA, codeB, voters: [], votingStartTime, votingEndTime }
  students: [], // List of registered students: [{firstName, lastName, nickname, sessionKey}]
  selectedTab: null // Name of the currently selected tab for this game
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

// --- Tab Management Functions ---
function getAllSheetNames_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    return sheets.map(sheet => sheet.getName());
  } catch (e) {
    Logger.log(`Error getting sheet names: ${e.toString()}`);
    return [];
  }
}

function showTabSelectionDialog_() {
  const sheetNames = getAllSheetNames_();
  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert('No sheets found in this spreadsheet.');
    return null;
  }

  // Create HTML for the dialog
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>Select Tab for Bracket Battle</h3>
      <p>Choose which tab to use for this game:</p>
      <select id="tabSelector" style="width: 100%; padding: 8px; margin: 10px 0;">
        ${sheetNames.map(name => `<option value="${name}">${name}</option>`).join('')}
      </select>
      <br><br>
      <button onclick="selectTab()" style="background: #4285f4; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer;">Select Tab</button>
      <button onclick="google.script.host.close()" style="background: #ccc; color: black; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; margin-left: 10px;">Cancel</button>
    </div>
    <script>
      function selectTab() {
        const selectedTab = document.getElementById('tabSelector').value;
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              google.script.host.close();
            } else {
              alert(result.message || 'Error selecting tab');
            }
          })
          .setSelectedTab(selectedTab);
      }
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Game Tab');
}

function setSelectedTab(tabName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tabName);
    
    if (!sheet) {
      return { success: false, message: 'Selected tab does not exist.' };
    }

    // Store the current active tab globally
    PropertiesService.getScriptProperties().setProperty('currentActiveTab', tabName);

    // Load or create game state for this tab
    loadGameStateFromProperties_(tabName);
    gameState.selectedTab = tabName;
    
    // Initialize/update the sheet structure
    setupTabStructure_(sheet);
    
    // Load prompts from the selected tab
    loadPromptsFromSheet_();
    
    // Continue with game setup (without showing additional alert)
    const setupSuccess = continueGameSetup(false);
    
    if (setupSuccess) {
      return { success: true, message: `Tab "${tabName}" selected and game set to "Waiting Room". Students can now register and join the web app.` };
    } else {
      return { success: false, message: `Tab "${tabName}" selected but not enough prompts found. Please add at least 2 prompts to Column A.` };
    }
  } catch (e) {
    Logger.log(`Error setting selected tab: ${e.toString()}`);
    return { success: false, message: 'Error selecting tab: ' + e.toString() };
  }
}

function setupTabStructure_(sheet) {
  try {
    // Check if headers exist in row 1
    const headerRange = sheet.getRange(1, 1, 1, 4);
    const headers = headerRange.getValues()[0];
    
    const expectedHeaders = ['Prompt', 'First Name', 'Last Name', 'Nickname'];
    const needsHeaders = headers.some((header, index) => header !== expectedHeaders[index]);
    
    if (needsHeaders) {
      // Set up headers
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log(`Set up headers for tab "${sheet.getName()}".`);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error setting up tab structure: ${e.toString()}`);
    return false;
  }
}

// --- Game State Persistence Functions ---
function getGameStateKey_(tabName) {
  return `gameState_${tabName}`;
}

function loadGameStateFromProperties_(tabName = null) {
  try {
    const key = tabName ? getGameStateKey_(tabName) : (gameState.selectedTab ? getGameStateKey_(gameState.selectedTab) : 'gameState');
    const savedStateString = PropertiesService.getScriptProperties().getProperty(key);
    
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
        selectedTab: tabName,
        ...loadedState 
      };
      Logger.log(`GameState loaded from PropertiesService for tab: ${gameState.selectedTab}`);
      
      if ((!gameState.prompts || gameState.prompts.length === 0) && gameState.status !== 'game_over') {
        Logger.log('Prompts empty or not present in loaded state, attempting to load from sheet.');
        loadPromptsFromSheet_();
      }
      if (!gameState.students) {
        gameState.students = [];
      }
    } else {
      Logger.log(`No saved game state found for tab: ${tabName}. Initializing new state.`);
      gameState = {
        status: 'setup',
        currentRound: 0,
        currentMatchupIndex: 0,
        prompts: [],
        bracket: [],
        activeMatchup: null,
        students: [],
        selectedTab: tabName
      };
      if (tabName) {
        loadPromptsFromSheet_();
      }
    }
  } catch (e) {
    Logger.log(`Error loading game state from PropertiesService: ${e.toString()}. Resetting to default.`);
    gameState = {
      status: 'setup',
      currentRound: 0,
      currentMatchupIndex: 0,
      prompts: [],
      bracket: [],
      activeMatchup: null,
      students: [],
      selectedTab: tabName
    };
    if (tabName) {
      loadPromptsFromSheet_();
    }
  }
}

function saveGameStateToProperties_() {
  if (!gameState || typeof gameState !== 'object' || !gameState.selectedTab) {
    Logger.log('CRITICAL_ERROR: Attempted to save game state with invalid global gameState or missing selectedTab. Aborting save.');
    return false;
  }
  try {
    // gameState.selectedTab is already validated by the check above
    const key = getGameStateKey_(gameState.selectedTab);
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(gameState));
    Logger.log(`GameState saved to PropertiesService for tab: ${gameState.selectedTab}`);
    return true; // Added return
  } catch (e) {
    Logger.log(`Error saving game state to PropertiesService: ${e.toString()}`);
    return false; // Added return
  }
}

// --- Sheet Management Functions ---
function addStudentToSheet_(student) {
  try {
    if (!gameState.selectedTab) {
      Logger.log('No selected tab to add student to.');
      return false;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(gameState.selectedTab);
    if (!sheet) {
      Logger.log(`Selected tab "${gameState.selectedTab}" not found.`);
      return false;
    }

    // Find the first empty row in columns B-D (starting from row 2)
    let targetRow = 2;
    const lastRow = sheet.getLastRow();
    
    for (let row = 2; row <= lastRow + 1; row++) {
      const firstName = sheet.getRange(row, 2).getValue();
      if (!firstName || firstName.toString().trim() === '') {
        targetRow = row;
        break;
      }
    }
    
    // Add student data
    sheet.getRange(targetRow, 2, 1, 3).setValues([[
      student.firstName,
      student.lastName,
      student.nickname
    ]]);
    
    Logger.log(`Added student ${student.nickname} to tab "${gameState.selectedTab}" at row ${targetRow}.`);
    return true;
  } catch (e) {
    Logger.log(`Error adding student to sheet: ${e.toString()}`);
    return false;
  }
}

function addRoundColumnHeader_(roundNumber) {
  try {
    if (!gameState.selectedTab) return false;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(gameState.selectedTab);
    if (!sheet) return false;

    const headerRow = 1;
    const roundColumnIndex = 5 + roundNumber - 1; // Columns E, F, G, etc.
    const headerName = `Round ${roundNumber}`;
    
    // Check if header already exists
    const currentHeader = sheet.getRange(headerRow, roundColumnIndex).getValue();
    if (currentHeader === headerName) {
      return true; // Already exists
    }
    
    sheet.getRange(headerRow, roundColumnIndex).setValue(headerName);
    sheet.getRange(headerRow, roundColumnIndex).setFontWeight('bold');
    
    Logger.log(`Added round column header for Round ${roundNumber} in tab "${gameState.selectedTab}".`);
    return true;
  } catch (e) {
    Logger.log(`Error adding round column header: ${e.toString()}`);
    return false;
  }
}

function recordVoteInSheet_(nickname, roundNumber, votedFor) {
  try {
    if (!gameState.selectedTab) return false;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(gameState.selectedTab);
    if (!sheet) return false;

    // Find the student's row
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false; // No data rows
    
    const nicknameColumn = 4; // Column D
    let studentRow = -1;
    
    for (let row = 2; row <= lastRow; row++) {
      const nickname_cell = sheet.getRange(row, nicknameColumn).getValue();
      if (nickname_cell && nickname_cell.toString().trim() === nickname) {
        studentRow = row;
        break;
      }
    }
    
    if (studentRow === -1) {
      Logger.log(`Student with nickname ${nickname} not found in sheet.`);
      return false;
    }
    
    const roundColumnIndex = 5 + roundNumber - 1; // Columns E, F, G, etc.
    
    // Ensure the round column header exists
    addRoundColumnHeader_(roundNumber);
    
    // Get existing votes for this round
    const existingVotes = sheet.getRange(studentRow, roundColumnIndex).getValue();
    let newVoteValue = votedFor;
    
    if (existingVotes && existingVotes.toString().trim() !== '') {
      // Append to existing votes with comma separator
      newVoteValue = existingVotes + ', ' + votedFor;
    }
    
    // Record the vote
    sheet.getRange(studentRow, roundColumnIndex).setValue(newVoteValue);
    
    Logger.log(`Recorded vote for ${nickname} in Round ${roundNumber}: ${newVoteValue}`);
    return true;
  } catch (e) {
    Logger.log(`Error recording vote in sheet: ${e.toString()}`);
    return false;
  }
}

function saveFinalResults_() {
  try {
    if (!gameState.selectedTab) return false;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(gameState.selectedTab);
    if (!sheet) return false;

    // Find the next blank column (starting from column E onwards)
    let resultsColumn = 5; // Start from column E
    const lastCol = sheet.getLastColumn();
    
    for (let col = 5; col <= lastCol + 1; col++) {
      const headerValue = sheet.getRange(1, col).getValue();
      if (!headerValue || headerValue.toString().trim() === '') {
        resultsColumn = col;
        break;
      }
    }

    // Create results summary
    const winner = findFinalWinner_(gameState.bracket);
    const results = [
      ['FINAL RESULTS'],
      [''],
      ['Winner:', winner ? winner.text : 'No winner determined'],
      [''],
      ['Bracket Summary:']
    ];

    // Add bracket information
    gameState.bracket.forEach((round, roundIndex) => {
      results.push([`Round ${roundIndex + 1}:`]);
      round.forEach((matchup, matchupIndex) => {
        const promptAText = matchup.promptA ? matchup.promptA.text : 'TBD';
        const promptBText = matchup.promptB ? (matchup.promptB.id === 'BYE_ID' ? 'BYE' : matchup.promptB.text) : 'TBD';
        const winnerText = matchup.winner ? matchup.winner.text : 'No winner';
        const votesText = `(${matchup.votesA || 0} vs ${matchup.votesB || 0})`;
        
        results.push([`  ${promptAText} vs ${promptBText} ${votesText} → ${winnerText}`]);
      });
      results.push(['']); // Empty row between rounds
    });

    // Write results to sheet
    const range = sheet.getRange(1, resultsColumn, results.length, 1);
    range.setValues(results);
    
    // Format the header
    sheet.getRange(1, resultsColumn).setFontWeight('bold').setFontSize(12);
    
    Logger.log(`Final results saved to column ${resultsColumn} in tab "${gameState.selectedTab}".`);
    return true;
  } catch (e) {
    Logger.log(`Error saving final results: ${e.toString()}`);
    return false;
  }
}

function findFinalWinner_(bracket) {
  if (!bracket || bracket.length === 0) return null;
  const lastRound = bracket[bracket.length - 1];
  if (!lastRound || lastRound.length === 0) return null;
  if (lastRound.length === 1 && lastRound[0].winner) return lastRound[0].winner;
  return null; 
}

function _showGameOverAlert(resultsSaved) {
  let gameOverMessage = 'Game Over! Final winner determined.';
  if (resultsSaved) {
    gameOverMessage += ' Results saved.';
  } else {
    gameOverMessage += ' WARNING: Failed to save detailed results to the spreadsheet. Please check logs.';
  }
  SpreadsheetApp.getUi().alert(gameOverMessage);
}

// --- Custom Menu ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bracket Battle Game')
    .addItem('Select Tab & Start Game', 'startGameMenuItem')
    .addItem('Launch Next Round/Prompts', 'launchNextRoundMenuItem')
    .addItem('Reset Current Game', 'resetGameMenuItem')
    .addSeparator()
    .addItem('Clear ALL Game Data', 'clearAllGameDataMenuItem')
    .addItem('Debug: Show Properties', 'debugShowPropertiesMenuItem')
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
  showTabSelectionDialog_();
}

function continueGameSetup(showAlert = true) {
  if (!gameState.selectedTab) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert('No tab selected. Please select a tab first.');
    }
    return false;
  }

  if (!gameState.prompts || gameState.prompts.length === 0) {
    Logger.log('continueGameSetup: Prompts are empty, attempting to load from sheet.');
    loadPromptsFromSheet_();
  }

  if (gameState.prompts.length < 2) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert(`Not enough prompts in the "${gameState.selectedTab}" tab to start a game. Please add at least 2 prompts to Column A.`);
    }
    return false;
  }

  gameState.status = 'waiting';
  gameState.currentRound = 0;
  gameState.currentMatchupIndex = 0;
  gameState.bracket = [];
  gameState.activeMatchup = null;
  gameState.students = []; // Reset students for new game
  setupInitialBracket_(); 
  
  const saveSuccess = saveGameStateToProperties_();
  if (!saveSuccess && showAlert) {
    SpreadsheetApp.getUi().alert("CRITICAL ERROR: Failed to save game state. Please try the operation again. If the issue persists, contact support or check script logs.");
    return false; // Indicate failure
  }

  if (showAlert) {
    SpreadsheetApp.getUi().alert(`Game set to "Waiting Room" for tab "${gameState.selectedTab}". Students can now register and join the web app.`);
  }
  
  return true;
}

function launchNextRoundMenuItem() {
  // Load the current active tab first
  const currentActiveTab = PropertiesService.getScriptProperties().getProperty('currentActiveTab');
  
  if (!currentActiveTab) {
    SpreadsheetApp.getUi().alert('No tab selected. Please select a tab first using "Select Tab & Start Game".');
    return;
  }

  // Load the game state for the active tab
  loadGameStateFromProperties_(currentActiveTab);

  if (gameState.status === 'setup') {
      SpreadsheetApp.getUi().alert('Please use "Select Tab & Start Game" first to initialize prompts and waiting room.');
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
            const resultsSaved = saveFinalResults_();
            _showGameOverAlert(resultsSaved);
        } else {
            gameState.currentRound++;
            gameState.currentMatchupIndex = 0; 
            setupNextRoundBracket_(); 

            if (gameState.status === 'game_over') { // status might be set by setupNextRoundBracket_
                // If game ended after setting up next round (e.g. only one winner advanced to become overall winner)
                // Ensure activeMatchup is null if game is truly over. setupNextRoundBracket_ already logs this.
                if (gameState.activeMatchup && gameState.bracket[gameState.bracket.length-1].length === 1 && gameState.bracket[gameState.bracket.length-1][0].winner) {
                   gameState.activeMatchup = null; // Ensure consistency
                }
                const resultsSaved = saveFinalResults_();
                _showGameOverAlert(resultsSaved);
            } else {
                const nextRoundSuccess = prepareNextMatchup_();
                if (nextRoundSuccess) {
                    gameState.status = 'voting';
                    SpreadsheetApp.getUi().alert(`Advanced to Round ${gameState.currentRound + 1}. Next matchup launched!`);
                } else {
                    // This 'else' implies prepareNextMatchup_ returned false.
                    // Check if the game ended because no more matchups could be prepared.
                    if(gameState.status === 'game_over'){ 
                        const resultsSaved = saveFinalResults_();
                        _showGameOverAlert(resultsSaved);
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
  const saveSuccess = saveGameStateToProperties_();
  if (!saveSuccess) {
    SpreadsheetApp.getUi().alert("CRITICAL ERROR: Failed to save game state after launching round/matchup. Please try the operation again. If the issue persists, contact support or check script logs.");
  }
}

function resetGameMenuItem() {
  // Load the current active tab first
  const currentActiveTab = PropertiesService.getScriptProperties().getProperty('currentActiveTab');
  
  if (!currentActiveTab) {
    SpreadsheetApp.getUi().alert('No tab selected. Please select a tab first using "Select Tab & Start Game".');
    return;
  }

  // Clear the game state for current tab only
  const key = getGameStateKey_(currentActiveTab);
  PropertiesService.getScriptProperties().deleteProperty(key);
  
  // Reset in-memory state
  gameState = {
    status: 'setup',
    currentRound: 0,
    currentMatchupIndex: 0,
    prompts: [],
    bracket: [],
    activeMatchup: null,
    students: [],
    selectedTab: currentActiveTab
  };
  
  loadPromptsFromSheet_();
  const saveSuccess = saveGameStateToProperties_();
  
  if (!saveSuccess) {
    SpreadsheetApp.getUi().alert(`Game for tab "${currentActiveTab}" has been reset, but CRITICAL ERROR: Failed to save this reset state. Please try resetting again or contact support.`);
  } else {
    SpreadsheetApp.getUi().alert(`Game has been reset for tab "${currentActiveTab}". Prompts reloaded from sheet.`);
  }
}

function clearAllGameDataMenuItem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear ALL Game Data',
    'This will completely clear ALL game data for ALL tabs and reset everything. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Clear ALL properties
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    // Clear all game-related properties
    Object.keys(allProperties).forEach(key => {
      if (key.startsWith('gameState_') || key === 'currentActiveTab') {
        properties.deleteProperty(key);
        Logger.log(`Cleared property: ${key}`);
      }
    });
    
    // Reset in-memory state completely
    gameState = {
      status: 'setup',
      currentRound: 0,
      currentMatchupIndex: 0,
      prompts: [],
      bracket: [],
      activeMatchup: null,
      students: [],
      selectedTab: null
    };
    
    SpreadsheetApp.getUi().alert('ALL game data has been completely cleared. All properties deleted. Please use "Select Tab & Start Game" to begin a new session.');
  }
}

function debugShowPropertiesMenuItem() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const gameProperties = {};
  
  // Filter to show only game-related properties
  Object.keys(properties).forEach(key => {
    if (key.startsWith('gameState_') || key === 'currentActiveTab') {
      gameProperties[key] = properties[key];
    }
  });
  
  let message = 'Current Game Properties:\n\n';
  
  if (Object.keys(gameProperties).length === 0) {
    message += 'No game properties found.';
  } else {
    Object.keys(gameProperties).forEach(key => {
      const value = gameProperties[key];
      if (key === 'currentActiveTab') {
        message += `${key}: ${value}\n`;
      } else {
        // For game states, just show basic info
        try {
          const gameData = JSON.parse(value);
          message += `${key}: Status=${gameData.status}, Students=${gameData.students ? gameData.students.length : 0}\n`;
        } catch (e) {
          message += `${key}: [Error parsing data]\n`;
        }
      }
    });
  }
  
  SpreadsheetApp.getUi().alert('Debug Properties', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- Core Game Logic (same as before) ---
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
  // Load the current active tab's game state
  const currentActiveTab = PropertiesService.getScriptProperties().getProperty('currentActiveTab');
  
  if (!currentActiveTab) {
    return { 
      success: false, 
      message: 'No game session active. Please wait for your teacher to start a game.' 
    };
  }

  loadGameStateFromProperties_(currentActiveTab);
  
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
  
  // Add to sheet first
  const sheetSuccess = addStudentToSheet_(newStudent);
  if (!sheetSuccess) {
    Logger.log(`CRITICAL: Failed to add student ${nickname} to sheet. Registration aborted.`);
    return {
      success: false,
      message: 'Failed to register student due to a server error. Please try again or contact the host.'
    };
  }
  
  // Add to game state only if sheet write was successful
  gameState.students.push(newStudent);
  const saveSuccess = saveGameStateToProperties_();

  if (!saveSuccess) {
    Logger.log('CRITICAL: saveGameStateToProperties_ failed after student registration. Subsequent game operations may rely on stale data until a save succeeds.');
    // Student is in gameState in memory, but it might not persist.
    // Sheet write was successful, so student is on the list.
    return {
      success: false, // Indicate overall operation might have issues.
      message: `Welcome, ${nickname}! You are on the list, but there was a server error saving your registration. Please inform the host.`,
      student: {
        firstName: cleanFirstName,
        lastName: cleanLastName,
        nickname: nickname
      }
    };
  }
  
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
  // Load the current active tab's game state
  const currentActiveTab = PropertiesService.getScriptProperties().getProperty('currentActiveTab');
  
  if (!currentActiveTab) {
    return { 
      success: false, 
      message: 'No game session active. Please wait for your teacher to start a game.' 
    };
  }

  loadGameStateFromProperties_(currentActiveTab);
  
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
  
  const saveSuccess = saveGameStateToProperties_();

  if (!saveSuccess) {
    Logger.log('CRITICAL: saveGameStateToProperties_ failed after student identity confirmation. Subsequent game operations may rely on stale data until a save succeeds.');
    return {
      success: false, // Indicate overall operation might have issues.
      message: `Welcome back, ${cleanNickname}! Your identity is confirmed, but there was a server error saving this session. Please inform the host.`,
      student: {
        firstName: existingStudent.firstName,
        lastName: existingStudent.lastName,
        nickname: existingStudent.nickname
      }
    };
  }
  
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
    // First, get the current active tab
    const currentActiveTab = PropertiesService.getScriptProperties().getProperty('currentActiveTab');
    
    if (!currentActiveTab) {
        // No active tab set, return default state
        return {
            status: 'setup',
            activeMatchup: null,
            bracket: [],
            currentRound: 0,
            currentMatchupIndex: 0,
            students: []
        };
    }

    // Load the game state for the active tab
    loadGameStateFromProperties_(currentActiveTab);

    // Check for automatic advancement due to timer expiry
    if (gameState.status === 'voting' && gameState.activeMatchup && gameState.activeMatchup.votingEndTime && Date.now() >= gameState.activeMatchup.votingEndTime) {
        Logger.log(`Automatic advancement: Voting time expired for matchup ${gameState.activeMatchup.promptA.text} vs ${gameState.activeMatchup.promptB.text}.`);
        
        gameState.status = 'round_over_pending_auto_advance';
        determineWinnerAndAdvance_();

        if (gameState.status === 'game_over') { 
            const resultsSaved = saveFinalResults_(); // Save results when game ends
            if (!resultsSaved) {
                Logger.log('CRITICAL: Game ended and saveFinalResults_ failed during getGameData auto-advance.');
            }
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
                        const resultsSaved = saveFinalResults_(); // Save results when game ends
                        if (!resultsSaved) {
                            Logger.log('CRITICAL: Game ended and saveFinalResults_ failed during getGameData auto-advance.');
                        }
                        Logger.log('Auto-advanced: Game Over! Final winner determined.');
                    } else {
                        gameState.currentRound++;
                        gameState.currentMatchupIndex = 0;
                        setupNextRoundBracket_();

                        if (gameState.status === 'game_over') { 
                            const resultsSaved = saveFinalResults_(); // Save results when game ends
                            if (!resultsSaved) {
                                Logger.log('CRITICAL: Game ended and saveFinalResults_ failed during getGameData auto-advance.');
                            }
                            Logger.log('Auto-advanced: Game Over! Determined after setting up next round.');
                        } else {
                            const nextRoundSuccess = prepareNextMatchup_();
                            if (nextRoundSuccess) {
                                gameState.status = 'voting'; 
                                Logger.log(`Auto-advanced: Advanced to Round ${gameState.currentRound + 1}. Next matchup launched after timer.`);
                            } else {
                                if (gameState.status === 'game_over') {
                                    const resultsSaved = saveFinalResults_(); // Save results when game ends
                                    if (!resultsSaved) {
                                        Logger.log('CRITICAL: Game ended and saveFinalResults_ failed during getGameData auto-advance.');
                                    }
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
        const saveSuccess = saveGameStateToProperties_();
        if (!saveSuccess) {
            Logger.log('CRITICAL: saveGameStateToProperties_ failed in getGameData auto-advance. Subsequent game operations may rely on stale data until a save succeeds.');
        }
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
  // Load the current active tab's game state
  const currentActiveTab = PropertiesService.getScriptProperties().getProperty('currentActiveTab');
  
  if (!currentActiveTab) {
    return { 
      success: false, 
      message: 'No game session active. Please wait for your teacher to start a game.' 
    };
  }

  loadGameStateFromProperties_(currentActiveTab);

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
  
  let votedForPromptText = '';
  let isVoteForA = false;
  let isVoteForB = false;

  if (code === gameState.activeMatchup.codeA) {
    votedForPromptText = gameState.activeMatchup.promptA.text;
    isVoteForA = true;
  } else if (code === gameState.activeMatchup.codeB) {
    votedForPromptText = gameState.activeMatchup.promptB.text;
    isVoteForB = true;
  } else {
    // Invalid code, should not happen if UI is correct, but good to handle
    return { 
        success: false, message: 'Invalid prompt code.',
        updatedMatchup: { 
            ...gameState.activeMatchup,
            currentUserHasVoted: gameState.activeMatchup.voters ? gameState.activeMatchup.voters.includes(studentNickname) : false
        }
    };
  }

  // Record vote in sheet BEFORE updating game state
  const roundNumber = gameState.activeMatchup.round + 1;
  const sheetVoteSuccess = recordVoteInSheet_(studentNickname, roundNumber, votedForPromptText);

  if (!sheetVoteSuccess) {
    Logger.log(`CRITICAL: Failed to record vote in sheet for ${studentNickname}, Round ${roundNumber}, Voted for: ${votedForPromptText}. Vote not counted.`);
    return {
      success: false,
      message: 'Your vote could not be recorded due to a server error. Please try again.'
    };
  }

  // Proceed with game state update only if sheet write was successful
  if (isVoteForA) {
    gameState.activeMatchup.votesA++;
    currentMatchupInBracket.votesA = gameState.activeMatchup.votesA;
  } else if (isVoteForB) {
    gameState.activeMatchup.votesB++;
    currentMatchupInBracket.votesB = gameState.activeMatchup.votesB;
  }

  currentMatchupInBracket.voters.push(studentNickname);
  if (!gameState.activeMatchup.voters) gameState.activeMatchup.voters = [];
  gameState.activeMatchup.voters.push(studentNickname);

  Logger.log(`Vote registered for ${code} by ${studentNickname}. Voters: ${currentMatchupInBracket.voters.length}. Votes: A:${gameState.activeMatchup.votesA}, B:${gameState.activeMatchup.votesB}`);
  const saveSuccess = saveGameStateToProperties_();

  const response = {
    success: true,
    message: 'Vote registered!',
    updatedMatchup: {
        ...gameState.activeMatchup,
        currentUserHasVoted: true
    }
  };

  if (!saveSuccess) {
      Logger.log('CRITICAL: saveGameStateToProperties_ failed after vote submission. Subsequent game operations may rely on stale data until a save succeeds.');
      response.success = false; // Indicate overall operation might have issues.
      response.message = 'Your vote was recorded, but there was a server error saving the game state. Please inform the host.';
  }

  return response;
}

// --- Google Sheet Interaction ---
function loadPromptsFromSheet_() {
  try {
    if (!gameState.selectedTab) {
        Logger.log('No selected tab to load prompts from.');
        gameState.prompts = [];
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(gameState.selectedTab);
    
    if (!sheet) {
      Logger.log(`Tab "${gameState.selectedTab}" not found in spreadsheet.`);
      gameState.prompts = []; 
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { 
        Logger.log(`No data in "${gameState.selectedTab}" tab (A2 onwards). Last row is ${lastRow}.`);
        gameState.prompts = [];
        return;
    }

    const range = sheet.getRange('A2:A' + lastRow); 
    const data = range.getValues()
                        .map(row => row[0]) 
                        .filter(cellValue => cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== ""); 

    if (data.length === 0) {
        Logger.log(`Filtered data is empty. No valid prompts found in Column A of "${gameState.selectedTab}" tab (A2 onwards).`);
        gameState.prompts = [];
        return;
    }
    
    gameState.prompts = data.map((text, index) => ({ 
        id: `prompt_${index + 1}_${Utilities.getUuid()}`, 
        text: text.toString().trim() 
    }));

    Logger.log(`Loaded ${gameState.prompts.length} prompts from tab "${gameState.selectedTab}".`);

  } catch (e) {
    Logger.log(`Error loading prompts from sheet: ${e.toString()}\nStack: ${e.stack}`);
    gameState.prompts = []; 
  }
}