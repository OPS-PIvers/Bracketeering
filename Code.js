// --- Global Variables & Configuration ---
const SPREADSHEET_ID = '1aVmXdoThzsfjsEtV-vys3movt5FPc2QYXAWJlntPSh4'; // Replace with your Google Sheet ID
const PROMPTS_SHEET_NAME = 'Prompts'; // Sheet name where prompts are listed
const GAME_STATE_SHEET_NAME = 'GameState'; // Sheet for storing game state (optional, can also use PropertiesService)
const VOTING_DURATION_SECONDS = 20; // Define voting duration

let gameState = {
  status: 'setup', // 'setup', 'waiting', 'voting', 'round_over', 'game_over'
  currentRound: 0,
  currentMatchupIndex: 0,
  prompts: [], // Will be loaded from sheet: [{id: 1, text: "Prompt A"}, {id: 2, text: "Prompt B"}, ...]
  bracket: [], // Structure: [[{prompt1, prompt2, winner, votes1, votes2, voters: [], votingStartTime, votingEndTime}], [{...next round...}]]
  activeMatchup: null // { promptA, promptB, votesA, votesB, codeA, codeB, voters: [], votingStartTime, votingEndTime }
};

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
        ...loadedState 
      };
      Logger.log('GameState loaded from PropertiesService.');
      if ((!gameState.prompts || gameState.prompts.length === 0) && gameState.status !== 'game_over') {
        Logger.log('Prompts empty or not present in loaded state, attempting to load from sheet.');
        loadPromptsFromSheet_();
      }
    } else {
      Logger.log('No saved game state found in PropertiesService. Initializing and loading prompts.');
      gameState = {
        status: 'setup',
        currentRound: 0,
        currentMatchupIndex: 0,
        prompts: [],
        bracket: [],
        activeMatchup: null
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
        activeMatchup: null
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
    // Avoid SpreadsheetApp.getUi().alert() in server-side logic that might run automatically
    Logger.log('Critical Error: Could not save game state. Please check logs for teacher action if issue persists.');
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
  setupInitialBracket_(); 

  SpreadsheetApp.getUi().alert('Game set to "Waiting Room". Students can now open the web app.');
  saveGameStateToProperties_();
}

function launchNextRoundMenuItem() {
  loadGameStateFromProperties_();

  if (gameState.status === 'setup') {
      SpreadsheetApp.getUi().alert('Please "Start Game" first to initialize prompts and waiting room.');
      return;
  }

  // If voting was active, determine winner. This also handles cases where teacher ends voting early.
  if (gameState.status === 'voting' && gameState.activeMatchup) {
    Logger.log('Teacher manually advancing from voting state.');
    determineWinnerAndAdvance_(); // This will set status to 'round_over'
  }

  if (gameState.status === 'game_over') {
    SpreadsheetApp.getUi().alert('Game is over. Please reset to start a new game.');
    return;
  }

  // At this point, status should be 'waiting', 'round_over', or just became 'round_over'
  const success = prepareNextMatchup_(); 
  if (success) {
    gameState.status = 'voting'; // prepareNextMatchup_ sets new activeMatchup with new timer
    SpreadsheetApp.getUi().alert(`Round ${gameState.currentRound + 1}, Matchup ${gameState.currentMatchupIndex + 1} launched!`);
  } else {
    // No more matchups in current round, or couldn't prepare one.
    // Check if all matchups in the current round have winners
    if (gameState.bracket[gameState.currentRound] && gameState.bracket[gameState.currentRound].every(m => m.winner)) {
        if (gameState.bracket[gameState.currentRound].length === 1 && gameState.bracket[gameState.currentRound][0].winner) {
            gameState.status = 'game_over';
            gameState.activeMatchup = null;
            SpreadsheetApp.getUi().alert('Game Over! Final winner determined.');
        } else {
            // Advance to the next round
            gameState.currentRound++;
            gameState.currentMatchupIndex = 0; 
            setupNextRoundBracket_(); 

            if (gameState.status === 'game_over') { // setupNextRoundBracket_ might set game_over
                SpreadsheetApp.getUi().alert('Game Over! Determined after setting up next round.');
            } else {
                const nextRoundSuccess = prepareNextMatchup_();
                if (nextRoundSuccess) {
                    gameState.status = 'voting'; // New matchup, new timer
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
      // This case implies current round is not finished, but prepareNextMatchup failed.
      // This might happen if launchNextRoundMenuItem is called when status is 'waiting' and no matchups are ready.
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
    activeMatchup: null
  };
  loadPromptsFromSheet_();
  SpreadsheetApp.getUi().alert('Game has been reset. Prompts reloaded from sheet.');
  saveGameStateToProperties_();
}

// --- Core Game Logic ---

function setupInitialBracket_() {
  if (!gameState.prompts || gameState.prompts.length === 0) {
      Logger.log('setupInitialBracket_: No prompts available. Attempting load.');
      loadPromptsFromSheet_();
      if (!gameState.prompts || gameState.prompts.length === 0) {
          Logger.log('Cannot set up bracket: No prompts found even after reload.');
          // Avoid UI alert here if it can be called automatically
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
      votingStartTime: now, // ADDED
      votingEndTime: votingEndTime // ADDED
    };
    // Also store voting end time in the bracket data for persistence
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
  if (!gameState.activeMatchup ) { // Removed status check here, as it might be called after timer
    Logger.log('determineWinnerAndAdvance_: No active matchup to determine winner.');
    return; 
  }
   if (gameState.status !== 'voting' && gameState.status !== 'round_over_pending_auto_advance') {
      // If status is already 'round_over' (e.g. teacher clicked advance after timer already processed it), do nothing.
      // Or if status is not voting (e.g. 'waiting'), do nothing.
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
    gameState.status = 'round_over'; // Set status to indicate the round for this matchup is over
    Logger.log(`Winner of R${round+1}, M${matchupIndexInRound+1} ("${promptA.text}" vs "${promptB.text}") is: "${winner.text}". Status set to round_over.`);
  } else {
    Logger.log(`Error in determineWinnerAndAdvance_: Could not find matchup in bracket at R${round}, M${matchupIndexInRound} to update winner.`);
  }
}


// --- Web App Callable Functions ---
function getGameData() {
    loadGameStateFromProperties_(); 

    // Check for automatic advancement due to timer expiry
    if (gameState.status === 'voting' && gameState.activeMatchup && gameState.activeMatchup.votingEndTime && Date.now() >= gameState.activeMatchup.votingEndTime) {
        Logger.log(`Automatic advancement: Voting time expired for matchup ${gameState.activeMatchup.promptA.text} vs ${gameState.activeMatchup.promptB.text}.`);
        
        gameState.status = 'round_over_pending_auto_advance'; // Temporary status
        determineWinnerAndAdvance_(); // This sets gameState.status to 'round_over'

        // Now, attempt to prepare the next state (next matchup or next round)
        if (gameState.status === 'game_over') { 
            Logger.log('Auto-advanced: Game is over.');
        } else {
            // gameState.status is 'round_over'
            const success = prepareNextMatchup_(); 
            if (success) {
                gameState.status = 'voting'; 
                Logger.log(`Auto-advanced: Round ${gameState.currentRound + 1}, Matchup ${gameState.currentMatchupIndex + 1} launched after timer.`);
            } else {
                // No more matchups in current round, or couldn't prepare one.
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
                                    gameState.status = 'round_over'; // Revert to round_over if stuck, teacher might need to intervene
                                }
                            }
                        }
                    }
                } else {
                     Logger.log('Auto-advanced: Could not prepare next matchup. Current round may not be complete or an error occurred. Status set to round_over.');
                     gameState.status = 'round_over'; // Revert to round_over, teacher might need to intervene
                }
            }
        }
        saveGameStateToProperties_(); // Save the new state after auto-advancement
    }
  
  let activeMatchupClient = null;
  if (gameState.activeMatchup) {
      const userKey = Session.getTemporaryActiveUserKey();
      const votersList = gameState.activeMatchup.voters || [];
      const currentUserHasVoted = votersList.includes(userKey);
      
      activeMatchupClient = {
          ...gameState.activeMatchup,
          // voters: votersList, // Client might not need the full list for display
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
    currentMatchupIndex: gameState.currentMatchupIndex 
  };
}

function submitVote(voteData) {
  loadGameStateFromProperties_();

  if (gameState.status !== 'voting' || !gameState.activeMatchup) {
    return { success: false, message: 'Voting is not currently active or no matchup is live.' };
  }

  // Check if voting time has expired on the server
  if (gameState.activeMatchup.votingEndTime && Date.now() >= gameState.activeMatchup.votingEndTime) {
      Logger.log(`Vote submitted by ${Session.getTemporaryActiveUserKey()} after voting period ended for ${gameState.activeMatchup.codeA} vs ${gameState.activeMatchup.codeB}.`);
      return { 
          success: false, 
          message: "Time's up! Voting for this matchup has ended.",
          alreadyVoted: false, // Not 'alreadyVoted', but 'timeUp'
          updatedMatchup: { // Return current state so UI can still update vote counts if needed
            ...gameState.activeMatchup,
            currentUserHasVoted: (gameState.activeMatchup.voters || []).includes(Session.getTemporaryActiveUserKey())
          }
      };
  }

  const userKey = Session.getTemporaryActiveUserKey();
  const { code } = voteData;
  let voteRegistered = false;
  
  const currentMatchupInBracket = gameState.bracket[gameState.activeMatchup.round][gameState.activeMatchup.matchupIndexInRound];

  if (!currentMatchupInBracket) {
      Logger.log(`Error in submitVote: Active matchup (R${gameState.activeMatchup.round}, M${gameState.activeMatchup.matchupIndexInRound}) not found in bracket.`);
      return { success: false, message: 'Internal error: Active matchup mismatch.'};
  }

  if (!currentMatchupInBracket.voters) currentMatchupInBracket.voters = [];
  if (!gameState.activeMatchup.voters) gameState.activeMatchup.voters = [];
  

  if (currentMatchupInBracket.voters.includes(userKey)) {
      Logger.log(`User ${userKey} has already voted in this matchup.`);
      return { 
          success: false, message: 'You have already voted in this matchup.',
          alreadyVoted: true, 
          updatedMatchup: { 
            ...gameState.activeMatchup,
            currentUserHasVoted: true 
          }
      };
  }
  
  if (code === gameState.activeMatchup.codeA) {
    gameState.activeMatchup.votesA++;
    currentMatchupInBracket.votesA = gameState.activeMatchup.votesA;
    voteRegistered = true;
  } else if (code === gameState.activeMatchup.codeB) {
    gameState.activeMatchup.votesB++;
    currentMatchupInBracket.votesB = gameState.activeMatchup.votesB;
    voteRegistered = true;
  }

  if (voteRegistered) {
    currentMatchupInBracket.voters.push(userKey);
    if (!gameState.activeMatchup.voters) gameState.activeMatchup.voters = []; 
    gameState.activeMatchup.voters.push(userKey);
    
    Logger.log(`Vote registered for ${code} by ${userKey}. Voters: ${currentMatchupInBracket.voters.length}. Votes: A:${gameState.activeMatchup.votesA}, B:${gameState.activeMatchup.votesB}`);
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
            currentUserHasVoted: gameState.activeMatchup.voters ? gameState.activeMatchup.voters.includes(userKey) : false
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
        // Avoid UI alert here
        gameState.prompts = [];
        return;
    }

    const sheet = ss.getSheetByName(PROMPTS_SHEET_NAME);
    if (!sheet) {
      Logger.log(`Sheet "${PROMPTS_SHEET_NAME}" not found in spreadsheet ID: ${ss.getId()}.`);
      gameState.prompts = []; 
      // Avoid UI alert here
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
    // Avoid UI alert here
    gameState.prompts = []; 
  }
}
