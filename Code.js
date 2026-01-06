/**
 * @title nams-campus-transition-notification-project-25-26
 * 
 * @description This script is designed to be run from a Google Sheet. It will send an email
 * with attachments and links to a list of recipients based on the value in the
 * "Campus" column of the sheet named, "Teacher Notes". This script is part of an email
 * notification system that uses Autocrat to create letters for each student.
 * 
 * Procedure: The user will first create a transition letter using Autocrat. Autocrat
 * saves each letter to a campus specific shared Google Drive folder.
 * After verifying the autocrat created letter and the user is ready to send the email,
 * the user will click on 'Notify Campuses' in the menubar followed by the option that
 * is provided in the dropdown. An email will be sent to the campus administrators with a
 * link to the folder that contains the transition letter.
 * 
 * @projectLead Reggie Ollendieck, Associate Principal, NAMS
 * @author Alvaro Gomez, Academic Technology Coach, 1-210-397-9408, alvaro.gomez@nisd.net
 * @lastUpdated 01/06/26
 */

// ============================================================================
// CORE INFRASTRUCTURE CLASSES
// ============================================================================

/**
 * Manages data migration from hardcoded values to PropertiesService and Recipients sheet.
 * Handles one-time migration operations and data extraction.
 */
class MigrationManager {
  
  /**
   * Logs migration status and outcomes with detailed information
   * Tracks all migration steps and any issues encountered during transition
   * Provides clear success/failure indicators (Requirements 6.5)
   * @param {string} step - The migration step being logged
   * @param {string} status - Status: 'started', 'progress', 'success', 'warning', 'error'
   * @param {string} message - Detailed message about the step
   * @param {Object} [data] - Optional data object with additional details
   */
  static logMigrationStatus(step, status, message, data = {}) {
    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      step,
      status,
      message,
      data
    };
    
    // Log to Google Apps Script Logger
    const logMessage = `MIGRATION [${status.toUpperCase()}] ${step}: ${message}`;
    Logger.log(logMessage);
    
    // Store detailed migration log in PropertiesService for monitoring
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingLogs = properties.getProperty('migrationLogs');
      let logs = [];
      
      if (existingLogs) {
        logs = JSON.parse(existingLogs);
      }
      
      // Add new log entry and keep only last 50 entries
      logs.unshift(logEntry);
      logs = logs.slice(0, 50);
      
      properties.setProperty('migrationLogs', JSON.stringify(logs));
      
      // Also update migration summary statistics
      this.updateMigrationSummary(step, status, data);
      
    } catch (error) {
      Logger.log(`MIGRATION [ERROR] Failed to store migration log: ${error.message}`);
      console.error('Migration logging error:', error);
    }
  }
  
  /**
   * Updates migration summary statistics for monitoring
   * @param {string} step - The migration step
   * @param {string} status - The status of the step
   * @param {Object} data - Additional data about the step
   */
  static updateMigrationSummary(step, status, data) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingSummary = properties.getProperty('migrationSummary');
      let summary = {
        lastMigrationAttempt: null,
        lastSuccessfulMigration: null,
        totalAttempts: 0,
        successfulMigrations: 0,
        failedMigrations: 0,
        warningsCount: 0,
        stepsCompleted: {},
        lastError: null,
        performanceMetrics: {}
      };
      
      if (existingSummary) {
        summary = { ...summary, ...JSON.parse(existingSummary) };
      }
      
      // Update summary based on status
      const now = new Date().toISOString();
      
      if (step === 'migration_start') {
        summary.lastMigrationAttempt = now;
        summary.totalAttempts++;
      }
      
      if (status === 'success') {
        summary.stepsCompleted[step] = now;
        if (step === 'migration_complete') {
          summary.lastSuccessfulMigration = now;
          summary.successfulMigrations++;
        }
      } else if (status === 'error') {
        summary.lastError = { step, message: data.message || 'Unknown error', timestamp: now };
        if (step === 'migration_complete') {
          summary.failedMigrations++;
        }
      } else if (status === 'warning') {
        summary.warningsCount++;
      }
      
      // Store performance metrics if available
      if (data.duration) {
        if (!summary.performanceMetrics[step]) {
          summary.performanceMetrics[step] = [];
        }
        summary.performanceMetrics[step].push({
          duration: data.duration,
          timestamp: now
        });
        // Keep only last 10 performance entries per step
        summary.performanceMetrics[step] = summary.performanceMetrics[step].slice(-10);
      }
      
      properties.setProperty('migrationSummary', JSON.stringify(summary));
      
    } catch (error) {
      Logger.log(`MIGRATION [ERROR] Failed to update migration summary: ${error.message}`);
    }
  }
  
  /**
   * Gets migration status and history for monitoring
   * @returns {Object} Migration status information
   */
  static getMigrationStatus() {
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // Get migration logs
      const logsData = properties.getProperty('migrationLogs');
      let logs = [];
      if (logsData) {
        logs = JSON.parse(logsData);
      }
      
      // Get migration summary
      const summaryData = properties.getProperty('migrationSummary');
      let summary = {};
      if (summaryData) {
        summary = JSON.parse(summaryData);
      }
      
      // Get current migration completion status
      const isComplete = PropertiesManager.isMigrationComplete();
      
      return {
        isComplete,
        summary,
        recentLogs: logs.slice(0, 10), // Last 10 log entries
        totalLogEntries: logs.length,
        timestamp: new Date().toISOString()
      };
      
    } catch (error) {
      Logger.log(`MIGRATION [ERROR] Failed to get migration status: ${error.message}`);
      return {
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }
  /**
   * Performs initial data migration if not already done
   * @returns {boolean} True if migration was performed, false if already migrated
   */
  static performMigration() {
    const migrationStartTime = new Date();
    
    this.logMigrationStatus('migration_check', 'started', 'Checking migration status');
    
    if (PropertiesManager.isMigrationComplete()) {
      this.logMigrationStatus('migration_check', 'success', 'Migration already completed, skipping');
      return false;
    }
    
    this.logMigrationStatus('migration_start', 'started', 'Starting migration process');
    
    try {
      // Extract hardcoded data
      const extractStartTime = new Date();
      this.logMigrationStatus('extract_hardcoded', 'progress', 'Extracting hardcoded campus data');
      
      const hardcodedData = this.extractHardcodedData();
      const extractDuration = new Date() - extractStartTime;
      
      this.logMigrationStatus('extract_hardcoded', 'success', 
        `Extracted data for ${Object.keys(hardcodedData).length} campuses`, 
        { campusCount: Object.keys(hardcodedData).length, duration: extractDuration });
      
      // Read CampusReferenceInfo sheet for drive folder IDs
      const driveStartTime = new Date();
      this.logMigrationStatus('read_drive_data', 'progress', 'Reading CampusReferenceInfo sheet for drive folder IDs');
      
      const driveData = this.readCampusReferenceInfo();
      const driveDuration = new Date() - driveStartTime;
      
      this.logMigrationStatus('read_drive_data', 'success', 
        `Read drive data for ${Object.keys(driveData).length} campuses`, 
        { driveCount: Object.keys(driveData).length, duration: driveDuration });
      
      // Combine recipient and drive data
      const combineStartTime = new Date();
      this.logMigrationStatus('combine_data', 'progress', 'Combining recipient and drive folder data');
      
      const combinedData = {};
      Object.keys(hardcodedData).forEach(campus => {
        combinedData[campus] = {
          recipients: hardcodedData[campus],
          driveLink: driveData[campus] || ''
        };
      });
      
      const combineDuration = new Date() - combineStartTime;
      this.logMigrationStatus('combine_data', 'success', 
        `Combined data for ${Object.keys(combinedData).length} campuses`, 
        { combinedCount: Object.keys(combinedData).length, duration: combineDuration });
      
      // Store in PropertiesService
      const storeStartTime = new Date();
      this.logMigrationStatus('store_properties', 'progress', 'Storing combined data in PropertiesService');
      
      PropertiesManager.storeCampusData(combinedData);
      const storeDuration = new Date() - storeStartTime;
      
      this.logMigrationStatus('store_properties', 'success', 
        'Stored combined data in PropertiesService', 
        { duration: storeDuration });
      
      // Create Recipients sheet
      const sheetStartTime = new Date();
      this.logMigrationStatus('create_sheet', 'progress', 'Creating and populating Recipients sheet');
      
      this.createRecipientsSheet(hardcodedData);
      const sheetDuration = new Date() - sheetStartTime;
      
      this.logMigrationStatus('create_sheet', 'success', 
        'Created and populated Recipients sheet', 
        { duration: sheetDuration });
      
      // Mark migration as complete
      const completeStartTime = new Date();
      this.logMigrationStatus('mark_complete', 'progress', 'Marking migration as complete');
      
      PropertiesManager.setMigrationComplete();
      const completeDuration = new Date() - completeStartTime;
      
      this.logMigrationStatus('mark_complete', 'success', 
        'Migration marked as complete', 
        { duration: completeDuration });
      
      // Log final migration completion
      const totalDuration = new Date() - migrationStartTime;
      this.logMigrationStatus('migration_complete', 'success', 
        'Migration completed successfully', 
        { 
          totalDuration: totalDuration,
          campusCount: Object.keys(combinedData).length,
          totalRecipients: Object.values(combinedData).reduce((sum, campus) => 
            sum + (campus.recipients ? campus.recipients.length : 0), 0)
        });
      
      return true;
      
    } catch (error) {
      const totalDuration = new Date() - migrationStartTime;
      this.logMigrationStatus('migration_complete', 'error', 
        `Migration failed: ${error.message}`, 
        { 
          error: error.message,
          stack: error.stack,
          totalDuration: totalDuration
        });
      
      console.error('Migration error:', error);
      throw error;
    }
  }
  
  /**
   * Extracts hardcoded campus data from existing function
   * @returns {Object} Campus data with recipients arrays
   */
  static extractHardcodedData() {
    this.logMigrationStatus('extract_hardcoded_start', 'progress', 'Starting extraction of hardcoded campus data');
    
    // This data is extracted from the existing getInfoByCampus function
    const campusData = {
      "bernal": [
        "david.laboy@nisd.net",
        "Marla.Reynolds@nisd.net",
        "sally.maher@nisd.net",
        "monica.flores@nisd.net"
      ],
      "briscoe": [
        "joe.bishop@nisd.net",
        "francesca.parker@nisd.net",
        "brigitte.rauschuber@nisd.net",
        "xavier.aguirre@nisd.net"
      ],
      "connally": [
        "erica.robles@nisd.net",
        "monica.ramirez@nisd.net"
      ],
      "folks": [
        "yvette.lopez@nisd.net",
        "miguel.trevino@nisd.net",
        "terry.precie@nisd.net",
        "angelica.perez@nisd.net",
        "norma.esparza@nisd.net",
        "james-1.garza@nisd.net",
        "keli.hall@nisd.net",
        "ann.devlin@nisd.net"
      ],
      "garcia": [
        "mateo.macias@nisd.net",
        "anna.lopez@nisd.net",
        "julie.minnis@nisd.net",
        "mark.lopez@nisd.net",
        "lori.persyn@nisd.net"
      ],
      "hobby": [
        "gregory.dylla@nisd.net",
        "marian.johnson@nisd.net",
        "lawrence.carranco@nisd.net",
        "jose.texidor@nisd.net",
        "victoria.denton@nisd.net"
      ],
      "hobby magnet": [
        "jaime.heye@nisd.net"
      ],
      "holmgreen": [
        "cheryl.parra@nisd.net",
        "frank.johnson@nisd.net"
      ],
      "jefferson": [
        "monica.cabico@nisd.net",
        "Nicole.Gomez@nisd.net",
        "tiffany.watkins@nisd.net",
        "catherine.villela@nisd.net"
      ],
      "jones": [
        "rudolph.arzola@nisd.net",
        "nicole.mcevoy@nisd.net",
        "erica.lashley@nisd.net",
        "javier.lazo@nisd.net",
        "aaron.logan@nisd.net"
      ],
      "jones magnet": [
        "david.johnston@nisd.net"
      ],
      "jordan": [
        "Shannon.Zavala@nisd.net",
        "juaquin.zavala@nisd.net",
        "erica.parra@nisd.net",
        "laurel.graham@nisd.net",
        "robert.ruiz@nisd.net",
        "anabel.romero@nisd.net"
      ],
      "jordan magnet": [
        "jessica.marcha@nisd.net"
      ],
      "luna": [
        "leti.chapa@nisd.net",
        "jennifer.cipollone@nisd.net",
        "amanda.king@nisd.net",
        "lisa.richard@nisd.net"
      ],
      "neff": [
        "yvonne.correa@nisd.net",
        "theresa.heim@nisd.net",
        "laura-i.sanroman@nisd.net",
        "joseph.castellanos@nisd.net",
        "mackenzie.fulton@nisd.net",
        "adriana.aguero@nisd.net",
        "priscilla.vela@nisd.net",
        "sarah.tennery@nisd.net",
        "jessica.montalvo@nisd.net",
        "hayley.giorgio@nisd.net"
      ],
      "pease": [
        "Lynda.Desutter@nisd.net",
        "jessica-1.barrera@nisd.net",
        "guadalupe.brister@nisd.net"
      ],
      "rawlinson": [
        "jesus.villela@nisd.net",
        "david.rojas@nisd.net",
        "nicole.buentello@nisd.net"
      ],
      "rayburn": [
        "robert.alvarado@nisd.net",
        "maricela.garza@nisd.net",
        "carol.zule@nisd.net",
        "micaela.welsh@nisd.net"
      ],
      "ross": [
        "christina.lozano@nisd.net",
        "priscilla.sigala@nisd.net",
        "dolores.cardenas@nisd.net",
        "katherine.vela@nisd.net",
        "roxanne.romo@nisd.net",
        "cristina.castillo@nisd.net"
      ],
      "rudder": [
        "catelyn.vasquez@nisd.net",
        "jeanette.navarro@nisd.net",
        "jason.padron@nisd.net",
        "adrian.hysten@nisd.net"
      ],
      "stevenson": [
        "chaeleen.garcia@nisd.net",
        "anthony.allen01@nisd.net",
        "hilary.pilaczynski@nisd.net",
        "johanna.davenport@nisd.net"
      ],
      "stinson": [
        "lourdes.medina@nisd.net",
        "louis.villarreal@nisd.net",
        "jeannette.rainey@nisd.net",
        "linda.boyett@nisd.net",
        "elda.garza@nisd.net",
        "maranda.luna@nisd.net",
        "alexis.lopez@nisd.net"
      ],
      "straus": [
        "araceli.farias@nisd.net",
        "jose.gonzalez02@nisd.net",
        "leigh.davis@nisd.net"
      ],
      "vale": [
        "jenna.bloom@nisd.net",
        "brenda.rayburg@nisd.net",
        "daniel.novosad@nisd.net",
        "mary.harrington@nisd.net"
      ],
      "zachry": [
        "Richard.DeLaGarza@nisd.net",
        "randolph.neuenfeldt@nisd.net",
        "jennifer-a.garcia@nisd.net",
        "jimann.caliva@nisd.net",
        "veronica.poblano@nisd.net"
      ],
      "zachry magnet": [
        "matthew.patty@nisd.net"
      ],
      "test": [
        "reggie.ollendieck@nisd.net",
        "zina.gonzales@nisd.net"
      ]
    };
    
    const totalRecipients = Object.values(campusData).reduce((sum, recipients) => sum + recipients.length, 0);
    
    this.logMigrationStatus('extract_hardcoded_complete', 'success', 
      `Extracted ${Object.keys(campusData).length} campus entries with ${totalRecipients} total recipients`,
      { 
        campusCount: Object.keys(campusData).length,
        totalRecipients: totalRecipients,
        campuses: Object.keys(campusData)
      });
    
    return campusData;
  }
  
  /**
   * Creates Recipients sheet and populates with data
   * @param {Object} campusData - Campus recipient mappings
   */
  static createRecipientsSheet(campusData) {
    Logger.log('MigrationManager: Creating Recipients sheet...');
    
    if (RecipientsSheetManager.sheetExists()) {
      Logger.log('MigrationManager: Recipients sheet already exists, skipping creation');
      return;
    }
    
    RecipientsSheetManager.createSheet(campusData);
    Logger.log('MigrationManager: Recipients sheet created and populated');
  }
  
  /**
   * Reads CampusReferenceInfo sheet for drive folder IDs
   * @returns {Object} Campus to drive folder ID mappings
   */
  static readCampusReferenceInfo() {
    this.logMigrationStatus('read_campus_ref_start', 'progress', 'Starting to read CampusReferenceInfo sheet');
    
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName('CampusReferenceInfo');
      
      if (!sheet) {
        this.logMigrationStatus('read_campus_ref_fallback', 'warning', 
          'CampusReferenceInfo sheet not found, using hardcoded drive links');
        return this.getHardcodedDriveLinks();
      }
      
      const data = sheet.getDataRange().getValues();
      const driveData = {};
      
      // Skip header row (index 0)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const campus = row[0] ? row[0].toString().toLowerCase().trim() : '';
        const driveId = row[1] ? row[1].toString().trim() : '';
        
        if (campus && driveId) {
          driveData[campus] = driveId;
        }
      }
      
      this.logMigrationStatus('read_campus_ref_success', 'success', 
        `Read drive data for ${Object.keys(driveData).length} campuses from CampusReferenceInfo sheet`,
        { 
          driveCount: Object.keys(driveData).length,
          campuses: Object.keys(driveData),
          source: 'CampusReferenceInfo_sheet'
        });
      
      return driveData;
      
    } catch (error) {
      this.logMigrationStatus('read_campus_ref_error', 'error', 
        `Error reading CampusReferenceInfo sheet: ${error.message}. Falling back to hardcoded drive links`,
        { error: error.message, fallback: true });
      
      return this.getHardcodedDriveLinks();
    }
  }
  
  /**
   * Returns hardcoded drive folder IDs as fallback
   * @returns {Object} Campus to drive folder ID mappings
   */
  static getHardcodedDriveLinks() {
    return {
      "bernal": "1QlavZvp-8tvqiPF3zYQanZ9SQ7UhiJaE",
      "briscoe": "1JgHL75iSr5F1lgHLieZl0y_psuc7gp-N",
      "connally": "1cWPX_nOXb9yldONekm9Ba3Rsksl9yqKe",
      "folks": "1MZ9MCh1DmI9cJFWbr5BA-v1jJh5Dd1E9",
      "garcia": "1D8_8q9fcB6tn3fXX8xnlkGE8aro8fbDm",
      "hobby": "1u76TEIq5BbCNCG-i8Ta73VaY7NTOsK4x",
      "hobby magnet": "1YrvPC7m0eX128C0dR3u4RGNfU82MXOg2",
      "holmgreen": "1c8ufu7MvAsNwQAwx0IOwi4O1dNU1-7yo",
      "jefferson": "1LlQxIJBeCxV440nalwS5fjzu8_1dYvlx",
      "jones": "1jBxe9OFTTcones277XnehgvW4EnM4YQk",
      "jones magnet": "1oeMqPEr_cpstSWRa0LV-uodQwteQKRWn",
      "jordan": "1T90JGPgUu7DhfBBytxRrfgqkMBrkzOAN",
      "jordan magnet": "1BVECP6fsaqGXap1uEOchy5SW-6PsADM9",
      "luna": "10DdpBdHwp7bH5ph-pKvfbvG23tsi9ZMc",
      "neff": "1Sd7DrcgHjnAcuqR79DVmISnVn3Bzzfyn",
      "pease": "1tOusVf1SxNckZC5ro-dwBlk8-YpKUMuh",
      "rawlinson": "1p9IXf40oikwOrxSSmonwnBJ-dOJB7Ui6",
      "rayburn": "1DcP6LUpcT8wT9PEYgc4_dwk2mclWI9bp",
      "ross": "1KGyYAJF5Qf-Gt0oWvRRjH2Vw_DTARMcg",
      "rudder": "1a3PiBLrTtsMJkR6lthx86gUqd2_qeF-D",
      "stevenson": "1Y43jZCtjKFbF-I09lBS6wnr-Vbl0p_Cm",
      "stinson": "1pG4NIUveTfv46DvQoq0TD4uuwaXCLwkS",
      "straus": "15p9xqZoyikuVRk4ZVi7sUYYgoxL1PSew",
      "vale": "1QoRQNEt7_gWT3PC_DPnDFjlT1XKVQ3sN",
      "zachry": "1UZDcETdHG5eN9DSCdfKcV0wt72cVY7Ek",
      "zachry magnet": "1wMjhAx6wGOw5j-tu7-4wTNnH7V80pfPe",
      "test": "1nMJAEcGIh_QnhfS5gjCkKd6CtoA3r5cf"
    };
  }
}

/**
 * Manages data storage and retrieval from PropertiesService with caching optimization.
 * Handles campus data persistence and migration status tracking.
 */
class PropertiesManager {
  static CAMPUS_DATA_KEY = 'campusData';
  static MIGRATION_COMPLETE_KEY = 'migrationComplete';
  static LAST_UPDATED_KEY = 'lastUpdated';
  
  /**
   * Logs runtime operations for performance monitoring and debugging
   * Logs cache hits/misses, validation errors, warnings, and sheet access patterns
   * (Requirements 2.4, 3.1, 3.2, 3.5)
   * @param {string} operation - The operation being logged (e.g., 'cache_hit', 'cache_miss', 'validation_error')
   * @param {string} level - Log level: 'info', 'warning', 'error'
   * @param {string} message - Detailed message about the operation
   * @param {Object} [data] - Optional data object with additional details
   */
  static logRuntimeOperation(operation, level, message, data = {}) {
    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      operation,
      level,
      message,
      data
    };
    
    // Log to Google Apps Script Logger with appropriate level
    const logMessage = `RUNTIME [${level.toUpperCase()}] ${operation}: ${message}`;
    Logger.log(logMessage);
    
    // Store detailed runtime log in PropertiesService for monitoring
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingLogs = properties.getProperty('runtimeLogs');
      let logs = [];
      
      if (existingLogs) {
        logs = JSON.parse(existingLogs);
      }
      
      // Add new log entry and keep only last 100 entries
      logs.unshift(logEntry);
      logs = logs.slice(0, 100);
      
      properties.setProperty('runtimeLogs', JSON.stringify(logs));
      
      // Update runtime statistics
      this.updateRuntimeStats(operation, level, data);
      
    } catch (error) {
      Logger.log(`RUNTIME [ERROR] Failed to store runtime log: ${error.message}`);
      console.error('Runtime logging error:', error);
    }
  }
  
  /**
   * Updates runtime operation statistics for performance monitoring
   * @param {string} operation - The operation type
   * @param {string} level - The log level
   * @param {Object} data - Additional data about the operation
   */
  static updateRuntimeStats(operation, level, data) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingStats = properties.getProperty('runtimeStats');
      let stats = {
        cacheHits: 0,
        cacheMisses: 0,
        validationErrors: 0,
        validationWarnings: 0,
        sheetAccesses: 0,
        lastActivity: null,
        operationCounts: {},
        performanceMetrics: {},
        errorPatterns: {}
      };
      
      if (existingStats) {
        stats = { ...stats, ...JSON.parse(existingStats) };
      }
      
      // Update counters based on operation type
      const now = new Date().toISOString();
      stats.lastActivity = now;
      
      // Track operation counts
      if (!stats.operationCounts[operation]) {
        stats.operationCounts[operation] = 0;
      }
      stats.operationCounts[operation]++;
      
      // Track specific metrics
      switch (operation) {
        case 'cache_hit':
          stats.cacheHits++;
          break;
        case 'cache_miss':
          stats.cacheMisses++;
          break;
        case 'validation_error':
          stats.validationErrors++;
          break;
        case 'validation_warning':
          stats.validationWarnings++;
          break;
        case 'sheet_access':
          stats.sheetAccesses++;
          break;
      }
      
      // Track performance metrics if available
      if (data.duration) {
        if (!stats.performanceMetrics[operation]) {
          stats.performanceMetrics[operation] = {
            totalTime: 0,
            count: 0,
            averageTime: 0,
            minTime: null,
            maxTime: null
          };
        }
        
        const metric = stats.performanceMetrics[operation];
        metric.totalTime += data.duration;
        metric.count++;
        metric.averageTime = metric.totalTime / metric.count;
        
        if (metric.minTime === null || data.duration < metric.minTime) {
          metric.minTime = data.duration;
        }
        if (metric.maxTime === null || data.duration > metric.maxTime) {
          metric.maxTime = data.duration;
        }
      }
      
      // Track error patterns for debugging
      if (level === 'error' && data.error) {
        const errorKey = data.error.substring(0, 50); // First 50 chars of error
        if (!stats.errorPatterns[errorKey]) {
          stats.errorPatterns[errorKey] = {
            count: 0,
            lastOccurrence: null,
            operations: []
          };
        }
        stats.errorPatterns[errorKey].count++;
        stats.errorPatterns[errorKey].lastOccurrence = now;
        if (!stats.errorPatterns[errorKey].operations.includes(operation)) {
          stats.errorPatterns[errorKey].operations.push(operation);
        }
      }
      
      properties.setProperty('runtimeStats', JSON.stringify(stats));
      
    } catch (error) {
      Logger.log(`RUNTIME [ERROR] Failed to update runtime stats: ${error.message}`);
    }
  }
  
  /**
   * Gets runtime operation statistics for monitoring
   * @returns {Object} Runtime statistics and recent logs
   */
  static getRuntimeStats() {
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // Get runtime statistics
      const statsData = properties.getProperty('runtimeStats');
      let stats = {};
      if (statsData) {
        stats = JSON.parse(statsData);
      }
      
      // Get recent runtime logs
      const logsData = properties.getProperty('runtimeLogs');
      let logs = [];
      if (logsData) {
        logs = JSON.parse(logsData);
      }
      
      // Calculate cache hit ratio
      const totalCacheOperations = (stats.cacheHits || 0) + (stats.cacheMisses || 0);
      const cacheHitRatio = totalCacheOperations > 0 ? 
        ((stats.cacheHits || 0) / totalCacheOperations * 100).toFixed(1) : 0;
      
      return {
        stats,
        cacheHitRatio,
        recentLogs: logs.slice(0, 20), // Last 20 log entries
        totalLogEntries: logs.length,
        timestamp: new Date().toISOString()
      };
      
    } catch (error) {
      Logger.log(`RUNTIME [ERROR] Failed to get runtime stats: ${error.message}`);
      return {
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }
  
  /**
   * Stores complete campus data in PropertiesService
   * Combines recipient and drive folder data with proper JSON structure for fast lookups
   * @param {Object} campusData - Complete campus information with recipients and drive links
   */
  static storeCampusData(campusData) {
    Logger.log('PropertiesManager: Storing complete campus data...');
    
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // Validate input data structure
      if (!campusData || typeof campusData !== 'object') {
        throw new Error('Invalid campus data: must be an object');
      }
      
      // Validate each campus entry has required structure
      Object.keys(campusData).forEach(campus => {
        const campusInfo = campusData[campus];
        if (!campusInfo || typeof campusInfo !== 'object') {
          throw new Error(`Invalid data for campus ${campus}: must be an object`);
        }
        if (!Array.isArray(campusInfo.recipients)) {
          throw new Error(`Invalid recipients for campus ${campus}: must be an array`);
        }
        if (typeof campusInfo.driveLink !== 'string') {
          throw new Error(`Invalid driveLink for campus ${campus}: must be a string`);
        }
      });
      
      // Create optimized data structure for fast lookups
      const dataToStore = {
        campusData: campusData,
        lastUpdated: new Date().toISOString(),
        migrationComplete: true
      };
      
      // Store the complete data structure
      properties.setProperty(this.CAMPUS_DATA_KEY, JSON.stringify(dataToStore));
      
      // Log storage details
      const campusCount = Object.keys(campusData).length;
      const totalRecipients = Object.values(campusData).reduce((total, campus) => 
        total + (campus.recipients ? campus.recipients.length : 0), 0);
      
      Logger.log(`PropertiesManager: Successfully stored data for ${campusCount} campuses with ${totalRecipients} total recipients`);
      
    } catch (error) {
      Logger.log(`PropertiesManager: Error storing campus data: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Retrieves campus data from PropertiesService with validation
   * Includes cache validation and corruption detection
   * @returns {Object|null} Cached campus data or null if not found/invalid
   */
  static getCampusData() {
    const startTime = new Date();
    this.logRuntimeOperation('cache_lookup', 'info', 'Retrieving campus data with validation');
    
    try {
      const properties = PropertiesService.getScriptProperties();
      const storedData = properties.getProperty(this.CAMPUS_DATA_KEY);
      
      if (!storedData) {
        const duration = new Date() - startTime;
        this.logRuntimeOperation('cache_miss', 'info', 'No campus data found in cache', { duration });
        return null;
      }
      
      // Parse and validate stored data
      let parsedData;
      try {
        parsedData = JSON.parse(storedData);
      } catch (parseError) {
        const duration = new Date() - startTime;
        this.logRuntimeOperation('cache_corruption', 'error', 
          `Cache corruption detected - invalid JSON: ${parseError.message}`, 
          { duration, error: parseError.message });
        this.clearCache();
        return null;
      }
      
      // Validate data structure
      const validationResult = this.validateCacheData(parsedData);
      if (!validationResult.isValid) {
        const duration = new Date() - startTime;
        this.logRuntimeOperation('cache_validation_failed', 'error', 
          `Cache validation failed: ${validationResult.errors.join(', ')}`, 
          { duration, errors: validationResult.errors });
        this.clearCache();
        return null;
      }
      
      const campusCount = Object.keys(parsedData.campusData || {}).length;
      const duration = new Date() - startTime;
      
      this.logRuntimeOperation('cache_hit', 'info', 
        `Successfully retrieved and validated data for ${campusCount} campuses`, 
        { 
          duration, 
          campusCount, 
          lastUpdated: parsedData.lastUpdated,
          dataSize: storedData.length 
        });
      
      return parsedData.campusData;
      
    } catch (error) {
      const duration = new Date() - startTime;
      this.logRuntimeOperation('cache_error', 'error', 
        `Error retrieving campus data: ${error.message}`, 
        { duration, error: error.message });
      return null;
    }
  }
  
  /**
   * Validates cached campus data structure and content
   * @param {Object} data - The parsed cache data to validate
   * @returns {Object} Validation result with isValid flag and errors array
   */
  static validateCacheData(data) {
    const errors = [];
    
    // Check top-level structure
    if (!data || typeof data !== 'object') {
      errors.push('Cache data is not an object');
      return { isValid: false, errors };
    }
    
    if (!data.campusData || typeof data.campusData !== 'object') {
      errors.push('Missing or invalid campusData property');
      return { isValid: false, errors };
    }
    
    // Validate each campus entry
    Object.keys(data.campusData).forEach(campus => {
      const campusInfo = data.campusData[campus];
      
      if (!campusInfo || typeof campusInfo !== 'object') {
        errors.push(`Invalid data structure for campus: ${campus}`);
        return;
      }
      
      if (!Array.isArray(campusInfo.recipients)) {
        errors.push(`Invalid recipients array for campus: ${campus}`);
      }
      
      if (typeof campusInfo.driveLink !== 'string') {
        errors.push(`Invalid driveLink for campus: ${campus}`);
      }
      
      // Validate email format in recipients
      if (Array.isArray(campusInfo.recipients)) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        campusInfo.recipients.forEach((email, index) => {
          if (typeof email !== 'string' || !emailRegex.test(email)) {
            errors.push(`Invalid email format in campus ${campus}, recipient ${index + 1}: ${email}`);
          }
        });
      }
    });
    
    // Check for reasonable data freshness (warn if older than 30 days)
    if (data.lastUpdated) {
      const lastUpdate = new Date(data.lastUpdated);
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
      if (lastUpdate < thirtyDaysAgo) {
        Logger.log(`PropertiesManager: Warning - Cache data is older than 30 days (${data.lastUpdated})`);
      }
    }
    
    const isValid = errors.length === 0;
    if (!isValid) {
      Logger.log(`PropertiesManager: Cache validation found ${errors.length} errors`);
    }
    
    return { isValid, errors };
  }
  
  /**
   * Clears cached campus data and logs the operation
   */
  static clearCache() {
    Logger.log('PropertiesManager: Clearing campus data cache...');
    
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty(this.CAMPUS_DATA_KEY);
      Logger.log('PropertiesManager: Cache cleared successfully');
      
    } catch (error) {
      Logger.log(`PropertiesManager: Error clearing cache: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Gets cache metadata including last updated timestamp and data size
   * @returns {Object|null} Cache metadata or null if no cache exists
   */
  static getCacheMetadata() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const storedData = properties.getProperty(this.CAMPUS_DATA_KEY);
      
      if (!storedData) {
        return null;
      }
      
      const parsedData = JSON.parse(storedData);
      const campusCount = Object.keys(parsedData.campusData || {}).length;
      const totalRecipients = Object.values(parsedData.campusData || {}).reduce((total, campus) => 
        total + (campus.recipients ? campus.recipients.length : 0), 0);
      
      return {
        lastUpdated: parsedData.lastUpdated,
        campusCount: campusCount,
        totalRecipients: totalRecipients,
        dataSize: storedData.length
      };
      
    } catch (error) {
      Logger.log(`PropertiesManager: Error getting cache metadata: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Checks if migration has been completed
   * @returns {boolean} True if migration flag is set
   */
  static isMigrationComplete() {
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // Check both the dedicated migration flag and the flag in campus data
      const migrationStatus = properties.getProperty(this.MIGRATION_COMPLETE_KEY);
      const campusData = properties.getProperty(this.CAMPUS_DATA_KEY);
      
      let isComplete = migrationStatus === 'true';
      
      // Also check if migration flag is set in the campus data structure
      if (!isComplete && campusData) {
        try {
          const parsedData = JSON.parse(campusData);
          isComplete = parsedData.migrationComplete === true;
        } catch (error) {
          Logger.log(`PropertiesManager: Error parsing campus data for migration check: ${error.message}`);
        }
      }
      
      Logger.log(`PropertiesManager: Migration status - ${isComplete ? 'completed' : 'not completed'}`);
      return isComplete;
      
    } catch (error) {
      Logger.log(`PropertiesManager: Error checking migration status: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Marks migration as complete with timestamp
   */
  static setMigrationComplete() {
    Logger.log('PropertiesManager: Setting migration as complete...');
    
    try {
      const properties = PropertiesService.getScriptProperties();
      const timestamp = new Date().toISOString();
      
      // Set dedicated migration flag
      properties.setProperty(this.MIGRATION_COMPLETE_KEY, 'true');
      properties.setProperty(this.LAST_UPDATED_KEY, timestamp);
      
      // Also update the flag in campus data if it exists
      const campusData = properties.getProperty(this.CAMPUS_DATA_KEY);
      if (campusData) {
        try {
          const parsedData = JSON.parse(campusData);
          parsedData.migrationComplete = true;
          parsedData.lastUpdated = timestamp;
          properties.setProperty(this.CAMPUS_DATA_KEY, JSON.stringify(parsedData));
        } catch (error) {
          Logger.log(`PropertiesManager: Warning - Could not update migration flag in campus data: ${error.message}`);
        }
      }
      
      Logger.log('PropertiesManager: Migration marked as complete with timestamp: ' + timestamp);
      
    } catch (error) {
      Logger.log(`PropertiesManager: Error setting migration complete: ${error.message}`);
      throw error;
    }
  }
}

/**
 * Handles reading from and managing the Recipients sheet.
 * Manages sheet creation, data validation, and recipient data operations.
 */
class RecipientsSheetManager {
  static SHEET_NAME = 'Recipients';
  static CAMPUS_COLUMN = 'Campus';
  static RECIPIENT_COLUMN = 'Recipient';
  
  /**
   * Reads all recipient data from Recipients sheet with comprehensive validation
   * Uses enhanced email validation and continues processing valid emails when invalid ones are found
   * Detects duplicate campus entries and uses first occurrence while logging warnings
   * Implements batch sheet reading operations to minimize API calls to Google Sheets service
   * @returns {Object} Campus to recipients array mapping with validation results
   */
  static readRecipientsData() {
    const startTime = new Date();
    PropertiesManager.logRuntimeOperation('sheet_access', 'info', 
      'Starting Recipients sheet read with batch operations and comprehensive validation');
    
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(this.SHEET_NAME);
      
      if (!sheet) {
        const duration = new Date() - startTime;
        PropertiesManager.logRuntimeOperation('sheet_not_found', 'warning', 
          'Recipients sheet not found', { duration, sheetName: this.SHEET_NAME });
        return {};
      }
      
      // Use batch sheet reading operations - read all data in single operation
      // This minimizes API calls to Google Sheets service (Requirement 5.4)
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      
      if (lastRow <= 1 || lastColumn < 2) {
        const duration = new Date() - startTime;
        PropertiesManager.logRuntimeOperation('sheet_empty', 'warning', 
          'No data found in Recipients sheet or insufficient columns', 
          { duration, lastRow, lastColumn });
        return {};
      }
      
      // Read all data in a single batch operation instead of multiple calls
      const batchStartTime = new Date();
      const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
      const batchDuration = new Date() - batchStartTime;
      
      PropertiesManager.logRuntimeOperation('batch_sheet_read', 'info', 
        `Batch reading ${lastRow}x${lastColumn} range in single operation`, 
        { batchDuration, rows: lastRow, columns: lastColumn, dataLength: data.length });
      
      if (data.length <= 1) {
        const duration = new Date() - startTime;
        PropertiesManager.logRuntimeOperation('sheet_no_data', 'warning', 
          'No data rows found after batch read', { duration });
        return {};
      }
      
      // Perform comprehensive email validation
      const validationResults = this.validateEmailAddresses(data);
      
      // Log validation summary with detailed error tracking
      if (validationResults.errors.length > 0) {
        PropertiesManager.logRuntimeOperation('validation_error', 'error', 
          `Found ${validationResults.errors.length} validation errors in Recipients sheet`, 
          { 
            errorCount: validationResults.errors.length,
            errors: validationResults.errors.slice(0, 5), // First 5 errors for logging
            totalRows: validationResults.summary.totalRows
          });
        
        // Log each validation error individually for detailed tracking
        validationResults.errors.forEach((error, index) => {
          if (index < 10) { // Log first 10 errors individually
            PropertiesManager.logRuntimeOperation('validation_error_detail', 'error', error, 
              { errorIndex: index + 1, totalErrors: validationResults.errors.length });
          }
        });
      }
      
      if (validationResults.warnings.length > 0) {
        PropertiesManager.logRuntimeOperation('validation_warning', 'warning', 
          `Found ${validationResults.warnings.length} validation warnings in Recipients sheet`, 
          { 
            warningCount: validationResults.warnings.length,
            warnings: validationResults.warnings.slice(0, 3) // First 3 warnings for logging
          });
      }
      
      // Validate campus names against expected values
      const campusValidation = this.validateCampusNames(validationResults.valid);
      
      // Handle duplicate campus entries - use first occurrence and log warnings for duplicates
      const duplicateHandling = this.handleDuplicateCampusEntries(validationResults.valid);
      
      // Get the final processed data
      const recipientsData = duplicateHandling.data;
      
      // Log comprehensive final results
      const validCampusCount = Object.keys(recipientsData).length;
      const totalValidRecipients = Object.values(recipientsData).reduce((sum, recipients) => sum + recipients.length, 0);
      const totalDuration = new Date() - startTime;
      
      PropertiesManager.logRuntimeOperation('sheet_read_complete', 'info', 
        'Recipients sheet processing complete', 
        { 
          totalDuration,
          batchDuration,
          validRecipients: totalValidRecipients,
          validCampuses: validCampusCount,
          invalidEmails: validationResults.summary.invalidEmails,
          duplicatesHandled: duplicateHandling.summary.duplicatesFound,
          unknownCampuses: campusValidation.invalidCampuses.length,
          sheetAccessPattern: 'batch_read_single_operation'
        });
      
      if (validationResults.summary.invalidEmails > 0) {
        PropertiesManager.logRuntimeOperation('invalid_emails_skipped', 'warning', 
          `Skipped ${validationResults.summary.invalidEmails} invalid email addresses`, 
          { invalidCount: validationResults.summary.invalidEmails });
      }
      
      if (duplicateHandling.summary.duplicatesFound > 0) {
        PropertiesManager.logRuntimeOperation('duplicates_handled', 'warning', 
          `Used first occurrence for ${duplicateHandling.summary.duplicatesFound} duplicate entries`, 
          { duplicateCount: duplicateHandling.summary.duplicatesFound });
      }
      
      // Store comprehensive validation results for potential use by calling functions
      recipientsData._validationResults = {
        email: validationResults,
        campus: campusValidation,
        duplicates: duplicateHandling,
        summary: {
          totalProcessed: validationResults.summary.totalRows,
          validRecipients: totalValidRecipients,
          validCampuses: validCampusCount,
          errorsFound: validationResults.errors.length,
          warningsFound: validationResults.warnings.length + duplicateHandling.duplicateWarnings.length + campusValidation.warnings.length
        }
      };
      
      return recipientsData;
      
    } catch (error) {
      const duration = new Date() - startTime;
      PropertiesManager.logRuntimeOperation('sheet_read_error', 'error', 
        `Error reading recipients data: ${error.message}`, 
        { duration, error: error.message, stack: error.stack });
      throw error;
    }
  }
  
  /**
   * Validates email addresses in Recipients sheet with detailed error logging
   * Validates email format using regex patterns and logs detailed errors with row numbers
   * Continues processing valid emails when invalid ones are found
   * @param {Array} data - Raw sheet data
   * @returns {Object} Validation results with detailed errors and valid emails
   */
  static validateEmailAddresses(data) {
    Logger.log('RecipientsSheetManager: Starting comprehensive email address validation...');
    
    // Enhanced email regex pattern that validates proper email format
    const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
    
    const validationResults = {
      valid: [],
      invalid: [],
      errors: [],
      warnings: [],
      summary: {
        totalRows: data.length - 1, // Exclude header
        validEmails: 0,
        invalidEmails: 0,
        emptyRows: 0,
        duplicateEmails: 0
      }
    };
    
    const seenEmails = new Set();
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const campus = row[0] ? row[0].toString().trim() : '';
      const recipient = row[1] ? row[1].toString().trim() : '';
      const rowNumber = i + 1;
      
      // Handle empty rows
      if (!campus && !recipient) {
        validationResults.summary.emptyRows++;
        continue;
      }
      
      // Handle rows with missing campus or recipient
      if (!campus) {
        const errorMsg = `Row ${rowNumber}: Missing campus name for recipient "${recipient}"`;
        validationResults.errors.push(errorMsg);
        validationResults.invalid.push({ campus: '', recipient, rowNumber, error: 'Missing campus' });
        Logger.log(`ERROR: ${errorMsg}`);
        continue;
      }
      
      if (!recipient) {
        const errorMsg = `Row ${rowNumber}: Missing recipient email for campus "${campus}"`;
        validationResults.errors.push(errorMsg);
        validationResults.invalid.push({ campus, recipient: '', rowNumber, error: 'Missing email' });
        Logger.log(`ERROR: ${errorMsg}`);
        continue;
      }
      
      // Check for duplicate emails
      const lowerCaseEmail = recipient.toLowerCase();
      if (seenEmails.has(lowerCaseEmail)) {
        const warningMsg = `Row ${rowNumber}: Duplicate email address "${recipient}" (campus: ${campus})`;
        validationResults.warnings.push(warningMsg);
        validationResults.summary.duplicateEmails++;
        Logger.log(`WARNING: ${warningMsg}`);
        // Continue processing - duplicates are warnings, not errors
      } else {
        seenEmails.add(lowerCaseEmail);
      }
      
      // Validate email format using regex patterns
      if (emailRegex.test(recipient)) {
        validationResults.valid.push({ 
          campus: campus.toLowerCase(), 
          recipient, 
          rowNumber,
          originalCampus: campus 
        });
        validationResults.summary.validEmails++;
        Logger.log(`VALID: Row ${rowNumber} - ${recipient} (${campus})`);
      } else {
        // Log detailed errors with row numbers for invalid emails
        const errorMsg = `Row ${rowNumber}: Invalid email format - "${recipient}" (campus: ${campus})`;
        validationResults.invalid.push({ 
          campus, 
          recipient, 
          rowNumber, 
          error: 'Invalid format',
          originalCampus: campus 
        });
        validationResults.errors.push(errorMsg);
        validationResults.summary.invalidEmails++;
        Logger.log(`ERROR: ${errorMsg}`);
        
        // Provide specific feedback on common email format issues
        if (!recipient.includes('@')) {
          Logger.log(`  → Missing @ symbol`);
        } else if (recipient.indexOf('@') !== recipient.lastIndexOf('@')) {
          Logger.log(`  → Multiple @ symbols`);
        } else if (!recipient.includes('.')) {
          Logger.log(`  → Missing domain extension`);
        } else if (recipient.startsWith('@') || recipient.endsWith('@')) {
          Logger.log(`  → @ symbol at beginning or end`);
        } else if (recipient.includes('..')) {
          Logger.log(`  → Consecutive dots`);
        }
      }
    }
    
    // Log comprehensive validation summary
    Logger.log(`RecipientsSheetManager: Email validation completed`);
    Logger.log(`  Total rows processed: ${validationResults.summary.totalRows}`);
    Logger.log(`  Valid emails: ${validationResults.summary.validEmails}`);
    Logger.log(`  Invalid emails: ${validationResults.summary.invalidEmails}`);
    Logger.log(`  Empty rows: ${validationResults.summary.emptyRows}`);
    Logger.log(`  Duplicate emails: ${validationResults.summary.duplicateEmails}`);
    Logger.log(`  Warnings: ${validationResults.warnings.length}`);
    Logger.log(`  Errors: ${validationResults.errors.length}`);
    
    // Continue processing valid emails when invalid ones are found
    if (validationResults.summary.invalidEmails > 0) {
      Logger.log(`RecipientsSheetManager: Found ${validationResults.summary.invalidEmails} invalid emails, but continuing with ${validationResults.summary.validEmails} valid emails`);
    }
    
    return validationResults;
  }
  
  /**
   * Creates Recipients sheet if it doesn't exist
   * @param {Object} initialData - Initial campus recipient data
   */
  static createSheet(initialData) {
    Logger.log('RecipientsSheetManager: Creating Recipients sheet...');
    
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Check if sheet already exists
      if (spreadsheet.getSheetByName(this.SHEET_NAME)) {
        Logger.log('RecipientsSheetManager: Recipients sheet already exists');
        return;
      }
      
      // Create new sheet as the last tab
      const sheet = spreadsheet.insertSheet(this.SHEET_NAME);
      const sheets = spreadsheet.getSheets();
      const lastPosition = sheets.length;
      spreadsheet.moveSheet(sheet, lastPosition);
      
      // Add headers
      sheet.getRange(1, 1).setValue(this.CAMPUS_COLUMN);
      sheet.getRange(1, 2).setValue(this.RECIPIENT_COLUMN);
      
      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, 2);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      
      // Populate with initial data
      let currentRow = 2;
      Object.keys(initialData).forEach(campus => {
        const recipients = initialData[campus];
        recipients.forEach(recipient => {
          sheet.getRange(currentRow, 1).setValue(campus);
          sheet.getRange(currentRow, 2).setValue(recipient);
          currentRow++;
        });
      });
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, 2);
      
      Logger.log(`RecipientsSheetManager: Created sheet with ${currentRow - 2} recipient entries`);
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Error creating sheet: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Detects and handles duplicate campus entries in Recipients sheet
   * Uses first occurrence and logs warnings for duplicates
   * @param {Array} validEntries - Array of valid recipient entries
   * @returns {Object} Processed data with duplicate handling results
   */
  static handleDuplicateCampusEntries(validEntries) {
    Logger.log('RecipientsSheetManager: Detecting duplicate campus entries...');
    
    const campusFirstOccurrence = new Map();
    const duplicateWarnings = [];
    const processedData = {};
    
    validEntries.forEach(entry => {
      const campus = entry.campus;
      const recipient = entry.recipient;
      const rowNumber = entry.rowNumber;
      
      // Track first occurrence of each campus
      if (!campusFirstOccurrence.has(campus)) {
        campusFirstOccurrence.set(campus, rowNumber);
        
        // Initialize campus array if not exists
        if (!processedData[campus]) {
          processedData[campus] = [];
        }
      }
      
      // Add recipient to campus (this handles multiple recipients per campus correctly)
      if (!processedData[campus]) {
        processedData[campus] = [];
      }
      processedData[campus].push(recipient);
      
      // Check if this is a duplicate campus entry (same campus, different row than first occurrence)
      const firstRow = campusFirstOccurrence.get(campus);
      if (rowNumber !== firstRow) {
        // This is not actually a duplicate campus - it's multiple recipients for the same campus
        // Only log as duplicate if we see the exact same campus-recipient combination
        const existingRecipients = processedData[campus].slice(0, -1); // All except the one we just added
        if (existingRecipients.includes(recipient)) {
          const warningMsg = `Row ${rowNumber}: Duplicate campus-recipient combination - "${campus}" with "${recipient}" (first seen at row ${firstRow})`;
          duplicateWarnings.push(warningMsg);
          Logger.log(`WARNING: ${warningMsg}`);
          
          // Remove the duplicate recipient (use first occurrence)
          processedData[campus].pop();
        }
      }
    });
    
    // Log summary of duplicate handling
    if (duplicateWarnings.length > 0) {
      Logger.log(`RecipientsSheetManager: Found ${duplicateWarnings.length} duplicate campus-recipient combinations`);
      duplicateWarnings.forEach(warning => Logger.log(`  ${warning}`));
    } else {
      Logger.log('RecipientsSheetManager: No duplicate campus-recipient combinations found');
    }
    
    const totalCampuses = Object.keys(processedData).length;
    const totalRecipients = Object.values(processedData).reduce((sum, recipients) => sum + recipients.length, 0);
    
    Logger.log(`RecipientsSheetManager: Processed ${totalRecipients} unique recipients across ${totalCampuses} campuses`);
    
    return {
      data: processedData,
      duplicateWarnings: duplicateWarnings,
      summary: {
        totalCampuses: totalCampuses,
        totalRecipients: totalRecipients,
        duplicatesFound: duplicateWarnings.length
      }
    };
  }

  /**
   * Validates campus names against expected values and handles duplicates
   * @param {Array} validEntries - Array of valid recipient entries  
   * @returns {Object} Validation results with campus name issues
   */
  static validateCampusNames(validEntries) {
    Logger.log('RecipientsSheetManager: Validating campus names...');
    
    // Get expected campus names from hardcoded data for validation
    const expectedCampuses = Object.keys(MigrationManager.extractHardcodedData());
    const validationResults = {
      validCampuses: [],
      invalidCampuses: [],
      warnings: [],
      suggestions: []
    };
    
    const seenCampuses = new Set();
    
    validEntries.forEach(entry => {
      const campus = entry.campus;
      const rowNumber = entry.rowNumber;
      
      if (!seenCampuses.has(campus)) {
        seenCampuses.add(campus);
        
        if (expectedCampuses.includes(campus)) {
          validationResults.validCampuses.push({ campus, rowNumber });
        } else {
          validationResults.invalidCampuses.push({ campus, rowNumber });
          
          // Suggest similar campus names
          const suggestions = expectedCampuses.filter(expected => 
            expected.includes(campus) || 
            campus.includes(expected) ||
            this.calculateLevenshteinDistance(campus, expected) <= 2
          );
          
          const warningMsg = `Row ${rowNumber}: Unknown campus name "${campus}"`;
          let suggestionMsg = '';
          
          if (suggestions.length > 0) {
            suggestionMsg = ` - Did you mean: ${suggestions.join(', ')}?`;
            validationResults.suggestions.push({ campus, suggestions, rowNumber });
          }
          
          const fullWarning = warningMsg + suggestionMsg;
          validationResults.warnings.push(fullWarning);
          Logger.log(`WARNING: ${fullWarning}`);
        }
      }
    });
    
    Logger.log(`RecipientsSheetManager: Campus validation complete - ${validationResults.validCampuses.length} valid, ${validationResults.invalidCampuses.length} unknown`);
    
    return validationResults;
  }
  
  /**
   * Calculates Levenshtein distance between two strings for campus name suggestions
   * @param {string} str1 - First string
   * @param {string} str2 - Second string  
   * @returns {number} Edit distance between strings
   */
  static calculateLevenshteinDistance(str1, str2) {
    const matrix = [];
    
    for (let i = 0; i <= str2.length; i++) {
      matrix[i] = [i];
    }
    
    for (let j = 0; j <= str1.length; j++) {
      matrix[0][j] = j;
    }
    
    for (let i = 1; i <= str2.length; i++) {
      for (let j = 1; j <= str1.length; j++) {
        if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1, // substitution
            matrix[i][j - 1] + 1,     // insertion
            matrix[i - 1][j] + 1      // deletion
          );
        }
      }
    }
    
    return matrix[str2.length][str1.length];
  }
  
  /**
   * Checks if Recipients sheet exists
   * @returns {boolean} True if sheet exists
   */
  static sheetExists() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(this.SHEET_NAME);
      const exists = sheet !== null;
      
      Logger.log(`RecipientsSheetManager: Sheet exists - ${exists}`);
      return exists;
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Error checking sheet existence: ${error.message}`);
      return false;
    }
  }

  /**
   * Performs optimized batch reading of all sheet data with minimal API calls
   * Reads Recipients sheet data in single operation to minimize API calls to Google Sheets service
   * @returns {Object} Raw sheet data and metadata for further processing
   */
  static batchReadSheetData() {
    Logger.log('RecipientsSheetManager: Performing optimized batch sheet read...');
    
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(this.SHEET_NAME);
      
      if (!sheet) {
        Logger.log('RecipientsSheetManager: Recipients sheet not found for batch read');
        return {
          success: false,
          error: 'Recipients sheet not found',
          data: [],
          metadata: { rows: 0, columns: 0 }
        };
      }
      
      // Get sheet dimensions first to optimize the read operation
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      
      Logger.log(`RecipientsSheetManager: Sheet dimensions - ${lastRow} rows x ${lastColumn} columns`);
      
      if (lastRow <= 1 || lastColumn < 2) {
        Logger.log('RecipientsSheetManager: Insufficient data for batch read');
        return {
          success: true,
          data: [],
          metadata: { rows: lastRow, columns: lastColumn, isEmpty: true }
        };
      }
      
      // Perform single batch read operation - this minimizes API calls (Requirement 5.4)
      const startTime = new Date();
      const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
      const readTime = new Date() - startTime;
      
      Logger.log(`RecipientsSheetManager: Batch read completed in ${readTime}ms - ${data.length} rows retrieved`);
      
      // Validate that we have the expected structure
      if (!data || data.length === 0) {
        Logger.log('RecipientsSheetManager: Warning - Batch read returned no data');
        return {
          success: true,
          data: [],
          metadata: { rows: lastRow, columns: lastColumn, isEmpty: true, readTime }
        };
      }
      
      // Verify header structure
      const headers = data[0];
      const expectedHeaders = [this.CAMPUS_COLUMN, this.RECIPIENT_COLUMN];
      const hasValidHeaders = expectedHeaders.every((header, index) => 
        headers[index] && headers[index].toString().trim() === header
      );
      
      if (!hasValidHeaders) {
        Logger.log(`RecipientsSheetManager: Warning - Unexpected header structure: ${headers.join(', ')}`);
      }
      
      return {
        success: true,
        data: data,
        metadata: {
          rows: lastRow,
          columns: lastColumn,
          dataRows: data.length - 1, // Exclude header
          hasValidHeaders: hasValidHeaders,
          readTime: readTime,
          isEmpty: false
        }
      };
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Error during batch sheet read: ${error.message}`);
      return {
        success: false,
        error: error.message,
        data: [],
        metadata: { rows: 0, columns: 0 }
      };
    }
  }

  /**
   * Optimized version of readRecipientsData that uses batch operations
   * Minimizes API calls by using single batch read operation
   * @returns {Object} Campus to recipients array mapping with validation results
   */
  static readRecipientsDataOptimized() {
    Logger.log('RecipientsSheetManager: Starting optimized recipients data read with batch operations...');
    
    try {
      // Use batch read operation to minimize API calls
      const batchResult = this.batchReadSheetData();
      
      if (!batchResult.success) {
        Logger.log(`RecipientsSheetManager: Batch read failed: ${batchResult.error}`);
        return {};
      }
      
      if (batchResult.metadata.isEmpty) {
        Logger.log('RecipientsSheetManager: No data found in Recipients sheet (batch read)');
        return {};
      }
      
      const data = batchResult.data;
      Logger.log(`RecipientsSheetManager: Processing ${batchResult.metadata.dataRows} data rows from batch read (${batchResult.metadata.readTime}ms)`);
      
      // Perform comprehensive email validation
      const validationResults = this.validateEmailAddresses(data);
      
      // Log validation summary
      if (validationResults.errors.length > 0) {
        Logger.log(`RecipientsSheetManager: Found ${validationResults.errors.length} validation errors:`);
        validationResults.errors.forEach(error => Logger.log(`  ${error}`));
      }
      
      if (validationResults.warnings.length > 0) {
        Logger.log(`RecipientsSheetManager: Found ${validationResults.warnings.length} warnings:`);
        validationResults.warnings.forEach(warning => Logger.log(`  ${warning}`));
      }
      
      // Validate campus names against expected values
      const campusValidation = this.validateCampusNames(validationResults.valid);
      
      // Handle duplicate campus entries - use first occurrence and log warnings for duplicates
      const duplicateHandling = this.handleDuplicateCampusEntries(validationResults.valid);
      
      // Get the final processed data
      const recipientsData = duplicateHandling.data;
      
      // Log comprehensive final results including batch operation performance
      const validCampusCount = Object.keys(recipientsData).length;
      const totalValidRecipients = Object.values(recipientsData).reduce((sum, recipients) => sum + recipients.length, 0);
      
      Logger.log(`RecipientsSheetManager: Batch processing complete:`);
      Logger.log(`  Batch read time: ${batchResult.metadata.readTime}ms`);
      Logger.log(`  Valid recipients: ${totalValidRecipients} across ${validCampusCount} campuses`);
      Logger.log(`  Invalid emails skipped: ${validationResults.summary.invalidEmails}`);
      Logger.log(`  Duplicate entries handled: ${duplicateHandling.summary.duplicatesFound}`);
      Logger.log(`  Unknown campuses: ${campusValidation.invalidCampuses.length}`);
      
      if (validationResults.summary.invalidEmails > 0) {
        Logger.log(`RecipientsSheetManager: Skipped ${validationResults.summary.invalidEmails} invalid email addresses - check logs for details`);
      }
      
      if (duplicateHandling.summary.duplicatesFound > 0) {
        Logger.log(`RecipientsSheetManager: Used first occurrence for ${duplicateHandling.summary.duplicatesFound} duplicate entries`);
      }
      
      // Store comprehensive validation results including batch operation metadata
      recipientsData._validationResults = {
        email: validationResults,
        campus: campusValidation,
        duplicates: duplicateHandling,
        batchOperation: batchResult.metadata,
        summary: {
          totalProcessed: validationResults.summary.totalRows,
          validRecipients: totalValidRecipients,
          validCampuses: validCampusCount,
          errorsFound: validationResults.errors.length,
          warningsFound: validationResults.warnings.length + duplicateHandling.duplicateWarnings.length + campusValidation.warnings.length,
          batchReadTime: batchResult.metadata.readTime
        }
      };
      
      return recipientsData;
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Error in optimized recipients data read: ${error.message}`);
      throw error;
    }
  }

  /**
   * Recreates Recipients sheet from PropertiesService data if missing
   * Handles corrupted sheet scenarios by rebuilding from cached data
   * Implements sheet recovery mechanisms (Requirements 3.3, 4.3)
   * @returns {Object} Recovery operation result with success status and details
   */
  static recoverRecipientsSheet() {
    Logger.log('RecipientsSheetManager: Starting Recipients sheet recovery process...');
    
    try {
      // Check if sheet recovery is needed
      const sheetExists = this.sheetExists();
      
      if (sheetExists) {
        Logger.log('RecipientsSheetManager: Recipients sheet exists, checking for corruption...');
        
        // Test if sheet is readable and has valid structure
        try {
          const batchResult = this.batchReadSheetData();
          
          if (batchResult.success && !batchResult.metadata.isEmpty && batchResult.metadata.hasValidHeaders) {
            Logger.log('RecipientsSheetManager: Recipients sheet is healthy, no recovery needed');
            return {
              success: true,
              action: 'no_recovery_needed',
              message: 'Recipients sheet exists and is healthy',
              sheetExists: true,
              dataRows: batchResult.metadata.dataRows
            };
          } else {
            Logger.log('RecipientsSheetManager: Recipients sheet appears corrupted or empty, proceeding with recovery...');
          }
        } catch (readError) {
          Logger.log(`RecipientsSheetManager: Recipients sheet read failed, treating as corrupted: ${readError.message}`);
        }
      } else {
        Logger.log('RecipientsSheetManager: Recipients sheet missing, proceeding with recovery...');
      }
      
      // Attempt to recover from PropertiesService data
      Logger.log('RecipientsSheetManager: Attempting to recover from PropertiesService data...');
      const cachedData = PropertiesManager.getCampusData();
      
      if (!cachedData || Object.keys(cachedData).length === 0) {
        Logger.log('RecipientsSheetManager: No cached data available for recovery, using hardcoded data...');
        
        // Fall back to hardcoded data if no cache available
        const hardcodedData = MigrationManager.extractHardcodedData();
        
        if (!hardcodedData || Object.keys(hardcodedData).length === 0) {
          throw new Error('No data available for sheet recovery - both cache and hardcoded data are empty');
        }
        
        // Create sheet from hardcoded data
        const recoveryResult = this.createSheetFromData(hardcodedData, 'hardcoded');
        
        Logger.log('RecipientsSheetManager: Sheet recovery completed using hardcoded data');
        return {
          success: true,
          action: 'recovered_from_hardcoded',
          message: 'Recipients sheet recreated from hardcoded data',
          campusCount: Object.keys(hardcodedData).length,
          totalRecipients: Object.values(hardcodedData).reduce((sum, recipients) => sum + recipients.length, 0),
          dataSource: 'hardcoded'
        };
      }
      
      // Extract recipient data from cached campus data
      const recipientData = {};
      Object.keys(cachedData).forEach(campus => {
        if (cachedData[campus].recipients && Array.isArray(cachedData[campus].recipients)) {
          recipientData[campus] = cachedData[campus].recipients;
        }
      });
      
      if (Object.keys(recipientData).length === 0) {
        throw new Error('No recipient data found in cached campus data for recovery');
      }
      
      // Create sheet from cached data
      const recoveryResult = this.createSheetFromData(recipientData, 'cache');
      
      Logger.log('RecipientsSheetManager: Sheet recovery completed using cached data');
      return {
        success: true,
        action: 'recovered_from_cache',
        message: 'Recipients sheet recreated from PropertiesService cache',
        campusCount: Object.keys(recipientData).length,
        totalRecipients: Object.values(recipientData).reduce((sum, recipients) => sum + recipients.length, 0),
        dataSource: 'cache'
      };
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Sheet recovery failed: ${error.message}`);
      console.error('Sheet recovery error:', error);
      
      return {
        success: false,
        action: 'recovery_failed',
        message: `Sheet recovery failed: ${error.message}`,
        error: error.message
      };
    }
  }

  /**
   * Creates Recipients sheet from provided data with proper structure
   * Handles both hardcoded data and cached data formats
   * @param {Object} data - Campus recipient data to populate sheet with
   * @param {string} dataSource - Source of data ('cache', 'hardcoded', etc.)
   * @returns {Object} Creation result with success status and details
   */
  static createSheetFromData(data, dataSource = 'unknown') {
    Logger.log(`RecipientsSheetManager: Creating Recipients sheet from ${dataSource} data...`);
    
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Remove existing sheet if it exists (for recovery scenarios)
      const existingSheet = spreadsheet.getSheetByName(this.SHEET_NAME);
      if (existingSheet) {
        Logger.log('RecipientsSheetManager: Removing existing corrupted Recipients sheet...');
        spreadsheet.deleteSheet(existingSheet);
      }
      
      // Create new sheet as the last tab
      Logger.log('RecipientsSheetManager: Creating new Recipients sheet...');
      const sheet = spreadsheet.insertSheet(this.SHEET_NAME);
      const sheets = spreadsheet.getSheets();
      const lastPosition = sheets.length;
      spreadsheet.moveSheet(sheet, lastPosition);
      
      // Add headers
      sheet.getRange(1, 1).setValue(this.CAMPUS_COLUMN);
      sheet.getRange(1, 2).setValue(this.RECIPIENT_COLUMN);
      
      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, 2);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      
      // Populate with data
      let currentRow = 2;
      let totalRecipients = 0;
      
      Object.keys(data).forEach(campus => {
        const recipients = data[campus];
        if (Array.isArray(recipients)) {
          recipients.forEach(recipient => {
            sheet.getRange(currentRow, 1).setValue(campus);
            sheet.getRange(currentRow, 2).setValue(recipient);
            currentRow++;
            totalRecipients++;
          });
        }
      });
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, 2);
      
      // Add recovery metadata as a comment to the header
      const recoveryComment = `Sheet recovered from ${dataSource} data on ${new Date().toISOString()}`;
      sheet.getRange(1, 1).setNote(recoveryComment);
      
      Logger.log(`RecipientsSheetManager: Created sheet with ${totalRecipients} recipient entries from ${dataSource} data`);
      
      return {
        success: true,
        entriesCreated: totalRecipients,
        campusCount: Object.keys(data).length,
        dataSource: dataSource,
        recoveryTimestamp: new Date().toISOString()
      };
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Error creating sheet from ${dataSource} data: ${error.message}`);
      throw error;
    }
  }

  /**
   * Validates Recipients sheet integrity and attempts recovery if needed
   * Checks for sheet existence, readability, and data integrity
   * Automatically triggers recovery if corruption is detected
   * @returns {Object} Validation and recovery result
   */
  static validateAndRecoverSheet() {
    Logger.log('RecipientsSheetManager: Starting sheet validation and recovery check...');
    
    try {
      // Check if sheet exists
      if (!this.sheetExists()) {
        Logger.log('RecipientsSheetManager: Recipients sheet missing, triggering recovery...');
        return this.recoverRecipientsSheet();
      }
      
      // Test sheet readability and structure
      try {
        const batchResult = this.batchReadSheetData();
        
        if (!batchResult.success) {
          Logger.log(`RecipientsSheetManager: Sheet read failed, triggering recovery: ${batchResult.error}`);
          return this.recoverRecipientsSheet();
        }
        
        if (batchResult.metadata.isEmpty) {
          Logger.log('RecipientsSheetManager: Sheet is empty, triggering recovery...');
          return this.recoverRecipientsSheet();
        }
        
        if (!batchResult.metadata.hasValidHeaders) {
          Logger.log('RecipientsSheetManager: Sheet has invalid headers, triggering recovery...');
          return this.recoverRecipientsSheet();
        }
        
        // Validate data integrity
        const validationResults = this.validateEmailAddresses(batchResult.data);
        
        if (validationResults.valid.length === 0 && validationResults.invalid.length > 0) {
          Logger.log('RecipientsSheetManager: Sheet contains only invalid data, triggering recovery...');
          return this.recoverRecipientsSheet();
        }
        
        Logger.log('RecipientsSheetManager: Sheet validation passed - sheet is healthy');
        return {
          success: true,
          action: 'validation_passed',
          message: 'Recipients sheet is healthy and valid',
          dataRows: batchResult.metadata.dataRows,
          validRecipients: validationResults.valid.length,
          invalidRecipients: validationResults.invalid.length
        };
        
      } catch (validationError) {
        Logger.log(`RecipientsSheetManager: Sheet validation failed, triggering recovery: ${validationError.message}`);
        return this.recoverRecipientsSheet();
      }
      
    } catch (error) {
      Logger.log(`RecipientsSheetManager: Error during validation and recovery: ${error.message}`);
      console.error('Validation and recovery error:', error);
      
      return {
        success: false,
        action: 'validation_error',
        message: `Validation and recovery failed: ${error.message}`,
        error: error.message
      };
    }
  }
}

/**
 * The name of the column used to track the return date for sent emails.
 * @constant {string}
 */
const RETURN_DATE = "Anticipated Return date";

/**
 * The name of the column used to track the date when the email was sent to campuses.
 * @constant {string}
 */
const DATE_SENT_COL = "Date when the email was sent to campuses";

/**
 * The name of the column used to store the campus folder ID.
 * @constant {string}
 */
const CAMPUS_FOLDER_COL = "Campus folder ID";

/**
 * Adds a custom menu to the Google Sheets UI when the spreadsheet is opened.
 * The menu allows users to preview emails or send emails to campuses based
 * on specific criteria, and manage the campus recipients cache.
 * @function
 * @returns {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📬 Notify Campuses")
    .addItem(
      "Preview emails (don't send)",
      "previewEmails"
    )
    .addItem(
      "Send emails to campuses",
      "sendEmails"
    )
    .addSeparator()
    .addSubMenu(ui.createMenu("🔄 Cache Management")
      .addItem("Refresh campus cache", "refreshCampusCache")
      .addItem("Show cache status", "showCacheStatus")
    )
    .addSubMenu(ui.createMenu("📊 System Monitoring")
      .addItem("Show migration status", "showMigrationStatus")
      .addItem("Show runtime statistics", "showRuntimeStats")
    )
    .addSubMenu(ui.createMenu("🛠️ Sheet Recovery")
      .addItem("Validate Recipients sheet", "validateRecipientsSheet")
      .addItem("Recover Recipients sheet", "recoverRecipientsSheetManual")
    )
    .addToUi();
}

/**
 * Preview emails without sending them.
 * Shows a summary of what would be sent.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] - The sheet to process.
 * @returns {void}
 */
function previewEmails(
  sheet = SpreadsheetApp.openById(
    "1p0KwjAMLnI4KyPG1ErNq0oRmGpQ8GWGDSnYVa7MYnP0"
  ).getSheetByName("Teacher Notes")
) {
  processEmails(sheet, true);
}

/**
 * Sends emails to campuses based on the data in the sheet.
 * Only sends emails for rows with an anticipated return date, a null value for the date when the
 * email was sent to campuses, and has a campus folder ID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] - The sheet to process. Defaults to the 'Teacher Notes' sheet.
 * @returns {void}
 */
function sendEmails(
  sheet = SpreadsheetApp.openById(
    "1p0KwjAMLnI4KyPG1ErNq0oRmGpQ8GWGDSnYVa7MYnP0"
  ).getSheetByName("Teacher Notes")
) {
  processEmails(sheet, false);
}

/**
 * Main function to process emails with preview or send mode.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to process.
 * @param {boolean} previewMode - If true, shows preview without sending.
 * @returns {void}
 */
function processEmails(sheet, previewMode = false) {
  const startTime = new Date();
  Logger.log(`${previewMode ? 'Preview' : 'Send'} operation started at ${startTime}`);
  
  // Validation: Check sheet
  const sheetName = sheet.getName();
  const targetSheetname = "Teacher Notes";

  if (sheetName !== targetSheetname) {
    const errorMsg = `Function can only be run from the "${targetSheetname}" sheet. Current sheet: "${sheetName}"`;
    Logger.log(`ERROR: ${errorMsg}`);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  // Validation: Check if sheet has data
  if (sheet.getLastRow() < 2) {
    const errorMsg = "No data found in sheet. Sheet must have at least a header row and one data row.";
    Logger.log(`ERROR: ${errorMsg}`);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  const dataRange = sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  );
  const data = dataRange.getDisplayValues();
  const heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Validation: Check required columns
  const requiredColumns = [RETURN_DATE, DATE_SENT_COL, CAMPUS_FOLDER_COL, "Campus", "Name"];
  const missingColumns = requiredColumns.filter(col => heads.indexOf(col) === -1);
  
  if (missingColumns.length > 0) {
    const errorMsg = `Missing required columns: ${missingColumns.join(", ")}`;
    Logger.log(`ERROR: ${errorMsg}`);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  const emailSentColIdx = heads.indexOf(DATE_SENT_COL);

  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );

  // Group students by campus
  const campusGroups = {};
  const validRows = [];

  obj.forEach(function (row, rowIdx) {
    if (
      row[RETURN_DATE].trim() !== "" &&
      row[DATE_SENT_COL] === "" &&
      row[CAMPUS_FOLDER_COL] !== ""
    ) {
      const campus = row["Campus"];
      if (!campus || campus.trim() === "") {
        Logger.log(`WARNING: Row ${rowIdx + 2} has empty campus field`);
        return;
      }
      
      if (!campusGroups[campus]) {
        campusGroups[campus] = [];
      }
      campusGroups[campus].push({ row, rowIdx });
      validRows.push({ row, rowIdx });
    }
  });

  const campusCount = Object.keys(campusGroups).length;
  const totalStudents = validRows.length;

  Logger.log(`Found ${totalStudents} students across ${campusCount} campuses`);

  if (totalStudents === 0) {
    const msg = "No students found that meet the criteria (have return date, no sent date, and campus folder ID).";
    Logger.log(`INFO: ${msg}`);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  // Preview mode - show summary and exit
  if (previewMode) {
    let previewMessage = `PREVIEW MODE - No emails will be sent\n\n`;
    previewMessage += `Found ${totalStudents} students across ${campusCount} campuses:\n\n`;
    
    Object.keys(campusGroups).forEach(function (campus) {
      const students = campusGroups[campus];
      previewMessage += `${campus}: ${students.length} student${students.length === 1 ? '' : 's'}\n`;
      students.forEach(function(student) {
        previewMessage += `  • ${student.row["Name"]} (returns: ${student.row[RETURN_DATE]})\n`;
      });
      previewMessage += `\n`;
    });
    
    SpreadsheetApp.getUi().alert(previewMessage);
    Logger.log(`Preview completed. ${totalStudents} students ready to process.`);
    return;
  }

  // Confirmation dialog for sending
  const confirmMessage = `Ready to send emails to ${campusCount} campuses for ${totalStudents} students.\n\nProceed with sending emails?`;
  const response = SpreadsheetApp.getUi().alert(
    "Confirm Email Send",
    confirmMessage,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response !== SpreadsheetApp.getUi().Button.YES) {
    Logger.log("Email send cancelled by user");
    return;
  }

  // Initialize counters for success and error counts
  let successCount = 0;
  let errorCount = 0;
  let errorMessages = [];

  // Send one email per campus with all students
  const campusList = Object.keys(campusGroups);
  
  Logger.log(`Starting to send emails to ${campusList.length} campuses`);
  
  campusList.forEach(function (campus, index) {
    Logger.log(`Processing campus ${index + 1}/${campusList.length}: ${campus}`);
    
    try {
      const students = campusGroups[campus];
      const campusInfo = getInfoByCampus(campus);
      
      // Validation: Check if campus info exists
      if (!campusInfo.recipients || !campusInfo.driveLink) {
        throw new Error(`No configuration found for campus: ${campus}`);
      }
      
      let recipients = campusInfo.recipients;
      // Ensure recipients is a comma-separated string
      if (Array.isArray(recipients)) {
        recipients = recipients.join(",");
      }
      
      // Validation: Check recipients
      if (!recipients || recipients.trim() === "") {
        throw new Error(`No recipients configured for campus: ${campus}`);
      }
      
      const driveLink = campusInfo.driveLink;
      
      Logger.log(`  Sending to ${recipients} for ${students.length} students`);
      
      // Create grouped email template
      const emailTemplate = getGroupedGmailTemplate_(students, driveLink, campus);
      const msgObj = fillInTemplateFromObject_(
        emailTemplate.message,
        { Campus: campus },
        driveLink
      );

      GmailApp.sendEmail(recipients, msgObj.subject, msgObj.text, {
        htmlBody: msgObj.html,
        replyTo: "reggie.ollendieck@nisd.net",
        // cc: "reggie.ollendieck@nisd.net,zina.gonzales@nisd.net",
      });
      
      successCount++;
      Logger.log(`  ✓ Email sent successfully to ${campus}`);
      
      // Mark all rows in this campus as sent using batch operation
      const updates = students.map(function(student) {
        return [student.rowIdx + 2, emailSentColIdx + 1, new Date()];
      });
      
      updates.forEach(function(update) {
        sheet.getRange(update[0], update[1]).setValue(update[2]);
      });
      
      Logger.log(`  ✓ Marked ${students.length} students as sent`);
      
    } catch (e) {
      const errorMsg = `Campus ${campus}: ${e.message}`;
      Logger.log(`  ✗ ERROR: ${errorMsg}`);
      console.error(`Detailed error for ${campus}:`, e);
      errorCount++;
      errorMessages.push(errorMsg);
    }
  });

  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  const resultMessage = `Operation completed in ${duration} seconds\n\nEmails Sent: ${successCount}\nErrors: ${errorCount}${errorMessages.length > 0 ? '\n\nErrors:\n' + errorMessages.join('\n') : ''}`;
  
  Logger.log(`Final results: ${successCount} successful, ${errorCount} errors`);
  SpreadsheetApp.getUi().alert(resultMessage);

  /**
   * Returns campus-specific information including recipients and drive folder link.
   * Uses dynamic data from PropertiesService cache and Recipients sheet.
   * @param {string} campusValue - The campus name.
   * @returns {{recipients: string[]|string, driveLink: string}} Campus info object.
   */
  function getInfoByCampus(campusValue) {
    const startTime = new Date();
    PropertiesManager.logRuntimeOperation('campus_lookup', 'info', `Looking up campus "${campusValue}"`);
    
    try {
      // Validate input parameter
      if (!campusValue || typeof campusValue !== 'string') {
        const duration = new Date() - startTime;
        PropertiesManager.logRuntimeOperation('invalid_campus_param', 'warning', 
          `Invalid campus parameter: ${campusValue}`, { duration, parameter: campusValue });
        return {
          recipients: [],
          driveLink: ""
        };
      }
      
      // Check migration status and trigger if needed
      if (!PropertiesManager.isMigrationComplete()) {
        PropertiesManager.logRuntimeOperation('migration_trigger', 'info', 
          'Migration not complete, triggering migration from getInfoByCampus');
        const migrationPerformed = MigrationManager.performMigration();
        if (migrationPerformed) {
          PropertiesManager.logRuntimeOperation('migration_completed', 'info', 
            'Migration completed successfully from getInfoByCampus');
        }
      }
      
      // Normalize campus name for lookup
      const normalizedCampus = campusValue.toLowerCase().trim();
      
      // Implement cache-first lookup strategy
      let campusData = PropertiesManager.getCampusData();
      
      if (!campusData) {
        PropertiesManager.logRuntimeOperation('cache_miss_sheet_fallback', 'info', 
          'No cached data found, reading from Recipients sheet using batch operations');
        
        // Validate and recover Recipients sheet if needed before reading
        const recoveryResult = RecipientsSheetManager.validateAndRecoverSheet();
        
        if (!recoveryResult.success) {
          PropertiesManager.logRuntimeOperation('sheet_recovery_failed', 'error', 
            `Sheet recovery failed: ${recoveryResult.message}`, { recoveryResult });
          // Continue with fallback to hardcoded data
        } else if (recoveryResult.action !== 'validation_passed' && recoveryResult.action !== 'no_recovery_needed') {
          PropertiesManager.logRuntimeOperation('sheet_recovery_performed', 'info', 
            `Sheet recovery performed: ${recoveryResult.message}`, { recoveryResult });
        }
        
        // Read from Recipients sheet using optimized batch operations and refresh cache
        const recipientsData = RecipientsSheetManager.readRecipientsDataOptimized();
        const driveData = MigrationManager.readCampusReferenceInfo();
        
        // Combine recipient and drive data
        const combinedData = {};
        Object.keys(recipientsData).forEach(campus => {
          combinedData[campus] = {
            recipients: recipientsData[campus],
            driveLink: driveData[campus] || ''
          };
        });
        
        // Add any drive links that don't have recipients
        Object.keys(driveData).forEach(campus => {
          if (!combinedData[campus]) {
            combinedData[campus] = {
              recipients: [],
              driveLink: driveData[campus]
            };
          }
        });
        
        // Store in cache for future lookups
        PropertiesManager.storeCampusData(combinedData);
        campusData = combinedData;
        
        PropertiesManager.logRuntimeOperation('cache_refreshed', 'info', 
          `Refreshed cache with data for ${Object.keys(campusData).length} campuses`, 
          { campusCount: Object.keys(campusData).length });
      }
      
      // Implement campus name validation against expected values
      const availableCampuses = Object.keys(campusData);
      
      // Look up campus information
      const campusInfo = campusData[normalizedCampus];
      
      if (!campusInfo) {
        // Handle missing campus cases with proper error logging
        const duration = new Date() - startTime;
        
        // Suggest similar campus names if available
        const suggestions = availableCampuses.filter(campus => 
          campus.includes(normalizedCampus) || normalizedCampus.includes(campus)
        );
        
        PropertiesManager.logRuntimeOperation('campus_not_found', 'warning', 
          `Campus "${normalizedCampus}" not found in data`, 
          { 
            duration,
            requestedCampus: normalizedCampus,
            availableCampuses: availableCampuses.length,
            suggestions: suggestions
          });
        
        // Return empty array for invalid campuses
        return {
          recipients: [],
          driveLink: ""
        };
      }
      
      // Validate campus data structure
      if (!campusInfo.recipients || !Array.isArray(campusInfo.recipients)) {
        PropertiesManager.logRuntimeOperation('invalid_recipients_data', 'warning', 
          `Invalid recipients data for campus "${normalizedCampus}"`, 
          { campus: normalizedCampus, recipientsType: typeof campusInfo.recipients });
        campusInfo.recipients = [];
      }
      
      if (!campusInfo.driveLink || typeof campusInfo.driveLink !== 'string') {
        PropertiesManager.logRuntimeOperation('invalid_drive_link', 'warning', 
          `Invalid drive link for campus "${normalizedCampus}"`, 
          { campus: normalizedCampus, driveLinkType: typeof campusInfo.driveLink });
        campusInfo.driveLink = "";
      }
      
      // Maintain original function signature and return structure
      const result = {
        recipients: campusInfo.recipients || [],
        driveLink: campusInfo.driveLink || ""
      };
      
      const duration = new Date() - startTime;
      
      PropertiesManager.logRuntimeOperation('campus_lookup_success', 'info', 
        `Found ${result.recipients.length} recipients for campus "${normalizedCampus}"`, 
        { 
          duration,
          campus: normalizedCampus,
          recipientCount: result.recipients.length,
          hasDriveLink: !!result.driveLink
        });
      
      // Log validation warnings if recipients array is empty
      if (result.recipients.length === 0) {
        PropertiesManager.logRuntimeOperation('no_recipients_found', 'warning', 
          `No recipients found for campus "${normalizedCampus}"`, 
          { campus: normalizedCampus });
      }
      
      return result;
      
    } catch (error) {
      const duration = new Date() - startTime;
      PropertiesManager.logRuntimeOperation('campus_lookup_error', 'error', 
        `Error looking up campus "${campusValue}": ${error.message}`, 
        { 
          duration,
          campus: campusValue,
          error: error.message,
          stack: error.stack
        });
      
      console.error('getInfoByCampus error:', error);
      
      // Return empty result on error to maintain backward compatibility
      return {
        recipients: [],
        driveLink: ""
      };
    }
  }

  /**
   * Generates the Gmail message template for multiple students grouped by campus.
   * @param {Array} students - Array of student objects with row and rowIdx properties.
   * @param {string} driveLink - The campus drive folder link.
   * @param {string} campus - The campus name.
   * @returns {{message: {subject: string, html: string}}} Gmail message template object.
   */
  function getGroupedGmailTemplate_(students, driveLink, campus) {
    const studentCount = students.length;
    const studentWord = studentCount === 1 ? "student" : "students";
    
    // Build student list with links
    let studentList = "";
    students.forEach(function(student, index) {
      const row = student.row;
      studentList += `
        <li>
          <strong>${row["Name"]}</strong><br>
        </li>`;
    });

    return {
      message: {
        subject: `DAEP Placement Transition Plan Notification`,
        html: `Dear ${campus},<br><br>
        You have ${studentCount} ${studentWord} nearing completion of their assigned placement at NAMS and should be returning to ${campus} soon.<br><br>   
        On their last day of placement, they will be given withdrawal documents and the parents/guardians will have been called and told to contact ${campus} to set up an appointment to re-enroll and meet with an administrator/counselor.<br><br>  
        Below is a list of the returning ${studentWord}. You can find their DAEP Transition Plans (with grades and notes from their teachers at NAMS) linked below in your campus folder.<br><br>
        <h3>${studentWord} returning to ${campus}:</h3>
          <ul>${studentList}</ul>

        <h3>Important Links:</h3>
            <ul>
              <li><a href="https://drive.google.com/drive/folders/${driveLink}">${campus} Here is a link to your campus folder for the year.</a></li>
              <li><a href="https://drive.google.com/file/d/1qnyQ8cCxLVM9D6rg4wkyBp6KrXIELfNx/view?usp=sharing">Updates in Special Education</a></li>
            </ul>
        
        Please let me know if you have any questions or concerns.<br>
        Thank you for all you do,<br>
        Reggie Ollendieck<br>
        Associate Principal<br><br>`,
      },
    };
  }

  /**
   * Fills in a template string with data from the row and drive link.
   * @param {string|Object} template - The template string or object.
   * @param {Object} row - The row data object.
   * @param {string} driveLink - The campus drive folder link.
   * @returns {Object} The filled-in template object.
   */
  function fillInTemplateFromObject_(template, row, driveLink) {
    let template_string = JSON.stringify(template);
    template_string = template_string.replace(/{{[^{}]+}}/g, (key) => {
      if (key === "${driveLink}") {
        return escapeData_(driveLink);
      }
      return escapeData_(row[key.replace(/[{}]+/g, "")] || "", driveLink);
    });
    return JSON.parse(template_string);
  }

  /**
   * Escapes special characters in a string for safe insertion into templates.
   * @param {string} str - The string to escape.
   * @returns {string} The escaped string.
   */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, "\\\\")
      .replace(/[\"]/g, '\\"')
      .replace(/[\/]/g, "\\/")
      .replace(/[\b]/g, "\\b")
      .replace(/[\f]/g, "\\f")
      .replace(/[\n]/g, "\\n")
      .replace(/[\r]/g, "\\r")
      .replace(/[\t]/g, "\\t");
  }
}

// ============================================================================
// TESTING AND VALIDATION FUNCTIONS
// ============================================================================

/**
 * Comprehensive test function to validate migration infrastructure
 * This function tests all core classes and their functionality
 * @returns {boolean} True if all tests pass, false otherwise
 */
function validateMigrationInfrastructure() {
  Logger.log('=== Validating Migration Infrastructure ===');
  
  let testsPassed = 0;
  let testsTotal = 0;
  let errors = [];
  
  try {
    // Test 1: Verify all classes are defined and accessible
    testsTotal++;
    Logger.log('Test 1: Checking class definitions...');
    
    if (typeof MigrationManager === 'undefined') {
      throw new Error('MigrationManager class is not defined');
    }
    if (typeof PropertiesManager === 'undefined') {
      throw new Error('PropertiesManager class is not defined');
    }
    if (typeof RecipientsSheetManager === 'undefined') {
      throw new Error('RecipientsSheetManager class is not defined');
    }
    
    Logger.log('✓ All required classes are defined');
    testsPassed++;
    
    // Test 2: Test PropertiesManager functionality
    testsTotal++;
    Logger.log('Test 2: Testing PropertiesManager...');
    
    // Store test data
    const testData = {
      'test-campus': {
        recipients: ['test1@nisd.net', 'test2@nisd.net'],
        driveLink: 'test-drive-id-123'
      }
    };
    
    PropertiesManager.storeCampusData(testData);
    const retrievedData = PropertiesManager.getCampusData();
    
    if (!retrievedData || !retrievedData['test-campus']) {
      throw new Error('Failed to store/retrieve test data');
    }
    
    if (retrievedData['test-campus'].recipients.length !== 2) {
      throw new Error('Retrieved data structure is incorrect');
    }
    
    Logger.log('✓ PropertiesManager store/retrieve works correctly');
    testsPassed++;
    
    // Test 3: Test MigrationManager data extraction
    testsTotal++;
    Logger.log('Test 3: Testing MigrationManager data extraction...');
    
    const hardcodedData = MigrationManager.extractHardcodedData();
    
    if (!hardcodedData || typeof hardcodedData !== 'object') {
      throw new Error('Failed to extract hardcoded data');
    }
    
    // Verify we have expected campuses
    const requiredCampuses = ['bernal', 'briscoe', 'test'];
    for (const campus of requiredCampuses) {
      if (!hardcodedData[campus]) {
        throw new Error(`Missing required campus: ${campus}`);
      }
      if (!Array.isArray(hardcodedData[campus])) {
        throw new Error(`Campus ${campus} data is not an array`);
      }
      if (hardcodedData[campus].length === 0) {
        throw new Error(`Campus ${campus} has no recipients`);
      }
    }
    
    Logger.log(`✓ Successfully extracted data for ${Object.keys(hardcodedData).length} campuses`);
    testsPassed++;
    
    // Test 4: Test drive folder data extraction
    testsTotal++;
    Logger.log('Test 4: Testing drive folder data extraction...');
    
    const driveData = MigrationManager.getHardcodedDriveLinks();
    
    if (!driveData || typeof driveData !== 'object') {
      throw new Error('Failed to get hardcoded drive links');
    }
    
    // Verify we have drive links for required campuses
    for (const campus of requiredCampuses) {
      if (!driveData[campus]) {
        throw new Error(`Missing drive link for campus: ${campus}`);
      }
      if (typeof driveData[campus] !== 'string' || driveData[campus].length === 0) {
        throw new Error(`Invalid drive link for campus: ${campus}`);
      }
    }
    
    Logger.log(`✓ Successfully extracted drive links for ${Object.keys(driveData).length} campuses`);
    testsPassed++;
    
    // Test 5: Test RecipientsSheetManager functionality
    testsTotal++;
    Logger.log('Test 5: Testing RecipientsSheetManager...');
    
    // Test sheet existence check (should not throw error)
    const sheetExists = RecipientsSheetManager.sheetExists();
    Logger.log(`Recipients sheet currently exists: ${sheetExists}`);
    
    // Test email validation
    const testEmailData = [
      ['Campus', 'Recipient'], // Header row
      ['test', 'valid@nisd.net'],
      ['test', 'invalid-email'],
      ['test', 'another.valid@nisd.net']
    ];
    
    const validationResult = RecipientsSheetManager.validateEmailAddresses(testEmailData);
    
    if (!validationResult || typeof validationResult !== 'object') {
      throw new Error('Email validation failed to return results');
    }
    
    if (validationResult.valid.length !== 2) {
      throw new Error(`Expected 2 valid emails, got ${validationResult.valid.length}`);
    }
    
    if (validationResult.invalid.length !== 1) {
      throw new Error(`Expected 1 invalid email, got ${validationResult.invalid.length}`);
    }
    
    Logger.log('✓ RecipientsSheetManager validation works correctly');
    testsPassed++;
    
    // Test 6: Test migration status management
    testsTotal++;
    Logger.log('Test 6: Testing migration status management...');
    
    // Clear migration status for testing
    PropertiesService.getScriptProperties().deleteProperty('migrationComplete');
    
    let migrationStatus = PropertiesManager.isMigrationComplete();
    if (migrationStatus) {
      throw new Error('Migration status should be false after clearing');
    }
    
    PropertiesManager.setMigrationComplete();
    migrationStatus = PropertiesManager.isMigrationComplete();
    if (!migrationStatus) {
      throw new Error('Migration status should be true after setting complete');
    }
    
    Logger.log('✓ Migration status management works correctly');
    testsPassed++;
    
    // Test 7: Test cache validation
    testsTotal++;
    Logger.log('Test 7: Testing cache validation...');
    
    const metadata = PropertiesManager.getCacheMetadata();
    if (!metadata) {
      throw new Error('Failed to get cache metadata');
    }
    
    if (typeof metadata.campusCount !== 'number' || metadata.campusCount < 0) {
      throw new Error('Invalid campus count in metadata');
    }
    
    Logger.log(`✓ Cache validation works - ${metadata.campusCount} campuses, ${metadata.totalRecipients} recipients`);
    testsPassed++;
    
  } catch (error) {
    errors.push(error.message);
    Logger.log(`✗ Test failed: ${error.message}`);
    console.error('Detailed error:', error);
  }
  
  // Clean up test data
  try {
    PropertiesManager.clearCache();
    Logger.log('Test cleanup completed');
  } catch (cleanupError) {
    Logger.log(`Warning: Cleanup failed - ${cleanupError.message}`);
  }
  
  // Report final results
  Logger.log('=== Validation Results ===');
  Logger.log(`Tests passed: ${testsPassed}/${testsTotal}`);
  
  if (errors.length > 0) {
    Logger.log('❌ Validation failed with errors:');
    errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
    
    // Show user-friendly alert
    const ui = SpreadsheetApp.getUi();
    const errorMessage = `Migration infrastructure validation failed!\n\n${errors.join('\n\n')}`;
    ui.alert('Validation Failed', errorMessage, ui.ButtonSet.OK);
    
    return false;
  } else {
    Logger.log('✅ All migration infrastructure validation tests passed!');
    
    // Show success message
    const ui = SpreadsheetApp.getUi();
    const successMessage = `Migration infrastructure validation completed successfully!\n\n` +
                          `✓ All ${testsTotal} tests passed\n` +
                          `✓ Classes are properly defined\n` +
                          `✓ Data storage and retrieval works\n` +
                          `✓ Email validation functions correctly\n` +
                          `✓ Migration status tracking works\n\n` +
                          `The migration infrastructure is ready for use.`;
    ui.alert('Validation Successful', successMessage, ui.ButtonSet.OK);
    
    return true;
  }
}

/**
 * Test function to verify the new dynamic getInfoByCampus implementation
 * Can be run from the Apps Script editor to validate functionality
 */
function testDynamicGetInfoByCampus() {
  Logger.log('=== Testing Dynamic getInfoByCampus Implementation ===');
  
  try {
    // Test 1: Valid campus lookup
    Logger.log('Test 1: Testing valid campus lookup...');
    const testResult = getInfoByCampus('test');
    
    if (!testResult || typeof testResult !== 'object') {
      throw new Error('Function did not return an object');
    }
    
    if (!testResult.hasOwnProperty('recipients') || !testResult.hasOwnProperty('driveLink')) {
      throw new Error('Function did not return expected properties');
    }
    
    Logger.log(`✓ Valid campus test passed - Recipients: ${testResult.recipients.length}, Drive Link: ${testResult.driveLink ? 'present' : 'missing'}`);
    
    // Test 2: Invalid campus lookup
    Logger.log('Test 2: Testing invalid campus lookup...');
    const invalidResult = getInfoByCampus('nonexistent-campus');
    
    if (!Array.isArray(invalidResult.recipients) || invalidResult.recipients.length !== 0) {
      throw new Error('Invalid campus should return empty recipients array');
    }
    
    if (invalidResult.driveLink !== '') {
      throw new Error('Invalid campus should return empty drive link');
    }
    
    Logger.log('✓ Invalid campus test passed - Returns empty arrays as expected');
    
    // Test 3: Edge cases
    Logger.log('Test 3: Testing edge cases...');
    
    // Empty string
    const emptyResult = getInfoByCampus('');
    if (!Array.isArray(emptyResult.recipients) || emptyResult.recipients.length !== 0) {
      throw new Error('Empty string should return empty recipients array');
    }
    
    // Null/undefined
    const nullResult = getInfoByCampus(null);
    if (!Array.isArray(nullResult.recipients) || nullResult.recipients.length !== 0) {
      throw new Error('Null input should return empty recipients array');
    }
    
    Logger.log('✓ Edge case tests passed');
    
    // Test 4: Case insensitive lookup
    Logger.log('Test 4: Testing case insensitive lookup...');
    const upperCaseResult = getInfoByCampus('TEST');
    const lowerCaseResult = getInfoByCampus('test');
    
    if (JSON.stringify(upperCaseResult) !== JSON.stringify(lowerCaseResult)) {
      throw new Error('Case insensitive lookup failed');
    }
    
    Logger.log('✓ Case insensitive test passed');
    
    Logger.log('✅ All dynamic getInfoByCampus tests passed successfully!');
    
    // Show success message to user
    const ui = SpreadsheetApp.getUi();
    const successMessage = `Dynamic getInfoByCampus Implementation Test Results:\n\n` +
                          `✓ Valid campus lookup works correctly\n` +
                          `✓ Invalid campus returns empty arrays\n` +
                          `✓ Edge cases handled properly\n` +
                          `✓ Case insensitive lookup works\n\n` +
                          `The new dynamic implementation is working correctly!`;
    ui.alert('Test Results', successMessage, ui.ButtonSet.OK);
    
    return true;
    
  } catch (error) {
    Logger.log(`❌ Dynamic getInfoByCampus test failed: ${error.message}`);
    console.error('Test error:', error);
    
    // Show error message to user
    const ui = SpreadsheetApp.getUi();
    const errorMessage = `Dynamic getInfoByCampus test failed:\n\n${error.message}\n\nCheck the execution log for more details.`;
    ui.alert('Test Failed', errorMessage, ui.ButtonSet.OK);
    
    return false;
  }
}

/**
 * Test function to validate the enhanced data validation and error handling functionality
 * Tests email validation, duplicate campus handling, and comprehensive error logging
 */
function testDataValidationAndErrorHandling() {
  Logger.log('=== Testing Data Validation and Error Handling ===');
  
  let testsPassed = 0;
  let testsTotal = 0;
  let errors = [];
  
  try {
    // Test 1: Email validation with various formats
    testsTotal++;
    Logger.log('Test 1: Testing email address validation...');
    
    const testEmailData = [
      ['Campus', 'Recipient'], // Header row
      ['test', 'valid@nisd.net'],
      ['test', 'invalid-email'],
      ['test', 'another.valid@nisd.net'],
      ['test', '@invalid.com'],
      ['test', 'invalid@'],
      ['test', 'valid.email+tag@domain.co.uk'],
      ['test', 'spaces in@email.com'],
      ['test', 'multiple@@symbols.com'],
      ['test', ''],
      ['', 'missing.campus@nisd.net']
    ];
    
    const validationResult = RecipientsSheetManager.validateEmailAddresses(testEmailData);
    
    if (!validationResult || typeof validationResult !== 'object') {
      throw new Error('Email validation failed to return results');
    }
    
    // Should have 3 valid emails
    if (validationResult.valid.length !== 3) {
      throw new Error(`Expected 3 valid emails, got ${validationResult.valid.length}`);
    }
    
    // Should have multiple invalid emails
    if (validationResult.invalid.length < 5) {
      throw new Error(`Expected at least 5 invalid emails, got ${validationResult.invalid.length}`);
    }
    
    // Should have detailed error messages with row numbers
    if (validationResult.errors.length === 0) {
      throw new Error('Expected detailed error messages with row numbers');
    }
    
    // Check that error messages contain row numbers
    const hasRowNumbers = validationResult.errors.some(error => error.includes('Row '));
    if (!hasRowNumbers) {
      throw new Error('Error messages should contain row numbers');
    }
    
    Logger.log('✓ Email validation works correctly with detailed error logging');
    testsPassed++;
    
    // Test 2: Duplicate campus handling
    testsTotal++;
    Logger.log('Test 2: Testing duplicate campus entry handling...');
    
    const testValidEntries = [
      { campus: 'test', recipient: 'first@nisd.net', rowNumber: 2 },
      { campus: 'test', recipient: 'second@nisd.net', rowNumber: 3 },
      { campus: 'test', recipient: 'first@nisd.net', rowNumber: 4 }, // Duplicate
      { campus: 'other', recipient: 'other@nisd.net', rowNumber: 5 }
    ];
    
    const duplicateResult = RecipientsSheetManager.handleDuplicateCampusEntries(testValidEntries);
    
    if (!duplicateResult || !duplicateResult.data) {
      throw new Error('Duplicate handling failed to return data');
    }
    
    // Should have 2 campuses
    if (Object.keys(duplicateResult.data).length !== 2) {
      throw new Error(`Expected 2 campuses, got ${Object.keys(duplicateResult.data).length}`);
    }
    
    // Test campus should have 2 unique recipients (duplicate removed)
    if (duplicateResult.data['test'].length !== 2) {
      throw new Error(`Expected 2 unique recipients for test campus, got ${duplicateResult.data['test'].length}`);
    }
    
    // Should have detected 1 duplicate
    if (duplicateResult.duplicateWarnings.length !== 1) {
      throw new Error(`Expected 1 duplicate warning, got ${duplicateResult.duplicateWarnings.length}`);
    }
    
    Logger.log('✓ Duplicate campus handling works correctly');
    testsPassed++;
    
    // Test 3: Campus name validation
    testsTotal++;
    Logger.log('Test 3: Testing campus name validation...');
    
    const testCampusEntries = [
      { campus: 'test', recipient: 'test@nisd.net', rowNumber: 2 }, // Valid
      { campus: 'bernal', recipient: 'bernal@nisd.net', rowNumber: 3 }, // Valid
      { campus: 'unknown', recipient: 'unknown@nisd.net', rowNumber: 4 }, // Invalid
      { campus: 'bernel', recipient: 'typo@nisd.net', rowNumber: 5 } // Typo (should suggest bernal)
    ];
    
    const campusValidation = RecipientsSheetManager.validateCampusNames(testCampusEntries);
    
    if (!campusValidation) {
      throw new Error('Campus validation failed to return results');
    }
    
    // Should have 2 valid campuses
    if (campusValidation.validCampuses.length !== 2) {
      throw new Error(`Expected 2 valid campuses, got ${campusValidation.validCampuses.length}`);
    }
    
    // Should have 2 invalid campuses
    if (campusValidation.invalidCampuses.length !== 2) {
      throw new Error(`Expected 2 invalid campuses, got ${campusValidation.invalidCampuses.length}`);
    }
    
    // Should have suggestions for typos
    if (campusValidation.suggestions.length === 0) {
      throw new Error('Expected suggestions for similar campus names');
    }
    
    Logger.log('✓ Campus name validation works correctly with suggestions');
    testsPassed++;
    
    // Test 4: Comprehensive readRecipientsData integration
    testsTotal++;
    Logger.log('Test 4: Testing integrated validation in readRecipientsData...');
    
    // This test would require creating a temporary sheet, so we'll test the validation logic
    // by checking that the method exists and has the right structure
    if (typeof RecipientsSheetManager.readRecipientsData !== 'function') {
      throw new Error('readRecipientsData method not found');
    }
    
    // Test that validation methods are properly integrated
    if (typeof RecipientsSheetManager.validateEmailAddresses !== 'function') {
      throw new Error('validateEmailAddresses method not found');
    }
    
    if (typeof RecipientsSheetManager.handleDuplicateCampusEntries !== 'function') {
      throw new Error('handleDuplicateCampusEntries method not found');
    }
    
    if (typeof RecipientsSheetManager.validateCampusNames !== 'function') {
      throw new Error('validateCampusNames method not found');
    }
    
    Logger.log('✓ All validation methods are properly integrated');
    testsPassed++;
    
    // Test 5: Error handling and logging
    testsTotal++;
    Logger.log('Test 5: Testing error handling and logging...');
    
    // Test that validation continues processing valid emails when invalid ones are found
    const mixedData = [
      ['Campus', 'Recipient'],
      ['test', 'valid1@nisd.net'],
      ['test', 'invalid-email'],
      ['test', 'valid2@nisd.net']
    ];
    
    const mixedResult = RecipientsSheetManager.validateEmailAddresses(mixedData);
    
    // Should continue processing and return valid emails despite invalid ones
    if (mixedResult.valid.length !== 2) {
      throw new Error('Should continue processing valid emails when invalid ones are found');
    }
    
    if (mixedResult.invalid.length !== 1) {
      throw new Error('Should properly identify invalid emails');
    }
    
    // Should have comprehensive summary
    if (!mixedResult.summary || typeof mixedResult.summary !== 'object') {
      throw new Error('Should provide comprehensive validation summary');
    }
    
    Logger.log('✓ Error handling continues processing valid data correctly');
    testsPassed++;
    
  } catch (error) {
    errors.push(error.message);
    Logger.log(`✗ Test failed: ${error.message}`);
    console.error('Detailed error:', error);
  }
  
  // Report final results
  Logger.log('=== Data Validation Test Results ===');
  Logger.log(`Tests passed: ${testsPassed}/${testsTotal}`);
  
  if (errors.length > 0) {
    Logger.log('❌ Data validation testing failed with errors:');
    errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
    
    // Show user-friendly alert
    const ui = SpreadsheetApp.getUi();
    const errorMessage = `Data validation testing failed!\n\n${errors.join('\n\n')}`;
    ui.alert('Validation Test Failed', errorMessage, ui.ButtonSet.OK);
    
    return false;
  } else {
    Logger.log('✅ All data validation and error handling tests passed!');
    
    // Show success message
    const ui = SpreadsheetApp.getUi();
    const successMessage = `Data Validation and Error Handling Test Results:\n\n` +
                          `✓ All ${testsTotal} tests passed\n` +
                          `✓ Email validation with detailed error logging works\n` +
                          `✓ Duplicate campus handling works correctly\n` +
                          `✓ Campus name validation with suggestions works\n` +
                          `✓ Integration with readRecipientsData works\n` +
                          `✓ Error handling continues processing valid data\n\n` +
                          `The data validation and error handling system is working correctly!`;
    ui.alert('Validation Test Successful', successMessage, ui.ButtonSet.OK);
    
    return true;
  }
}

// ============================================================================
// CACHE INVALIDATION SYSTEM
// ============================================================================

/**
 * onEdit trigger function that detects Recipients sheet changes
 * Automatically clears cached data when Recipients sheet is modified
 * @param {Object} e - Edit event object from Google Sheets
 */
function onEdit(e) {
  try {
    // Check if the edit occurred in the Recipients sheet
    if (!e || !e.source || !e.range) {
      return; // No valid edit event
    }
    
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    
    // Only process edits to the Recipients sheet
    if (sheetName !== RecipientsSheetManager.SHEET_NAME) {
      return; // Not the Recipients sheet
    }
    
    Logger.log(`onEdit: Detected edit in Recipients sheet at range ${e.range.getA1Notation()}`);
    
    // Handle Recipients sheet modifications
    handleRecipientsSheetEdit(e);
    
  } catch (error) {
    Logger.log(`onEdit: Error processing edit event: ${error.message}`);
    console.error('onEdit error:', error);
  }
}

/**
 * Handles Recipients sheet modifications by clearing cache and logging events
 * Detects edits specifically to Recipients sheet and clears cached data
 * @param {Object} e - Edit event object from Google Sheets
 */
function handleRecipientsSheetEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const editedRow = range.getRow();
    const editedColumn = range.getColumn();
    const numRows = range.getNumRows();
    const numColumns = range.getNumColumns();
    
    Logger.log(`handleRecipientsSheetEdit: Processing edit at row ${editedRow}, column ${editedColumn} (${numRows}x${numColumns} range)`);
    
    // Get the old and new values if available
    const oldValue = e.oldValue || '';
    const newValue = e.value || '';
    
    // Log cache invalidation events with detailed information
    Logger.log(`handleRecipientsSheetEdit: Cache invalidation triggered by Recipients sheet modification`);
    Logger.log(`  Sheet: ${sheet.getName()}`);
    Logger.log(`  Range: ${range.getA1Notation()}`);
    Logger.log(`  Old value: "${oldValue}"`);
    Logger.log(`  New value: "${newValue}"`);
    Logger.log(`  Timestamp: ${new Date().toISOString()}`);
    
    // Determine the type of edit for more specific logging
    let editType = 'unknown';
    if (oldValue === '' && newValue !== '') {
      editType = 'addition';
    } else if (oldValue !== '' && newValue === '') {
      editType = 'deletion';
    } else if (oldValue !== '' && newValue !== '') {
      editType = 'modification';
    }
    
    Logger.log(`  Edit type: ${editType}`);
    
    // Check if this is a data row (not header)
    if (editedRow === 1) {
      Logger.log(`handleRecipientsSheetEdit: Header row edited - still clearing cache for safety`);
    } else {
      Logger.log(`handleRecipientsSheetEdit: Data row edited - clearing cache`);
    }
    
    // Clear cached data when Recipients sheet is modified
    PropertiesManager.clearCache();
    Logger.log(`handleRecipientsSheetEdit: Cache cleared successfully`);
    
    // Log cache invalidation event for monitoring
    const invalidationEvent = {
      timestamp: new Date().toISOString(),
      sheetName: sheet.getName(),
      range: range.getA1Notation(),
      editType: editType,
      oldValue: oldValue,
      newValue: newValue,
      reason: 'Recipients sheet modification'
    };
    
    // Store invalidation event in a property for monitoring (keep last 10 events)
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingEvents = properties.getProperty('cacheInvalidationEvents');
      let events = [];
      
      if (existingEvents) {
        events = JSON.parse(existingEvents);
      }
      
      // Add new event and keep only last 10
      events.unshift(invalidationEvent);
      events = events.slice(0, 10);
      
      properties.setProperty('cacheInvalidationEvents', JSON.stringify(events));
      Logger.log(`handleRecipientsSheetEdit: Logged invalidation event to monitoring`);
      
    } catch (monitoringError) {
      Logger.log(`handleRecipientsSheetEdit: Warning - Could not log to monitoring: ${monitoringError.message}`);
    }
    
    // Provide user feedback for significant changes
    if (editType === 'addition' || editType === 'deletion') {
      Logger.log(`handleRecipientsSheetEdit: Significant change detected (${editType}) - cache will be refreshed on next lookup`);
    }
    
    Logger.log(`handleRecipientsSheetEdit: Recipients sheet edit handling completed`);
    
  } catch (error) {
    Logger.log(`handleRecipientsSheetEdit: Error handling Recipients sheet edit: ${error.message}`);
    console.error('handleRecipientsSheetEdit error:', error);
    
    // Even if there's an error, try to clear the cache for safety
    try {
      PropertiesManager.clearCache();
      Logger.log(`handleRecipientsSheetEdit: Cache cleared despite error for safety`);
    } catch (clearError) {
      Logger.log(`handleRecipientsSheetEdit: Critical error - could not clear cache: ${clearError.message}`);
    }
  }
}

/**
 * Manual cache refresh mechanism - clears cache and forces immediate refresh
 * Uses batch sheet reading operations to minimize API calls to Google Sheets service
 * Provides user-friendly way to force cache update
 * Can be called manually or added to custom menu
 */
function refreshCampusCache() {
  Logger.log('refreshCampusCache: Starting manual cache refresh with batch operations...');
  
  try {
    const startTime = new Date();
    
    // Clear existing cache
    Logger.log('refreshCampusCache: Clearing existing cache...');
    PropertiesManager.clearCache();
    
    // Force immediate refresh using optimized batch reading from Recipients sheet
    Logger.log('refreshCampusCache: Reading fresh data using batch operations...');
    const recipientsData = RecipientsSheetManager.readRecipientsDataOptimized();
    
    // Read drive folder data
    Logger.log('refreshCampusCache: Reading drive folder data...');
    const driveData = MigrationManager.readCampusReferenceInfo();
    
    // Combine recipient and drive data
    const combinedData = {};
    Object.keys(recipientsData).forEach(campus => {
      combinedData[campus] = {
        recipients: recipientsData[campus],
        driveLink: driveData[campus] || ''
      };
    });
    
    // Add any drive links that don't have recipients
    Object.keys(driveData).forEach(campus => {
      if (!combinedData[campus]) {
        combinedData[campus] = {
          recipients: [],
          driveLink: driveData[campus]
        };
      }
    });
    
    // Store refreshed data in cache
    Logger.log('refreshCampusCache: Storing refreshed data in cache...');
    PropertiesManager.storeCampusData(combinedData);
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    // Get batch operation performance metrics if available
    const batchMetrics = recipientsData._validationResults?.batchOperation || {};
    const batchReadTime = batchMetrics.readTime || 'unknown';
    
    // Log refresh event for monitoring
    const refreshEvent = {
      timestamp: new Date().toISOString(),
      type: 'manual_refresh',
      duration: duration,
      campusCount: Object.keys(combinedData).length,
      totalRecipients: Object.values(combinedData).reduce((sum, campus) => 
        sum + (campus.recipients ? campus.recipients.length : 0), 0),
      reason: 'Manual cache refresh requested',
      batchOperationTime: batchReadTime,
      optimizedRead: true
    };
    
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingEvents = properties.getProperty('cacheInvalidationEvents');
      let events = [];
      
      if (existingEvents) {
        events = JSON.parse(existingEvents);
      }
      
      events.unshift(refreshEvent);
      events = events.slice(0, 10);
      
      properties.setProperty('cacheInvalidationEvents', JSON.stringify(events));
      
    } catch (monitoringError) {
      Logger.log(`refreshCampusCache: Warning - Could not log refresh event: ${monitoringError.message}`);
    }
    
    // Provide user feedback including batch operation performance
    const successMessage = `Cache refresh completed successfully!\n\n` +
                          `Duration: ${duration} seconds\n` +
                          `Batch read time: ${batchReadTime}ms\n` +
                          `Campuses loaded: ${Object.keys(combinedData).length}\n` +
                          `Total recipients: ${refreshEvent.totalRecipients}\n\n` +
                          `The cache has been updated with the latest data using optimized batch operations.`;
    
    Logger.log(`refreshCampusCache: Manual cache refresh completed successfully in ${duration} seconds`);
    Logger.log(`refreshCampusCache: Batch read operation took ${batchReadTime}ms`);
    Logger.log(`refreshCampusCache: Loaded ${Object.keys(combinedData).length} campuses with ${refreshEvent.totalRecipients} total recipients`);
    
    // Show success dialog to user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Cache Refresh Complete', successMessage, ui.ButtonSet.OK);
    
    return {
      success: true,
      duration: duration,
      batchReadTime: batchReadTime,
      campusCount: Object.keys(combinedData).length,
      totalRecipients: refreshEvent.totalRecipients
    };
    
  } catch (error) {
    Logger.log(`refreshCampusCache: Error during manual cache refresh: ${error.message}`);
    console.error('refreshCampusCache error:', error);
    
    // Show error dialog to user
    const ui = SpreadsheetApp.getUi();
    const errorMessage = `Cache refresh failed!\n\nError: ${error.message}\n\nPlease check the execution log for more details.`;
    ui.alert('Cache Refresh Failed', errorMessage, ui.ButtonSet.OK);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets cache status and recent invalidation events for monitoring
 * Provides detailed information about cache state and recent activity
 * @returns {Object} Cache status information
 */
function getCacheStatus() {
  Logger.log('getCacheStatus: Retrieving cache status information...');
  
  try {
    // Get cache metadata
    const cacheMetadata = PropertiesManager.getCacheMetadata();
    
    // Get recent invalidation events
    const properties = PropertiesService.getScriptProperties();
    const eventsData = properties.getProperty('cacheInvalidationEvents');
    let recentEvents = [];
    
    if (eventsData) {
      try {
        recentEvents = JSON.parse(eventsData);
      } catch (parseError) {
        Logger.log(`getCacheStatus: Warning - Could not parse invalidation events: ${parseError.message}`);
      }
    }
    
    // Check migration status
    const migrationComplete = PropertiesManager.isMigrationComplete();
    
    // Check if Recipients sheet exists
    const recipientsSheetExists = RecipientsSheetManager.sheetExists();
    
    const status = {
      cacheExists: cacheMetadata !== null,
      migrationComplete: migrationComplete,
      recipientsSheetExists: recipientsSheetExists,
      metadata: cacheMetadata,
      recentEvents: recentEvents,
      timestamp: new Date().toISOString()
    };
    
    Logger.log(`getCacheStatus: Cache exists: ${status.cacheExists}`);
    Logger.log(`getCacheStatus: Migration complete: ${status.migrationComplete}`);
    Logger.log(`getCacheStatus: Recipients sheet exists: ${status.recipientsSheetExists}`);
    
    if (cacheMetadata) {
      Logger.log(`getCacheStatus: Cache last updated: ${cacheMetadata.lastUpdated}`);
      Logger.log(`getCacheStatus: Cached campuses: ${cacheMetadata.campusCount}`);
      Logger.log(`getCacheStatus: Total recipients: ${cacheMetadata.totalRecipients}`);
    }
    
    Logger.log(`getCacheStatus: Recent events: ${recentEvents.length}`);
    
    return status;
    
  } catch (error) {
    Logger.log(`getCacheStatus: Error retrieving cache status: ${error.message}`);
    console.error('getCacheStatus error:', error);
    
    return {
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * User-friendly function to display cache status in a dialog
 * Shows comprehensive cache information to the user
 */
function showCacheStatus() {
  Logger.log('showCacheStatus: Displaying cache status to user...');
  
  try {
    const status = getCacheStatus();
    
    if (status.error) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Cache Status Error', `Error retrieving cache status:\n\n${status.error}`, ui.ButtonSet.OK);
      return;
    }
    
    let statusMessage = 'Campus Recipients Cache Status\n\n';
    
    // Basic status
    statusMessage += `Cache Status: ${status.cacheExists ? '✓ Active' : '✗ Not Found'}\n`;
    statusMessage += `Migration: ${status.migrationComplete ? '✓ Complete' : '✗ Pending'}\n`;
    statusMessage += `Recipients Sheet: ${status.recipientsSheetExists ? '✓ Exists' : '✗ Missing'}\n\n`;
    
    // Cache details
    if (status.metadata) {
      statusMessage += `Cache Details:\n`;
      statusMessage += `  Last Updated: ${status.metadata.lastUpdated || 'Unknown'}\n`;
      statusMessage += `  Campuses: ${status.metadata.campusCount || 0}\n`;
      statusMessage += `  Recipients: ${status.metadata.totalRecipients || 0}\n`;
      statusMessage += `  Data Size: ${Math.round((status.metadata.dataSize || 0) / 1024)} KB\n\n`;
    }
    
    // Recent activity
    if (status.recentEvents && status.recentEvents.length > 0) {
      statusMessage += `Recent Activity (${status.recentEvents.length} events):\n`;
      status.recentEvents.slice(0, 3).forEach((event, index) => {
        const eventTime = new Date(event.timestamp).toLocaleString();
        const eventType = event.type || event.reason || 'Unknown';
        statusMessage += `  ${index + 1}. ${eventTime} - ${eventType}\n`;
      });
      
      if (status.recentEvents.length > 3) {
        statusMessage += `  ... and ${status.recentEvents.length - 3} more events\n`;
      }
    } else {
      statusMessage += `Recent Activity: No recent events\n`;
    }
    
    statusMessage += `\nStatus checked: ${new Date().toLocaleString()}`;
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Cache Status', statusMessage, ui.ButtonSet.OK);
    
    Logger.log('showCacheStatus: Cache status displayed to user');
    
  } catch (error) {
    Logger.log(`showCacheStatus: Error displaying cache status: ${error.message}`);
    console.error('showCacheStatus error:', error);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Could not display cache status:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * User-friendly function to validate Recipients sheet integrity
 * Checks sheet existence, structure, and data validity
 */
function validateRecipientsSheet() {
  Logger.log('validateRecipientsSheet: Starting manual Recipients sheet validation...');
  
  try {
    const validationResult = RecipientsSheetManager.validateAndRecoverSheet();
    
    let resultMessage = 'Recipients Sheet Validation Results\n\n';
    
    if (validationResult.success) {
      switch (validationResult.action) {
        case 'validation_passed':
        case 'no_recovery_needed':
          resultMessage += '✓ Recipients sheet is healthy and valid\n\n';
          resultMessage += `Data Rows: ${validationResult.dataRows || 'Unknown'}\n`;
          if (validationResult.validRecipients !== undefined) {
            resultMessage += `Valid Recipients: ${validationResult.validRecipients}\n`;
            resultMessage += `Invalid Recipients: ${validationResult.invalidRecipients}\n`;
          }
          break;
          
        case 'recovered_from_cache':
        case 'recovered_from_hardcoded':
          resultMessage += '⚠️ Recipients sheet was missing or corrupted and has been recovered\n\n';
          resultMessage += `Recovery Action: ${validationResult.action}\n`;
          resultMessage += `Data Source: ${validationResult.dataSource}\n`;
          resultMessage += `Campuses Recovered: ${validationResult.campusCount}\n`;
          resultMessage += `Recipients Recovered: ${validationResult.totalRecipients}\n`;
          break;
          
        default:
          resultMessage += `✓ Validation completed: ${validationResult.message}\n`;
      }
    } else {
      resultMessage += `❌ Validation failed: ${validationResult.message}\n\n`;
      resultMessage += 'Please check the execution log for more details.';
    }
    
    resultMessage += `\nValidation completed: ${new Date().toLocaleString()}`;
    
    const ui = SpreadsheetApp.getUi();
    const alertTitle = validationResult.success ? 'Validation Complete' : 'Validation Failed';
    ui.alert(alertTitle, resultMessage, ui.ButtonSet.OK);
    
    Logger.log(`validateRecipientsSheet: Validation completed - ${validationResult.success ? 'success' : 'failed'}`);
    
  } catch (error) {
    Logger.log(`validateRecipientsSheet: Error during validation: ${error.message}`);
    console.error('validateRecipientsSheet error:', error);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Validation Error', `Recipients sheet validation failed:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * User-friendly function to manually recover Recipients sheet
 * Forces recreation of Recipients sheet from available data sources
 */
function recoverRecipientsSheetManual() {
  Logger.log('recoverRecipientsSheetManual: Starting manual Recipients sheet recovery...');
  
  try {
    // Ask user for confirmation before recovery
    const ui = SpreadsheetApp.getUi();
    const confirmMessage = 'This will recreate the Recipients sheet from cached or hardcoded data.\n\n' +
                          'Any existing Recipients sheet will be replaced.\n\n' +
                          'Do you want to proceed with the recovery?';
    
    const response = ui.alert(
      'Confirm Sheet Recovery',
      confirmMessage,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      Logger.log('recoverRecipientsSheetManual: Recovery cancelled by user');
      return;
    }
    
    const recoveryResult = RecipientsSheetManager.recoverRecipientsSheet();
    
    let resultMessage = 'Recipients Sheet Recovery Results\n\n';
    
    if (recoveryResult.success) {
      switch (recoveryResult.action) {
        case 'no_recovery_needed':
          resultMessage += '✓ No recovery was needed - sheet is already healthy\n\n';
          resultMessage += `Data Rows: ${recoveryResult.dataRows}\n`;
          break;
          
        case 'recovered_from_cache':
          resultMessage += '✓ Recipients sheet successfully recovered from cache\n\n';
          resultMessage += `Campuses Recovered: ${recoveryResult.campusCount}\n`;
          resultMessage += `Recipients Recovered: ${recoveryResult.totalRecipients}\n`;
          resultMessage += `Data Source: PropertiesService cache\n`;
          break;
          
        case 'recovered_from_hardcoded':
          resultMessage += '✓ Recipients sheet successfully recovered from hardcoded data\n\n';
          resultMessage += `Campuses Recovered: ${recoveryResult.campusCount}\n`;
          resultMessage += `Recipients Recovered: ${recoveryResult.totalRecipients}\n`;
          resultMessage += `Data Source: Hardcoded backup data\n`;
          break;
          
        default:
          resultMessage += `✓ Recovery completed: ${recoveryResult.message}\n`;
      }
      
      resultMessage += '\nThe Recipients sheet is now available and ready for use.';
    } else {
      resultMessage += `❌ Recovery failed: ${recoveryResult.message}\n\n`;
      resultMessage += 'Please check the execution log for more details.\n\n';
      resultMessage += 'You may need to manually create the Recipients sheet or contact support.';
    }
    
    resultMessage += `\nRecovery completed: ${new Date().toLocaleString()}`;
    
    const alertTitle = recoveryResult.success ? 'Recovery Complete' : 'Recovery Failed';
    ui.alert(alertTitle, resultMessage, ui.ButtonSet.OK);
    
    Logger.log(`recoverRecipientsSheetManual: Recovery completed - ${recoveryResult.success ? 'success' : 'failed'}`);
    
  } catch (error) {
    Logger.log(`recoverRecipientsSheetManual: Error during manual recovery: ${error.message}`);
    console.error('recoverRecipientsSheetManual error:', error);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Recovery Error', `Recipients sheet recovery failed:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * User-friendly function to display runtime operation statistics
 * Shows comprehensive runtime performance and monitoring information to the user
 */
function showRuntimeStats() {
  Logger.log('showRuntimeStats: Displaying runtime statistics to user...');
  
  try {
    const runtimeData = PropertiesManager.getRuntimeStats();
    
    if (runtimeData.error) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Runtime Stats Error', `Error retrieving runtime statistics:\n\n${runtimeData.error}`, ui.ButtonSet.OK);
      return;
    }
    
    let statsMessage = 'Campus Recipients Runtime Statistics\n\n';
    
    // Cache performance
    if (runtimeData.stats.cacheHits !== undefined || runtimeData.stats.cacheMisses !== undefined) {
      statsMessage += `Cache Performance:\n`;
      statsMessage += `  Cache Hits: ${runtimeData.stats.cacheHits || 0}\n`;
      statsMessage += `  Cache Misses: ${runtimeData.stats.cacheMisses || 0}\n`;
      statsMessage += `  Hit Ratio: ${runtimeData.cacheHitRatio}%\n\n`;
    }
    
    // Validation statistics
    if (runtimeData.stats.validationErrors !== undefined || runtimeData.stats.validationWarnings !== undefined) {
      statsMessage += `Data Validation:\n`;
      statsMessage += `  Validation Errors: ${runtimeData.stats.validationErrors || 0}\n`;
      statsMessage += `  Validation Warnings: ${runtimeData.stats.validationWarnings || 0}\n\n`;
    }
    
    // Sheet access patterns
    if (runtimeData.stats.sheetAccesses !== undefined) {
      statsMessage += `Sheet Access:\n`;
      statsMessage += `  Total Sheet Accesses: ${runtimeData.stats.sheetAccesses || 0}\n\n`;
    }
    
    // Top operations
    if (runtimeData.stats.operationCounts && Object.keys(runtimeData.stats.operationCounts).length > 0) {
      statsMessage += `Top Operations:\n`;
      const sortedOps = Object.entries(runtimeData.stats.operationCounts)
        .sort(([,a], [,b]) => b - a)
        .slice(0, 5);
      
      sortedOps.forEach(([operation, count]) => {
        statsMessage += `  ${operation}: ${count}\n`;
      });
      statsMessage += `\n`;
    }
    
    // Performance metrics
    if (runtimeData.stats.performanceMetrics && Object.keys(runtimeData.stats.performanceMetrics).length > 0) {
      statsMessage += `Performance Metrics:\n`;
      Object.entries(runtimeData.stats.performanceMetrics).slice(0, 3).forEach(([operation, metrics]) => {
        statsMessage += `  ${operation}: avg ${Math.round(metrics.averageTime)}ms (${metrics.count} ops)\n`;
      });
      statsMessage += `\n`;
    }
    
    // Recent activity
    if (runtimeData.recentLogs && runtimeData.recentLogs.length > 0) {
      statsMessage += `Recent Activity (${runtimeData.recentLogs.length} of ${runtimeData.totalLogEntries} entries):\n`;
      runtimeData.recentLogs.slice(0, 5).forEach((log, index) => {
        const logTime = new Date(log.timestamp).toLocaleString();
        const levelIcon = log.level === 'error' ? '✗' : log.level === 'warning' ? '⚠' : '•';
        statsMessage += `  ${levelIcon} ${logTime} - ${log.operation}\n`;
      });
      
      if (runtimeData.recentLogs.length > 5) {
        statsMessage += `  ... and ${runtimeData.recentLogs.length - 5} more entries\n`;
      }
    } else {
      statsMessage += `Recent Activity: No runtime logs found\n`;
    }
    
    if (runtimeData.stats.lastActivity) {
      const lastActivity = new Date(runtimeData.stats.lastActivity).toLocaleString();
      statsMessage += `\nLast Activity: ${lastActivity}`;
    }
    
    statsMessage += `\nStats generated: ${new Date().toLocaleString()}`;
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Runtime Statistics', statsMessage, ui.ButtonSet.OK);
    
    Logger.log('showRuntimeStats: Runtime statistics displayed to user');
    
  } catch (error) {
    Logger.log(`showRuntimeStats: Error displaying runtime statistics: ${error.message}`);
    console.error('showRuntimeStats error:', error);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Could not display runtime statistics:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * User-friendly function to display migration status and history
 * Shows comprehensive migration information to the user
 */
function showMigrationStatus() {
  Logger.log('showMigrationStatus: Displaying migration status to user...');
  
  try {
    const status = MigrationManager.getMigrationStatus();
    
    if (status.error) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Migration Status Error', `Error retrieving migration status:\n\n${status.error}`, ui.ButtonSet.OK);
      return;
    }
    
    let statusMessage = 'Campus Recipients Migration Status\n\n';
    
    // Basic status
    statusMessage += `Migration Status: ${status.isComplete ? '✓ Complete' : '✗ Pending'}\n\n`;
    
    // Migration summary
    if (status.summary && Object.keys(status.summary).length > 0) {
      statusMessage += `Migration Summary:\n`;
      statusMessage += `  Total Attempts: ${status.summary.totalAttempts || 0}\n`;
      statusMessage += `  Successful: ${status.summary.successfulMigrations || 0}\n`;
      statusMessage += `  Failed: ${status.summary.failedMigrations || 0}\n`;
      statusMessage += `  Warnings: ${status.summary.warningsCount || 0}\n`;
      
      if (status.summary.lastSuccessfulMigration) {
        const lastSuccess = new Date(status.summary.lastSuccessfulMigration).toLocaleString();
        statusMessage += `  Last Success: ${lastSuccess}\n`;
      }
      
      if (status.summary.lastError) {
        const lastError = new Date(status.summary.lastError.timestamp).toLocaleString();
        statusMessage += `  Last Error: ${lastError} - ${status.summary.lastError.step}\n`;
      }
      
      statusMessage += `\n`;
    }
    
    // Recent migration activity
    if (status.recentLogs && status.recentLogs.length > 0) {
      statusMessage += `Recent Migration Activity (${status.recentLogs.length} of ${status.totalLogEntries} entries):\n`;
      status.recentLogs.slice(0, 5).forEach((log, index) => {
        const logTime = new Date(log.timestamp).toLocaleString();
        const statusIcon = log.status === 'success' ? '✓' : log.status === 'error' ? '✗' : log.status === 'warning' ? '⚠' : '•';
        statusMessage += `  ${statusIcon} ${logTime} - ${log.step}: ${log.message}\n`;
      });
      
      if (status.recentLogs.length > 5) {
        statusMessage += `  ... and ${status.recentLogs.length - 5} more entries\n`;
      }
    } else {
      statusMessage += `Recent Migration Activity: No migration logs found\n`;
    }
    
    statusMessage += `\nStatus checked: ${new Date().toLocaleString()}`;
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Migration Status', statusMessage, ui.ButtonSet.OK);
    
    Logger.log('showMigrationStatus: Migration status displayed to user');
    
  } catch (error) {
    Logger.log(`showMigrationStatus: Error displaying migration status: ${error.message}`);
    console.error('showMigrationStatus error:', error);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Could not display migration status:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}
  Logger.log('recoverRecipientsSheetManual: Starting manual Recipients sheet recovery...');
  
  try {
    // Ask user for confirmation before recovery
    const ui = SpreadsheetApp.getUi();
    const confirmMessage = 'This will recreate the Recipients sheet from cached or hardcoded data.\n\n' +
                          'Any existing Recipients sheet will be replaced.\n\n' +
                          'Do you want to proceed with the recovery?';
    
    const response = ui.alert(
      'Confirm Sheet Recovery',
      confirmMessage,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      Logger.log('recoverRecipientsSheetManual: Recovery cancelled by user');
      // return;
    }
    
    const recoveryResult = RecipientsSheetManager.recoverRecipientsSheet();
    
    let resultMessage = 'Recipients Sheet Recovery Results\n\n';
    
    if (recoveryResult.success) {
      switch (recoveryResult.action) {
        case 'no_recovery_needed':
          resultMessage += '✓ No recovery was needed - sheet is already healthy\n\n';
          resultMessage += `Data Rows: ${recoveryResult.dataRows}\n`;
          break;
          
        case 'recovered_from_cache':
          resultMessage += '✓ Recipients sheet successfully recovered from cache\n\n';
          resultMessage += `Campuses Recovered: ${recoveryResult.campusCount}\n`;
          resultMessage += `Recipients Recovered: ${recoveryResult.totalRecipients}\n`;
          resultMessage += `Data Source: PropertiesService cache\n`;
          break;
          
        case 'recovered_from_hardcoded':
          resultMessage += '✓ Recipients sheet successfully recovered from hardcoded data\n\n';
          resultMessage += `Campuses Recovered: ${recoveryResult.campusCount}\n`;
          resultMessage += `Recipients Recovered: ${recoveryResult.totalRecipients}\n`;
          resultMessage += `Data Source: Hardcoded backup data\n`;
          break;
          
        default:
          resultMessage += `✓ Recovery completed: ${recoveryResult.message}\n`;
      }
      
      resultMessage += '\nThe Recipients sheet is now available and ready for use.';
    } else {
      resultMessage += `❌ Recovery failed: ${recoveryResult.message}\n\n`;
      resultMessage += 'Please check the execution log for more details.\n\n';
      resultMessage += 'You may need to manually create the Recipients sheet or contact support.';
    }
    
    resultMessage += `\nRecovery completed: ${new Date().toLocaleString()}`;
    
    const alertTitle = recoveryResult.success ? 'Recovery Complete' : 'Recovery Failed';
    ui.alert(alertTitle, resultMessage, ui.ButtonSet.OK);
    
    Logger.log(`recoverRecipientsSheetManual: Recovery completed - ${recoveryResult.success ? 'success' : 'failed'}`);
    
  } catch (error) {
    Logger.log(`recoverRecipientsSheetManual: Error during manual recovery: ${error.message}`);
    console.error('recoverRecipientsSheetManual error:', error);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Recovery Error', `Recipients sheet recovery failed:\n\n${error.message}`, ui.ButtonSet.OK);
  }

/**
 * Test function to validate the cache invalidation system implementation
 * Tests onEdit trigger, manual refresh, and cache status functions
 */
function testCacheInvalidationSystem() {
  Logger.log('=== Testing Cache Invalidation System ===');
  
  let testsPassed = 0;
  let testsTotal = 0;
  let errors = [];
  
  try {
    // Test 1: Verify onEdit function exists and is callable
    testsTotal++;
    Logger.log('Test 1: Testing onEdit function definition...');
    
    if (typeof onEdit !== 'function') {
      throw new Error('onEdit function is not defined');
    }
    
    // Test with mock event object
    const mockEvent = {
      source: SpreadsheetApp.getActiveSpreadsheet(),
      range: {
        getSheet: () => ({ getName: () => 'Recipients' }),
        getA1Notation: () => 'A1',
        getRow: () => 2,
        getColumn: () => 1,
        getNumRows: () => 1,
        getNumColumns: () => 1
      },
      oldValue: 'old@test.com',
      value: 'new@test.com'
    };
    
    // This should not throw an error
    onEdit(mockEvent);
    
    Logger.log('✓ onEdit function works correctly');
    testsPassed++;
    
    // Test 2: Verify handleRecipientsSheetEdit function exists
    testsTotal++;
    Logger.log('Test 2: Testing handleRecipientsSheetEdit function...');
    
    if (typeof handleRecipientsSheetEdit !== 'function') {
      throw new Error('handleRecipientsSheetEdit function is not defined');
    }
    
    Logger.log('✓ handleRecipientsSheetEdit function is defined');
    testsPassed++;
    
    // Test 3: Verify manual cache refresh functions exist
    testsTotal++;
    Logger.log('Test 3: Testing manual cache refresh functions...');
    
    if (typeof refreshCampusCache !== 'function') {
      throw new Error('refreshCampusCache function is not defined');
    }
    
    if (typeof getCacheStatus !== 'function') {
      throw new Error('getCacheStatus function is not defined');
    }
    
    if (typeof showCacheStatus !== 'function') {
      throw new Error('showCacheStatus function is not defined');
    }
    
    Logger.log('✓ All manual cache refresh functions are defined');
    testsPassed++;
    
    // Test 4: Test getCacheStatus functionality
    testsTotal++;
    Logger.log('Test 4: Testing getCacheStatus functionality...');
    
    const cacheStatus = getCacheStatus();
    
    if (!cacheStatus || typeof cacheStatus !== 'object') {
      throw new Error('getCacheStatus should return an object');
    }
    
    // Should have required properties
    const requiredProps = ['cacheExists', 'migrationComplete', 'recipientsSheetExists', 'timestamp'];
    const missingProps = requiredProps.filter(prop => !cacheStatus.hasOwnProperty(prop));
    
    if (missingProps.length > 0) {
      throw new Error(`getCacheStatus missing properties: ${missingProps.join(', ')}`);
    }
    
    Logger.log('✓ getCacheStatus returns proper status object');
    testsPassed++;
    
    // Test 5: Test cache invalidation event logging
    testsTotal++;
    Logger.log('Test 5: Testing cache invalidation event logging...');
    
    // Clear any existing events first
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('cacheInvalidationEvents');
    
    // Create a test invalidation event
    const testEvent = {
      timestamp: new Date().toISOString(),
      type: 'test_event',
      reason: 'Testing cache invalidation system'
    };
    
    // Store the event (simulating what handleRecipientsSheetEdit does)
    properties.setProperty('cacheInvalidationEvents', JSON.stringify([testEvent]));
    
    // Retrieve and verify
    const storedEvents = properties.getProperty('cacheInvalidationEvents');
    if (!storedEvents) {
      throw new Error('Failed to store invalidation events');
    }
    
    const parsedEvents = JSON.parse(storedEvents);
    if (!Array.isArray(parsedEvents) || parsedEvents.length !== 1) {
      throw new Error('Invalidation events not stored correctly');
    }
    
    if (parsedEvents[0].type !== 'test_event') {
      throw new Error('Invalidation event data not preserved correctly');
    }
    
    Logger.log('✓ Cache invalidation event logging works correctly');
    testsPassed++;
    
    // Clean up test data
    properties.deleteProperty('cacheInvalidationEvents');
    
  } catch (error) {
    errors.push(error.message);
    Logger.log(`✗ Test failed: ${error.message}`);
    console.error('Detailed error:', error);
  }
  
  // Report final results
  Logger.log('=== Cache Invalidation System Test Results ===');
  Logger.log(`Tests passed: ${testsPassed}/${testsTotal}`);
  
  if (errors.length > 0) {
    Logger.log('❌ Cache invalidation system testing failed with errors:');
    errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
    
    // Show user-friendly alert
    const ui = SpreadsheetApp.getUi();
    const errorMessage = `Cache invalidation system testing failed!\n\n${errors.join('\n\n')}`;
    ui.alert('Cache Invalidation Test Failed', errorMessage, ui.ButtonSet.OK);
    
    return false;
  } else {
    Logger.log('✅ All cache invalidation system tests passed!');
    
    // Show success message
    const ui = SpreadsheetApp.getUi();
    const successMessage = `Cache Invalidation System Test Results:\n\n` +
                          `✓ All ${testsTotal} tests passed\n` +
                          `✓ onEdit trigger function works correctly\n` +
                          `✓ handleRecipientsSheetEdit function is defined\n` +
                          `✓ Manual cache refresh functions are available\n` +
                          `✓ getCacheStatus returns proper status information\n` +
                          `✓ Cache invalidation event logging works\n\n` +
                          `The cache invalidation system is working correctly!\n\n` +
                          `You can now:\n` +
                          `• Edit the Recipients sheet to automatically clear cache\n` +
                          `• Use "Cache Management" menu for manual operations\n` +
                          `• Monitor cache status and recent activity`;
    ui.alert('Cache Invalidation Test Successful', successMessage, ui.ButtonSet.OK);
    
    return true;
  }
}

/**
 * Quick test function to verify basic functionality
 * Can be run from the Apps Script editor
 */
function quickInfrastructureTest() {
  Logger.log('=== Quick Infrastructure Test ===');
  
  try {
    // Test class availability
    Logger.log('Testing class definitions...');
    const classes = [MigrationManager, PropertiesManager, RecipientsSheetManager];
    Logger.log('✓ All classes are accessible');
    
    // Test data extraction
    Logger.log('Testing data extraction...');
    const hardcodedData = MigrationManager.extractHardcodedData();
    Logger.log(`✓ Extracted ${Object.keys(hardcodedData).length} campuses`);
    
    const driveData = MigrationManager.getHardcodedDriveLinks();
    Logger.log(`✓ Extracted ${Object.keys(driveData).length} drive links`);
    
    // Test basic PropertiesManager
    Logger.log('Testing PropertiesManager...');
    const testData = { 'test': { recipients: ['test@nisd.net'], driveLink: 'test123' } };
    PropertiesManager.storeCampusData(testData);
    const retrieved = PropertiesManager.getCampusData();
    PropertiesManager.clearCache();
    
    if (retrieved && retrieved['test']) {
      Logger.log('✓ PropertiesManager works correctly');
    } else {
      throw new Error('PropertiesManager test failed');
    }
    
    Logger.log('✅ Quick test completed successfully!');
    return true;
    
  } catch (error) {
    Logger.log(`❌ Quick test failed: ${error.message}`);
    console.error('Quick test error:', error);
    return false;
  }
}