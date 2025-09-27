/**
 * Scheduler Management
 * Handles time-driven triggers for dataset refreshes
 */

// Trigger name prefix
const TRIGGER_PREFIX = 'QBO_Dataset_';
const TRIGGER_FUNCTION = 'scheduledDatasetRun';

/**
 * Creates a schedule trigger for a dataset
 */
function createScheduleTrigger(datasetId, schedule) {
  try {
    if (!schedule || !schedule.enabled) {
      return null;
    }
    
    // Remove existing trigger if any
    removeScheduleTrigger(datasetId);
    
    // Create new trigger based on frequency
    const triggerBuilder = ScriptApp.newTrigger(TRIGGER_FUNCTION);
    let trigger;
    
    switch (schedule.freq) {
      case 'hourly':
        trigger = triggerBuilder
          .timeBased()
          .everyHours(1)
          .create();
        break;
        
      case 'daily':
        const dailyHour = parseInt(schedule.timeOfDay.split(':')[0]) || 9;
        trigger = triggerBuilder
          .timeBased()
          .everyDays(1)
          .atHour(dailyHour)
          .create();
        break;
        
      case 'weekly':
        const weeklyHour = parseInt(schedule.timeOfDay.split(':')[0]) || 9;
        const dayOfWeek = schedule.dayOfWeek || ScriptApp.WeekDay.MONDAY;
        trigger = triggerBuilder
          .timeBased()
          .onWeekDay(dayOfWeek)
          .atHour(weeklyHour)
          .create();
        break;
        
      case 'monthly':
        const monthlyHour = parseInt(schedule.timeOfDay.split(':')[0]) || 9;
        const dayOfMonth = schedule.dayOfMonth || 1;
        trigger = triggerBuilder
          .timeBased()
          .onMonthDay(dayOfMonth)
          .atHour(monthlyHour)
          .create();
        break;
        
      default:
        throw new Error(`Unsupported schedule frequency: ${schedule.freq}`);
    }
    
    // Store trigger ID with dataset ID mapping
    const triggerId = trigger.getUniqueId();
    PropertiesService.getUserProperties().setProperty(
      TRIGGER_PREFIX + datasetId,
      triggerId
    );
    
    // Also store reverse mapping for trigger handler
    PropertiesService.getUserProperties().setProperty(
      'trigger_' + triggerId,
      datasetId
    );
    
    logAction('create_schedule_trigger', {
      dataset_id: datasetId,
      schedule_freq: schedule.freq,
      schedule_time: schedule.timeOfDay,
      trigger_id: triggerId
    });
    
    return triggerId;
  } catch (error) {
    console.error('Error creating schedule trigger:', error);
    logAction('create_schedule_trigger_error', {
      dataset_id: datasetId,
      error: error.toString()
    });
    throw error;
  }
}

/**
 * Removes a schedule trigger for a dataset
 */
function removeScheduleTrigger(datasetId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const triggerId = userProperties.getProperty(TRIGGER_PREFIX + datasetId);
    
    if (!triggerId) {
      return false;
    }
    
    // Find and delete the trigger
    const triggers = ScriptApp.getProjectTriggers();
    let found = false;
    
    triggers.forEach(trigger => {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        found = true;
      }
    });
    
    // Clean up properties
    userProperties.deleteProperty(TRIGGER_PREFIX + datasetId);
    userProperties.deleteProperty('trigger_' + triggerId);
    
    logAction('remove_schedule_trigger', {
      dataset_id: datasetId,
      trigger_id: triggerId,
      found: found
    });
    
    return found;
  } catch (error) {
    console.error('Error removing schedule trigger:', error);
    logAction('remove_schedule_trigger_error', {
      dataset_id: datasetId,
      error: error.toString()
    });
    return false;
  }
}

/**
 * Handler function called by time-driven triggers
 */
function scheduledDatasetRun(e) {
  const startTime = new Date();
  let datasetId = null;
  let triggerId = null;
  
  try {
    // Get trigger ID from event
    triggerId = e ? e.triggerUid : null;
    
    if (!triggerId) {
      console.error('No trigger ID in event');
      return;
    }
    
    // Get dataset ID from trigger mapping
    const userProperties = PropertiesService.getUserProperties();
    datasetId = userProperties.getProperty('trigger_' + triggerId);
    
    if (!datasetId) {
      console.error('No dataset ID found for trigger:', triggerId);
      return;
    }
    
    // Get dataset
    const dataset = getDatasetById(datasetId);
    if (!dataset) {
      console.error('Dataset not found:', datasetId);
      // Remove orphaned trigger
      removeScheduleTrigger(datasetId);
      return;
    }
    
    // Check if schedule is still enabled
    if (!dataset.schedule || !dataset.schedule.enabled) {
      console.log('Schedule disabled for dataset:', datasetId);
      removeScheduleTrigger(datasetId);
      return;
    }
    
    // Check for realm-level lock to prevent concurrent runs
    const lockService = LockService.getUserLock();
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    const lockKey = 'qbo_realm_' + realmId;
    
    // Try to acquire lock (wait up to 10 seconds)
    const hasLock = lockService.tryLock(10000);
    
    if (!hasLock) {
      logAction('scheduled_run_skipped', {
        dataset_id: datasetId,
        dataset_name: dataset.name,
        reason: 'Could not acquire realm lock',
        trigger_id: triggerId
      });
      return;
    }
    
    try {
      // Log scheduled run start
      logAction('scheduled_run_start', {
        dataset_id: datasetId,
        dataset_name: dataset.name,
        dataset_type: dataset.type,
        schedule_freq: dataset.schedule.freq,
        trigger_id: triggerId,
        schedule_owner_email: Session.getActiveUser().getEmail()
      });
      
      // Run the dataset
      const jobResult = runDataset(datasetId);
      
      // Log scheduled run complete
      const elapsedMs = new Date().getTime() - startTime.getTime();
      logAction('scheduled_run_complete', {
        dataset_id: datasetId,
        dataset_name: dataset.name,
        dataset_type: dataset.type,
        schedule_freq: dataset.schedule.freq,
        trigger_id: triggerId,
        job_id: jobResult.id,
        elapsed_ms: elapsedMs,
        status: jobResult.status,
        rows: jobResult.result ? jobResult.result.lastWrite.rows : 0,
        cols: jobResult.result ? jobResult.result.lastWrite.cols : 0
      });
      
    } finally {
      // Release lock
      lockService.releaseLock();
    }
    
  } catch (error) {
    console.error('Error in scheduled dataset run:', error);
    
    const elapsedMs = new Date().getTime() - startTime.getTime();
    logAction('scheduled_run_error', {
      dataset_id: datasetId,
      trigger_id: triggerId,
      error_message: error.toString(),
      elapsed_ms: elapsedMs,
      stack: error.stack
    });
    
    // Send error notification if configured
    if (datasetId) {
      notifyScheduleError(datasetId, error);
    }
  }
}

/**
 * Gets all scheduled datasets
 */
function getScheduledDatasets() {
  try {
    const datasets = getDatasets();
    return datasets.filter(d => d.schedule && d.schedule.enabled);
  } catch (error) {
    console.error('Error getting scheduled datasets:', error);
    return [];
  }
}

/**
 * Gets schedule status for UI
 */
function getScheduleStatus() {
  try {
    const datasets = getDatasets();
    const triggers = ScriptApp.getProjectTriggers();
    const userProperties = PropertiesService.getUserProperties();
    
    const MAX_TIME_TRIGGERS = 20;
    const clockTriggers = triggers.filter(t => t.getHandlerFunction() === TRIGGER_FUNCTION);
    
    const status = {
      totalDatasets: datasets.length,
      scheduledDatasets: 0,
      activeTriggers: clockTriggers.length,
      maxTriggers: MAX_TIME_TRIGGERS,
      remainingTriggers: Math.max(0, MAX_TIME_TRIGGERS - clockTriggers.length),
      schedules: [],
      orphanedSchedules: 0
    };
    
    datasets.forEach(dataset => {
      if (dataset.schedule && dataset.schedule.enabled) {
        status.scheduledDatasets++;
        
        const triggerId = userProperties.getProperty(TRIGGER_PREFIX + dataset.id);
        const trigger = triggers.find(t => t.getUniqueId() === triggerId);
        
        const triggerExists = !!trigger;
        status.schedules.push({
          datasetId: dataset.id,
          datasetName: dataset.name,
          frequency: dataset.schedule.freq,
          timeOfDay: dataset.schedule.timeOfDay,
          dayOfWeek: dataset.schedule.dayOfWeek || null,
          dayOfMonth: dataset.schedule.dayOfMonth || null,
          triggerId: triggerId,
          triggerExists: triggerExists,
          nextRun: trigger ? getNextRunTime(trigger) : null
        });

        if (!triggerExists) {
          status.orphanedSchedules++;
        }
      }
    });
    
    status.limitWarning = status.remainingTriggers <= 2;
    status.limitExceeded = status.remainingTriggers === 0;
    
    return status;
  } catch (error) {
    console.error('Error getting schedule status:', error);
    return {
      error: error.toString()
    };
  }
}

/**
 * Gets next run time for a trigger
 */
function getNextRunTime(trigger) {
  try {
    // This is an approximation as Apps Script doesn't provide exact next run time
    const eventType = trigger.getEventType();
    const now = new Date();
    
    if (eventType === ScriptApp.EventType.CLOCK) {
      // For time-based triggers, we can make educated guesses
      // This is simplified - actual implementation would need more logic
      return 'Next scheduled run';
    }
    
    return 'Unknown';
  } catch (error) {
    return 'Error determining next run';
  }
}

/**
 * Validates schedule configuration
 */
function validateSchedule(schedule) {
  if (!schedule) {
    return { valid: false, error: 'Schedule configuration required' };
  }
  
  const validFrequencies = ['hourly', 'daily', 'weekly', 'monthly'];
  if (!validFrequencies.includes(schedule.freq)) {
    return { valid: false, error: `Invalid frequency. Must be one of: ${validFrequencies.join(', ')}` };
  }
  
  if (schedule.freq !== 'hourly' && !schedule.timeOfDay) {
    return { valid: false, error: 'Time of day required for daily, weekly, and monthly schedules' };
  }
  
  if (schedule.timeOfDay) {
    const timeRegex = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/;
    if (!timeRegex.test(schedule.timeOfDay)) {
      return { valid: false, error: 'Invalid time format. Use HH:MM (24-hour format)' };
    }
  }
  
  if (schedule.freq === 'weekly' && schedule.dayOfWeek) {
    const validDays = Object.values(ScriptApp.WeekDay);
    if (!validDays.includes(schedule.dayOfWeek)) {
      return { valid: false, error: 'Invalid day of week' };
    }
  }
  
  if (schedule.freq === 'monthly' && schedule.dayOfMonth) {
    const day = parseInt(schedule.dayOfMonth);
    if (isNaN(day) || day < 1 || day > 31) {
      return { valid: false, error: 'Invalid day of month. Must be between 1 and 31' };
    }
  }
  
  return { valid: true };
}

/**
 * Sends error notification for failed scheduled run
 */
function notifyScheduleError(datasetId, error) {
  try {
    const dataset = getDatasetById(datasetId);
    if (!dataset) return;
    
    const userEmail = Session.getActiveUser().getEmail();
    const notificationEnabled = PropertiesService.getUserProperties().getProperty('SCHEDULE_ERROR_NOTIFICATIONS') === 'true';
    
    if (!notificationEnabled || !userEmail) {
      return;
    }
    
    // Log notification attempt
    logAction('schedule_error_notification', {
      dataset_id: datasetId,
      dataset_name: dataset.name,
      user_email: userEmail,
      error_message: error.toString()
    });
    
    // In a real implementation, you might send an email or create a notification
    // For now, we just log it
    console.log('Schedule error notification:', {
      dataset: dataset.name,
      error: error.toString()
    });
    
  } catch (notifyError) {
    console.error('Error sending schedule error notification:', notifyError);
  }
}

/**
 * Manually triggers a scheduled dataset (for testing)
 */
function testScheduledRun(datasetId) {
  try {
    const dataset = getDatasetById(datasetId);
    if (!dataset) {
      throw new Error('Dataset not found');
    }
    
    // Simulate trigger event
    const mockEvent = {
      triggerUid: 'test-trigger-' + Utilities.getUuid()
    };
    
    // Temporarily store mapping
    PropertiesService.getUserProperties().setProperty(
      'trigger_' + mockEvent.triggerUid,
      datasetId
    );
    
    try {
      // Run the scheduled function
      scheduledDatasetRun(mockEvent);
      
      return {
        success: true,
        message: 'Test scheduled run completed'
      };
    } finally {
      // Clean up temporary mapping
      PropertiesService.getUserProperties().deleteProperty(
        'trigger_' + mockEvent.triggerUid
      );
    }
  } catch (error) {
    console.error('Error in test scheduled run:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Cleans up orphaned triggers
 */
function cleanupOrphanedTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const userProperties = PropertiesService.getUserProperties();
    const datasets = getDatasets();
    const datasetMap = {};
    const scheduledDatasets = [];
    const triggerMap = {};
    let removed = 0;
    let repaired = 0;

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === TRIGGER_FUNCTION) {
        triggerMap[trigger.getUniqueId()] = trigger;
      }
    });

    datasets.forEach(dataset => {
      datasetMap[dataset.id] = dataset;
      const propertyKey = TRIGGER_PREFIX + dataset.id;
      if (dataset.schedule && dataset.schedule.enabled) {
        scheduledDatasets.push(dataset);
      } else {
        const existingTriggerId = userProperties.getProperty(propertyKey);
        if (existingTriggerId) {
          const existingTrigger = triggerMap[existingTriggerId];
          if (existingTrigger) {
            ScriptApp.deleteTrigger(existingTrigger);
            delete triggerMap[existingTriggerId];
            removed++;
          }
          userProperties.deleteProperty(propertyKey);
          userProperties.deleteProperty('trigger_' + existingTriggerId);
        }
      }
    });

    Object.keys(triggerMap).forEach(triggerId => {
      const trigger = triggerMap[triggerId];
      const datasetId = userProperties.getProperty('trigger_' + triggerId);
      const dataset = datasetId ? datasetMap[datasetId] : null;

      if (!dataset || !dataset.schedule || !dataset.schedule.enabled) {
        ScriptApp.deleteTrigger(trigger);
        userProperties.deleteProperty('trigger_' + triggerId);
        if (datasetId) {
          userProperties.deleteProperty(TRIGGER_PREFIX + datasetId);
        }
        delete triggerMap[triggerId];
        removed++;
      }
    });

    scheduledDatasets.forEach(dataset => {
      const propertyKey = TRIGGER_PREFIX + dataset.id;
      const storedTriggerId = userProperties.getProperty(propertyKey);
      const triggerExists = storedTriggerId && triggerMap[storedTriggerId];

      if (!triggerExists) {
        const newTriggerId = createScheduleTrigger(dataset.id, dataset.schedule);
        if (newTriggerId) {
          repaired++;
          triggerMap[newTriggerId] = null;
        }
      }
    });

    logAction('cleanup_orphaned_triggers', {
      removed: removed,
      repaired: repaired
    });

    return { removed: removed, repaired: repaired };
  } catch (error) {
    console.error('Error cleaning up orphaned triggers:', error);
    logAction('cleanup_orphaned_triggers_error', { error: error.toString() });
    return { removed: 0, repaired: 0, error: error.toString() };
  }
}
