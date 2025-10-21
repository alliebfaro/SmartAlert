// Office.js On Send Event Handlers for Messages and Appointments
// Shows sender and recipients with 20-second auto-send timer

Office.initialize = function () {
  console.log('Office Add-in initialized');
};

// Handler for OnMessageSend event
function onMessageSendHandler(event) {
  console.log('OnMessageSend triggered');
  const item = Office.context.mailbox.item;
  
  // Gather all required information
  Promise.all([
    getFromAddress(item),
    getRecipients(item)
  ]).then(([fromInfo, recipients]) => {
    showSendConfirmationDialog(fromInfo, recipients, 'email', event);
  }).catch(error => {
    console.error('Error gathering email info:', error);
    // On error, allow send to proceed
    event.completed({ allowEvent: true });
  });
}

// Handler for OnAppointmentSend event
function onAppointmentSendHandler(event) {
  console.log('OnAppointmentSend triggered');
  const item = Office.context.mailbox.item;
  
  // Gather all required information
  Promise.all([
    getFromAddress(item),
    getRecipients(item)
  ]).then(([fromInfo, recipients]) => {
    showSendConfirmationDialog(fromInfo, recipients, 'appointment', event);
  }).catch(error => {
    console.error('Error gathering appointment info:', error);
    // On error, allow send to proceed
    event.completed({ allowEvent: true });
  });
}

// Helper to ensure event.completed is only called once
function tryCompleteEvent(event, completionArgs) {
  try {
    if (!event || typeof event.completed !== 'function') {
      console.warn('tryCompleteEvent: invalid event object', event);
      return;
    }
    if (!event._completed) {
      event._completed = true;
      event.completed(completionArgs);
    } else {
      console.warn('Event already completed, skipping duplicate completion.');
    }
  } catch (err) {
    console.error('Error completing event:', err);
  }
}

// Get sender information as a Promise (defensive)
function getFromAddress(item) {
  return new Promise((resolve) => {
    try {
      if (!item) {
        resolve({ name: 'Unknown', email: '' });
        return;
      }

      // Some hosts expose a simple from object
      if (item.from && (typeof item.from === 'object') && !item.from.getAsync) {
        resolve({
          name: item.from.displayName || item.from.name || 'Unknown',
          email: item.from.emailAddress || item.from.address || ''
        });
        return;
      }

      // If API is available, use it
      if (item.from && typeof item.from.getAsync === 'function') {
        item.from.getAsync((result) => {
          try {
            if (result && result.status === Office.AsyncResultStatus.Succeeded && result.value) {
              resolve({
                name: result.value.displayName || 'Unknown',
                email: result.value.emailAddress || ''
              });
            } else {
              console.warn('getFromAddress: getAsync did not succeed', result && result.error);
              resolve({ name: 'Unknown', email: '' });
            }
          } catch (err) {
            console.error('getFromAddress inner error:', err);
            resolve({ name: 'Unknown', email: '' });
          }
        });
        return;
      }

      // Fallback
      resolve({ name: 'Unknown', email: '' });
    } catch (err) {
      console.error('getFromAddress unexpected error:', err);
      resolve({ name: 'Unknown', email: '' });
    }
  });
}

// Get all recipients (To, Cc, Bcc) as a Promise (defensive)
function getRecipients(item) {
  return new Promise((resolve) => {
    try {
      const recipientPromises = [];

      const pushIfApi = (propName, typeLabel) => {
        if (item && item[propName] && typeof item[propName].getAsync === 'function') {
          recipientPromises.push(new Promise((res) => {
            item[propName].getAsync((result) => {
              if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
                res({ type: typeLabel, recipients: result.value });
              } else {
                res({ type: typeLabel, recipients: [] });
              }
            });
          }));
        } else {
          // API not available for this property -> resolve empty
          recipientPromises.push(Promise.resolve({ type: typeLabel, recipients: [] }));
        }
      };

      // Message-style recipients
      pushIfApi('to', 'To');
      pushIfApi('cc', 'Cc');
      pushIfApi('bcc', 'Bcc');

      // Appointment-style attendees (if present)
      if (item && item.requiredAttendees && typeof item.requiredAttendees.getAsync === 'function') {
        recipientPromises.push(new Promise((res) => {
          item.requiredAttendees.getAsync((result) => {
            if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
              res({ type: 'Required', recipients: result.value });
            } else {
              res({ type: 'Required', recipients: [] });
            }
          });
        }));
      }

      if (item && item.optionalAttendees && typeof item.optionalAttendees.getAsync === 'function') {
        recipientPromises.push(new Promise((res) => {
          item.optionalAttendees.getAsync((result) => {
            if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
              res({ type: 'Optional', recipients: result.value });
            } else {
              res({ type: 'Optional', recipients: [] });
            }
          });
        }));
      }

      Promise.all(recipientPromises).then(results => resolve(results)).catch(err => {
        console.error('getRecipients Promise.all error:', err);
        resolve([]);
      });
    } catch (err) {
      console.error('getRecipients unexpected error:', err);
      resolve([]);
    }
  });
}

// Show the confirmation dialog with timer (safe messaging + single completion)
function showSendConfirmationDialog(fromInfo, recipientsData, itemType, event) {
  const dialogData = {
    fromName: fromInfo.name,
    fromEmail: fromInfo.email,
    recipients: recipientsData,
    itemType: itemType,
    autoSendSeconds: 20
  };

  const dialogUrl = new URL('sendConfirmDialog.html', window.location.href).href;
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 60, width: 45, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('Dialog failed to open:', asyncResult.error && asyncResult.error.message);
        tryCompleteEvent(event, { allowEvent: true });
        return;
      }

      const dialog = asyncResult.value;
      let messageSent = false;

      const finish = (allow, opts) => {
        tryCompleteEvent(event, Object.assign({ allowEvent: !!allow }, opts || {}));
        try { dialog.close(); } catch (e) { /* ignore */ }
      };

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        try {
          const response = JSON.parse(arg.message);
          if (response.action === 'send') {
            console.log('Send confirmed:', response.reason);
            finish(true);
          } else if (response.action === 'cancel') {
            console.log('Send cancelled by user');
            finish(false, { errorMessage: 'Send cancelled by user.' });
          } else {
            console.warn('Unknown dialog response, allowing send', response);
            finish(true);
          }
        } catch (err) {
          console.error('Error parsing dialog response:', err);
          finish(true);
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.log('Dialog event received:', arg);
        // treat any dialog event as cancellation fallback
        finish(false, { errorMessage: 'Send cancelled (dialog closed).' });
      });

      // Retry messaging a few times
      let attempts = 0;
      const tryMessageChild = () => {
        try {
          dialog.messageChild(JSON.stringify(dialogData));
          messageSent = true;
        } catch (err) {
          attempts++;
          if (attempts < 5) {
            setTimeout(tryMessageChild, 250);
          } else {
            console.warn('Failed to message dialog child, allowing send as fallback');
            finish(true);
          }
        }
      };
      tryMessageChild();
    }
  );
}

// Register the functions
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
//Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);

// Make functions available globally
if (typeof global !== 'undefined') {
  global.onMessageSendHandler = onMessageSendHandler;
//  global.onAppointmentSendHandler = onAppointmentSendHandler;
}
