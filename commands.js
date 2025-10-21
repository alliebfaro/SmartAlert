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

// Get sender information as a Promise
function getFromAddress(item) {
  return new Promise((resolve, reject) => {
    item.from.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({
          name: result.value.displayName,
          email: result.value.emailAddress
        });
      } else {
        reject(result.error);
      }
    });
  });
}

// Get all recipients (To, Cc, Bcc) as a Promise
function getRecipients(item) {
  return new Promise((resolve, reject) => {
    const recipientPromises = [];
    
    // Get To recipients
    recipientPromises.push(
      new Promise((res) => {
        item.to.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            res({ type: 'To', recipients: result.value });
          } else {
            res({ type: 'To', recipients: [] });
          }
        });
      })
    );
    
    // Get Cc recipients
    recipientPromises.push(
      new Promise((res) => {
        item.cc.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            res({ type: 'Cc', recipients: result.value });
          } else {
            res({ type: 'Cc', recipients: [] });
          }
        });
      })
    );
    
    // Get Bcc recipients (if available)
    if (item.bcc) {
      recipientPromises.push(
        new Promise((res) => {
          item.bcc.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              res({ type: 'Bcc', recipients: result.value });
            } else {
              res({ type: 'Bcc', recipients: [] });
            }
          });
        })
      );
    }
    
    Promise.all(recipientPromises).then(results => {
      resolve(results);
    }).catch(reject);
  });
}

// Show the confirmation dialog with timer
function showSendConfirmationDialog(fromInfo, recipientsData, itemType, event) {
  const dialogData = {
    fromName: fromInfo.name,
    fromEmail: fromInfo.email,
    recipients: recipientsData,
    itemType: itemType,
    autoSendSeconds: 20
  };
  
  // Open the dialog
  Office.context.ui.displayDialogAsync(
    'https://alliebfaro.github.io/SmartAlert/sendConfirmDialog.html',
    { height: 60, width: 45, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('Dialog failed to open:', asyncResult.error.message);
        // If dialog fails, allow send to continue
        event.completed({ allowEvent: true });
        return;
      }
      
      const dialog = asyncResult.value;
      let dialogMessageSent = false;
      
      // Handle messages from dialog
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();
        
        try {
          const response = JSON.parse(arg.message);
          
          if (response.action === 'send') {
            // User clicked Send or timer expired
            console.log('Send confirmed:', response.reason);
            event.completed({ allowEvent: true });
          } else if (response.action === 'cancel') {
            // User clicked Cancel
            console.log('Send cancelled by user');
            event.completed({ 
              allowEvent: false, 
              errorMessage: 'Send cancelled by user.' 
            });
          }
        } catch (error) {
          console.error('Error parsing dialog response:', error);
          event.completed({ allowEvent: true });
        }
      });
      
      // Handle dialog being closed without response
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.log('Dialog event:', arg);
        if (arg.error === 12006) { // Dialog closed by user
          dialog.close();
          event.completed({ 
            allowEvent: false, 
            errorMessage: 'Send cancelled.' 
          });
        }
      });
      
      // Send data to dialog after a brief delay to ensure it's ready
      setTimeout(() => {
        if (!dialogMessageSent) {
          dialog.messageChild(JSON.stringify(dialogData));
          dialogMessageSent = true;
        }
      }, 500);
    }
  );
}

// Register the functions
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);

// Make functions available globally
if (typeof global !== 'undefined') {
  global.onMessageSendHandler = onMessageSendHandler;
  global.onAppointmentSendHandler = onAppointmentSendHandler;
}
