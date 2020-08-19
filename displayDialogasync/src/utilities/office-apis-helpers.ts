// import { getGlobal } from '../commands/commands';
import {  showNotification ,getGlobal} from '../taskpane/taskpane';
import * as OfficeHelpers from '@microsoft/office-js-helpers';
export const SetRuntimeVisibleHelper = (visible: boolean) => {
  let p: any;
  if (visible) {
    p = Office.addin.showAsTaskpane();
  } else {
    p = Office.addin.hide();
  }

  return p
    .then(() => {
      return visible;
    })
    .catch(error => {
      return error.code;
    });
};

export const SetStartupBehaviorHelper = (isStarting: boolean) => {
  if (isStarting) {
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } else {
    Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  }
  let g = getGlobal() as any;
  g.isStartOnDocOpen = isStarting;
};

export function updateRibbon() {
  // Update ribbon based on state tracking
  const g = getGlobal() as any;

  // @ts-ignore
  OfficeRuntime.ui
    .getRibbon()
    // @ts-ignore
    .then(ribbon => {
      ribbon.requestUpdate({
        tabs: [
          {
            id: 'Contoso.Tab1',
            controls: [
              {
                id: 'Contoso.LogoutButton',
                enabled: true
              },
              {
                id: 'Contoso.IButton',
                enabled: true
              }]
          }
        ]
      });
    });
}

var dialog;

export async function loginpopup() { 
let g = getGlobal() as any;
if (OfficeHelpers.Authenticator.isAuthDialog())
return;
  Office.context.ui.displayDialogAsync(
    window.location.origin + "/Dialog.html",
    { height: 60, width: 60, promptBeforeOpen: true },
    dialogCallback);
  function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {

      // In addition to general system errors, there are 3 specific errors for 
      // displayDialogAsync that you can handle individually.
      switch (asyncResult.error.code) {
        case 12004:
          console.log("Domain is not trusted");
          break;
        case 12005:
          console.log("HTTPS is required");
          break;
        case 12007:
          console.log("A dialog is already opened.");
          break;
        default:
          console.log(asyncResult.error.message);
          break;
      }
    }
    else {
      dialog = asyncResult.value;
      /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
      

    }
  }


  function messageHandler(arg: any) { 
    // g.state.setConnected(true); 
    // Office.context.ui.messageParent("jsonMessage");
    if (arg.message != "test") {
      // //$(".loader").show();    
      // var accessToken = JSON.parse(arg.message).value.split("#")[1].split("&")[1].split("=")[1];    
      // localStorage.setItem('accessToken', accessToken);  
     

      dialog.close(); 
      showNotification();
    }; 
  
  }
 
}  //end


export function generateCustomFunction(selectedOption: string) {
  try {
    Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      let range = context.workbook.getSelectedRange();

      //let selectedOption = 'Communications';

      range.values = [['=CONTOSOSHARE.GETDATA("' + selectedOption + '")']];
      range.format.autofitColumns();
      return context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

//This will check if state is initialized, and if not, initialize it.
//Useful as there are multiple entry points that need the state and it is not clear which one will get called first.
export async function ensureStateInitialized(isOfficeInitializing: boolean) {
  console.log('ensureInitialize called');
  let g = getGlobal() as any;
  let initValue = false;
  if (isOfficeInitializing) {
    //we are being called in response to Office Initialize
    if (g.state !== undefined) {
      if (g.state.isInitialized === false) {
        g.state.isInitialized = true;
      }
    }
    if (g.state === undefined) {
      initValue = true;
    }
  }

  if (g.state === undefined) {
    g.state = {
      isStartOnDocOpen: false,
      isSignedIn: false,
      isTaskpaneOpen: false,
      isConnected: false,
      isSyncEnabled: false,
      isConnectInProgress: false,
      isFirstSyncCall: true,
      isSumEnabled: false,
      isInitialized: initValue,
      updateRct: () => {},
      setTaskpaneStatus: (opened: boolean) => {
        g.state.isTaskpaneOpen = opened;
        updateRibbon();
      },
      setConnected: (connected: boolean) => {
        g.state.isConnected = connected;

        if (connected) {
          if (g.state.updateRct !== null) {
            g.state.updateRct('true');
          }
        } else {
          if (g.state.updateRct !== null) {
            g.state.updateRct('false');
          }
        }
        updateRibbon();
      }
    };
    // console.log("init value:" + initValue);
    // console.log("is initialized: " + g.state.isInitialized);
    // monitorSheetChanges();

    let addinState = await Office.addin.getStartupBehavior();
    console.log('load state is:');
    console.log('load state' + addinState);
    if (addinState === Office.StartupBehavior.load) {
      g.state.isStartOnDocOpen = true;
    }
    if (localStorage.getItem('loggedIn') === 'yes') {
      g.state.isSignedIn = true;
    }
  }
  updateRibbon();
}

async function onTableSelectionChange(event) {
  let g = getGlobal() as any;
  return Excel.run(context => {
    return context.sync().then(() => {
      console.log('Table section changed...');
      console.log('Change type of event: ' + event.changeType);
      console.log('Address of event: ' + event.address);
      console.log('Source of event: ' + event.source);
      g.state.selectionAddress = event.address;
      if (event.address === '' && g.state.isSumEnabled === true) {
        g.state.isSumEnabled = false;
        updateRibbon();
      } else if (g.state.isSumEnabled === false && event.address !== '') {
        g.state.isSumEnabled = true;
        updateRibbon();
      }
    });
  });
}

export async function monitorSheetChanges() {
  try {
    let g = getGlobal() as any;
    if (g.state === undefined) {
      return;
    }
    if (g.state.isInitialized) {
      await Excel.run(async context => {
        let table = context.workbook.tables.getItem('ExpensesTable');
        if (table !== undefined) {
          table.onSelectionChanged.add(onTableSelectionChange);
          await context.sync();
          updateRibbon();
        } else {
          g.state.isSumEnabled = false;
          updateRibbon();
        }
      });
    }
  } catch (error) {
    console.error(error);
  }
}
