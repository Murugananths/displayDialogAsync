/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// import {getGlobal} from '../commands/commands';
// import AuthService from '../AuthService/AuthService';
import {SetRuntimeVisibleHelper,SetStartupBehaviorHelper,updateRibbon, loginpopup} from '../utilities/office-apis-helpers';
import * as OfficeHelpers from '@microsoft/office-js-helpers';
// The initialize function must be run each time a new page is loaded

var clickEvent; 
(function () {
  Office.initialize = function (reason) { 
    $("document").ready(function () {
      // showNotification();
    });
  };
})();

// async function SilentRenew() {

//   loginpopup();

//   // tkn.signinSilentCallback();   
// }


async function getLoginfirst(event: Office.AddinCommands.Event) {
  // Your code goes here
   
 
  clickEvent = event;  
  loginpopup();
  localStorage.setItem("isEvent",event.source.id.toString());  
   SetRuntimeVisibleHelper(true);  
   updateRibbon();  

  event.completed();

} 


export function showNotification()
{
    // writeToDoc(text);
    //Required, call event.completed to let the platform know you are done processing.
    clickEvent.completed(true);
}
 

function getLogout(event: Office.AddinCommands.Event) {
  // Your code goes here
  localStorage.setItem("isEvent",event.source.id.toString());
 
   SetRuntimeVisibleHelper(true);   
   updateRibbon();
  
   
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}


const g = getGlobal() as any;
g.getLoginfirst = getLoginfirst;
g.getLogout=getLogout;

 export function getGlobal() { 
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window : 
    (typeof global !== "undefined") ? global :
    undefined;
}




