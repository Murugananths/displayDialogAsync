// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */
// import {SetRuntimeVisibleHelper,SetStartupBehaviorHelper,updateRibbon, loginpopup} from '../utilities/office-apis-helpers';

// var clickEvent;
// (function () {
//   Office.initialize = function (reason) {
//      // If you need to initialize something you can do so here.
//   };
// })();

// function getLoginfirst(event: Office.AddinCommands.Event) {
//   // Your code goes here
//   clickEvent = event;
//   localStorage.setItem("isEvent",event.source.id.toString());    
//    updateRibbon();  
//    loginpopup(); 
// }

// export function showNotification(text)
// {   
//     clickEvent.completed();
// }


// function getLogout(event: Office.AddinCommands.Event) {
//   // Your code goes here
//   localStorage.setItem("isEvent",event.source.id.toString());
  
//    SetRuntimeVisibleHelper(false);  
//    updateRibbon();
//   // Be sure to indicate when the add-in command function is complete
//   event.completed();
// }


// export function getGlobal() { 
//   return (typeof self !== "undefined") ? self :
//     (typeof window !== "undefined") ? window : 
//     (typeof global !== "undefined") ? global :
//     undefined;
// }

// const g = getGlobal() as any;

// // the add-in command functions need to be available in global scope

// // g.getLoginfirst = getLoginfirst;
// // g.getLogout=getLogout;
