/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("call").onclick = test;
    document.getElementById("testeventid").onclick = getCalendarEventIdAfterSave;
  }
});

// Function to get the Calendar Event ID after saving the event
//Made by Chatgbt.
function getCalendarEventIdAfterSave() {
  Office.context.mailbox.item.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          const item = Office.context.mailbox.item;
          //const myvarbasedvar = item.getItemIdAsync()
          //console.log(myvarbasedvar)
          console.log(item)
          if (item) {
              console.log("Calendar Event ID: " + item);
              // Further processing with eventId
          } else {
              console.error("Event ID not available even after save." + item);
              // Handle the case where itemId is not available
          }
      } else {
          console.error("Error during save: ", result.error);
      }
  });
}

// ... Rest of your existing code for 'test' and 'run' functions ...



//David made this.
//Test function to see if the frontend can communicate with the backend.
async function test() {
  const res = await fetch("https://timereg-api.azurewebsites.net/hello")
  const data = await res.json()

  const node = document.getElementById("returned-message-backend")
  node.innerHTML = data.value
}

//14:54. 22/11/2023. A large portion of this function has been taken from Chatgbt.
//Victor has edited large sections of this function so it fits our needs.
export async function run() { //Run fucntion to send the project ID to the backend. And show a respond feedback to user.
  const projectId = document.querySelector("input#project-id").value;

  try {

    const response = await fetch('https://timereg-api.azurewebsites.net/test/' + projectId, {
    });

    if (!response.ok) { // If response status code is an error (4xx or 5xx)
       
      throw new Error(`HTTP Error! Status: ${response.status}`);
    }

    document.getElementById("returned-message-backend").innerHTML = "Success!!!!!!!!!!!!!";
    return response.json(); // or .text() if the response is not JSON
  } catch (error) {
    
    // Select all elements with the given class name and set their innerHTML
    console.log(error)

    //Troels made this switch case.
    //This switch case can be expanded to handle more errors.
    switch (error.message.replace(/\D/g, '')) {
      case "400": document.getElementById("returned-message-backend").innerHTML = "Fejl i projekt ID. Prøv igen";
        break;
      case "404": document.getElementById("returned-message-backend").innerHTML = "Intern server fejl. Prøv igen";
        break;
      default: document.getElementById("returned-message-backend").innerHTML = "Genneral fejl. IK prøv igen";
        break;
    }
  }
}