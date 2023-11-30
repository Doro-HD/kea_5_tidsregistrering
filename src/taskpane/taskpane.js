/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { async } from "regenerator-runtime";


/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    getTime();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("call").onclick = test;
  }
});

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

    //Tilføj de populated felter (startDate, endDate, startTime, endTime)
    //til et fetch kald, så de kan sendes med til backenden, sammen med projekt ID'et.

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

//This function gets the start and end time of the appointment.
//Made by Troels.
export async function getTime() {

  Office.context.mailbox.item.start.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${result.error.message}`);
      return;
    }

    console.log(`Appointment starts: ${result.value}`);
    document.getElementById("startTime").innerHTML = result.value.toTimeString().split(' ')[0];
    document.getElementById("startDate").innerHTML = result.value.toLocaleDateString();
  });

  Office.context.mailbox.item.end.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${result.error.message}`);
      return;
    }

    console.log(`Appointment ends: ${result.value}`);
    document.getElementById("endTime").innerHTML = result.value.toTimeString().split(' ')[0];
    document.getElementById("endDate").innerHTML = result.value.toLocaleDateString();
  });
}