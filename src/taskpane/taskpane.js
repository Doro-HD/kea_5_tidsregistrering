/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const baseURL = "https://timereg-api.azurewebsites.net"


Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("testeventid").onclick = getCalendarEventIdAfterSave;
  }
});



//Made by Victor, Troels and David.
async function getCalendarEventIdAfterSave() {
  const eventIdString = await myTestFunction();

  let headers = new Headers()
  headers.append("Content-Type", "application/json; charset=utf-8")
  headers.append("Accept", "application/json")

  const jsonBody = JSON.stringify({ eventid: eventIdString })

  try {
    const data = await fetch(baseURL + "/appointment", {
      method: 'post',
      headers: headers,
      body: jsonBody
    })

    if (!data.ok) {
      throw new Error(`HTTP Error! Status: ${response.status}`)
    }

    document.getElementById("returned-message-backend").innerHTML = "Successful registreret";
    console.log("Registreret")

  } catch (error) {

    console.error(error)

    errorHandler(error);
  }

  console.log(eventIdString)
}

//Made by Victor, Troels and David.
function myTestFunction() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getItemIdAsync(result => {
      resolve(result.value)
    })
  })
}


//14:54. 22/11/2023. A large portion of this function has been taken from Chatgbt.
//Victor has edited large sections of this function so it fits our needs.
export async function run() { //Run fucntion to send the project ID to the backend. And show a respond feedback to user.
  const projectId = document.querySelector("input#project-id").value;

  try {

    const response = await fetch(baseURL + '/test/' + projectId, {
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
    errorHandler(error);
  }
}


function errorHandler(error) {
  switch (error.message.replace(/\D/g, '')) {
    case "400": document.getElementById("returned-message-backend").innerHTML = "Fejl i projekt ID. Prøv igen";
      break;
    case "404": document.getElementById("returned-message-backend").innerHTML = "Intern server fejl. Prøv igen";
      break;
    default: document.getElementById("returned-message-backend").innerHTML = "Genneral fejl. IK prøv igen";
      break;
  }
}