/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

const localUrl = "http://localhost:5040"
const baseURL = "https://timereg-api.azurewebsites.net"


Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    getInfo();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("the-event-id").onclick = getCalendarEventIdAfterSave;
  }
});

//Made by Victor, Troels and David.
async function getCalendarEventIdAfterSave() {
  const eventIdString = await getEventId();

  let headers = new Headers()
  headers.append("Content-Type", "application/json; charset=utf-8")
  headers.append("Accept", "application/json")

  const jsonBody = JSON.stringify({
    id: eventIdString,
    name: "TestStringBody"
  })

  try {
    const data = await fetch(localUrl + "/appointment", {
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
function getEventId() {
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
    errorHandler(error);
  }
}

//Made by Troels.
//En lille metode, der tager en fejl, og viser en besked til brugeren.
//Smed den over i sin egen metode, så den kan genbruges. frem for at skrive den samme kode flere gange.
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


//Made by Troels.
async function getInfo() {

  //Henter start og slut tidspunk på mødet
  //========================================================================================
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
  //========================================================================================


  //Henter titlen på mødet
  //========================================================================================
  Office.context.mailbox.item.subject.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${result.error.message}`);
      return;
    }
    console.log(`Appointment subject: ${result.value}`);
    document.getElementById("subjectLine").innerHTML = result.value;
  });
  //========================================================================================

  //Henter mødelederens email
  //========================================================================================
  Office.context.mailbox.item.organizer.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${result.error.message}`);
      return;
    }
    console.log(`Appointment organizer: ${result.value}`);
    document.getElementById("emailAddress").innerHTML = result.value.emailAddress;
  });
  //========================================================================================
}