/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

let ifEventIsPresent;
const baseURL = "https://timereg-api.azurewebsites.net"


Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    getInfo();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("the-event-id").onclick = sendJsonDataToBackend;
    getEventFromBackend();
  }
});


async function getEventFromBackend(){

  let headers = new Headers()
  headers.append("Content-Type", "application/json; charset=utf-8")
  headers.append("Accept", "application/json")

  const jsonBody = JSON.stringify({
    Id: eventIdString, //Aktivitets ID

  })

  try {
    const data = await fetch(baseURL + "/appointment", {
      method: 'get',
      headers: headers,
      body: jsonBody
    })

    if (!data.ok) {
      throw new Error(`HTTP Error! Status: ${response.status}`)
    }

    //Unpack json object here
    document.getElementById("returned-message-backend").innerHTML = "Successful registreret";

  } catch (error) {
    console.error(error)
    errorHandler(error);
  }

}



//Made by Victor, Troels and David.
async function sendJsonDataToBackend() {

  const values = getInfo();

  const projectid = document.querySelector("input#project-id").value;
  values[4] = projectid;


  console.log(values);

  let eventIdString;
  if (Office.context.mailbox.item.itemId == undefined) {
    eventIdString = await getEventId();
  } else {
    eventIdString = Office.context.mailbox.item.itemId;
  }


  let headers = new Headers()
  headers.append("Content-Type", "application/json; charset=utf-8")
  headers.append("Accept", "application/json")




  const jsonBody = JSON.stringify({
    Id: eventIdString, //Aktivitets ID
    //Felterne skal ændres til hvad de hedder i databasen.
    Subject: values[2], //Subjectline/Navn på møde
    UserEmail: values[3], //Email på bruger
    ProjectId: values[4], //Projekt ID
    AppointmentStart: values[0], //Møde start
    AppointmentEnd: values[1], //Møde slut

  })

  console.log(jsonBody)

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
function getEventId() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getItemIdAsync(result => {
      resolve(result.value)
    })
  })
}


//Made by Troels.
//En lille metode, der tager en fejl, og viser en besked til brugeren.
//Smed den over i sin egen metode, så den kan genbruges. frem for at skrive den samme kode flere gange.
//This switch case can be expanded to handle more errors.
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
function getInfo() {

  const values = [];

  if (Office.context.mailbox.item.itemId == undefined) {

    //Henter start og slut tidspunk på mødet
    //========================================================================================
    Office.context.mailbox.item.start.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        return;
      }

      values[0] = result.value;
      console.log(`Appointment starts: ${result.value}`);
      /* values[0] =  */document.getElementById("startTime").innerHTML = result.value.toTimeString().split(' ')[0];
      /* values[1] =  */document.getElementById("startDate").innerHTML = result.value.toLocaleDateString();
    });

    Office.context.mailbox.item.end.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        return;
      }

      values[1] = result.value;
      console.log(`Appointment ends: ${result.value}`);
      /* values[2] =  */document.getElementById("endTime").innerHTML = result.value.toTimeString().split(' ')[0];
      /* values[3] =  */document.getElementById("endDate").innerHTML = result.value.toLocaleDateString();
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
      values[2] = document.getElementById("subjectLine").innerHTML = result.value;
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
      values[3] = document.getElementById("emailAddress").innerHTML = result.value.emailAddress;
    });
    //========================================================================================
  } else {

    //Sætter felterne til at være de samme som mødet, hvis det er et møde man er inviteret til.
    document.getElementById("startTime").innerHTML = Office.context.mailbox.item.end.toTimeString().split(' ')[0];
    document.getElementById("startDate").innerHTML = Office.context.mailbox.item.end.toLocaleDateString();
    document.getElementById("endTime").innerHTML = Office.context.mailbox.item.end.toTimeString().split(' ')[0];
    document.getElementById("endDate").innerHTML = Office.context.mailbox.item.end.toLocaleDateString();
    values[0] = Office.context.mailbox.item.start;
    values[1] = Office.context.mailbox.item.end;
    values[2] = document.getElementById("subjectLine").innerHTML = Office.context.mailbox.item.subject;
    values[3] = document.getElementById("emailAddress").innerHTML = Office.context.mailbox.userProfile.emailAddress;

  }

  console.log(values);
  return values;
} 