/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

let isEventPresent;
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

//Made by Victor and Troels.
async function getEventFromBackend() {

  let eventIdString;
  if (Office.context.mailbox.item.itemId == undefined) {
    eventIdString = await getEventId();
  } else {
    eventIdString = Office.context.mailbox.item.itemId;
  }

  let headers = new Headers()
  headers.append("Content-Type", "application/json; charset=utf-8")
  headers.append("Accept", "application/json")

  const encodedIdString = encodeURIComponent(eventIdString);

  try {
    const response = await fetch(baseURL + "/appointment/query?id=" + encodedIdString, {})

    if (!response.ok) {
      throw new Error()
    }


    const data = await response.json()


    if (data != null) {
      isEventPresent = true;
      if (data.projectId.trim() != "") {
        document.getElementById("grabbed-data").textContent = "Tidlligere bogført projekt ID: " + data.projectId;
        document.getElementById("project-id").value = data.projectId;
      } else {
        document.getElementById("intet-grabbed-data").textContent = "Ingen tidlligere bogført projekt ID.";
      }
    } else {
      isEventPresent = false;
      console.log("No event found");
      document.getElementById("returned-message-backend").textContent = "No event found";
    }


    //console.log(data.projectId);

  } catch (error) {
    console.error("Ingen event fundet i databasen som allerede eksisterer.")
    isEventPresent = false;
  }


}



//Made by Victor, Troels and David.
async function sendJsonDataToBackend() {
  console.log(isEventPresent)
  const values = getInfo();

  const projectid = document.querySelector("input#project-id").value;
  values[4] = projectid;


  let eventIdString;
  if (Office.context.mailbox.item.itemId == undefined) {
    eventIdString = await getEventId();
  } else {
    eventIdString = Office.context.mailbox.item.itemId;
  }

  const jsonBody = JSON.stringify({
    Id: eventIdString, //Aktivitets ID
    //Felterne skal ændres til hvad de hedder i databasen.
    Subject: values[2], //Subjectline/Navn på møde
    UserEmail: values[3], //Email på bruger
    ProjectId: values[4], //Projekt ID
    AppointmentStart: values[0], //Møde start
    AppointmentEnd: values[1], //Møde slut

  })

  let headers = new Headers()
  headers.append("Content-Type", "application/json; charset=utf-8")
  headers.append("Accept", "application/json")


  if (isEventPresent == false) {

    try {
      const data = await fetch(baseURL + "/appointment", {
        method: 'post',
        headers: headers,
        body: jsonBody
      })

      if (!data.ok) {
        throw new Error(`HTTP Error! Status: ${data.status}`)
      }

      const node = document.getElementById("returned-message-backend");
      node.style.color = "green";
      node.textContent = "Successful registreret!";
      console.log("Registreret!")

    } catch (error) {

      console.error(error)

      errorHandler(error);
    }

  } else {

    try {
      const data = await fetch(baseURL + "/appointment", {
        method: 'put',
        headers: headers,
        body: jsonBody
      })

      if (!data.ok) {
        throw new Error(`HTTP Error! Status: ${data.status}`)
      }

      const node = document.getElementById("returned-message-backend");
      node.style.color = "green";
      node.textContent = "Successful opdateret!";
      console.log("Opdateret!")

    } catch (error) {

      console.error(error)

      errorHandler(error);
    }
  }
  //console.log(jsonBody)
  //console.log(eventIdString)
}



//Made by Victor, Troels and David.
function getEventId() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getItemIdAsync(result => {
      resolve(result.value)
    })
  })
}


//Made by Troels, Victor and supervised by David
//En lille metode, der tager en fejl, og viser en besked til brugeren.
//Smed den over i sin egen metode, så den kan genbruges. frem for at skrive den samme kode flere gange.
//This switch case can be expanded to handle more errors.
function errorHandler(error) {
  console.log(error.message.replace(/\D/g, ''));
  const node = document.getElementById("returned-message-backend");
  node.style.color = "red";
  switch (error.message.replace(/\D/g, '')) {
    case "400": node.textContent = "Fejl. Gem aktivitet og prøv igen.";
      break;
    case "404": node.textContent = "Intern server fejl. Prøv igen.";
      break;
    default: node.textContent = "Generel fejl. Prøv igen.";
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
      //console.log(`Appointment starts: ${result.value}`);
      document.getElementById("startDate").textContent = result.value.toLocaleDateString();
      document.getElementById("startTime").textContent = result.value.toTimeString().split(' ')[0];
    });

    Office.context.mailbox.item.end.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        return;
      }

      values[1] = result.value;
      //console.log(`Appointment ends: ${result.value}`);
      document.getElementById("endTime").textContent = result.value.toTimeString().split(' ')[0];
      document.getElementById("endDate").textContent = result.value.toLocaleDateString();
    });
    //========================================================================================


    //Henter titlen på mødet
    //========================================================================================
    Office.context.mailbox.item.subject.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        return;
      }
      //console.log(`Appointment subject: ${result.value}`);
      values[2] = document.getElementById("subjectLine").textContent = result.value;
    });
    //========================================================================================

    //Henter mødelederens email
    //========================================================================================
    Office.context.mailbox.item.organizer.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        return;
      }
      //console.log(`Appointment organizer: ${result.value}`);
      values[3] = document.getElementById("emailAddress").textContent = result.value.emailAddress;
    });
    //========================================================================================
  } else {

    //Sætter felterne til at være de samme som mødet, hvis det er et møde man er inviteret til.
    document.getElementById("startTime").textContent = Office.context.mailbox.item.end.toTimeString().split(' ')[0];
    document.getElementById("startDate").textContent = Office.context.mailbox.item.end.toLocaleDateString();
    document.getElementById("endTime").textContent = Office.context.mailbox.item.end.toTimeString().split(' ')[0];
    document.getElementById("endDate").textContent = Office.context.mailbox.item.end.toLocaleDateString();
    values[0] = Office.context.mailbox.item.start;
    values[1] = Office.context.mailbox.item.end;
    values[2] = document.getElementById("subjectLine").textContent = Office.context.mailbox.item.subject;
    values[3] = document.getElementById("emailAddress").textContent = Office.context.mailbox.userProfile.emailAddress;

  }

  //console.log(values);
  return values;
} 