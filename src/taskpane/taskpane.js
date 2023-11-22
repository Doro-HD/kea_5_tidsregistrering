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
  }
});

export async function run() {


  const projectId = document.querySelector("input#project-id").value

  try {
    const data = await fetch('https://timereg-api.azurewebsites.net/test/' + projectId, {
        method: 'post', //Denne skal jo være en post, men end pointed modtager kun get.
        headers: {
          'Content-Type': 'text/plain', // Specify the content type as plain text
        },
    }).then(res => handleHttpErrors(res))

    document.getElementById("returned-message-backend").innerHTML = "";
    console.log("Added " + projectId)
} catch (err) {
    //document.getElementById("returned-message-backend").innerHTML = (err.apiError.response);
    console.error(err)

}
//DOMPurify.sanitize


  // Get a reference to the current message
  // const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  //document.getElementById("returned-message-backend").innerHTML = "<b>Status:</b> Gemt! (Hardcoded) <br/>";

}
/*

//Export function taget fra vores 3.semester.
async function register() {

    const projectId = document.querySelector("input#project-id").value

    //let headers = new Headers()
    //headers.append("Content-Type", "application/json; charset=utf-8")
    //headers.append("Accept", "application/json")

    //const jsonBody = JSON.stringify({username: usernameInput, password: passwordInput})
    
    try {
        const data = await fetch("timereg-api.azurewebsites.net/test/", {
            method: 'post',
            headers: {
              'Content-Type': 'text/plain', // Specify the content type as plain text
            },
            body: projectId
        }).then(res => handleHttpErrors(res))

        document.getElementById("returned-message-backend").innerHTML = "";
        window.router.navigate("/")
        alert("Projekt Id added!")
        console.log("Added")
    } catch (err) {
        document.getElementById("returned-message-backend").innerHTML = DOMPurify.sanitize(err.apiError.response);
        console.error(err)
    
    }
}




}
*/

//Error function taget fra vores 3.semester.
    async function handleHttpErrors(res) {
      console.log("I httperrorfunctionen" + res)
    if (!res.ok) {
      const errorResponse = await res.json();
      const error = new Error(errorResponse.message)
      error.apiError = errorResponse
      throw error
    }
    
    return res.json()
  }