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
  // Get a reference to the current message
 // const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("returned-message-backend").innerHTML = "<b>Status:</b> Gemt! (Hardcoded) <br/>";



/*
Export function taget fra vores 3.semester.
async function register() {

    const usernameInput = document.querySelector("input#username").value
    const passwordInput = document.querySelector("input#password").value

    let headers = new Headers()
    headers.append("Content-Type", "application/json; charset=utf-8")
    headers.append("Accept", "application/json")

    const jsonBody = JSON.stringify({username: usernameInput, password: passwordInput})
    
    try {
        const data = await fetch(baseURL + "/api/register", {
            method: 'post',
            headers: headers,
            body: jsonBody
        }).then(res => handleHttpErrors(res))

        document.getElementById("error-on-register").innerHTML = "";
        window.router.navigate("/")
        alert("Bruger oprettet!")
        console.log("Registreret")
    } catch (err) {
        document.getElementById("error-on-register").innerHTML = DOMPurify.sanitize(err.apiError.response);
        console.error(err)
    
    }

*/





}

export async function testFunction(event) {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById(item.body).innerHTML = "Hello";
}
