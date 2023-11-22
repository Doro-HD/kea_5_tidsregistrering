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
  }
});

async function test() {
  const res = await fetch("https://timereg-api.azurewebsites.net/hello")
  const data = await res.json()

  const node = document.querySelector("#item-subject")
  node.textContent = data.value
}

//14:54. 22/11/2023. Meget af nedenstående er taget fra Chatgbt
export async function run() {
  const projectId = document.querySelector("input#project-id").value;

  try {

    const response = await fetch('https://timereg-api.azurewebsites.net/test/' + projectId, {
      /* method: 'get', // Usually how the fetch call should look like
      headers: {
        'Content-Type': 'text/plain',
      }, */
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

  //Error function taget fra vores 3.semester.
/*   async function handleHttpErrors(res) {
    console.log("I httperrorfunctionen" + res.ok)
    if (!res.ok) {
      const errorResponse = await res.json();
      const error = new Error(errorResponse.message)
      error.apiError = errorResponse
      throw error
    }
    
    return res.json()
  } */


//DOMPurify.sanitize

  // Get a reference to the current message
  // const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  //document.getElementById("returned-message-backend").innerHTML = "<b>Status:</b> Gemt! (Hardcoded) <br/>";


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

    }).then(res => handleHttpErrors(res))

    document.getElementById("returned-message-backend").innerHTML = "Sucess!!!!!!!!!!!!!";
    console.log("Added")
} catch (err) {
    console.log("This is the error number:" + err.message)
    document.getElementById("returned-message-backend").innerHTML = (err);//.apiError.response
    console.error(err)
//DOMPurify.sanitize
}


}
*/
