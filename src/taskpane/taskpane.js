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

//14:54. 22/11/2023. Meget af nedenstående er taget fra Chatgbt
export async function run() {
  const projectId = document.querySelector("input#project-id").value;
  const className = 'your-error-class'; // Replace with your actual error class name

  try {
<<<<<<< HEAD
    const response = await fetch('https://timereg-api.azurewebsites.net/test/' + projectId, {
      method: 'get',
      headers: {
        'Content-Type': 'text/plain',
      },
    });

    if (!response.ok) { // If response status code is an error (4xx or 5xx)
      throw new Error(`HTTP Error! Status: ${response.status}`);
    }

    document.getElementById("returned-message-backend").innerHTML = "Success!!!!!!!!!!!!!";
    return response.json(); // or .text() if the response is not JSON
  } catch (error) {
    // Select all elements with the given class name and set their innerHTML
    console.log(error)
    
    const elements = document.querySelectorAll(`.${className}`);
    elements.forEach(element => {
      element.innerHTML = error.message;
    });
  }
}
=======
    await fetch('https://timereg-api.azurewebsites.net/test/' + projectId, {
      method: 'get',
      headers: {
        'Content-Type': 'text/plain',
      }
    }).then(res => handleHttpErrors(res))

    document.getElementById("returned-message-backend").innerHTML = "Sucess!!!!!!!!!!!!!";
    console.log("Added")
  } catch (err) {
    document.getElementById("returned-message-backend").innerHTML = (err);//.apiError.response
    console.error(err)

  }


  //Error function taget fra vores 3.semester.
  async function handleHttpErrors(res) {
    console.log("I httperrorfunctionen" + res.ok)
    if (!res.ok) {
      const errorResponse = await res.json();
      const error = new Error(errorResponse.message)
      error.apiError = errorResponse
      throw error
    }
    
    return res.json()
  }
}

//DOMPurify.sanitize

>>>>>>> d7e30671f4bb56abdaf54ba32e439de9ea2c6f69

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
