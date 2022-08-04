/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { SharedMap } from "fluid-framework";
import { InsecureTokenProvider } from "@fluidframework/test-client-utils";
import { AzureClient, AzureConnectionConfig, AzureFunctionTokenProvider, LOCAL_MODE_TENANT_ID } from "@fluidframework/azure-client";

export const userValueKey = "user-value-key";
export const qValueKey = "q-value-key";



(function () {
	"use strict";

  Office.onReady(function(info) {
      console.trace("sample trace");

      const clientProps = {
        connection: {
            // tenantId: LOCAL_MODE_TENANT_ID,
            // tenantId: "52029d34-7724-49b0-9dd8-fb73f760454d",
            tenantId: "b482b942-933a-4437-b904-ad1487a46590",
            tokenProvider: new InsecureTokenProvider("7d7af34b9651252249d0c3676243d2b2", { id: "userId" }),
            // tokenProvider: new AzureFunctionTokenProvider("https://fluidaddin.azurewebsites.net" + "/api/GetAzureToken", { }),
            orderer: "https://alfred.westus2.fluidrelay.azure.com",
            storage: "https://historian.westus2.fluidrelay.azure.com",
            // orderer: "http://localhost:7070",
            // storage: "http://localhost:7070",
        },
      };
      
      const client = new AzureClient(clientProps);
      
      
      
      const containerSchema = {
          initialObjects: { userID: SharedMap, nameList: SharedMap, answerList: SharedMap, prompt: SharedMap }
      };
      const root = document.getElementById("content");
      
      const createNewDice = async () => {
          const { container, services } = await client.createContainer(containerSchema);
          container.initialObjects.userID.set(userValueKey, 1);
          container.initialObjects.prompt.set(qValueKey, "");
          console.log("New User: " + container.initialObjects.userID.get(userValueKey));
          const id = await container.attach();
          setTimeout(saveSession(id, root), 10000);
          renderDiceRoller(id, root, container.initialObjects.userID, container.initialObjects.nameList, container.initialObjects.answerList, container.initialObjects.prompt);
          return id;
      }
      
      const loadExistingDice = async (id) => {
          const { container } = await client.getContainer(id, containerSchema);
          const participants = container.initialObjects.userID.get(userValueKey) + 1;

          await container.initialObjects.userID.set(userValueKey, participants);

          console.log("New User: " + container.initialObjects.userID.get(userValueKey));
          renderDiceRoller(id, root, container.initialObjects.userID, container.initialObjects.nameList, container.initialObjects.answerList, container.initialObjects.prompt);
      }
      
      async function start() {
        let sid = Office.context.document.settings.get('session');
        console.log(sid);

        if (sid != undefined) {
          console.log("not undefined");
          console.log("sid2 = " + sid);
          await loadExistingDice(sid);
        } else {
          console.log("undefined");
          const id = await createNewDice();

          location.hash = id;
        }

        // if (location.hash) {
        //   await loadExistingDice(location.hash.substring(1))
        // } else {
        //   const id = await createNewDice();
        //   location.hash = id;
        // }
      }
      
      start().catch((error) => console.error(error));
      
      
      // Define the view
      
      const template = document.createElement("template");
      
      template.innerHTML = `
        </br>
        <div class="seshInfo">
          <div class="sesh"></div>
          <div class="userInfo"></div>
          <div class="success"></div>
        </div>
        <div class="teachview">
          <button class="super"> Teacher? </button>
          </br>
          <div class="role"></div>
          </br>
          <textarea class="qinput" rows="5" cols="20"></textarea>
          </br>
          <button class="post"> Post </button>
        </div>
        <div class="wrapper">
          <div class="question"></div>
        </div>
        <div class="wrapper">
          </br>
          <textarea class="nameinput" rows="1" cols="25" placeholder="Enter your name"></textarea>
          </br>
          <textarea class="textinput" rows="10" cols="50" placeholder="Enter your answer"></textarea>
          </br>
          <button class="submit"> Submit </button>
          </br>
          <div class="answer"></div>
          </br>
          </br>
          </br>
          <div class="ansTitle"><b>Answers</b></div>
          </br>
        </div>
        <div class="name-table">
          <table>
            <tr>
              <td>
                <b>Name</b>
              </td>
              <td>
                <div class="nom-1"></div>
              </td>
              <td>
                <div class="nom-2"></div>
              </td>
              <td>
                <div class="nom-3"></div>
              </td>
              <td>
                <div class="nom-4"></div>
              </td>
              <td>
                <div class="nom-5"></div>
              </td>
              <td>
                <div class="nom-6"></div>
              </td>
              <td>
                <div class="nom-7"></div>
              </td>
              <td>
                <div class="nom-8"></div>
              </td>
              <td>
                <div class="nom-9"></div>
              </td>
            </tr>
          </table>
        </div>
        <div class="table1">
          <table>
            <tr>
              <td>
                <b>Answer</b>
              </td>
              <td>
                <div class="ans-1"></div>
              </td>
              <td>
                <div class="ans-2"></div>
              </td>
              <td>
                <div class="ans-3"></div>
              </td>
              <td>
                <div class="ans-4"></div>
              </td>
              <td>
                <div class="ans-5"></div>
              </td>
              <td>
                <div class="ans-6"></div>
              </td>
              <td>
                <div class="ans-7"></div>
              </td>
              <td>
                <div class="ans-8"></div>
              </td>
              <td>
                <div class="ans-9"></div>
              </td>
            </tr>
          </table>
        </div>
      `
      
      const renderDiceRoller = (id, elem, userID, nameList, answerList, prompt) => {

          const partNum = userID.get(userValueKey).toString();

          elem.appendChild(template.content.cloneNode(true));

          const session = elem.querySelector(".sesh");
          session.textContent = "Session ID: " + id;

          const userInfo = elem.querySelector(".userInfo");
          userInfo.textContent = "User Number: " + partNum;

          const teacherButton = elem.querySelector(".super");
          const role = elem.querySelector(".role");

          const submitButton = elem.querySelector(".submit");
          const textbox = elem.querySelector(".textinput");

          const namebox = elem.querySelector(".nameinput");

          const question = elem.querySelector(".question");
          const qinput = elem.querySelector(".qinput");
          const postButton = elem.querySelector(".post");

          qinput.style.display = "none";
          postButton.style.display = "none";

          teacherButton.onclick = () => updateTeacher();

          const updateTeacher = () => {
            role.textContent = "Teacher Mode";
            qinput.style.display = "block";
            postButton.style.display = "block";
          };

          submitButton.onclick = () => {
            nameList.set(partNum, namebox.value);
            answerList.set(partNum, textbox.value);
          };

          const updateText = () => {
            for(let i = 1; i <= 9; i++) {
              const myText = elem.querySelector(".ans-" + i.toString());
              const myName = elem.querySelector(".nom-" + i.toString());
              myText.textContent = answerList.get(i.toString());
              myName.textContent = nameList.get(i.toString());
            };
          };
          updateText();

          postButton.onclick = () => prompt.set(qValueKey, qinput.value)

          const updateQuestion = () => {
            question.textContent = prompt.get(qValueKey);
          };
          updateQuestion();
      
          // Use the changed event to trigger the rerender whenever the value changes.
          answerList.on("valueChanged", updateText);
          prompt.on("valueChanged", updateQuestion);
      }

  });

  function saveSession(id, elem) {
    Office.context.document.settings.set('session', id);
    Office.context.document.settings.saveAsync(function (asyncResult) {
      console.log('Settings saved with status: ' + asyncResult.status);
      elem.querySelector(".success").textContent = 'Syncing ' + asyncResult.status;
    });
  }

})();
