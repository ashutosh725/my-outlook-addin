/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT license.
 *
 * Outlook Executive Email Add-in
 */

/* global document, Office */

Office.onReady(function (info) {

  if (info.host === Office.HostType.Outlook) {

    var button = document.getElementById("sendToFlow");

    if (button) {
      button.onclick = run;
    }

  }

});


export async function run() {

  var statusElement =
    document.getElementById("status");


  try {

    // ==========================================
    // CHECK STATUS ELEMENT
    // ==========================================

    if (!statusElement) {

      console.error(
        "Status element not found."
      );

      return;

    }


    statusElement.textContent =
      "Getting email details...";


    // ==========================================
    // GET OUTLOOK EMAIL
    // ==========================================

    var item =
      Office.context.mailbox.item;


    if (!item) {

      throw new Error(
        "No email is currently open."
      );

    }


    // ==========================================
    // GET OUTLOOK MESSAGE ID
    // ==========================================

    var messageId =
      item.itemId || "";


    console.log(
      "Outlook Message ID:",
      messageId
    );


    if (!messageId) {

      throw new Error(
        "Unable to get Outlook Message ID."
      );

    }


    // ==========================================
    // GET FORM ELEMENTS
    // ==========================================

    var categoryElement =
      document.getElementById("category");


    var priorityElement =
      document.getElementById("priority");


    var dueDateElement =
      document.getElementById("dueDate");


    var instructionsElement =
      document.getElementById("instructions");


    // ==========================================
    // CHECK FORM ELEMENTS
    // ==========================================

    if (
      !categoryElement ||
      !priorityElement ||
      !dueDateElement ||
      !instructionsElement
    ) {

      throw new Error(
        "One or more form fields could not be found."
      );

    }


    // ==========================================
    // GET USER INPUTS
    // ==========================================

    var category =
      categoryElement.value;


    var priority =
      priorityElement.value;


    var dueDate =
      dueDateElement.value;


    var instructions =
      instructionsElement.value.trim();


    // ==========================================
    // VALIDATION
    // ==========================================

    if (!category) {

      statusElement.textContent =
        "Please select a category.";

      return;

    }


    if (!priority) {

      statusElement.textContent =
        "Please select a priority.";

      return;

    }


    if (!dueDate) {

      statusElement.textContent =
        "Please select a due date.";

      return;

    }


    if (!instructions) {

      statusElement.textContent =
        "Please enter additional instructions.";

      return;

    }


    // ==========================================
    // CONSOLE LOG
    // ==========================================

    console.log(
      "Category:",
      category
    );

    console.log(
      "Priority:",
      priority
    );

    console.log(
      "Due Date:",
      dueDate
    );

    console.log(
      "Instructions:",
      instructions
    );

    console.log(
      "Message ID:",
      messageId
    );


    // ==========================================
    // GET EMAIL SUBJECT
    // ==========================================

    var subject =
      item.subject || "";


    console.log(
      "Subject:",
      subject
    );


    // ==========================================
    // GET EMAIL BODY
    // ==========================================

    statusElement.textContent =
      "Getting email body...";


    item.body.getAsync(
      Office.CoercionType.Html,
      async function (result) {

        try {

          // ==========================================
          // CHECK BODY RESULT
          // ==========================================

          if (
            result.status !==
            Office.AsyncResultStatus.Succeeded
          ) {

            statusElement.textContent =
              "Unable to get email body.";

            console.error(
              "Email body error:",
              result.error
            );

            return;

          }


          var body =
            result.value;


          console.log(
            "Email body:",
            body
          );


          // ==========================================
          // POWER AUTOMATE URL
          // ==========================================

          var powerAutomateUrl =
            "https://defaulte1c709c847fe4dc0a35429338962b7.81.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/22/workflows/57d3ef023e9a41eba012f8b6f737002f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QqiZz3uwnxRMLhHqoZhdtbB6FgSjhFOl-PP0fp7k1KQ";


          // ==========================================
          // SEND DATA TO POWER AUTOMATE
          // ==========================================

          statusElement.textContent =
            "Sending to Power Automate...";


          var response =
            await fetch(
              powerAutomateUrl,
              {

                method: "POST",

                headers: {
                  "Content-Type":
                    "application/json"
                },

                body: JSON.stringify({

                  // ==================================
                  // EMAIL INFORMATION
                  // ==================================

                  subject:
                    subject,

                  body:
                    body,


                  // ==================================
                  // OUTLOOK MESSAGE ID
                  // ==================================

                  messageId:
                    messageId,


                  // ==================================
                  // CUSTOM CATEGORY
                  // ==================================

                  category:
                    category,


                  // ==================================
                  // USER INPUT
                  // ==================================

                  priority:
                    priority,

                  dueDate:
                    dueDate,

                  instructions:
                    instructions

                })

              }
            );


          // ==========================================
          // CHECK RESPONSE
          // ==========================================

          if (!response.ok) {

            var errorMessage =
              "Power Automate returned " +
              response.status;


            console.error(
              "Power Automate HTTP Error:",
              response.status
            );


            try {

              var errorText =
                await response.text();


              if (errorText) {

                console.error(
                  "Power Automate response:",
                  errorText
                );

              }

            } catch (readError) {

              console.error(
                "Unable to read error response:",
                readError
              );

            }


            throw new Error(
              errorMessage
            );

          }


          // ==========================================
          // SUCCESS
          // ==========================================

          statusElement.textContent =
            "Successfully sent to Power Automate.";


          console.log(
            "Successfully sent to Power Automate."
          );


          // ==========================================
          // SHOW SENT DATA
          // ==========================================

          console.log(
            "Sent data:",
            {

              subject:
                subject,

              body:
                body,

              messageId:
                messageId,

              category:
                category,

              priority:
                priority,

              dueDate:
                dueDate,

              instructions:
                instructions

            }
          );


          // ==========================================
          // CLEAR FORM
          // ==========================================

          categoryElement.value =
            "";

          priorityElement.value =
            "";

          dueDateElement.value =
            "";

          instructionsElement.value =
            "";


        } catch (error) {

          console.error(
            "Power Automate Error:",
            error
          );


          var errorMessage =
            error && error.message
              ? error.message
              : String(error);


          statusElement.textContent =
            "Power Automate error: " +
            errorMessage;

        }

      }
    );


  } catch (error) {

    console.error(
      "Error:",
      error
    );


    var errorMessage =
      error && error.message
        ? error.message
        : String(error);


    if (statusElement) {

      statusElement.textContent =
        "Error: " +
        errorMessage;

    }

  }

}