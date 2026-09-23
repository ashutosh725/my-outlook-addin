/*
 * Outlook Executive Email Add-in
 *
 * Handles:
 * - Category based field visibility
 * - Conditional form validation
 * - Outlook email subject/body/message ID
 * - Outlook Conversation ID
 * - Sending data to Power Automate
 * - Success / Error status styling
 */

 /* global document, Office */


// ==========================================================
// OFFICE READY
// ==========================================================

Office.onReady(function (info) {

  if (info.host !== Office.HostType.Outlook) {
    return;
  }


  // ========================================================
  // SEND BUTTON
  // ========================================================

  var button =
    document.getElementById("sendToFlow");

  if (button) {
    button.onclick = run;
  }


  // ========================================================
  // CATEGORY ELEMENTS
  // ========================================================

  var categoryElement =
    document.getElementById("category");

  var dueDateContainer =
    document.getElementById("dueDateContainer");

  var waitingFields =
    document.getElementById("waitingFields");

  var respondFromElement =
    document.getElementById("respondFrom");

  var daysElement =
    document.getElementById("days");

  var instructionsRequiredStar =
    document.getElementById("instructionsRequiredStar");


  // ========================================================
  // UPDATE FIELD VISIBILITY
  // ========================================================

  if (
    categoryElement &&
    dueDateContainer &&
    waitingFields
  ) {

    function updateCategoryFields() {

      var selectedCategory =
        categoryElement.value;


      // ====================================================
      // DUE DATE VISIBILITY
      // ====================================================

      /*
       * Show Due Date:
       * - Tim's To-Dos
       * - Rachel's To-Dos
       *
       * Hide Due Date:
       * - Tim — Please Read
       * - Waiting on Others
       */

      if (
        selectedCategory === "Tim — Please Read" ||
        selectedCategory === "Waiting on Others"
      ) {

        dueDateContainer.style.display =
          "none";

      } else {

        dueDateContainer.style.display =
          "block";
      }


      // ====================================================
      // WAITING ON OTHERS FIELDS
      // ====================================================

      /*
       * Show:
       * - Respond From
       * - Days
       *
       * Only for:
       * - Waiting on Others
       */

      if (
        selectedCategory === "Waiting on Others"
      ) {

        waitingFields.style.display =
          "block";

      } else {

        waitingFields.style.display =
          "none";


        // Clear fields when category changes

        if (respondFromElement) {
          respondFromElement.value = "";
        }

        if (daysElement) {
          daysElement.value = "";
        }
      }


      // ====================================================
      // ADDITIONAL INSTRUCTIONS STAR
      // ====================================================

      /*
       * Additional Instructions required only for:
       *
       * - Tim's To-Dos
       * - Rachel's To-Dos
       *
       * Optional for:
       *
       * - Tim — Please Read
       * - Waiting on Others
       */

      var isTaskCategory =
        selectedCategory === "Tim's To-Dos" ||
        selectedCategory === "Rachel's To-Dos";


      if (instructionsRequiredStar) {

        instructionsRequiredStar.style.display =
          isTaskCategory
            ? "inline"
            : "none";
      }
    }


    // ======================================================
    // CATEGORY CHANGE EVENT
    // ======================================================

    categoryElement.addEventListener(
      "change",
      updateCategoryFields
    );


    // ======================================================
    // INITIAL STATE
    // ======================================================

    updateCategoryFields();
  }
});


// ==========================================================
// STATUS HELPERS
// ==========================================================

function showSuccess(message) {

  var statusElement =
    document.getElementById("status");

  if (!statusElement) {
    return;
  }

  statusElement.textContent =
    message;

  statusElement.style.display =
    "block";

  statusElement.style.color =
    "#107c10";

  statusElement.style.backgroundColor =
    "#e8f5e9";

  statusElement.style.border =
    "1px solid #a5d6a7";

  statusElement.style.padding =
    "10px";

  statusElement.style.borderRadius =
    "4px";
}


function showError(message) {

  var statusElement =
    document.getElementById("status");

  if (!statusElement) {
    return;
  }

  statusElement.textContent =
    message;

  statusElement.style.display =
    "block";

  statusElement.style.color =
    "#d13438";

  statusElement.style.backgroundColor =
    "#fde7e9";

  statusElement.style.border =
    "1px solid #f1aeb5";

  statusElement.style.padding =
    "10px";

  statusElement.style.borderRadius =
    "4px";
}


function showInfo(message) {

  var statusElement =
    document.getElementById("status");

  if (!statusElement) {
    return;
  }

  statusElement.textContent =
    message;

  statusElement.style.display =
    "block";

  statusElement.style.color =
    "#333333";

  statusElement.style.backgroundColor =
    "#f3f3f3";

  statusElement.style.border =
    "1px solid #d0d0d0";

  statusElement.style.padding =
    "10px";

  statusElement.style.borderRadius =
    "4px";
}


// ==========================================================
// MAIN FUNCTION
// ==========================================================

export async function run() {

  var statusElement =
    document.getElementById("status");


  try {

    // ======================================================
    // CHECK STATUS ELEMENT
    // ======================================================

    if (!statusElement) {

      console.error(
        "Status element not found."
      );

      return;
    }


    // Clear previous status

    statusElement.textContent =
      "";

    statusElement.style.display =
      "none";


    showInfo(
      "Getting email details..."
    );


    // ======================================================
    // GET CURRENT OUTLOOK EMAIL
    // ======================================================

    var item =
      Office.context.mailbox.item;


    if (!item) {

      throw new Error(
        "No email is currently open."
      );
    }


    // ======================================================
    // GET OUTLOOK MESSAGE ID
    // ======================================================

    var messageId =
      item.itemId || "";


    console.log(
      "Outlook Message ID:",
      messageId
    );


    // ======================================================
    // GET OUTLOOK CONVERSATION ID
    // ======================================================

    var conversationId =
      item.conversationId || "";


    console.log(
      "Outlook Conversation ID:",
      conversationId
    );


    if (!messageId) {

      throw new Error(
        "Unable to get Outlook Message ID."
      );
    }


    // ======================================================
    // GET FORM ELEMENTS
    // ======================================================

    var categoryElement =
      document.getElementById("category");

    var priorityElement =
      document.getElementById("priority");

    var dueDateElement =
      document.getElementById("dueDate");

    var instructionsElement =
      document.getElementById("instructions");

    var respondFromElement =
      document.getElementById("respondFrom");

    var daysElement =
      document.getElementById("days");


    // ======================================================
    // CHECK FORM ELEMENTS
    // ======================================================

    if (
      !categoryElement ||
      !priorityElement ||
      !dueDateElement ||
      !instructionsElement ||
      !respondFromElement ||
      !daysElement
    ) {

      throw new Error(
        "One or more form fields could not be found."
      );
    }


    // ======================================================
    // GET USER INPUTS
    // ======================================================

    var category =
      categoryElement.value.trim();

    var priority =
      priorityElement.value.trim();

    var dueDate =
      dueDateElement.value.trim();

    var instructions =
      instructionsElement.value.trim();

    var respondFrom =
      respondFromElement.value.trim();

    var days =
      daysElement.value.trim();


    // ======================================================
    // CATEGORY VALIDATION
    // ======================================================

    if (!category) {

      showError(
        "Please select a category."
      );

      return;
    }


    // ======================================================
    // PRIORITY VALIDATION
    // ======================================================

    if (!priority) {

      showError(
        "Please select a priority."
      );

      return;
    }


    // ======================================================
    // DUE DATE VALIDATION
    // ======================================================

    /*
     * Due Date required for:
     *
     * - Tim's To-Dos
     * - Rachel's To-Dos
     *
     * Due Date NOT required for:
     *
     * - Tim — Please Read
     * - Waiting on Others
     */

    if (
      category !== "Tim — Please Read" &&
      category !== "Waiting on Others" &&
      !dueDate
    ) {

      showError(
        "Please select a due date."
      );

      return;
    }


    // ======================================================
    // WAITING ON OTHERS VALIDATION
    // ======================================================

    if (
      category === "Waiting on Others"
    ) {

      // ----------------------------------------------------
      // Conversation ID required
      // ----------------------------------------------------

      if (!conversationId) {

        showError(
          "Unable to get email conversation ID."
        );

        console.error(
          "Conversation ID is empty."
        );

        return;
      }


      // ----------------------------------------------------
      // Respond From required
      // ----------------------------------------------------

      if (!respondFrom) {

        showError(
          "Please enter Respond From."
        );

        return;
      }


      // ----------------------------------------------------
      // Email format validation
      // ----------------------------------------------------

      var emailRegex =
        /^[^\s@]+@[^\s@]+\.[^\s@]+$/;


      if (!emailRegex.test(respondFrom)) {

        showError(
          "Please enter a valid email address."
        );

        return;
      }


      // ----------------------------------------------------
      // Days required
      // ----------------------------------------------------

      if (!days) {

        showError(
          "Please enter number of days."
        );

        return;
      }


      // ----------------------------------------------------
      // Days must be greater than 0
      // ----------------------------------------------------

      if (Number(days) <= 0) {

        showError(
          "Days must be greater than 0."
        );

        return;
      }


      // ----------------------------------------------------
      // Days must be a valid number
      // ----------------------------------------------------

      if (
        !Number.isFinite(
          Number(days)
        )
      ) {

        showError(
          "Please enter a valid number of days."
        );

        return;
      }
    }


    // ======================================================
    // INSTRUCTIONS VALIDATION
    // ======================================================

    /*
     * Required ONLY for:
     *
     * - Tim's To-Dos
     * - Rachel's To-Dos
     *
     * Optional for:
     *
     * - Tim — Please Read
     * - Waiting on Others
     */

    var isTaskCategory =
      category === "Tim's To-Dos" ||
      category === "Rachel's To-Dos";


    if (
      isTaskCategory &&
      !instructions
    ) {

      showError(
        "Please enter Additional Instructions."
      );

      return;
    }


    // ======================================================
    // CONSOLE LOG
    // ======================================================

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
      "Respond From:",
      respondFrom
    );

    console.log(
      "Days:",
      days
    );

    console.log(
      "Message ID:",
      messageId
    );

    console.log(
      "Conversation ID:",
      conversationId
    );


    // ======================================================
    // GET EMAIL SUBJECT
    // ======================================================

    var subject =
      item.subject || "";


    console.log(
      "Subject:",
      subject
    );


    // ======================================================
    // GET EMAIL BODY
    // ======================================================

    showInfo(
      "Getting email body..."
    );


    item.body.getAsync(
      Office.CoercionType.Html,
      async function (result) {

        try {

          // ==================================================
          // CHECK BODY RESULT
          // ==================================================

          if (
            result.status !==
            Office.AsyncResultStatus.Succeeded
          ) {

            showError(
              "Unable to get email body."
            );

            console.error(
              "Email body error:",
              result.error
            );

            return;
          }


          // ==================================================
          // EMAIL BODY
          // ==================================================

          var body =
            result.value;


          console.log(
            "Email body:",
            body
          );


          // ==================================================
          // POWER AUTOMATE URL
          // ==================================================

          /*
           * IMPORTANT:
           *
           * Replace this value with your current
           * Power Automate HTTP trigger URL.
           *
           * Do not commit the URL with the "sig"
           * parameter to a public repository.
           */

          var powerAutomateUrl =
            "https://defaulte1c709c847fe4dc0a35429338962b7.81.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/22/workflows/57d3ef023e9a41eba012f8b6f737002f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QqiZz3uwnxRMLhHqoZhdtbB6FgSjhFOl-PP0fp7k1KQ";


          // ==================================================
          // SEND DATA TO POWER AUTOMATE
          // ==================================================

          showInfo(
            "Sending to Power Automate..."
          );


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

                  // ========================================
                  // EMAIL INFORMATION
                  // ========================================

                  subject:
                    subject,

                  body:
                    body,


                  // ========================================
                  // OUTLOOK MESSAGE ID
                  // ========================================

                  messageId:
                    messageId,


                  // ========================================
                  // OUTLOOK CONVERSATION ID
                  // ========================================

                  conversationId:
                    conversationId,


                  // ========================================
                  // CATEGORY
                  // ========================================

                  category:
                    category,


                  // ========================================
                  // USER INPUT
                  // ========================================

                  priority:
                    priority,

                  dueDate:
                    dueDate,

                  instructions:
                    instructions,


                  // ========================================
                  // WAITING ON OTHERS
                  // ========================================

                  respondFrom:
                    respondFrom,

                  days:
                    days
                })
              }
            );


          // ==================================================
          // CHECK POWER AUTOMATE RESPONSE
          // ==================================================

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


          // ==================================================
          // SUCCESS
          // ==================================================

          showSuccess(
            "Successfully sent to Power Automate."
          );


          console.log(
            "Successfully sent to Power Automate."
          );


          // ==================================================
          // SHOW SENT DATA
          // ==================================================

          console.log(
            "Sent data:",
            {

              subject:
                subject,

              body:
                body,

              messageId:
                messageId,

              conversationId:
                conversationId,

              category:
                category,

              priority:
                priority,

              dueDate:
                dueDate,

              instructions:
                instructions,

              respondFrom:
                respondFrom,

              days:
                days
            }
          );


          // ==================================================
          // CLEAR FORM
          // ==================================================

          categoryElement.value =
            "";

          priorityElement.value =
            "";

          dueDateElement.value =
            "";

          instructionsElement.value =
            "";

          respondFromElement.value =
            "";

          daysElement.value =
            "";


          // ==================================================
          // RESET FIELD VISIBILITY
          // ==================================================

          var dueDateContainer =
            document.getElementById(
              "dueDateContainer"
            );

          var waitingFields =
            document.getElementById(
              "waitingFields"
            );

          var instructionsRequiredStar =
            document.getElementById(
              "instructionsRequiredStar"
            );


          if (dueDateContainer) {

            dueDateContainer.style.display =
              "block";
          }


          if (waitingFields) {

            waitingFields.style.display =
              "none";
          }


          if (instructionsRequiredStar) {

            instructionsRequiredStar.style.display =
              "none";
          }


        } catch (error) {

          // ==================================================
          // POWER AUTOMATE ERROR
          // ==================================================

          console.error(
            "Power Automate Error:",
            error
          );


          var errorMessage =
            error && error.message
              ? error.message
              : String(error);


          showError(
            "Power Automate error: " +
            errorMessage
          );
        }
      }
    );


  } catch (error) {

    // ======================================================
    // GENERAL ERROR
    // ======================================================

    console.error(
      "Error:",
      error
    );


    var errorMessage =
      error && error.message
        ? error.message
        : String(error);


    showError(
      "Error: " +
      errorMessage
    );
  }
}