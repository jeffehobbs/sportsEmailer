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
  const item = Office.context.mailbox.item;
  const sender = item.from.emailAddress;  // Update to get sender's email address
  const senderName = item.from.displayName; // Update to get sender's name
  const recipients = item.to; // This is an array of recipients
  const mainRecipient = recipients[0].emailAddress; // This is the main recipient's email
  const mainRecipientName = recipients[0].displayName; // This is the main recipient's name

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Generated email:</b> <br/>" + item.subject;
  item.body.getAsync(Office.CoercionType.Text, async function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Make a call to the server to get metrics
      try {
        // Create a JSON object containing the email to send
        const emailData = {
          email: result.value, 
          mainRecipient,
          mainRecipientName,
          sender,
          senderName
        };

        // Make a POST request to the /metrics endpoint
        const response = await fetch(`http://localhost:5000/metrics`, {
          method: 'POST', // Set method to POST
          headers: {
            'Content-Type': 'application/json' // Set content type to application/json
          },
          body: JSON.stringify(emailData) // Stringify the email data to send in the request body
        });

        const data = await response.json(); // Parse the JSON response
        const metrics = data.metrics; // Extract metrics from the response

        // Display the metrics in the task pane
        document.getElementById("item-body").innerHTML = metrics;
      } catch (error) {
        console.error("Error making API call:", error);
      }

    } else {
      console.error("Error getting email body: " + result.error.message);
    }
  });
}


document.getElementById('copyButton').addEventListener('click', function () {
  // Create a range and selection to select the text block
  var range = document.createRange();
  var selection = window.getSelection();

  // Clear current selection if any
  selection.removeAllRanges();

  // Select the text content of textBlock element
  range.selectNodeContents(document.getElementById('textBlock'));

  // Add the new range
  selection.addRange(range);

  try {
    // Execute the copy command
    var successful = document.execCommand('copy');
    var msg = successful ? 'successful' : 'unsuccessful';
    console.log('Copy command was ' + msg);
  } catch (err) {
    console.log('Oops, unable to copy');
  }

  // Remove the selections - NOTE: Should be done after the copy command
  selection.removeAllRanges();
});
