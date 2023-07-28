/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import * as $ from "jquery";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

Office.onReady(function () {
  // Ensure the DOM is ready
  $(document).ready(function () {
    // Wire up the button click event
    $("#searchButton").click(searchWord);
  });
  $("#searchInput").keyup(function (event) {
    if (event.keyCode === 13) {
      $("#searchButton").click(searchWord);
    }
  });
});

function searchWord() {
  // Get the search term from the input field
  var searchTerm = $("#searchInput").val().trim();

  // Exit early if the search term is empty
  if (!searchTerm) {
    return;
  }

  // Call the Word API to search for the word
  Word.run(function (context) {
    // Get the current document body
    var body = context.document.body;

    // Load the body and search for the word
    context.load(body, "text");
    return context.sync().then(function () {
      // Perform case-insensitive search
      var regex = new RegExp("\\b" + searchTerm + "\\b", "gi");
      var matches = body.text.match(regex);

      // Display the number of entries
      if (matches && matches.length > 0) {
        $("#result").text("Number of entries: " + matches.length);
      } else {
        $("#result").text("No matches found.");
      }
    });
  }).catch(function (error) {
    // Handle errors
    $("#result").text("Error: " + error.message);
  });
}
