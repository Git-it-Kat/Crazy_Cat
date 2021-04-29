/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("addSignatures").onclick = addSign;
     //console.log ("added new code");

    // Create a "close" button and append it to each list item
    var myNodelist = document.getElementsByTagName("LI");
    var i;

    for (i = 0; i < myNodelist.length; i++) {
      var span = document.createElement("SPAN");
      //hex num represents "x"
      var txt = document.createTextNode("\u00D7");
      span.className = "close";
      span.appendChild(txt);
      myNodelist[i].appendChild(span);
    }
    // Click on a close button to hide the current list item
    addCloseEvent();
  }
});