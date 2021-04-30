/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("addSignatures").onclick = addSign;

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
      // func add sign
    function addSign(){
      var sign = document.querySelector('[name="email"]').value;

        if (sign === "")
          return;

      var ul = document.getElementById("myList");
      var li = document.createElement("li");
      li.appendChild(document.createTextNode(sign));
      ul.appendChild(li);

      var span = document.createElement("SPAN");
      var txt = document.createTextNode("\u00D7");
      span.className = "close";
      span.appendChild(txt);
      li.appendChild(span);
      addCloseEvent();
      document.querySelector('[name="email"]').value = "";
    }
    //function add closeevent
    function addCloseEvent(){
      // Click on a close button to hide the current list item
      var i;
      var close = document.getElementsByClassName("close");

      for (i = 0; i < close.length; i++) {
        close[i].onclick = function() {
        var div = this.parentElement;
        div.style.display = "none";
        }
      }
    }
});