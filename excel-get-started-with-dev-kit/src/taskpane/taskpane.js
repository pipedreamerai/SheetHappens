/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* eslint-disable prettier/prettier */
/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideload = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideload) sideload.style.display = "none";
    if (appBody) appBody.classList.remove("is-hidden");
  }
});
