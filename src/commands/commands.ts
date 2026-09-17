import { insertBlueParagraphInWord } from "./word";

/* global Office */

// Register the add-in commands with the Office host application.
Office.onReady(async (info) => {
  switch (info.host) {
    case Office.HostType.Word:
      Office.actions.associate("action", insertBlueParagraphInWord);
      break;
    default: {
      throw new Error(`${info.host} not supported.`);
    }
  }
});
