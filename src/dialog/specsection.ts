/*
 * Specification Section dialog.
 *
 * Shown as an Office Dialog (a real popup window) so the spec section has room, instead of
 * being cramped in the task pane. On ready it tells the parent it is listening, then renders
 * whatever section payload the parent sends (and re-renders if the parent sends another).
 */

/* global Office, document, HTMLElement */

interface SectionMessage {
  section: string;
  outline: string;
  itemName: string;
  title: string;
  text: string;
}

function render(data: SectionMessage): void {
  const crumb = document.getElementById("crumb") as HTMLElement;
  const title = document.getElementById("title") as HTMLElement;
  const body = document.getElementById("body") as HTMLElement;
  crumb.textContent = `${data.section} ${data.outline} — ${data.itemName}`.trim();
  title.textContent = data.title || "Specification Section";
  if (data.text) {
    body.textContent = data.text;
  } else {
    body.innerHTML = '<span class="empty">No section text was found for this item.</span>';
  }
  document.title = data.title || "Specification Section";
}

Office.onReady(() => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    try {
      render(JSON.parse((arg as { message: string }).message) as SectionMessage);
    } catch {
      // ignore malformed messages
    }
  });
  // Tell the parent we're ready to receive the section payload.
  Office.context.ui.messageParent("ready");
});
