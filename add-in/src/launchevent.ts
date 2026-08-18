import { onMessageSendHandler } from "./onsend";

interface MailboxEvent {
  completed: (options: { allowEvent: boolean; errorMessage?: string }) => void;
}

/**
 * Register the OnMessageSend handler AFTER Office.js is fully initialized.
 * This prevents "Office.js has not fully loaded" errors in the Smart Alerts runtime.
 */
Office.onReady(function () {
  Office.actions.associate("onMessageSendHandler", onMessageSendLaunch);
});

/**
 * OnMessageSend launch event handler.
 *
 * Strategy: always allow the send. Journaling is fire-and-forget.
 * A 2-second safety net ensures the send always proceeds regardless
 * of what happens in the handler.
 */
function onMessageSendLaunch(event: MailboxEvent): void {
  let completed = false;

  // Safety net: allow send after 2 seconds no matter what
  const timer = setTimeout(() => {
    if (!completed) {
      completed = true;
      event.completed({ allowEvent: true });
    }
  }, 2000);

  // Fire the handler as a promise. If it completes before the timer,
  // use its result. Otherwise the timer handles it.
  onMessageSendHandler()
    .then((result) => {
      if (!completed) {
        completed = true;
        clearTimeout(timer);
        event.completed({ allowEvent: result.allowEvent !== false });
      }
    })
    .catch(() => {
      if (!completed) {
        completed = true;
        clearTimeout(timer);
        event.completed({ allowEvent: true });
      }
    });
}